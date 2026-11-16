using System.Diagnostics;
using System.Reflection;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.Setup;
using MicroadsOffice.UI;

namespace MicroadsOffice;

/// <summary>
/// One executable for the whole suite:
///   MicroadsOffice.exe                      → Microads Office hub (choose Word / Excel / PDF)
///   MicroadsOffice.exe --app=word|excel|pdf → that application
///   MicroadsOffice.exe file.docx|xlsx|pdf   → the matching application opens the file
///   MicroadsOffice_Setup_x.y.z.exe          → installer (any file name containing "setup"), also --install
///   MicroadsOffice.exe --uninstall          → removes shortcuts, file associations and the program
///   setup option --no-auto-update            → installs with automatic update checks switched off (no network access on its own)
/// CONTRACT with the web installer (src/MicroadsOffice.WebSetup — a frozen binary that never changes) and with winget:
///   "--setup --web-setup --lang=vi|en [--silent] [--dir PATH]" and "--install --silent --dir PATH" must keep working in every version.
/// </summary>
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        // Build helper: writes the logo images of the Microsoft Store package (used by tools/store/make_msix.ps1).
        if (args.Length == 2 && args[0] == "--export-store-assets") return StoreAssets.Export(args[1]);

        // The Microsoft Store edition is installed, updated and removed by Windows: no built-in installer there.
        bool storeEdition = StorePackage.IsPackaged;
        bool setupMode = !storeEdition && (args.Any(a => a.Equals("--install", StringComparison.OrdinalIgnoreCase) || a.Equals("--setup", StringComparison.OrdinalIgnoreCase))
            || Path.GetFileName(Environment.ProcessPath ?? "").Contains("setup", StringComparison.OrdinalIgnoreCase));
        // The installer must not leave a "Data" folder next to the downloaded setup file: it only reads
        // the language of an existing installation.
        if (setupMode && Environment.GetEnvironmentVariable("MICROADS_DATA_DIR") is not { Length: > 0 })
            AppSettings.BaseDirectory = Setup.Installer.DefaultDirectory;
        else ConfigureDataDirectory();
        PdfiumLoader.Register(Assembly.GetExecutingAssembly());
        AppearanceRenderer.PlatformVersion = Product.VersionText;
        UpdateService.CurrentVersion = Product.Version;

        ApplicationConfiguration.Initialize();
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) => ReportError(e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, e) => { if (e.ExceptionObject is Exception ex) ReportError(ex); };

        var opts = args.Where(a => a.StartsWith("--")).ToList();
        string? Opt(string name) => opts.FirstOrDefault(o => o.StartsWith("--" + name + "=", StringComparison.OrdinalIgnoreCase))?.Split('=', 2)[1];
        bool Has(string name) => opts.Any(o => o.Equals("--" + name, StringComparison.OrdinalIgnoreCase));
        string? ArgAfter(string name)
        {
            int i = Array.FindIndex(args, a => a.Equals("--" + name, StringComparison.OrdinalIgnoreCase));
            return i >= 0 && i + 1 < args.Length ? args[i + 1] : Opt(name);
        }

        // ------------------------------------------------ installer / uninstaller
        var exeName = Path.GetFileName(Environment.ProcessPath ?? "");
        if (Has("uninstall") && !storeEdition)
        {
            OfficeShell.Init(AppKind.Hub);
            return Installer.RunUninstall(Has("silent"));
        }
        if (setupMode)
        {
            OfficeShell.Init(AppKind.Hub);
            return Installer.RunSetup(Has("silent"), ArgAfter("dir"), Opt("lang"), Has("no-auto-update"));
        }

        // ------------------------------------------------ which application?
        var files = args.Where(a => !a.StartsWith("--") && File.Exists(a)).Select(Path.GetFullPath).ToList();
        var kind = Product.ParseKind(Opt("app"));
        if (kind == null && files.Count > 0) kind = Product.KindForFile(files[0]) ?? AppKind.Pdf;
        kind ??= AppKind.Hub;

        // files for other apps (e.g. "open with" on a mixed selection) go to their own process
        foreach (var group in files.GroupBy(f => Product.KindForFile(f) ?? AppKind.Pdf).Where(g => g.Key != kind && kind != AppKind.Hub))
            Launch(group.Key, group);
        files = files.Where(f => (Product.KindForFile(f) ?? AppKind.Pdf) == kind || kind == AppKind.Hub).ToList();
        if (kind == AppKind.Hub && files.Count > 0) { foreach (var g in files.GroupBy(f => Product.KindForFile(f) ?? AppKind.Pdf)) Launch(g.Key, g); return 0; }

        OfficeShell.Init(kind.Value);

        // single instance per application and data folder: forward files to the running window
        var key = SingleInstance.KeyFor(AppSettings.DataDirectory + "|" + kind);
        var mutex = new Mutex(true, key, out bool firstInstance);
        if (!firstInstance && SingleInstance.TrySend(key, files)) return 0;

        OfficeShell.EnsureLanguageChosen(null);

        Form form = kind switch
        {
            AppKind.Word => MicroadsOffice.Word.WordApp.CreateMainForm(files.ToArray()),
            AppKind.Excel => MicroadsOffice.Excel.ExcelApp.CreateMainForm(files.ToArray()),
            AppKind.Pdf => PdfApp.CreateMainForm(files.ToArray()),
            _ => new Hub.HubForm()
        };
        if (form.Icon == null && OfficeShell.AppIcon != null) form.Icon = OfficeShell.AppIcon;
        bool fromShell = files.Count > 0 && !Has("no-promo");
        form.Shown += (_, _) =>
        {
            OfficeShell.StartBackgroundServices(form);
            if (fromShell) OfficeShell.OnFileOpenedFromShell(form);
        };
        SingleInstance.Listen(key, paths =>
        {
            try
            {
                form.BeginInvoke(() =>
                {
                    if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
                    form.Activate();
                    if (form is IOfficeAppWindow w && paths.Count > 0)
                    {
                        w.OpenFiles(paths, fromShell: true);
                        OfficeShell.OnFileOpenedFromShell(form);
                    }
                });
            }
            catch { }
        });
        Application.Run(form);

        var restart = OfficeShell.PendingRestartArgs;
        try { mutex.ReleaseMutex(); } catch { }
        mutex.Dispose();
        if (restart != null) UpdateService.Restart(restart + " --no-promo");
        return 0;
    }

    /// <summary>Starts another Microads Office application with files.</summary>
    public static void Launch(AppKind kind, IEnumerable<string> files)
    {
        try
        {
            var a = $"--app={kind.ToString().ToLowerInvariant()} " + string.Join(" ", files.Select(f => "\"" + f + "\""));
            Process.Start(new ProcessStartInfo(Environment.ProcessPath!, a) { UseShellExecute = false });
        }
        catch { }
    }

    /// <summary>
    /// Settings live in "Data" next to the executable (portable / per-user install). If that folder is
    /// not writable (e.g. installed for all users in Program Files) the user's AppData is used instead.
    /// </summary>
    private static void ConfigureDataDirectory()
    {
        var overrideDir = Environment.GetEnvironmentVariable("MICROADS_DATA_DIR");
        if (!string.IsNullOrWhiteSpace(overrideDir)) { AppSettings.BaseDirectory = overrideDir; return; }
        if (StorePackage.IsPackaged)
        {
            // Store edition: the install folder is read-only; LocalAppData is redirected by Windows into the
            // package's private storage and removed together with the app.
            AppSettings.BaseDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microads Office");
            return;
        }
        var exeDir = AppContext.BaseDirectory;
        AppSettings.BaseDirectory = exeDir;
        try
        {
            Directory.CreateDirectory(AppSettings.DataDirectory);
            var probe = Path.Combine(AppSettings.DataDirectory, ".write-test");
            File.WriteAllText(probe, "ok");
            File.Delete(probe);
        }
        catch
        {
            AppSettings.BaseDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microads Office");
        }
    }

    private static void ReportError(Exception ex)
    {
        try
        {
            var dir = Path.Combine(AppSettings.DataDirectory, "Logs");
            Directory.CreateDirectory(dir);
            File.AppendAllText(Path.Combine(dir, "error.log"), $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] v{Product.VersionText} {OfficeShell.App}: {ex}\n\n");
        }
        catch { }
        Msg.Warn(null, Loc.T("Đã xảy ra lỗi không mong muốn:\n", "An unexpected error occurred:\n") + ex.Message +
            Loc.T("\n\nỨng dụng vẫn tiếp tục hoạt động. Chi tiết được ghi tại Data/Logs/error.log.", "\n\nThe application keeps running. Details were written to Data/Logs/error.log."));
    }
}
