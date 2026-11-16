using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice.Setup;

public sealed class InstallOptions
{
    public string Directory { get; set; } = Installer.DefaultDirectory;
    public bool DesktopShortcuts { get; set; } = true;
    public bool StartMenu { get; set; } = true;
    public bool FileAssociations { get; set; } = true;
    /// <summary>Automatic update checks (the only network access the apps make on their own). Shown in setup with the privacy policy.</summary>
    public bool AutoUpdate { get; set; } = OfficeConfig.Current.AutoCheckUpdates;
    public string Language { get; set; } = Loc.Lang;
}

/// <summary>
/// Per-user installation (no administrator rights needed, like VS Code "user setup"):
/// copies the single executable, creates Start menu / Desktop shortcuts for Word, Excel and PDF,
/// registers file types for "Open with" and an entry in Settings › Apps for uninstalling.
/// Environment variable MICROADS_SETUP_SANDBOX=folder redirects shortcuts and registry writes
/// to a sandbox (used by the automated tests so the real system is never touched).
/// </summary>
public static class Installer
{
    public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\MicroadsOffice";
    public const string CapabilitiesKey = @"Software\Microads\Office\Capabilities";

    public static string DefaultDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Microads Office");

    private static string? Sandbox => Environment.GetEnvironmentVariable("MICROADS_SETUP_SANDBOX") is { Length: > 0 } s ? s : null;

    private static RegistryKey Root => Sandbox != null
        ? Registry.CurrentUser.CreateSubKey(@"Software\MicroadsOfficeSetupSandbox")
        : Registry.CurrentUser;

    public static string StartMenuDir => Sandbox != null ? Path.Combine(Sandbox, "StartMenu", "Microads Office")
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "Microads Office");

    public static string DesktopDir => Sandbox != null ? Path.Combine(Sandbox, "Desktop")
        : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

    public static readonly (AppKind Kind, string ProgId, string[] Exts, string TypeVi, string TypeEn)[] Apps =
    [
        (AppKind.Word, "MicroadsOffice.Word.Document", Product.WordExtensions, "Tài liệu Microads Office Word", "Microads Office Word document"),
        (AppKind.Excel, "MicroadsOffice.Excel.Sheet", Product.ExcelExtensions, "Bảng tính Microads Office Excel", "Microads Office Excel workbook"),
        (AppKind.Pdf, "MicroadsOffice.Pdf.Document", Product.PdfExtensions, "Tài liệu PDF", "PDF document"),
    ];

    // ------------------------------------------------------------------ entry points

    public static int RunSetup(bool silent, string? dir, string? lang, bool noAutoUpdate = false)
    {
        if (lang != null) Loc.Lang = Loc.Normalize(lang);
        if (silent)
        {
            var o = new InstallOptions { Directory = dir ?? DefaultDirectory, Language = Loc.Lang };
            if (noAutoUpdate) o.AutoUpdate = false;
            try { Install(o, null); return 0; }
            catch (Exception ex) { LogError(ex); return 1; }
        }
        using var f = new SetupForm(dir, noAutoUpdate);
        Application.Run(f);
        return f.ExitCode;
    }

    public static int RunUninstall(bool silent)
    {
        var dir = AppContext.BaseDirectory.TrimEnd('\\');
        bool removeData = false;
        if (!silent)
        {
            if (!Msg.Confirm(null, Loc.T("Gỡ cài đặt Microads Office (Word, Excel, PDF) khỏi máy tính này?", "Uninstall Microads Office (Word, Excel, PDF) from this computer?"),
                    Loc.T("Gỡ cài đặt", "Uninstall"), Loc.T("Hủy", "Cancel"))) return 1;
            removeData = Msg.Confirm(null, Loc.T("Xóa luôn dữ liệu cá nhân (cấu hình, mẫu chữ ký, ID số tự tạo) trong thư mục Data?",
                    "Also delete personal data (settings, signature styles, self-created digital IDs) in the Data folder?"),
                Loc.T("Xóa dữ liệu", "Delete data"), Loc.T("Giữ lại", "Keep"));
        }
        Uninstall(dir, removeData);
        if (!silent) Msg.Info(null, Loc.T("Đã gỡ Microads Office. Cảm ơn bạn đã sử dụng!\nmicroads.vn", "Microads Office has been removed. Thank you for using it!\nmicroads.vn"));
        return 0;
    }

    // ------------------------------------------------------------------ install

    /// <summary>Installs the running executable. <paramref name="progress"/> receives (percent, status).</summary>
    public static string Install(InstallOptions o, Action<int, string>? progress)
    {
        void P(int pct, string vi, string en) => progress?.Invoke(pct, Loc.T(vi, en));
        var src = Environment.ProcessPath ?? throw new InvalidOperationException("no process path");
        var dir = Path.GetFullPath(o.Directory);
        System.IO.Directory.CreateDirectory(dir);
        var exe = Path.Combine(dir, Product.ExeName);

        P(10, "Đang chép chương trình...", "Copying program files...");
        if (!string.Equals(Path.GetFullPath(src), Path.GetFullPath(exe), StringComparison.OrdinalIgnoreCase))
        {
            if (File.Exists(exe))
            {
                // A running copy can be renamed but not overwritten.
                var old = exe + ".old-" + Environment.ProcessId;
                try { File.Move(exe, old, true); } catch { }
            }
            // Write the program as a new file (like MSI / Inno Setup) instead of File.Copy: CopyFile also copies the
            // ":Zone.Identifier" stream of the downloaded setup, which would make Windows show its download
            // warning again every time an installed app is started. The user already chose to run the setup.
            using (var input = new FileStream(src, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var output = new FileStream(exe, FileMode.Create, FileAccess.Write, FileShare.None))
                input.CopyTo(output);
        }

        P(35, "Đang tạo biểu tượng...", "Creating icons...");
        var iconDir = Path.Combine(dir, "Icons");
        System.IO.Directory.CreateDirectory(iconDir);
        foreach (var k in new[] { AppKind.Hub, AppKind.Word, AppKind.Excel, AppKind.Pdf })
            BrandArt.WriteIcon(IconPath(dir, k), k);

        P(50, "Đang tạo lối tắt...", "Creating shortcuts...");
        if (o.StartMenu)
        {
            System.IO.Directory.CreateDirectory(StartMenuDir);
            foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf, AppKind.Hub })
                CreateShortcut(Path.Combine(StartMenuDir, Product.AppName(k) + ".lnk"), exe, ArgsFor(k), IconPath(dir, k), Product.AppDescription(k));
        }
        if (o.DesktopShortcuts)
        {
            System.IO.Directory.CreateDirectory(DesktopDir);
            foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf })
                CreateShortcut(Path.Combine(DesktopDir, Product.AppName(k) + ".lnk"), exe, ArgsFor(k), IconPath(dir, k), Product.AppDescription(k));
        }

        P(70, "Đang đăng ký loại tệp...", "Registering file types...");
        RegisterFileTypes(dir, exe, o.FileAssociations);

        P(85, "Đang đăng ký gỡ cài đặt...", "Registering the uninstaller...");
        using (var k = Root.CreateSubKey(UninstallKey))
        {
            k.SetValue("DisplayName", "Microads Office");
            k.SetValue("DisplayVersion", Product.VersionText);
            k.SetValue("Publisher", Loc.T("Công ty TNHH Công nghệ Microads", "Microads Technology Co., Ltd."));
            k.SetValue("DisplayIcon", IconPath(dir, AppKind.Hub));
            k.SetValue("InstallLocation", dir);
            k.SetValue("UninstallString", $"\"{exe}\" --uninstall");
            k.SetValue("QuietUninstallString", $"\"{exe}\" --uninstall --silent");
            k.SetValue("URLInfoAbout", Product.Website);
            k.SetValue("URLUpdateInfo", Product.UpdatePage);
            k.SetValue("HelpLink", Product.Website);
            k.SetValue("NoModify", 1, RegistryValueKind.DWord);
            k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            k.SetValue("EstimatedSize", (int)(new FileInfo(exe).Length / 1024), RegistryValueKind.DWord);
            k.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
        }

        P(95, "Đang lưu cài đặt ngôn ngữ...", "Saving language settings...");
        var prevBase = AppSettings.BaseDirectory;
        AppSettings.BaseDirectory = dir;
        try { OfficeConfig.Update(c => { c.Language = Loc.Normalize(o.Language); c.AutoCheckUpdates = o.AutoUpdate; }); }
        finally { AppSettings.BaseDirectory = prevBase; }

        if (Sandbox == null) SHChangeNotify(0x08000000, 0, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED
        P(100, "Hoàn tất", "Done");
        return exe;
    }

    public static string IconPath(string dir, AppKind k) => Path.Combine(dir, "Icons", k.ToString().ToLowerInvariant() + ".ico");

    public static string ArgsFor(AppKind k) => k == AppKind.Hub ? "" : "--app=" + k.ToString().ToLowerInvariant();

    private static void RegisterFileTypes(string dir, string exe, bool takeDefaults)
    {
        using var classes = Root.CreateSubKey(@"Software\Classes");
        using (var cap = Root.CreateSubKey(CapabilitiesKey))
        {
            cap.SetValue("ApplicationName", "Microads Office");
            cap.SetValue("ApplicationDescription", Loc.T("Bộ ứng dụng văn phòng miễn phí: Word, Excel, PDF — microads.vn", "Free office suite: Word, Excel, PDF — microads.vn"));
            cap.SetValue("ApplicationIcon", IconPath(dir, AppKind.Hub));
        }
        using (var reg = Root.CreateSubKey(@"Software\RegisteredApplications"))
            reg.SetValue("Microads Office", CapabilitiesKey);
        using var fa = Root.CreateSubKey(CapabilitiesKey + @"\FileAssociations");
        foreach (var (kind, progId, exts, vi, en) in Apps)
        {
            using (var p = classes.CreateSubKey(progId))
            {
                p.SetValue("", Loc.T(vi, en));
                p.SetValue("FriendlyTypeName", Loc.T(vi, en));
                using (var ic = p.CreateSubKey("DefaultIcon")) ic.SetValue("", IconPath(dir, kind));
                using (var cmd = p.CreateSubKey(@"shell\open\command")) cmd.SetValue("", $"\"{exe}\" --app={kind.ToString().ToLowerInvariant()} \"%1\"");
                using (var app = p.CreateSubKey("Application"))
                {
                    app.SetValue("ApplicationName", Product.AppName(kind));
                    app.SetValue("ApplicationCompany", "Microads");
                }
            }
            foreach (var ext in exts)
            {
                using var e = classes.CreateSubKey(ext);
                using (var owp = e.CreateSubKey("OpenWithProgids")) owp.SetValue(progId, Array.Empty<byte>(), RegistryValueKind.None);
                // Become the default only for extensions that have no handler yet (Windows protects user choices).
                if (takeDefaults && string.IsNullOrEmpty(Registry.ClassesRoot.OpenSubKey(ext)?.GetValue("") as string) && Sandbox == null)
                    e.SetValue("", progId);
                else if (takeDefaults && Sandbox != null) e.SetValue("", progId);
                fa.SetValue(ext, progId);
            }
        }
        using (var a = classes.CreateSubKey(@"Applications\" + Product.ExeName))
        {
            a.SetValue("FriendlyAppName", "Microads Office");
            using (var cmd = a.CreateSubKey(@"shell\open\command")) cmd.SetValue("", $"\"{exe}\" \"%1\"");
            using var st = a.CreateSubKey("SupportedTypes");
            foreach (var ext in Apps.SelectMany(x => x.Exts)) st.SetValue(ext, "");
        }
    }

    // ------------------------------------------------------------------ uninstall

    public static void Uninstall(string dir, bool removeData)
    {
        foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf, AppKind.Hub })
        {
            TryDelete(Path.Combine(StartMenuDir, Product.AppName(k) + ".lnk"));
            TryDelete(Path.Combine(DesktopDir, Product.AppName(k) + ".lnk"));
        }
        try { if (System.IO.Directory.Exists(StartMenuDir)) System.IO.Directory.Delete(StartMenuDir, true); } catch { }
        try
        {
            using var classes = Root.CreateSubKey(@"Software\Classes");
            foreach (var (_, progId, exts, _, _) in Apps)
            {
                classes.DeleteSubKeyTree(progId, false);
                foreach (var ext in exts)
                {
                    using var e = classes.OpenSubKey(ext, true);
                    if (e == null) continue;
                    using (var owp = e.OpenSubKey("OpenWithProgids", true)) owp?.DeleteValue(progId, false);
                    if (e.GetValue("") as string == progId) e.DeleteValue("", false);
                }
            }
            classes.DeleteSubKeyTree(@"Applications\" + Product.ExeName, false);
            using (var reg = Root.OpenSubKey(@"Software\RegisteredApplications", true)) reg?.DeleteValue("Microads Office", false);
            Root.DeleteSubKeyTree(@"Software\Microads\Office", false);
            Root.DeleteSubKeyTree(UninstallKey, false);
        }
        catch { }
        if (Sandbox == null) SHChangeNotify(0x08000000, 0, IntPtr.Zero, IntPtr.Zero);

        // The running executable cannot delete itself: a hidden helper removes the folder after we exit.
        var targets = new List<string> { Path.Combine(dir, Product.ExeName), Path.Combine(dir, "Icons"), Path.Combine(dir, "pdfium.dll") };
        if (removeData) targets.Add(Path.Combine(dir, "Data"));
        var script = string.Join(" & ", targets.Select(t => $"(if exist \"{t}\\*\" (rmdir /s /q \"{t}\") else (del /f /q \"{t}\"))"));
        script += $" & del /f /q \"{dir}\\{Product.ExeName}.old-*\" & rmdir \"{dir}\"";
        try
        {
            Process.Start(new ProcessStartInfo("cmd.exe", $"/c ping 127.0.0.1 -n 3 > nul & {script}")
            {
                CreateNoWindow = true, UseShellExecute = false, WorkingDirectory = Path.GetTempPath(), WindowStyle = ProcessWindowStyle.Hidden
            });
        }
        catch { }
    }

    private static void TryDelete(string f) { try { if (File.Exists(f)) File.Delete(f); } catch { } }

    // ------------------------------------------------------------------ helpers

    public static void CreateShortcut(string lnk, string target, string args, string icon, string description)
    {
        var t = Type.GetTypeFromProgID("WScript.Shell") ?? throw new InvalidOperationException("WScript.Shell unavailable");
        dynamic shell = Activator.CreateInstance(t)!;
        try
        {
            dynamic sc = shell.CreateShortcut(lnk);
            sc.TargetPath = target;
            sc.Arguments = args;
            sc.IconLocation = icon + ",0";
            sc.Description = description;
            sc.WorkingDirectory = Path.GetDirectoryName(target);
            sc.Save();
            Marshal.FinalReleaseComObject(sc);
        }
        finally { Marshal.FinalReleaseComObject(shell); }
    }

    public static bool IsInstalledAt(string dir) => File.Exists(Path.Combine(dir, Product.ExeName));

    public static string? InstalledLocation()
    {
        using var k = Root.OpenSubKey(UninstallKey);
        return k?.GetValue("InstallLocation") as string;
    }

    public static void LogError(Exception ex)
    {
        try
        {
            var log = Path.Combine(Path.GetTempPath(), "MicroadsOffice_Setup.log");
            File.AppendAllText(log, $"[{DateTime.Now:s}] {ex}\n");
        }
        catch { }
    }

    [DllImport("shell32.dll")] private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
