using System.Diagnostics;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;

namespace MicroadsOffice.UI;

/// <summary>Implemented by the main window of every Microads Office application.</summary>
public interface IOfficeAppWindow
{
    /// <summary>Opens files (from the command line, another instance, drag &amp; drop or the hub).</summary>
    void OpenFiles(IReadOnlyList<string> paths, bool fromShell);

    /// <summary>Saved documents currently open (used to reopen them after a language switch / update restart).</summary>
    IReadOnlyList<string> OpenDocumentPaths { get; }
}

/// <summary>
/// Services shared by Word, Excel, PDF and the hub: identity, language, about, promo notice,
/// update checks (manual and "update everyone" orders from microads.vn/office).
/// </summary>
public static class OfficeShell
{
    public static AppKind App { get; private set; } = AppKind.Hub;
    public static Icon? AppIcon { get; private set; }

    /// <summary>Arguments to start again with after the main window closed (language switch / update).</summary>
    public static string? PendingRestartArgs { get; private set; }

    private static System.Windows.Forms.Timer? _updateTimer;
    private static bool _checking;
    private static bool _updateApplied;

    public static void Init(AppKind app)
    {
        App = app;
        var cfg = OfficeConfig.Reload();
        Loc.Lang = Loc.Normalize(cfg.Language);
        var (accent, dark) = BrandArt.AccentOf(app);
        Theme.UseAccent(accent, dark);
        try
        {
            using var g = Graphics.FromHwnd(IntPtr.Zero);
            Dpi.Factor = g.DpiX / 96f;
        }
        catch { }
        try { AppIcon = BrandArt.CreateIcon(app); DialogBase.AppIcon = AppIcon; } catch { }
    }

    // ------------------------------------------------------------------ language

    /// <summary>First run: asks the user to choose Vietnamese or English (big flag buttons).</summary>
    public static void EnsureLanguageChosen(IWin32Window? owner)
    {
        if (OfficeConfig.Current.Language != null) return;
        using var d = new LanguageDialog(firstRun: true);
        d.ShowDialog(owner);
        SetLanguage(d.Selected);
    }

    public static void SetLanguage(string lang)
    {
        lang = Loc.Normalize(lang);
        OfficeConfig.Update(c => c.Language = lang);
        Loc.Lang = lang;
    }

    /// <summary>Switches language; the window is closed (asking to save) and reopened in the new language with the same documents.</summary>
    public static void ChangeLanguage(Form main, string lang)
    {
        lang = Loc.Normalize(lang);
        if (lang == Loc.Lang) return;
        var old = Loc.Lang;
        SetLanguage(lang);
        RestartApp(main, () => { SetLanguage(old); });
    }

    /// <summary>Closes the main window (normal save prompts apply) and starts the app again with its open documents.</summary>
    public static void RestartApp(Form main, Action? onCancelled = null)
    {
        var docs = (main as IOfficeAppWindow)?.OpenDocumentPaths ?? [];
        PendingRestartArgs = $"--app={App.ToString().ToLowerInvariant()} " + string.Join(" ", docs.Select(p => "\"" + p + "\""));
        main.Close();
        if (!main.IsDisposed && main.Visible)
        {
            // closing was cancelled by the user
            PendingRestartArgs = null;
            onCancelled?.Invoke();
        }
    }

    /// <summary>Flag drop-down for the ribbon header: shows the current language, lets the user switch.</summary>
    public static ToolStripDropDownButton CreateLanguageButton(Form main)
    {
        var b = new ToolStripDropDownButton(Flags.LanguageName(Loc.Lang), Flags.Small(Loc.Lang))
        {
            Name = "btnLanguage",
            AccessibleName = Loc.T("Ngôn ngữ", "Language"),
            ToolTipText = Loc.T("Đổi ngôn ngữ / Change language", "Change language / Đổi ngôn ngữ"),
            DisplayStyle = ToolStripItemDisplayStyle.ImageAndText,
            ImageScaling = ToolStripItemImageScaling.None,
            Alignment = ToolStripItemAlignment.Right
        };
        b.DropDown = new ContextMenuStrip { Renderer = new RibbonRenderer(false), ImageScalingSize = new Size(Dpi.Scale(21), Dpi.Scale(14)) };
        foreach (var lang in new[] { Loc.Vietnamese, Loc.English })
        {
            var l = lang;
            var it = new ToolStripMenuItem(Flags.LanguageName(l), Flags.Small(l), (_, _) => ChangeLanguage(main, l))
            {
                Name = "miLang_" + l, Checked = Loc.Lang == l, ImageScaling = ToolStripItemImageScaling.None
            };
            b.DropDownItems.Add(it);
        }
        return b;
    }

    /// <summary>"microads.vn" link for the ribbon header.</summary>
    public static ToolStripButton CreateWebsiteButton()
    {
        var p = PromoInfo.Effective;
        var b = new ToolStripButton(p.WebsiteLabel)
        {
            Name = "btnWebsite",
            ForeColor = Theme.AccentDark,
            Alignment = ToolStripItemAlignment.Right,
            ToolTipText = p.Company + " — " + p.Website,
            AccessibleName = "microads.vn"
        };
        b.Click += (_, _) => OpenWebsite();
        return b;
    }

    public static ToolStripButton CreateUpdateButton(IWin32Window owner)
    {
        var b = new ToolStripButton(Icons.Small(Icons.Sync))
        {
            Name = "btnCheckUpdateHeader",
            DisplayStyle = ToolStripItemDisplayStyle.Image,
            Alignment = ToolStripItemAlignment.Right,
            ToolTipText = Loc.T("Kiểm tra cập nhật", "Check for updates"),
            AccessibleName = Loc.T("Kiểm tra cập nhật", "Check for updates")
        };
        b.Click += (_, _) => CheckForUpdatesInteractive(owner);
        return b;
    }

    /// <summary>Standard right-side header items: language flag, update, microads.vn.</summary>
    public static ToolStripItem[] CreateHeaderItems(Form main) =>
        [CreateLanguageButton(main), CreateUpdateButton(main), CreateWebsiteButton()];

    public static void OpenWebsite(string? url = null)
    {
        url ??= PromoInfo.Effective.Website;
        if (!PromoInfo.IsSafeLink(url)) url = Product.Website;
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { }
    }

    // ------------------------------------------------------------------ dialogs

    public static void ShowAbout(IWin32Window owner)
    {
        using var d = new AboutOfficeDialog();
        d.ShowDialog(owner);
    }

    public static void ShowSettings(Form main)
    {
        using var d = new OfficeSettingsDialog();
        if (d.ShowDialog(main) == DialogResult.OK && d.SelectedLanguage != Loc.Lang)
            ChangeLanguage(main, d.SelectedLanguage);
    }

    /// <summary>Shows the Microads introduction when a document is opened with Word / Excel / PDF (Explorer, "Open with", another instance).</summary>
    public static void OnFileOpenedFromShell(Form owner)
    {
        var p = PromoInfo.Effective;
        if (!p.Enabled || !p.ShowOnFileOpen) return;
        ShowPromo(owner);
    }

    /// <summary>Shows the Microads introduction every time Word, Excel or PDF starts (switchable from microads.vn/office/admin).</summary>
    public static void OnAppStarted(Form owner)
    {
        var p = PromoInfo.Effective;
        if (!p.Enabled || !p.ShowOnStart) return;
        ShowPromo(owner);
    }

    private static void ShowPromo(Form owner)
    {
        if (Environment.GetEnvironmentVariable("MICROADS_NO_PROMO") == "1") return;
        try
        {
            if (owner.IsDisposed) return;
            owner.BeginInvoke(() =>
            {
                if (owner.IsDisposed) return;
                if (Application.OpenForms.OfType<PromoPopup>().Any()) return;
                var popup = new PromoPopup();
                popup.Show(owner);
            });
        }
        catch { }
    }

    // ------------------------------------------------------------------ updates

    public static void CheckForUpdatesInteractive(IWin32Window owner)
    {
        using var d = new UpdateDialog(owner as Form);
        d.ShowDialog(owner);
    }

    /// <summary>
    /// Starts background work after the main window is shown: removes leftovers of old updates,
    /// refreshes promo content, honours "update everyone" orders, and informs about new versions.
    /// </summary>
    public static void StartBackgroundServices(Form main)
    {
        UpdateService.CleanupOldFiles();
        var just = OfficeConfig.Current.JustUpdatedTo;
        if (!string.IsNullOrEmpty(just))
        {
            OfficeConfig.Update(c => c.JustUpdatedTo = null);
            if (just == Product.VersionText)
                Toast.Show(main, Loc.F("Microads Office đã được cập nhật lên phiên bản {0}.", "Microads Office has been updated to version {0}.", just), null, null);
        }
        if (Environment.GetEnvironmentVariable("MICROADS_NO_UPDATE") == "1") return;
        // Privacy: with "check automatically" switched off (installer or Settings) the app makes NO network request on its own;
        // "Check for updates" still works when the user clicks it.
        if (!OfficeConfig.Current.AutoCheckUpdates) return;
        _ = BackgroundCheckAsync(main);
        _updateTimer?.Dispose();
        _updateTimer = new System.Windows.Forms.Timer { Interval = 2 * 60 * 60 * 1000 };
        _updateTimer.Tick += (_, _) => _ = BackgroundCheckAsync(main);
        _updateTimer.Start();
        main.FormClosed += (_, _) => { _updateTimer?.Dispose(); _updateTimer = null; };
    }

    private static async Task BackgroundCheckAsync(Form main)
    {
        if (_checking || _updateApplied || !OfficeConfig.Current.AutoCheckUpdates) return;
        _checking = true;
        try
        {
            await Task.Delay(1500);
            var m = await Task.Run(() => UpdateService.FetchManifestAsync());
            // Microsoft Store edition: the manifest still refreshes promo content, but updates come from the Store.
            if (StorePackage.IsPackaged) return;
            if (m?.Latest == null || main.IsDisposed) return;
            if (UpdateService.IsForced(m) && !OfficeConfig.Current.AutoInstallForcedUpdates)
            {
                // "Update everyone" order, but the user prefers not to be interrupted: notice only, they click Update.
                if (!main.IsDisposed)
                    Toast.Show(main, Loc.F("Microads khuyên cập nhật lên phiên bản {0}.", "Microads recommends updating to version {0}.", m.Latest.Version),
                        Loc.T("Cập nhật", "Update"), () => CheckForUpdatesInteractive(main));
                return;
            }
            if (UpdateService.IsForced(m))
            {
                // "Update everyone" order from microads.vn/office: install silently, then offer a restart.
                var bytes = await Task.Run(() => UpdateService.DownloadAsync(m.Latest, null, CancellationToken.None));
                var res = UpdateService.Apply(bytes, m.Latest.Version, out _);
                if (res == ApplyResult.Failed) return;
                var cmd = m.ForceUpdate!.CommandId;
                OfficeConfig.Update(c => c.AppliedForceCommand = cmd);
                _updateApplied = true;
                if (res == ApplyResult.Applied && !main.IsDisposed)
                    Toast.Show(main, Loc.F("Đã tải và cài bản cập nhật {0}. Khởi động lại để dùng phiên bản mới.", "Update {0} has been installed. Restart to use the new version.", m.Latest.Version),
                        Loc.T("Khởi động lại", "Restart"), () => RestartApp(main));
                return;
            }
            if (OfficeConfig.Current.AutoCheckUpdates && UpdateService.IsNewer(m.Latest) && OfficeConfig.Current.SkippedVersion != m.Latest.Version && !main.IsDisposed)
            {
                Toast.Show(main, Loc.F("Đã có phiên bản mới {0} của Microads Office.", "Microads Office {0} is available.", m.Latest.Version),
                    Loc.T("Cập nhật", "Update"), () => CheckForUpdatesInteractive(main));
            }
        }
        catch (Exception ex) { LogUpdate("background check: " + ex); /* offline or invalid update: try again later */ }
        finally { _checking = false; }
    }

    internal static void MarkUpdateApplied() => _updateApplied = true;

    /// <summary>Update diagnostics for support (Data/Logs/update.log, kept small).</summary>
    public static void LogUpdate(string line)
    {
        try
        {
            var dir = Path.Combine(AppSettings.DataDirectory, "Logs");
            Directory.CreateDirectory(dir);
            var f = Path.Combine(dir, "update.log");
            if (File.Exists(f) && new FileInfo(f).Length > 256 * 1024) File.Delete(f);
            File.AppendAllText(f, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {App} v{Product.VersionText}: {line}{Environment.NewLine}");
        }
        catch { }
    }
}
