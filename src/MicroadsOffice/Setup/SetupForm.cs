using System.Diagnostics;
using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice.Setup;

/// <summary>Installer window: Language → Options → Installing → Finished (language flags + open apps).</summary>
public sealed class SetupForm : Form
{
    private readonly Panel _side, _pages, _buttons;
    private readonly Button _back, _next, _cancel;
    private int _step;
    private string _lang = Loc.Lang;
    private readonly InstallOptions _opts = new();
    private string? _installedExe;
    public int ExitCode { get; private set; } = 1;

    // option controls (kept to read values after a language switch)
    private TextBox? _dirBox;
    private CheckBox? _desktop, _assoc, _launch, _autoUpdate;
    private ProgressBar? _bar;
    private Label? _status;

    public SetupForm(string? dir, bool noAutoUpdate = false)
    {
        if (dir != null) _opts.Directory = dir;
        if (noAutoUpdate) _opts.AutoUpdate = false;
        Name = "setupForm";
        Text = Loc.T("Cài đặt Microads Office", "Microads Office Setup");
        Font = Theme.UiFont;
        AutoScaleMode = AutoScaleMode.None;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(Dpi.Scale(760), Dpi.Scale(480));
        BackColor = Color.White;
        Icon = BrandArt.CreateIcon(AppKind.Hub);

        _side = new Panel { Dock = DockStyle.Left, Width = Dpi.Scale(250), Name = "setupSide" };
        _side.Paint += PaintSide;
        _buttons = new Panel { Dock = DockStyle.Bottom, Height = Dpi.Scale(58), BackColor = Theme.PanelBack };
        _pages = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Dpi.Scale(30), Dpi.Scale(24), Dpi.Scale(30), Dpi.Scale(10)), Name = "setupPages" };
        _next = MkButton("btnSetupNext", true);
        _back = MkButton("btnSetupBack", false);
        _cancel = MkButton("btnSetupCancel", false);
        _next.Click += (_, _) => Next();
        _back.Click += (_, _) => { if (_step > 0) ShowStep(_step - 1); };
        _cancel.Click += (_, _) => Close();
        _buttons.Controls.AddRange([_next, _back, _cancel]);
        _buttons.Resize += (_, _) => LayoutButtons();
        Controls.Add(_pages);
        Controls.Add(_buttons);
        Controls.Add(_side);
        ShowStep(0);
    }

    private static Button MkButton(string name, bool primary)
    {
        var b = new Button
        {
            Name = name, Size = new Size(Dpi.Scale(120), Dpi.Scale(34)), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
            BackColor = primary ? BrandArt.Red : Color.White, ForeColor = primary ? Color.White : Theme.Ink,
            Font = primary ? new Font("Segoe UI Semibold", 10f) : new Font("Segoe UI", 10f)
        };
        b.FlatAppearance.BorderSize = primary ? 0 : 1;
        b.FlatAppearance.BorderColor = Theme.Border;
        return b;
    }

    private void LayoutButtons()
    {
        int y = (_buttons.Height - _next.Height) / 2, x = _buttons.Width - Dpi.Scale(20);
        foreach (var b in new[] { _next, _back, _cancel }.Where(b => b.Visible))
        {
            x -= b.Width;
            b.Location = new Point(x, y);
            x -= Dpi.Scale(10);
        }
    }

    private void PaintSide(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        var r = _side.ClientRectangle;
        using (var b = new LinearGradientBrush(r, Color.FromArgb(34, 36, 44), Color.FromArgb(20, 21, 26), LinearGradientMode.Vertical)) g.FillRectangle(b, r);
        using (var bar = new LinearGradientBrush(new Rectangle(0, 0, r.Width, Dpi.Scale(5)), BrandArt.Red, BrandArt.Green, LinearGradientMode.Horizontal))
            g.FillRectangle(bar, 0, 0, r.Width, Dpi.Scale(5));
        using (var icon = BrandArt.RenderSuiteIcon(Dpi.Scale(84))) g.DrawImage(icon, Dpi.Scale(26), Dpi.Scale(34));
        using var title = new Font("Segoe UI Semibold", 17f);
        using var small = new Font("Segoe UI", 9f);
        g.DrawString("Microads Office", title, Brushes.White, Dpi.Scale(24), Dpi.Scale(128));
        using var sub = new SolidBrush(Color.FromArgb(170, 255, 255, 255));
        g.DrawString(Loc.F("Phiên bản {0}", "Version {0}", Product.VersionText), small, sub, Dpi.Scale(27), Dpi.Scale(164));
        int y = Dpi.Scale(206);
        foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf })
        {
            using var ic = BrandArt.RenderAppIcon(Dpi.Scale(32), k);
            g.DrawImage(ic, Dpi.Scale(26), y);
            using var nf = new Font("Segoe UI Semibold", 10f);
            g.DrawString(Product.AppName(k), nf, Brushes.White, Dpi.Scale(66), y - Dpi.Scale(2));
            g.DrawString(Product.AppDescription(k), small, sub, Dpi.Scale(66), y + Dpi.Scale(16));
            y += Dpi.Scale(48);
        }
        var free = new Rectangle(Dpi.Scale(24), r.Height - Dpi.Scale(112), r.Width - Dpi.Scale(48), Dpi.Scale(28));
        using (var fb = new SolidBrush(BrandArt.Green)) using (var fp = BrandArt.RoundRect(free, Dpi.Scale(14))) g.FillPath(fb, fp);
        TextRenderer.DrawText(g, Product.AlwaysFree.ToUpperInvariant(), new Font("Segoe UI Semibold", 9f), free, Color.White, TextFormatFlags.NoPrefix | TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        TextRenderer.DrawText(g, Product.PublishedBy, small, new Rectangle(Dpi.Scale(24), r.Height - Dpi.Scale(76), r.Width - Dpi.Scale(40), Dpi.Scale(40)), Color.FromArgb(200, 200, 205), TextFormatFlags.NoPrefix | TextFormatFlags.WordBreak);
        TextRenderer.DrawText(g, "microads.vn", new Font("Segoe UI Semibold", 10f), new Point(Dpi.Scale(24), r.Height - Dpi.Scale(34)), Color.FromArgb(255, 120, 115), TextFormatFlags.NoPrefix);
    }

    private void Next()
    {
        switch (_step)
        {
            case 0: ShowStep(1); break;
            case 1: ReadOptions(); ShowStep(2); _ = RunInstall(); break;
            case 3: Finish(); break;
        }
    }

    private void ReadOptions()
    {
        if (_dirBox != null) _opts.Directory = _dirBox.Text.Trim();
        if (_desktop != null) _opts.DesktopShortcuts = _desktop.Checked;
        if (_assoc != null) _opts.FileAssociations = _assoc.Checked;
        if (_autoUpdate != null) _opts.AutoUpdate = _autoUpdate.Checked;
        if (_launch != null) _launchAfter = _launch.Checked;
        _opts.Language = _lang;
    }

    private void ShowStep(int step)
    {
        if (_step == 1 && step != 1) ReadOptions();
        _step = step;
        Text = Loc.T("Cài đặt Microads Office", "Microads Office Setup");
        _pages.SuspendLayout();
        foreach (Control c in _pages.Controls.Cast<Control>().ToList()) c.Dispose();
        _next.Visible = step != 2;
        _back.Visible = step == 1;
        _cancel.Visible = step < 2;
        _back.Text = Loc.T("Quay lại", "Back");
        _cancel.Text = Loc.T("Hủy", "Cancel");
        _next.Enabled = true;
        switch (step)
        {
            case 0: BuildLanguagePage(); _next.Text = Loc.T("Tiếp tục", "Next"); break;
            case 1: BuildOptionsPage(); _next.Text = Loc.T("Cài đặt", "Install"); break;
            case 2: BuildProgressPage(); break;
            case 3: BuildFinishPage(); _next.Text = Loc.T("Hoàn tất", "Finish"); break;
        }
        _pages.ResumeLayout();
        LayoutButtons();
        _side.Invalidate();
        AcceptButton = _next;
    }

    private Label Title(string text) => new()
    {
        Text = text, Dock = DockStyle.Top, Height = Dpi.Scale(44), Font = new Font("Segoe UI Semibold", 15f), ForeColor = Theme.Ink, Name = "lblSetupTitle"
    };

    private Label Para(string text, int h) => new() { Text = text, Dock = DockStyle.Top, Height = Dpi.Scale(h), ForeColor = Theme.SubtleInk, Font = new Font("Segoe UI", 10f) };

    private void BuildLanguagePage()
    {
        var row = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(176), WrapContents = false, Padding = new Padding(0, Dpi.Scale(6), 0, 0) };
        foreach (var l in new[] { Loc.Vietnamese, Loc.English })
        {
            var card = new LanguageCardButton(l, l == _lang) { Margin = new Padding(0, 0, Dpi.Scale(18), 0) };
            card.Click += (_, _) => { _lang = l; Loc.Lang = l; ShowStep(0); };
            row.Controls.Add(card);
        }
        _pages.Controls.Add(Para(Loc.T("Bạn có thể đổi lại ngôn ngữ bất cứ lúc nào bằng nút lá cờ trong ứng dụng.", "You can change the language at any time with the flag button inside the apps."), 40));
        _pages.Controls.Add(row);
        _pages.Controls.Add(Para("Microads Office Word · Excel · PDF — " + Product.AlwaysFree.ToLowerInvariant(), 34));
        _pages.Controls.Add(Title("Chọn ngôn ngữ · Choose language"));
    }

    private void BuildOptionsPage()
    {
        var lbl = new Label { Text = Loc.T("Thư mục cài đặt", "Install folder"), Dock = DockStyle.Top, Height = Dpi.Scale(24), Font = Theme.UiFontBold };
        var dirRow = new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(36) };
        _dirBox = new TextBox { Name = "txtInstallDir", Text = _opts.Directory, Font = new Font("Segoe UI", 10f), Location = new Point(0, Dpi.Scale(4)), Width = Dpi.Scale(330) };
        var browse = new Button { Name = "btnBrowseDir", Text = Loc.T("Chọn...", "Browse..."), Location = new Point(Dpi.Scale(338), Dpi.Scale(2)), Size = new Size(Dpi.Scale(86), Dpi.Scale(30)), FlatStyle = FlatStyle.Flat };
        browse.FlatAppearance.BorderColor = Theme.Border;
        browse.Click += (_, _) =>
        {
            using var d = new FolderBrowserDialog { SelectedPath = _dirBox.Text, UseDescriptionForTitle = true, Description = Loc.T("Chọn thư mục cài đặt", "Choose the install folder") };
            if (d.ShowDialog(this) == DialogResult.OK) _dirBox.Text = Path.Combine(d.SelectedPath, d.SelectedPath.EndsWith("Microads Office", StringComparison.OrdinalIgnoreCase) ? "" : "Microads Office").TrimEnd('\\');
        };
        dirRow.Controls.AddRange([_dirBox, browse]);
        CheckBox Chk(string name, string text, bool val) => new() { Name = name, Text = text, Checked = val, Dock = DockStyle.Top, Height = Dpi.Scale(30), Font = new Font("Segoe UI", 10f) };
        _desktop = Chk("chkDesktop", Loc.T("Tạo lối tắt Word, Excel, PDF trên màn hình Desktop", "Create Word, Excel and PDF shortcuts on the Desktop"), _opts.DesktopShortcuts);
        _assoc = Chk("chkAssoc", Loc.T("Mở tệp .docx, .xlsx, .csv, .pdf bằng Microads Office", "Open .docx, .xlsx, .csv and .pdf files with Microads Office"), _opts.FileAssociations);
        _launch = Chk("chkLaunch", Loc.T("Mở Microads Office sau khi cài", "Start Microads Office after installing"), true);
        _autoUpdate = Chk("chkAutoUpdate", Loc.T("Thông báo khi có phiên bản mới (kiểm tra microads.vn, chỉ gửi số phiên bản)", "Notify me about new versions (checks microads.vn, sends only the version number)"), _opts.AutoUpdate);
        var privacyText = Loc.T("Bỏ chọn thì ứng dụng không tự kết nối Internet. Xem chính sách quyền riêng tư", "When unchecked the apps never connect to the Internet on their own. See the privacy policy");
        var privacyLink = Loc.T("chính sách quyền riêng tư", "privacy policy");
        var privacy = new LinkLabel
        {
            Name = "lnkSetupPrivacy", Text = privacyText, Dock = DockStyle.Top, Height = Dpi.Scale(24), Padding = new Padding(Dpi.Scale(18), 0, 0, 0),
            Font = new Font("Segoe UI", 9f), LinkColor = BrandArt.Red, LinkArea = new LinkArea(privacyText.Length - privacyLink.Length, privacyLink.Length)
        };
        privacy.LinkClicked += (_, _) => { try { Process.Start(new ProcessStartInfo(Product.PrivacyUrl) { UseShellExecute = true }); } catch { } };
        var note = Para(Loc.T("Cài cho người dùng hiện tại — không cần quyền quản trị. Dung lượng khoảng 70 MB.",
            "Installs for the current user — no administrator rights needed. About 70 MB."), 44);
        _pages.Controls.Add(note);
        _pages.Controls.Add(privacy);
        _pages.Controls.Add(_autoUpdate);
        _pages.Controls.Add(_launch);
        _pages.Controls.Add(_assoc);
        _pages.Controls.Add(_desktop);
        _pages.Controls.Add(new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(12) });
        _pages.Controls.Add(dirRow);
        _pages.Controls.Add(lbl);
        _pages.Controls.Add(Title(Loc.T("Tùy chọn cài đặt", "Install options")));
    }

    private bool _launchAfter = true;

    private void BuildProgressPage()
    {
        _status = new Label { Name = "lblSetupStatus", Dock = DockStyle.Top, Height = Dpi.Scale(30), Font = new Font("Segoe UI", 10f), Text = Loc.T("Đang chuẩn bị...", "Preparing...") };
        _bar = new ProgressBar { Name = "progressSetup", Dock = DockStyle.Top, Height = Dpi.Scale(18) };
        _pages.Controls.Add(_status);
        _pages.Controls.Add(_bar);
        _pages.Controls.Add(new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(20) });
        _pages.Controls.Add(Title(Loc.T("Đang cài đặt...", "Installing...")));
    }

    private async Task RunInstall()
    {
        var o = _opts;
        try
        {
            var progress = new Progress<(int, string)>(p => { if (_bar != null && !_bar.IsDisposed) { _bar.Value = p.Item1; _status!.Text = p.Item2; } });
            var ip = (IProgress<(int, string)>)progress;
            _installedExe = await Task.Run(() => Installer.Install(o, (pct, s) => ip.Report((pct, s))));
            ExitCode = 0;
            await Task.Delay(300);
            ShowStep(3);
        }
        catch (Exception ex)
        {
            Installer.LogError(ex);
            Msg.Error(this, Loc.T("Không cài đặt được: ", "Installation failed: ") + ex.Message);
            ShowStep(1);
        }
    }

    private void BuildFinishPage()
    {
        var langLbl = new Label { Text = Loc.T("Ngôn ngữ hiển thị:", "Display language:"), Dock = DockStyle.Top, Height = Dpi.Scale(26), Font = Theme.UiFontBold };
        var flags = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(48), WrapContents = false, Name = "finishLanguage" };
        foreach (var l in new[] { Loc.Vietnamese, Loc.English })
        {
            var b = new Button
            {
                Name = "btnFinishLang_" + l, Text = "  " + Flags.LanguageName(l), Image = Flags.Get(l, Dpi.Scale(30), Dpi.Scale(20)), ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText, Size = new Size(Dpi.Scale(150), Dpi.Scale(40)), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                BackColor = l == _lang ? Theme.Hover : Color.White, Font = new Font("Segoe UI", 10f), Margin = new Padding(0, 0, Dpi.Scale(10), 0)
            };
            b.FlatAppearance.BorderColor = l == _lang ? BrandArt.Red : Theme.Border;
            b.FlatAppearance.BorderSize = l == _lang ? 2 : 1;
            b.Click += (_, _) => { _lang = l; Loc.Lang = l; SaveLanguage(); ShowStep(3); };
            flags.Controls.Add(b);
        }
        var openLbl = new Label { Text = Loc.T("Mở ứng dụng:", "Open an app:"), Dock = DockStyle.Top, Height = Dpi.Scale(30), Font = Theme.UiFontBold, Padding = new Padding(0, Dpi.Scale(8), 0, 0) };
        var apps = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(64), WrapContents = false };
        foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf })
        {
            var kind = k;
            var b = new Button
            {
                Name = "btnOpen" + k, Text = " " + Product.AppShortName(k), Image = BrandArt.RenderAppIcon(Dpi.Scale(32), k), TextImageRelation = TextImageRelation.ImageBeforeText,
                Size = new Size(Dpi.Scale(140), Dpi.Scale(52)), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Font = new Font("Segoe UI Semibold", 10.5f), Margin = new Padding(0, 0, Dpi.Scale(10), 0)
            };
            b.FlatAppearance.BorderColor = Theme.Border;
            b.Click += (_, _) => { StartInstalled(kind); _launchAfter = false; };
            apps.Controls.Add(b);
        }
        _pages.Controls.Add(apps);
        _pages.Controls.Add(openLbl);
        _pages.Controls.Add(flags);
        _pages.Controls.Add(langLbl);
        _pages.Controls.Add(Para(Loc.T("Microads Office đã được cài đặt. Các lối tắt Word, Excel, PDF có trong menu Start" + (_opts.DesktopShortcuts ? " và trên Desktop." : "."),
            "Microads Office is installed. Word, Excel and PDF shortcuts are in the Start menu" + (_opts.DesktopShortcuts ? " and on the Desktop." : ".")), 52));
        _pages.Controls.Add(Title(Loc.T("Cài đặt hoàn tất ✓", "Setup complete ✓")));
    }

    private void SaveLanguage()
    {
        var prev = MicroadsOffice.Core.AppSettings.BaseDirectory;
        MicroadsOffice.Core.AppSettings.BaseDirectory = _opts.Directory;
        try { OfficeConfig.Update(c => c.Language = _lang); }
        finally { MicroadsOffice.Core.AppSettings.BaseDirectory = prev; }
    }

    private void StartInstalled(AppKind k)
    {
        if (_installedExe == null) return;
        try { Process.Start(new ProcessStartInfo(_installedExe, Installer.ArgsFor(k)) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(_installedExe)! }); } catch { }
    }

    private void Finish()
    {
        if (_launchAfter) StartInstalled(AppKind.Hub);
        Close();
    }

    /// <summary>Big flag card for the language page.</summary>
    private sealed class LanguageCardButton : Button
    {
        private readonly string _l;
        private readonly bool _sel;
        public LanguageCardButton(string l, bool selected)
        {
            _l = l; _sel = selected;
            Name = "btnSetupLang_" + l;
            AccessibleName = Flags.LanguageName(l);
            Size = new Size(Dpi.Scale(190), Dpi.Scale(160));
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Color.White);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(1, 1, Width - 3, Height - 3);
            using (var p = BrandArt.RoundRect(r, Dpi.Scale(10)))
            {
                using var b = new SolidBrush(_sel ? Color.FromArgb(253, 236, 235) : Color.White);
                g.FillPath(b, p);
                using var pen = new Pen(_sel ? BrandArt.Red : Theme.Border, _sel ? Dpi.Scale(2) : 1);
                g.DrawPath(pen, p);
            }
            int fw = Dpi.Scale(108), fh = Dpi.Scale(72);
            g.DrawImage(Flags.Get(_l, fw, fh), (Width - fw) / 2, Dpi.Scale(22));
            TextRenderer.DrawText(g, Flags.LanguageName(_l), new Font("Segoe UI Semibold", 12.5f), new Rectangle(0, Dpi.Scale(104), Width, Dpi.Scale(32)), Theme.Ink, TextFormatFlags.NoPrefix | TextFormatFlags.HorizontalCenter);
            if (_sel) g.DrawImage(Icons.Get(Icons.Check, Dpi.Scale(20), BrandArt.Red), Width - Dpi.Scale(30), Dpi.Scale(10));
        }
    }
}
