using System.Diagnostics;
using System.Drawing.Drawing2D;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;

namespace MicroadsOffice.UI;

/// <summary>Large clickable card with a flag and the language name.</summary>
internal static class Mnemonics
{
    /// <summary>Promo / company texts may contain '&amp;' (e.g. "Security &amp; safety"): show it literally.</summary>
    public static void Off(Control root)
    {
        foreach (Control c in root.Controls)
        {
            if (c is Label l) l.UseMnemonic = false;
            Off(c);
        }
    }
}

internal sealed class LanguageCard : Button
{
    public string Lang { get; }
    private bool _selected, _hover;
    public bool Selected { get => _selected; set { _selected = value; Invalidate(); } }

    public LanguageCard(string lang)
    {
        Lang = lang;
        Name = "lang_" + lang;
        AccessibleName = Flags.LanguageName(lang);
        AccessibleRole = AccessibleRole.PushButton;
        Size = new Size(Dpi.Scale(180), Dpi.Scale(150));
        Cursor = Cursors.Hand;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Color.White;
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
    }

    protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? Color.White);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var r = new Rectangle(1, 1, Width - 3, Height - 3);
        using (var path = BrandArt.RoundRect(r, Dpi.Scale(10)))
        {
            using var back = new SolidBrush(_selected ? Theme.Hover : _hover ? Color.FromArgb(248, 248, 250) : Color.White);
            g.FillPath(back, path);
            using var pen = new Pen(_selected ? Theme.Accent : Theme.Border, _selected ? Dpi.Scale(2) : 1);
            g.DrawPath(pen, path);
        }
        int fw = Dpi.Scale(96), fh = Dpi.Scale(64);
        var flag = Flags.Get(Lang, fw, fh);
        g.DrawImage(flag, (Width - fw) / 2, Dpi.Scale(22));
        TextRenderer.DrawText(g, Flags.LanguageName(Lang), new Font("Segoe UI Semibold", 12f), new Rectangle(0, Dpi.Scale(96), Width, Dpi.Scale(30)), Theme.Ink,
            TextFormatFlags.NoPrefix | TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        if (_selected)
            g.DrawImage(Icons.Get(Icons.Check, Dpi.Scale(18), Theme.Accent), Width - Dpi.Scale(28), Dpi.Scale(10));
    }
}

/// <summary>Language picker (shown right after installation / first run, and from Settings).</summary>
public sealed class LanguageDialog : DialogBase
{
    public string Selected { get; private set; } = Loc.Lang;

    public LanguageDialog(bool firstRun) : base("Chọn ngôn ngữ · Choose your language", 460, firstRun ? 300 : 280)
    {
        Name = "languageDialog";
        var title = new Label
        {
            Text = "Chọn ngôn ngữ hiển thị\nChoose the display language",
            Dock = DockStyle.Top, Height = Dpi.Scale(52), Font = new Font("Segoe UI Semibold", 11f), ForeColor = Theme.Ink, TextAlign = ContentAlignment.MiddleCenter
        };
        var row = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(Dpi.Scale(22), Dpi.Scale(8), 0, 0) };
        foreach (var l in new[] { Loc.Vietnamese, Loc.English })
        {
            var card = new LanguageCard(l) { Selected = l == Selected, Margin = new Padding(Dpi.Scale(8)) };
            card.Click += (_, _) =>
            {
                Selected = card.Lang;
                foreach (var c in row.Controls.OfType<LanguageCard>()) c.Selected = c == card;
                if (firstRun) { DialogResult = DialogResult.OK; Close(); }
            };
            row.Controls.Add(card);
        }
        Body.Controls.Add(row);
        Body.Controls.Add(title);
        var hint = new Label
        {
            Text = "Có thể đổi lại bất cứ lúc nào bằng nút lá cờ trên thanh công cụ · You can change it later with the flag button.",
            Dock = DockStyle.Bottom, Height = Dpi.Scale(34), ForeColor = Theme.SubtleInk, Font = Theme.SmallFont, TextAlign = ContentAlignment.MiddleCenter
        };
        Body.Controls.Add(hint);
        if (firstRun) AddButton("Tiếp tục · Continue", "btnLangContinue", DialogResult.OK, true);
        else AddOkCancel(Loc.T("Áp dụng", "Apply"), Loc.T("Hủy", "Cancel"));
    }
}

/// <summary>About box: logo, version, publisher, "always free", services and microads.vn.</summary>
public sealed class AboutOfficeDialog : DialogBase
{
    public AboutOfficeDialog() : base(Loc.T("Giới thiệu ", "About ") + Product.AppName(OfficeShell.App), 560, 520)
    {
        Name = "aboutDialog";
        var p = PromoInfo.Effective;
        var logo = new PictureBox
        {
            Image = BrandArt.RenderWordmark(Dpi.Scale(76), false, OfficeShell.App), SizeMode = PictureBoxSizeMode.Zoom,
            Dock = DockStyle.Top, Height = Dpi.Scale(84), Name = "picLogo"
        };
        var text = new Label
        {
            Name = "lblAbout", Dock = DockStyle.Top, Height = Dpi.Scale(96), Font = new Font("Segoe UI", 10f),
            Text = $"{Product.AppName(OfficeShell.App)} — {Product.AppDescription(OfficeShell.App)}\n" +
                   Loc.F("Phiên bản {0}", "Version {0}", Product.VersionText) + "\n" +
                   p.Company + "\n" + "© " + DateTime.Now.Year + " Microads. " + Loc.T("Bảo lưu mọi quyền.", "All rights reserved.")
        };
        var free = new Label
        {
            Name = "lblFree", Text = Product.PublishedBy + " — " + Product.AlwaysFree.ToLower(Loc.IsVi ? new System.Globalization.CultureInfo("vi-VN") : System.Globalization.CultureInfo.InvariantCulture),
            Dock = DockStyle.Top, Height = Dpi.Scale(30), ForeColor = Color.White, BackColor = Theme.AccentGreen, TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI Semibold", 10f)
        };
        var services = new Label
        {
            Name = "lblServices", Dock = DockStyle.Top, Height = Dpi.Scale(130), Padding = new Padding(0, Dpi.Scale(8), 0, 0),
            Text = p.Tagline + ":\n" + string.Join("\n", p.Services.Items.Select(s => "   •  " + s)), ForeColor = Theme.Ink
        };
        var link = new LinkLabel
        {
            Name = "lnkWebsite", Text = p.WebsiteLabel, Dock = DockStyle.Top, Height = Dpi.Scale(30), Font = new Font("Segoe UI Semibold", 12f),
            LinkColor = Theme.AccentDark, TextAlign = ContentAlignment.MiddleLeft
        };
        link.LinkClicked += (_, _) => OfficeShell.OpenWebsite(p.Website);
        var contact = new Label
        {
            Name = "lblContact", Dock = DockStyle.Top, Height = Dpi.Scale(22), ForeColor = Theme.SubtleInk,
            Text = string.Join("   ", new[] { string.IsNullOrWhiteSpace(p.Phone) ? null : "☎ " + p.Phone, string.IsNullOrWhiteSpace(p.Email) ? null : "✉ " + p.Email }.Where(x => x != null))
        };
        var oss = new Label
        {
            Dock = DockStyle.Bottom, Height = Dpi.Scale(22), ForeColor = Theme.SubtleInk, Font = Theme.SmallFont,
            Text = Loc.T("Thành phần mã nguồn mở: PDFium (BSD-3-Clause), .NET (MIT).", "Open-source components: PDFium (BSD-3-Clause), .NET (MIT).")
        };
        Body.Controls.Add(oss);
        Body.Controls.Add(contact);
        Body.Controls.Add(link);
        Body.Controls.Add(services);
        Body.Controls.Add(free);
        Body.Controls.Add(text);
        Body.Controls.Add(logo);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
        AddButton(Loc.T("Kiểm tra cập nhật", "Check for updates"), "btnAboutUpdate", DialogResult.None, false, (_, _) => OfficeShell.CheckForUpdatesInteractive(this));
        Mnemonics.Off(this);
        AddButton(p.CallToAction.ToString(), "btnAboutVisit", DialogResult.None, false, (_, _) => OfficeShell.OpenWebsite(p.Website));
    }
}

/// <summary>
/// Promotional notice shown when a document is opened from a folder. It can be closed immediately
/// (X button, "Close" button, Esc) — there is no countdown.
/// </summary>
public sealed class PromoPopup : Form
{
    public PromoPopup()
    {
        var p = PromoInfo.Effective;
        Name = "promoPopup";
        Text = "Microads Office";
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.White;
        Font = Theme.UiFont;
        KeyPreview = true;
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(Dpi.Scale(480), Dpi.Scale(370));
        DoubleBuffered = true;
        AccessibleName = Loc.T("Thông báo Microads Office", "Microads Office notice");

        var close = new Button
        {
            Name = "btnPromoX", Text = "✕", FlatStyle = FlatStyle.Flat, Size = new Size(Dpi.Scale(36), Dpi.Scale(32)),
            Location = new Point(ClientSize.Width - Dpi.Scale(40), Dpi.Scale(8)), Font = new Font("Segoe UI", 11f), ForeColor = Theme.SubtleInk,
            AccessibleName = Loc.T("Tắt thông báo", "Close notice"), Cursor = Cursors.Hand, TabStop = false
        };
        close.FlatAppearance.BorderSize = 0;
        close.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 17, 35);
        close.MouseEnter += (_, _) => close.ForeColor = Color.White;
        close.MouseLeave += (_, _) => close.ForeColor = Theme.SubtleInk;
        close.Click += (_, _) => Close();
        Controls.Add(close);

        var logo = new PictureBox
        {
            Image = BrandArt.RenderWordmark(Dpi.Scale(64), false, null), SizeMode = PictureBoxSizeMode.Zoom,
            Bounds = new Rectangle(Dpi.Scale(24), Dpi.Scale(22), Dpi.Scale(300), Dpi.Scale(64)), Name = "picPromoLogo"
        };
        Controls.Add(logo);

        var free = new Label
        {
            Name = "lblPromoFree", Text = Product.AlwaysFree.ToUpperInvariant(), AutoSize = false, TextAlign = ContentAlignment.MiddleCenter,
            BackColor = Theme.AccentGreen, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 8.5f),
            Bounds = new Rectangle(Dpi.Scale(330), Dpi.Scale(42), Dpi.Scale(100), Dpi.Scale(22))
        };
        Controls.Add(free);

        var head = new Label
        {
            Name = "lblPromoHeadline", Text = p.Headline.ToString(), Font = new Font("Segoe UI Semibold", 10.5f), ForeColor = Theme.Ink,
            Bounds = new Rectangle(Dpi.Scale(24), Dpi.Scale(96), Dpi.Scale(432), Dpi.Scale(44))
        };
        Controls.Add(head);
        var body = new Label
        {
            Name = "lblPromoServices", ForeColor = Theme.Ink,
            Text = p.Tagline + ":\n" + string.Join("\n", p.Services.Items.Select(s => "  ✓  " + s)),
            Bounds = new Rectangle(Dpi.Scale(24), Dpi.Scale(144), Dpi.Scale(432), Dpi.Scale(130))
        };
        Controls.Add(body);

        var visit = new Button
        {
            Name = "btnPromoVisit", Text = p.CallToAction + "  →", FlatStyle = FlatStyle.Flat, BackColor = BrandArt.Red, ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 10f), Bounds = new Rectangle(Dpi.Scale(24), Dpi.Scale(296), Dpi.Scale(250), Dpi.Scale(40)), Cursor = Cursors.Hand
        };
        visit.FlatAppearance.BorderSize = 0;
        visit.Click += (_, _) => { OfficeShell.OpenWebsite(p.Website); Close(); };
        Controls.Add(visit);
        var dismiss = new Button
        {
            Name = "btnPromoClose", Text = Loc.T("Đóng", "Close"), FlatStyle = FlatStyle.Flat, BackColor = Color.White, ForeColor = Theme.Ink,
            Bounds = new Rectangle(Dpi.Scale(290), Dpi.Scale(296), Dpi.Scale(166), Dpi.Scale(40)), Cursor = Cursors.Hand
        };
        dismiss.FlatAppearance.BorderColor = Theme.Border;
        dismiss.Click += (_, _) => Close();
        Controls.Add(dismiss);
        var site = new LinkLabel
        {
            Name = "lnkPromoSite", Text = p.WebsiteLabel, AutoSize = true, LinkColor = Theme.SubtleInk, Location = new Point(Dpi.Scale(24), Dpi.Scale(344)),
            Font = Theme.SmallFont
        };
        site.LinkClicked += (_, _) => OfficeShell.OpenWebsite(p.Website);
        Controls.Add(site);
        Mnemonics.Off(this);
        AcceptButton = dismiss;
        CancelButton = dismiss;
        KeyDown += (_, e) => { if (e.KeyCode == Keys.Escape) Close(); };
    }

    protected override void OnShown(EventArgs e)
    {
        if (Owner != null)
        {
            var o = Owner.Bounds;
            Location = new Point(o.Left + (o.Width - Width) / 2, o.Top + (o.Height - Height) / 2);
        }
        base.OnShown(e);
        Activate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        using var top = new LinearGradientBrush(new Rectangle(0, 0, Width, Dpi.Scale(6)), BrandArt.Red, BrandArt.Green, LinearGradientMode.Horizontal);
        g.FillRectangle(top, 0, 0, Width, Dpi.Scale(6));
        using var pen = new Pen(Theme.Border);
        g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
    }

    protected override CreateParams CreateParams
    {
        get { var cp = base.CreateParams; cp.ClassStyle |= 0x00020000; /* CS_DROPSHADOW */ return cp; }
    }
}

/// <summary>Small non-blocking notification in the bottom-right corner of the window.</summary>
public sealed class Toast : Form
{
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 20000 };

    private Toast(string message, string? action, Action? onAction)
    {
        Name = "toast";
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = Color.White;
        Font = Theme.UiFont;
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(Dpi.Scale(380), Dpi.Scale(action != null ? 108 : 76));
        var icon = new PictureBox { Image = BrandArt.RenderAppIcon(Dpi.Scale(32), OfficeShell.App), SizeMode = PictureBoxSizeMode.AutoSize, Location = new Point(Dpi.Scale(14), Dpi.Scale(16)) };
        var lbl = new Label { Name = "lblToast", Text = message, Bounds = new Rectangle(Dpi.Scale(56), Dpi.Scale(14), Dpi.Scale(282), Dpi.Scale(52)), ForeColor = Theme.Ink };
        var x = new Button { Name = "btnToastClose", Text = "✕", FlatStyle = FlatStyle.Flat, Bounds = new Rectangle(ClientSize.Width - Dpi.Scale(32), Dpi.Scale(4), Dpi.Scale(28), Dpi.Scale(26)), TabStop = false, AccessibleName = Loc.T("Đóng", "Close") };
        x.FlatAppearance.BorderSize = 0;
        x.Click += (_, _) => Close();
        Controls.AddRange([icon, lbl, x]);
        if (action != null)
        {
            var b = new Button
            {
                Name = "btnToastAction", Text = action, FlatStyle = FlatStyle.Flat, BackColor = Theme.Accent, ForeColor = Color.White, AutoSize = true,
                Location = new Point(Dpi.Scale(56), Dpi.Scale(68)), MinimumSize = new Size(Dpi.Scale(110), Dpi.Scale(30)), Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += (_, _) => { Close(); onAction?.Invoke(); };
            Controls.Add(b);
        }
        else _timer.Start();
        _timer.Tick += (_, _) => Close();
        Mnemonics.Off(this);
    }

    protected override bool ShowWithoutActivation => true;

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        using var acc = new SolidBrush(Theme.Accent);
        e.Graphics.FillRectangle(acc, 0, 0, Dpi.Scale(4), Height);
        using var pen = new Pen(Theme.Border);
        e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
    }

    protected override void Dispose(bool disposing) { if (disposing) _timer.Dispose(); base.Dispose(disposing); }

    public static void Show(Form owner, string message, string? action, Action? onAction)
    {
        if (owner.IsDisposed) return;
        void Do()
        {
            foreach (var old in Application.OpenForms.OfType<Toast>().ToList()) old.Close();
            var t = new Toast(message, action, onAction);
            var area = owner.RectangleToScreen(owner.ClientRectangle);
            t.Location = new Point(area.Right - t.Width - Dpi.Scale(16), area.Bottom - t.Height - Dpi.Scale(40));
            t.Show(owner);
        }
        if (owner.InvokeRequired) owner.BeginInvoke(Do); else Do();
    }
}

/// <summary>"Check for updates": shows the installed / latest version, downloads, verifies and installs.</summary>
public sealed class UpdateDialog : DialogBase
{
    private readonly Label _status, _notes;
    private readonly ProgressBar _bar;
    private readonly Button _update, _later, _page;
    private readonly Form? _main;
    private ReleaseInfo? _latest;
    private CancellationTokenSource? _cts;
    private bool _installed;

    public UpdateDialog(Form? main) : base(Loc.T("Kiểm tra cập nhật", "Check for updates"), 500, 330)
    {
        Name = "updateDialog";
        _main = main;
        var head = new Label
        {
            Name = "lblCurrent", Dock = DockStyle.Top, Height = Dpi.Scale(48), Font = new Font("Segoe UI Semibold", 11f),
            Text = Product.SuiteName + "\n" + Loc.F("Phiên bản đang dùng: {0}", "Installed version: {0}", Product.VersionText)
        };
        _status = new Label { Name = "lblUpdateStatus", Dock = DockStyle.Top, Height = Dpi.Scale(48), Text = Loc.T("Đang kiểm tra phiên bản mới tại microads.vn/office ...", "Checking microads.vn/office for a new version ...") };
        _bar = new ProgressBar { Name = "progressUpdate", Dock = DockStyle.Top, Height = Dpi.Scale(16), Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30 };
        _notes = new Label { Name = "lblUpdateNotes", Dock = DockStyle.Fill, ForeColor = Theme.SubtleInk, Padding = new Padding(0, Dpi.Scale(10), 0, 0) };
        Body.Controls.Add(_notes);
        Body.Controls.Add(_bar);
        Body.Controls.Add(_status);
        Body.Controls.Add(head);
        _later = AddButton(Loc.T("Đóng", "Close"), "btnUpdateClose", DialogResult.Cancel);
        _update = AddButton(Loc.T("Cập nhật ngay", "Update now"), "btnUpdateNow", DialogResult.None, true, async (_, _) => await DoUpdate());
        _page = AddButton(Loc.T("Mở trang tải", "Open download page"), "btnUpdatePage", DialogResult.None, false, (_, _) => OfficeShell.OpenWebsite(Product.UpdatePage));
        Mnemonics.Off(this);
        _update.Visible = false;
        CancelButton = _later;
        Shown += async (_, _) => await Check();
        FormClosing += (_, _) => _cts?.Cancel();
    }

    private async Task Check()
    {
        if (StorePackage.IsPackaged)
        {
            // Store edition: Microsoft Store installs updates automatically; offer to open the Store page.
            _ = Task.Run(() => UpdateService.FetchManifestAsync()); // refresh promo content only
            _bar.Style = ProgressBarStyle.Continuous;
            _bar.Value = 100;
            _status.Text = Loc.T("Bản Microsoft Store được cập nhật tự động qua Microsoft Store.", "The Microsoft Store edition is updated automatically by the Microsoft Store.");
            _status.ForeColor = Theme.OkGreen;
            _update.Text = Loc.T("Mở Microsoft Store", "Open Microsoft Store");
            _update.Visible = true;
            _storeMode = true;
            return;
        }
        var m = await Task.Run(() => UpdateService.FetchManifestAsync());
        if (IsDisposed) return;
        _bar.Style = ProgressBarStyle.Continuous;
        _bar.Value = 0;
        if (m != null && m.Latest == null)
        {
            _status.Text = Loc.T("Bạn đang dùng phiên bản mới nhất. ✓", "You are using the latest version. ✓");
            _status.ForeColor = Theme.OkGreen;
            _bar.Value = 100;
            return;
        }
        if (m?.Latest == null)
        {
            _status.Text = Loc.T("Không kết nối được máy chủ cập nhật. Vui lòng kiểm tra Internet rồi thử lại.", "Could not reach the update server. Please check your Internet connection and try again.");
            _status.ForeColor = Theme.WarnAmber;
            return;
        }
        _latest = m.Latest;
        if (UpdateService.IsNewer(m.Latest))
        {
            _status.Text = Loc.F("Đã có phiên bản mới: {0} (phát hành {1}).", "A new version is available: {0} (released {1}).", m.Latest.Version, m.Latest.Date);
            _status.ForeColor = Theme.AccentDark;
            _notes.Text = m.Latest.Notes?.ToString() ?? "";
            _update.Visible = true;
            AcceptButton = _update;
            _update.Focus();
        }
        else
        {
            _status.Text = Loc.T("Bạn đang dùng phiên bản mới nhất. ✓", "You are using the latest version. ✓");
            _status.ForeColor = Theme.OkGreen;
            _bar.Value = 100;
        }
    }

    private bool _storeMode;

    private async Task DoUpdate()
    {
        if (_storeMode) { StorePackage.OpenStorePage(); return; }
        if (_latest == null) return;
        if (_installed) { DialogResult = DialogResult.OK; Close(); if (_main != null) OfficeShell.RestartApp(_main); return; }
        _update.Enabled = false;
        _status.Text = Loc.T("Đang tải bản cập nhật ...", "Downloading the update ...");
        _status.ForeColor = Theme.Ink;
        _cts = new CancellationTokenSource();
        try
        {
            var prog = new Progress<double>(v => { if (!IsDisposed) _bar.Value = (int)(v * 100); });
            var bytes = await UpdateService.DownloadAsync(_latest, prog, _cts.Token);
            _status.Text = Loc.T("Đang cài đặt ...", "Installing ...");
            var res = UpdateService.Apply(bytes, _latest.Version, out var err);
            switch (res)
            {
                case ApplyResult.Applied:
                    OfficeShell.MarkUpdateApplied();
                    _installed = true;
                    _bar.Value = 100;
                    _status.Text = Loc.F("Đã cài phiên bản {0}. Khởi động lại để sử dụng.", "Version {0} installed. Restart to use it.", _latest.Version);
                    _status.ForeColor = Theme.OkGreen;
                    _update.Text = Loc.T("Khởi động lại ngay", "Restart now");
                    _update.Enabled = _main != null;
                    _later.Text = Loc.T("Để sau", "Later");
                    break;
                case ApplyResult.NeedsElevation:
                    _status.Text = Loc.T("Trình cài đặt đang chạy với quyền quản trị để hoàn tất cập nhật.", "The installer is running with administrator rights to finish the update.");
                    break;
                default:
                    _status.Text = Loc.T("Không cài được bản cập nhật: ", "Could not install the update: ") + err;
                    _status.ForeColor = Theme.ErrRed;
                    _update.Enabled = true;
                    break;
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            _status.Text = ex.Message;
            _status.ForeColor = Theme.ErrRed;
            _update.Enabled = true;
        }
    }
}

/// <summary>Settings shared by every app: language (flags), automatic update checks, data folder.</summary>
public sealed class OfficeSettingsDialog : DialogBase
{
    public string SelectedLanguage { get; private set; } = Loc.Lang;

    public OfficeSettingsDialog() : base(Loc.T("Cài đặt Microads Office", "Microads Office settings"), 500, OfficeShell.AiSettingsHandler != null ? 470 : 420)
    {
        Name = "settingsDialog";
        var langTitle = new Label { Text = Loc.T("Ngôn ngữ / Language", "Language / Ngôn ngữ"), Dock = DockStyle.Top, Height = Dpi.Scale(26), Font = Theme.UiFontBold };
        var row = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(170), WrapContents = false, Padding = new Padding(Dpi.Scale(40), 0, 0, 0) };
        foreach (var l in new[] { Loc.Vietnamese, Loc.English })
        {
            var card = new LanguageCard(l) { Selected = l == SelectedLanguage, Margin = new Padding(Dpi.Scale(8)) };
            card.Click += (_, _) =>
            {
                SelectedLanguage = card.Lang;
                foreach (var c in row.Controls.OfType<LanguageCard>()) c.Selected = c == card;
            };
            row.Controls.Add(card);
        }
        var auto = new CheckBox
        {
            Name = "chkAutoUpdate", Text = Loc.T("Thông báo khi có phiên bản mới (kiểm tra microads.vn, chỉ gửi số phiên bản)", "Notify me about new versions (checks microads.vn, sends only the version number)"),
            Checked = OfficeConfig.Current.AutoCheckUpdates, Dock = DockStyle.Top, Height = Dpi.Scale(30)
        };
        var autoInstall = new CheckBox
        {
            Name = "chkAutoInstall", Text = Loc.T("Tự cài bản cập nhật khi Microads yêu cầu cập nhật toàn bộ", "Install updates by itself when Microads asks everyone to update"),
            Checked = OfficeConfig.Current.AutoInstallForcedUpdates, Dock = DockStyle.Top, Height = Dpi.Scale(30)
        };
        auto.CheckedChanged += (_, _) => autoInstall.Enabled = auto.Checked;
        autoInstall.Enabled = auto.Checked;
        var privacy = new LinkLabel
        {
            Name = "lnkPrivacy", Dock = DockStyle.Top, Height = Dpi.Scale(24), Padding = new Padding(Dpi.Scale(18), 0, 0, 0),
            Text = Loc.T("Tắt mục này thì ứng dụng không tự kết nối Internet. Chính sách quyền riêng tư", "When off, the apps never connect to the Internet on their own. Privacy policy"), LinkColor = Theme.AccentDark
        };
        privacy.LinkArea = new LinkArea(privacy.Text.Length - Loc.T("Chính sách quyền riêng tư", "Privacy policy").Length, Loc.T("Chính sách quyền riêng tư", "Privacy policy").Length);
        privacy.LinkClicked += (_, _) => OfficeShell.OpenWebsite(Product.PrivacyUrl);
        var check = new Button { Name = "btnSettingsCheckUpdate", Text = Loc.T("Kiểm tra cập nhật ngay", "Check for updates now"), AutoSize = true, Dock = DockStyle.Top, FlatStyle = FlatStyle.Flat };
        check.FlatAppearance.BorderColor = Theme.Border;
        check.Click += (_, _) => OfficeShell.CheckForUpdatesInteractive(this);
        var data = new LinkLabel
        {
            Name = "lnkDataFolder", Dock = DockStyle.Top, Height = Dpi.Scale(40), Padding = new Padding(0, Dpi.Scale(10), 0, 0),
            Text = Loc.T("Thư mục dữ liệu: ", "Data folder: ") + AppSettings.DataDirectory, LinkColor = Theme.AccentDark
        };
        data.LinkClicked += (_, _) => { try { Directory.CreateDirectory(AppSettings.DataDirectory); Process.Start(new ProcessStartInfo(AppSettings.DataDirectory) { UseShellExecute = true }); } catch { } };
        var ai = new Button
        {
            Name = "btnSettingsAi", Text = Loc.T("Trợ lý AI (ChatGPT, Claude)…", "AI assistant (ChatGPT, Claude)…"), AutoSize = true, Dock = DockStyle.Top, FlatStyle = FlatStyle.Flat,
            Visible = OfficeShell.AiSettingsHandler != null, Margin = new Padding(0, Dpi.Scale(6), 0, 0)
        };
        ai.FlatAppearance.BorderColor = Theme.Border;
        ai.Click += (_, _) => OfficeShell.AiSettingsHandler?.Invoke(this);
        var ver = new Label { Dock = DockStyle.Top, Height = Dpi.Scale(24), ForeColor = Theme.SubtleInk, Text = Product.SuiteName + " " + Product.VersionText + " · " + Product.PublishedBy };
        Body.Controls.Add(ver);
        Body.Controls.Add(data);
        Body.Controls.Add(ai);
        Body.Controls.Add(check);
        Body.Controls.Add(autoInstall);
        Body.Controls.Add(privacy);
        Body.Controls.Add(auto);
        Body.Controls.Add(row);
        Body.Controls.Add(langTitle);
        AddOkCancel(Loc.T("Lưu", "Save"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            var a = auto.Checked;
            var ai2 = autoInstall.Checked;
            OfficeConfig.Update(c => { c.AutoCheckUpdates = a; c.AutoInstallForcedUpdates = ai2; });
            DialogResult = DialogResult.OK;
            Close();
        });
    }
}
