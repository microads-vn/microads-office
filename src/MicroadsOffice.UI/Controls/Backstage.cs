using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.UI;

/// <summary>
/// Full-window "File" view (like Office backstage): a coloured navigation column on the left with
/// pages (Home, New, Open, Info, Print, Export...) and plain commands (Save, Close...).
/// The app decides what each page contains.
/// </summary>
public sealed class Backstage : UserControl
{
    private readonly Panel _nav;
    private readonly Panel _content;
    private readonly FlowLayoutPanel _navItems;
    private readonly FlowLayoutPanel _navBottom;
    private readonly Dictionary<string, (Control Content, NavButton Button)> _pages = new();
    private string? _current;

    /// <summary>Raised when the user leaves the backstage (back arrow / Esc).</summary>
    public event EventHandler? BackRequested;
    public string? CurrentPage => _current;

    public Backstage(AppKind app)
    {
        Name = "backstage";
        Dock = DockStyle.Fill;
        BackColor = Color.White;
        Visible = false;
        _nav = new Panel { Dock = DockStyle.Left, Width = Dpi.Scale(200), BackColor = Theme.AccentDark, Name = "backstageNav" };
        _nav.Paint += (_, e) =>
        {
            using var b = new LinearGradientBrush(_nav.ClientRectangle, Theme.Accent, Theme.AccentDark, LinearGradientMode.Vertical);
            e.Graphics.FillRectangle(b, _nav.ClientRectangle);
        };
        var back = new NavButton("", Icons.Get(Icons.Back, Dpi.Scale(20), Color.White)) { Name = "btnBackstageBack", AccessibleName = Loc.T("Quay lại", "Back"), Height = Dpi.Scale(48) };
        back.Click += (_, _) => BackRequested?.Invoke(this, EventArgs.Empty);
        _navItems = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, BackColor = Color.Transparent };
        _navBottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, BackColor = Color.Transparent };
        _navItems.Controls.Add(back);
        _nav.Controls.Add(_navItems);
        _nav.Controls.Add(_navBottom);
        _content = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Name = "backstageContent", Padding = new Padding(Dpi.Scale(36), Dpi.Scale(24), Dpi.Scale(24), Dpi.Scale(12)) };
        Controls.Add(_content);
        Controls.Add(_nav);
    }

    public void AddPage(string name, string text, string glyph, Control content, bool bottom = false)
    {
        var b = new NavButton(text, Icons.Get(glyph, Dpi.Scale(18), Color.White)) { Name = "bs" + name };
        b.Click += (_, _) => ShowPage(name);
        content.Dock = DockStyle.Fill;
        content.Visible = false;
        _content.Controls.Add(content);
        (bottom ? _navBottom : _navItems).Controls.Add(b);
        _pages[name] = (content, b);
    }

    public NavButton AddCommand(string name, string text, string glyph, Action onClick, bool bottom = false)
    {
        var b = new NavButton(text, Icons.Get(glyph, Dpi.Scale(18), Color.White)) { Name = "bs" + name };
        b.Click += (_, _) => onClick();
        (bottom ? _navBottom : _navItems).Controls.Add(b);
        return b;
    }

    public void AddSeparator(bool bottom = false)
    {
        var sep = new Panel { Height = 1, Width = Dpi.Scale(168), BackColor = Color.FromArgb(90, 255, 255, 255), Margin = new Padding(Dpi.Scale(16), Dpi.Scale(6), 0, Dpi.Scale(6)) };
        (bottom ? _navBottom : _navItems).Controls.Add(sep);
    }

    public void ShowPage(string name)
    {
        if (!_pages.ContainsKey(name)) return;
        foreach (var (key, (content, button)) in _pages)
        {
            content.Visible = key == name;
            button.Selected = key == name;
        }
        _current = name;
        PageShown?.Invoke(this, name);
    }

    /// <summary>Raised whenever a page becomes visible (refresh recent lists etc.).</summary>
    public event EventHandler<string>? PageShown;

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape) { BackRequested?.Invoke(this, EventArgs.Empty); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    /// <summary>Navigation entry of the backstage column.</summary>
    public sealed class NavButton : Button
    {
        private bool _hover, _selected;
        private readonly Image _icon;
        public bool Selected { get => _selected; set { _selected = value; Invalidate(); } }

        public NavButton(string text, Image icon)
        {
            Text = text;
            _icon = icon;
            AccessibleName = text;
            Size = new Size(Dpi.Scale(200), Dpi.Scale(40));
            Margin = Padding.Empty;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            ForeColor = Color.White;
            Font = new Font("Segoe UI", 10.5f);
            Cursor = Cursors.Hand;
            TabStop = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var b = new LinearGradientBrush(Parent!.Parent!.ClientRectangle, Theme.Accent, Theme.AccentDark, LinearGradientMode.Vertical))
            {
                var st = g.Save();
                g.TranslateTransform(-Left - Parent.Left, -Top - Parent.Top);
                g.FillRectangle(b, Parent.Parent.ClientRectangle);
                g.Restore(st);
            }
            if (_selected || _hover)
            {
                using var hb = new SolidBrush(Color.FromArgb(_selected ? 60 : 35, 255, 255, 255));
                g.FillRectangle(hb, ClientRectangle);
            }
            if (_selected) g.FillRectangle(Brushes.White, 0, 0, Dpi.Scale(4), Height);
            int isz = Dpi.Scale(18);
            g.DrawImage(_icon, Dpi.Scale(18), (Height - isz) / 2, isz, isz);
            TextRenderer.DrawText(g, Text, Font, new Rectangle(Dpi.Scale(48), 0, Width - Dpi.Scale(52), Height), Color.White, TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            if (Focused && ShowFocusCues) ControlPaint.DrawFocusRectangle(g, Rectangle.Inflate(ClientRectangle, -2, -2));
        }
    }
}

/// <summary>Template tile for the start screen / New page.</summary>
public sealed record DocTemplate(string Id, string Title, Func<int, int, Bitmap> Thumbnail);

/// <summary>
/// Start screen shared by Word and Excel (Office "Home" page): greeting, template tiles
/// (Blank first), recent documents, Open button and the Microads footer.
/// </summary>
public sealed class StartScreen : UserControl
{
    private readonly AppKind _app;
    private readonly FlowLayoutPanel _templates;
    private readonly Panel _recent;
    private readonly Action<DocTemplate> _onTemplate;
    private readonly Action<string> _onOpenRecent;
    private readonly Action _onBrowse;
    private readonly bool _compactTemplates;

    public StartScreen(AppKind app, IEnumerable<DocTemplate> templates, Action<DocTemplate> onTemplate, Action<string> onOpenRecent, Action onBrowse, bool compactTemplates = true)
    {
        _app = app;
        _onTemplate = onTemplate;
        _onOpenRecent = onOpenRecent;
        _onBrowse = onBrowse;
        _compactTemplates = compactTemplates;
        Name = "startScreen";
        Dock = DockStyle.Fill;
        BackColor = Color.White;
        AutoScroll = true;
        HorizontalScroll.Enabled = false;
        HorizontalScroll.Visible = false;

        var hello = new Label
        {
            Text = Greeting(), AutoSize = true, Font = new Font("Segoe UI Light", 22f), ForeColor = Theme.Ink,
            Location = new Point(0, 0), Name = "lblGreeting"
        };
        var newTitle = new Label { Text = Loc.T("Mới", "New"), AutoSize = true, Font = new Font("Segoe UI Semibold", 12f), ForeColor = Theme.Ink, Name = "lblNew" };
        _templates = new FlowLayoutPanel { AutoSize = true, WrapContents = true, FlowDirection = FlowDirection.LeftToRight, Name = "templateList", MaximumSize = new Size(Dpi.Scale(compactTemplates ? 900 : 760), 0) };
        foreach (var t in templates) _templates.Controls.Add(MakeTile(t));
        var recentTitle = new Label { Text = Loc.T("Gần đây", "Recent"), AutoSize = true, Font = new Font("Segoe UI Semibold", 12f), ForeColor = Theme.Ink, Name = "lblRecent" };
        var browse = new Button
        {
            Text = "  " + Loc.T("Mở tệp khác...", "Open other files..."), Name = "btnStartBrowse", AutoSize = true, FlatStyle = FlatStyle.Flat,
            Image = Icons.Get(Icons.Open, Dpi.Scale(18), Theme.AccentDark), TextImageRelation = TextImageRelation.ImageBeforeText, ForeColor = Theme.AccentDark,
            Font = new Font("Segoe UI Semibold", 10f), Cursor = Cursors.Hand, Padding = new Padding(Dpi.Scale(6), Dpi.Scale(2), Dpi.Scale(6), Dpi.Scale(2))
        };
        browse.FlatAppearance.BorderColor = Theme.Border;
        browse.Click += (_, _) => _onBrowse();
        _recent = new Panel { Name = "recentList", AutoSize = false };
        var footer = new PromoStrip { Name = "promoStrip" };

        Controls.AddRange([hello, newTitle, _templates, recentTitle, browse, _recent, footer]);
        Resize += (_, _) => DoLayout();
        VisibleChanged += (_, _) => { if (Visible) RefreshRecent(); };
        RefreshRecent();
        DoLayout();

        void DoLayout()
        {
            int x = Dpi.Scale(48), y = Dpi.Scale(28), w = Math.Max(Dpi.Scale(160), ClientSize.Width - x * 2 - SystemInformation.VerticalScrollBarWidth);
            hello.Location = new Point(x, y); y += hello.Height + Dpi.Scale(14);
            newTitle.Location = new Point(x, y); y += newTitle.Height + Dpi.Scale(6);
            _templates.MaximumSize = new Size(w, 0);
            _templates.Location = new Point(x - Dpi.Scale(6), y); y += _templates.Height + Dpi.Scale(16);
            recentTitle.Location = new Point(x, y);
            browse.Location = new Point(x + w - browse.Width, y - Dpi.Scale(4));
            y += recentTitle.Height + Dpi.Scale(8);
            _recent.Location = new Point(x, y);
            _recent.Width = w;
            y += _recent.Height + Dpi.Scale(18);
            footer.Width = w;
            footer.Location = new Point(x, Math.Max(y, ClientSize.Height - footer.Height - Dpi.Scale(12)));
        }
    }

    private static string Greeting()
    {
        int h = DateTime.Now.Hour;
        return h < 11 ? Loc.T("Chào buổi sáng", "Good morning") : h < 13 ? Loc.T("Chào buổi trưa", "Good afternoon") : h < 18 ? Loc.T("Chào buổi chiều", "Good afternoon") : Loc.T("Chào buổi tối", "Good evening");
    }

    private Control MakeTile(DocTemplate t)
    {
        int tw = Dpi.Scale(_compactTemplates ? 126 : 150), th = Dpi.Scale(_compactTemplates ? 164 : 120);
        var tile = new Button
        {
            Name = "tpl_" + t.Id, Size = new Size(tw + Dpi.Scale(12), th + Dpi.Scale(40)), FlatStyle = FlatStyle.Flat, BackColor = Color.White,
            Margin = new Padding(Dpi.Scale(6)), Cursor = Cursors.Hand, AccessibleName = t.Title, Text = ""
        };
        tile.FlatAppearance.BorderSize = 0;
        tile.FlatAppearance.MouseOverBackColor = Theme.Hover;
        Bitmap? thumb = null;
        try { thumb = t.Thumbnail(tw, th); } catch { }
        tile.Paint += (_, e) =>
        {
            var g = e.Graphics;
            var r = new Rectangle(Dpi.Scale(6), Dpi.Scale(6), tw, th);
            if (thumb != null) g.DrawImage(thumb, r);
            using var pen = new Pen(Theme.Border);
            g.DrawRectangle(pen, r);
            TextRenderer.DrawText(g, t.Title, Theme.UiFont, new Rectangle(0, r.Bottom + Dpi.Scale(4), tile.Width, Dpi.Scale(30)), Theme.Ink,
                TextFormatFlags.NoPrefix | TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak | TextFormatFlags.EndEllipsis);
        };
        tile.Click += (_, _) => _onTemplate(t);
        return tile;
    }

    public void RefreshRecent()
    {
        _recent.SuspendLayout();
        foreach (Control c in _recent.Controls.Cast<Control>().ToList()) c.Dispose();
        var list = OfficeConfig.Reload() is { } ? OfficeConfig.RecentFor(_app) : [];
        int y = 0, i = 0;
        int rowH = Dpi.Scale(44);
        foreach (var path in list.Take(12))
        {
            var p = path;
            bool exists = File.Exists(p);
            var row = new Button
            {
                Name = "recent" + i++, Location = new Point(0, y), Size = new Size(_recent.Width, rowH), FlatStyle = FlatStyle.Flat, BackColor = Color.White,
                Cursor = Cursors.Hand, AccessibleName = Path.GetFileName(p), Text = "", Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };
            row.FlatAppearance.BorderSize = 0;
            row.FlatAppearance.MouseOverBackColor = Theme.Hover;
            var date = exists ? File.GetLastWriteTime(p).ToString("dd/MM/yyyy HH:mm") : Loc.T("(không tìm thấy)", "(not found)");
            row.Paint += (_, e) =>
            {
                var g = e.Graphics;
                var kind = Product.KindForFile(p) ?? _app;
                using var ico = BrandArt.RenderAppIcon(Dpi.Scale(28), kind);
                g.DrawImage(ico, Dpi.Scale(6), (rowH - Dpi.Scale(28)) / 2);
                TextRenderer.DrawText(g, Path.GetFileName(p), Theme.UiFontBold, new Rectangle(Dpi.Scale(44), Dpi.Scale(4), row.Width - Dpi.Scale(200), Dpi.Scale(20)), exists ? Theme.Ink : Theme.SubtleInk, TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
                TextRenderer.DrawText(g, Path.GetDirectoryName(p) ?? "", Theme.SmallFont, new Rectangle(Dpi.Scale(44), Dpi.Scale(22), row.Width - Dpi.Scale(200), Dpi.Scale(18)), Theme.SubtleInk, TextFormatFlags.NoPrefix | TextFormatFlags.PathEllipsis);
                TextRenderer.DrawText(g, date, Theme.SmallFont, new Rectangle(row.Width - Dpi.Scale(150), 0, Dpi.Scale(144), rowH), Theme.SubtleInk, TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            };
            row.Click += (_, _) => _onOpenRecent(p);
            var menu = new ContextMenuStrip { Renderer = new RibbonRenderer(false) };
            menu.Items.Add(Loc.T("Mở", "Open"), null, (_, _) => _onOpenRecent(p));
            menu.Items.Add(Loc.T("Mở thư mục chứa tệp", "Open file location"), null, (_, _) =>
            {
                try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + p + "\""); } catch { }
            });
            menu.Items.Add(Loc.T("Xóa khỏi danh sách", "Remove from list"), null, (_, _) => { OfficeConfig.RemoveRecent(_app, p); RefreshRecent(); });
            row.ContextMenuStrip = menu;
            _recent.Controls.Add(row);
            y += rowH;
        }
        if (i == 0)
        {
            _recent.Controls.Add(new Label
            {
                Text = Loc.T("Chưa có tài liệu nào được mở gần đây. Bấm \"Mở tệp khác...\" hoặc kéo-thả tệp vào cửa sổ.",
                             "No recent documents yet. Click \"Open other files...\" or drag files into this window."),
                AutoSize = true, ForeColor = Theme.SubtleInk, Location = new Point(Dpi.Scale(4), Dpi.Scale(6)), Name = "lblNoRecent"
            });
            y = Dpi.Scale(36);
        }
        _recent.Height = y;
        _recent.ResumeLayout();
        PerformLayout();
        OnResize(EventArgs.Empty);
    }
}

/// <summary>Thin branded footer: "Published by Microads ... always free · microads.vn".</summary>
public sealed class PromoStrip : Control
{
    public PromoStrip()
    {
        Height = Dpi.Scale(56);
        Cursor = Cursors.Hand;
        DoubleBuffered = true;
        var p = PromoInfo.Effective;
        AccessibleName = p.WebsiteLabel;
        AccessibleRole = AccessibleRole.Link;
        Click += (_, _) => OfficeShell.OpenWebsite(p.Website);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var p = PromoInfo.Effective;
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        using (var path = BrandArt.RoundRect(r, Dpi.Scale(8)))
        {
            using var b = new SolidBrush(Color.FromArgb(248, 249, 251));
            g.FillPath(b, path);
            using var pen = new Pen(Theme.Border);
            g.DrawPath(pen, path);
        }
        BrandArt.DrawMark(g, new RectangleF(Dpi.Scale(14), Dpi.Scale(14), Dpi.Scale(35), Dpi.Scale(28)));
        TextRenderer.DrawText(g, Product.PublishedBy + " · " + Product.AlwaysFree, Theme.UiFontBold, new Rectangle(Dpi.Scale(60), Dpi.Scale(8), Width - Dpi.Scale(200), Dpi.Scale(20)), Theme.Ink, TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        TextRenderer.DrawText(g, p.Tagline + ": " + string.Join(" · ", p.Services.Items), Theme.SmallFont, new Rectangle(Dpi.Scale(60), Dpi.Scale(28), Width - Dpi.Scale(200), Dpi.Scale(20)), Theme.SubtleInk, TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        TextRenderer.DrawText(g, p.WebsiteLabel + "  →", new Font("Segoe UI Semibold", 11f), new Rectangle(Width - Dpi.Scale(170), 0, Dpi.Scale(160), Height), BrandArt.Red, TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
    }
}
