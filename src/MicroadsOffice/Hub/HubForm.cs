using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice.Hub;

/// <summary>Microads Office home: launch Word / Excel / PDF, open any document, recent files of every app.</summary>
public sealed class HubForm : Form, IOfficeAppWindow
{
    private readonly FlowLayoutPanel _recent;

    public HubForm()
    {
        Name = "hubForm";
        Text = "Microads Office";
        Font = Theme.UiFont;
        AutoScaleMode = AutoScaleMode.None;
        BackColor = Color.White;
        StartPosition = FormStartPosition.CenterScreen;
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1366, 728);
        ClientSize = new Size(Math.Min(Dpi.Scale(980), wa.Width - 40), Math.Min(Dpi.Scale(640), wa.Height - 40));
        MinimumSize = new Size(Dpi.Scale(760), Dpi.Scale(540));
        Icon = OfficeShell.AppIcon;
        AllowDrop = true;

        var header = new ToolStrip
        {
            Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, Renderer = new RibbonRenderer(true), BackColor = Theme.RibbonBack,
            Padding = new Padding(Dpi.Scale(8), Dpi.Scale(4), Dpi.Scale(8), Dpi.Scale(4)), Name = "hubHeader", ImageScalingSize = new Size(Dpi.Scale(16), Dpi.Scale(16))
        };
        var about = new ToolStripButton(Loc.T("Giới thiệu", "About")) { Name = "btnAbout", Alignment = ToolStripItemAlignment.Right };
        about.Click += (_, _) => OfficeShell.ShowAbout(this);
        var settings = new ToolStripButton(Loc.T("Cài đặt", "Settings")) { Name = "btnSettings", Alignment = ToolStripItemAlignment.Right };
        settings.Click += (_, _) => OfficeShell.ShowSettings(this);
        var title = new ToolStripLabel("Microads Office") { Font = new Font("Segoe UI Semibold", 11f), ForeColor = Theme.Ink };
        header.Items.Add(new ToolStripLabel(BrandArt.RenderSuiteIcon(Dpi.Scale(22))) { ImageScaling = ToolStripItemImageScaling.None });
        header.Items.Add(title);
        foreach (var it in OfficeShell.CreateHeaderItems(this)) header.Items.Add(it);
        header.Items.Add(settings);
        header.Items.Add(about);

        var body = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Name = "hubBody", BackColor = Color.White };
        var hello = new Label { Text = Loc.T("Bạn muốn làm gì hôm nay?", "What would you like to do today?"), AutoSize = true, Font = new Font("Segoe UI Light", 22f), ForeColor = Theme.Ink };
        var cards = new FlowLayoutPanel { AutoSize = true, WrapContents = true, Name = "hubApps" };
        foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf }) cards.Controls.Add(new AppCard(k, () => Program.Launch(k, [])));
        var open = new Button
        {
            Name = "btnHubOpen", Text = "  " + Loc.T("Mở tệp...", "Open a file..."), Image = Icons.Get(Icons.Open, Dpi.Scale(20), Color.White), TextImageRelation = TextImageRelation.ImageBeforeText,
            AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = BrandArt.Red, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 10.5f), Cursor = Cursors.Hand,
            Padding = new Padding(Dpi.Scale(10), Dpi.Scale(4), Dpi.Scale(10), Dpi.Scale(4))
        };
        open.FlatAppearance.BorderSize = 0;
        open.Click += (_, _) => BrowseAndOpen();
        var recentTitle = new Label { Text = Loc.T("Tài liệu gần đây", "Recent documents"), AutoSize = true, Font = new Font("Segoe UI Semibold", 12f), ForeColor = Theme.Ink };
        _recent = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Name = "hubRecent" };
        var promo = new PromoStrip();
        body.Controls.AddRange([hello, cards, open, recentTitle, _recent, promo]);
        void DoLayout()
        {
            int x = Dpi.Scale(40), y = Dpi.Scale(22), w = body.ClientSize.Width - x * 2;
            hello.Location = new Point(x, y); y += hello.Height + Dpi.Scale(16);
            cards.MaximumSize = new Size(w + Dpi.Scale(12), 0);
            cards.Location = new Point(x - Dpi.Scale(6), y); y += cards.Height + Dpi.Scale(14);
            open.Location = new Point(x, y); y += open.Height + Dpi.Scale(22);
            recentTitle.Location = new Point(x, y); y += recentTitle.Height + Dpi.Scale(6);
            _recent.Location = new Point(x, y); y += _recent.Height + Dpi.Scale(16);
            promo.Width = w;
            promo.Location = new Point(x, Math.Max(y, body.ClientSize.Height - promo.Height - Dpi.Scale(14)));
        }
        body.Resize += (_, _) => DoLayout();
        Activated += (_, _) => { FillRecent(); DoLayout(); };
        Controls.Add(body);
        Controls.Add(header);
        FillRecent();
        DoLayout();

        DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        DragDrop += (_, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] f) OpenFiles(f, false); };
    }

    public IReadOnlyList<string> OpenDocumentPaths => [];

    public void OpenFiles(IReadOnlyList<string> paths, bool fromShell)
    {
        foreach (var g in paths.Where(File.Exists).GroupBy(p => Product.KindForFile(p)))
        {
            if (g.Key == null)
            {
                Msg.Warn(this, Loc.T("Microads Office chưa hỗ trợ loại tệp này: ", "This file type is not supported: ") + string.Join(", ", g.Select(Path.GetFileName)));
                continue;
            }
            Program.Launch(g.Key.Value, g);
        }
    }

    private void BrowseAndOpen()
    {
        using var d = new OpenFileDialog
        {
            Title = Loc.T("Mở tài liệu", "Open document"), Multiselect = true,
            Filter = Loc.T("Tất cả tài liệu hỗ trợ", "All supported documents") + "|*.docx;*.docm;*.dotx;*.rtf;*.txt;*.xlsx;*.xlsm;*.xltx;*.csv;*.tsv;*.pdf|" +
                     "Word (*.docx;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.rtf;*.txt|Excel (*.xlsx;*.csv)|*.xlsx;*.xlsm;*.xltx;*.csv;*.tsv|PDF (*.pdf)|*.pdf"
        };
        if (d.ShowDialog(this) == DialogResult.OK) OpenFiles(d.FileNames, false);
    }

    private void FillRecent()
    {
        var cfg = OfficeConfig.Reload();
        var items = cfg.Recent.SelectMany(kv => kv.Value.Select((p, i) => (p, i))).GroupBy(x => x.p, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First()).OrderBy(x => x.i).Select(x => x.p).Where(File.Exists).Take(8).ToList();
        _recent.SuspendLayout();
        foreach (Control c in _recent.Controls.Cast<Control>().ToList()) c.Dispose();
        int n = 0;
        foreach (var p in items)
        {
            var path = p;
            var kind = Product.KindForFile(path) ?? AppKind.Pdf;
            var b = new Button
            {
                Name = "hubRecent" + n++, Text = "  " + Path.GetFileName(path) + "    —    " + Path.GetDirectoryName(path), Image = BrandArt.RenderAppIcon(Dpi.Scale(24), kind),
                TextImageRelation = TextImageRelation.ImageBeforeText, AutoSize = true, FlatStyle = FlatStyle.Flat, TextAlign = ContentAlignment.MiddleLeft, Cursor = Cursors.Hand,
                ImageAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 0, 0, Dpi.Scale(2))
            };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Theme.Hover;
            b.Click += (_, _) => Program.Launch(kind, [path]);
            _recent.Controls.Add(b);
        }
        if (n == 0) _recent.Controls.Add(new Label { Text = Loc.T("(chưa có)", "(none yet)"), AutoSize = true, ForeColor = Theme.SubtleInk });
        _recent.ResumeLayout();
    }

    /// <summary>Large tile for one application.</summary>
    private sealed class AppCard : Button
    {
        private readonly AppKind _k;
        private bool _hover;

        public AppCard(AppKind k, Action onClick)
        {
            _k = k;
            Name = "hubApp" + k;
            AccessibleName = Product.AppName(k);
            Size = new Size(Dpi.Scale(282), Dpi.Scale(120));
            Margin = new Padding(Dpi.Scale(6));
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            Click += (_, _) => onClick();
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Color.White);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var (acc, dark) = BrandArt.AccentOf(_k);
            var r = new Rectangle(1, 1, Width - 3, Height - 3);
            using (var p = BrandArt.RoundRect(r, Dpi.Scale(10)))
            {
                using var b = new SolidBrush(_hover ? Theme.Blend(acc, Color.White, 0.92f) : Color.FromArgb(250, 250, 252));
                g.FillPath(b, p);
                using var pen = new Pen(_hover ? acc : Theme.Border, _hover ? 2 : 1);
                g.DrawPath(pen, p);
            }
            using (var bar = new SolidBrush(acc)) g.FillRectangle(bar, Dpi.Scale(14), Height - Dpi.Scale(8), Width - Dpi.Scale(28), Dpi.Scale(3));
            using (var ic = BrandArt.RenderAppIcon(Dpi.Scale(64), _k)) g.DrawImage(ic, Dpi.Scale(16), Dpi.Scale(24));
            TextRenderer.DrawText(g, Product.AppShortName(_k), new Font("Segoe UI Semibold", 16f), new Point(Dpi.Scale(92), Dpi.Scale(26)), dark, TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(g, Product.AppDescription(_k), new Font("Segoe UI", 9.5f), new Rectangle(Dpi.Scale(94), Dpi.Scale(62), Width - Dpi.Scale(100), Dpi.Scale(40)), Theme.SubtleInk, TextFormatFlags.NoPrefix | TextFormatFlags.WordBreak);
        }
    }
}
