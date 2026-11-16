using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Editor;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.UI;

/// <summary>Horizontal ruler (cm) with draggable first-line, hanging/left and right indent markers.</summary>
public sealed class RulerControl : Control
{
    private readonly WordEditor _ed;
    private int _drag = -1; // 0 first line, 1 hanging, 2 left (both), 3 right
    private float _dragX;
    private ParaProps? _start;

    public RulerControl(WordEditor ed)
    {
        _ed = ed;
        Name = "ruler";
        Height = Dpi.Scale(24);
        Dock = DockStyle.Top;
        DoubleBuffered = true;
        BackColor = Color.FromArgb(232, 234, 237);
        AccessibleName = Loc.T("Thước kẻ", "Ruler");
        ed.SelectionChanged += (_, _) => Invalidate();
        ed.ZoomChanged += (_, _) => Invalidate();
        ed.Scroll += (_, _) => Invalidate();
        ed.Resize += (_, _) => Invalidate();
        ed.DocumentChanged += (_, _) => Invalidate();
    }

    private float PageLeft => _ed.PageLeftPx;
    private float K => _ed.PxPerPoint;

    private (float Left, float Right) ContentEdges()
    {
        var ps = _ed.Document.Page;
        var cell = _ed.CurrentCell;
        if (cell?.Row?.Table is Table t)
        {
            var pl = _ed.LayoutResult.LineOf(_ed.Caret);
            if (pl != null && !pl.Clip.IsEmpty) return (PageLeft + (pl.Clip.Left + DocLayout.CellPad) * K, PageLeft + (pl.Clip.Right - DocLayout.CellPad) * K);
        }
        return (PageLeft + ps.MarginLeft * K, PageLeft + (ps.Width - ps.MarginRight) * K);
    }

    private PointF[] Markers()
    {
        var pp = _ed.CurrentParaProps;
        var (l, r) = ContentEdges();
        return
        [
            new PointF(l + (pp.IndentLeft + pp.FirstLine) * K, 0),
            new PointF(l + pp.IndentLeft * K, 1),
            new PointF(l + pp.IndentLeft * K, 2),
            new PointF(r - pp.IndentRight * K, 3)
        ];
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var ps = _ed.Document.Page;
        float pl = PageLeft, pr = PageLeft + ps.Width * K;
        int top = Dpi.Scale(4), h = Height - Dpi.Scale(8);
        using (var gray = new SolidBrush(Color.FromArgb(210, 212, 216))) g.FillRectangle(gray, pl, top, pr - pl, h);
        var (cl, cr) = ContentEdges();
        g.FillRectangle(Brushes.White, cl, top, cr - cl, h);
        using var pen = new Pen(Color.FromArgb(110, 110, 115));
        using var f = new Font("Segoe UI", 6.5f);
        float cm = 72f / 2.54f * K;
        for (int i = -30; i < 60; i++)
        {
            float x = cl + i * cm / 2;
            if (x < pl || x > pr) continue;
            if (i % 2 == 0)
            {
                if (i != 0) TextRenderer.DrawText(g, Math.Abs(i / 2).ToString(), f, new Rectangle((int)x - 10, top, 20, h), Color.FromArgb(80, 80, 85), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            else g.DrawLine(pen, x, top + h / 2 - 2, x, top + h / 2 + 2);
        }
        var m = Markers();
        using var mb = new SolidBrush(Color.FromArgb(245, 245, 247));
        using var mp = new Pen(Color.FromArgb(90, 90, 95));
        float s = Dpi.Scale(5);
        void Tri(float x, bool down, float y)
        {
            var pts = down ? new[] { new PointF(x - s, y), new PointF(x + s, y), new PointF(x, y + s) } : new[] { new PointF(x - s, y), new PointF(x + s, y), new PointF(x, y - s) };
            g.FillPolygon(mb, pts); g.DrawPolygon(mp, pts);
        }
        Tri(m[0].X, true, top - 1);
        Tri(m[1].X, false, top + h - s + 1);
        g.FillRectangle(mb, m[2].X - s, top + h - s + 1, 2 * s, s); g.DrawRectangle(mp, m[2].X - s, top + h - s + 1, 2 * s, s);
        Tri(m[3].X, false, top + h + 1);
    }

    private int HitMarker(Point p)
    {
        var m = Markers();
        int top = Dpi.Scale(4), h = Height - Dpi.Scale(8), tol = Dpi.Scale(6);
        if (Math.Abs(p.X - m[0].X) <= tol && p.Y < top + h / 2) return 0;
        if (Math.Abs(p.X - m[2].X) <= tol && p.Y > top + h - Dpi.Scale(5)) return 2;
        if (Math.Abs(p.X - m[1].X) <= tol && p.Y >= top + h / 2) return 1;
        if (Math.Abs(p.X - m[3].X) <= tol && p.Y >= top + h / 2) return 3;
        return -1;
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        _drag = HitMarker(e.Location);
        _dragX = e.X;
        _start = _ed.CurrentParaProps;
        Capture = _drag >= 0;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        Cursor = HitMarker(e.Location) >= 0 || _drag >= 0 ? Cursors.SizeWE : Cursors.Default;
        if (_drag >= 0) Invalidate();
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (_drag < 0 || _start == null) return;
        float dpt = (e.X - _dragX) / K;
        // snap to 0.25 cm
        float snap = 72f / 2.54f / 4f;
        dpt = (float)Math.Round(dpt / snap) * snap;
        var s = _start;
        int d = _drag;
        _drag = -1;
        Capture = false;
        if (Math.Abs(dpt) < 0.01f) return;
        _ed.ApplyPara(p => d switch
        {
            0 => p with { FirstLine = s.FirstLine + dpt },
            1 => p with { IndentLeft = Math.Max(0, s.IndentLeft + dpt), FirstLine = s.FirstLine - dpt },
            2 => p with { IndentLeft = Math.Max(0, s.IndentLeft + dpt) },
            _ => p with { IndentRight = Math.Max(0, s.IndentRight - dpt) }
        });
        _ed.Focus();
    }

    /// <summary>Applies an indent drag programmatically (used by tests).</summary>
    public void SimulateDrag(int marker, float points)
    {
        var s = _ed.CurrentParaProps;
        _ed.ApplyPara(p => marker switch
        {
            0 => p with { FirstLine = s.FirstLine + points },
            1 => p with { IndentLeft = Math.Max(0, s.IndentLeft + points), FirstLine = s.FirstLine - points },
            2 => p with { IndentLeft = Math.Max(0, s.IndentLeft + points) },
            _ => p with { IndentRight = Math.Max(0, s.IndentRight - points) }
        });
    }
}

/// <summary>Zoom slider for the status bar (10% – 500%, 100% in the middle).</summary>
public sealed class ZoomSlider : Control
{
    private float _value = 1f;
    private bool _drag;
    public event EventHandler? ValueChanged;

    public ZoomSlider()
    {
        Name = "zoomSlider";
        Size = new Size(Dpi.Scale(120), Dpi.Scale(20));
        DoubleBuffered = true;
        AccessibleName = Loc.T("Thanh trượt thu phóng", "Zoom slider");
        AccessibleRole = AccessibleRole.Slider;
    }

    public float Value { get => _value; set { _value = Math.Clamp(value, 0.1f, 5f); Invalidate(); } }

    private float ToPos(float z) => z <= 1 ? (z - 0.1f) / 0.9f * 0.5f : 0.5f + (z - 1) / 4f * 0.5f;
    private float FromPos(float p) => p <= 0.5f ? 0.1f + p / 0.5f * 0.9f : 1 + (p - 0.5f) / 0.5f * 4f;

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        int pad = Dpi.Scale(6);
        float w = Width - 2 * pad, y = Height / 2f;
        using var p = new Pen(Color.FromArgb(120, 120, 125));
        g.DrawLine(p, pad, y, pad + w, y);
        g.DrawLine(p, pad + w / 2, y - 4, pad + w / 2, y + 4);
        float x = pad + ToPos(_value) * w;
        using var b = new SolidBrush(Theme.Ink);
        g.FillRectangle(b, x - 2, y - 6, 4, 12);
    }

    private void SetFromMouse(int mx)
    {
        int pad = Dpi.Scale(6);
        float pos = Math.Clamp((mx - pad) / (float)(Width - 2 * pad), 0, 1);
        var z = FromPos(pos);
        if (Math.Abs(z - 1) < 0.06f) z = 1;
        Value = (float)Math.Round(z, 2);
        ValueChanged?.Invoke(this, EventArgs.Empty);
    }

    protected override void OnMouseDown(MouseEventArgs e) { _drag = true; Capture = true; SetFromMouse(e.X); }
    protected override void OnMouseMove(MouseEventArgs e) { if (_drag) SetFromMouse(e.X); }
    protected override void OnMouseUp(MouseEventArgs e) { _drag = false; Capture = false; }
}

/// <summary>Quick styles gallery for the Home tab.</summary>
public sealed class StyleGallery : Panel
{
    private readonly List<RibbonButton> _tiles = new();
    private readonly Func<Document> _doc;
    private int _first;
    private const int VisibleCount = 4;
    public event Action<string>? StylePicked;
    public string? Current { get; set; }

    public StyleGallery(Func<Document> doc)
    {
        _doc = doc;
        Name = "styleGallery";
        DoubleBuffered = true;
        BackColor = Theme.RibbonBack;
        int tw = Dpi.Scale(74), th = Dpi.Scale(70);
        Size = new Size(tw * VisibleCount + Dpi.Scale(20), Dpi.Scale(72));
        foreach (var id in Document.GalleryStyles)
        {
            var sid = id;
            var b = new StyleTile(this, sid) { Name = "style_" + sid, Size = new Size(tw - 2, th), ToolTipText = StyleName(sid) };
            b.Click += (_, _) => StylePicked?.Invoke(sid);
            _tiles.Add(b);
            Controls.Add(b);
        }
        var up = Arrow("btnStylesUp", "▲", () => { _first = Math.Max(0, _first - VisibleCount); Arrange(); });
        var down = Arrow("btnStylesDown", "▼", () => { _first = Math.Min(Math.Max(0, _tiles.Count - VisibleCount), _first + VisibleCount); Arrange(); });
        var more = Arrow("btnStylesMore", "▾", () =>
        {
            var menu = new ContextMenuStrip { Renderer = new RibbonRenderer(false), Name = "menuStyles" };
            foreach (var id in Document.GalleryStyles) { var sid = id; menu.Items.Add(new ToolStripMenuItem(StyleName(sid), null, (_, _) => StylePicked?.Invoke(sid)) { Name = "miStyle_" + sid, Checked = Current == sid }); }
            menu.Show(this, new Point(0, Height));
        });
        up.Location = new Point(tw * VisibleCount, 0); down.Location = new Point(tw * VisibleCount, Dpi.Scale(24)); more.Location = new Point(tw * VisibleCount, Dpi.Scale(48));
        Controls.AddRange([up, down, more]);
        Arrange();
    }

    private string StyleName(string id) => _doc().Styles.TryGetValue(id, out var s) ? s.Name : id;

    private Button Arrow(string name, string text, Action a)
    {
        var b = new RibbonButton(text, null, RibbonButtonSize.Icon) { Name = name, Size = new Size(Dpi.Scale(18), Dpi.Scale(24)), Text = "", ToolTipText = name == "btnStylesMore" ? Loc.T("Tất cả kiểu", "All styles") : "" };
        b.Paint += (_, e) => TextRenderer.DrawText(e.Graphics, text, Theme.SmallFont, b.ClientRectangle, Theme.SubtleInk, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        b.Click += (_, _) => a();
        return b;
    }

    private void Arrange()
    {
        int tw = Dpi.Scale(74);
        for (int i = 0; i < _tiles.Count; i++)
        {
            bool vis = i >= _first && i < _first + VisibleCount;
            _tiles[i].Visible = vis;
            if (vis) _tiles[i].Location = new Point((i - _first) * tw, 1);
        }
    }

    public void RefreshState(string? current)
    {
        Current = current;
        foreach (var t in _tiles) t.Checked = t.Name == "style_" + current;
        int idx = _tiles.FindIndex(t => t.Checked);
        if (idx >= 0 && (idx < _first || idx >= _first + VisibleCount)) { _first = idx / VisibleCount * VisibleCount; Arrange(); }
        foreach (var t in _tiles) t.Invalidate();
    }

    private sealed class StyleTile : RibbonButton
    {
        private readonly StyleGallery _g;
        private readonly string _id;
        public StyleTile(StyleGallery g, string id) : base("", null, RibbonButtonSize.Icon) { _g = g; _id = id; AccessibleName = g.StyleName(id); }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            var doc = _g._doc();
            if (!doc.Styles.TryGetValue(_id, out var s)) return;
            var cp = s.Char;
            using var border = new Pen(Checked ? Theme.Accent : Theme.Border);
            g.DrawRectangle(border, 1, 1, Width - 3, Height - 3);
            float size = Math.Clamp(cp.Size * 0.85f, 8, 15);
            using var f = new Font(cp.FontName, size, (cp.Bold ? FontStyle.Bold : 0) | (cp.Italic ? FontStyle.Italic : 0), GraphicsUnit.Point);
            TextRenderer.DrawText(g, "AaBbCc", f, new Rectangle(2, 4, Width - 4, Height - Dpi.Scale(26)), cp.Color ?? Color.Black,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(g, s.Name, Theme.SmallFont, new Rectangle(2, Height - Dpi.Scale(22), Width - 4, Dpi.Scale(18)), Theme.SubtleInk,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.EndEllipsis);
        }
    }
}

/// <summary>Drop-down grid for quick table insertion (up to 10 × 8).</summary>
public sealed class TableGridDropDown : ToolStripDropDown
{
    public TableGridDropDown(Action<int, int> onPick, Action onDialog)
    {
        Name = "tableGrid";
        var panel = new GridPanel(onPick, () => { Close(); onDialog(); }, () => Close());
        var host = new ToolStripControlHost(panel) { AutoSize = false, Size = panel.Size, Margin = Padding.Empty, Padding = Padding.Empty };
        Items.Add(host);
    }

    private sealed class GridPanel : Panel
    {
        private int _r, _c;
        private readonly Action<int, int> _pick;
        private readonly Action _dialog, _close;
        private const int Cols = 10, Rows = 8;
        private int Cell => Dpi.Scale(18);

        public GridPanel(Action<int, int> pick, Action dialog, Action close)
        {
            _pick = pick; _dialog = dialog; _close = close;
            DoubleBuffered = true;
            BackColor = Color.White;
            Size = new Size(Cols * Cell + Dpi.Scale(16), Rows * Cell + Dpi.Scale(64));
            var link = new LinkLabel { Text = Loc.T("Chèn bảng...", "Insert Table..."), Name = "lnkInsertTableDialog", AutoSize = true, Location = new Point(Dpi.Scale(8), Rows * Cell + Dpi.Scale(38)), LinkColor = Theme.AccentDark };
            link.LinkClicked += (_, _) => _dialog();
            Controls.Add(link);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int c = (e.X - Dpi.Scale(8)) / Cell + 1, r = (e.Y - Dpi.Scale(28)) / Cell + 1;
            c = Math.Clamp(c, 0, Cols); r = Math.Clamp(r, 0, Rows);
            if (e.Y < Dpi.Scale(28) || e.Y > Dpi.Scale(28) + Rows * Cell) { c = 0; r = 0; }
            if (c != _c || r != _r) { _c = c; _r = r; Invalidate(); }
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (_r > 0 && _c > 0) { _close(); _pick(_r, _c); }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var title = _r > 0 ? Loc.F("Bảng {0} × {1}", "{0} × {1} Table", _c, _r) : Loc.T("Chèn bảng", "Insert Table");
            TextRenderer.DrawText(g, title, Theme.UiFontBold, new Point(Dpi.Scale(8), Dpi.Scale(6)), Theme.Ink);
            using var pen = new Pen(Color.FromArgb(170, 170, 175));
            using var hot = new SolidBrush(Theme.Pressed);
            using var hotPen = new Pen(Theme.AccentDark);
            for (int r = 0; r < Rows; r++)
                for (int c = 0; c < Cols; c++)
                {
                    var rect = new Rectangle(Dpi.Scale(8) + c * Cell, Dpi.Scale(28) + r * Cell, Cell - 3, Cell - 3);
                    bool h = r < _r && c < _c;
                    if (h) g.FillRectangle(hot, rect);
                    g.DrawRectangle(h ? hotPen : pen, rect);
                }
        }
    }
}
