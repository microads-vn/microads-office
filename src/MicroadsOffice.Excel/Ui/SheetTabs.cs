using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Sheet tabs at the bottom (click to switch, double-click to rename, right-click menu, + to add, drag to reorder).</summary>
public sealed class SheetTabs : Control
{
    private Workbook? _wb;
    private readonly List<(int Index, Rectangle Rect)> _tabs = new();
    private int _first;
    private Rectangle _add, _left, _right;
    private int _dragFrom = -1;

    public event EventHandler<int>? SheetSelected;
    public event EventHandler? AddClicked;
    public event EventHandler<int>? RenameRequested;
    public event EventHandler<(int Index, Point Location)>? MenuRequested;
    public event EventHandler<(int From, int To)>? MoveRequested;

    public SheetTabs()
    {
        Name = "sheetTabs";
        Height = Dpi.Scale(30);
        Dock = DockStyle.Bottom;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        AccessibleName = Loc.T("Thẻ trang tính", "Sheet tabs");
    }

    public Workbook? Book { get => _wb; set { _wb = value; _first = 0; Invalidate(); } }

    public Rectangle? TabRect(int index) { var t = _tabs.FirstOrDefault(x => x.Index == index); return t.Rect.IsEmpty ? null : t.Rect; }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Theme.TabStripBack);
        _tabs.Clear();
        if (_wb == null) return;
        int h = Height, bw = Dpi.Scale(24);
        _left = new Rectangle(Dpi.Scale(4), 0, bw, h);
        _right = new Rectangle(_left.Right, 0, bw, h);
        foreach (var (r, glyph) in new[] { (_left, Icons.Prev), (_right, Icons.Next) })
            g.DrawImage(Icons.Get(glyph, Dpi.Scale(12), Theme.SubtleInk), r.X + (r.Width - Dpi.Scale(12)) / 2, (h - Dpi.Scale(12)) / 2);
        int x = _right.Right + Dpi.Scale(6);
        var font = Theme.UiFont;
        var bold = Theme.UiFontBold;
        int maxX = Width - Dpi.Scale(40);
        for (int i = _first; i < _wb.Sheets.Count; i++)
        {
            var s = _wb.Sheets[i];
            if (s.Hidden) continue;
            bool active = i == _wb.ActiveSheet;
            int w = TextRenderer.MeasureText(s.Name, active ? bold : font).Width + Dpi.Scale(24);
            if (x + w > maxX && _tabs.Count > 0) break;
            var r = new Rectangle(x, 0, w, h - Dpi.Scale(3));
            using (var b = new SolidBrush(active ? Color.White : Theme.TabStripBack)) g.FillRectangle(b, r);
            if (active)
            {
                using var ab = new SolidBrush(Theme.Accent);
                g.FillRectangle(ab, r.X, r.Bottom - Dpi.Scale(3), r.Width, Dpi.Scale(3));
            }
            if (s.TabColor != 0)
            {
                using var tb = new SolidBrush(Color.FromArgb(255, Color.FromArgb(s.TabColor)));
                g.FillRectangle(tb, r.X + 2, active ? r.Bottom - Dpi.Scale(6) : r.Bottom - Dpi.Scale(4), r.Width - 4, Dpi.Scale(3));
            }
            using (var p = new Pen(Theme.Border)) g.DrawLine(p, r.Right, Dpi.Scale(6), r.Right, h - Dpi.Scale(8));
            TextRenderer.DrawText(g, s.Name, active ? bold : font, r, active ? Theme.AccentDark : Theme.Ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            _tabs.Add((i, r));
            x += w;
        }
        _add = new Rectangle(x + Dpi.Scale(6), Dpi.Scale(3), Dpi.Scale(24), h - Dpi.Scale(8));
        g.DrawImage(XlIcons.Get("addSheet", Dpi.Scale(16)), _add.X + Dpi.Scale(4), _add.Y + (_add.Height - Dpi.Scale(16)) / 2);
    }

    public int IndexAt(Point p) => _tabs.FirstOrDefault(t => t.Rect.Contains(p)).Rect.IsEmpty ? -1 : _tabs.First(t => t.Rect.Contains(p)).Index;

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (_wb == null) return;
        if (_add.Contains(e.Location)) { AddClicked?.Invoke(this, EventArgs.Empty); return; }
        if (_left.Contains(e.Location)) { _first = Math.Max(0, _first - 1); Invalidate(); return; }
        if (_right.Contains(e.Location)) { _first = Math.Min(_wb.Sheets.Count - 1, _first + 1); Invalidate(); return; }
        int i = IndexAt(e.Location);
        if (i < 0) return;
        if (e.Button == MouseButtons.Right) { SheetSelected?.Invoke(this, i); MenuRequested?.Invoke(this, (i, e.Location)); return; }
        _dragFrom = i;
        SheetSelected?.Invoke(this, i);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (_dragFrom >= 0)
        {
            int to = IndexAt(e.Location);
            if (to >= 0 && to != _dragFrom) MoveRequested?.Invoke(this, (_dragFrom, to));
        }
        _dragFrom = -1;
    }

    protected override void OnMouseDoubleClick(MouseEventArgs e)
    {
        base.OnMouseDoubleClick(e);
        int i = IndexAt(e.Location);
        if (i >= 0) RenameRequested?.Invoke(this, i);
        else if (!_add.Contains(e.Location)) AddClicked?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Makes sure the active tab is visible.</summary>
    public void Reveal()
    {
        if (_wb == null) return;
        if (_wb.ActiveSheet < _first) _first = _wb.ActiveSheet;
        Invalidate();
        Update();
        if (_tabs.Count > 0 && _wb.ActiveSheet > _tabs[^1].Index) { _first = Math.Min(_wb.ActiveSheet, _first + (_wb.ActiveSheet - _tabs[^1].Index)); Invalidate(); }
    }
}
