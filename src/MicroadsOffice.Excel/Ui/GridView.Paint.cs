using System.Drawing.Drawing2D;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

public sealed partial class GridView
{
    private CfEvaluator? _cf;
    private int _cfVersion = -1;
    private Sheet? _cfSheet;
    private readonly Dictionary<string, Image> _pictureCache = new();
    internal static readonly Color[] RefColors = [Color.FromArgb(0, 112, 192), Color.FromArgb(192, 0, 0), Color.FromArgb(112, 48, 160), Color.FromArgb(0, 128, 0), Color.FromArgb(204, 102, 0), Color.FromArgb(0, 128, 128)];

    protected override void OnPaint(PaintEventArgs e)
    {
        if (_doc == null) return;
        var g = e.Graphics;
        var s = Sheet;
        if (_rows.Count == 0 || _cols.Count == 0) ComputeLayout();
        g.Clear(Color.White);
        if (_cf == null || _cfVersion != Book.Version || _cfSheet != s) { _cf = new CfEvaluator(s); _cfVersion = Book.Version; _cfSheet = s; }
        var rows = _rows.Select(r => new CellPainter.Line(r.Index, r.Pos, r.Size)).ToList();
        var cols = _cols.Select(c => new CellPainter.Line(c.Index, c.Pos, c.Size)).ToList();
        var area = CellArea;
        var st = g.Save();
        g.SetClip(area);
        CellPainter.Paint(g, s, rows, cols, Scale, (float)s.Zoom, s.ShowGridlines, true, ShowFormulas, _cf, area);
        PaintAdornments(g);
        PaintSelection(g);
        PaintObjects(g);
        g.Restore(st);
        PaintHeaders(g);
        // frozen pane separators
        using (var fp = new Pen(Color.FromArgb(160, 163, 168)))
        {
            if (s.FreezeRows > 0) g.DrawLine(fp, 0, _frozenBottom - 1, Width - BarW, _frozenBottom - 1);
            if (s.FreezeCols > 0) g.DrawLine(fp, _frozenRight - 1, 0, _frozenRight - 1, Height - BarH);
        }
        // corner between scroll bars
        using (var cb = new SolidBrush(SystemColors.Control)) g.FillRectangle(cb, Width - BarW, Height - BarH, BarW, BarH);
    }

    private void PaintHeaders(Graphics g)
    {
        var s = Sheet;
        if (!s.ShowHeadings) return;
        int hh = ColHeaderH, hw = RowHeaderW;
        using var back = new SolidBrush(Color.FromArgb(245, 245, 247));
        using var selBack = new SolidBrush(Theme.Blend(Theme.Accent, Color.White, 0.82f));
        using var fullBack = new SolidBrush(Theme.Blend(Theme.Accent, Color.White, 0.55f));
        using var line = new Pen(Color.FromArgb(213, 215, 219));
        using var accent = new Pen(Theme.Accent, Math.Max(2, Dpi.Scale(2)));
        var font = CellPainter.GetFont("Segoe UI", 9f * (float)s.Zoom, FontStyle.Regular);
        var sels = AllSelections.ToList();
        bool ColSel(int c) => sels.Any(r => c >= r.C1 && c <= r.C2);
        bool RowSel(int r) => sels.Any(x => r >= x.R1 && r <= x.R2);
        bool WholeCol(int c) => sels.Any(r => c >= r.C1 && c <= r.C2 && r.R1 == 0 && r.R2 == Addr.MaxRows - 1);
        bool WholeRow(int r) => sels.Any(x => r >= x.R1 && r <= x.R2 && x.C1 == 0 && x.C2 == Addr.MaxCols - 1);
        g.FillRectangle(back, 0, 0, Width - BarW, hh);
        g.FillRectangle(back, 0, 0, hw, Height - BarH);
        foreach (var c in _cols)
        {
            var r = new Rectangle(c.Pos, 0, c.Size, hh);
            bool sel = ColSel(c.Index);
            if (sel) g.FillRectangle(WholeCol(c.Index) ? fullBack : selBack, r);
            g.DrawLine(line, r.Right - 1, 0, r.Right - 1, hh - 1);
            TextRenderer.DrawText(g, Addr.ColName(c.Index), font, r, sel ? Theme.AccentDark : Color.FromArgb(70, 72, 76), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            if (sel) g.DrawLine(accent, r.Left, hh - 1, r.Right - 1, hh - 1);
            if (s.HiddenCols.Contains(c.Index + 1)) g.DrawLine(new Pen(Theme.AccentDark, 2), r.Right - 1, 2, r.Right - 1, hh - 2);
        }
        foreach (var rw in _rows)
        {
            var r = new Rectangle(0, rw.Pos, hw, rw.Size);
            bool sel = RowSel(rw.Index);
            if (sel) g.FillRectangle(WholeRow(rw.Index) ? fullBack : selBack, r);
            g.DrawLine(line, 0, r.Bottom - 1, hw - 1, r.Bottom - 1);
            TextRenderer.DrawText(g, (rw.Index + 1).ToString(), font, r, sel ? Theme.AccentDark : Color.FromArgb(70, 72, 76), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            if (sel) g.DrawLine(accent, hw - 1, r.Top, hw - 1, r.Bottom - 1);
            if (s.HiddenRows.Contains(rw.Index + 1)) g.DrawLine(new Pen(Theme.AccentDark, 2), 2, r.Bottom - 1, hw - 2, r.Bottom - 1);
        }
        g.DrawLine(line, 0, hh - 1, Width - BarW, hh - 1);
        g.DrawLine(line, hw - 1, 0, hw - 1, Height - BarH);
        // select-all triangle
        using var tri = new SolidBrush(Color.FromArgb(180, 182, 186));
        int t = Math.Min(hh, hw) - Dpi.Scale(6);
        g.FillPolygon(tri, [new Point(hw - 3, hh - 3 - t), new Point(hw - 3, hh - 3), new Point(hw - 3 - t, hh - 3)]);
    }

    private void PaintAdornments(Graphics g)
    {
        var s = Sheet;
        // note markers
        using var red = new SolidBrush(Color.FromArgb(217, 37, 37));
        foreach (var k in s.Notes.Keys)
        {
            var (r, c) = Sheet.FromKey(k);
            if (CellRect(r, c) is not Rectangle cr) continue;
            int t = Dpi.Scale(6);
            g.FillPolygon(red, [new Point(cr.Right - t - 1, cr.Top), new Point(cr.Right - 1, cr.Top), new Point(cr.Right - 1, cr.Top + t)]);
        }
        // autofilter buttons
        if (s.Filter is AutoFilter f)
        {
            for (int c = f.Range.C1; c <= f.Range.C2; c++)
            {
                if (FilterButtonRect(f.Range.R1, c) is not Rectangle b) continue;
                bool active = f.Allowed.ContainsKey(c) || f.Conditions.ContainsKey(c);
                g.FillRectangle(Brushes.White, b);
                using (var p = new Pen(Color.FromArgb(160, 160, 165))) g.DrawRectangle(p, b.X, b.Y, b.Width - 1, b.Height - 1);
                var ink = active ? Theme.AccentDark : Color.FromArgb(80, 80, 85);
                using var ib = new SolidBrush(ink);
                int cx = b.X + b.Width / 2, cy = b.Y + b.Height / 2;
                if (active)
                    g.FillPolygon(ib, [new Point(cx - 4, cy - 4), new Point(cx + 4, cy - 4), new Point(cx + 1, cy), new Point(cx + 1, cy + 4), new Point(cx - 1, cy + 3), new Point(cx - 1, cy)]);
                else g.FillPolygon(ib, [new Point(cx - 4, cy - 2), new Point(cx + 4, cy - 2), new Point(cx, cy + 3)]);
            }
        }
        // copy marquee
        if (CopyMarquee is CellRange cm && RangeRect(cm) is Rectangle mr)
        {
            using var p = new Pen(Theme.Accent, Dpi.Scale(2)) { DashStyle = DashStyle.Dash };
            g.DrawRectangle(p, mr.X, mr.Y, mr.Width - 2, mr.Height - 2);
        }
        // formula references while editing
        if (IsEditing && _editText.StartsWith('='))
        {
            int i = 0;
            foreach (var (rg, sheetName) in EditRefs())
            {
                if (sheetName != null && !string.Equals(sheetName, s.Name, StringComparison.OrdinalIgnoreCase)) { i++; continue; }
                if (RangeRect(rg) is Rectangle rr)
                {
                    var col = RefColors[i % RefColors.Length];
                    using var b = new SolidBrush(Color.FromArgb(28, col));
                    g.FillRectangle(b, rr);
                    using var p = new Pen(col, Dpi.Scale(2));
                    g.DrawRectangle(p, rr.X, rr.Y, rr.Width - 1, rr.Height - 1);
                }
                i++;
            }
        }
    }

    private void PaintSelection(Graphics g)
    {
        if (SelectedObject != null) return;
        var sel = Selection;
        // spill range of the active cell's dynamic-array formula: thin blue frame (like Excel)
        var spillFrom = Sheet.SpillSource(Active.Row, Active.Col);
        var spillArea = spillFrom is var (sr, sc, _) ? Sheet.SpillArea(sr, sc) : Sheet.SpillArea(Active.Row, Active.Col);
        if (spillArea is CellRange sa && RangeRect(sa) is Rectangle spr)
        {
            using var sp = new Pen(Color.FromArgb(68, 114, 196), Math.Max(1, Dpi.Scale(1)));
            g.DrawRectangle(sp, spr.X, spr.Y, spr.Width - 1, spr.Height - 1);
        }
        using var tint = new SolidBrush(Color.FromArgb(40, Theme.Accent));
        foreach (var x in ExtraRanges) if (RangeRect(x) is Rectangle er) g.FillRectangle(tint, er);
        if (RangeRect(sel) is Rectangle r)
        {
            if (!sel.IsSingle && Sheet.MergeAt(sel.R1, sel.C1) != sel)
            {
                g.FillRectangle(tint, r);
                if (CellRect(Active.Row, Active.Col) is Rectangle ar && !IsEditing) g.FillRectangle(Brushes.White, ar.X, ar.Y, ar.Width - 1, ar.Height - 1);
                // redraw active cell content on top of the white box
                if (CellRect(Active.Row, Active.Col) is Rectangle ar2 && !IsEditing)
                {
                    var rows = _rows.Where(x => x.Index >= Active.Row && x.Index <= (Sheet.MergeAt(Active.Row, Active.Col)?.R2 ?? Active.Row)).Select(x => new CellPainter.Line(x.Index, x.Pos, x.Size)).ToList();
                    var cols = _cols.Where(x => x.Index >= Active.Col && x.Index <= (Sheet.MergeAt(Active.Row, Active.Col)?.C2 ?? Active.Col)).Select(x => new CellPainter.Line(x.Index, x.Pos, x.Size)).ToList();
                    var st = g.Save();
                    g.SetClip(ar2);
                    CellPainter.Paint(g, Sheet, rows, cols, Scale, (float)Sheet.Zoom, Sheet.ShowGridlines, true, ShowFormulas, _cf, ar2);
                    g.Restore(st);
                }
            }
            using var p = new Pen(Theme.Accent, Dpi.Scale(2));
            g.DrawRectangle(p, r.X, r.Y, r.Width - 1, r.Height - 1);
            if (FillHandleRect() is Rectangle fh && !IsEditing)
            {
                g.FillRectangle(Brushes.White, Rectangle.Inflate(fh, 1, 1));
                using var b = new SolidBrush(Theme.Accent);
                g.FillRectangle(b, fh);
            }
        }
        if (_fillTarget is CellRange ft && RangeRect(ft) is Rectangle fr)
        {
            using var p = new Pen(Color.FromArgb(120, 120, 125)) { DashStyle = DashStyle.Dash };
            g.DrawRectangle(p, fr.X, fr.Y, fr.Width - 1, fr.Height - 1);
        }
        if (ValidationArrowRect() is Rectangle va)
        {
            g.FillRectangle(Brushes.White, va);
            using (var p = new Pen(Color.FromArgb(160, 160, 165))) g.DrawRectangle(p, va.X, va.Y, va.Width - 1, va.Height - 1);
            int cx = va.X + va.Width / 2, cy = va.Y + va.Height / 2;
            g.FillPolygon(Brushes.DimGray, [new Point(cx - 4, cy - 2), new Point(cx + 4, cy - 2), new Point(cx, cy + 3)]);
        }
    }

    private void PaintObjects(Graphics g)
    {
        var s = Sheet;
        foreach (var pic in s.Pictures)
        {
            var r = SheetToScreen(pic.Bounds);
            if (!r.IntersectsWith(ClientRectangle)) continue;
            var img = PictureImage(pic);
            if (img != null) g.DrawImage(img, r);
            if (ReferenceEquals(SelectedObject, pic)) DrawHandles(g, r);
        }
        foreach (var ch in s.Charts)
        {
            var r = SheetToScreen(ch.Bounds);
            if (!r.IntersectsWith(ClientRectangle)) continue;
            ChartRenderer.Draw(g, r, s, ch, Scale);
            if (ReferenceEquals(SelectedObject, ch)) DrawHandles(g, r);
        }
    }

    internal Image? PictureImage(Picture pic)
    {
        var key = pic.Name + ":" + pic.Data.Length + ":" + pic.Data.GetHashCode();
        if (_pictureCache.TryGetValue(key, out var img)) return img;
        try { img = Image.FromStream(new MemoryStream(pic.Data)); _pictureCache[key] = img; return img; }
        catch { return null; }
    }

    private static void DrawHandles(Graphics g, RectangleF r)
    {
        using var p = new Pen(Theme.Accent, 1.5f);
        g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height);
        foreach (var h in HandleRects(r))
        {
            g.FillEllipse(Brushes.White, h);
            g.DrawEllipse(p, h);
        }
    }

    internal static RectangleF[] HandleRects(RectangleF r)
    {
        float s = Dpi.Scale(8);
        float[] xs = [r.Left, r.Left + r.Width / 2, r.Right], ys = [r.Top, r.Top + r.Height / 2, r.Bottom];
        var list = new List<RectangleF>();
        // order: 0 TL, 1 T, 2 TR, 3 R, 4 BR, 5 B, 6 BL, 7 L
        (int X, int Y)[] idx = [(0, 0), (1, 0), (2, 0), (2, 1), (2, 2), (1, 2), (0, 2), (0, 1)];
        foreach (var (x, y) in idx) list.Add(new RectangleF(xs[x] - s / 2, ys[y] - s / 2, s, s));
        return list.ToArray();
    }

    private Hit? HitObject(Point p)
    {
        var s = Sheet;
        if (SelectedObject != null)
        {
            var b = SelectedObject is Chart sc ? sc.Bounds : SelectedObject is Picture sp ? sp.Bounds : RectangleF.Empty;
            var r = SheetToScreen(b);
            var handles = HandleRects(r);
            for (int i = 0; i < handles.Length; i++) if (RectangleF.Inflate(handles[i], 2, 2).Contains(p)) return new Hit(HitArea.ObjectHandle, -1, -1, i, SelectedObject);
        }
        for (int i = s.Charts.Count - 1; i >= 0; i--)
            if (SheetToScreen(s.Charts[i].Bounds).Contains(p) && p.X > RowHeaderW && p.Y > ColHeaderH) return new Hit(HitArea.Object, -1, -1, -1, s.Charts[i]);
        for (int i = s.Pictures.Count - 1; i >= 0; i--)
            if (SheetToScreen(s.Pictures[i].Bounds).Contains(p) && p.X > RowHeaderW && p.Y > ColHeaderH) return new Hit(HitArea.Object, -1, -1, -1, s.Pictures[i]);
        return null;
    }
}
