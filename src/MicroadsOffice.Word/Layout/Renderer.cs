using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.Layout;

public sealed class RenderOptions
{
    public DocPos? SelStart, SelEnd;
    public bool ShowMarks;
    public bool Print;
    public bool ShowGridlines = true;
    /// <summary>Misspelled ranges per paragraph (offset, length).</summary>
    public Func<Paragraph, IReadOnlyList<(int Start, int Len)>>? Misspelled;
    public Color SelectionColor = Color.FromArgb(110, 66, 133, 244);
    /// <summary>Selected inline image (draws resize handles).</summary>
    public ImageObject? SelectedImage;
}

/// <summary>Draws one laid-out page. The graphics must use points with the page origin at (0,0).</summary>
public static class PageRenderer
{
    /// <summary>
    /// Maps world units to points so that one point covers <paramref name="pxPerPoint"/> device pixels and the page
    /// origin is at device pixel (<paramref name="ox"/>, <paramref name="oy"/>). Fonts sized in points then render at
    /// exactly the size used by the layout engine.
    /// </summary>
    public static void Setup(Graphics g, float ox, float oy, float pxPerPoint)
    {
        g.PageUnit = GraphicsUnit.Point;
        float k = 72f / g.DpiX;
        g.TranslateTransform(ox * k, oy * k);
        g.ScaleTransform(pxPerPoint * k, pxPerPoint * k);
    }

    public static void DrawPage(Graphics g, DocLayout L, int pageIndex, RenderOptions o)
    {
        var doc = L.Doc;
        var ps = doc.Page;
        var page = L.Pages[pageIndex];
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = PixelOffsetMode.Half;

        using (var bg = new SolidBrush(doc.PageColor ?? Color.White)) g.FillRectangle(bg, 0, 0, ps.Width, ps.Height);
        if (!string.IsNullOrWhiteSpace(doc.Watermark)) DrawWatermark(g, ps, doc.Watermark!);

        // paragraph shading / borders (body paragraphs)
        foreach (var grp in page.Lines.Where(l => l.Line.Para.Cell == null && (l.Line.Para.Props.Shading != null || l.Line.Para.Props.Border)).GroupBy(l => l.Line.Para))
        {
            var pp = grp.Key.Props;
            float top = grp.Min(l => l.Y), bottom = grp.Max(l => l.Y + l.Line.Height);
            var r = new RectangleF(ps.MarginLeft + pp.IndentLeft - 3, top - 2, ps.ContentWidth - pp.IndentLeft - pp.IndentRight + 6, bottom - top + 4);
            if (pp.Shading is Color sh) using (var b = new SolidBrush(sh)) g.FillRectangle(b, r);
            if (pp.Border) using (var pen = new Pen(Color.Black, 0.75f)) g.DrawRectangle(pen, r.X, r.Y, r.Width, r.Height);
        }

        // table cells
        foreach (var c in page.Cells)
        {
            if (c.Cell.Shading is Color sh) using (var b = new SolidBrush(sh)) g.FillRectangle(b, c.Rect);
        }
        foreach (var c in page.Cells)
        {
            var r = c.Rect;
            if (c.Table.Borders)
            {
                using var pen = new Pen(c.Table.BorderColor, Math.Max(0.25f, c.Table.BorderWidth));
                g.DrawLine(pen, r.Left, r.Top, r.Left, r.Bottom);
                g.DrawLine(pen, r.Right, r.Top, r.Right, r.Bottom);
                if (c.DrawTop) g.DrawLine(pen, r.Left, r.Top, r.Right, r.Top);
                if (c.DrawBottom) g.DrawLine(pen, r.Left, r.Bottom, r.Right, r.Bottom);
            }
            else if (!o.Print && o.ShowGridlines)
            {
                using var pen = new Pen(Color.FromArgb(150, 120, 160, 220), 0.5f) { DashStyle = DashStyle.Dot };
                g.DrawRectangle(pen, r.X, r.Y, r.Width, r.Height);
            }
        }

        foreach (var pl in page.HeaderLines) DrawLine(g, L, pl, o, pageIndex, header: true);
        foreach (var pl in page.FooterLines) DrawLine(g, L, pl, o, pageIndex, header: true);
        foreach (var pl in page.Lines) DrawLine(g, L, pl, o, pageIndex, header: false);
    }

    private static void DrawWatermark(Graphics g, PageSetup ps, string text)
    {
        var st = g.Save();
        g.TranslateTransform(ps.Width / 2, ps.Height / 2);
        float angle = -(float)(Math.Atan2(ps.Height, ps.Width) * 180 / Math.PI);
        g.RotateTransform(angle);
        float diag = (float)Math.Sqrt(ps.Width * ps.Width + ps.Height * ps.Height);
        using var f0 = new Font("Arial", 100, FontStyle.Bold, GraphicsUnit.Point);
        var size = g.MeasureString(text, f0, PointF.Empty, TextMeasurer.Format);
        float fs = Math.Clamp(100 * diag * 0.62f / Math.Max(1, size.Width), 20, 160);
        using var f = new Font("Arial", fs, FontStyle.Bold, GraphicsUnit.Point);
        var sz = g.MeasureString(text, f, PointF.Empty, TextMeasurer.Format);
        using var b = new SolidBrush(Color.FromArgb(48, 128, 128, 128));
        g.DrawString(text, f, b, -sz.Width / 2, -sz.Height / 2, TextMeasurer.Format);
        g.Restore(st);
    }

    private static void DrawLine(Graphics g, DocLayout L, PlacedLine pl, RenderOptions o, int pageIndex, bool header)
    {
        var line = pl.Line;
        var m = L.M;
        GraphicsState? clipState = null;
        if (!pl.Clip.IsEmpty)
        {
            clipState = g.Save();
            g.SetClip(pl.Clip, CombineMode.Intersect);
        }
        float baseY = pl.Y + line.Baseline;

        // selection
        if (!o.Print && !header && o.SelStart is DocPos s0 && o.SelEnd is DocPos s1 && s0 != s1)
        {
            int pi = L.Doc.IndexOf(line.Para);
            var a = s0 < s1 ? s0 : s1; var b = s0 < s1 ? s1 : s0;
            if (pi >= a.Para && pi <= b.Para)
            {
                int from = pi == a.Para ? Math.Max(a.Offset, line.Start) : line.Start;
                int to = pi == b.Para ? Math.Min(b.Offset, line.End) : line.End;
                bool includesMark = pi < b.Para && line.IsLast;
                if (to > from || includesMark)
                {
                    float x0 = L.XInLine(line, from), x1 = to > from ? L.XInLine(line, to) : x0;
                    if (includesMark) x1 += 5;
                    using var sb = new SolidBrush(o.SelectionColor);
                    g.FillRectangle(sb, pl.X + x0, pl.Y, Math.Max(1, x1 - x0), line.Height);
                }
            }
        }

        // highlights (behind text)
        foreach (var s in line.Segs)
        {
            if (s.Props.Highlight is Color hc && s.Kind != SegKind.Break)
            {
                var mt = m.Metrics(s.Props);
                float top = s.Obj is ImageObject im ? baseY - im.Height : baseY - mt.Asc;
                float bottom = baseY + (s.Obj is ImageObject ? 0 : mt.Desc);
                using var hb = new SolidBrush(hc);
                g.FillRectangle(hb, pl.X + s.X, top, s.W, bottom - top);
            }
        }

        if (line.Label != null)
        {
            var f = m.GetFont(line.LabelProps);
            var asc = m.Metrics(line.LabelProps).Asc;
            using var lb = new SolidBrush(line.LabelProps.Color ?? Color.Black);
            g.DrawString(line.Label, f, lb, pl.X + line.LabelX, baseY - asc, TextMeasurer.Format);
        }

        foreach (var s in line.Segs)
        {
            float x = pl.X + s.X;
            var p = s.Props;
            var color = p.Color ?? (p.Link != null ? Color.FromArgb(5, 99, 193) : Color.Black);
            if (o.Print && color.GetBrightness() > 0.95f && L.Doc.PageColor == null) color = Color.Black;
            float shift = p.VertAlign == VertAlign.Superscript ? -p.Size * 0.33f : p.VertAlign == VertAlign.Subscript ? p.Size * 0.12f : 0;
            switch (s.Kind)
            {
                case SegKind.Text:
                {
                    var f = m.GetFont(p);
                    var asc = m.Metrics(p).Asc;
                    using var br = new SolidBrush(color);
                    var text = p.AllCaps ? s.Text.ToUpperInvariant() : s.Text;
                    if (Math.Abs(s.W - m.Width(s.Text, p)) > 0.5f && text.Length > 1)
                    {
                        // justified text segment that was stretched: draw char by char
                        for (int i = 0; i < text.Length; i++)
                            g.DrawString(text[i].ToString(), f, br, pl.X + s.XAt(m, s.Start + i), baseY - asc + shift, TextMeasurer.Format);
                    }
                    else g.DrawString(text, f, br, x, baseY - asc + shift, TextMeasurer.Format);
                    break;
                }
                case SegKind.Object when s.Obj is ImageObject img:
                {
                    try { g.DrawImage(img.Image, x, baseY - img.Height, img.Width, img.Height); }
                    catch { using var pen = new Pen(Color.Gray); g.DrawRectangle(pen, x, baseY - img.Height, img.Width, img.Height); }
                    if (!o.Print && o.SelectedImage == img)
                    {
                        using var pen = new Pen(Color.FromArgb(66, 133, 244), 1f);
                        var r = new RectangleF(x, baseY - img.Height, img.Width, img.Height);
                        g.DrawRectangle(pen, r.X, r.Y, r.Width, r.Height);
                        foreach (var h in Handles(r)) { g.FillRectangle(Brushes.White, h); g.DrawRectangle(pen, h.X, h.Y, h.Width, h.Height); }
                    }
                    break;
                }
                case SegKind.Object when s.Obj is FieldObject fo:
                {
                    var txt = fo.Kind switch { FieldKind.Page => (pageIndex + 1).ToString(), FieldKind.NumPages => L.Pages.Count.ToString(), _ => DateTime.Now.ToString("dd/MM/yyyy") };
                    var f = m.GetFont(p);
                    using var br = new SolidBrush(color);
                    g.DrawString(txt, f, br, x, baseY - m.Metrics(p).Asc + shift, TextMeasurer.Format);
                    if (!o.Print && o.ShowMarks) using (var fb = new SolidBrush(Color.FromArgb(40, 0, 0, 0))) g.FillRectangle(fb, x, pl.Y, s.W, line.Height);
                    break;
                }
                case SegKind.Tab when s.Leader:
                {
                    using var pen = new Pen(color, 0.8f) { DashStyle = DashStyle.Dot };
                    g.DrawLine(pen, x + 2, baseY - 1, x + s.W - 2, baseY - 1);
                    break;
                }
            }
            // underline / strike
            if (s.Kind is SegKind.Text or SegKind.Space or SegKind.Tab && !(s.Kind == SegKind.Space && s == line.Segs[^1]))
            {
                if (p.Underline != UnderlineStyle.None || p.Link != null && p.Underline == UnderlineStyle.None && false)
                {
                    float uy = baseY + Math.Max(1f, p.Size * 0.11f) + shift;
                    float th = Math.Max(0.5f, p.Size * (p.Underline == UnderlineStyle.Thick ? 0.1f : 0.055f));
                    using var pen = new Pen(color, th);
                    if (p.Underline == UnderlineStyle.Dotted) pen.DashStyle = DashStyle.Dot;
                    if (p.Underline == UnderlineStyle.Wavy) DrawWave(g, color, x, uy, s.W, th);
                    else g.DrawLine(pen, x, uy, x + s.W, uy);
                    if (p.Underline == UnderlineStyle.Double) g.DrawLine(pen, x, uy + th * 2.2f, x + s.W, uy + th * 2.2f);
                }
                if (p.Strike)
                {
                    float sy = baseY - p.Size * 0.28f + shift;
                    using var pen = new Pen(color, Math.Max(0.5f, p.Size * 0.055f));
                    g.DrawLine(pen, x, sy, x + s.W, sy);
                }
            }
            if (!o.Print && o.ShowMarks)
            {
                using var mb = new SolidBrush(Color.FromArgb(120, 70, 110, 190));
                using var mf = new Font("Segoe UI Symbol", Math.Max(5, p.Size * 0.8f), GraphicsUnit.Point);
                if (s.Kind == SegKind.Space)
                    for (int i = s.Start; i < s.End; i++) g.DrawString("·", mf, mb, pl.X + s.XAt(m, i) + (s.XAt(m, i + 1) - s.XAt(m, i)) / 2 - p.Size * 0.12f, baseY - m.Metrics(p).Asc, TextMeasurer.Format);
                else if (s.Kind == SegKind.Tab) g.DrawString("→", mf, mb, x + s.W / 2 - p.Size * 0.3f, baseY - m.Metrics(p).Asc, TextMeasurer.Format);
                else if (s.Kind == SegKind.Break)
                {
                    bool pb = s.Text == Run.PageBreak.ToString();
                    if (pb)
                    {
                        using var pen = new Pen(Color.FromArgb(140, 70, 110, 190), 0.6f) { DashStyle = DashStyle.Dot };
                        float lx = pl.X + line.Left, rx = pl.X + line.Right;
                        g.DrawLine(pen, lx, pl.Y + line.Height / 2, rx, pl.Y + line.Height / 2);
                        using var sf = new Font("Segoe UI", 7, GraphicsUnit.Point);
                        var t = MicroadsOffice.Core.Office.Loc.T(" Ngắt trang ", " Page Break ");
                        var tw = g.MeasureString(t, sf).Width;
                        g.FillRectangle(Brushes.White, (lx + rx - tw) / 2, pl.Y, tw, line.Height);
                        g.DrawString(t, sf, mb, (lx + rx - tw) / 2, pl.Y + line.Height / 2 - 5);
                    }
                    else g.DrawString("↵", mf, mb, x, baseY - m.Metrics(p).Asc, TextMeasurer.Format);
                }
            }
        }

        // misspellings
        if (!o.Print && !header && o.Misspelled != null)
        {
            var bad = o.Misspelled(line.Para);
            foreach (var (st, len) in bad)
            {
                int a = Math.Max(st, line.Start), b = Math.Min(st + len, line.End);
                if (b <= a) continue;
                float x0 = L.XInLine(line, a), x1 = L.XInLine(line, b);
                DrawWave(g, Color.FromArgb(220, 30, 30), pl.X + x0, baseY + 2.2f, x1 - x0, 0.6f);
            }
        }

        if (!o.Print && o.ShowMarks && line.IsLast && !header)
        {
            var mp = line.Segs.Count > 0 ? line.Segs[^1].Props : line.Para.MarkProps;
            using var mb = new SolidBrush(Color.FromArgb(120, 70, 110, 190));
            using var mf = new Font("Segoe UI", Math.Max(5, mp.Size * 0.85f), GraphicsUnit.Point);
            float ex = line.Segs.Count > 0 ? (line.Segs[^1].Kind == SegKind.Break ? line.Segs[^1].X + 8 : line.Segs[^1].X + line.Segs[^1].W) : L.XInLine(line, line.Start);
            g.DrawString("¶", mf, mb, pl.X + ex + 1, baseY - m.Metrics(mp).Asc, TextMeasurer.Format);
        }
        if (clipState != null) g.Restore(clipState);
    }

    private static void DrawWave(Graphics g, Color c, float x, float y, float w, float th)
    {
        if (w <= 0) return;
        using var pen = new Pen(c, Math.Max(0.5f, th));
        var pts = new List<PointF>();
        float step = 1.5f;
        for (float i = 0; i <= w; i += step) pts.Add(new PointF(x + i, y + ((int)(i / step) % 2 == 0 ? 0 : 1.2f)));
        if (pts.Count > 1) g.DrawLines(pen, pts.ToArray());
    }

    public static RectangleF[] Handles(RectangleF r)
    {
        float h = 5;
        float cx = r.X + r.Width / 2, cy = r.Y + r.Height / 2;
        return
        [
            new(r.Left - h / 2, r.Top - h / 2, h, h), new(cx - h / 2, r.Top - h / 2, h, h), new(r.Right - h / 2, r.Top - h / 2, h, h),
            new(r.Left - h / 2, cy - h / 2, h, h), new(r.Right - h / 2, cy - h / 2, h, h),
            new(r.Left - h / 2, r.Bottom - h / 2, h, h), new(cx - h / 2, r.Bottom - h / 2, h, h), new(r.Right - h / 2, r.Bottom - h / 2, h, h)
        ];
    }
}
