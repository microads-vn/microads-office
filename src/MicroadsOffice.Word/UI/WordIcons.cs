using System.Drawing.Drawing2D;
using MicroadsOffice.UI;

namespace MicroadsOffice.Word.UI;

/// <summary>Icons for Word commands: Segoe Fluent/MDL2 glyphs where they exist, otherwise small vector drawings.</summary>
internal static class WI
{
    public const string Bold = "", Italic = "", Underline = "", FontColor = "", Highlight = "",
        AlignLeft = "", AlignCenter = "", AlignRight = "", Bullets = "", Cut = "", Copy = "",
        Paste = "", Undo = "", Redo = "", Save = "", Print = "", Picture = "", Link = "",
        Table = "", Calendar = "", GrowFont = "", ShrinkFont = "", Sort = "", Spell = "",
        Brush = "", Erase = "", Find = "", Replace = "", Page = "", Zoom = "", Ruler = "",
        Help = "", Info = "", Settings = "", Globe = "", Sync = "", Delete = "", Add = "",
        Open = "", NewDoc = "", Pdf = "", Export = "", Close = "", Fullscreen = "", Select = "",
        Calculator = "", Home = "", Doc = "", ArrowUp = "", ArrowDown = "", Nav = "";

    public static Bitmap G(string glyph, bool large) => large ? Icons.Large(glyph) : Icons.Small(glyph);

    private static float S(bool large) => large ? Dpi.Scale(32) : Dpi.Scale(16);

    public static Bitmap P(string key, bool large, Action<Graphics, float, Color> paint) =>
        Icons.Painted("w_" + key, (int)S(large), Theme.Ink, paint);

    private static Pen Pen(Color c, float s, float w = 1.3f) => new(c, Math.Max(1f, s / 16f * w)) { StartCap = LineCap.Round, EndCap = LineCap.Round };

    /// <summary>Lines like a paragraph (x0,x1 per line as fractions).</summary>
    private static void Lines(Graphics g, float s, Color c, (float A, float B)[] rows, float top = 0.2f, float step = 0.2f)
    {
        using var p = Pen(c, s);
        for (int i = 0; i < rows.Length; i++) g.DrawLine(p, s * rows[i].A, s * (top + i * step), s * rows[i].B, s * (top + i * step));
    }

    private static void Text(Graphics g, float s, Color c, string t, float size, FontStyle st = FontStyle.Regular, float x = 0.5f, float y = 0.5f, string font = "Segoe UI")
    {
        using var f = new Font(font, s * size, st, GraphicsUnit.Pixel);
        using var b = new SolidBrush(c);
        using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString(t, f, b, new RectangleF(s * x - s, s * y - s, 2 * s, 2 * s), sf);
    }

    public static Bitmap Justify(bool l) => P("justify", l, (g, s, c) => Lines(g, s, c, [(0.12f, 0.88f), (0.12f, 0.88f), (0.12f, 0.88f), (0.12f, 0.88f)]));
    public static Bitmap Numbering(bool l) => P("numbering", l, (g, s, c) =>
    {
        Text(g, s, c, "1", 0.3f, FontStyle.Bold, 0.16f, 0.22f); Text(g, s, c, "2", 0.3f, FontStyle.Bold, 0.16f, 0.5f); Text(g, s, c, "3", 0.3f, FontStyle.Bold, 0.16f, 0.78f);
        Lines(g, s, c, [(0.38f, 0.9f), (0.38f, 0.9f), (0.38f, 0.9f)], 0.22f, 0.28f);
    });
    public static Bitmap Multilevel(bool l) => P("multilevel", l, (g, s, c) =>
    {
        Text(g, s, c, "1", 0.26f, FontStyle.Bold, 0.14f, 0.2f); Text(g, s, c, "a", 0.26f, FontStyle.Bold, 0.34f, 0.5f); Text(g, s, c, "i", 0.26f, FontStyle.Bold, 0.54f, 0.8f);
        using var p = Pen(c, s);
        g.DrawLine(p, s * 0.3f, s * 0.2f, s * 0.9f, s * 0.2f); g.DrawLine(p, s * 0.5f, s * 0.5f, s * 0.9f, s * 0.5f); g.DrawLine(p, s * 0.7f, s * 0.8f, s * 0.9f, s * 0.8f);
    });
    public static Bitmap IndentMore(bool l) => P("indentmore", l, (g, s, c) =>
    {
        Lines(g, s, c, [(0.1f, 0.9f), (0.45f, 0.9f), (0.45f, 0.9f), (0.1f, 0.9f)], 0.16f, 0.22f);
        using var b = new SolidBrush(Theme.Accent);
        g.FillPolygon(b, [new PointF(s * 0.1f, s * 0.32f), new PointF(s * 0.32f, s * 0.48f), new PointF(s * 0.1f, s * 0.64f)]);
    });
    public static Bitmap IndentLess(bool l) => P("indentless", l, (g, s, c) =>
    {
        Lines(g, s, c, [(0.1f, 0.9f), (0.45f, 0.9f), (0.45f, 0.9f), (0.1f, 0.9f)], 0.16f, 0.22f);
        using var b = new SolidBrush(Theme.Accent);
        g.FillPolygon(b, [new PointF(s * 0.32f, s * 0.32f), new PointF(s * 0.1f, s * 0.48f), new PointF(s * 0.32f, s * 0.64f)]);
    });
    public static Bitmap LineSpacing(bool l) => P("linespacing", l, (g, s, c) =>
    {
        Lines(g, s, c, [(0.42f, 0.92f), (0.42f, 0.92f), (0.42f, 0.92f), (0.42f, 0.92f)], 0.18f, 0.21f);
        using var p = Pen(Theme.Accent, s);
        g.DrawLine(p, s * 0.2f, s * 0.12f, s * 0.2f, s * 0.88f);
        g.DrawLine(p, s * 0.08f, s * 0.26f, s * 0.2f, s * 0.12f); g.DrawLine(p, s * 0.32f, s * 0.26f, s * 0.2f, s * 0.12f);
        g.DrawLine(p, s * 0.08f, s * 0.74f, s * 0.2f, s * 0.88f); g.DrawLine(p, s * 0.32f, s * 0.74f, s * 0.2f, s * 0.88f);
    });
    public static Bitmap Pilcrow(bool l) => P("pilcrow", l, (g, s, c) => Text(g, s, c, "¶", 0.85f, FontStyle.Regular, 0.5f, 0.5f));
    public static Bitmap Strike(bool l) => P("strike", l, (g, s, c) =>
    {
        Text(g, s, c, "ab", 0.62f, FontStyle.Regular, 0.5f, 0.46f);
        using var p = Pen(c, s, 1.1f);
        g.DrawLine(p, s * 0.1f, s * 0.52f, s * 0.9f, s * 0.52f);
    });
    public static Bitmap Subscript(bool l) => P("sub", l, (g, s, c) => { Text(g, s, c, "x", 0.7f, FontStyle.Regular, 0.38f, 0.42f); using var b = new SolidBrush(Theme.Accent); Text(g, s, Theme.Accent, "2", 0.42f, FontStyle.Bold, 0.78f, 0.72f); });
    public static Bitmap Superscript(bool l) => P("sup", l, (g, s, c) => { Text(g, s, c, "x", 0.7f, FontStyle.Regular, 0.38f, 0.58f); Text(g, s, Theme.Accent, "2", 0.42f, FontStyle.Bold, 0.78f, 0.24f); });
    public static Bitmap ChangeCase(bool l) => P("case", l, (g, s, c) => Text(g, s, c, "Aa", 0.58f, FontStyle.Regular));
    public static Bitmap ClearFormat(bool l) => P("clearfmt", l, (g, s, c) =>
    {
        Text(g, s, c, "A", 0.74f, FontStyle.Regular, 0.4f, 0.46f);
        using var b = new SolidBrush(Color.FromArgb(232, 80, 120));
        g.FillRectangle(b, s * 0.55f, s * 0.6f, s * 0.38f, s * 0.26f);
    });
    public static Bitmap Shading(bool l) => P("shading", l, (g, s, c) =>
    {
        using var p = Pen(c, s);
        g.DrawPolygon(p, [new PointF(s * 0.2f, s * 0.45f), new PointF(s * 0.5f, s * 0.15f), new PointF(s * 0.78f, s * 0.43f), new PointF(s * 0.48f, s * 0.73f)]);
        using var b = new SolidBrush(c);
        g.FillEllipse(b, s * 0.74f, s * 0.5f, s * 0.14f, s * 0.2f);
    });
    public static Bitmap Borders(bool l) => P("borders", l, (g, s, c) =>
    {
        using var dot = new Pen(Color.FromArgb(150, c), Math.Max(1, s / 16f)) { DashStyle = DashStyle.Dot };
        g.DrawRectangle(dot, s * 0.14f, s * 0.14f, s * 0.72f, s * 0.72f);
        g.DrawLine(dot, s * 0.5f, s * 0.14f, s * 0.5f, s * 0.86f); g.DrawLine(dot, s * 0.14f, s * 0.5f, s * 0.86f, s * 0.5f);
        using var p = Pen(c, s, 1.6f);
        g.DrawLine(p, s * 0.14f, s * 0.86f, s * 0.86f, s * 0.86f);
    });
    public static Bitmap PageBreak(bool l) => P("pagebreak", l, (g, s, c) =>
    {
        using var p = Pen(c, s);
        g.DrawLines(p, [new PointF(s * 0.22f, s * 0.38f), new PointF(s * 0.22f, s * 0.08f), new PointF(s * 0.78f, s * 0.08f), new PointF(s * 0.78f, s * 0.38f)]);
        g.DrawLines(p, [new PointF(s * 0.22f, s * 0.62f), new PointF(s * 0.22f, s * 0.92f), new PointF(s * 0.78f, s * 0.92f), new PointF(s * 0.78f, s * 0.62f)]);
        using var d = new Pen(Theme.Accent, Math.Max(1, s / 16f)) { DashStyle = DashStyle.Dash };
        g.DrawLine(d, s * 0.06f, s * 0.5f, s * 0.94f, s * 0.5f);
    });
    private static void PageShape(Graphics g, float s, Color c)
    {
        using var p = Pen(c, s);
        g.DrawRectangle(p, s * 0.2f, s * 0.08f, s * 0.6f, s * 0.84f);
    }
    public static Bitmap Header(bool l) => P("header", l, (g, s, c) => { PageShape(g, s, c); using var b = new SolidBrush(Theme.Accent); g.FillRectangle(b, s * 0.28f, s * 0.15f, s * 0.44f, s * 0.12f); Lines(g, s, Color.FromArgb(150, c), [(0.3f, 0.7f), (0.3f, 0.7f)], 0.5f, 0.16f); });
    public static Bitmap Footer(bool l) => P("footer", l, (g, s, c) => { PageShape(g, s, c); using var b = new SolidBrush(Theme.Accent); g.FillRectangle(b, s * 0.28f, s * 0.73f, s * 0.44f, s * 0.12f); Lines(g, s, Color.FromArgb(150, c), [(0.3f, 0.7f), (0.3f, 0.7f)], 0.25f, 0.16f); });
    public static Bitmap PageNumber(bool l) => P("pagenum", l, (g, s, c) => { PageShape(g, s, c); Text(g, s, Theme.Accent, "#", 0.4f, FontStyle.Bold, 0.5f, 0.72f); });
    public static Bitmap Symbol(bool l) => P("symbol", l, (g, s, c) => Text(g, s, c, "Ω", 0.8f, FontStyle.Regular, 0.5f, 0.52f));
    public static Bitmap HLine(bool l) => P("hline", l, (g, s, c) => { using var p = Pen(c, s, 1.8f); g.DrawLine(p, s * 0.08f, s * 0.5f, s * 0.92f, s * 0.5f); Lines(g, s, Color.FromArgb(120, c), [(0.2f, 0.8f), (0.2f, 0.8f)], 0.25f, 0.5f); });
    public static Bitmap Margins(bool l) => P("margins", l, (g, s, c) =>
    {
        PageShape(g, s, c);
        using var d = new Pen(Theme.Accent, Math.Max(1, s / 16f)) { DashStyle = DashStyle.Dot };
        g.DrawRectangle(d, s * 0.3f, s * 0.2f, s * 0.4f, s * 0.6f);
    });
    public static Bitmap Orientation(bool l) => P("orient", l, (g, s, c) =>
    {
        using var p = Pen(c, s);
        g.DrawRectangle(p, s * 0.1f, s * 0.12f, s * 0.42f, s * 0.6f);
        using var a = Pen(Theme.Accent, s);
        g.DrawRectangle(a, s * 0.36f, s * 0.48f, s * 0.56f, s * 0.38f);
    });
    public static Bitmap PaperSize(bool l) => P("paper", l, (g, s, c) => { PageShape(g, s, c); Text(g, s, Theme.Accent, "A4", 0.3f, FontStyle.Bold); });
    public static Bitmap Watermark(bool l) => P("watermark", l, (g, s, c) =>
    {
        PageShape(g, s, c);
        var st = g.Save();
        g.TranslateTransform(s * 0.5f, s * 0.5f); g.RotateTransform(-55);
        Text(g, s, Color.FromArgb(160, Theme.Accent), "A", 0.5f, FontStyle.Bold, 0, 0);
        g.Restore(st);
    });
    public static Bitmap PageColor(bool l) => P("pagecolor", l, (g, s, c) =>
    {
        using var b = new SolidBrush(Color.FromArgb(255, 236, 179));
        g.FillRectangle(b, s * 0.2f, s * 0.08f, s * 0.6f, s * 0.84f);
        PageShape(g, s, c);
    });
    public static Bitmap Toc(bool l) => P("toc", l, (g, s, c) =>
    {
        PageShape(g, s, c);
        using var d = new Pen(c, Math.Max(1, s / 20f)) { DashStyle = DashStyle.Dot };
        for (int i = 0; i < 4; i++) { float y = s * (0.25f + i * 0.17f); g.DrawLine(d, s * 0.28f, y, s * 0.66f, y); Text(g, s, Theme.Accent, (i + 1).ToString(), 0.16f, FontStyle.Bold, 0.71f, (0.25f + i * 0.17f)); }
    });
    public static Bitmap UpdateToc(bool l) => P("updtoc", l, (g, s, c) => { Text(g, s, c, "", 0.6f, FontStyle.Regular, 0.5f, 0.5f, Icons.FontName); });
    public static Bitmap WordCount(bool l) => P("wordcount", l, (g, s, c) => { Text(g, s, c, "ABC", 0.3f, FontStyle.Bold, 0.5f, 0.3f); Text(g, s, Theme.Accent, "123", 0.3f, FontStyle.Bold, 0.5f, 0.68f); });
    public static Bitmap OnePage(bool l) => P("onepage", l, (g, s, c) => PageShape(g, s, c));
    public static Bitmap PageWidth(bool l) => P("pagewidth", l, (g, s, c) =>
    {
        PageShape(g, s, c);
        using var p = Pen(Theme.Accent, s);
        g.DrawLine(p, s * 0.05f, s * 0.5f, s * 0.95f, s * 0.5f);
        g.DrawLine(p, s * 0.05f, s * 0.5f, s * 0.16f, s * 0.4f); g.DrawLine(p, s * 0.05f, s * 0.5f, s * 0.16f, s * 0.6f);
        g.DrawLine(p, s * 0.95f, s * 0.5f, s * 0.84f, s * 0.4f); g.DrawLine(p, s * 0.95f, s * 0.5f, s * 0.84f, s * 0.6f);
    });
    public static Bitmap Zoom100(bool l) => P("zoom100", l, (g, s, c) => Text(g, s, c, "100", 0.34f, FontStyle.Bold));
    public static Bitmap PrintLayout(bool l) => P("printlayout", l, (g, s, c) => { PageShape(g, s, c); Lines(g, s, Color.FromArgb(150, c), [(0.3f, 0.7f), (0.3f, 0.7f), (0.3f, 0.6f)], 0.3f, 0.14f); });
    public static Bitmap NavPane(bool l) => P("navpane", l, (g, s, c) =>
    {
        using var p = Pen(c, s);
        g.DrawRectangle(p, s * 0.1f, s * 0.15f, s * 0.8f, s * 0.7f);
        using var b = new SolidBrush(Theme.Accent);
        g.FillRectangle(b, s * 0.14f, s * 0.2f, s * 0.24f, s * 0.6f);
    });
    public static Bitmap RowAbove(bool l) => TableArrow("rowabove", l, 0);
    public static Bitmap RowBelow(bool l) => TableArrow("rowbelow", l, 1);
    public static Bitmap ColLeft(bool l) => TableArrow("colleft", l, 2);
    public static Bitmap ColRight(bool l) => TableArrow("colright", l, 3);

    private static Bitmap TableArrow(string key, bool l, int dir) => P(key, l, (g, s, c) =>
    {
        using var p = Pen(c, s, 1f);
        var r = dir switch { 0 => new RectangleF(0.15f, 0.45f, 0.7f, 0.42f), 1 => new RectangleF(0.15f, 0.13f, 0.7f, 0.42f), 2 => new RectangleF(0.45f, 0.15f, 0.42f, 0.7f), _ => new RectangleF(0.13f, 0.15f, 0.42f, 0.7f) };
        g.DrawRectangle(p, s * r.X, s * r.Y, s * r.Width, s * r.Height);
        if (dir < 2) g.DrawLine(p, s * 0.5f, s * r.Top, s * 0.5f, s * r.Bottom); else g.DrawLine(p, s * r.Left, s * 0.5f, s * r.Right, s * 0.5f);
        using var b = new SolidBrush(Theme.Accent);
        var plus = dir switch { 0 => new PointF(0.5f, 0.2f), 1 => new PointF(0.5f, 0.8f), 2 => new PointF(0.2f, 0.5f), _ => new PointF(0.8f, 0.5f) };
        Text(g, s, Theme.Accent, "+", 0.5f, FontStyle.Bold, plus.X, plus.Y);
    });

    public static Bitmap DeleteRow(bool l) => P("delrow", l, (g, s, c) => { using var p = Pen(c, s, 1f); g.DrawRectangle(p, s * 0.12f, s * 0.3f, s * 0.76f, s * 0.4f); Text(g, s, Theme.ErrRed, "×", 0.6f, FontStyle.Bold, 0.5f, 0.48f); });
    public static Bitmap DeleteCol(bool l) => P("delcol", l, (g, s, c) => { using var p = Pen(c, s, 1f); g.DrawRectangle(p, s * 0.3f, s * 0.12f, s * 0.4f, s * 0.76f); Text(g, s, Theme.ErrRed, "×", 0.6f, FontStyle.Bold, 0.5f, 0.48f); });
    public static Bitmap DeleteTable(bool l) => P("deltable", l, (g, s, c) => { using var p = Pen(c, s, 1f); g.DrawRectangle(p, s * 0.12f, s * 0.12f, s * 0.76f, s * 0.76f); g.DrawLine(p, s * 0.12f, s * 0.5f, s * 0.88f, s * 0.5f); g.DrawLine(p, s * 0.5f, s * 0.12f, s * 0.5f, s * 0.88f); Text(g, s, Theme.ErrRed, "×", 0.62f, FontStyle.Bold, 0.72f, 0.7f); });
    public static Bitmap Merge(bool l) => P("merge", l, (g, s, c) =>
    {
        using var p = Pen(c, s, 1f);
        g.DrawRectangle(p, s * 0.1f, s * 0.25f, s * 0.8f, s * 0.5f);
        using var a = Pen(Theme.Accent, s);
        g.DrawLine(a, s * 0.22f, s * 0.5f, s * 0.42f, s * 0.5f); g.DrawLine(a, s * 0.58f, s * 0.5f, s * 0.78f, s * 0.5f);
        g.DrawLine(a, s * 0.34f, s * 0.42f, s * 0.42f, s * 0.5f); g.DrawLine(a, s * 0.34f, s * 0.58f, s * 0.42f, s * 0.5f);
        g.DrawLine(a, s * 0.66f, s * 0.42f, s * 0.58f, s * 0.5f); g.DrawLine(a, s * 0.66f, s * 0.58f, s * 0.58f, s * 0.5f);
    });
    public static Bitmap CellAlign(bool l, int v) => P("cellalign" + v, l, (g, s, c) =>
    {
        using var p = Pen(Color.FromArgb(150, c), s, 0.8f);
        g.DrawRectangle(p, s * 0.1f, s * 0.1f, s * 0.8f, s * 0.8f);
        float y = v == 0 ? 0.25f : v == 1 ? 0.43f : 0.62f;
        Lines(g, s, c, [(0.25f, 0.75f), (0.25f, 0.75f)], y, 0.16f);
    });
    public static Bitmap ResetPicture(bool l) => P("resetpic", l, (g, s, c) => { Text(g, s, c, "", 0.7f, FontStyle.Regular, 0.42f, 0.5f, Icons.FontName); Text(g, s, Theme.Accent, "", 0.38f, FontStyle.Regular, 0.78f, 0.78f, Icons.FontName); });
    public static Bitmap StyleGlyph(bool l) => P("styles", l, (g, s, c) => Text(g, s, c, "Aa", 0.55f, FontStyle.Bold));
}
