using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.Core.Branding;

/// <summary>Vector re-creation of the Microads brand mark and the Microads Office application icon.</summary>
public static class BrandArt
{
    public static readonly Color Red = Color.FromArgb(232, 49, 47);
    public static readonly Color RedDark = Color.FromArgb(196, 32, 36);
    public static readonly Color Green = Color.FromArgb(43, 168, 74);
    public static readonly Color Ink = Color.FromArgb(24, 24, 27);
    public static readonly Color Blue = Color.FromArgb(37, 99, 235);
    public static readonly Color BlueDark = Color.FromArgb(29, 78, 196);
    public static readonly Color GreenDark = Color.FromArgb(28, 128, 56);

    /// <summary>Accent colour of each application.</summary>
    public static (Color Accent, Color Dark) AccentOf(AppKind k) => k switch
    {
        AppKind.Word => (Blue, BlueDark),
        AppKind.Excel => (Green, GreenDark),
        _ => (Red, RedDark)
    };

    public static string BandText(AppKind k) => k switch { AppKind.Word => "DOC", AppKind.Excel => "XLS", AppKind.Pdf => "PDF", _ => "" };

    /// <summary>Draws the italic "N" mark (red, green cap) into the given rectangle (aspect 100:80).</summary>
    public static void DrawMark(Graphics g, RectangleF r)
    {
        float sx = r.Width / 100f, sy = r.Height / 80f;
        PointF P(float x, float y) => new(r.X + x * sx, r.Y + y * sy);
        var state = g.Save();
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        using var red = new SolidBrush(Red);
        using var green = new SolidBrush(Green);
        using (var path = new GraphicsPath(FillMode.Winding))
        {
            path.AddPolygon([P(0, 80), P(26, 80), P(48, 0), P(22, 0)]);         // left stem
            path.AddPolygon([P(22, 0), P(48, 0), P(78, 80), P(52, 80)]);        // diagonal
            path.AddPolygon([P(52, 80), P(78, 80), P(100, 0), P(74, 0)]);       // right stem
            g.FillPath(red, path);
        }
        // green cap on the right stem (top ~ 32%)
        float capY = 26;
        float xl = 74 - 22 * capY / 80f, xr = 100 - 22 * capY / 80f;
        g.FillPolygon(green, [P(74, 0), P(100, 0), P(xr, capY), P(xl, capY)]);
        g.Restore(state);
    }

    /// <summary>Application icon: a white document sheet (green folded corner) carrying the brand mark and a red "PDF" band.</summary>
    public static Bitmap RenderAppIcon(int size) => RenderAppIcon(size, AppKind.Pdf);

    public static Bitmap RenderAppIcon(int size, AppKind kind)
    {
        if (kind == AppKind.Hub) return RenderSuiteIcon(size);
        var (bandColor, bandDark) = AccentOf(kind);
        var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
        float s = size / 256f;
        bool tiny = size <= 24;

        // Sheet with folded corner
        float left = 34 * s, top = 10 * s, right = 222 * s, bottom = 246 * s, fold = 58 * s, rad = 14 * s;
        if (tiny) { left = 20 * s; right = 236 * s; top = 6 * s; bottom = 250 * s; fold = 64 * s; }
        using (var sheet = new GraphicsPath())
        {
            sheet.AddLine(left + rad, top, right - fold, top);
            sheet.AddLine(right - fold, top, right, top + fold);
            sheet.AddLine(right, top + fold, right, bottom - rad);
            sheet.AddArc(right - 2 * rad, bottom - 2 * rad, 2 * rad, 2 * rad, 0, 90);
            sheet.AddArc(left, bottom - 2 * rad, 2 * rad, 2 * rad, 90, 90);
            sheet.AddArc(left, top, 2 * rad, 2 * rad, 180, 90);
            sheet.CloseFigure();
            if (!tiny)
            {
                using var shadow = new SolidBrush(Color.FromArgb(40, 0, 0, 0));
                var st = g.Save();
                g.TranslateTransform(4 * s, 5 * s);
                g.FillPath(shadow, sheet);
                g.Restore(st);
            }
            using var white = new SolidBrush(Color.White);
            g.FillPath(white, sheet);
            using var border = new Pen(Color.FromArgb(205, 208, 214), Math.Max(1f, 4 * s));
            g.DrawPath(border, sheet);
        }
        using (var foldBrush = new SolidBrush(Color.FromArgb(222, 226, 231)))
            g.FillPolygon(foldBrush, [new PointF(right - fold, top), new PointF(right, top + fold), new PointF(right - fold, top + fold)]);

        // Brand mark
        var markRect = tiny ? new RectangleF(40 * s, 70 * s, 176 * s, 140 * s) : new RectangleF(62 * s, 66 * s, 132 * s, 105.6f * s);
        DrawMark(g, markRect);

        // Coloured band with the document type (DOC / XLS / PDF). On tiny sizes a plain stripe keeps the colour cue.
        if (!tiny)
        {
            var band = new RectangleF(left, 186 * s, right - left, 44 * s);
            using var bandBrush = new LinearGradientBrush(band, bandColor, bandDark, LinearGradientMode.Vertical);
            g.FillRectangle(bandBrush, band);
            using var f = new Font("Segoe UI", 30 * s, FontStyle.Bold, GraphicsUnit.Pixel);
            using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(BandText(kind), f, Brushes.White, band, sf);
        }
        else
        {
            using var stripe = new SolidBrush(bandColor);
            g.FillRectangle(stripe, left, 214 * s, right - left, 34 * s);
        }
        return bmp;
    }

    /// <summary>Suite icon (Microads Office hub): rounded tile with the brand mark and three app colour dots.</summary>
    public static Bitmap RenderSuiteIcon(int size)
    {
        var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        float s = size / 256f;
        var tile = new RectangleF(12 * s, 12 * s, 232 * s, 232 * s);
        using (var path = RoundRect(tile, 48 * s))
        {
            using var white = new SolidBrush(Color.White);
            g.FillPath(white, path);
            using var border = new Pen(Color.FromArgb(205, 208, 214), Math.Max(1f, 5 * s));
            g.DrawPath(border, path);
        }
        DrawMark(g, size <= 24 ? new RectangleF(30 * s, 46 * s, 196 * s, 156.8f * s) : new RectangleF(48 * s, 50 * s, 160 * s, 128 * s));
        if (size > 24)
        {
            float d = 26 * s, y = 196 * s, cx = 128 * s;
            foreach (var (c, i) in new[] { (Blue, -1), (Green, 0), (Red, 1) })
            {
                using var b = new SolidBrush(c);
                g.FillEllipse(b, cx + i * 40 * s - d / 2, y - d / 2, d, d);
            }
        }
        return bmp;
    }

    public static GraphicsPath RoundRect(RectangleF r, float rad)
    {
        var p = new GraphicsPath();
        float d = rad * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    /// <summary>Horizontal logo: mark + "Microads Office" wordmark.</summary>
    public static Bitmap RenderWordmark(int height, bool darkBackground = false) => RenderWordmark(height, darkBackground, null);

    /// <summary>Horizontal logo "MICROADS Office" (+ optional app name such as "Word") with "microads.vn" underneath.</summary>
    public static Bitmap RenderWordmark(int height, bool darkBackground, AppKind? app)
    {
        string office = app is AppKind.Word or AppKind.Excel or AppKind.Pdf ? "Office " + Product.AppShortName(app.Value) : "Office";
        var officeColor = app is AppKind.Word or AppKind.Excel ? AccentOf(app.Value).Accent : Red;
        float s = height / 100f;
        using var fBold = new Font("Segoe UI", 46 * s, FontStyle.Bold, GraphicsUnit.Pixel);
        using var fLight = new Font("Segoe UI Light", 46 * s, FontStyle.Regular, GraphicsUnit.Pixel);
        using var fSmall = new Font("Segoe UI", 15 * s, FontStyle.Regular, GraphicsUnit.Pixel);
        int width;
        using (var tmp = new Bitmap(1, 1))
        using (var tg = Graphics.FromImage(tmp))
        {
            var w1 = tg.MeasureString("MICROADS", fBold).Width;
            var w2 = tg.MeasureString(" " + office, fLight).Width;
            width = (int)(100 * s + w1 + w2 + 10 * s);
        }
        var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
        DrawMark(g, new RectangleF(4 * s, 18 * s, 80 * s, 64 * s));
        var ink = darkBackground ? Color.White : Ink;
        using var inkBrush = new SolidBrush(ink);
        float x = 92 * s;
        g.DrawString("MICROADS", fBold, inkBrush, x, 8 * s);
        var w = g.MeasureString("MICROADS", fBold).Width;
        using var redBrush = new SolidBrush(officeColor);
        g.DrawString(office, fLight, redBrush, x + w - 8 * s, 8 * s);
        using var sub = new SolidBrush(darkBackground ? Color.FromArgb(200, 255, 255, 255) : Color.FromArgb(110, 110, 120));
        g.DrawString("microads.vn", fSmall, sub, x + 4 * s, 70 * s);
        return bmp;
    }

    /// <summary>Writes a multi-resolution .ico (PNG-compressed entries).</summary>
    public static void WriteIcon(string path, params int[] sizes) => WriteIcon(path, AppKind.Pdf, sizes);

    public static void WriteIcon(string path, AppKind kind, params int[] sizes)
    {
        if (sizes.Length == 0) sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256];
        var images = sizes.Select(sz =>
        {
            using var b = RenderAppIcon(sz, kind);
            using var ms = new MemoryStream();
            b.Save(ms, ImageFormat.Png);
            return (sz, ms.ToArray());
        }).ToList();
        using var fs = File.Create(path);
        using var w = new BinaryWriter(fs);
        w.Write((short)0); w.Write((short)1); w.Write((short)images.Count);
        int offset = 6 + 16 * images.Count;
        foreach (var (sz, data) in images)
        {
            w.Write((byte)(sz >= 256 ? 0 : sz)); w.Write((byte)(sz >= 256 ? 0 : sz));
            w.Write((byte)0); w.Write((byte)0);
            w.Write((short)1); w.Write((short)32);
            w.Write(data.Length); w.Write(offset);
            offset += data.Length;
        }
        foreach (var (_, data) in images) w.Write(data);
    }

    public static Icon CreateIcon() => CreateIcon(AppKind.Pdf);

    public static Icon CreateIcon(AppKind kind)
    {
        using var ms = new MemoryStream();
        // build in memory
        var sizes = new[] { 16, 24, 32, 48, 256 };
        var images = sizes.Select(sz =>
        {
            using var b = RenderAppIcon(sz, kind);
            using var m = new MemoryStream();
            b.Save(m, ImageFormat.Png);
            return (sz, m.ToArray());
        }).ToList();
        var w = new BinaryWriter(ms);
        w.Write((short)0); w.Write((short)1); w.Write((short)images.Count);
        int offset = 6 + 16 * images.Count;
        foreach (var (sz, data) in images)
        {
            w.Write((byte)(sz >= 256 ? 0 : sz)); w.Write((byte)(sz >= 256 ? 0 : sz));
            w.Write((byte)0); w.Write((byte)0); w.Write((short)1); w.Write((short)32);
            w.Write(data.Length); w.Write(offset);
            offset += data.Length;
        }
        foreach (var (_, data) in images) w.Write(data);
        ms.Position = 0;
        return new Icon(ms);
    }
}
