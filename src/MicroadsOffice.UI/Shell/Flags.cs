using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.UI;

/// <summary>Vector national flags used by the language switcher (no image files).</summary>
public static class Flags
{
    private static readonly Dictionary<(string, int, int), Bitmap> Cache = new();

    public static Bitmap Get(string lang, int width, int height)
    {
        var key = (Loc.Normalize(lang), width, height);
        if (Cache.TryGetValue(key, out var b)) return b;
        b = new Bitmap(width, height);
        using (var g = Graphics.FromImage(b))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            var r = new RectangleF(0, 0, width, height);
            if (key.Item1 == Loc.English) DrawUk(g, r); else DrawVietnam(g, r);
            using var pen = new Pen(Color.FromArgb(60, 0, 0, 0));
            g.DrawRectangle(pen, 0, 0, width - 1, height - 1);
        }
        Cache[key] = b;
        return b;
    }

    /// <summary>Small flag for menus / ribbon header (3:2).</summary>
    public static Bitmap Small(string lang) => Get(lang, Dpi.Scale(21), Dpi.Scale(14));

    public static void DrawVietnam(Graphics g, RectangleF r)
    {
        using var red = new SolidBrush(Color.FromArgb(218, 37, 29));
        using var yellow = new SolidBrush(Color.FromArgb(255, 205, 0));
        g.FillRectangle(red, r);
        float cx = r.X + r.Width / 2, cy = r.Y + r.Height / 2;
        float outer = r.Height * 0.32f, inner = outer * 0.382f;
        var pts = new PointF[10];
        for (int i = 0; i < 10; i++)
        {
            double a = -Math.PI / 2 + i * Math.PI / 5;
            float rad = i % 2 == 0 ? outer : inner;
            pts[i] = new PointF(cx + (float)(Math.Cos(a) * rad), cy + (float)(Math.Sin(a) * rad));
        }
        g.FillPolygon(yellow, pts);
    }

    public static void DrawUk(Graphics g, RectangleF r)
    {
        var state = g.Save();
        g.SetClip(r);
        using var blue = new SolidBrush(Color.FromArgb(1, 33, 105));
        using var red = new SolidBrush(Color.FromArgb(200, 16, 46));
        g.FillRectangle(blue, r);
        float h = r.Height;
        using (var white = new Pen(Color.White, h / 5f))
        {
            g.DrawLine(white, r.Left, r.Top, r.Right, r.Bottom);
            g.DrawLine(white, r.Right, r.Top, r.Left, r.Bottom);
        }
        using (var redPen = new Pen(Color.FromArgb(200, 16, 46), h / 12f))
        {
            g.DrawLine(redPen, r.Left, r.Top, r.Right, r.Bottom);
            g.DrawLine(redPen, r.Right, r.Top, r.Left, r.Bottom);
        }
        float cw = h / 3f, rw = h / 5f;
        g.FillRectangle(Brushes.White, r.Left, r.Top + (h - cw) / 2, r.Width, cw);
        g.FillRectangle(Brushes.White, r.Left + (r.Width - cw) / 2, r.Top, cw, h);
        g.FillRectangle(red, r.Left, r.Top + (h - rw) / 2, r.Width, rw);
        g.FillRectangle(red, r.Left + (r.Width - rw) / 2, r.Top, rw, h);
        g.Restore(state);
    }

    public static string LanguageName(string lang) => Loc.Normalize(lang) == Loc.English ? "English" : "Tiếng Việt";
}
