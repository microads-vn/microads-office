using System.Drawing.Drawing2D;
using System.Drawing.Text;
using MicroadsOffice.Core.Branding;

namespace MicroadsOffice.UI;

public static class Theme
{
    /// <summary>Accent of the running application (Word = blue, Excel = green, PDF = Microads red). Set once at start-up via <see cref="UseAccent"/>.</summary>
    public static Color Accent { get; private set; } = BrandArt.Red;
    public static Color AccentDark { get; private set; } = BrandArt.RedDark;
    /// <summary>Light tint of the accent used for hover backgrounds.</summary>
    public static Color Hover { get; private set; } = Color.FromArgb(253, 232, 231);
    public static Color Pressed { get; private set; } = Color.FromArgb(250, 205, 203);
    public static Color PressedBorder { get; private set; } = Color.FromArgb(230, 140, 135);

    public static void UseAccent(Color accent, Color dark)
    {
        Accent = accent; AccentDark = dark;
        Hover = Blend(accent, Color.White, 0.88f);
        Pressed = Blend(accent, Color.White, 0.76f);
        PressedBorder = Blend(accent, Color.White, 0.45f);
    }

    public static Color Blend(Color a, Color b, float t)
        => Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));

    public static readonly Color AccentGreen = BrandArt.Green;
    public static readonly Color Ink = Color.FromArgb(32, 33, 36);
    public static readonly Color SubtleInk = Color.FromArgb(95, 99, 104);
    public static readonly Color RibbonBack = Color.FromArgb(250, 250, 251);
    public static readonly Color TabStripBack = Color.FromArgb(243, 243, 245);
    public static readonly Color Border = Color.FromArgb(221, 223, 227);
    public static readonly Color ViewerBack = Color.FromArgb(222, 224, 228);
    public static readonly Color PanelBack = Color.FromArgb(247, 247, 248);
    public static readonly Color Selection = Color.FromArgb(90, 51, 153, 255);
    public static readonly Color OkGreen = Color.FromArgb(30, 142, 62);
    public static readonly Color WarnAmber = Color.FromArgb(214, 140, 0);
    public static readonly Color ErrRed = Color.FromArgb(200, 30, 30);

    public static readonly Font UiFont = new("Segoe UI", 9f);
    public static readonly Font UiFontBold = new("Segoe UI Semibold", 9f);
    public static readonly Font SmallFont = new("Segoe UI", 8.25f);
    public static readonly Font TitleFont = new("Segoe UI Semibold", 13f);
}

/// <summary>Icon factory based on the Windows "Segoe Fluent Icons" / "Segoe MDL2 Assets" fonts (no image files needed).</summary>
public static class Icons
{
    public const string Open = "", Save = "", SaveAs = "", Print = "", ZoomIn = "", ZoomOut = "",
        FitPage = "", FitWidth = "", FullScreen = "", ExitFullScreen = "", Rotate = "", Hand = "",
        SelectText = "", Copy = "", Search = "", First = "", Last = "", Prev = "", Next = "",
        Sign = "", Certificate = "", Key = "", Shield = "", Picture = "", Style = "",
        Settings = "", Info = "", Help = "", Thumbnails = "", Bookmarks = "", SinglePage = "",
        Continuous = "", NavPane = "", Recent = "", Close = "", Exit = "", Properties = "",
        Check = "", Unknown = "", Error = "", Pen = "", Add = "", Delete = "", Document = "",
        Signatures = "", GoTo = "",        Sync = "", More = "", Launcher = "", Globe = "", Undo = "", Redo = "",
        Cut = "", Paste = "", NewDoc = "", Home = "", Link = "", Table = "",
        Calendar = "", Font = "", Bold = "", Italic = "", Underline = "",
        AlignLeft = "", AlignCenter = "", AlignRight = "", Bullets = "", Sort = "",
        Filter = "", Chart = "", Comment = "", Lock = "", Mail = "", Phone = "",
        Back = "", Pdf = "", Export = "", Image = "", Highlight = "", FontColor = "",
        Erase = "", Replace = "", Freeze = "", Function = "", Sum = "", Ruler = "",
        Page = "", Header = "", Symbol = "", Columns = "", Spell = "", Count = "",
        Grid = "", Paint = "", Clear = "", Fill = "", Merge = "", Wrap = "";

    private static readonly Dictionary<(string, int, int), Bitmap> PaintedCache = new();

    /// <summary>
    /// Vector icon drawn by code (for commands without a suitable font glyph). The painter receives the
    /// graphics, the icon size in pixels and the ink colour; results are cached.
    /// </summary>
    public static Bitmap Painted(string key, int size, Color color, Action<Graphics, float, Color> paint)
    {
        var k = (key, size, color.ToArgb());
        if (PaintedCache.TryGetValue(k, out var b)) return b;
        b = new Bitmap(size, size);
        using (var g = Graphics.FromImage(b))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            paint(g, size, color);
        }
        PaintedCache[k] = b;
        return b;
    }

    private static readonly Dictionary<(string, int, int, bool), Bitmap> Cache = new();
    private static string? _fontName;

    public static string FontName
    {
        get
        {
            if (_fontName != null) return _fontName;
            using var fonts = new InstalledFontCollection();
            _fontName = fonts.Families.Any(f => f.Name == "Segoe Fluent Icons") ? "Segoe Fluent Icons"
                : fonts.Families.Any(f => f.Name == "Segoe MDL2 Assets") ? "Segoe MDL2 Assets" : "Segoe UI Symbol";
            return _fontName;
        }
    }

    public static Bitmap Get(string glyph, int size, Color color, bool mirror = false)
    {
        var key = (glyph, size, color.ToArgb(), mirror);
        if (Cache.TryGetValue(key, out var b)) return b;
        b = new Bitmap(size, size);
        using (var g = Graphics.FromImage(b))
        {
            g.Clear(Color.Transparent);
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (mirror) { g.TranslateTransform(size, 0); g.ScaleTransform(-1, 1); }
            if (glyph.Length <= 2 && glyph[0] >= '')
            {
                using var f = new Font(FontName, size * 0.78f, GraphicsUnit.Pixel);
                using var br = new SolidBrush(color);
                using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(glyph, f, br, new RectangleF(0, 1, size, size), sf);
            }
            else
            {
                // text icon such as "1:1"
                using var f = new Font("Segoe UI", size * 0.42f, FontStyle.Bold, GraphicsUnit.Pixel);
                using var br = new SolidBrush(color);
                using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                using var pen = new Pen(color, Math.Max(1, size / 16f));
                g.DrawRectangle(pen, size * 0.12f, size * 0.2f, size * 0.76f, size * 0.6f);
                g.DrawString(glyph, f, br, new RectangleF(0, 0, size, size), sf);
            }
        }
        Cache[key] = b;
        return b;
    }

    public static Bitmap Large(string glyph, Color? c = null, bool mirror = false) => Get(glyph, Dpi.Scale(28), c ?? Theme.Ink, mirror);
    public static Bitmap Small(string glyph, Color? c = null, bool mirror = false) => Get(glyph, Dpi.Scale(16), c ?? Theme.Ink, mirror);
}

public static class Dpi
{
    public static float Factor { get; set; } = 1f;
    public static int Scale(int v) => (int)Math.Round(v * Factor);
}
