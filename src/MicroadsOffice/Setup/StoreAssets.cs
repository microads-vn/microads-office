using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.Setup;

/// <summary>Logo images required by the MSIX manifest (Start menu tiles, taskbar, Store listing).</summary>
internal static class StoreAssets
{
    public static int Export(string dir)
    {
        Directory.CreateDirectory(dir);
        foreach (var (kind, name) in new[] { (AppKind.Hub, "Office"), (AppKind.Word, "Word"), (AppKind.Excel, "Excel"), (AppKind.Pdf, "Pdf") })
        {
            Save(Path.Combine(dir, name + "44.png"), 44, kind, 1.0f);
            Save(Path.Combine(dir, name + "150.png"), 150, kind, 0.66f);
            // unplated taskbar / Start list icons
            foreach (var t in new[] { 16, 24, 32, 48, 256 })
                Save(Path.Combine(dir, $"{name}44.targetsize-{t}_altform-unplated.png"), t, kind, 1.0f);
        }
        Save(Path.Combine(dir, "StoreLogo.png"), 50, AppKind.Hub, 1.0f);
        using (var wide = new Bitmap(310, 150, PixelFormat.Format32bppArgb))
        {
            using var g = Graphics.FromImage(wide);
            g.Clear(Color.Transparent);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var icon = BrandArt.RenderSuiteIcon(100);
            g.DrawImage(icon, 105, 25, 100, 100);
            wide.Save(Path.Combine(dir, "OfficeWide310.png"), ImageFormat.Png);
        }
        return 0;
    }

    /// <summary>Icon centred on a transparent square; <paramref name="scale"/> leaves Start-tile padding.</summary>
    private static void Save(string path, int size, AppKind kind, float scale)
    {
        using var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Transparent);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            int inner = Math.Max(8, (int)Math.Round(size * scale));
            using var icon = BrandArt.RenderAppIcon(Math.Max(inner, 32), kind);
            int off = (size - inner) / 2;
            g.DrawImage(icon, off, off, inner, inner);
        }
        bmp.Save(path, ImageFormat.Png);
    }
}
