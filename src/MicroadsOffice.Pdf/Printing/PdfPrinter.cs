using System.Drawing.Printing;
using MicroadsOffice.Core.Pdfium;

namespace MicroadsOffice.Printing;

public enum PrintScaling { Fit, ActualSize, ShrinkOversized }

public sealed class PrintOptions
{
    public string PrinterName { get; set; } = new PrinterSettings().PrinterName;
    public short Copies { get; set; } = 1;
    public bool Collate { get; set; } = true;
    public List<int> Pages { get; set; } = new();
    public PrintScaling Scaling { get; set; } = PrintScaling.Fit;
    public bool AutoRotate { get; set; } = true;
    public bool AutoCenter { get; set; } = true;
    public bool Grayscale { get; set; }
    public bool Annotations { get; set; } = true;
    public bool PrintToFile { get; set; }
    public string? OutputFile { get; set; }
    public PageSettings? PageSettings { get; set; }
    /// <summary>Full driver settings chosen with "Thuộc tính..." (duplex, tray, quality...).</summary>
    public PrinterSettings? DriverSettings { get; set; }
}

public static class PdfPrinter
{
    /// <summary>Parses "1-3, 5, 8-" style ranges (1-based) into 0-based page indexes.</summary>
    public static List<int>? ParseRange(string text, int pageCount)
    {
        var result = new List<int>();
        if (string.IsNullOrWhiteSpace(text)) return null;
        foreach (var raw in text.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries))
        {
            var part = raw.Trim();
            int dash = part.IndexOf('-');
            if (dash < 0)
            {
                if (!int.TryParse(part, out int p) || p < 1 || p > pageCount) return null;
                result.Add(p - 1);
            }
            else
            {
                var a = part[..dash].Trim();
                var b = part[(dash + 1)..].Trim();
                int from = a.Length == 0 ? 1 : int.TryParse(a, out int x) ? x : -1;
                int to = b.Length == 0 ? pageCount : int.TryParse(b, out int y) ? y : -1;
                if (from < 1 || to < 1 || from > pageCount || to > pageCount) return null;
                if (from <= to) for (int i = from; i <= to; i++) result.Add(i - 1);
                else for (int i = from; i >= to; i--) result.Add(i - 1);
            }
        }
        return result.Count > 0 ? result : null;
    }

    /// <summary>
    /// Computes where a PDF page goes on the paper. All rectangles are in 1/100 inch.
    /// </summary>
    public static RectangleF Layout(SizeF pagePts, RectangleF printable, PrintOptions o, bool paperLandscape, out bool rotate)
    {
        float w = pagePts.Width / 72f * 100f, h = pagePts.Height / 72f * 100f;
        bool pageLandscape = w > h;
        rotate = o.AutoRotate && pageLandscape != (printable.Width > printable.Height);
        if (rotate) (w, h) = (h, w);
        float fit = Math.Min(printable.Width / w, printable.Height / h);
        float k = o.Scaling switch
        {
            PrintScaling.Fit => fit,
            PrintScaling.ActualSize => 1f,
            _ => Math.Min(1f, fit)
        };
        float dw = w * k, dh = h * k;
        float x = o.AutoCenter ? printable.X + (printable.Width - dw) / 2 : printable.X;
        float y = o.AutoCenter ? printable.Y + (printable.Height - dh) / 2 : printable.Y;
        return new RectangleF(x, y, dw, dh);
    }

    public static PrintDocument CreateDocument(PdfiumDocument doc, PrintOptions o, string docName)
    {
        // Automated-test safety net: never let a test run reach a physical printer.
        var guard = Environment.GetEnvironmentVariable("MICROADS_TEST_PRINT_TO_FILE");
        if (!string.IsNullOrWhiteSpace(guard))
        {
            o.PrinterName = "Microsoft Print to PDF";
            o.PrintToFile = true;
            if (string.IsNullOrWhiteSpace(o.OutputFile)) o.OutputFile = guard;
        }
        var pd = new PrintDocument { DocumentName = docName };
        if (o.DriverSettings != null && o.DriverSettings.PrinterName == o.PrinterName) pd.PrinterSettings = (PrinterSettings)o.DriverSettings.Clone();
        pd.PrinterSettings.PrinterName = o.PrinterName;
        if (o.PageSettings != null) pd.DefaultPageSettings = (PageSettings)o.PageSettings.Clone();
        pd.DefaultPageSettings.PrinterSettings = pd.PrinterSettings;
        bool manualCopies = o.Copies > 1 && o.Copies > pd.PrinterSettings.MaximumCopies;
        if (!manualCopies)
        {
            pd.PrinterSettings.Copies = o.Copies;
            pd.PrinterSettings.Collate = o.Collate;
        }
        if (o.PrintToFile && !string.IsNullOrWhiteSpace(o.OutputFile))
        {
            pd.PrinterSettings.PrintToFile = true;
            pd.PrinterSettings.PrintFileName = o.OutputFile;
        }
        var sequence = new List<int>();
        if (manualCopies)
        {
            if (o.Collate) for (int c = 0; c < o.Copies; c++) sequence.AddRange(o.Pages);
            else foreach (var p in o.Pages) for (int c = 0; c < o.Copies; c++) sequence.Add(p);
        }
        else sequence.AddRange(o.Pages);

        int index = 0;
        pd.BeginPrint += (_, _) => index = 0;
        pd.QueryPageSettings += (_, e) =>
        {
            // keep orientation as chosen; auto-rotate handles landscape pages
        };
        pd.PrintPage += (_, e) =>
        {
            if (index >= sequence.Count) { e.HasMorePages = false; return; }
            int page = sequence[index];
            var g = e.Graphics!;
            float hmx = e.PageSettings.HardMarginX, hmy = e.PageSettings.HardMarginY;
            var paper = e.PageBounds; // 1/100 inch, already oriented
            // Real printable area (margins may differ per side); PrintableArea is not rotated for landscape.
            var pa = e.PageSettings.PrintableArea;
            if (e.PageSettings.Landscape) pa = new RectangleF(pa.Y, pa.X, pa.Height, pa.Width);
            var printable = new RectangleF(pa.X - hmx, pa.Y - hmy, pa.Width, pa.Height);
            if (printable.Width <= 0 || printable.Height <= 0) printable = new RectangleF(0, 0, paper.Width, paper.Height);
            var dest = Layout(doc.GetPageSize(page), printable, o, paper.Width > paper.Height, out bool rotate);
            bool preview = pd.PrintController?.IsPreview ?? false;
            int dpi = preview ? 150 : Math.Clamp(e.PageSettings.PrinterResolution.X > 0 ? e.PageSettings.PrinterResolution.X : 300, 150, 300);
            int pw = Math.Max(1, (int)(dest.Width / 100f * dpi)), ph = Math.Max(1, (int)(dest.Height / 100f * dpi));
            // limit memory for huge pages
            double maxPixels = 60_000_000;
            if ((double)pw * ph > maxPixels)
            {
                double s = Math.Sqrt(maxPixels / ((double)pw * ph));
                pw = (int)(pw * s); ph = (int)(ph * s);
            }
            using var bmp = doc.Render(page, pw, ph, rotate ? 1 : 0, null, o.Annotations, printing: true, grayscale: o.Grayscale);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.DrawImage(bmp, dest);
            index++;
            e.HasMorePages = index < sequence.Count;
        };
        return pd;
    }
}
