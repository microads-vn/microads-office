using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.Text;

namespace MicroadsOffice.UI;

/// <summary>
/// Exports any <see cref="PrintDocument"/> (Word pages, Excel sheets) to PDF.
/// Preferred path: the built-in "Microsoft Print to PDF" printer (text stays selectable/searchable,
/// fonts are embedded). Fallback when that printer was removed: pages rendered at 200 dpi into an image PDF.
/// </summary>
public static class PdfExporter
{
    public const string MsPdfPrinter = "Microsoft Print to PDF";

    public static bool HasPdfPrinter()
    {
        try { return PrinterSettings.InstalledPrinters.Cast<string>().Any(p => string.Equals(p, MsPdfPrinter, StringComparison.OrdinalIgnoreCase)); }
        catch { return false; }
    }

    /// <summary>Writes the document to <paramref name="path"/>. Throws on failure.</summary>
    public static void Export(PrintDocument doc, string path, bool forceImage = false)
    {
        if (File.Exists(path)) File.Delete(path);
        if (!forceImage && HasPdfPrinter())
        {
            try
            {
                var ps = new PrinterSettings { PrinterName = MsPdfPrinter, PrintToFile = true, PrintFileName = path };
                if (ps.IsValid)
                {
                    var landscape = doc.DefaultPageSettings.Landscape;
                    var paper = doc.DefaultPageSettings.PaperSize;
                    var margins = doc.DefaultPageSettings.Margins;
                    doc.PrinterSettings = ps;
                    doc.DefaultPageSettings.Landscape = landscape;
                    var match = ps.PaperSizes.Cast<PaperSize>().FirstOrDefault(p => Math.Abs(p.Width - paper.Width) <= 3 && Math.Abs(p.Height - paper.Height) <= 3);
                    doc.DefaultPageSettings.PaperSize = match ?? paper;
                    doc.DefaultPageSettings.Margins = margins;
                    doc.PrintController = new StandardPrintController();
                    doc.Print();
                    if (WaitForFile(path, TimeSpan.FromSeconds(60))) return;
                }
            }
            catch { /* fall back to image PDF */ }
        }
        ExportAsImages(doc, path);
    }

    private static bool WaitForFile(string path, TimeSpan timeout)
    {
        var until = DateTime.UtcNow + timeout;
        long last = -1;
        int stable = 0;
        while (DateTime.UtcNow < until)
        {
            try
            {
                if (File.Exists(path))
                {
                    var len = new FileInfo(path).Length;
                    if (len > 0 && len == last) { if (++stable >= 3) { using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None); return true; } }
                    else stable = 0;
                    last = len;
                }
            }
            catch (IOException) { stable = 0; }
            Thread.Sleep(200);
            Application.DoEvents();
        }
        return File.Exists(path) && new FileInfo(path).Length > 0;
    }

    /// <summary>Renders pages through a preview controller and writes them as JPEG images into a PDF.</summary>
    public static void ExportAsImages(PrintDocument doc, string path, int dpi = 200)
    {
        var controller = new PreviewPrintController { UseAntiAlias = true };
        doc.PrintController = controller;
        doc.Print();
        var pages = controller.GetPreviewPageInfo();
        var images = new List<(byte[] Jpeg, int W, int H, float PtW, float PtH)>();
        var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
        using var ep = new EncoderParameters(1);
        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);
        foreach (var p in pages)
        {
            // PhysicalSize is in 1/100 inch
            float ptW = p.PhysicalSize.Width * 0.72f, ptH = p.PhysicalSize.Height * 0.72f;
            int pw = (int)(p.PhysicalSize.Width / 100f * dpi), ph = (int)(p.PhysicalSize.Height / 100f * dpi);
            using var bmp = new Bitmap(pw, ph);
            bmp.SetResolution(dpi, dpi);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.DrawImage(p.Image, 0, 0, pw, ph);
            }
            using var ms = new MemoryStream();
            bmp.Save(ms, codec, ep);
            images.Add((ms.ToArray(), pw, ph, ptW, ptH));
            p.Image.Dispose();
        }
        WriteImagePdf(path, images);
    }

    private static void WriteImagePdf(string path, List<(byte[] Jpeg, int W, int H, float PtW, float PtH)> pages)
    {
        using var fs = File.Create(path);
        var offsets = new List<long>();
        void W(string s) { var b = Encoding.Latin1.GetBytes(s); fs.Write(b, 0, b.Length); }
        string F(float v) => v.ToString("0.###", CultureInfo.InvariantCulture);
        W("%PDF-1.4\n%âãÏÓ\n");
        int n = pages.Count;
        // objects: 1 catalog, 2 pages, then per page: page, content, image
        int ObjPage(int i) => 3 + i * 3;
        void Begin(int id) { while (offsets.Count < id) offsets.Add(0); offsets[id - 1] = fs.Position; W($"{id} 0 obj\n"); }
        Begin(1); W("<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Begin(2); W("<< /Type /Pages /Count " + n + " /Kids [" + string.Join(" ", Enumerable.Range(0, n).Select(i => $"{ObjPage(i)} 0 R")) + "] >>\nendobj\n");
        W("");
        for (int i = 0; i < n; i++)
        {
            var (jpeg, w, h, ptW, ptH) = pages[i];
            int pg = ObjPage(i), ct = pg + 1, im = pg + 2;
            Begin(pg);
            W($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {F(ptW)} {F(ptH)}] /Resources << /XObject << /Im0 {im} 0 R >> >> /Contents {ct} 0 R >>\nendobj\n");
            var content = $"q {F(ptW)} 0 0 {F(ptH)} 0 0 cm /Im0 Do Q";
            Begin(ct); W($"<< /Length {content.Length} >>\nstream\n{content}\nendstream\nendobj\n");
            Begin(im); W($"<< /Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {jpeg.Length} >>\nstream\n");
            fs.Write(jpeg, 0, jpeg.Length);
            W("\nendstream\nendobj\n");
        }
        long xref = fs.Position;
        W($"xref\n0 {offsets.Count + 1}\n0000000000 65535 f \n");
        foreach (var o in offsets) W($"{o:D10} 00000 n \n");
        W($"trailer\n<< /Size {offsets.Count + 1} /Root 1 0 R /Info << /Producer (Microads Office - microads.vn) >> >>\nstartxref\n{xref}\n%%EOF\n");
    }
}
