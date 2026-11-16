using System.Drawing.Printing;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.IO;

/// <summary>Opens / saves documents by file extension.</summary>
public static class WordFiles
{
    public static string OpenFilter => Loc.T("Tài liệu Word", "Word documents") + " (*.docx;*.docm;*.dotx;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.rtf;*.txt;*.doc|" +
                                       "Word (*.docx)|*.docx;*.docm;*.dotx|RTF (*.rtf)|*.rtf|" + Loc.T("Văn bản thuần", "Plain text") + " (*.txt)|*.txt|" +
                                       Loc.T("Tất cả tệp", "All files") + " (*.*)|*.*";

    public static string SaveFilter => Loc.T("Tài liệu Word", "Word document") + " (*.docx)|*.docx|" +
                                       "PDF (*.pdf)|*.pdf|" +
                                       Loc.T("Định dạng RTF", "Rich Text Format") + " (*.rtf)|*.rtf|" +
                                       Loc.T("Trang web", "Web page") + " (*.html)|*.html|" +
                                       Loc.T("Văn bản thuần", "Plain text") + " (*.txt)|*.txt";

    public static Document Open(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".doc":
                throw new NotSupportedException(Loc.T(
                    "Tệp .doc (Word 97-2003) dùng định dạng nhị phân cũ chưa được hỗ trợ.\nHãy mở tệp bằng Word/WPS và lưu lại thành .docx, hoặc lưu dưới dạng .rtf, rồi mở bằng Microads Office Word.",
                    "The .doc format (Word 97-2003) is an old binary format that is not supported.\nPlease save it as .docx (or .rtf) and open it again in Microads Office Word."));
            case ".rtf":
                return RtfReader.Read(File.ReadAllText(path, System.Text.Encoding.Latin1));
            case ".txt":
            case ".text":
            case ".csv":
                return TextIO.Read(path);
            default:
                // .docx / .docm / .dotx and unknown extensions that are zip packages
                using (var fs = File.OpenRead(path))
                {
                    var sig = new byte[4];
                    int n = fs.Read(sig, 0, 4);
                    if (n == 4 && sig[0] == 0xD0 && sig[1] == 0xCF)
                        throw new NotSupportedException(Loc.T("Tệp này là định dạng Word 97-2003 (.doc) chưa được hỗ trợ. Hãy lưu lại thành .docx.", "This is a Word 97-2003 (.doc) file which is not supported. Please save it as .docx."));
                    if (n == 4 && sig[0] == '{' && sig[1] == '\\') { fs.Close(); return RtfReader.Read(File.ReadAllText(path, System.Text.Encoding.Latin1)); }
                    if (!(n >= 2 && sig[0] == 'P' && sig[1] == 'K')) { fs.Close(); return TextIO.Read(path); }
                    fs.Position = 0;
                    return DocxReader.Read(fs);
                }
        }
    }

    /// <summary>Saves by extension. PDF uses the print pipeline.</summary>
    public static void Save(Document d, string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".rtf": File.WriteAllText(path, RtfWriter.Write(d), System.Text.Encoding.ASCII); break;
            case ".txt": TextIO.Write(d, path); break;
            case ".htm":
            case ".html": File.WriteAllText(path, HtmlWriter.Write(d), new System.Text.UTF8Encoding(false)); break;
            case ".pdf": ExportPdf(d, path); break;
            default: DocxWriter.Write(d, path); break;
        }
    }

    public static void ExportPdf(Document d, string path, bool forceImage = false)
    {
        using var m = new TextMeasurer();
        var layout = new DocLayout(d, m);
        using var pd = new WordPrintDocument(layout);
        pd.DocumentName = string.IsNullOrEmpty(d.Title) ? Path.GetFileNameWithoutExtension(path) : d.Title;
        MicroadsOffice.UI.PdfExporter.Export(pd, path, forceImage);
    }

    /// <summary>True when the format keeps all formatting (only .docx does).</summary>
    public static bool IsNative(string? path) => path != null && Path.GetExtension(path).ToLowerInvariant() is ".docx" or ".docm";
}

/// <summary>Prints laid-out pages (also used for PDF export and print preview).</summary>
public sealed class WordPrintDocument : PrintDocument
{
    private readonly DocLayout _layout;
    private List<int> _pages = new();
    private int _next;

    /// <summary>Zero-based page indices to print; null = all.</summary>
    public List<int>? PageList { get; set; }

    public WordPrintDocument(DocLayout layout)
    {
        _layout = layout;
        var ps = layout.Doc.Page;
        DefaultPageSettings.Landscape = ps.Landscape;
        int w = (int)Math.Round(Math.Min(ps.Width, ps.Height) / 72f * 100), h = (int)Math.Round(Math.Max(ps.Width, ps.Height) / 72f * 100);
        try
        {
            var match = PrinterSettings.IsValid ? PrinterSettings.PaperSizes.Cast<PaperSize>().FirstOrDefault(p => Math.Abs(p.Width - w) <= 3 && Math.Abs(p.Height - h) <= 3) : null;
            DefaultPageSettings.PaperSize = match ?? new PaperSize(ps.PaperName, w, h);
        }
        catch { DefaultPageSettings.PaperSize = new PaperSize(ps.PaperName, w, h); }
        DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
        OriginAtMargins = false;
    }

    protected override void OnBeginPrint(PrintEventArgs e)
    {
        base.OnBeginPrint(e);
        var all = Enumerable.Range(0, _layout.Pages.Count);
        if (PageList != null) _pages = PageList.Where(i => i >= 0 && i < _layout.Pages.Count).ToList();
        else if (PrinterSettings.PrintRange == PrintRange.SomePages)
            _pages = all.Where(i => i + 1 >= PrinterSettings.FromPage && i + 1 <= PrinterSettings.ToPage).ToList();
        else _pages = all.ToList();
        if (_pages.Count == 0) _pages.Add(0);
        _next = 0;
    }

    protected override void OnPrintPage(PrintPageEventArgs e)
    {
        base.OnPrintPage(e);
        var g = e.Graphics!;
        g.ResetTransform();
        g.PageUnit = GraphicsUnit.Point;
        // the graphics origin is the printable area; move back to the physical page corner
        if (!e.PageSettings.PrinterSettings.PrintToFile || true)
        {
            float hx = e.PageSettings.HardMarginX / 100f * 72f, hy = e.PageSettings.HardMarginY / 100f * 72f;
            if (PrintController is not PreviewPrintController) g.TranslateTransform(-hx, -hy);
        }
        // rotate when the printer gives a portrait page for a landscape document (and vice versa)
        var ps = _layout.Doc.Page;
        float pw = e.PageBounds.Width / 100f * 72f, ph = e.PageBounds.Height / 100f * 72f;
        if ((ps.Width > ps.Height) != (pw > ph) && Math.Abs(pw - ph) > 10)
        {
            g.TranslateTransform(pw, 0);
            g.RotateTransform(90);
        }
        PageRenderer.DrawPage(g, _layout, _pages[_next], new RenderOptions { Print = true });
        _next++;
        e.HasMorePages = _next < _pages.Count;
    }

    /// <summary>Parses "1-3, 5, 8-" (1-based) into zero-based page indices.</summary>
    public static List<int>? ParseRange(string text, int pageCount)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var list = new SortedSet<int>();
        foreach (var part in text.Split(',', ';'))
        {
            var p = part.Trim();
            if (p.Length == 0) continue;
            var dash = p.IndexOf('-');
            if (dash < 0)
            {
                if (!int.TryParse(p, out var n) || n < 1 || n > pageCount) return [];
                list.Add(n - 1);
                continue;
            }
            var a = p[..dash].Trim(); var b = p[(dash + 1)..].Trim();
            int from = a.Length == 0 ? 1 : int.TryParse(a, out var x) ? x : -1;
            int to = b.Length == 0 ? pageCount : int.TryParse(b, out var y) ? y : -1;
            if (from < 1 || to < from || from > pageCount) return [];
            for (int i = from; i <= Math.Min(to, pageCount); i++) list.Add(i - 1);
        }
        return list.ToList();
    }
}
