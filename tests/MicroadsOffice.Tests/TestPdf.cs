using System.Text;

namespace MicroadsOffice.Tests;

/// <summary>Generates small, valid PDFs (classic xref) for tests.</summary>
public static class TestPdf
{
    public static byte[] Create(int pages = 3, int[]? rotations = null, bool withEmptySigField = false, bool withLink = false,
        bool withOutline = false)
    {
        var objs = new List<string>();
        // 1 catalog, 2 pages, 3 font, then pages/content pairs
        int firstPage = 4;
        var kids = new List<string>();
        for (int i = 0; i < pages; i++) kids.Add($"{firstPage + i * 2} 0 R");
        int next = firstPage + pages * 2;
        int fieldObj = withEmptySigField ? next++ : 0;
        int linkObj = withLink ? next++ : 0;
        int outlinesObj = withOutline ? next++ : 0;
        int outlineItem = withOutline ? next++ : 0;

        var catalog = new StringBuilder("<< /Type /Catalog /Pages 2 0 R");
        if (withEmptySigField) catalog.Append($" /AcroForm << /Fields [{fieldObj} 0 R] >>");
        if (withOutline) catalog.Append($" /Outlines {outlinesObj} 0 R /PageMode /UseOutlines");
        catalog.Append(" >>");
        objs.Add(catalog.ToString());
        objs.Add($"<< /Type /Pages /Kids [{string.Join(" ", kids)}] /Count {pages} >>");
        objs.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        for (int i = 0; i < pages; i++)
        {
            int rot = rotations != null && i < rotations.Length ? rotations[i] : 0;
            var annots = new List<string>();
            if (withEmptySigField && i == 0) annots.Add($"{fieldObj} 0 R");
            if (withLink && i == 0) annots.Add($"{linkObj} 0 R");
            string annotStr = annots.Count > 0 ? $" /Annots [{string.Join(" ", annots)}]" : "";
            objs.Add($"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Rotate {rot} /Resources << /Font << /F1 3 0 R >> >> /Contents {firstPage + i * 2 + 1} 0 R{annotStr} >>");
            var content = $"BT /F1 28 Tf 72 760 Td (Microads Test Page {i + 1}) Tj ET\n" +
                          $"BT /F1 14 Tf 72 720 Td (Hop dong dich vu so {i + 1:D3} - Contract searchable text) Tj ET\n" +
                          $"BT /F1 12 Tf 72 690 Td (Lorem ipsum dolor sit amet, consectetur adipiscing elit.) Tj ET\n" +
                          "0 0.4 0.8 RG 2 w 72 100 451 60 re S\n";
            objs.Add($"<< /Length {content.Length} >>\nstream\n{content}endstream");
        }
        if (withEmptySigField)
            objs.Add($"<< /Type /Annot /Subtype /Widget /FT /Sig /T (ApproverSignature) /Rect [300 100 520 170] /F 4 /P {firstPage} 0 R >>");
        if (withLink)
            objs.Add($"<< /Type /Annot /Subtype /Link /Rect [72 755 400 790] /Border [0 0 0] /Dest [{firstPage + (pages - 1) * 2} 0 R /XYZ 0 842 0] >>");
        if (withOutline)
        {
            objs.Add($"<< /Type /Outlines /First {outlineItem} 0 R /Last {outlineItem} 0 R /Count 1 >>");
            objs.Add($"<< /Title (Chapter 2) /Parent {outlinesObj} 0 R /Dest [{firstPage + 2} 0 R /Fit] >>");
        }

        var sb = new MemoryStream();
        void W(string s) { var b = Encoding.Latin1.GetBytes(s); sb.Write(b, 0, b.Length); }
        W("%PDF-1.7\n%âãÏÓ\n");
        var offsets = new List<long>();
        for (int i = 0; i < objs.Count; i++)
        {
            offsets.Add(sb.Length);
            W($"{i + 1} 0 obj\n{objs[i]}\nendobj\n");
        }
        long xref = sb.Length;
        W($"xref\n0 {objs.Count + 1}\n0000000000 65535 f\r\n");
        foreach (var o in offsets) W($"{o:D10} 00000 n\r\n");
        W($"trailer\n<< /Size {objs.Count + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        return sb.ToArray();
    }
}
