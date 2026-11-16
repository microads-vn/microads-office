using System.Drawing;
using System.Drawing.Imaging;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using MicroadsOffice.Word.IO;
using MicroadsOffice.Word.Model;
using WRun = MicroadsOffice.Word.Model.Run;

namespace MicroadsOffice.Word.Tests;

/// <summary>DOCX / RTF / HTML / TXT / PDF reading and writing.</summary>
public static class IoTests
{
    public static byte[] Png(int w, int h, Color c)
    {
        using var b = new Bitmap(w, h);
        using (var g = Graphics.FromImage(b)) { g.Clear(c); g.FillEllipse(Brushes.White, 4, 4, w - 8, h - 8); }
        using var ms = new MemoryStream();
        b.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    /// <summary>A document using every feature the writer supports.</summary>
    public static Document Rich()
    {
        var d = Document.CreateBlank("Times New Roman", 13);
        d.Blocks.Clear();
        var n = d.Styles[StyleIds.Normal];
        d.Blocks.Add(new Paragraph("Tiêu đề chính", d.Styles[StyleIds.Heading1].Char, d.Styles[StyleIds.Heading1].Para));
        var p = new Paragraph("", n.Char, n.Para with { Align = ParaAlign.Justify, IndentLeft = 20, IndentRight = 10, FirstLine = 28, SpaceBefore = 6, SpaceAfter = 12, LineSpacing = 1.5f, Shading = Color.FromArgb(255, 242, 204), Border = true });
        p.Runs.Add(new WRun("Đậm ", n.Char with { Bold = true }));
        p.Runs.Add(new WRun("nghiêng ", n.Char with { Italic = true, Color = Color.FromArgb(192, 0, 0) }));
        p.Runs.Add(new WRun("gạch đôi ", n.Char with { Underline = UnderlineStyle.Double, FontName = "Arial", Size = 16 }));
        p.Runs.Add(new WRun("ngang ", n.Char with { Strike = true, Highlight = Color.Yellow }));
        p.Runs.Add(new WRun("x", n.Char));
        p.Runs.Add(new WRun("2", n.Char with { VertAlign = VertAlign.Superscript }));
        p.Runs.Add(new WRun(" H", n.Char));
        p.Runs.Add(new WRun("2", n.Char with { VertAlign = VertAlign.Subscript }));
        p.Runs.Add(new WRun("O\tTab" + WRun.LineBreak + "dòng mới ", n.Char));
        p.Runs.Add(new WRun("microads.vn", n.Char with { Link = "https://microads.vn", Color = Color.FromArgb(5, 99, 193), Underline = UnderlineStyle.Single }));
        d.Blocks.Add(p);
        for (int i = 0; i < 3; i++) d.Blocks.Add(new Paragraph("Mục " + (i + 1), n.Char, n.Para with { List = ListType.Decimal, ListId = 3, ListLevel = i == 1 ? 1 : 0, IndentLeft = 36 + (i == 1 ? 18 : 0), FirstLine = -18 }));
        d.Blocks.Add(new Paragraph("Chấm tròn", n.Char, n.Para with { List = ListType.Bullet, ListId = 4, IndentLeft = 36, FirstLine = -18 }));
        var t = Table.Create(3, 3, 400, n.Char);
        t.Rows[0].Cells[0].Paragraphs[0].Runs.Add(new WRun("Ô gộp", n.Char with { Bold = true }));
        t.Rows[0].Cells[0].Shading = Color.FromArgb(221, 235, 247);
        t.Rows[0].Cells[0].GridSpan = 2; t.Rows[0].Cells.RemoveAt(1);
        t.Rows[1].Cells[2].VMerge = VMerge.Restart; t.Rows[2].Cells[2].VMerge = VMerge.Continue;
        t.Rows[1].Cells[0].Paragraphs[0].Runs.Add(new WRun("A2", n.Char));
        t.Rows[2].Cells[1].VAlign = CellVAlign.Center;
        d.Blocks.Add(t);
        var ip = new Paragraph("", n.Char, n.Para with { Align = ParaAlign.Center });
        ip.Runs.Add(new WRun(WRun.ObjectChar.ToString(), n.Char, new ImageObject(Png(120, 80, Color.SteelBlue), "png", 90, 60)));
        d.Blocks.Add(ip);
        d.Blocks.Add(new Paragraph("Trang sau" + WRun.PageBreak + "sau ngắt trang", n.Char, n.Para with { LineRule = LineRule.Exact, LineSpacing = 18 }));
        d.Header.Add(new Paragraph("Đầu trang Microads", n.Char with { Size = 10 }, new ParaProps { Align = ParaAlign.Right, SpaceAfter = 0 }));
        var f = new Paragraph("Trang ", n.Char with { Size = 10 }, new ParaProps { Align = ParaAlign.Center, SpaceAfter = 0 });
        f.Runs.Add(new WRun(WRun.ObjectChar.ToString(), n.Char with { Size = 10 }, new FieldObject(FieldKind.Page)));
        f.Runs.Add(new WRun(" / ", n.Char with { Size = 10 }));
        f.Runs.Add(new WRun(WRun.ObjectChar.ToString(), n.Char with { Size = 10 }, new FieldObject(FieldKind.NumPages)));
        d.Footer.Add(f);
        d.Watermark = "BẢN NHÁP";
        d.PageColor = Color.FromArgb(250, 250, 240);
        d.Page = new PageSetup { Width = 841.9f, Height = 595.3f, MarginTop = 50, MarginBottom = 40, MarginLeft = 70, MarginRight = 30, PaperName = "A4" };
        d.Title = "Kiểm thử"; d.Author = "Microads QA"; d.Subject = "DOCX"; d.Keywords = "word,test";
        d.StructureChanged();
        return d;
    }

    public static void Run(Runner r)
    {
        var dir = Path.Combine(T.Out, "io");
        Directory.CreateDirectory(dir);

        r.Test("IO01 DOCX round-trip keeps text, formatting, lists, tables, image, sections", () =>
        {
            var d = Rich();
            var path = Path.Combine(dir, "roundtrip.docx");
            DocxWriter.Write(d, path);
            // every part must be well-formed XML
            using (var z = ZipFile.OpenRead(path))
                foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels")))
                    using (var s = e.Open()) XDocument.Load(s);
            var b = DocxReader.Read(path);
            var ps = b.Paragraphs;
            T.Eq("Tiêu đề chính", ps[0].Text);
            T.Eq(1, ps[0].Props.OutlineLevel, "heading level kept");
            var p = b.Blocks.OfType<Paragraph>().ElementAt(1);
            T.Eq(ParaAlign.Justify, p.Props.Align);
            T.Near(20, p.Props.IndentLeft, 0.1f, "indent left"); T.Near(28, p.Props.FirstLine, 0.1f, "first line");
            T.Near(1.5f, p.Props.LineSpacing, 0.01f, "line spacing"); T.Near(12, p.Props.SpaceAfter, 0.1f, "space after");
            T.True(p.Props.Border && p.Props.Shading?.ToArgb() == Color.FromArgb(255, 242, 204).ToArgb(), "border + shading");
            T.True(p.Runs[0].Props.Bold && p.Runs[0].Text == "Đậm ", "bold run");
            T.True(p.Runs.Any(x => x.Props.Italic && x.Props.Color?.R == 192), "italic red");
            T.True(p.Runs.Any(x => x.Props.Underline == UnderlineStyle.Double && x.Props.FontName == "Arial" && x.Props.Size == 16), "double underline Arial 16");
            T.True(p.Runs.Any(x => x.Props.Strike && x.Props.Highlight != null), "strike + highlight");
            T.True(p.Runs.Any(x => x.Props.VertAlign == VertAlign.Superscript) && p.Runs.Any(x => x.Props.VertAlign == VertAlign.Subscript), "super/sub");
            T.True(p.Text.Contains('\t') && p.Text.Contains(WRun.LineBreak), "tab and line break");
            T.True(p.Runs.Any(x => x.Props.Link == "https://microads.vn" && x.Text == "microads.vn"), "hyperlink");
            var lists = b.Paragraphs.Where(x => x.Props.List != ListType.None).ToList();
            T.Eq(4, lists.Count, "list paragraphs");
            T.Eq(ListType.Decimal, lists[0].Props.List); T.Eq(1, lists[1].Props.ListLevel); T.Eq(ListType.Bullet, lists[3].Props.List);
            var t = b.Blocks.OfType<Table>().Single();
            T.Eq(3, t.Rows.Count); T.Eq(3, t.ColumnWidths.Count);
            T.Eq(2, t.Rows[0].Cells[0].GridSpan, "gridSpan");
            T.Eq(VMerge.Restart, t.Rows[1].Cells[2].VMerge); T.Eq(VMerge.Continue, t.Rows[2].Cells[2].VMerge);
            T.True(t.Rows[0].Cells[0].Shading?.ToArgb() == Color.FromArgb(221, 235, 247).ToArgb(), "cell shading");
            T.Eq(CellVAlign.Center, t.Rows[2].Cells[1].VAlign);
            T.True(t.Borders, "table borders");
            var img = b.Paragraphs.SelectMany(x => x.Runs).Select(x => x.Object).OfType<ImageObject>().Single();
            T.Near(90, img.Width, 0.1f, "image width"); T.Near(60, img.Height, 0.1f, "image height");
            T.True(b.Paragraphs.Any(x => x.Text.Contains(WRun.PageBreak)), "page break");
            T.True(b.Paragraphs.Any(x => x.Props.LineRule == LineRule.Exact && Math.Abs(x.Props.LineSpacing - 18) < 0.1), "exact line spacing");
            T.Eq("Đầu trang Microads", b.Header[0].Text);
            T.Eq(ParaAlign.Right, b.Header[0].Props.Align);
            T.Eq(2, b.Footer[0].Runs.Count(x => x.Object is FieldObject), "PAGE + NUMPAGES fields");
            T.Eq("BẢN NHÁP", b.Watermark);
            T.True(b.PageColor?.ToArgb() == Color.FromArgb(250, 250, 240).ToArgb(), "page colour");
            T.True(b.Page.Landscape, "landscape"); T.Near(70, b.Page.MarginLeft, 0.1f, "margin left"); T.Near(50, b.Page.MarginTop, 0.1f, "margin top");
            T.Eq("Kiểm thử", b.Title); T.Eq("Microads QA", b.Author);
        });

        r.Test("IO02 read a Word-authored style DOCX (styles, theme fonts, numbering, complex fields, sdt)", () =>
        {
            var path = Path.Combine(dir, "wordstyle.docx");
            HandCrafted(path);
            var d = DocxReader.Read(path);
            var ps = d.Paragraphs;
            T.Eq("Báo cáo năm", ps[0].Text);
            T.Eq(StyleIds.Heading1, ps[0].Props.Style, "Heading1 detected by name");
            T.True(ps[0].Runs[0].Props.Bold && Math.Abs(ps[0].Runs[0].Props.Size - 16) < 0.01, "heading style formatting resolved");
            T.Eq("Calibri", ps[1].Runs[0].Props.FontName, "theme minor font");
            T.Near(11, ps[1].Runs[0].Props.Size, 0.01f, "docDefaults size");
            T.Near(8, ps[1].Props.SpaceAfter, 0.01f, "docDefaults spacing");
            T.Eq(ListType.Bullet, ps[2].Props.List, "numbering bullet");
            T.Eq(ListType.Decimal, ps[3].Props.List, "numbering decimal");
            T.True(ps[4].Runs.Any(x => x.Object is FieldObject { Kind: FieldKind.Page }), "complex PAGE field");
            T.False(ps[4].Text.Contains("PAGE"), "instruction text hidden");
            T.True(ps[5].Text.Contains("nội dung trong sdt"), "sdt content unwrapped");
            T.True(ps[5].Runs.Any(x => x.Props.Link == "https://example.com"), "hyperlink relation");
            var t = d.Blocks.OfType<Table>().Single();
            T.Eq(2, t.Rows.Count); T.True(t.Borders, "TableGrid style borders");
            T.Near(144, t.ColumnWidths[0], 0.1f, "grid width");
            T.Eq(VMerge.Continue, t.Rows[1].Cells[0].VMerge);
            T.Near(PageSetup.Mm(30), d.Page.MarginLeft, 0.2f, "sectPr margin");
            T.Eq("Header text", d.Header[0].Text);
        });

        r.Test("IO03 TXT write/read with Vietnamese (UTF-8)", () =>
        {
            var d = Rich();
            var path = Path.Combine(dir, "out.txt");
            TextIO.Write(d, path);
            var txt = File.ReadAllText(path);
            T.True(txt.Contains("Tiêu đề chính") && txt.Contains("1. Mục 1") && txt.Contains("Ô gộp"), "text content");
            var back = WordFiles.Open(path);
            T.True(back.Paragraphs.Any(p => p.Text == "Tiêu đề chính"), "re-read txt");
        });

        r.Test("IO04 RTF write and read back through the rich edit control", () =>
        {
            var d = Rich();
            var path = Path.Combine(dir, "out.rtf");
            WordFiles.Save(d, path);
            var rtf = File.ReadAllText(path);
            T.True(rtf.StartsWith("{\\rtf1") && rtf.Contains("\\trowd") && rtf.Contains("\\pngblip"), "rtf structure");
            var back = WordFiles.Open(path);
            var p = back.Paragraphs.FirstOrDefault(x => x.Text.StartsWith("Đậm"));
            T.True(p != null, "paragraph found: " + string.Join("|", back.Paragraphs.Take(4).Select(x => x.Text)));
            T.True(p!.Runs[0].Props.Bold, "bold survives RTF");
            T.True(back.Paragraphs.Any(x => x.Text.Contains("Tiêu đề chính")), "vietnamese text survives RTF");
        });

        r.Test("IO05 HTML export", () =>
        {
            var path = Path.Combine(dir, "out.html");
            WordFiles.Save(Rich(), path);
            var html = File.ReadAllText(path);
            T.True(html.Contains("<h1") && html.Contains("Tiêu đề chính") && html.Contains("<table") && html.Contains("colspan=\"2\"") && html.Contains("data:image/png;base64") && html.Contains("href=\"https://microads.vn\""), "html content");
        });

        r.Test("IO06 PDF export (Microsoft Print to PDF, text kept)", () =>
        {
            var path = Path.Combine(dir, "out.pdf");
            WordFiles.ExportPdf(Rich(), path);
            var bytes = File.ReadAllBytes(path);
            T.True(bytes.Length > 1500 && Encoding.ASCII.GetString(bytes, 0, 5) == "%PDF-", "valid pdf header, size " + bytes.Length);
        });

        r.Test("IO07 PDF export fallback (image PDF) works without the PDF printer", () =>
        {
            var path = Path.Combine(dir, "out-image.pdf");
            WordFiles.ExportPdf(Rich(), path, forceImage: true);
            var s = Encoding.Latin1.GetString(File.ReadAllBytes(path));
            T.True(s.StartsWith("%PDF-") && s.Contains("/DCTDecode") && s.TrimEnd().EndsWith("%%EOF"), "image pdf structure");
            T.True(System.Text.RegularExpressions.Regex.Matches(s, "/Type /Page ").Count >= 2, "two pages (page break)");
        });

        r.Test("IO08 .doc gives a friendly message, templates build and save", () =>
        {
            var doc = Path.Combine(dir, "old.doc");
            File.WriteAllBytes(doc, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
            try { WordFiles.Open(doc); T.True(false, "should throw"); }
            catch (NotSupportedException ex) { T.True(ex.Message.Contains(".docx"), "message suggests .docx"); }
            foreach (var t in Templates.All)
            {
                var d = Templates.Create(t.Id);
                var p = Path.Combine(dir, "tpl-" + t.Id + ".docx");
                DocxWriter.Write(d, p);
                var b = DocxReader.Read(p);
                T.Eq(d.PlainText.Replace("\n", ""), b.PlainText.Replace("\n", ""), "template text round-trip " + t.Id);
                using var thumb = t.Thumbnail(120, 160);
                T.True(thumb.Width == 120, "thumbnail");
            }
            var cv = Templates.Create("congvan");
            T.Near(PageSetup.Mm(30), cv.Page.MarginLeft, 0.1f, "NĐ30 left margin 30mm");
            T.True(cv.PlainText.Contains("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") && cv.PlainText.Contains("Độc lập - Tự do - Hạnh phúc"), "quốc hiệu, tiêu ngữ");
        });
    }

    private static void HandCrafted(string path)
    {
        const string W = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>" +
                "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>" +
                "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
                "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/></Relationships>",
            ["word/theme/theme1.xml"] = "<?xml version=\"1.0\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>",
            ["word/styles.xml"] = $"<?xml version=\"1.0\"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>" +
                "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
                "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:outlineLvl w:val=\"0\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>" +
                "<w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/><w:left w:val=\"single\" w:sz=\"4\"/><w:bottom w:val=\"single\" w:sz=\"4\"/><w:right w:val=\"single\" w:sz=\"4\"/><w:insideH w:val=\"single\" w:sz=\"4\"/><w:insideV w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPr></w:style></w:styles>",
            ["word/numbering.xml"] = $"<?xml version=\"1.0\"?><w:numbering {W}><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum>" +
                "<w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num><w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>",
            ["word/header1.xml"] = $"<?xml version=\"1.0\"?><w:hdr {W}><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>",
            ["word/document.xml"] = $"<?xml version=\"1.0\"?><w:document {W}><w:body>" +
                "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Báo cáo năm</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Đoạn thường </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>nghiêng</w:t></w:r></w:p>" +
                "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Gạch đầu dòng</w:t></w:r></w:p>" +
                "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>Đánh số</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Trang </w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> PAGE \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>" +
                "<w:sdt><w:sdtContent><w:p><w:r><w:t xml:space=\"preserve\">nội dung trong sdt </w:t></w:r><w:hyperlink r:id=\"rId4\"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt>" +
                "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2880\"/><w:gridCol w:w=\"2880\"/></w:tblGrid>" +
                "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>" +
                "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                "<w:p/><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId5\"/><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:right=\"850\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>"
        };
        if (File.Exists(path)) File.Delete(path);
        using var z = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var (name, xml) in parts)
        {
            var e = z.CreateEntry(name);
            using var w = new StreamWriter(e.Open(), new UTF8Encoding(false));
            w.Write(xml);
        }
    }
}

