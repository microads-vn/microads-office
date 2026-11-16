using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.IO;

/// <summary>Writes a Word 2007+ (.docx, Office Open XML) package.</summary>
public static class DocxWriter
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    private sealed class Ctx
    {
        public readonly List<(string Id, string Type, string Target, bool External)> Rels = new();
        public readonly List<(string Name, byte[] Data)> Media = new();
        public readonly Dictionary<int, (ListType[] Levels, int NumId)> Lists = new();
        public int NextRel = 10;
        public int NextImage = 1;
        public int DocPr = 1;

        public string AddRel(string type, string target, bool external = false)
        {
            var id = "rId" + NextRel++;
            Rels.Add((id, type, target, external));
            return id;
        }
    }

    public static void Write(Document d, string path)
    {
        var tmp = path + ".tmp";
        using (var fs = File.Create(tmp)) Write(d, fs);
        File.Move(tmp, path, true);
    }

    public static void Write(Document d, Stream output)
    {
        var ctx = new Ctx();
        CollectLists(d, ctx);
        string? headerRel = null, footerRel = null;
        string? headerXml = null, footerXml = null;
        if (d.Header.Count > 0 || !string.IsNullOrWhiteSpace(d.Watermark))
        {
            headerXml = Part("hdr", d.Header, ctx, d.Watermark);
            headerRel = ctx.AddRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header1.xml");
        }
        if (d.Footer.Count > 0)
        {
            footerXml = Part("ftr", d.Footer, ctx, null);
            footerRel = ctx.AddRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer1.xml");
        }
        var body = new StringBuilder();
        foreach (var b in d.Blocks) WriteBlock(body, b, ctx);
        body.Append(SectPr(d, headerRel, footerRel));
        var docXml = new StringBuilder();
        docXml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        docXml.Append($"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
                      "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">");
        if (d.PageColor is Color pc) docXml.Append($"<w:background w:color=\"{Editing.Hex(pc)}\"/>");
        docXml.Append("<w:body>").Append(body).Append("</w:body></w:document>");

        using var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        void Add(string name, string content)
        {
            var e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using var s = new StreamWriter(e.Open(), new UTF8Encoding(false));
            s.Write(content);
        }
        var ct = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/><Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>" +
            "<Default Extension=\"gif\" ContentType=\"image/gif\"/><Default Extension=\"bmp\" ContentType=\"image/bmp\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
            "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>" +
            "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>" +
            "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
            "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        if (headerXml != null) ct.Append("<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>");
        if (footerXml != null) ct.Append("<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>");
        ct.Append("</Types>");
        Add("[Content_Types].xml", ct.ToString());
        Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
            "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>");
        Add("docProps/core.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
            $"<dc:title>{X(d.Title)}</dc:title><dc:subject>{X(d.Subject)}</dc:subject><dc:creator>{X(d.Author)}</dc:creator><cp:keywords>{X(d.Keywords)}</cp:keywords>" +
            $"<cp:lastModifiedBy>Microads Office</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">{d.Created.ToUniversalTime():yyyy-MM-ddTHH:mm:ssZ}</dcterms:created>" +
            $"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:modified></cp:coreProperties>");
        Add("docProps/app.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
            "<Application>Microads Office Word (microads.vn)</Application></Properties>");
        Add("word/styles.xml", Styles(d));
        Add("word/settings.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:settings xmlns:w=\"{W}\">" +
            (d.PageColor != null ? "<w:displayBackgroundShape/>" : "") + "<w:defaultTabStop w:val=\"720\"/><w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat></w:settings>");
        Add("word/numbering.xml", Numbering(ctx));
        Add("word/document.xml", docXml.ToString());
        if (headerXml != null) Add("word/header1.xml", headerXml);
        if (footerXml != null) Add("word/footer1.xml", footerXml);
        var rels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>" +
            "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>");
        // header/footer parts share the document relationships for images and links (written below as part rels too)
        foreach (var (id, type, target, ext) in ctx.Rels)
            rels.Append($"<Relationship Id=\"{id}\" Type=\"{type}\" Target=\"{X(target)}\"{(ext ? " TargetMode=\"External\"" : "")}/>");
        rels.Append("</Relationships>");
        Add("word/_rels/document.xml.rels", rels.ToString());
        // header/footer relationship parts (images/links used in them reuse the same ids)
        var partRels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        foreach (var (id, type, target, ext) in ctx.Rels.Where(r => !r.Type.EndsWith("/header") && !r.Type.EndsWith("/footer")))
            partRels.Append($"<Relationship Id=\"{id}\" Type=\"{type}\" Target=\"{X(target)}\"{(ext ? " TargetMode=\"External\"" : "")}/>");
        partRels.Append("</Relationships>");
        if (headerXml != null) Add("word/_rels/header1.xml.rels", partRels.ToString());
        if (footerXml != null) Add("word/_rels/footer1.xml.rels", partRels.ToString());
        foreach (var (name, data) in ctx.Media)
        {
            var e = zip.CreateEntry("word/media/" + name, CompressionLevel.NoCompression);
            using var s = e.Open();
            s.Write(data);
        }
    }

    private static string X(string? s) => SecurityElement.Escape(s ?? "") ?? "";
    private static string Tw(float pt) => ((int)Math.Round(pt * 20)).ToString(Inv);
    private static long Emu(float pt) => (long)Math.Round(pt * 12700);

    private static void CollectLists(Document d, Ctx ctx)
    {
        int num = 1;
        foreach (var p in d.Paragraphs.Concat(d.Header).Concat(d.Footer))
        {
            var pp = p.Props;
            if (pp.List == ListType.None) continue;
            if (!ctx.Lists.TryGetValue(pp.ListId, out var info))
            {
                var levels = new ListType[9];
                for (int i = 0; i < 9; i++) levels[i] = pp.List == ListType.Bullet ? ListType.Bullet : (i % 3) switch { 0 => ListType.Decimal, 1 => ListType.LowerAlpha, _ => ListType.LowerRoman };
                info = (levels, num++);
                ctx.Lists[pp.ListId] = info;
            }
            info.Levels[Math.Clamp(pp.ListLevel, 0, 8)] = pp.List;
        }
    }

    private static string Numbering(Ctx ctx)
    {
        var sb = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:numbering xmlns:w=\"{W}\">");
        foreach (var (_, (levels, num)) in ctx.Lists)
        {
            sb.Append($"<w:abstractNum w:abstractNumId=\"{num}\"><w:multiLevelType w:val=\"hybridMultilevel\"/>");
            for (int i = 0; i < 9; i++)
            {
                var t = levels[i];
                string fmt = t switch { ListType.Bullet => "bullet", ListType.LowerAlpha => "lowerLetter", ListType.UpperAlpha => "upperLetter", ListType.LowerRoman => "lowerRoman", ListType.UpperRoman => "upperRoman", _ => "decimal" };
                string text = t == ListType.Bullet ? Layout.DocLayout.Label(ListType.Bullet, i, 1) : t == ListType.LowerAlpha ? $"%{i + 1})" : $"%{i + 1}.";
                sb.Append($"<w:lvl w:ilvl=\"{i}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{fmt}\"/><w:lvlText w:val=\"{X(text)}\"/><w:lvlJc w:val=\"left\"/>" +
                          $"<w:pPr><w:ind w:left=\"{Tw(36 + 18 * i)}\" w:hanging=\"{Tw(18)}\"/></w:pPr>");
                if (t == ListType.Bullet) sb.Append("<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr>");
                sb.Append("</w:lvl>");
            }
            sb.Append("</w:abstractNum>");
        }
        foreach (var (_, (_, num)) in ctx.Lists) sb.Append($"<w:num w:numId=\"{num}\"><w:abstractNumId w:val=\"{num}\"/></w:num>");
        sb.Append("</w:numbering>");
        return sb.ToString();
    }

    private static string Styles(Document d)
    {
        var normal = d.Styles[StyleIds.Normal];
        var sb = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"{W}\">");
        sb.Append("<w:docDefaults><w:rPrDefault>").Append(RPr(normal.Char, null)).Append("</w:rPrDefault><w:pPrDefault><w:pPr>")
          .Append($"<w:spacing w:after=\"{Tw(normal.Para.SpaceAfter)}\" w:line=\"{(int)(normal.Para.LineSpacing * 240)}\" w:lineRule=\"auto\"/>").Append("</w:pPr></w:pPrDefault></w:docDefaults>");
        foreach (var s in d.Styles.Values)
        {
            bool isDefault = s.Id == StyleIds.Normal;
            sb.Append($"<w:style w:type=\"paragraph\" w:styleId=\"{X(s.Id)}\"{(isDefault ? " w:default=\"1\"" : "")}><w:name w:val=\"{X(StyleName(s))}\"/>");
            if (!isDefault) sb.Append("<w:basedOn w:val=\"Normal\"/>");
            if (s.Para.OutlineLevel > 0 || s.Id is StyleIds.Title or StyleIds.Subtitle) sb.Append("<w:next w:val=\"Normal\"/>");
            sb.Append("<w:qFormat/>");
            sb.Append(PPr(s.Para with { Style = "" }, null, includeStyle: false));
            sb.Append(RPr(s.Char, null));
            sb.Append("</w:style>");
        }
        sb.Append("<w:style w:type=\"character\" w:styleId=\"Hyperlink\"><w:name w:val=\"Hyperlink\"/><w:rPr><w:color w:val=\"0563C1\"/><w:u w:val=\"single\"/></w:rPr></w:style>");
        sb.Append("<w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:name w:val=\"Normal Table\"/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><w:left w:w=\"108\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style>");
        sb.Append("</w:styles>");
        return sb.ToString();
    }

    private static string StyleName(StyleDef s) => s.Id switch
    {
        StyleIds.Heading1 => "heading 1", StyleIds.Heading2 => "heading 2", StyleIds.Heading3 => "heading 3", StyleIds.Toc1 => "toc 1", StyleIds.Toc2 => "toc 2",
        StyleIds.Toc3 => "toc 3", StyleIds.TocHeading => "TOC Heading", StyleIds.NoSpacing => "No Spacing", StyleIds.ListParagraph => "List Paragraph", _ => s.NameEn
    };

    private static string Part(string root, List<Paragraph> paras, Ctx ctx, string? watermark)
    {
        var sb = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:{root} xmlns:w=\"{W}\" xmlns:r=\"{R}\" " +
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
            "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">");
        var list = paras.Count > 0 ? paras : [new Paragraph()];
        for (int i = 0; i < list.Count; i++)
        {
            string? extra = i == 0 && !string.IsNullOrWhiteSpace(watermark) ? WatermarkRun(watermark!) : null;
            WriteParagraph(sb, list[i], ctx, extra);
        }
        sb.Append($"</w:{root}>");
        return sb.ToString();
    }

    private static string WatermarkRun(string text) =>
        "<w:r><w:pict><v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\"><v:path textpathok=\"t\" o:connecttype=\"custom\"/><v:textpath on=\"t\" fitshape=\"t\"/></v:shapetype>" +
        "<v:shape id=\"PowerPlusWaterMarkObject\" o:spid=\"_x0000_s2049\" type=\"#_x0000_t136\" style=\"position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin\" o:allowincell=\"f\" fillcolor=\"silver\" stroked=\"f\">" +
        $"<v:fill opacity=\".5\"/><v:textpath style=\"font-family:&quot;Arial&quot;;font-size:1pt\" string=\"{X(text)}\"/></v:shape></w:pict></w:r>";

    private static string SectPr(Document d, string? headerRel, string? footerRel)
    {
        var p = d.Page;
        var sb = new StringBuilder("<w:sectPr>");
        if (headerRel != null) sb.Append($"<w:headerReference w:type=\"default\" r:id=\"{headerRel}\"/>");
        if (footerRel != null) sb.Append($"<w:footerReference w:type=\"default\" r:id=\"{footerRel}\"/>");
        sb.Append($"<w:pgSz w:w=\"{Tw(p.Width)}\" w:h=\"{Tw(p.Height)}\"{(p.Landscape ? " w:orient=\"landscape\"" : "")}/>");
        sb.Append($"<w:pgMar w:top=\"{Tw(p.MarginTop)}\" w:right=\"{Tw(p.MarginRight)}\" w:bottom=\"{Tw(p.MarginBottom)}\" w:left=\"{Tw(p.MarginLeft)}\" w:header=\"{Tw(p.HeaderDistance)}\" w:footer=\"{Tw(p.FooterDistance)}\" w:gutter=\"0\"/>");
        sb.Append("<w:cols w:space=\"720\"/></w:sectPr>");
        return sb.ToString();
    }

    private static void WriteBlock(StringBuilder sb, Block b, Ctx ctx)
    {
        if (b is Paragraph p) WriteParagraph(sb, p, ctx, null);
        else if (b is Table t) WriteTable(sb, t, ctx);
    }

    private static void WriteTable(StringBuilder sb, Table t, Ctx ctx)
    {
        sb.Append("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>");
        if (t.Align != ParaAlign.Left) sb.Append($"<w:jc w:val=\"{(t.Align == ParaAlign.Center ? "center" : "right")}\"/>");
        if (t.Borders)
        {
            var c = Editing.Hex(t.BorderColor);
            int sz = Math.Max(2, (int)Math.Round(t.BorderWidth * 8));
            sb.Append("<w:tblBorders>");
            foreach (var side in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
                sb.Append($"<w:{side} w:val=\"single\" w:sz=\"{sz}\" w:space=\"0\" w:color=\"{c}\"/>");
            sb.Append("</w:tblBorders>");
        }
        sb.Append("<w:tblLayout w:type=\"fixed\"/><w:tblLook w:val=\"04A0\"/></w:tblPr><w:tblGrid>");
        foreach (var w in t.ColumnWidths) sb.Append($"<w:gridCol w:w=\"{Tw(w)}\"/>");
        sb.Append("</w:tblGrid>");
        foreach (var row in t.Rows)
        {
            sb.Append("<w:tr>");
            if (row.MinHeight > 0 || row.HeaderRow)
            {
                sb.Append("<w:trPr>");
                if (row.MinHeight > 0) sb.Append($"<w:trHeight w:val=\"{Tw(row.MinHeight)}\"/>");
                if (row.HeaderRow) sb.Append("<w:tblHeader/>");
                sb.Append("</w:trPr>");
            }
            int g = 0;
            foreach (var cell in row.Cells)
            {
                int span = Math.Max(1, cell.GridSpan);
                float w = t.ColumnWidths.Skip(g).Take(span).Sum();
                g += span;
                sb.Append($"<w:tc><w:tcPr><w:tcW w:w=\"{Tw(w)}\" w:type=\"dxa\"/>");
                if (span > 1) sb.Append($"<w:gridSpan w:val=\"{span}\"/>");
                if (cell.VMerge == VMerge.Restart) sb.Append("<w:vMerge w:val=\"restart\"/>");
                else if (cell.VMerge == VMerge.Continue) sb.Append("<w:vMerge/>");
                if (cell.Shading is Color sh) sb.Append($"<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{Editing.Hex(sh)}\"/>");
                if (cell.VAlign != CellVAlign.Top) sb.Append($"<w:vAlign w:val=\"{(cell.VAlign == CellVAlign.Center ? "center" : "bottom")}\"/>");
                sb.Append("</w:tcPr>");
                var paras = cell.Paragraphs.Count > 0 ? cell.Paragraphs : [new Paragraph()];
                foreach (var p in paras) WriteParagraph(sb, p, ctx, null);
                sb.Append("</w:tc>");
            }
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
    }

    private static string PPr(ParaProps pp, Ctx? ctx, bool includeStyle = true, CharProps? mark = null)
    {
        var sb = new StringBuilder("<w:pPr>");
        if (includeStyle && !string.IsNullOrEmpty(pp.Style) && pp.Style != StyleIds.Normal) sb.Append($"<w:pStyle w:val=\"{X(pp.Style)}\"/>");
        if (pp.KeepWithNext) sb.Append("<w:keepNext/>");
        if (pp.PageBreakBefore) sb.Append("<w:pageBreakBefore/>");
        if (pp.List != ListType.None && ctx != null && ctx.Lists.TryGetValue(pp.ListId, out var li))
            sb.Append($"<w:numPr><w:ilvl w:val=\"{Math.Clamp(pp.ListLevel, 0, 8)}\"/><w:numId w:val=\"{li.NumId}\"/></w:numPr>");
        if (pp.Border)
        {
            sb.Append("<w:pBdr>");
            foreach (var side in new[] { "top", "left", "bottom", "right" }) sb.Append($"<w:{side} w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/>");
            sb.Append("</w:pBdr>");
        }
        if (pp.Shading is Color sh) sb.Append($"<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{Editing.Hex(sh)}\"/>");
        if (pp.RightTabLeader) sb.Append("<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/></w:tabs>");
        string line = pp.LineRule == LineRule.Multiple ? $"w:line=\"{(int)Math.Round(pp.LineSpacing * 240)}\" w:lineRule=\"auto\"" : $"w:line=\"{Tw(pp.LineSpacing)}\" w:lineRule=\"{(pp.LineRule == LineRule.Exact ? "exact" : "atLeast")}\"";
        sb.Append($"<w:spacing w:before=\"{Tw(pp.SpaceBefore)}\" w:after=\"{Tw(pp.SpaceAfter)}\" {line}/>");
        string first = pp.FirstLine >= 0 ? $"w:firstLine=\"{Tw(pp.FirstLine)}\"" : $"w:hanging=\"{Tw(-pp.FirstLine)}\"";
        sb.Append($"<w:ind w:left=\"{Tw(pp.IndentLeft)}\" w:right=\"{Tw(pp.IndentRight)}\" {first}/>");
        sb.Append($"<w:jc w:val=\"{pp.Align switch { ParaAlign.Center => "center", ParaAlign.Right => "right", ParaAlign.Justify => "both", _ => "left" }}\"/>");
        if (pp.OutlineLevel > 0 && !includeStyle) sb.Append($"<w:outlineLvl w:val=\"{pp.OutlineLevel - 1}\"/>");
        if (mark != null) sb.Append(RPr(mark, null));
        sb.Append("</w:pPr>");
        return sb.ToString();
    }

    private static string RPr(CharProps c, string? rStyle)
    {
        var sb = new StringBuilder("<w:rPr>");
        if (rStyle != null) sb.Append($"<w:rStyle w:val=\"{rStyle}\"/>");
        sb.Append($"<w:rFonts w:ascii=\"{X(c.FontName)}\" w:hAnsi=\"{X(c.FontName)}\" w:cs=\"{X(c.FontName)}\" w:eastAsia=\"{X(c.FontName)}\"/>");
        sb.Append(c.Bold ? "<w:b/><w:bCs/>" : "<w:b w:val=\"0\"/>");
        sb.Append(c.Italic ? "<w:i/><w:iCs/>" : "<w:i w:val=\"0\"/>");
        if (c.AllCaps) sb.Append("<w:caps/>");
        if (c.Strike) sb.Append("<w:strike/>");
        if (c.Color is Color col) sb.Append($"<w:color w:val=\"{Editing.Hex(col)}\"/>");
        int hs = (int)Math.Round(c.Size * 2);
        sb.Append($"<w:sz w:val=\"{hs}\"/><w:szCs w:val=\"{hs}\"/>");
        if (c.Underline != UnderlineStyle.None)
            sb.Append($"<w:u w:val=\"{c.Underline switch { UnderlineStyle.Double => "double", UnderlineStyle.Dotted => "dotted", UnderlineStyle.Thick => "thick", UnderlineStyle.Wavy => "wave", _ => "single" }}\"/>");
        if (c.Highlight is Color h) sb.Append($"<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{Editing.Hex(h)}\"/>");
        if (c.VertAlign != VertAlign.Baseline) sb.Append($"<w:vertAlign w:val=\"{(c.VertAlign == VertAlign.Superscript ? "superscript" : "subscript")}\"/>");
        sb.Append("</w:rPr>");
        return sb.ToString();
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, Ctx ctx, string? extraRun)
    {
        sb.Append("<w:p>").Append(PPr(p.Props, ctx, true, p.MarkProps));
        if (extraRun != null) sb.Append(extraRun);
        string? openLink = null;
        foreach (var r in p.Runs)
        {
            if (r.Props.Link != openLink)
            {
                if (openLink != null) sb.Append("</w:hyperlink>");
                openLink = r.Props.Link;
                if (openLink != null)
                {
                    var id = ctx.AddRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", openLink, true);
                    sb.Append($"<w:hyperlink r:id=\"{id}\" w:history=\"1\">");
                }
            }
            if (r.Object is FieldObject f)
            {
                sb.Append($"<w:fldSimple w:instr=\" {f.Instruction} \"><w:r>").Append(RPr(r.Props, null)).Append("<w:t>1</w:t></w:r></w:fldSimple>");
                continue;
            }
            if (r.Object is ImageObject img)
            {
                var name = $"image{ctx.NextImage++}.{(img.Extension == "jpeg" ? "jpg" : img.Extension)}";
                ctx.Media.Add((name, img.Data));
                var rid = ctx.AddRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/" + name);
                int id = ctx.DocPr++;
                long cx = Emu(img.Width), cy = Emu(img.Height);
                sb.Append("<w:r>").Append(RPr(r.Props with { Link = null }, null)).Append("<w:drawing>")
                  .Append($"<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:docPr id=\"{id}\" name=\"Picture {id}\"/>")
                  .Append("<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
                  .Append($"<pic:pic><pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"{name}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>")
                  .Append($"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>")
                  .Append("</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>");
                continue;
            }
            sb.Append("<w:r>").Append(RPr(r.Props, r.Props.Link != null ? "Hyperlink" : null));
            var buf = new StringBuilder();
            void Flush() { if (buf.Length > 0) { sb.Append("<w:t xml:space=\"preserve\">").Append(X(buf.ToString())).Append("</w:t>"); buf.Clear(); } }
            foreach (var ch in r.Text)
            {
                if (ch == '\t') { Flush(); sb.Append("<w:tab/>"); }
                else if (ch == Run.LineBreak) { Flush(); sb.Append("<w:br/>"); }
                else if (ch == Run.PageBreak) { Flush(); sb.Append("<w:br w:type=\"page\"/>"); }
                else if (ch < 32) { }
                else buf.Append(ch);
            }
            Flush();
            sb.Append("</w:r>");
        }
        if (openLink != null) sb.Append("</w:hyperlink>");
        sb.Append("</w:p>");
    }
}
