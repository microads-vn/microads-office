using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ValidationType = MicroadsOffice.Excel.Model.ValidationType;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.IO;

/// <summary>Writes .xlsx (SpreadsheetML) without third-party libraries.</summary>
public static class XlsxWriter
{
    internal const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    internal const string NsDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    internal const string NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
    internal const string NsC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const int EmuPerPx = 9525;
    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public static void Write(Workbook wb, string path)
    {
        var tmp = path + ".tmp";
        using (var fs = File.Create(tmp))
            Write(wb, fs);
        File.Move(tmp, path, true);
    }

    public static void Write(Workbook wb, Stream output)
    {
        using var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        var sst = new List<string>();
        var sstIndex = new Dictionary<string, int>();
        int Sst(string s) { if (!sstIndex.TryGetValue(s, out var i)) { i = sst.Count; sst.Add(s); sstIndex[s] = i; } return i; }
        var dxfs = new List<ConditionalFormat>();
        var contentOverrides = new List<(string Part, string Type)>();
        int chartCounter = 0, drawingCounter = 0, commentCounter = 0, imageCounter = 0;
        var imageExts = new HashSet<string>();

        for (int si = 0; si < wb.Sheets.Count; si++)
        {
            var s = wb.Sheets[si];
            var rels = new List<(string Id, string Type, string Target, bool External)>();
            string? drawingRel = null, vmlRel = null;
            // drawing (charts + pictures)
            if (s.Charts.Count > 0 || s.Pictures.Count > 0)
            {
                drawingCounter++;
                var drels = new List<(string Id, string Type, string Target)>();
                var dsb = new StringBuilder();
                dsb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"").Append(NsDrawing).Append("\" xmlns:a=\"").Append(NsA).Append("\" xmlns:r=\"").Append(NsRel).Append("\" xmlns:c=\"").Append(NsC).Append("\">");
                int objId = 2;
                foreach (var ch in s.Charts)
                {
                    chartCounter++;
                    var rid = "rId" + (drels.Count + 1);
                    drels.Add((rid, NsRel + "/chart", "../charts/chart" + chartCounter + ".xml"));
                    dsb.Append(Anchor(s, ch.Bounds));
                    dsb.Append("<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"").Append(objId++).Append("\" name=\"").Append(Esc(ch.Name)).Append("\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>")
                       .Append("<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"").Append(NsC).Append("\"><c:chart r:id=\"").Append(rid).Append("\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>");
                    Put(zip, "xl/charts/chart" + chartCounter + ".xml", ChartXml(s, ch));
                    contentOverrides.Add(("/xl/charts/chart" + chartCounter + ".xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"));
                }
                foreach (var pic in s.Pictures)
                {
                    imageCounter++;
                    var ext = pic.Ext.ToLowerInvariant().TrimStart('.');
                    if (ext == "jpg") ext = "jpeg";
                    imageExts.Add(ext);
                    var rid = "rId" + (drels.Count + 1);
                    drels.Add((rid, NsRel + "/image", "../media/image" + imageCounter + "." + ext));
                    PutBytes(zip, "xl/media/image" + imageCounter + "." + ext, pic.Data);
                    dsb.Append(Anchor(s, pic.Bounds, oneCell: true));
                    dsb.Append("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"").Append(objId++).Append("\" name=\"").Append(Esc(pic.Name)).Append("\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>")
                       .Append("<xdr:blipFill><a:blip r:embed=\"").Append(rid).Append("\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>")
                       .Append("<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"").Append((long)(pic.Bounds.Width * EmuPerPx)).Append("\" cy=\"").Append((long)(pic.Bounds.Height * EmuPerPx)).Append("\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
                }
                dsb.Append("</xdr:wsDr>");
                Put(zip, "xl/drawings/drawing" + drawingCounter + ".xml", dsb.ToString());
                Put(zip, "xl/drawings/_rels/drawing" + drawingCounter + ".xml.rels", RelsXml(drels.Select(r => (r.Id, r.Type, r.Target, false))));
                contentOverrides.Add(("/xl/drawings/drawing" + drawingCounter + ".xml", "application/vnd.openxmlformats-officedocument.drawing+xml"));
                drawingRel = "rId" + (rels.Count + 1);
                rels.Add((drawingRel, NsRel + "/drawing", "../drawings/drawing" + drawingCounter + ".xml", false));
            }
            if (s.Notes.Count > 0)
            {
                commentCounter++;
                Put(zip, "xl/comments" + commentCounter + ".xml", CommentsXml(s));
                Put(zip, "xl/drawings/vmlDrawing" + commentCounter + ".vml", VmlXml(s));
                contentOverrides.Add(("/xl/comments" + commentCounter + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"));
                rels.Add(("rId" + (rels.Count + 1), NsRel + "/comments", "../comments" + commentCounter + ".xml", false));
                vmlRel = "rId" + (rels.Count + 1);
                rels.Add((vmlRel, NsRel + "/vmlDrawing", "../drawings/vmlDrawing" + commentCounter + ".vml", false));
            }
            var linkRels = new Dictionary<long, string>();
            foreach (var (k, url) in s.Links)
            {
                if (url.StartsWith('#')) continue;
                var id = "rId" + (rels.Count + 1);
                rels.Add((id, NsRel + "/hyperlink", url, true));
                linkRels[k] = id;
            }
            Put(zip, $"xl/worksheets/sheet{si + 1}.xml", SheetXml(wb, s, Sst, dxfs, linkRels, drawingRel, vmlRel));
            if (rels.Count > 0) Put(zip, $"xl/worksheets/_rels/sheet{si + 1}.xml.rels", RelsXml(rels));
        }

        Put(zip, "xl/workbook.xml", WorkbookXml(wb));
        var wbRels = new List<(string, string, string, bool)>();
        for (int i = 0; i < wb.Sheets.Count; i++) wbRels.Add(("rId" + (i + 1), NsRel + "/worksheet", "worksheets/sheet" + (i + 1) + ".xml", false));
        int n = wb.Sheets.Count;
        wbRels.Add(("rId" + (n + 1), NsRel + "/styles", "styles.xml", false));
        wbRels.Add(("rId" + (n + 2), NsRel + "/sharedStrings", "sharedStrings.xml", false));
        wbRels.Add(("rId" + (n + 3), NsRel + "/theme", "theme/theme1.xml", false));
        Put(zip, "xl/_rels/workbook.xml.rels", RelsXml(wbRels));
        Put(zip, "xl/styles.xml", StylesXml(wb, dxfs));
        Put(zip, "xl/theme/theme1.xml", ThemeXml);
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"" + NsMain + "\" count=\"" + sst.Count + "\" uniqueCount=\"" + sst.Count + "\">");
        foreach (var s in sst) sb.Append("<si><t xml:space=\"preserve\">").Append(Esc(s)).Append("</t></si>");
        sb.Append("</sst>");
        Put(zip, "xl/sharedStrings.xml", sb.ToString());
        Put(zip, "docProps/core.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
            + "<dc:title>" + Esc(wb.Title) + "</dc:title><dc:creator>" + Esc(wb.Author) + "</dc:creator><dcterms:created xsi:type=\"dcterms:W3CDTF\">" + DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", Inv)
            + "</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", Inv) + "</dcterms:modified></cp:coreProperties>");
        Put(zip, "docProps/app.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Microads Office Excel</Application><Company>microads.vn</Company></Properties>");
        Put(zip, "_rels/.rels", RelsXml([("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml", false),
            ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml", false),
            ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml", false)]));
        var ct = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");
        foreach (var e in imageExts) ct.Append("<Default Extension=\"").Append(e).Append("\" ContentType=\"image/").Append(e).Append("\"/>");
        ct.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        for (int i = 0; i < wb.Sheets.Count; i++) ct.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i + 1).Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        ct.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>")
          .Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>")
          .Append("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>")
          .Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>")
          .Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        foreach (var (p, t) in contentOverrides) ct.Append("<Override PartName=\"").Append(p).Append("\" ContentType=\"").Append(t).Append("\"/>");
        ct.Append("</Types>");
        Put(zip, "[Content_Types].xml", ct.ToString());
    }

    // ------------------------------------------------------------------ helpers

    internal static string Esc(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            switch (ch)
            {
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '&': sb.Append("&amp;"); break;
                case '"': sb.Append("&quot;"); break;
                default:
                    if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') sb.Append("_x").Append(((int)ch).ToString("X4")).Append('_');
                    else sb.Append(ch);
                    break;
            }
        }
        return sb.ToString();
    }

    private static void Put(ZipArchive zip, string name, string content)
    {
        var e = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var w = new StreamWriter(e.Open(), new UTF8Encoding(false));
        w.Write(content);
    }

    private static void PutBytes(ZipArchive zip, string name, byte[] data)
    {
        var e = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var w = e.Open();
        w.Write(data);
    }

    private static string RelsXml(IEnumerable<(string Id, string Type, string Target, bool External)> rels)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"" + NsPkgRel + "\">");
        foreach (var (id, type, target, ext) in rels)
            sb.Append("<Relationship Id=\"").Append(id).Append("\" Type=\"").Append(type).Append("\" Target=\"").Append(Esc(target)).Append('"').Append(ext ? " TargetMode=\"External\"" : "").Append("/>");
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    internal static string Hex(int argb) => ((uint)argb).ToString("X8");

    private static string F(double d) => d.ToString("0.########", Inv);

    // ------------------------------------------------------------------ workbook

    private static string WorkbookXml(Workbook wb)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"" + NsMain + "\" xmlns:r=\"" + NsRel + "\">");
        sb.Append("<workbookPr/><bookViews><workbookView activeTab=\"").Append(Math.Clamp(wb.ActiveSheet, 0, wb.Sheets.Count - 1)).Append("\"/></bookViews><sheets>");
        for (int i = 0; i < wb.Sheets.Count; i++)
            sb.Append("<sheet name=\"").Append(Esc(wb.Sheets[i].Name)).Append("\" sheetId=\"").Append(i + 1).Append('"').Append(wb.Sheets[i].Hidden ? " state=\"hidden\"" : "").Append(" r:id=\"rId").Append(i + 1).Append("\"/>");
        sb.Append("</sheets>");
        var names = new StringBuilder();
        foreach (var (k, v) in wb.Names) names.Append("<definedName name=\"").Append(Esc(k)).Append("\">").Append(Esc(v)).Append("</definedName>");
        for (int i = 0; i < wb.Sheets.Count; i++)
        {
            var s = wb.Sheets[i];
            var q = Addr.QuoteSheet(s.Name);
            if (s.Page.PrintArea is CellRange pa)
                names.Append("<definedName name=\"_xlnm.Print_Area\" localSheetId=\"").Append(i).Append("\">").Append(Esc(q + "!" + Abs(pa))).Append("</definedName>");
            if (s.Page.PrintTitleRows is (int r1, int r2))
                names.Append("<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"").Append(i).Append("\">").Append(Esc(q + "!$" + (r1 + 1) + ":$" + (r2 + 1))).Append("</definedName>");
            if (s.Filter != null)
                names.Append("<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"").Append(i).Append("\" hidden=\"1\">").Append(Esc(q + "!" + Abs(s.Filter.Range))).Append("</definedName>");
        }
        if (names.Length > 0) sb.Append("<definedNames>").Append(names).Append("</definedNames>");
        sb.Append("<calcPr calcId=\"191029\"").Append(wb.ManualCalc ? " calcMode=\"manual\"" : "").Append(" fullCalcOnLoad=\"1\"/></workbook>");
        return sb.ToString();
    }

    private static string Abs(CellRange r) => Addr.ToA1(r.R1, r.C1, true, true) + (r.IsSingle ? "" : ":" + Addr.ToA1(r.R2, r.C2, true, true));

    // ------------------------------------------------------------------ sheet

    private static string SheetXml(Workbook wb, Sheet s, Func<string, int> sst, List<ConditionalFormat> dxfs, Dictionary<long, string> linkRels, string? drawingRel, string? vmlRel)
    {
        var sb = new StringBuilder(4096 + s.CellCount * 40);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"").Append(NsMain).Append("\" xmlns:r=\"").Append(NsRel).Append("\">");
        bool fit = s.Page.FitToWidth;
        if (s.TabColor != 0 || fit)
        {
            sb.Append("<sheetPr>");
            if (s.TabColor != 0) sb.Append("<tabColor rgb=\"").Append(Hex(s.TabColor)).Append("\"/>");
            if (fit) sb.Append("<pageSetUpPr fitToPage=\"1\"/>");
            sb.Append("</sheetPr>");
        }
        var used = s.UsedRange();
        sb.Append("<dimension ref=\"").Append(used.ToString()).Append("\"/>");
        sb.Append("<sheetViews><sheetView workbookViewId=\"0\"");
        if (!s.ShowGridlines) sb.Append(" showGridLines=\"0\"");
        if (!s.ShowHeadings) sb.Append(" showRowColHeaders=\"0\"");
        if (Math.Abs(s.Zoom - 1) > 0.001) sb.Append(" zoomScale=\"").Append((int)Math.Round(s.Zoom * 100)).Append('"');
        if (s.Book.Sheets.IndexOf(s) == s.Book.ActiveSheet) sb.Append(" tabSelected=\"1\"");
        sb.Append('>');
        if (s.FreezeRows > 0 || s.FreezeCols > 0)
        {
            var pane = s.FreezeRows > 0 && s.FreezeCols > 0 ? "bottomRight" : s.FreezeRows > 0 ? "bottomLeft" : "topRight";
            sb.Append("<pane");
            if (s.FreezeCols > 0) sb.Append(" xSplit=\"").Append(s.FreezeCols).Append('"');
            if (s.FreezeRows > 0) sb.Append(" ySplit=\"").Append(s.FreezeRows).Append('"');
            sb.Append(" topLeftCell=\"").Append(Addr.ToA1(s.FreezeRows, s.FreezeCols)).Append("\" activePane=\"").Append(pane).Append("\" state=\"frozen\"/>");
        }
        sb.Append("<selection activeCell=\"").Append(s.ActiveCell.ToString()).Append("\" sqref=\"").Append(s.Selection.Normalized.ToString()).Append("\"/>");
        sb.Append("</sheetView></sheetViews>");
        sb.Append("<sheetFormatPr defaultRowHeight=\"15\"/>");
        // columns
        var colKeys = s.ColWidths.Keys.Concat(s.HiddenCols).Concat(s.ColStyles.Keys).Distinct().OrderBy(x => x).ToList();
        if (colKeys.Count > 0)
        {
            sb.Append("<cols>");
            foreach (var c in colKeys)
            {
                double w = Math.Max(0, s.RawColWidth(c) / 7.0);
                sb.Append("<col min=\"").Append(c + 1).Append("\" max=\"").Append(c + 1).Append("\" width=\"").Append(F(Math.Round(w, 4))).Append("\" customWidth=\"1\"");
                if (s.HiddenCols.Contains(c)) sb.Append(" hidden=\"1\"");
                if (s.ColStyles.TryGetValue(c, out var cs)) sb.Append(" style=\"").Append(cs).Append('"');
                sb.Append("/>");
            }
            sb.Append("</cols>");
        }
        // data
        sb.Append("<sheetData>");
        var byRow = new SortedDictionary<int, List<(int Col, Cell Cell)>>();
        foreach (var (k, cell) in s.Cells)
        {
            var (r, c) = Sheet.FromKey(k);
            if (!byRow.TryGetValue(r, out var l)) byRow[r] = l = new();
            l.Add((c, cell));
        }
        foreach (var r in s.RowHeights.Keys.Concat(s.AutoHeights.Keys).Concat(s.HiddenRows).Concat(s.FilteredRows).Concat(s.RowStyles.Keys)) if (!byRow.ContainsKey(r)) byRow[r] = new();
        foreach (var (r, list) in byRow)
        {
            sb.Append("<row r=\"").Append(r + 1).Append('"');
            if (s.RowHeights.TryGetValue(r, out var h)) sb.Append(" ht=\"").Append(F(Math.Round(h * 0.75, 2))).Append("\" customHeight=\"1\"");
            else if (s.AutoHeights.TryGetValue(r, out var ah)) sb.Append(" ht=\"").Append(F(Math.Round(ah * 0.75, 2))).Append('"');
            if (s.HiddenRows.Contains(r) || s.FilteredRows.Contains(r)) sb.Append(" hidden=\"1\"");
            if (s.RowStyles.TryGetValue(r, out var rs)) sb.Append(" s=\"").Append(rs).Append("\" customFormat=\"1\"");
            sb.Append('>');
            list.Sort((a, b) => a.Col.CompareTo(b.Col));
            foreach (var (c, cell) in list)
            {
                sb.Append("<c r=\"").Append(Addr.ColName(c)).Append(r + 1).Append('"');
                if (cell.StyleId != 0) sb.Append(" s=\"").Append(cell.StyleId).Append('"');
                var v = cell.Value;
                if (cell.Formula != null)
                {
                    if (v.IsText) sb.Append(" t=\"str\"");
                    else if (v.IsBool) sb.Append(" t=\"b\"");
                    else if (v.IsError) sb.Append(" t=\"e\"");
                    sb.Append("><f>").Append(Esc(cell.Formula)).Append("</f>");
                    if (!v.IsBlank) sb.Append("<v>").Append(Esc(v.IsNumber ? F(v.Num) : v.IsBool ? (v.Num != 0 ? "1" : "0") : v.Str!)).Append("</v>");
                    sb.Append("</c>");
                    continue;
                }
                switch (v.Kind)
                {
                    case ValueKind.Blank: sb.Append("/>"); break;
                    case ValueKind.Number: sb.Append("><v>").Append(v.Num.ToString("R", Inv)).Append("</v></c>"); break;
                    case ValueKind.Text: sb.Append(" t=\"s\"><v>").Append(sst(v.Str!)).Append("</v></c>"); break;
                    case ValueKind.Bool: sb.Append(" t=\"b\"><v>").Append(v.Num != 0 ? 1 : 0).Append("</v></c>"); break;
                    case ValueKind.Error: sb.Append(" t=\"e\"><v>").Append(Esc(v.Str!)).Append("</v></c>"); break;
                }
            }
            sb.Append("</row>");
        }
        sb.Append("</sheetData>");
        if (s.Protected)
        {
            sb.Append("<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"");
            if (!string.IsNullOrEmpty(s.ProtectionPasswordHash)) sb.Append(" password=\"").Append(s.ProtectionPasswordHash).Append('"');
            sb.Append("/>");
        }
        if (s.Filter != null)
        {
            sb.Append("<autoFilter ref=\"").Append(s.Filter.Range.ToString()).Append('"');
            var cols = s.Filter.Allowed.Keys.Concat(s.Filter.Conditions.Keys).Distinct().OrderBy(x => x).ToList();
            if (cols.Count == 0) sb.Append("/>");
            else
            {
                sb.Append('>');
                foreach (var col in cols)
                {
                    sb.Append("<filterColumn colId=\"").Append(col - s.Filter.Range.C1).Append("\">");
                    if (s.Filter.Allowed.TryGetValue(col, out var allowed))
                    {
                        sb.Append("<filters").Append(allowed.Contains("") ? " blank=\"1\"" : "").Append('>');
                        foreach (var a in allowed.Where(x => x.Length > 0)) sb.Append("<filter val=\"").Append(Esc(a)).Append("\"/>");
                        sb.Append("</filters>");
                    }
                    else if (s.Filter.Conditions.TryGetValue(col, out var cond))
                    {
                        var (op, val) = CondToCustom(cond);
                        sb.Append("<customFilters><customFilter operator=\"").Append(op).Append("\" val=\"").Append(Esc(val)).Append("\"/></customFilters>");
                    }
                    sb.Append("</filterColumn>");
                }
                sb.Append("</autoFilter>");
            }
        }
        if (s.Merges.Count > 0)
        {
            sb.Append("<mergeCells count=\"").Append(s.Merges.Count).Append("\">");
            foreach (var m in s.Merges) sb.Append("<mergeCell ref=\"").Append(m.ToString()).Append("\"/>");
            sb.Append("</mergeCells>");
        }
        int prio = 1;
        foreach (var cf in s.CondFormats)
        {
            sb.Append("<conditionalFormatting sqref=\"").Append(cf.Range.ToString()).Append("\">");
            int dxfId = -1;
            if (cf.Type is not (CfType.DataBar or CfType.ColorScale2 or CfType.ColorScale3)) { dxfs.Add(cf); dxfId = dxfs.Count - 1; }
            string dxfAttr = dxfId >= 0 ? " dxfId=\"" + dxfId + "\"" : "";
            string first = Addr.ToA1(cf.Range.R1, cf.Range.C1);
            switch (cf.Type)
            {
                case CfType.GreaterThan: case CfType.LessThan: case CfType.Between: case CfType.EqualTo:
                {
                    var op = cf.Type switch { CfType.GreaterThan => "greaterThan", CfType.LessThan => "lessThan", CfType.Between => "between", _ => "equal" };
                    sb.Append("<cfRule type=\"cellIs\"").Append(dxfAttr).Append(" priority=\"").Append(prio++).Append("\" operator=\"").Append(op).Append("\"><formula>").Append(Esc(CfFormula(cf.Value1))).Append("</formula>");
                    if (cf.Type == CfType.Between) sb.Append("<formula>").Append(Esc(CfFormula(cf.Value2))).Append("</formula>");
                    sb.Append("</cfRule>");
                    break;
                }
                case CfType.TextContains:
                    sb.Append("<cfRule type=\"containsText\"").Append(dxfAttr).Append(" priority=\"").Append(prio++).Append("\" operator=\"containsText\" text=\"").Append(Esc(cf.Value1)).Append("\"><formula>")
                      .Append(Esc($"NOT(ISERROR(SEARCH(\"{cf.Value1.Replace("\"", "\"\"")}\",{first})))")).Append("</formula></cfRule>");
                    break;
                case CfType.Duplicate:
                    sb.Append("<cfRule type=\"duplicateValues\"").Append(dxfAttr).Append(" priority=\"").Append(prio++).Append("\"/>");
                    break;
                case CfType.Top10: case CfType.Bottom10:
                    sb.Append("<cfRule type=\"top10\"").Append(dxfAttr).Append(" priority=\"").Append(prio++).Append("\" rank=\"").Append(cf.Rank).Append('"').Append(cf.Type == CfType.Bottom10 ? " bottom=\"1\"" : "").Append("/>");
                    break;
                case CfType.AboveAverage: case CfType.BelowAverage:
                    sb.Append("<cfRule type=\"aboveAverage\"").Append(dxfAttr).Append(" priority=\"").Append(prio++).Append('"').Append(cf.Type == CfType.BelowAverage ? " aboveAverage=\"0\"" : "").Append("/>");
                    break;
                case CfType.Expression:
                    sb.Append("<cfRule type=\"expression\"").Append(dxfAttr).Append(" priority=\"").Append(prio++).Append("\"><formula>").Append(Esc(cf.Value1.TrimStart('='))).Append("</formula></cfRule>");
                    break;
                case CfType.DataBar:
                    sb.Append("<cfRule type=\"dataBar\" priority=\"").Append(prio++).Append("\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"").Append(Hex(cf.Color1)).Append("\"/></dataBar></cfRule>");
                    break;
                case CfType.ColorScale2:
                    sb.Append("<cfRule type=\"colorScale\" priority=\"").Append(prio++).Append("\"><colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"").Append(Hex(cf.Color1)).Append("\"/><color rgb=\"").Append(Hex(cf.Color3)).Append("\"/></colorScale></cfRule>");
                    break;
                case CfType.ColorScale3:
                    sb.Append("<cfRule type=\"colorScale\" priority=\"").Append(prio++).Append("\"><colorScale><cfvo type=\"min\"/><cfvo type=\"percentile\" val=\"50\"/><cfvo type=\"max\"/><color rgb=\"").Append(Hex(cf.Color1))
                      .Append("\"/><color rgb=\"").Append(Hex(cf.Color2)).Append("\"/><color rgb=\"").Append(Hex(cf.Color3)).Append("\"/></colorScale></cfRule>");
                    break;
            }
            sb.Append("</conditionalFormatting>");
        }
        if (s.Validations.Count > 0)
        {
            sb.Append("<dataValidations count=\"").Append(s.Validations.Count).Append("\">");
            foreach (var dv in s.Validations)
            {
                var type = dv.Type switch { ValidationType.List => "list", ValidationType.WholeNumber => "whole", ValidationType.Decimal => "decimal", ValidationType.Date => "date", ValidationType.TextLength => "textLength", _ => "none" };
                var op = dv.Op switch
                {
                    ValidationOp.NotBetween => "notBetween", ValidationOp.Equal => "equal", ValidationOp.NotEqual => "notEqual", ValidationOp.Greater => "greaterThan",
                    ValidationOp.Less => "lessThan", ValidationOp.GreaterOrEqual => "greaterThanOrEqual", ValidationOp.LessOrEqual => "lessThanOrEqual", _ => "between"
                };
                sb.Append("<dataValidation type=\"").Append(type).Append('"');
                if (dv.Type != ValidationType.List && dv.Type != ValidationType.Any && dv.Op != ValidationOp.Between) sb.Append(" operator=\"").Append(op).Append('"');
                if (dv.AllowBlank) sb.Append(" allowBlank=\"1\"");
                sb.Append(" showErrorMessage=\"1\"");
                if (dv.ErrorMessage.Length > 0) sb.Append(" error=\"").Append(Esc(dv.ErrorMessage)).Append('"');
                sb.Append(" sqref=\"").Append(dv.Range.ToString()).Append("\">");
                if (dv.Formula1.Length > 0)
                {
                    var f1 = dv.Type == ValidationType.List && !dv.Formula1.StartsWith('=') ? "\"" + dv.Formula1.Replace("\"", "\"\"") + "\"" : dv.Formula1.TrimStart('=');
                    sb.Append("<formula1>").Append(Esc(f1)).Append("</formula1>");
                }
                if (dv.Formula2.Length > 0) sb.Append("<formula2>").Append(Esc(dv.Formula2.TrimStart('='))).Append("</formula2>");
                sb.Append("</dataValidation>");
            }
            sb.Append("</dataValidations>");
        }
        if (s.Links.Count > 0)
        {
            sb.Append("<hyperlinks>");
            foreach (var (k, url) in s.Links)
            {
                var (r, c) = Sheet.FromKey(k);
                sb.Append("<hyperlink ref=\"").Append(Addr.ToA1(r, c)).Append('"');
                if (linkRels.TryGetValue(k, out var rid)) sb.Append(" r:id=\"").Append(rid).Append('"');
                else sb.Append(" location=\"").Append(Esc(url.TrimStart('#'))).Append('"');
                sb.Append("/>");
            }
            sb.Append("</hyperlinks>");
        }
        var p = s.Page;
        if (p.PrintGridlines || p.PrintHeadings || p.CenterHorizontally)
        {
            sb.Append("<printOptions");
            if (p.CenterHorizontally) sb.Append(" horizontalCentered=\"1\"");
            if (p.PrintHeadings) sb.Append(" headings=\"1\"");
            if (p.PrintGridlines) sb.Append(" gridLines=\"1\"");
            sb.Append("/>");
        }
        sb.Append("<pageMargins left=\"").Append(F(p.MarginLeft)).Append("\" right=\"").Append(F(p.MarginRight)).Append("\" top=\"").Append(F(p.MarginTop)).Append("\" bottom=\"").Append(F(p.MarginBottom)).Append("\" header=\"0.3\" footer=\"0.3\"/>");
        sb.Append("<pageSetup paperSize=\"").Append(p.PaperSize).Append("\" orientation=\"").Append(p.Landscape ? "landscape" : "portrait").Append('"');
        if (p.Scale != 100) sb.Append(" scale=\"").Append(p.Scale).Append('"');
        if (p.FitToWidth) sb.Append(" fitToWidth=\"1\" fitToHeight=\"0\"");
        sb.Append("/>");
        if (drawingRel != null) sb.Append("<drawing r:id=\"").Append(drawingRel).Append("\"/>");
        if (vmlRel != null) sb.Append("<legacyDrawing r:id=\"").Append(vmlRel).Append("\"/>");
        sb.Append("</worksheet>");
        return sb.ToString();
    }

    private static string CfFormula(string v)
    {
        v = v.Trim();
        if (v.StartsWith('=')) return v[1..];
        if (double.TryParse(v, NumberStyles.Float, Inv, out _)) return v;
        if (ValueParser.TryCoerceNumber(v, out var d)) return F(d);
        return "\"" + v.Replace("\"", "\"\"") + "\"";
    }

    private static (string Op, string Val) CondToCustom(string cond)
    {
        if (cond.StartsWith("contains:")) return ("equal", "*" + cond[9..] + "*");
        if (cond.StartsWith("begins:")) return ("equal", cond[7..] + "*");
        if (cond.StartsWith("ends:")) return ("equal", "*" + cond[5..]);
        foreach (var (p, op) in new[] { (">=", "greaterThanOrEqual"), ("<=", "lessThanOrEqual"), ("<>", "notEqual"), (">", "greaterThan"), ("<", "lessThan"), ("=", "equal") })
            if (cond.StartsWith(p)) return (op, cond[p.Length..]);
        return ("equal", cond);
    }

    private static string Anchor(Sheet s, RectangleF b, bool oneCell = false)
    {
        var (c1, ox1) = PxToCol(s, b.X);
        var (r1, oy1) = PxToRow(s, b.Y);
        var (c2, ox2) = PxToCol(s, b.Right);
        var (r2, oy2) = PxToRow(s, b.Bottom);
        return "<xdr:twoCellAnchor" + (oneCell ? " editAs=\"oneCell\"" : "") + "><xdr:from><xdr:col>" + c1 + "</xdr:col><xdr:colOff>" + (long)(ox1 * EmuPerPx) + "</xdr:colOff><xdr:row>" + r1 + "</xdr:row><xdr:rowOff>" + (long)(oy1 * EmuPerPx)
            + "</xdr:rowOff></xdr:from><xdr:to><xdr:col>" + c2 + "</xdr:col><xdr:colOff>" + (long)(ox2 * EmuPerPx) + "</xdr:colOff><xdr:row>" + r2 + "</xdr:row><xdr:rowOff>" + (long)(oy2 * EmuPerPx) + "</xdr:rowOff></xdr:to>";
    }

    internal static (int, double) PxToCol(Sheet s, double x)
    {
        int c = 0;
        while (c < Addr.MaxCols - 1) { var w = s.ColWidth(c); if (x < w) break; x -= w; c++; }
        return (c, Math.Max(0, x));
    }

    internal static (int, double) PxToRow(Sheet s, double y)
    {
        int r = 0;
        while (r < Addr.MaxRows - 1) { var h = s.RowHeight(r); if (y < h) break; y -= h; r++; }
        return (r, Math.Max(0, y));
    }

    // ------------------------------------------------------------------ comments

    private static string CommentsXml(Sheet s)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<comments xmlns=\"" + NsMain + "\"><authors><author>Microads Office</author></authors><commentList>");
        foreach (var (k, text) in s.Notes)
        {
            var (r, c) = Sheet.FromKey(k);
            sb.Append("<comment ref=\"").Append(Addr.ToA1(r, c)).Append("\" authorId=\"0\"><text><t xml:space=\"preserve\">").Append(Esc(text)).Append("</t></text></comment>");
        }
        sb.Append("</commentList></comments>");
        return sb.ToString();
    }

    private static string VmlXml(Sheet s)
    {
        var sb = new StringBuilder("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">"
            + "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>");
        int id = 1025;
        foreach (var k in s.Notes.Keys)
        {
            var (r, c) = Sheet.FromKey(k);
            sb.Append("<v:shape id=\"_x0000_s").Append(id++).Append("\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:2pt;width:120pt;height:60pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">")
              .Append("<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>")
              .Append("<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>").Append(c + 1).Append(", 15, ").Append(r).Append(", 2, ").Append(c + 3).Append(", 15, ").Append(r + 4).Append(", 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>")
              .Append(r).Append("</x:Row><x:Column>").Append(c).Append("</x:Column></x:ClientData></v:shape>");
        }
        sb.Append("</xml>");
        return sb.ToString();
    }

    // ------------------------------------------------------------------ chart

    private static string ChartXml(Sheet s, Chart ch)
    {
        var (cats, series) = ChartData.Layout(ch);
        string q = Addr.QuoteSheet(s.Name) + "!";
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<c:chartSpace xmlns:c=\"" + NsC + "\" xmlns:a=\"" + NsA + "\" xmlns:r=\"" + NsRel + "\"><c:roundedCorners val=\"0\"/><c:chart>");
        if (ch.Title.Length > 0)
            sb.Append("<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>").Append(Esc(ch.Title)).Append("</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"0\"/>");
        else sb.Append("<c:autoTitleDeleted val=\"1\"/>");
        sb.Append("<c:plotArea><c:layout/>");
        string tag = ch.Type switch
        {
            ChartType.Column or ChartType.StackedColumn or ChartType.Bar => "c:barChart",
            ChartType.Line => "c:lineChart",
            ChartType.Pie => "c:pieChart",
            ChartType.Doughnut => "c:doughnutChart",
            ChartType.Area => "c:areaChart",
            _ => "c:scatterChart"
        };
        sb.Append('<').Append(tag).Append('>');
        if (ch.Type is ChartType.Column or ChartType.StackedColumn or ChartType.Bar)
            sb.Append("<c:barDir val=\"").Append(ch.Type == ChartType.Bar ? "bar" : "col").Append("\"/><c:grouping val=\"").Append(ch.Type == ChartType.StackedColumn ? "stacked" : "clustered").Append("\"/>");
        else if (ch.Type is ChartType.Line or ChartType.Area) sb.Append("<c:grouping val=\"standard\"/>");
        else if (ch.Type == ChartType.Scatter) sb.Append("<c:scatterStyle val=\"lineMarker\"/>");
        sb.Append("<c:varyColors val=\"").Append(ch.Type is ChartType.Pie or ChartType.Doughnut ? 1 : 0).Append("\"/>");
        int idx = 0;
        foreach (var ser in series)
        {
            sb.Append("<c:ser><c:idx val=\"").Append(idx).Append("\"/><c:order val=\"").Append(idx).Append("\"/>");
            if (ser.NameCell is CellPos np) sb.Append("<c:tx><c:strRef><c:f>").Append(Esc(q + Addr.ToA1(np.Row, np.Col, true, true))).Append("</c:f></c:strRef></c:tx>");
            var color = ChartData.Palette[idx % ChartData.Palette.Length];
            if (ch.Type is not (ChartType.Pie or ChartType.Doughnut))
            {
                if (ch.Type == ChartType.Line || ch.Type == ChartType.Scatter)
                    sb.Append("<c:spPr><a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"").Append(Hex(color.ToArgb())[2..]).Append("\"/></a:solidFill></a:ln></c:spPr>");
                else sb.Append("<c:spPr><a:solidFill><a:srgbClr val=\"").Append(Hex(color.ToArgb())[2..]).Append("\"/></a:solidFill></c:spPr>");
            }
            if (ch.Type == ChartType.Bar || ch.Type is ChartType.Column or ChartType.StackedColumn) sb.Append("<c:invertIfNegative val=\"0\"/>");
            if (ch.ShowDataLabels) sb.Append("<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
            if (cats is CellRange cr)
            {
                var catTag = ch.Type == ChartType.Scatter ? "c:xVal" : "c:cat";
                sb.Append('<').Append(catTag).Append("><c:").Append(ch.Type == ChartType.Scatter ? "numRef" : "strRef").Append("><c:f>").Append(Esc(q + Abs(cr))).Append("</c:f></c:").Append(ch.Type == ChartType.Scatter ? "numRef" : "strRef").Append("></").Append(catTag).Append('>');
            }
            var valTag = ch.Type == ChartType.Scatter ? "c:yVal" : "c:val";
            sb.Append('<').Append(valTag).Append("><c:numRef><c:f>").Append(Esc(q + Abs(ser.Values))).Append("</c:f></c:numRef></").Append(valTag).Append('>');
            if (ch.Type == ChartType.Line) sb.Append("<c:smooth val=\"0\"/>");
            if (ch.Type == ChartType.Scatter) sb.Append("<c:smooth val=\"0\"/>");
            sb.Append("</c:ser>");
            idx++;
        }
        if (ch.Type is ChartType.Column or ChartType.StackedColumn or ChartType.Bar) sb.Append("<c:gapWidth val=\"150\"/>").Append(ch.Type == ChartType.StackedColumn ? "<c:overlap val=\"100\"/>" : "");
        if (ch.Type == ChartType.Line) sb.Append("<c:marker val=\"1\"/>");
        if (ch.Type == ChartType.Doughnut) sb.Append("<c:firstSliceAng val=\"0\"/><c:holeSize val=\"55\"/>");
        if (ch.Type == ChartType.Pie) sb.Append("<c:firstSliceAng val=\"0\"/>");
        bool axes = ch.Type is not (ChartType.Pie or ChartType.Doughnut);
        if (axes) sb.Append("<c:axId val=\"111\"/><c:axId val=\"222\"/>");
        sb.Append("</").Append(tag).Append('>');
        if (axes)
        {
            string ax1 = ch.Type == ChartType.Scatter ? "c:valAx" : "c:catAx";
            string pos1 = ch.Type == ChartType.Bar ? "l" : "b", pos2 = ch.Type == ChartType.Bar ? "b" : "l";
            sb.Append('<').Append(ax1).Append("><c:axId val=\"111\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"").Append(pos1).Append("\"/>")
              .Append(ch.Type == ChartType.Scatter ? "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>" : "").Append("<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"222\"/><c:crosses val=\"autoZero\"/>")
              .Append(ch.Type == ChartType.Scatter ? "<c:crossBetween val=\"midCat\"/>" : "<c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>").Append("</").Append(ax1).Append('>');
            sb.Append("<c:valAx><c:axId val=\"222\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"").Append(pos2).Append("\"/><c:majorGridlines/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>")
              .Append("<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"111\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>");
        }
        sb.Append("</c:plotArea>");
        if (ch.ShowLegend) sb.Append("<c:legend><c:legendPos val=\"").Append(ch.Type is ChartType.Pie or ChartType.Doughnut ? "r" : "b").Append("\"/><c:overlay val=\"0\"/></c:legend>");
        sb.Append("<c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>");
        return sb.ToString();
    }

    // ------------------------------------------------------------------ styles

    private static string StylesXml(Workbook wb, List<ConditionalFormat> dxfs)
    {
        var styles = wb.Styles.All;
        var numFmts = new Dictionary<string, int>();
        int nextFmt = 164;
        var fonts = new List<string>(); var fontIdx = new Dictionary<string, int>();
        var fills = new List<string> { "<fill><patternFill patternType=\"none\"/></fill>", "<fill><patternFill patternType=\"gray125\"/></fill>" };
        var fillIdx = new Dictionary<string, int>();
        var borders = new List<string> { "<border><left/><right/><top/><bottom/><diagonal/></border>" };
        var borderIdx = new Dictionary<string, int> { [borders[0]] = 0 };
        var xfs = new StringBuilder();
        int Font(CellStyle st)
        {
            var x = new StringBuilder("<font>");
            if (st.Bold) x.Append("<b/>");
            if (st.Italic) x.Append("<i/>");
            if (st.Strike) x.Append("<strike/>");
            if (st.Underline) x.Append("<u/>");
            x.Append("<sz val=\"").Append(F(st.FontSize)).Append("\"/>");
            if (st.FontColor != 0) x.Append("<color rgb=\"").Append(Hex(st.FontColor)).Append("\"/>"); else x.Append("<color theme=\"1\"/>");
            x.Append("<name val=\"").Append(Esc(st.FontName)).Append("\"/><family val=\"2\"/>");
            if (st.FontName == CellStyle.DefaultFont) x.Append("<scheme val=\"minor\"/>");
            x.Append("</font>");
            var key = x.ToString();
            if (!fontIdx.TryGetValue(key, out var i)) { i = fonts.Count; fonts.Add(key); fontIdx[key] = i; }
            return i;
        }
        int Fill(CellStyle st)
        {
            if (st.Fill == 0) return 0;
            var key = "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"" + Hex(st.Fill) + "\"/><bgColor indexed=\"64\"/></patternFill></fill>";
            if (!fillIdx.TryGetValue(key, out var i)) { i = fills.Count; fills.Add(key); fillIdx[key] = i; }
            return i;
        }
        string Side(string tag, BorderLine b)
        {
            if (b.IsNone) return "<" + tag + "/>";
            var style = b.Style switch { LineStyle.Medium => "medium", LineStyle.Thick => "thick", LineStyle.Dashed => "dashed", LineStyle.Dotted => "dotted", LineStyle.Double => "double", LineStyle.Hair => "hair", _ => "thin" };
            return "<" + tag + " style=\"" + style + "\"><color rgb=\"" + Hex(b.Argb == 0 ? Color.Black.ToArgb() : b.Argb) + "\"/></" + tag + ">";
        }
        int Border(CellStyle st)
        {
            var key = "<border>" + Side("left", st.Left) + Side("right", st.Right) + Side("top", st.Top) + Side("bottom", st.Bottom) + "<diagonal/></border>";
            if (!borderIdx.TryGetValue(key, out var i)) { i = borders.Count; borders.Add(key); borderIdx[key] = i; }
            return i;
        }
        Font(CellStyle.Default);
        foreach (var st in styles)
        {
            int numId;
            var bi = NumFmt.BuiltInId(st.NumberFormat);
            if (bi != null) numId = bi.Value;
            else if (!numFmts.TryGetValue(st.NumberFormat, out numId)) { numId = nextFmt++; numFmts[st.NumberFormat] = numId; }
            int fo = Font(st), fi = Fill(st), bo = Border(st);
            xfs.Append("<xf numFmtId=\"").Append(numId).Append("\" fontId=\"").Append(fo).Append("\" fillId=\"").Append(fi).Append("\" borderId=\"").Append(bo).Append("\" xfId=\"0\"");
            if (numId != 0) xfs.Append(" applyNumberFormat=\"1\"");
            if (fo != 0) xfs.Append(" applyFont=\"1\"");
            if (fi != 0) xfs.Append(" applyFill=\"1\"");
            if (bo != 0) xfs.Append(" applyBorder=\"1\"");
            bool align = st.HAlign != HAlign.General || st.VAlign != VAlign.Bottom || st.Wrap || st.Indent > 0;
            bool prot = !st.Locked || st.HiddenFormula;
            if (align) xfs.Append(" applyAlignment=\"1\"");
            if (prot) xfs.Append(" applyProtection=\"1\"");
            if (!align && !prot) { xfs.Append("/>"); continue; }
            xfs.Append('>');
            if (align)
            {
                xfs.Append("<alignment");
                if (st.HAlign != HAlign.General) xfs.Append(" horizontal=\"").Append(st.HAlign switch { HAlign.Left => "left", HAlign.Center => "center", HAlign.Right => "right", HAlign.Justify => "justify", _ => "centerContinuous" }).Append('"');
                if (st.VAlign != VAlign.Bottom) xfs.Append(" vertical=\"").Append(st.VAlign == VAlign.Center ? "center" : "top").Append('"');
                if (st.Wrap) xfs.Append(" wrapText=\"1\"");
                if (st.Indent > 0) xfs.Append(" indent=\"").Append(st.Indent).Append('"');
                xfs.Append("/>");
            }
            if (prot) xfs.Append("<protection locked=\"").Append(st.Locked ? 1 : 0).Append("\" hidden=\"").Append(st.HiddenFormula ? 1 : 0).Append("\"/>");
            xfs.Append("</xf>");
        }
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"" + NsMain + "\">");
        if (numFmts.Count > 0)
        {
            sb.Append("<numFmts count=\"").Append(numFmts.Count).Append("\">");
            foreach (var (code, id) in numFmts) sb.Append("<numFmt numFmtId=\"").Append(id).Append("\" formatCode=\"").Append(Esc(code)).Append("\"/>");
            sb.Append("</numFmts>");
        }
        sb.Append("<fonts count=\"").Append(fonts.Count).Append("\">").Append(string.Concat(fonts)).Append("</fonts>");
        sb.Append("<fills count=\"").Append(fills.Count).Append("\">").Append(string.Concat(fills)).Append("</fills>");
        sb.Append("<borders count=\"").Append(borders.Count).Append("\">").Append(string.Concat(borders)).Append("</borders>");
        sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
        sb.Append("<cellXfs count=\"").Append(styles.Count).Append("\">").Append(xfs).Append("</cellXfs>");
        sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
        sb.Append("<dxfs count=\"").Append(dxfs.Count).Append("\">");
        foreach (var cf in dxfs)
        {
            sb.Append("<dxf>");
            if (cf.FontColor != 0 || cf.Bold) { sb.Append("<font>"); if (cf.Bold) sb.Append("<b/>"); if (cf.FontColor != 0) sb.Append("<color rgb=\"").Append(Hex(cf.FontColor)).Append("\"/>"); sb.Append("</font>"); }
            if (cf.Fill != 0) sb.Append("<fill><patternFill><bgColor rgb=\"").Append(Hex(cf.Fill)).Append("\"/></patternFill></fill>");
            sb.Append("</dxf>");
        }
        sb.Append("</dxfs><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/></styleSheet>");
        return sb.ToString();
    }

    private const string ThemeXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\">"
        + "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>"
        + "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>"
        + "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme>"
        + "<a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>"
        + "<a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>"
        + "<a:lnStyleLst><a:ln w=\"6350\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>"
        + "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
        + "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>";
}

/// <summary>Splits a chart's data range into categories and series.</summary>
public static class ChartData
{
    public static readonly Color[] Palette =
    [
        Color.FromArgb(68, 114, 196), Color.FromArgb(237, 125, 49), Color.FromArgb(165, 165, 165), Color.FromArgb(255, 192, 0),
        Color.FromArgb(91, 155, 213), Color.FromArgb(112, 173, 71), Color.FromArgb(38, 68, 120), Color.FromArgb(158, 72, 14)
    ];

    public sealed record Series(CellPos? NameCell, CellRange Values);

    /// <summary>First row = series names, first column = categories (or transposed when SeriesInRows).</summary>
    public static (CellRange? Categories, List<Series> Series) Layout(Chart ch)
    {
        var d = ch.Data;
        var list = new List<Series>();
        if (!ch.SeriesInRows)
        {
            bool hasCats = d.Cols > 1;
            bool hasNames = d.Rows > 1;
            int firstRow = hasNames ? d.R1 + 1 : d.R1;
            CellRange? cats = hasCats ? new CellRange(firstRow, d.C1, d.R2, d.C1) : null;
            for (int c = hasCats ? d.C1 + 1 : d.C1; c <= d.C2; c++)
                list.Add(new Series(hasNames ? new CellPos(d.R1, c) : null, new CellRange(firstRow, c, d.R2, c)));
            return (cats, list);
        }
        else
        {
            bool hasCats = d.Rows > 1;
            bool hasNames = d.Cols > 1;
            int firstCol = hasNames ? d.C1 + 1 : d.C1;
            CellRange? cats = hasCats ? new CellRange(d.R1, firstCol, d.R1, d.C2) : null;
            for (int r = hasCats ? d.R1 + 1 : d.R1; r <= d.R2; r++)
                list.Add(new Series(hasNames ? new CellPos(r, d.C1) : null, new CellRange(r, firstCol, r, d.C2)));
            return (cats, list);
        }
    }
}
