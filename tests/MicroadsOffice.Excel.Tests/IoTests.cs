using System.IO.Compression;
using System.Text;
using MicroadsOffice.Excel;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.IO;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

public static class IoTests
{
    public static Workbook Sample()
    {
        var wb = Workbook.CreateBlank(2);
        var s = wb.Sheets[0];
        s.Name = "Bán hàng";
        SheetOps.SetInput(s, 0, 0, "Sản phẩm"); SheetOps.SetInput(s, 0, 1, "Quý 1"); SheetOps.SetInput(s, 0, 2, "Quý 2"); SheetOps.SetInput(s, 0, 3, "Tổng");
        string[] p = ["Bàn phím", "Chuột", "Màn hình", "Tai nghe"];
        double[] q1 = [120, 340, 80, 150], q2 = [140, 310, 95, 170];
        for (int i = 0; i < 4; i++)
        {
            s.SetValue(i + 1, 0, CellValue.Text(p[i]));
            s.SetValue(i + 1, 1, CellValue.Number(q1[i]));
            s.SetValue(i + 1, 2, CellValue.Number(q2[i]));
            s.SetFormula(i + 1, 3, $"SUM(B{i + 2}:C{i + 2})");
        }
        s.SetFormula(5, 3, "SUM(D2:D5)");
        SheetOps.SetInput(s, 5, 0, "Cộng");
        var head = wb.Styles.Intern(CellStyle.Default with { Bold = true, Fill = System.Drawing.Color.FromArgb(226, 239, 218).ToArgb(), HAlign = HAlign.Center, Bottom = BorderLine.Thin() });
        for (int c = 0; c < 4; c++) s.SetStyle(0, c, head);
        var money = wb.Styles.Intern(CellStyle.Default with { NumberFormat = "#,##0 \"₫\"", FontColor = System.Drawing.Color.FromArgb(31, 78, 121).ToArgb(), Italic = true });
        s.SetStyle(5, 3, money);
        s.ColWidths[0] = 140; s.ColWidths[3] = 90; s.RowHeights[0] = 28; s.HiddenCols.Add(7);
        s.FreezeRows = 1; s.FreezeCols = 1;
        s.Merges.Add(new CellRange(7, 0, 7, 3));
        SheetOps.SetInput(s, 7, 0, "Báo cáo doanh số — Microads");
        s.CondFormats.Add(new ConditionalFormat { Range = new CellRange(1, 3, 4, 3), Type = CfType.GreaterThan, Value1 = "300", Fill = System.Drawing.Color.FromArgb(255, 199, 206).ToArgb(), FontColor = System.Drawing.Color.FromArgb(156, 0, 6).ToArgb() });
        s.CondFormats.Add(new ConditionalFormat { Range = new CellRange(1, 1, 4, 1), Type = CfType.DataBar });
        s.Validations.Add(new DataValidation { Range = new CellRange(1, 4, 4, 4), Type = ValidationType.List, Formula1 = "Đạt,Chưa đạt" });
        s.Notes[Sheet.Key(1, 0)] = "Hàng bán chạy";
        s.Links[Sheet.Key(9, 0)] = "https://microads.vn";
        SheetOps.SetInput(s, 9, 0, "microads.vn");
        s.Charts.Add(new Chart { Type = ChartType.Column, Title = "Doanh số theo quý", Data = new CellRange(0, 0, 4, 2), Bounds = new System.Drawing.RectangleF(420, 30, 420, 260) });
        s.Page.Landscape = true; s.Page.FitToWidth = true; s.Page.PrintArea = new CellRange(0, 0, 5, 3); s.Page.PrintTitleRows = (0, 0); s.Page.PrintGridlines = true;
        s.TabColor = System.Drawing.Color.FromArgb(43, 168, 74).ToArgb();
        SheetOps.SetInput(s, 11, 0, "15/03/2024");
        SheetOps.SetInput(s, 11, 1, "=DOCSO(D6)");
        wb.Names["TongCong"] = "'Bán hàng'!$D$6";
        var s2 = wb.Sheets[1];
        s2.Name = "Ghi chú";
        s2.SetFormula(0, 0, "TongCong*2");
        s2.SetValue(1, 0, CellValue.Bool(true));
        s2.SetValue(2, 0, CellValue.Error(Errors.NA));
        wb.Recalculate();
        return wb;
    }

    public static void Run(TestRunner t, string outDir)
    {
        Format.NumFmt.UseLocale(true);
        t.Test("XLSX round-trip (values, formulas, styles, layout, CF, DV, notes, links, charts, print)", () =>
        {
            var wb = Sample();
            var path = Path.Combine(outDir, "roundtrip.xlsx");
            XlsxWriter.Write(wb, path);
            var r = XlsxReader.Read(path);
            var a = wb.Sheets[0];
            var b = r.Sheets[0];
            T.Eq("Bán hàng", b.Name); T.Eq("Ghi chú", r.Sheets[1].Name);
            foreach (var (k, cell) in a.AllCells)
            {
                var (row, col) = Sheet.FromKey(k);
                var other = b.GetCell(row, col) ?? throw new AssertException("missing cell " + Addr.ToA1(row, col));
                T.Eq(cell.Formula, other.Formula, "formula " + Addr.ToA1(row, col));
                T.True(cell.Value == other.Value, $"value {Addr.ToA1(row, col)}: {cell.Value} vs {other.Value}");
                T.True(wb.Styles[cell.StyleId] == r.Styles[other.StyleId], "style " + Addr.ToA1(row, col) + ": " + wb.Styles[cell.StyleId] + " vs " + r.Styles[other.StyleId]);
            }
            T.Eq(1405.0, b.GetValue(5, 3).Num, "total");
            T.Eq(140.0, b.ColWidths[0]); T.Eq(28.0, b.RowHeights[0]); T.True(b.HiddenCols.Contains(7), "hidden col");
            T.Eq(1, b.FreezeRows); T.Eq(1, b.FreezeCols);
            T.True(b.Merges.Contains(new CellRange(7, 0, 7, 3)), "merge");
            T.Eq(2, b.CondFormats.Count); T.Eq(CfType.GreaterThan, b.CondFormats[0].Type); T.Eq("300", b.CondFormats[0].Value1); T.Eq(a.CondFormats[0].Fill, b.CondFormats[0].Fill);
            T.Eq(CfType.DataBar, b.CondFormats[1].Type);
            T.Eq(1, b.Validations.Count); T.Eq("Đạt,Chưa đạt", b.Validations[0].Formula1); T.Eq(ValidationType.List, b.Validations[0].Type);
            T.Eq("Hàng bán chạy", b.Notes[Sheet.Key(1, 0)]);
            T.Eq("https://microads.vn", b.Links[Sheet.Key(9, 0)]);
            T.Eq(1, b.Charts.Count); T.Eq(ChartType.Column, b.Charts[0].Type); T.Eq("Doanh số theo quý", b.Charts[0].Title); T.Eq(new CellRange(0, 0, 4, 2), b.Charts[0].Data);
            T.True(Math.Abs(b.Charts[0].Bounds.X - 420) < 2 && Math.Abs(b.Charts[0].Bounds.Width - 420) < 2, "chart bounds " + b.Charts[0].Bounds);
            T.True(b.Page.Landscape && b.Page.FitToWidth && b.Page.PrintGridlines, "page setup");
            T.Eq(new CellRange(0, 0, 5, 3), b.Page.PrintArea!.Value); T.Eq((0, 0), b.Page.PrintTitleRows!.Value);
            T.Eq(a.TabColor, b.TabColor);
            T.Eq("'Bán hàng'!$D$6", r.Names["TongCong"]);
            T.Eq(2810.0, r.Sheets[1].GetValue(0, 0).Num, "defined name formula");
            T.True(r.Sheets[1].GetValue(2, 0).IsError, "error value");
            T.True(b.GetValue(11, 1).Str!.StartsWith("Một nghìn bốn trăm linh năm"), "DOCSO result " + b.GetValue(11, 1));
            T.Eq("dd/mm/yyyy", b.StyleAt(11, 0).NumberFormat);
            // structural validity: every part is well-formed XML
            using var z = ZipFile.OpenRead(path);
            foreach (var e in z.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels")))
            {
                using var st = e.Open();
                System.Xml.Linq.XDocument.Load(st);
            }
            T.True(z.GetEntry("xl/charts/chart1.xml") != null && z.GetEntry("xl/comments1.xml") != null && z.GetEntry("xl/drawings/vmlDrawing1.vml") != null, "parts present");
        });
        t.Test("XLSX hand-crafted file: shared strings, shared formulas, inline strings, theme colours", () =>
        {
            var path = Path.Combine(outDir, "handmade.xlsx");
            WriteHandmade(path);
            var wb = XlsxReader.Read(path);
            var s = wb.Sheets[0];
            T.Eq("Hà Nội", s.GetValue(0, 0).Str);
            T.Eq("Rich text", s.GetValue(0, 1).Str);
            T.Eq("A2*2", s.GetCell(1, 1)!.Formula);
            T.Eq("A3*2", s.GetCell(2, 1)!.Formula, "shared formula expanded");
            T.Eq("A4*2", s.GetCell(3, 1)!.Formula);
            T.Eq(60.0, s.GetValue(3, 1).Num, "calculated");
            T.Eq("inline", s.GetValue(4, 0).Str);
            T.True(s.StyleAt(0, 0).Bold, "bold font");
            T.Eq(System.Drawing.Color.FromArgb(0x44, 0x72, 0xC4).ToArgb(), s.StyleAt(0, 0).Fill, "theme fill colour");
            T.Eq("0.00%", s.StyleAt(1, 0).NumberFormat);
        });
        t.Test("CSV round-trip with Vietnamese text and delimiter detection", () =>
        {
            var wb = Workbook.CreateBlank();
            var s = wb.Sheets[0];
            SheetOps.SetInput(s, 0, 0, "Họ tên"); SheetOps.SetInput(s, 0, 1, "Ghi chú");
            SheetOps.SetInput(s, 1, 0, "Nguyễn Thị Ánh"); SheetOps.SetInput(s, 1, 1, "có dấu \"phẩy\", và xuống\ndòng");
            s.SetValue(2, 0, CellValue.Number(1234.5));
            var path = Path.Combine(outDir, "data.csv");
            CsvIO.Write(s, path, ',');
            var bytes = File.ReadAllBytes(path);
            T.True(bytes[0] == 0xEF && bytes[1] == 0xBB, "UTF-8 BOM");
            var r = CsvIO.Read(path);
            var rs = r.Sheets[0];
            T.Eq("Nguyễn Thị Ánh", rs.GetValue(1, 0).Str);
            T.Eq("có dấu \"phẩy\", và xuống\ndòng", rs.GetValue(1, 1).Str);
            T.Eq(1234.5, rs.GetValue(2, 0).Num);
            File.WriteAllText(Path.Combine(outDir, "semi.csv"), "a;b;c\n1;2;3\n", new UTF8Encoding(false));
            var r2 = CsvIO.Read(Path.Combine(outDir, "semi.csv"));
            T.Eq(3.0, r2.Sheets[0].GetValue(1, 2).Num, "semicolon detected");
            var tsv = Path.Combine(outDir, "data.tsv");
            CsvIO.Write(s, tsv, '\t');
            T.Eq("Nguyễn Thị Ánh", CsvIO.Read(tsv).Sheets[0].GetValue(1, 0).Str);
        });
        t.Test("HTML export", () =>
        {
            var path = Path.Combine(outDir, "export.html");
            CsvIO.WriteHtml(Sample(), path);
            var html = File.ReadAllText(path);
            T.True(html.Contains("Bàn phím") && html.Contains("colspan=\"4\"") && html.Contains("microads.vn"), "html content");
        });
    }

    private static void WriteHandmade(string path)
    {
        using var fs = File.Create(path);
        using var z = new ZipArchive(fs, ZipArchiveMode.Create);
        void Put(string n, string c) { using var w = new StreamWriter(z.CreateEntry(n).Open(), new UTF8Encoding(false)); w.Write(c); }
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Put("[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>");
        Put("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        Put("xl/workbook.xml", $"<?xml version=\"1.0\"?><workbook xmlns=\"{ns}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
        Put("xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\"/></Relationships>");
        Put("xl/sharedStrings.xml", $"<?xml version=\"1.0\"?><sst xmlns=\"{ns}\"><si><t>Hà Nội</t></si><si><r><rPr><b/></rPr><t>Rich </t></r><r><t>text</t></r></si></sst>");
        Put("xl/styles.xml", $"<?xml version=\"1.0\"?><styleSheet xmlns=\"{ns}\"><fonts><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
            + "<fills><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"4\"/></patternFill></fill></fills>"
            + "<borders><border/></borders><cellXfs><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\"/><xf numFmtId=\"10\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs></styleSheet>");
        Put("xl/worksheets/sheet1.xml", $"<?xml version=\"1.0\"?><worksheet xmlns=\"{ns}\"><sheetData>"
            + "<row r=\"1\"><c r=\"A1\" t=\"s\" s=\"1\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\" s=\"2\"><v>10</v></c><c r=\"B2\"><f t=\"shared\" ref=\"B2:B4\" si=\"0\">A2*2</f><v>20</v></c></row>"
            + "<row r=\"3\"><c r=\"A3\"><v>20</v></c><c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>40</v></c></row>"
            + "<row r=\"4\"><c r=\"A4\"><v>30</v></c><c r=\"B4\"><f t=\"shared\" si=\"0\"/></c></row>"
            + "<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>inline</t></is></c></row>"
            + "</sheetData></worksheet>");
    }
}
