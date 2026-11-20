using System.IO.Compression;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.IO;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

/// <summary>Dynamic arrays (spill), LET/LAMBDA, array functions, element-wise evaluation and their .xlsx form.</summary>
public static class EngineTestsDyn
{
    private static Workbook NewData()
    {
        var wb = Workbook.CreateBlank(1);
        var s = wb.Sheets[0];
        string[] fruits = ["apple", "banana", "cherry", "date", "elder"];
        for (int i = 0; i < 5; i++)
        {
            s.SetValue(i, 0, CellValue.Number(i + 1));
            s.SetValue(i, 1, CellValue.Number((i + 1) * 10));
            s.SetValue(i, 2, CellValue.Text(fruits[i]));
        }
        wb.Names["ADD1"] = "LAMBDA(x,x+1)";
        wb.Names["FACTR"] = "LAMBDA(n,IF(n<=1,1,n*FACTR(n-1)))";
        wb.Recalculate();
        return wb;
    }

    public static void Run(TestRunner t)
    {
        NumFmt.UseLocale(false);
        var wb = NewData();
        CellValue Ev(string f) => wb.Calc.EvaluateFormula(wb.Sheets[0], 30, 30, f);
        void N(string f, double expected)
        {
            var v = Ev(f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > 1e-9 * Math.Max(1, Math.Abs(expected))) throw new AssertException($"{f} => {v.Kind}:{v} expected {expected}");
        }
        void S(string f, string expected) { var v = Ev(f); if (v.ToString() != expected) throw new AssertException($"{f} => {v.Kind}:'{v}' expected '{expected}'"); }
        void E(string f, string code) { var v = Ev(f); if (!v.IsError || v.Str != code) throw new AssertException($"{f} => {v} expected {code}"); }

        t.Test("Catalog lists every function (Insert Function, menus, autocomplete)", () =>
        {
            var noEntry = Functions.Table.Keys.Where(k => FunctionCatalog.Find(k) == null && k is not ("TRUE" or "FALSE")).OrderBy(k => k).ToList();
            var noImpl = FunctionCatalog.All.Where(f => !Functions.Table.ContainsKey(f.Name)).Select(f => f.Name).ToList();
            T.True(noEntry.Count == 0, "functions without catalog entry: " + string.Join(" ", noEntry));
            T.True(noImpl.Count == 0, "catalog entries without implementation: " + string.Join(" ", noImpl));
            T.True(FunctionCatalog.All.All(f => f.DescVi.Length > 0 && f.DescEn.Length > 0), "descriptions");
            t.Log($"      functions: {Functions.Table.Count} implemented, {FunctionCatalog.All.Count} in the catalog");
        });
        t.Test("AI formula check accepts LET/LAMBDA variables, rejects unknown functions", () =>
        {
            MicroadsOffice.Excel.AI.SheetActions.CheckFormula("=LET(f,LAMBDA(x,x*2),f(3))");
            MicroadsOffice.Excel.AI.SheetActions.CheckFormula("=SUM(FILTER(A2:A9,B2:B9>0))+SUM(D2#)");
            MicroadsOffice.Excel.AI.SheetActions.CheckFormula("=LAMBDA(a,[b],a)(1)");
            bool threw = false;
            try { MicroadsOffice.Excel.AI.SheetActions.CheckFormula("=NOSUCHFN(1)"); } catch (FormatException) { threw = true; }
            T.True(threw, "unknown function rejected");
            T.True(MicroadsOffice.Excel.AI.SheetActions.FunctionList().Contains("XLOOKUP, XMATCH"), "AI function list");
        });
        t.Test("Recalculation speed with element-wise functions (20 000 formulas)", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            for (int i = 0; i < 10000; i++)
            {
                s.SetValue(i, 0, CellValue.Number(i * 1.37));
                s.SetValue(i, 1, CellValue.Text("item" + i));
                s.SetFormula(i, 2, $"ROUND(A{i + 1}*1.1,2)+LEN(B{i + 1})");
                s.SetFormula(i, 3, $"IF(C{i + 1}>100,LEFT(B{i + 1},3),UPPER(B{i + 1}))");
            }
            var sw = System.Diagnostics.Stopwatch.StartNew();
            w.Recalculate();
            sw.Stop();
            T.Eq(Math.Round(9999 * 1.37 * 1.1, 2) + 8, s.GetValue(9999, 2).Num, "value");
            t.Log($"      recalculation of 20 000 formulas: {sw.ElapsedMilliseconds} ms");
            T.True(sw.ElapsedMilliseconds < 5000, "fast enough");
        });
        t.Test("LET / LAMBDA", () =>
        {
            N("LET(x,2,y,3,x*y)", 6); N("LET(x,2,LET(x,5,x)+x)", 7); N("LET(a,A1:A5,SUM(a*2))", 30);
            N("LAMBDA(x,x+1)(5)", 6); N("LAMBDA(x,y,x*y)(3,4)", 12); N("ADD1(4)", 5); N("FACTR(5)", 120);
            N("LET(f,LAMBDA(x,x*x),f(7))", 49); E("LAMBDA(x,x)", Errors.Calc); E("LAMBDA(x,y,x+y)(1)", Errors.Value);
            N("LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1)", 1); N("LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1,2)", 3);
            E("LET(x,1)", Errors.Value);
        });
        t.Test("MAP / REDUCE / SCAN / BYROW / BYCOL / MAKEARRAY", () =>
        {
            N("SUM(MAP(A1:A5,LAMBDA(x,x*2)))", 30); N("SUM(MAP(A1:A5,B1:B5,LAMBDA(x,y,x*y)))", 550);
            N("REDUCE(0,A1:A5,LAMBDA(a,v,a+v))", 15); N("REDUCE(1,A1:A5,LAMBDA(a,v,a*v))", 120);
            N("INDEX(SCAN(0,A1:A5,LAMBDA(a,v,a+v)),3)", 6); N("INDEX(SCAN(0,A1:A5,LAMBDA(a,v,a+v)),5)", 15);
            N("INDEX(BYROW(A1:B5,LAMBDA(r,SUM(r))),2)", 22); N("INDEX(BYCOL(A1:B5,SUM),1,2)", 150); N("ROWS(BYROW(A1:B5,SUM))", 5);
            N("INDEX(MAKEARRAY(2,3,LAMBDA(r,c,r*10+c)),2,3)", 23); N("COLUMNS(MAKEARRAY(2,3,LAMBDA(r,c,0)))", 3);
        });
        t.Test("FILTER / SORT / SORTBY / UNIQUE", () =>
        {
            N("SUM(FILTER(B1:B5,A1:A5>2))", 120); N("ROWS(FILTER(A1:C5,A1:A5>3))", 2); S("INDEX(FILTER(A1:C5,A1:A5>3),1,3)", "date");
            S("FILTER(A1:A5,A1:A5>9,\"none\")", "none"); E("FILTER(A1:A5,A1:A5>9)", Errors.Calc); E("FILTER(A1:A5,{1;0})", Errors.Value);
            N("SUM(FILTER(A1:A5,(A1:A5>1)*(A1:A5<5)))", 9);
            N("INDEX(SORT(A1:A5,1,-1),1)", 5); S("INDEX(SORT(C1:C5,1,-1),1)", "elder"); S("INDEX(SORT(A1:C5,3,-1),1,3)", "elder");
            N("INDEX(SORT({3,1,2},1,1,TRUE),1,1)", 1); S("INDEX(SORTBY(C1:C5,B1:B5,-1),1)", "elder");
            N("ROWS(UNIQUE({1;2;2;3;1}))", 3); N("UNIQUE({1;2;2;3;1},FALSE,TRUE)", 3); N("ROWS(UNIQUE({\"a\";\"A\";\"b\"}))", 2);
            N("COLUMNS(UNIQUE({1,1,2},TRUE))", 2);
        });
        t.Test("Array shaping: VSTACK, HSTACK, TAKE, DROP, CHOOSE*, TOCOL, WRAP*, EXPAND", () =>
        {
            N("ROWS(VSTACK({1,2},{3,4}))", 2); N("INDEX(VSTACK({1,2},{3,4}),2,1)", 3); E("INDEX(VSTACK({1,2},{3}),2,2)", Errors.NA);
            N("COLUMNS(HSTACK({1;2},{3;4}))", 2); N("INDEX(HSTACK({1;2},{3;4}),2,2)", 4);
            N("SUM(TAKE(A1:A5,2))", 3); N("SUM(TAKE(A1:A5,-2))", 9); N("SUM(DROP(A1:A5,1))", 14); N("DROP(A1:A5,-4)", 1); E("DROP(A1:A5,5)", Errors.Calc);
            S("INDEX(CHOOSECOLS(A1:C5,-1),1,1)", "apple"); N("SUM(CHOOSEROWS(A1:A5,{1,5}))", 6); E("CHOOSECOLS(A1:C5,4)", Errors.Value);
            N("INDEX(TOCOL({1,2;3,4}),3)", 3); N("INDEX(TOCOL({1,2;3,4},0,TRUE),2)", 3); N("COLUMNS(TOROW({1,2;3,4}))", 4);
            N("ROWS(WRAPROWS(SEQUENCE(5),2))", 3); E("INDEX(WRAPROWS(SEQUENCE(5),2),3,2)", Errors.NA); N("INDEX(WRAPCOLS(SEQUENCE(5),2),1,3)", 5); E("INDEX(WRAPCOLS(SEQUENCE(5),2),2,3)", Errors.NA);
            N("SUM(EXPAND({1,2},2,3,0))", 3); N("ROWS(EXPAND({1,2},2,3,0))", 2); E("EXPAND({1,2},1,1)", Errors.Value);
        });
        t.Test("Element-wise evaluation of single-value functions", () =>
        {
            N("SUMPRODUCT(LEN(C1:C5))", 26); N("SUMPRODUCT(--(LEFT(C1:C5,1)=\"b\"))", 1); N("SUM(COUNTIF(C1:C5,{\"apple\",\"date\"}))", 2);
            N("SUM(LARGE(A1:A5,{1,2}))", 9); N("SUM(IFERROR(1/(A1:A5-3),0))", 0); N("SUM(ROW(A2:A4))", 9); N("SUM(COLUMN(A1:C1))", 6);
            N("INDEX(CHOOSE({1,2},A1:A5,B1:B5),2,2)", 20); N("VLOOKUP(\"cherry\",CHOOSE({1,2},C1:C5,A1:A5),2,0)", 3);
            N("SUM(ROUND(A1:A5/3,0))", 5); N("ROWS(UPPER(C1:C5))", 5); S("INDEX(UPPER(C1:C5),2)", "BANANA");
            N("SUM(--ISNUMBER(SEARCH(\"an\",C1:C5)))", 1); N("SUM(MATCH({\"date\",\"apple\"},C1:C5,0))", 5);
            N("INDEX({1,2;3,4},0,2)", 2); N("SUM(INDEX({1,2;3,4},0,2))", 6); N("SUM(INDEX({1,2;3,4},2,0))", 7);
        });
        t.Test("GROUPBY / PIVOTBY", () =>
        {
            var g = wb.Sheets[0];
            string[] reg = ["Bắc", "Nam", "Bắc", "Trung", "Nam", "Bắc"];
            double[] amt = [10, 20, 30, 40, 50, 60];
            string[] q = ["Q1", "Q1", "Q2", "Q2", "Q1", "Q2"];
            for (int i = 0; i < 6; i++) { g.SetValue(10 + i, 0, CellValue.Text(reg[i])); g.SetValue(10 + i, 1, CellValue.Number(amt[i])); g.SetValue(10 + i, 2, CellValue.Text(q[i])); }
            wb.Recalculate();
            N("ROWS(GROUPBY(A11:A16,B11:B16,SUM))", 4); S("INDEX(GROUPBY(A11:A16,B11:B16,SUM),1,1)", "Bắc"); N("INDEX(GROUPBY(A11:A16,B11:B16,SUM),1,2)", 100);
            S("INDEX(GROUPBY(A11:A16,B11:B16,SUM),4,1)", "Total"); N("INDEX(GROUPBY(A11:A16,B11:B16,SUM),4,2)", 210);
            N("ROWS(GROUPBY(A11:A16,B11:B16,SUM,0,0))", 3); N("INDEX(GROUPBY(A11:A16,B11:B16,SUM,0,0,-2),1,2)", 100);
            N("INDEX(GROUPBY(A11:A16,B11:B16,COUNT,0,0),2,2)", 2); N("INDEX(GROUPBY(A11:A16,B11:B16,LAMBDA(x,MAX(x)),0,0),3,2)", 40);
            // PIVOTBY: header row (Q1, Q2, Total) then Bắc / Nam / Trung / Total
            N("ROWS(PIVOTBY(A11:A16,C11:C16,B11:B16,SUM))", 5); S("INDEX(PIVOTBY(A11:A16,C11:C16,B11:B16,SUM),1,2)", "Q1");
            N("INDEX(PIVOTBY(A11:A16,C11:C16,B11:B16,SUM),2,2)", 10); N("INDEX(PIVOTBY(A11:A16,C11:C16,B11:B16,SUM),2,3)", 90);
            N("INDEX(PIVOTBY(A11:A16,C11:C16,B11:B16,SUM),5,4)", 210);
            for (int i = 0; i < 6; i++) for (int k = 0; k < 3; k++) g.ClearCell(10 + i, k, true, false);
            wb.Recalculate();
        });
        t.Test("Spill into neighbouring cells, #SPILL!, A1# references", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            s.SetFormula(20, 10, "SEQUENCE(5)");          // K21
            s.SetFormula(5, 5, "SUM(K21:K25)");            // above the formula: read before it is calculated
            s.SetFormula(5, 6, "SUM(K21#)");
            s.SetFormula(5, 7, "ROWS(K21#)");
            w.Recalculate();
            T.Eq(5.0, s.GetValue(24, 10).Num, "K25");
            T.Eq(15.0, s.GetValue(5, 5).Num, "SUM over spilled cells");
            T.Eq(15.0, s.GetValue(5, 6).Num, "SUM(K21#)");
            T.Eq(5.0, s.GetValue(5, 7).Num, "ROWS(K21#)");
            T.True(s.GetCell(22, 10)!.IsSpilled && s.SpillSource(22, 10)?.Row == 20, "K23 spilled from K21");
            T.Eq("SEQUENCE(5)", s.SpillSource(22, 10)!.Value.Cell.Formula, "source formula");
            // shrink: old values disappear
            s.SetFormula(20, 10, "SEQUENCE(3,1,10)");
            w.Recalculate();
            T.Eq(33.0, s.GetValue(5, 5).Num, "SUM after shrink");
            T.True(s.GetCell(24, 10) == null, "K25 cleared");
            // grow
            s.SetFormula(20, 10, "SEQUENCE(6,1,10)");
            w.Recalculate();
            T.Eq(60.0, s.GetValue(5, 5).Num, "SUM after growth (first five)");
            T.Eq(75.0, s.GetValue(5, 6).Num, "SUM(K21#) after growth");
            // typing into the spill range blocks it
            s.SetValue(22, 10, CellValue.Text("x"));
            w.Recalculate();
            T.Eq(Errors.Spill, s.GetValue(20, 10).Str, "#SPILL!");
            T.Eq(Errors.Spill, s.GetValue(5, 6).Str, "SUM(K21#) of blocked");
            T.True(s.GetCell(21, 10) == null, "blocked formula spills nothing");
            s.ClearCell(22, 10, true, false);
            w.Recalculate();
            T.Eq(75.0, s.GetValue(5, 6).Num, "spills again after clearing");
            // deleting the formula removes its values
            s.ClearCell(20, 10, true, false);
            w.Recalculate();
            T.True(s.GetCell(21, 10) == null && s.GetValue(5, 5).Num == 0, "values removed with the formula");
            // a single value does not spill; a range reference spills
            s.SetValue(0, 0, CellValue.Number(1)); s.SetValue(1, 0, CellValue.Number(2));
            s.SetFormula(0, 2, "A1:A2*10");
            w.Recalculate();
            T.Eq(20.0, s.GetValue(1, 2).Num, "A1:A2*10 spills");
            // undo snapshot (sheet clone) keeps the spill links
            var clone = s.Clone(w);
            T.True(clone.SpillSource(1, 2)?.Cell == clone.GetCell(0, 2), "clone keeps spill source");
        });
        t.Test("Copy/paste of a spill range: formula spills again, Paste Values keeps every value", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            s.SetFormula(0, 0, "SEQUENCE(4,2)");
            w.Recalculate();
            var clip = SheetOps.Copy(s, new CellRange(0, 0, 3, 1), false);
            SheetOps.Paste(s, clip, CellRange.Single(0, 4));
            w.Recalculate();
            T.Eq("SEQUENCE(4,2)", s.GetCell(0, 4)!.Formula, "formula pasted");
            T.Eq(8.0, s.GetValue(3, 5).Num, "pasted formula spills (no #SPILL!)");
            T.True(s.GetCell(3, 5)!.IsSpilled, "pasted area is a spill range");
            SheetOps.Paste(s, clip, CellRange.Single(10, 0), PasteMode.Values);
            w.Recalculate();
            T.Eq(8.0, s.GetValue(13, 1).Num, "values pasted");
            T.True(s.GetCell(13, 1)!.Formula == null && !s.GetCell(13, 1)!.IsSpilled, "plain values");
            // pasting values over a spill range blocks it
            SheetOps.Paste(s, SheetOps.Copy(s, CellRange.Single(10, 0), false), CellRange.Single(2, 4), PasteMode.Values);
            w.Recalculate();
            T.Eq(Errors.Spill, s.GetValue(0, 4).Str, "#SPILL! after pasting into the range");
        });
        t.Test("Dynamic arrays and new functions in .xlsx", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            for (int i = 0; i < 4; i++) s.SetValue(i, 0, CellValue.Number(i + 1));
            s.SetFormula(0, 2, "SEQUENCE(3)");
            s.SetFormula(0, 3, "LET(x,A1:A4,SUM(x*2))");
            s.SetFormula(0, 4, "FILTER(A1:A4,A1:A4>2)");
            s.SetFormula(0, 5, "XLOOKUP(3,A1:A4,A1:A4)");
            s.SetFormula(0, 6, "SUM(C1#)");
            s.SetFormula(0, 7, "SUM(A1:A4)");
            w.Names["SQ"] = "LAMBDA(v,v*v)";
            w.Recalculate();
            using var ms = new MemoryStream();
            XlsxWriter.Write(w, ms);
            ms.Position = 0;
            string sheetXml, wbXml;
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true))
            {
                sheetXml = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()).ReadToEnd();
                wbXml = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open()).ReadToEnd();
                T.True(zip.GetEntry("xl/metadata.xml") != null, "metadata.xml");
                var ct = new StreamReader(zip.GetEntry("[Content_Types].xml")!.Open()).ReadToEnd();
                T.True(ct.Contains("sheetMetadata+xml"), "metadata content type");
            }
            T.True(sheetXml.Contains("<c r=\"C1\" cm=\"1\"><f t=\"array\" ref=\"C1:C3\">_xlfn.SEQUENCE(3)</f>"), "dynamic array cell: " + sheetXml);
            T.True(sheetXml.Contains("_xlfn.LET(_xlpm.x,A1:A4,SUM(_xlpm.x*2))"), "LET prefixes");
            T.True(sheetXml.Contains("_xlfn._xlws.FILTER(A1:A4,A1:A4&gt;2)"), "FILTER prefix");
            T.True(sheetXml.Contains("_xlfn.XLOOKUP("), "XLOOKUP prefix");
            T.True(sheetXml.Contains("SUM(_xlfn.ANCHORARRAY(C1))"), "spill reference");
            T.True(sheetXml.Contains("<f>SUM(A1:A4)</f>"), "plain formula unchanged");
            T.True(wbXml.Contains("_xlfn.LAMBDA(_xlpm.v,_xlpm.v*_xlpm.v)"), "LAMBDA name prefixes");
            ms.Position = 0;
            var back = XlsxReader.Read(ms);
            var b = back.Sheets[0];
            T.Eq("SEQUENCE(3)", b.GetCell(0, 2)!.Formula, "prefix removed");
            T.Eq("LET(x,A1:A4,SUM(x*2))", b.GetCell(0, 3)!.Formula, "LET read back");
            T.Eq("SUM(C1#)", b.GetCell(0, 6)!.Formula, "ANCHORARRAY read back");
            T.Eq("LAMBDA(v,v*v)", back.Names["SQ"], "name read back");
            T.Eq(3.0, b.GetValue(2, 2).Num, "spilled after load");
            T.Eq(6.0, b.GetValue(0, 6).Num, "SUM(C1#) after load");
            T.Eq(4.0, b.GetValue(1, 4).Num, "FILTER spilled after load");
            T.Eq(20.0, b.GetValue(0, 3).Num, "LET after load");
            T.Eq(9.0, back.Calc.EvaluateFormula(b, 9, 9, "SQ(3)").Num, "LAMBDA name after load");
        });
        t.Test("Legacy formulas (no dynamic-array flag) keep implicit intersection", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            for (int i = 0; i < 3; i++) s.SetValue(i, 0, CellValue.Number((i + 1) * 100));
            s.SetFormula(1, 2, "A1:A3");
            w.Recalculate();
            using var ms = new MemoryStream();
            XlsxWriter.Write(w, ms);
            // rewrite the file as an Excel 2016 file would store it: no cm/array attributes
            ms.Position = 0;
            var copy = new MemoryStream();
            using (var src = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true))
            using (var dst = new ZipArchive(copy, ZipArchiveMode.Create, leaveOpen: true))
                foreach (var e in src.Entries)
                {
                    var text = new StreamReader(e.Open()).ReadToEnd();
                    if (e.FullName == "xl/worksheets/sheet1.xml")
                        text = System.Text.RegularExpressions.Regex.Replace(text, "<c r=\"C2\"[^>]*>.*?</c>", "<c r=\"C2\"><f>A1:A3</f><v>200</v></c>");
                    if (e.FullName == "xl/worksheets/sheet1.xml")
                        text = System.Text.RegularExpressions.Regex.Replace(text, "<c r=\"C[34]\"[^>]*>.*?</c>", "");
                    using var wr = new StreamWriter(dst.CreateEntry(e.FullName).Open());
                    wr.Write(text);
                }
            copy.Position = 0;
            var back = XlsxReader.Read(copy);
            var b = back.Sheets[0];
            T.Eq(200.0, b.GetValue(1, 2).Num, "implicit intersection in row 2");
            T.True(b.GetCell(2, 2) == null, "nothing spilled");
            b.SetFormula(1, 2, "A1:A3");
            back.Recalculate();
            T.Eq(300.0, b.GetValue(3, 2).Num, "re-entered formula spills");
        });
        t.Test("Older-Excel formulas: implicit intersection inside functions, array functions still see arrays", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            string[] codes = ["a", "bb", "ccc", "dddd"];
            for (int i = 0; i < 4; i++) { s.SetValue(i, 0, CellValue.Text(codes[i])); s.SetValue(i, 3, CellValue.Text(codes[i])); s.SetValue(i, 4, CellValue.Number((i + 1) * 100)); }
            void Legacy(int r, int c, string f) { s.SetFormula(r, c, f); s.GetCell(r, c)!.Implicit = true; }
            Legacy(2, 1, "LEN(A1:A4)");                       // row 3 → LEN(A3)
            Legacy(2, 2, "VLOOKUP(A1:A4,D1:E4,2,FALSE)");     // row 3 → 300
            Legacy(2, 5, "SUMPRODUCT(LEN(A1:A4))");           // array context → 10
            Legacy(2, 6, "A1:A4&\"!\"");                       // operator → A3&"!"
            Legacy(2, 7, "SUMPRODUCT(--(LEN(A:A)>1))");       // 3
            w.Recalculate();
            T.Eq(3.0, s.GetValue(2, 1).Num, "LEN implicit");
            T.Eq(300.0, s.GetValue(2, 2).Num, "VLOOKUP implicit");
            T.Eq(10.0, s.GetValue(2, 5).Num, "SUMPRODUCT array");
            T.Eq("ccc!", s.GetValue(2, 6).Str, "operator implicit");
            T.Eq(3.0, s.GetValue(2, 7).Num, "SUMPRODUCT whole column");
            T.True(s.GetCell(3, 1) == null, "no spill");
            // copying keeps the behaviour
            var clip = SheetOps.Copy(s, CellRange.Single(2, 1), false);
            SheetOps.Paste(s, clip, CellRange.Single(3, 1));
            w.Recalculate();
            T.Eq("LEN(A2:A5)", s.GetCell(3, 1)!.Formula, "pasted formula shifted");
            T.Eq(4.0, s.GetValue(3, 1).Num, "pasted legacy formula → LEN(A4)");
        });
        t.Test("Review fixes: whole-column spill, single-result FILTER in files, recursion, errors, name scope, sort", () =>
        {
            var w = Workbook.CreateBlank(1);
            var s = w.Sheets[0];
            for (int i = 0; i < 5; i++) s.SetValue(i, 0, CellValue.Number(i + 1));
            s.SetFormula(1, 1, "A:A*2");
            for (int k = 0; k < 3; k++) w.Recalculate();
            T.Eq(5, s.GetCell(1, 1)!.SpillRows, "whole-column spill does not grow");
            s.SetFormula(0, 3, "FILTER(A:A,A:A>3)");
            w.Recalculate();
            T.Eq(2, s.GetCell(0, 3)!.SpillRows, "FILTER over whole column");
            s.SetFormula(0, 4, "SUM(B:B)");
            w.Recalculate();
            T.Eq(30.0, s.GetValue(0, 4).Num, "SUM over spilled column");
            // FILTER with one match keeps its dynamic-array flag through save/load
            var w2 = Workbook.CreateBlank(1);
            var s2 = w2.Sheets[0];
            s2.SetValue(0, 0, CellValue.Text("x")); s2.SetValue(1, 0, CellValue.Text("y")); s2.SetValue(2, 0, CellValue.Text("z"));
            for (int i = 0; i < 3; i++) s2.SetValue(i, 1, CellValue.Number(i + 1));
            s2.SetFormula(0, 3, "FILTER(B1:B3,A1:A3=\"x\")");
            w2.Recalculate();
            using var ms = new MemoryStream();
            XlsxWriter.Write(w2, ms);
            ms.Position = 0;
            var back = XlsxReader.Read(ms);
            var b = back.Sheets[0];
            T.True(!b.GetCell(0, 3)!.Implicit, "still a dynamic-array formula");
            b.SetValue(2, 0, CellValue.Text("x"));
            back.Recalculate();
            T.Eq(3.0, b.GetValue(1, 3).Num, "spills after a second match");
            // runaway recursion: error, no crash, no hang
            w.Names["RR"] = "LAMBDA(n,IF(ISERROR(RR(n+1)),RR(n+1),0))";
            w.Names["SELF"] = "SELF(1)";
            var sw = System.Diagnostics.Stopwatch.StartNew();
            T.True(w.Calc.EvaluateFormula(s, 20, 20, "RR(1)").IsError, "runaway recursion → error");
            T.True(w.Calc.EvaluateFormula(s, 20, 20, "SELF+1").IsError, "name defined through itself → error");
            T.True(sw.ElapsedMilliseconds < 10000, "recursion stops quickly: " + sw.ElapsedMilliseconds);
            w.Names["FACT2"] = "LAMBDA(n,IF(n<=1,1,n*FACT2(n-1)))";
            T.Eq(Enumerable.Range(1, 150).Aggregate(1.0, (x, y) => x * y), w.Calc.EvaluateFormula(s, 20, 20, "FACT2(150)").Num, "deep recursion still works");
            // errors in arguments reach ISERROR & co.
            T.True(w.Calc.EvaluateFormula(s, 20, 20, "ISERROR(NoSuchName)").BoolValue, "ISERROR(NoSuchName)");
            T.True(w.Calc.EvaluateFormula(s, 20, 20, "ISNA(NoSuchName)").IsBool, "ISNA(NoSuchName) is a logical");
            T.Eq(5.0, w.Calc.EvaluateFormula(s, 20, 20, "ERROR.TYPE(FOOBAR())").Num, "ERROR.TYPE(#NAME?)");
            // a named LAMBDA does not see the caller's LET variables
            w.Names["XX"] = "1";
            w.Names["ADDX"] = "LAMBDA(y,y+XX)";
            T.Eq(2.0, w.Calc.EvaluateFormula(s, 20, 20, "LET(XX,100,ADDX(1))").Num, "lexical scope of named lambda");
            // sorting through a spill range is refused (like Excel)
            T.True(SheetOps.TouchesSpill(s, new CellRange(0, 0, 9, 1)), "spill detected");
            bool refused = false;
            try { SheetOps.Sort(s, new CellRange(0, 0, 9, 1), [new SheetOps.SortKey(0, false)], false); } catch (InvalidOperationException) { refused = true; }
            T.True(refused, "sort refused");
            T.Eq("@(A1:A2+1)", FileFormula.FromFile("_xlfn.SINGLE(A1:A2+1)"), "@ keeps parentheses");
        });
    }
}
