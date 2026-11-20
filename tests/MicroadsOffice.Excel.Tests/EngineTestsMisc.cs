using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

/// <summary>
/// Engine tests of the functions in FunctionsMisc.cs. Expected values come from Microsoft's documented examples
/// or were computed independently (numpy/scipy, exact arithmetic) — never from this implementation.
/// </summary>
public static class EngineTestsMisc
{
    private static Workbook NewData()
    {
        var wb = Workbook.CreateBlank(3);
        var s = wb.Sheets[0];
        // AGGREGATE documentation data (A1:B11)
        CellValue[] colA = [CellValue.Error(Errors.Div0), N(72), N(30), CellValue.Error(Errors.Num), N(31), N(96), N(32), N(81), N(33), N(53), N(34)];
        double[] colB = [82, 65, 95, 63, 53, 71, 55, 83, 100, 91, 89];
        for (int i = 0; i < 11; i++) { s.SetValue(i, 0, colA[i]); s.SetValue(i, 1, N(colB[i])); }
        // database (D1:H7) and criteria (J1:O3) from the DAVERAGE/DSUM documentation
        string[] head = ["Tree", "Height", "Age", "Yield", "Profit"];
        object[][] rows =
        [
            ["Apple", 18, 20, 14, 105], ["Pear", 12, 12, 10, 96], ["Cherry", 13, 14, 9, 105],
            ["Apple", 14, 15, 10, 75], ["Pear", 9, 8, 8, 76.8], ["Apple", 8, 9, 6, 45]
        ];
        for (int j = 0; j < 5; j++) s.SetValue(0, 3 + j, CellValue.Text(head[j]));
        for (int i = 0; i < rows.Length; i++)
            for (int j = 0; j < 5; j++)
                s.SetValue(i + 1, 3 + j, rows[i][j] is string t ? CellValue.Text(t) : N(Convert.ToDouble(rows[i][j])));
        string[] critHead = ["Tree", "Height", "Age", "Yield", "Profit", "Height"];
        for (int j = 0; j < 6; j++) s.SetValue(0, 9 + j, CellValue.Text(critHead[j]));
        s.SetValue(1, 9, CellValue.Text("=Apple")); s.SetValue(1, 10, CellValue.Text(">10")); s.SetValue(1, 14, CellValue.Text("<16"));
        s.SetValue(2, 9, CellValue.Text("=Pear"));
        // computed criterion (Q1:Q2), "begins with" text criterion (S1:S2), numeric criterion (U1:U2)
        s.SetValue(0, 16, CellValue.Text("Old")); s.SetFormula(1, 16, "F2>12");
        s.SetValue(0, 18, CellValue.Text("Tree")); s.SetValue(1, 18, CellValue.Text("Ch"));
        s.SetValue(0, 20, CellValue.Text("Age")); s.SetValue(1, 20, N(15));
        s.SetValue(0, 24, CellValue.Text("Tree")); s.SetValue(1, 24, CellValue.Text("=Kiwi"));
        // VARA/VARPA data with a logical and a text (W1:W3); PHONETIC text (C1)
        s.SetValue(0, 22, N(1)); s.SetValue(1, 22, CellValue.Bool(true)); s.SetValue(2, 22, CellValue.Text("x"));
        s.SetValue(0, 2, CellValue.Text("東京"));

        // hidden / filtered rows and nested SUBTOTAL / AGGREGATE for AGGREGATE options
        var h = wb.Sheets[1];
        h.Name = "Hid";
        for (int i = 0; i < 5; i++) h.SetValue(i, 0, N(i + 1));
        h.HiddenRows.Add(1);    // value 2 hidden manually
        h.FilteredRows.Add(3);  // value 4 filtered out
        h.SetFormula(5, 0, "SUBTOTAL(9,A1:A2)");   // = 3
        h.SetFormula(6, 0, "AGGREGATE(9,4,A1:A5)"); // = 15

        // multiple regression data (Reg!A1:E11): y, x1..x4
        var r = wb.Sheets[2];
        r.Name = "Reg";
        double[] y = [142000, 144000, 151000, 150000, 139000, 169000, 126000, 142900, 163000, 169000, 149000];
        double[] x1 = [2310, 2333, 2356, 2379, 2402, 2425, 2448, 2471, 2494, 2517, 2540];
        double[] x2 = [2, 2, 3, 3, 2, 4, 2, 2, 3, 4, 2];
        double[] x3 = [2, 2, 1.5, 2, 3, 2, 1.5, 2, 3, 4, 3];
        double[] x4 = [20, 12, 33, 43, 53, 23, 99, 34, 23, 55, 22];
        for (int i = 0; i < 11; i++)
        {
            r.SetValue(i, 0, N(y[i])); r.SetValue(i, 1, N(x1[i])); r.SetValue(i, 2, N(x2[i]));
            r.SetValue(i, 3, N(x3[i])); r.SetValue(i, 4, N(x4[i]));
        }
        wb.Calc.FixedNow = new DateTime(2026, 11, 16, 10, 30, 0);
        wb.Recalculate();
        return wb;
    }

    private static CellValue N(double d) => CellValue.Number(d);

    private static CellValue Ev(Workbook wb, string f) => wb.Calc.EvaluateFormula(wb.Sheets[0], 49, 25, f);

    public static void Run(TestRunner t)
    {
        NumFmt.UseLocale(false);
        var wb = NewData();
        int asserts = 0;
        void Num(string f, double expected, double tol = 1e-9)
        {
            asserts++;
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > tol * Math.Max(1, Math.Abs(expected)))
                throw new AssertException($"{f} => {v.Kind}:{v} expected {expected}");
        }
        void Str(string f, string expected)
        {
            asserts++;
            var v = Ev(wb, f);
            if (!v.IsText || v.Str != expected) throw new AssertException($"{f} => {v.Kind}:'{v}' expected '{expected}'");
        }
        void Bool(string f, bool expected) { asserts++; var v = Ev(wb, f); if (!v.IsBool || v.BoolValue != expected) throw new AssertException($"{f} => {v} expected {expected}"); }
        void E(string f, string code) { asserts++; var v = Ev(wb, f); if (!v.IsError || v.Str != code) throw new AssertException($"{f} => {v.Kind}:{v} expected {code}"); }

        t.Test("Misc: trig & hyperbolic", () =>
        {
            Num("ACOSH(1)", 0); Num("ACOSH(10)", 2.993222846126381); E("ACOSH(0.5)", Errors.Num);
            Num("ASINH(-2.5)", -1.6472311463710958); Num("ATANH(-0.1)", -0.1003353477310756); E("ATANH(1)", Errors.Num);
            Num("ACOT(2)", 0.46364760900080615); Num("ACOT(-2)", 2.677945044588987); Num("ACOT(0)", Math.PI / 2);
            Num("ACOTH(6)", 0.16823611831060648); E("ACOTH(1)", Errors.Num);
            Num("COT(30)", -0.15611995216165922); E("COT(0)", Errors.Div0); E("COT(134217728)", Errors.Num);
            Num("COTH(2)", 1.0373147207275482); E("COTH(0)", Errors.Div0);
            Num("CSC(15)", 1.5377805615408537); E("CSC(0)", Errors.Div0); Num("CSCH(1.5)", 0.46964244059522464); E("CSCH(0)", Errors.Div0);
            Num("SEC(45)", 1.9035944074044246); Num("SECH(45)", 5.725037161098787e-20); E("SEC(134217728)", Errors.Num); E("SEC(\"x\")", Errors.Value);
        });
        t.Test("Misc: number systems & combinatorics", () =>
        {
            Num("ARABIC(\"LVII\")", 57); Num("ARABIC(\"mcmxii\")", 1912); Num("ARABIC(\"\")", 0); Num("ARABIC(\"-MMXI\")", -2011); E("ARABIC(\"ABC\")", Errors.Value);
            Str("BASE(7,2)", "111"); Str("BASE(100,16)", "64"); Str("BASE(15,2,10)", "0000001111"); Str("BASE(35,36)", "Z");
            E("BASE(-1,2)", Errors.Num); E("BASE(7,1)", Errors.Num); E("BASE(7,37)", Errors.Num);
            Num("DECIMAL(\"FF\",16)", 255); Num("DECIMAL(111,2)", 7); Num("DECIMAL(\"zap\",36)", 45745); E("DECIMAL(\"G\",16)", Errors.Num); E("DECIMAL(\"1\",1)", Errors.Num);
            Num("COMBINA(4,3)", 20); Num("COMBINA(10,3)", 220); Num("COMBINA(0,0)", 1); E("COMBINA(-1,2)", Errors.Num);
            Num("FACTDOUBLE(6)", 48); Num("FACTDOUBLE(7)", 105); Num("FACTDOUBLE(0)", 1); E("FACTDOUBLE(-1)", Errors.Num);
            Num("MULTINOMIAL(2,3,4)", 1260); Num("MULTINOMIAL(5)", 1); E("MULTINOMIAL(-1,2)", Errors.Num);
            Num("SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666666666666664,-0.001388888888888889})", 0.7071032148228457); E("SERIESSUM(1,0,1,{1,\"a\"})", Errors.Value);
            Num("SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", -55); Num("SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 521);
            Num("SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 79); E("SUMXMY2({1,2},{1,2,3})", Errors.NA);
            Num("PERCENTOF({1,2},{1,2,3,4})", 0.3); Num("PERCENTOF(B1:B2,B1:B11)", 147.0 / 847); E("PERCENTOF(1,0)", Errors.Div0);
        });
        t.Test("Misc: matrices & arrays", () =>
        {
            Num("MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})", 88); Num("MDETERM({3,6;1,1})", -3); Num("MDETERM({1,2;2,4})", 0);
            E("MDETERM({1,2,3;4,5,6})", Errors.Value); E("MDETERM({1,\"a\";2,3})", Errors.Value);
            Num("INDEX(MINVERSE({4,-1;2,0}),1,2)", 0.5); Num("INDEX(MINVERSE({4,-1;2,0}),2,1)", -1); Num("INDEX(MINVERSE({4,-1;2,0}),2,2)", 2);
            Num("INDEX(MINVERSE({1,2,1;3,4,-1;0,2,0}),1,3)", -0.75); Num("INDEX(MINVERSE({1,2,1;3,4,-1;0,2,0}),3,1)", 0.75); E("MINVERSE({1,2;2,4})", Errors.Num);
            Num("INDEX(MMULT({1,3;7,2},{2,0;0,2}),2,1)", 14); Num("SUM(MMULT({1,3;7,2},{2,0;0,2}))", 26); Num("ROWS(MMULT({1,2,3},{1;2;3}))", 1);
            Num("MMULT({1,2,3},{1;2;3})", 14); E("MMULT({1,2},{1,2})", Errors.Value);
            Num("SUM(MUNIT(3))", 3); Num("INDEX(MUNIT(3),2,2)", 1); Num("INDEX(MUNIT(3),1,2)", 0); Num("ROWS(MUNIT(4))", 4); E("MUNIT(0)", Errors.Value);
            Num("ROWS(RANDARRAY(2,3))", 2); Num("COLUMNS(RANDARRAY(2,3))", 3); Num("SUM(RANDARRAY(2,2,3,3,TRUE))", 12);
            asserts++;
            var rv = Ev(wb, "MAX(RANDARRAY(50,1,10,20))");
            var rmin = Ev(wb, "MIN(RANDARRAY(50,1,10,20))");
            if (!(rv.Num < 20 && rmin.Num >= 10)) throw new AssertException("RANDARRAY range");
            E("RANDARRAY(2,2,5,1)", Errors.Value); E("RANDARRAY(1,1,1.5,3,TRUE)", Errors.Value); E("RANDARRAY(-1)", Errors.Value);
        });
        t.Test("Misc: AGGREGATE", () =>
        {
            // Microsoft documentation example
            Num("AGGREGATE(4,6,A1:A11)", 96); Num("AGGREGATE(14,6,A1:A11,3)", 72); E("AGGREGATE(15,6,A1:A11)", Errors.Value);
            Num("AGGREGATE(12,6,A1:A11,B1:B11)", 68); E("AGGREGATE(9,4,A1:A11)", Errors.Div0); Num("AGGREGATE(2,0,A1:A11)", 9);
            Num("AGGREGATE(3,6,A1:A11)", 9); Num("AGGREGATE(3,4,A1:A11)", 11);
            // hidden (row 2), filtered (row 4), nested SUBTOTAL (A6 = 3) and AGGREGATE (A7 = 15)
            Num("AGGREGATE(9,4,Hid!A1:A6)", 18); Num("AGGREGATE(9,0,Hid!A1:A6)", 15); Num("AGGREGATE(9,5,Hid!A1:A6)", 12);
            Num("AGGREGATE(9,1,Hid!A1:A6)", 9); Num("AGGREGATE(9,4,Hid!A1:A7)", 33); Num("AGGREGATE(9,3,Hid!A1:A7)", 9);
            Num("AGGREGATE(2,1,Hid!A1:A7)", 3); Num("AGGREGATE(1,1,Hid!A1:A5)", 3); Num("AGGREGATE(6,1,Hid!A1:A5)", 15);
            Num("AGGREGATE(15,5,Hid!A1:A6,4)", 5); E("AGGREGATE(15,1,Hid!A1:A6,4)", Errors.Num); Num("AGGREGATE(14,1,Hid!A1:A7,1)", 5);
            Num("AGGREGATE(16,4,{1,2,3,4,5},0.25)", 2); Num("AGGREGATE(17,4,{1,2,3,4,5},3)", 4); Num("AGGREGATE(13,4,Hid!A1:A6)", 3);
            Num("AGGREGATE(19,4,{6,7,15,36,39,40,41,42,43,47,49},1)", 15); Num("AGGREGATE(18,6,{1,2,3,4,#N/A},0.5)", 2.5);
            Num("AGGREGATE(7,4,Hid!A1:A5)", Math.Sqrt(2.5)); Num("AGGREGATE(11,4,Hid!A1:A5)", 2);
            E("AGGREGATE(9,6,{1,2})", Errors.Value); E("AGGREGATE(20,0,A1:A3)", Errors.Value); E("AGGREGATE(9,8,A1:A3)", Errors.Value);
        });
        t.Test("Misc: descriptive statistics", () =>
        {
            Num("COVAR({3,2,4,5,6},{9,7,12,15,17})", 5.2); Num("COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})", 5.2);
            Num("COVARIANCE.S({2,4,8},{5,11,12})", 9.666666666666668); E("COVARIANCE.S({1},{2})", Errors.Div0); E("COVAR({1,2},{1,2,3})", Errors.NA);
            Num("DEVSQ(4,5,8,7,11,4,3)", 48); E("DEVSQ(C1)", Errors.Num);
            Num("KURT(3,4,5,2,3,4,5,6,4,7)", -0.15179963720841405); Num("KURT({1,2,3,4,5,6,7,8,9,30})", 7.503962839309041); E("KURT(1,2,3)", Errors.Div0);
            Num("SKEW(3,4,5,2,3,4,5,6,4,7)", 0.3595430714067974); Num("SKEW.P(3,4,5,2,3,4,5,6,4,7)", 0.30319333935414394); E("SKEW(1,2)", Errors.Div0); E("SKEW.P(2,2,2)", Errors.Div0);
            Num("RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 0.057950191570881195); Num("STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 3.305718950210041);
            E("STEYX({1,2},{1,2})", Errors.Div0); E("RSQ({1,2},{1,2,3})", Errors.NA);
            Num("TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)", 3.7777777777777777); E("TRIMMEAN({1,2},1)", Errors.Num);
            Num("VARA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)", 754.2666666666667); Num("VARPA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)", 678.84);
            Num("STDEVPA(1345,1301,1368,1322,1310,1370,1318,1350,1303,1299)", 26.054558142482477);
            Num("VARA(W1:W3)", 1.0 / 3); Num("VARPA(W1:W3)", 2.0 / 9); Num("STDEVPA(W1:W3)", Math.Sqrt(2.0 / 9)); E("VARA(1)", Errors.Div0);
            Num("INDEX(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),1)", 1); Num("INDEX(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),3)", 4);
            Num("INDEX(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),4)", 2); Num("ROWS(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}))", 4);
            Num("INDEX(FREQUENCY({1,2,3,4,5},{4,2}),1)", 2); Num("INDEX(FREQUENCY({1,2,3,4,5},{4,2}),2)", 2); Num("FREQUENCY({1,2,3},A20)", 3);
            Num("ROWS(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}))", 3); Num("INDEX(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}),2)", 2); Num("INDEX(MODE.MULT({5,7,7,5,9}),1)", 5);
            E("MODE.MULT({1,2,3})", Errors.NA);
        });
        t.Test("Misc: percent ranks & quartiles", () =>
        {
            const string d = "{13,12,11,8,4,3,2,1,1,1}";
            Num($"PERCENTRANK.INC({d},2)", 0.333); Num($"PERCENTRANK.INC({d},4)", 0.555); Num($"PERCENTRANK.INC({d},8)", 0.666);
            Num($"PERCENTRANK.INC({d},5)", 0.583); Num($"PERCENTRANK({d},4)", 0.555); Num($"PERCENTRANK.INC({d},13)", 1); Num($"PERCENTRANK.INC({d},5,5)", 0.58333);
            E($"PERCENTRANK.INC({d},14)", Errors.NA); E($"PERCENTRANK.INC({d},5,0)", Errors.Num);
            const string e = "{1,2,3,6,6,6,7,8,9}";
            Num($"PERCENTRANK.EXC({e},7)", 0.7); Num($"PERCENTRANK.EXC({e},5.43)", 0.381); Num($"PERCENTRANK.EXC({e},5.43,1)", 0.3); E($"PERCENTRANK.EXC({e},0.5)", Errors.NA);
            Num("QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)", 15); Num("QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)", 43);
            E("QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},0)", Errors.Num); E("QUARTILE.EXC({1,2},1)", Errors.Num);
        });
        t.Test("Misc: LINEST, LOGEST, TREND, GROWTH", () =>
        {
            Num("INDEX(LINEST({1,9,5,7},{0,4,2,3}),1)", 2); Num("INDEX(LINEST({1,9,5,7},{0,4,2,3}),2)", 1); Num("INDEX(LINEST({3,5,7}),1)", 2);
            Num("COLUMNS(LINEST({1,9,5,7},{0,4,2,3}))", 2); Num("ROWS(LINEST({1,9,5,7},{0,4,2,3},TRUE,TRUE))", 5);
            // multiple regression with statistics (numpy reference values)
            const string L = "LINEST(Reg!A1:A11,Reg!B1:E11,TRUE,TRUE)";
            Num($"ROWS({L})", 5); Num($"COLUMNS({L})", 5);
            Num($"INDEX({L},1,1)", -234.2371644712075); Num($"INDEX({L},1,2)", 2553.210660391566); Num($"INDEX({L},1,3)", 12529.768167086753);
            Num($"INDEX({L},1,4)", 27.64138736602013); Num($"INDEX({L},1,5)", 52317.830507291816);
            Num($"INDEX({L},2,1)", 13.268011475500272, 1e-8); Num($"INDEX({L},2,4)", 5.429374041545011, 1e-8); Num($"INDEX({L},2,5)", 12237.361602861636, 1e-8);
            Num($"INDEX({L},3,1)", 0.9967479933845103); Num($"INDEX({L},3,2)", 970.5784629285005, 1e-8); E($"INDEX({L},3,3)", Errors.NA);
            Num($"INDEX({L},4,1)", 459.75367422540273, 1e-8); Num($"INDEX({L},4,2)", 6); Num($"INDEX({L},5,1)", 1732393319.2292686, 1e-8);
            Num($"INDEX({L},5,2)", 5652135.316203903, 1e-7); E($"INDEX({L},5,5)", Errors.NA);
            const string L0 = "LINEST(Reg!A1:A11,Reg!B1:E11,FALSE,TRUE)";
            Num($"INDEX({L0},1,1)", -250.13503238418343, 1e-8); Num($"INDEX({L0},1,4)", 50.71193750640334, 1e-8); Num($"INDEX({L0},1,5)", 0);
            E($"INDEX({L0},2,5)", Errors.NA); Num($"INDEX({L0},2,4)", 1.114377349765304, 1e-8); Num($"INDEX({L0},3,1)", 0.9999076735088338);
            Num($"INDEX({L0},4,1)", 18952.72317334794, 1e-7); Num($"INDEX({L0},4,2)", 7);
            Num("INDEX(LINEST({1,9,5,7},{0,4,2,3},FALSE,TRUE),1,1)", 2.310344827586207); Num("INDEX(LINEST({1,9,5,7},{0,4,2,3},FALSE,TRUE),2,1)", 0.11778104, 1e-7);
            Num("INDEX(LINEST({1,9,5,7},{0,4,2,3},FALSE,TRUE),5,1)", 154.79310344827587); Num("INDEX(LINEST({1,9,5,7},{0,4,2,3},FALSE,TRUE),4,2)", 3);
            // collinear x columns: one is dropped (coefficient 0) and df grows
            const string C = "LINEST({1;2;3;5},{1,2;2,4;3,6;4,8},TRUE,TRUE)";
            Num($"INDEX({C},1,2)+2*INDEX({C},1,1)", 1.3); Num($"INDEX({C},4,2)", 2); Num($"INDEX({C},1,3)", -0.5);
            E("LINEST({1,2,3},{1,2})", Errors.Ref); E("LINEST({1,\"a\",3})", Errors.Value);
            // LOGEST / GROWTH (documentation data)
            const string G = "LOGEST({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},TRUE,TRUE)";
            Num($"INDEX({G},1,1)", 1.4632756281161754); Num($"INDEX({G},1,2)", 495.30477015872697, 1e-8); Num($"INDEX({G},2,1)", 0.0026334028914250527, 1e-7);
            Num($"INDEX({G},3,1)", 0.9998086197758176); Num($"INDEX({G},4,2)", 4); E("LOGEST({1,-2,3})", Errors.Num);
            Num("INDEX(GROWTH({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},{17;18}),1)", 320196.718363471, 1e-8);
            Num("INDEX(GROWTH({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},{17;18}),2)", 468536.05418404634, 1e-8);
            Num("INDEX(GROWTH({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16}),1)", 32618.20377353963, 1e-8);
            // TREND (documentation data) and multi-variable TREND
            const string Y = "{133890,135000,135790,137300,138130,139100,139900,141120,141890,143230,144000,145290}";
            Num($"INDEX(TREND({Y},{{1,2,3,4,5,6,7,8,9,10,11,12}},{{13,14,15,16,17}}),1)", 146171.51515151505);
            Num($"INDEX(TREND({Y},{{1,2,3,4,5,6,7,8,9,10,11,12}},{{13,14,15,16,17}}),5)", 150244.2424242423);
            Num($"COLUMNS(TREND({Y},{{1,2,3,4,5,6,7,8,9,10,11,12}},{{13,14,15,16,17}}))", 5); Num($"INDEX(TREND({Y}),1)", 133953.33333333334);
            Num("TREND(Reg!A1:A11,Reg!B1:E11,{2600,3,2,25})", 161025.23436920735, 1e-8); E("TREND(Reg!A1:A11,Reg!B1:E11,{2600,3})", Errors.Ref);
        });
        t.Test("Misc: text", () =>
        {
            Str("ARRAYTOTEXT({TRUE,#VALUE!;1234.01234,\"Seattle\";\"Hello\",\"1,123\"})", "TRUE, #VALUE!, 1234.01234, Seattle, Hello, 1,123");
            Str("ARRAYTOTEXT({TRUE,#VALUE!;1234.01234,\"Seattle\";\"Hello\",\"1,123\"},1)", "{TRUE,#VALUE!;1234.01234,\"Seattle\";\"Hello\",\"1,123\"}");
            E("ARRAYTOTEXT({1,2},2)", Errors.Value);
            Str("VALUETOTEXT(\"Hello\")", "Hello"); Str("VALUETOTEXT(\"Hello\",1)", "\"Hello\""); Str("VALUETOTEXT(1234.01234)", "1234.01234");
            Str("VALUETOTEXT(TRUE,1)", "TRUE"); Str("VALUETOTEXT(#VALUE!)", "#VALUE!");
            const string q = "\"Do. Or do not. There is no try. -Anonymous\"";
            Str($"INDEX(TEXTSPLIT({q},\".\"),2)", " Or do not"); Num($"COLUMNS(TEXTSPLIT({q},\".\"))", 4);
            Str($"INDEX(TEXTSPLIT({q},{{\".\",\"-\"}}),4)", " "); Str($"INDEX(TEXTSPLIT({q},{{\".\",\"-\"}}),5)", "Anonymous");
            Str("INDEX(TEXTSPLIT(\"1,2,3;4,5,6\",\",\",\";\"),2,3)", "6");
            Num("ROWS(TEXTSPLIT(\"1,2,3;4,5,6\",\",\",\";\"))", 2); Num("ROWS(TEXTSPLIT(\"a;b;c\",,\";\"))", 3);
            Num($"ROWS(TEXTSPLIT({q},\" \",\".\",TRUE))", 4); Num($"COLUMNS(TEXTSPLIT({q},\" \",\".\",TRUE))", 4);
            E($"INDEX(TEXTSPLIT({q},\" \",\".\",TRUE),1,2)", Errors.NA); Str($"INDEX(TEXTSPLIT({q},\" \",\".\",TRUE,,\"-\"),1,2)", "-");
            Str($"INDEX(TEXTSPLIT({q},\" \",\".\",TRUE),3,4)", "try");
            Num("COLUMNS(TEXTSPLIT(\"a,,b\",\",\"))", 3); Num("COLUMNS(TEXTSPLIT(\"a,,b\",\",\",,TRUE))", 2);
            Num("COLUMNS(TEXTSPLIT(\"aXbxc\",\"x\"))", 2); Num("COLUMNS(TEXTSPLIT(\"aXbxc\",\"x\",,,1))", 3); E("TEXTSPLIT(\"abc\",\"\")", Errors.Value);
            Bool("REGEXTEST(\"alpha\",\"^a.*a$\")", true); Bool("REGEXTEST(\"ABC\",\"abc\")", false); Bool("REGEXTEST(\"ABC\",\"abc\",1)", true);
            Bool("REGEXTEST(\"(555) 123-4567\",\"^\\(\\d{3}\\) \\d{3}-\\d{4}$\")", true); E("REGEXTEST(\"a\",\"(\")", Errors.Value); E("REGEXTEST(\"a\",\"a\",2)", Errors.Value);
            Num("SUMPRODUCT(--REGEXTEST(D2:D7,\"^A\"))", 3);
            Str("REGEXEXTRACT(\"DylanWilliams\",\"[A-Z][a-z]+\")", "Dylan"); Str("INDEX(REGEXEXTRACT(\"DylanWilliams\",\"[A-Z][a-z]+\",1),2)", "Williams");
            Num("ROWS(REGEXEXTRACT(\"DylanWilliams\",\"[A-Z][a-z]+\",1))", 2); Str("INDEX(REGEXEXTRACT(\"2024-03-15\",\"(\\d+)-(\\d+)-(\\d+)\",2),1,3)", "15");
            Num("COLUMNS(REGEXEXTRACT(\"2024-03-15\",\"(\\d+)-(\\d+)-(\\d+)\",2))", 3); E("REGEXEXTRACT(\"abc\",\"\\d\")", Errors.NA);
            Str("REGEXEXTRACT(\"abc\",\"B\",0,1)", "b"); E("REGEXEXTRACT(\"abc\",\"b\",3)", Errors.Value);
            Str("REGEXREPLACE(\"Sonia Rees\",\"(\\w+) (\\w+)\",\"$2, $1\")", "Rees, Sonia"); Str("REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\")", "a#b#c#");
            Str("REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\",2)", "a1b#c3"); Str("REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\",-1)", "a1b2c#"); Str("REGEXREPLACE(\"a1b2c3\",\"\\d\",\"#\",5)", "a1b2c3");
            Str("REGEXREPLACE(\"ABC\",\"b\",\"x\",0,1)", "AxC"); Str("REGEXREPLACE(\"ABC\",\"b\",\"x\")", "ABC");
            Str("ASC(\"ＡＢＣ　１２３\")", "ABC 123"); Str("ASC(\"ガパア\")", "ｶﾞﾊﾟｱ"); Str("DBCS(\"ABC 123\")", "ＡＢＣ　１２３"); Str("JIS(\"ｶﾞｷﾞﾊﾟｱ\")", "ガギパア"); Str("ASC(\"Việt\")", "Việt");
            Str("BAHTTEXT(1234)", "หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน"); Str("BAHTTEXT(0)", "ศูนย์บาทถ้วน"); Str("BAHTTEXT(21.25)", "ยี่สิบเอ็ดบาทยี่สิบห้าสตางค์");
            Str("BAHTTEXT(11000000)", "สิบเอ็ดล้านบาทถ้วน"); Str("BAHTTEXT(-1.5)", "ลบหนึ่งบาทห้าสิบสตางค์"); Str("BAHTTEXT(0.5)", "ห้าสิบสตางค์"); E("BAHTTEXT(\"x\")", Errors.Value);
            Num("FINDB(\"c\",\"abcabc\",4)", 6); Num("SEARCHB(\"C\",\"abc\")", 3); Str("LEFTB(\"Microads\",5)", "Micro"); Str("RIGHTB(\"Microads\",3)", "ads");
            Str("MIDB(\"Microads\",3,3)", "cro"); Num("LENB(\"abc\")", 3); Str("REPLACEB(\"abcdef\",2,3,\"X\")", "aXef"); E("FINDB(\"C\",\"abc\")", Errors.Value);
            Str("PHONETIC(C1)", "東京"); Str("PHONETIC(C1:C3)", "東京");
        });
        t.Test("Misc: NETWORKDAYS.INTL, WORKDAY.INTL, INFO", () =>
        {
            Num("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))", 22); Num("NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))", -21);
            Num("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{\"2006/1/2\",\"2006/1/16\"})", 22);
            Num("NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"0010001\",{\"2006/1/2\",\"2006/1/16\"})", 20);
            Num("NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),1,{45359,45360})", 20); Num("NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),2)", 22);
            Num("NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),\"1111111\")", 0); E("NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),8)", Errors.Num);
            E("NETWORKDAYS.INTL(DATE(2024,3,1),DATE(2024,3,31),\"11\")", Errors.Value);
            Num("WORKDAY.INTL(DATE(2012,1,1),90,11)", 41013); Num("WORKDAY.INTL(DATE(2012,1,1),30,17)", 40944); E("WORKDAY.INTL(DATE(2012,1,1),30,0)", Errors.Num);
            Num("WORKDAY.INTL(DATE(2024,3,15),-10,1,DATE(2024,3,11))", 45351); Num("WORKDAY.INTL(DATE(2024,3,16),5)", 45373);
            E("WORKDAY.INTL(DATE(2024,3,16),5,\"1111111\")", Errors.Value); Num("WORKDAY.INTL(DATE(2024,3,15),0)", 45366);
            Num("INFO(\"numfile\")", 3); Str("INFO(\"system\")", "pcdos"); Str("INFO(\"release\")", "16.0"); Str("INFO(\"recalc\")", "Automatic");
            Str("INFO(\"origin\")", "$A:$A$1"); E("INFO(\"bogus\")", Errors.Value);
            asserts++;
            var os = Ev(wb, "INFO(\"osversion\")"); if (!os.IsText || !os.Str!.StartsWith("Windows")) throw new AssertException("INFO osversion: " + os);
            asserts++;
            var dir = Ev(wb, "INFO(\"directory\")"); if (!dir.IsText || !dir.Str!.EndsWith(Path.DirectorySeparatorChar)) throw new AssertException("INFO directory: " + dir);
        });
        t.Test("Misc: database functions", () =>
        {
            // Microsoft documentation examples (database D1:H7, criteria J1:O3)
            Num("DCOUNT(D1:H7,\"Age\",J1:O2)", 1); Num("DCOUNTA(D1:H7,\"Profit\",J1:O2)", 1); Num("DMAX(D1:H7,\"Profit\",J1:J3)", 105);
            Num("DMIN(D1:H7,\"Profit\",J1:K2)", 75); Num("DSUM(D1:H7,\"Profit\",J1:J2)", 225); Num("DSUM(D1:H7,\"Profit\",J1:O2)", 75);
            Num("DPRODUCT(D1:H7,\"Yield\",J1:O2)", 10); Num("DAVERAGE(D1:H7,\"Yield\",J1:K2)", 12); Num("DAVERAGE(D1:H7,3,D1:H7)", 13);
            Num("DSTDEV(D1:H7,\"Yield\",J1:J3)", 2.9664793948382653); Num("DSTDEVP(D1:H7,\"Yield\",J1:J3)", 2.6532998322843198);
            Num("DVAR(D1:H7,\"Yield\",J1:J3)", 8.8); Num("DVARP(D1:H7,\"Yield\",J1:J3)", 7.04); E("DGET(D1:H7,\"Yield\",J1:J3)", Errors.Num);
            Num("DGET(D1:H7,\"Yield\",J1:O2)", 10); Str("DGET(D1:H7,\"Tree\",U1:U2)", "Apple"); Num("DGET(D1:H7,\"Yield\",U1:U2)", 10); E("DGET(D1:H7,\"Yield\",Y1:Y2)", Errors.Value);
            Num("DCOUNT(D1:H7,,J1:J3)", 5); Num("DCOUNTA(D1:H7,,J1:J2)", 3); Num("DCOUNTA(D1:H7,1,J1:J3)", 5); Num("DCOUNT(D1:H7,\"Tree\",J1:J3)", 0);
            Num("DSUM(D1:H7,\"profit\",S1:S2)", 105); Num("DCOUNT(D1:H7,\"Age\",Q1:Q2)", 3); Num("DSUM(D1:H7,5,Q1:Q2)", 285);
            E("DSUM(D1:H7,\"Weight\",J1:J2)", Errors.Value); E("DSUM(D1:H7,6,J1:J2)", Errors.Value); E("DAVERAGE(D1:H7,\"Yield\",Y1:Y2)", Errors.Div0); Num("DSUM(D1:H7,\"Yield\",Y1:Y2)", 0); Num("DMAX(D1:H7,\"Yield\",Y1:Y2)", 0);
            E("DSUM(D1:H7,\"Yield\",J1:J1)", Errors.Value); E("DSUM(D1:H7,,J1:J2)", Errors.Value);
        });
        t.Test("Misc: ENCODEURL, FILTERXML", () =>
        {
            Str("ENCODEURL(\"http://contoso.sharepoint.com/teams/Finance/Documents/April Reports/Profit and Loss Statement.xlsx\")",
                "http%3A%2F%2Fcontoso.sharepoint.com%2Fteams%2FFinance%2FDocuments%2FApril%20Reports%2FProfit%20and%20Loss%20Statement.xlsx");
            Str("ENCODEURL(\"Việt Nam\")", "Vi%E1%BB%87t%20Nam");
            const string x = "\"<r><i>1</i><i>2</i><i>x</i><a k='v'/></r>\"";
            Num($"SUM(FILTERXML({x},\"//i\"))", 3); Str($"INDEX(FILTERXML({x},\"//i\"),3)", "x"); Num($"ROWS(FILTERXML({x},\"//i\"))", 3);
            Str($"FILTERXML({x},\"//a/@k\")", "v"); Num($"FILTERXML({x},\"count(//i)\")", 3); Num($"FILTERXML({x},\"//i[2]\")", 2);
            E($"FILTERXML({x},\"//zz\")", Errors.Value); E("FILTERXML(\"<r>\",\"//r\")", Errors.Value); E($"FILTERXML({x},\"//[\")", Errors.Value);
        });
        t.Log($"      misc function assertions executed: {asserts}");
    }
}
