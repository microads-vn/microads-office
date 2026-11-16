using System.Globalization;
using MicroadsOffice.Excel;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.IO;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

public static class EngineTests
{
    private static Workbook NewData()
    {
        var wb = Workbook.CreateBlank(2);
        var s = wb.Sheets[0];
        for (int i = 0; i < 5; i++)
        {
            s.SetValue(i, 0, CellValue.Number(i + 1));
            s.SetValue(i, 1, CellValue.Number((i + 1) * 10));
        }
        string[] fruits = ["apple", "banana", "cherry", "date", "elder"];
        for (int i = 0; i < 5; i++) s.SetValue(i, 2, CellValue.Text(fruits[i]));
        s.SetValue(0, 3, CellValue.Bool(true));
        s.SetValue(0, 4, CellValue.Text("5"));
        s.SetValue(0, 5, CellValue.Number(DateSerial.ToSerial(new DateTime(2024, 3, 15))));
        wb.Sheets[1].Name = "Data 2";
        wb.Sheets[1].SetValue(0, 0, CellValue.Number(100));
        wb.Sheets[1].SetValue(1, 0, CellValue.Number(200));
        wb.Names["Rate"] = "0.1";
        wb.Names["Prices"] = "Sheet1!$B$1:$B$5";
        wb.Calc.FixedNow = new DateTime(2026, 11, 16, 10, 30, 0);
        wb.Recalculate();
        return wb;
    }

    private static CellValue Ev(Workbook wb, string f) => wb.Calc.EvaluateFormula(wb.Sheets[0], 9, 9, f);

    public static void Run(TestRunner t)
    {
        NumFmt.UseLocale(false); // en-US style for deterministic text assertions
        var wb = NewData();
        int asserts = 0;
        void N(string f, double expected, double tol = 1e-9)
        {
            asserts++;
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > tol * Math.Max(1, Math.Abs(expected)))
                throw new AssertException($"{f} => {v.Kind}:{v} expected {expected}");
        }
        void S(string f, string expected)
        {
            asserts++;
            var v = Ev(wb, f);
            if (v.ToString() != expected) throw new AssertException($"{f} => {v.Kind}:'{v}' expected '{expected}'");
        }
        void B(string f, bool expected) { asserts++; var v = Ev(wb, f); if (!v.IsBool || v.BoolValue != expected) throw new AssertException($"{f} => {v} expected {expected}"); }
        void E(string f, string code) { asserts++; var v = Ev(wb, f); if (!v.IsError || v.Str != code) throw new AssertException($"{f} => {v} expected {code}"); }

        t.Test("Formula: operators & precedence", () =>
        {
            N("1+2*3", 7); N("(1+2)*3", 9); N("2^3^2", 64); N("-2^2", 4); N("10/4", 2.5); N("50%", 0.5); N("2*50%", 1);
            S("\"a\"&\"b\"&1", "ab1"); B("1<2", true); B("\"abc\"=\"ABC\"", true); B("2<>2", false); B("\"b\">\"a\"", true); B("1<\"a\"", true);
            N("A1+B1", 11); N("A1:A5*0+1", 1); N("\"5\"+1", 6); N("TRUE+1", 2); N("E1*2", 10); E("1/0", Errors.Div0); E("\"x\"+1", Errors.Value);
            N("'Data 2'!A1+'Data 2'!A2", 300); N("Rate*100", 10); E("NoSuchName+1", Errors.Name); N("SUM({1,2,3;4,5,6})", 21);
        });
        t.Test("Formula: math & trig", () =>
        {
            N("SUM(A1:A5)", 15); N("SUM(A1:B5)", 165); N("SUM(A:A)", 15); N("SUM(A1,B1,5)", 16); N("SUM(Prices)", 150);
            N("SUMIF(A1:A5,\">2\")", 12); N("SUMIF(C1:C5,\"b*\",B1:B5)", 20); N("SUMIFS(B1:B5,A1:A5,\">=2\",A1:A5,\"<=4\")", 90);
            N("SUMPRODUCT(A1:A5,B1:B5)", 550); N("SUMPRODUCT((A1:A5>2)*B1:B5)", 120); N("SUMSQ(1,2,3)", 14); N("PRODUCT(A1:A4)", 24);
            N("ROUND(2.345,2)", 2.35); N("ROUND(-2.5,0)", -3); N("ROUND(1234,-2)", 1200); N("ROUNDUP(2.341,2)", 2.35); N("ROUNDDOWN(2.349,2)", 2.34);
            N("TRUNC(-2.7)", -2); N("INT(-2.3)", -3); N("MOD(10,3)", 1); N("MOD(-3,2)", 1); E("MOD(1,0)", Errors.Div0); N("ABS(-5)", 5);
            N("SQRT(16)", 4); E("SQRT(-1)", Errors.Num); N("POWER(2,10)", 1024); N("EXP(0)", 1); N("LN(EXP(2))", 2); N("LOG(1000)", 3); N("LOG(8,2)", 3);
            N("LOG10(100)", 2); N("PI()", Math.PI); N("CEILING(4.2,1)", 5); N("CEILING(4.2,0.5)", 4.5); N("FLOOR(4.7,1)", 4); N("FLOOR.MATH(-4.5)", -5);
            N("SIGN(-3)", -1); N("FACT(5)", 120); N("COMBIN(5,2)", 10); N("PERMUT(5,2)", 20); N("GCD(12,18)", 6); N("LCM(4,6)", 12);
            N("EVEN(3)", 4); N("ODD(2)", 3); N("MROUND(10,3)", 9); N("QUOTIENT(7,2)", 3); N("SIN(0)", 0); N("COS(0)", 1); N("DEGREES(PI())", 180);
            N("RADIANS(180)", Math.PI); N("ATAN2(1,1)", Math.PI / 4); S("ROMAN(2024)", "MMXXIV"); N("SUBTOTAL(9,A1:A5)", 15); N("SUBTOTAL(1,A1:A5)", 3);
            asserts++; var r = Ev(wb, "RAND()"); if (!(r.IsNumber && r.Num >= 0 && r.Num < 1)) throw new AssertException("RAND");
            asserts++; var rb = Ev(wb, "RANDBETWEEN(1,6)"); if (!(rb.Num >= 1 && rb.Num <= 6 && rb.Num == Math.Floor(rb.Num))) throw new AssertException("RANDBETWEEN");
        });
        t.Test("Formula: statistics", () =>
        {
            N("AVERAGE(A1:A5)", 3); N("AVERAGEIF(A1:A5,\">2\")", 4); N("AVERAGEIFS(B1:B5,A1:A5,\"<3\")", 15); N("COUNT(A1:E5)", 10); N("COUNTA(C1:C5)", 5);
            N("COUNTBLANK(A1:A6)", 1); N("COUNTIF(C1:C5,\"*an*\")", 1); N("COUNTIF(A1:A5,3)", 1); N("COUNTIFS(A1:A5,\">1\",B1:B5,\"<50\")", 3);
            N("MAX(B1:B5)", 50); N("MIN(B1:B5)", 10); N("MAXIFS(B1:B5,A1:A5,\"<4\")", 30); N("MINIFS(B1:B5,A1:A5,\">2\")", 30); N("MEDIAN(1,3,2,4)", 2.5);
            N("MODE(1,2,2,3)", 2); N("STDEV(A1:A5)", Math.Sqrt(2.5)); N("STDEV.P(A1:A5)", Math.Sqrt(2)); N("VAR(A1:A5)", 2.5); N("VAR.P(A1:A5)", 2);
            N("LARGE(B1:B5,2)", 40); N("SMALL(B1:B5,2)", 20); N("RANK(30,B1:B5)", 3); N("RANK(30,B1:B5,1)", 3); N("PERCENTILE(A1:A5,0.25)", 2);
            N("QUARTILE(A1:A5,3)", 4); N("CORREL(A1:A5,B1:B5)", 1); N("AVEDEV(A1:A5)", 1.2); N("GEOMEAN(1,4)", 2); N("SLOPE(B1:B5,A1:A5)", 10);
            N("INTERCEPT(B1:B5,A1:A5)", 0); N("FORECAST(6,B1:B5,A1:A5)", 60); E("AVERAGE(C1:C5)", Errors.Div0);
        });
        t.Test("Formula: logical", () =>
        {
            S("IF(A1>0,\"pos\",\"neg\")", "pos"); B("IF(FALSE,1)", false); N("IF(TRUE,,1)", 0);
            S("IFS(A3=1,\"one\",A3=3,\"three\")", "three"); S("IFERROR(1/0,\"err\")", "err"); N("IFERROR(5,0)", 5); S("IFNA(NA(),\"none\")", "none");
            B("AND(TRUE,A1=1)", true); B("OR(FALSE,A1=2)", false); B("NOT(TRUE)", false); B("XOR(TRUE,TRUE)", false); B("AND(D1:D2)", true);
            S("SWITCH(A2,1,\"a\",2,\"b\",\"z\")", "b"); S("SWITCH(9,1,\"a\",\"z\")", "z"); N("CHOOSE(2,10,20,30)", 20); N("SUM(CHOOSE(1,A1:A5,B1:B5))", 15);
            B("TRUE()", true); B("FALSE()", false);
        });
        t.Test("Formula: lookup & reference", () =>
        {
            N("VLOOKUP(3,A1:B5,2,FALSE)", 30); N("VLOOKUP(3.5,A1:B5,2)", 30); E("VLOOKUP(9,A1:B5,2,FALSE)", Errors.NA); S("VLOOKUP(\"ch*\",C1:C5,1,FALSE)", "cherry");
            N("HLOOKUP(2,{1,2,3;10,20,30},2,FALSE)", 20); N("LOOKUP(4,A1:A5,B1:B5)", 40); N("MATCH(30,B1:B5,0)", 3); N("MATCH(35,B1:B5)", 3); N("MATCH(\"date\",C1:C5,0)", 4);
            N("XLOOKUP(\"date\",C1:C5,B1:B5)", 40); S("XLOOKUP(\"zz\",C1:C5,C1:C5,\"none\")", "none"); N("XLOOKUP(35,B1:B5,A1:A5,,-1)", 3); N("XLOOKUP(35,B1:B5,A1:A5,,1)", 4);
            N("XMATCH(40,B1:B5)", 4); N("INDEX(B1:B5,4)", 40); N("INDEX(A1:B5,2,2)", 20); N("SUM(INDEX(A1:B5,0,2))", 150); N("INDEX({1,2;3,4},2,1)", 3);
            N("SUM(OFFSET(A1,1,0,3,1))", 9); N("INDIRECT(\"B2\")", 20); N("INDIRECT(\"'Data 2'!A2\")", 200); N("ROW(A7)", 7); N("COLUMN(C1)", 3); N("ROW()", 10);
            N("ROWS(A1:B5)", 5); N("COLUMNS(A1:C1)", 3); S("ADDRESS(2,3)", "$C$2"); S("ADDRESS(2,3,4)", "C2"); S("ADDRESS(1,1,1,TRUE,\"Data 2\")", "'Data 2'!$A$1");
            N("SUM(TRANSPOSE(A1:A5))", 15); N("INDEX(A1:B5,MATCH(\"cherry\",C1:C5,0),2)", 30);
        });
        t.Test("Formula: text", () =>
        {
            S("LEFT(\"Microads\",5)", "Micro"); S("RIGHT(\"Microads\",3)", "ads"); S("MID(\"Microads\",3,3)", "cro"); N("LEN(\"Việt Nam\")", 8);
            S("LOWER(\"ABC\")", "abc"); S("UPPER(\"việt\")", "VIỆT"); S("PROPER(\"nguyễn văn an\")", "Nguyễn Văn An"); S("TRIM(\"  a   b  \")", "a b");
            S("CONCAT(A1:A3)", "123"); S("CONCATENATE(\"a\",1,\"b\")", "a1b"); S("TEXTJOIN(\", \",TRUE,C1:C3)", "apple, banana, cherry");
            S("SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c"); S("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c"); S("REPLACE(\"abcdef\",2,3,\"X\")", "aXef");
            N("FIND(\"c\",\"abcabc\")", 3); N("FIND(\"c\",\"abcabc\",4)", 6); E("FIND(\"C\",\"abc\")", Errors.Value); N("SEARCH(\"C\",\"abc\")", 3); N("SEARCH(\"b?d\",\"abcd\")", 2);
            S("TEXT(1234.5,\"#,##0.00\")", "1,234.50"); S("TEXT(0.256,\"0.0%\")", "25.6%"); S("TEXT(F1,\"dd/mm/yyyy\")", "15/03/2024"); S("TEXT(F1,\"yyyy-mm-dd\")", "2024-03-15");
            N("VALUE(\"12.5\")", 12.5); N("VALUE(\"50%\")", 0.5); N("NUMBERVALUE(\"1.234,5\",\",\",\".\")", 1234.5); S("REPT(\"ab\",3)", "ababab");
            B("EXACT(\"a\",\"A\")", false); S("CHAR(65)", "A"); N("CODE(\"A\")", 65); S("CLEAN(\"a\"&CHAR(7)&\"b\")", "ab"); S("FIXED(1234.567,1)", "1,234.6");
            S("T(\"x\")", "x"); S("T(5)", ""); S("TEXTBEFORE(\"a@b.com\",\"@\")", "a"); S("TEXTAFTER(\"a@b.com\",\"@\")", "b.com"); S("UNICHAR(8364)", "€");
        });
        t.Test("Formula: date & time", () =>
        {
            N("DATE(2024,3,15)", DateSerial.ToSerial(new DateTime(2024, 3, 15))); N("DATE(2024,14,1)", DateSerial.ToSerial(new DateTime(2025, 2, 1)));
            N("YEAR(F1)", 2024); N("MONTH(F1)", 3); N("DAY(F1)", 15); N("TIME(12,0,0)", 0.5); N("HOUR(0.75)", 18); N("MINUTE(TIME(1,23,0))", 23);
            N("SECOND(TIME(0,0,45))", 45); N("WEEKDAY(F1)", 6); N("WEEKDAY(F1,2)", 5); N("WEEKNUM(DATE(2024,1,7))", 2); N("EDATE(F1,1)", DateSerial.ToSerial(new DateTime(2024, 4, 15)));
            N("EOMONTH(F1,0)", DateSerial.ToSerial(new DateTime(2024, 3, 31))); N("DATEDIF(DATE(2000,5,10),DATE(2024,3,15),\"Y\")", 23);
            N("DATEDIF(DATE(2024,1,31),DATE(2024,3,15),\"M\")", 1); N("DATEDIF(DATE(2024,1,1),DATE(2024,3,1),\"D\")", 60); N("DAYS(DATE(2024,3,1),DATE(2024,1,1))", 60);
            N("NETWORKDAYS(DATE(2024,3,1),DATE(2024,3,31))", 21); N("WORKDAY(DATE(2024,3,15),1)", DateSerial.ToSerial(new DateTime(2024, 3, 18)));
            N("DATEVALUE(\"2024-03-15\")", DateSerial.ToSerial(new DateTime(2024, 3, 15))); N("TIMEVALUE(\"6:00\")", 0.25); N("YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))", 0.5);
            N("TODAY()", DateSerial.ToSerial(new DateTime(2026, 11, 16))); N("NOW()", DateSerial.ToSerial(new DateTime(2026, 11, 16, 10, 30, 0)), 1e-9);
            N("DAY(60)", 29); N("DATE(1900,3,1)", 61); N("ISOWEEKNUM(DATE(2024,1,1))", 1);
        });
        t.Test("Formula: financial", () =>
        {
            N("PMT(0.1/12,12,1000000)", -87915.887230, 1e-6); N("PMT(0,10,1000)", -100); N("FV(0.05,10,-100)", 1257.789253, 1e-6);
            N("PV(0.05,10,-100)", 772.173493, 1e-6); N("NPV(0.1,100,100,100)", 248.685199, 1e-6); N("IRR({-100,60,60})", 0.130662, 1e-5);
            N("RATE(12,-87915.887230,1000000)*12", 0.1, 1e-6); N("NPER(0.01,-100,1000)", 10.588644, 1e-6); N("IPMT(0.1/12,1,12,1000000)", -8333.333333, 1e-6);
            N("PPMT(0.1/12,1,12,1000000)", -79582.553897, 1e-6); N("SLN(10000,1000,5)", 1800); N("SYD(10000,1000,5,1)", 3000); N("DDB(10000,1000,5,1)", 4000);
        });
        t.Test("Formula: information", () =>
        {
            B("ISBLANK(A9)", true); B("ISNUMBER(A1)", true); B("ISTEXT(C1)", true); B("ISNONTEXT(A1)", true); B("ISLOGICAL(D1)", true); B("ISERROR(1/0)", true);
            B("ISERR(NA())", false); B("ISNA(NA())", true); B("ISEVEN(4)", true); B("ISODD(4)", false); B("ISREF(A1)", true); B("ISREF(5)", false);
            E("NA()", Errors.NA); N("TYPE(\"a\")", 2); N("TYPE(1)", 1); N("TYPE(A1:A3)", 64); N("N(TRUE)", 1); N("N(\"x\")", 0); N("ERROR.TYPE(1/0)", 2);
            N("SHEETS()", 2);
        });
        t.Test("DOCSO: Vietnamese number reading", () =>
        {
            S("DOCSO(120000)", "Một trăm hai mươi nghìn đồng"); S("DOCSO(1005)", "Một nghìn không trăm linh năm đồng"); S("DOCSO(15)", "Mười lăm đồng");
            S("DOCSO(21)", "Hai mươi mốt đồng"); S("DOCSO(1000000)", "Một triệu đồng"); S("DOCSO(0)", "Không đồng"); S("DOCSO(105)", "Một trăm linh năm đồng");
            S("DOCSO(2500000000)", "Hai tỷ năm trăm triệu đồng"); S("DOCSO(35,\"\")", "Ba mươi lăm"); S("DOCSO(-5)", "Âm năm đồng");
        });
        t.Test("Formula: errors & circular references", () =>
        {
            E("#REF!+1", Errors.Ref); E("SUM(1,#N/A)", Errors.NA); E("UNKNOWNFN(1)", Errors.Name); S("IFERROR(#VALUE!,\"ok\")", "ok");
            var w2 = Workbook.CreateBlank();
            var s = w2.Sheets[0];
            s.SetFormula(0, 0, "B1+1"); s.SetFormula(0, 1, "A1+1");
            w2.Recalculate();
            T.True(w2.Calc.Circular.Count > 0, "circular detected");
            s.SetFormula(0, 1, "5"); w2.Recalculate();
            T.Eq(6.0, s.GetValue(0, 0).Num, "circular resolved");
            T.Eq(0, w2.Calc.Circular.Count);
            // long dependency chain (stack safety)
            var w3 = Workbook.CreateBlank();
            var s3 = w3.Sheets[0];
            s3.SetValue(0, 0, CellValue.Number(1));
            for (int i = 1; i < 20000; i++) s3.SetFormula(i, 0, $"A{i}+1");
            s3.SetFormula(0, 1, "A20000");
            w3.Recalculate();
            T.Eq(20000.0, s3.GetValue(0, 1).Num, "chain");
            // reverse order chain (deep recursion)
            var w4 = Workbook.CreateBlank();
            var s4 = w4.Sheets[0];
            s4.SetValue(19999, 0, CellValue.Number(1));
            for (int i = 0; i < 19999; i++) s4.SetFormula(i, 0, $"A{i + 2}+1");
            w4.Recalculate();
            T.Eq(20000.0, s4.GetValue(0, 0).Num, "reverse chain");
        });
        t.Test("Reference adjustment: copy, insert, delete, rename", () =>
        {
            T.Eq("B2+$A$1+C$3", RefAdjust.Shift("A1+$A$1+B$3", 1, 1));
            T.Eq("#REF!+1", RefAdjust.Shift("A1+1", -1, 0));
            T.Eq("SUM(A1:A12)", RefAdjust.InsertDelete("SUM(A1:A10)", "Sheet1", "Sheet1", true, 4, 2));
            T.Eq("SUM(A1:A8)", RefAdjust.InsertDelete("SUM(A1:A10)", "Sheet1", "Sheet1", true, 2, -2));
            T.Eq("#REF!*2", RefAdjust.InsertDelete("A5*2", "Sheet1", "Sheet1", true, 4, -1));
            T.Eq("D1+'Data 2'!B1", RefAdjust.InsertDelete("C1+'Data 2'!B1", "Sheet1", "Sheet1", false, 1, 1));
            T.Eq("C1+'Data 2'!C1", RefAdjust.InsertDelete("C1+'Data 2'!B1", "Sheet1", "Data 2", false, 1, 1));
            T.Eq("Mới!A1", RefAdjust.RenameSheet("Old!A1", "Old", "Mới"));
            var w = NewData();
            var s = w.Sheets[0];
            s.SetFormula(6, 0, "SUM(A1:A5)");
            SheetOps.InsertRows(s, 2, 1);
            T.Eq("SUM(A1:A6)", s.GetCell(7, 0)!.Formula);
            w.Recalculate();
            T.Eq(15.0, s.GetValue(7, 0).Num, "sum after insert");
            SheetOps.DeleteRows(s, 2, 1);
            T.Eq("SUM(A1:A5)", s.GetCell(6, 0)!.Formula);
            SheetOps.InsertCols(s, 0, 1);
            T.Eq("SUM(B1:B5)", s.GetCell(6, 1)!.Formula);
            T.Eq(10.0, s.GetValue(0, 2).Num, "B moved to C");
            SheetOps.RenameSheet(w, w.Sheets[1], "Kho");
            s.SetFormula(8, 1, "Kho!A1*2"); w.Recalculate();
            T.Eq(200.0, s.GetValue(8, 1).Num, "renamed sheet ref");
            SheetOps.DeleteSheet(w, w.Sheets[1]);
            T.True(s.GetCell(8, 1)!.Formula!.Contains("#REF!"), "deleted sheet -> #REF!");
        });
        t.Test("Fill series & copy/paste", () =>
        {
            var w = Workbook.CreateBlank();
            var s = w.Sheets[0];
            s.SetValue(0, 0, CellValue.Number(1)); s.SetValue(1, 0, CellValue.Number(3));
            SheetOps.Fill(s, new CellRange(0, 0, 1, 0), new CellRange(0, 0, 5, 0));
            T.Eq(11.0, s.GetValue(5, 0).Num, "arith series");
            s.SetValue(0, 1, CellValue.Text("Thứ Hai"));
            SheetOps.Fill(s, new CellRange(0, 1, 0, 1), new CellRange(0, 1, 7, 1));
            T.Eq("Thứ Ba", s.GetValue(1, 1).Str); T.Eq("Thứ Hai", s.GetValue(7, 1).Str);
            s.SetValue(0, 2, CellValue.Text("Jan"));
            SheetOps.Fill(s, new CellRange(0, 2, 0, 2), new CellRange(0, 2, 2, 2));
            T.Eq("Mar", s.GetValue(2, 2).Str);
            s.SetValue(0, 3, CellValue.Text("Item 1"));
            SheetOps.Fill(s, new CellRange(0, 3, 0, 3), new CellRange(0, 3, 3, 3));
            T.Eq("Item 4", s.GetValue(3, 3).Str);
            s.SetValue(0, 4, CellValue.Number(DateSerial.ToSerial(new DateTime(2024, 1, 31))));
            s.SetStyle(0, 4, w.Styles.Intern(CellStyle.Default with { NumberFormat = "dd/mm/yyyy" }));
            SheetOps.Fill(s, new CellRange(0, 4, 0, 4), new CellRange(0, 4, 2, 4));
            T.Eq(DateSerial.ToSerial(new DateTime(2024, 2, 2)), s.GetValue(2, 4).Num, "date step");
            s.SetFormula(0, 5, "A1*2");
            SheetOps.Fill(s, new CellRange(0, 5, 0, 5), new CellRange(0, 5, 3, 5));
            T.Eq("A4*2", s.GetCell(3, 5)!.Formula);
            w.Recalculate();
            T.Eq(14.0, s.GetValue(3, 5).Num);
            // fill right
            s.SetValue(10, 0, CellValue.Number(5)); s.SetValue(10, 1, CellValue.Number(10));
            SheetOps.Fill(s, new CellRange(10, 0, 10, 1), new CellRange(10, 0, 10, 4));
            T.Eq(25.0, s.GetValue(10, 4).Num, "fill right");
            // copy/paste with formula shift
            var clip = SheetOps.Copy(s, new CellRange(0, 5, 1, 5), false);
            SheetOps.Paste(s, clip, CellRange.Single(20, 6));
            T.Eq("B21*2", s.GetCell(20, 6)!.Formula);
            // paste values
            w.Recalculate();
            var clip2 = SheetOps.Copy(s, CellRange.Single(0, 5), false);
            SheetOps.Paste(s, clip2, CellRange.Single(30, 0), PasteMode.Values);
            T.True(s.GetCell(30, 0)!.Formula == null && s.GetValue(30, 0).Num == 2, "paste values");
            // transpose
            var clip3 = SheetOps.Copy(s, new CellRange(0, 0, 2, 0), false);
            SheetOps.Paste(s, clip3, CellRange.Single(40, 0), PasteMode.All, transpose: true);
            T.Eq(5.0, s.GetValue(40, 2).Num, "transpose");
            // cut moves references
            s.SetValue(50, 0, CellValue.Number(7)); s.SetFormula(50, 1, "A51*3");
            var cut = SheetOps.Copy(s, CellRange.Single(50, 0), true);
            SheetOps.Paste(s, cut, CellRange.Single(55, 0));
            T.Eq("A56*3", s.GetCell(50, 1)!.Formula);
            T.True(s.GetCell(50, 0) == null, "cut source cleared");
        });
        t.Test("Sort, filter, remove duplicates, text to columns, find/replace", () =>
        {
            var w = Workbook.CreateBlank();
            var s = w.Sheets[0];
            SheetOps.SetInput(s, 0, 0, "Tên"); SheetOps.SetInput(s, 0, 1, "Điểm");
            string[] names = ["Bình", "An", "Cường", "An", "Dũng"];
            double[] pts = [7, 9, 5, 9, 8];
            for (int i = 0; i < 5; i++) { s.SetValue(i + 1, 0, CellValue.Text(names[i])); s.SetValue(i + 1, 1, CellValue.Number(pts[i])); }
            SheetOps.Sort(s, new CellRange(0, 0, 5, 1), [new SheetOps.SortKey(1, false), new SheetOps.SortKey(0, true)], hasHeader: true);
            T.Eq("An", s.GetValue(1, 0).Str); T.Eq(9.0, s.GetValue(1, 1).Num); T.Eq("Cường", s.GetValue(5, 0).Str);
            T.Eq("Tên", s.GetValue(0, 0).Str, "header kept");
            SheetOps.Sort(s, new CellRange(0, 0, 5, 1), [new SheetOps.SortKey(0, true)], hasHeader: true);
            T.Eq("An", s.GetValue(1, 0).Str); T.Eq("Bình", s.GetValue(3, 0).Str, "vi collation: An, An, Bình");
            s.Filter = new AutoFilter { Range = new CellRange(0, 0, 5, 1) };
            s.Filter.Conditions[1] = ">7";
            SheetOps.ApplyFilter(s);
            T.Eq(2, s.FilteredRows.Count, "filtered rows");
            w.Recalculate();
            T.Eq(26.0, w.Calc.EvaluateFormula(s, 10, 0, "SUBTOTAL(9,B2:B6)").Num, "subtotal ignores filtered");
            s.Filter.Conditions.Clear();
            s.Filter.Allowed[0] = new HashSet<string> { "An" };
            SheetOps.ApplyFilter(s);
            T.Eq(3, s.FilteredRows.Count, "value filter");
            s.Filter = null; s.FilteredRows.Clear();
            int removed = SheetOps.RemoveDuplicates(s, new CellRange(0, 0, 5, 1), [0, 1], true);
            T.Eq(1, removed, "duplicates removed");
            T.Eq(4, Enumerable.Range(1, 5).Count(r => !s.GetValue(r, 0).IsBlank));
            s.SetValue(10, 4, CellValue.Text("Nguyễn;Văn;An"));
            SheetOps.TextToColumns(s, 4, 10, 10, ';', false);
            T.Eq("An", s.GetValue(10, 6).Str, "text to columns");
            var found = SheetOps.FindAll(w, s, new SheetOps.FindOptions("an", InFormulas: false)).ToList();
            T.True(found.Count >= 2, "find");
            int rep = SheetOps.ReplaceAll(w, s, new SheetOps.FindOptions("Văn", MatchCase: true, InFormulas: false), "Thị");
            T.Eq(1, rep); T.Eq("Thị", s.GetValue(10, 5).Str);
        });
        t.Test("Number formatting (en & vi)", () =>
        {
            void F(double v, string code, string exp) { var r = NumFmt.FormatNumber(v, code, out _); if (r != exp) throw new AssertException($"{v} '{code}' => '{r}' expected '{exp}'"); asserts++; }
            NumFmt.UseLocale(false);
            F(1234.5, "General", "1234.5"); F(1234567.891, "#,##0.00", "1,234,567.89"); F(-1234, "#,##0", "-1,234"); F(0.125, "0.0%", "12.5%"); F(5, "00000", "00005");
            F(1234.5, "#,##0 \"₫\"", "1,235 ₫"); F(-5, "#,##0;(#,##0)", "(5)"); F(0, "#,##0;(#,##0);\"-\"", "-"); F(12345.678, "0.00E+00", "1.23E+04");
            F(0.75, "# ?/?", "3/4"); F(1.5, "# ?/?", "1 1/2"); F(1234567, "#,##0,\"K\"", "1,235K"); F(3.14159, "0.###", "3.142"); F(3, "0.0#", "3.0");
            var d = DateSerial.ToSerial(new DateTime(2024, 3, 5, 14, 7, 9));
            F(d, "dd/mm/yyyy", "05/03/2024"); F(d, "d-mmm-yy", "5-Mar-24"); F(d, "hh:mm:ss", "14:07:09"); F(d, "h:mm AM/PM", "2:07 PM"); F(d, "yyyy-mm-dd hh:mm", "2024-03-05 14:07");
            F(1.5, "[h]:mm", "36:00"); F(d, "dddd", "Tuesday"); F(d, "mmmm", "March");
            T.Eq("[Red]", "[Red]");
            var col = NumFmt.FormatNumber(-3, "0;[Red]-0", out var color);
            T.Eq("-3", col); T.True(color == System.Drawing.Color.Red, "red colour");
            T.Eq("0.00", NumFmt.AdjustDecimals("0.0", 1, 1)); T.Eq("0", NumFmt.AdjustDecimals("0.0", 1, -1)); T.Eq("#,##0.0", NumFmt.AdjustDecimals("#,##0", 1, 1));
            NumFmt.UseLocale(true);
            F(1234567.891, "#,##0.00", "1.234.567,89"); F(1234.5, "General", "1234,5"); F(d, "dd/mm/yyyy", "05/03/2024");
            var v = ValueParser.Parse("1.234,5", out _); T.Eq(1234.5, v.Num, "vi input");
            v = ValueParser.Parse("1.5", out _); T.Eq(1.5, v.Num, "vi input with dot decimal");
            v = ValueParser.Parse("15/3/2024", out var fmt); T.Eq(DateSerial.ToSerial(new DateTime(2024, 3, 15)), v.Num, "vi date"); T.Eq("dd/mm/yyyy", fmt);
            v = ValueParser.Parse("12%", out fmt); T.Eq(0.12, v.Num, "percent"); T.Eq("0%", fmt);
            v = ValueParser.Parse("1.000.000 ₫", out fmt); T.Eq(1000000.0, v.Num, "vnd"); T.True(fmt!.Contains("₫"), "vnd fmt");
            NumFmt.UseLocale(false);
            v = ValueParser.Parse("3/15/2024", out _); T.Eq(DateSerial.ToSerial(new DateTime(2024, 3, 15)), v.Num, "en date");
            v = ValueParser.Parse("1,234.5", out _); T.Eq(1234.5, v.Num, "en input");
            v = ValueParser.Parse("abc", out _); T.True(v.IsText, "text");
            v = ValueParser.Parse("'0123", out _); T.Eq("0123", v.Str, "apostrophe text");
        });
        t.Log($"      formula/format assertions executed: {asserts}");
        if (asserts < 150) throw new Exception("too few assertions: " + asserts);
    }
}
