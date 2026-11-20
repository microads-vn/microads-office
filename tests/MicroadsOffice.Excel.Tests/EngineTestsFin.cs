using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Tests;

/// <summary>
/// Financial functions. Expected values are Microsoft's documented examples or were computed independently in Python
/// (scipy) from Microsoft's documented formulas.
/// </summary>
public static class EngineTestsFin
{
    private static Workbook NewData()
    {
        var wb = Workbook.CreateBlank(1);
        var s = wb.Sheets[0];
        // MIRR / XNPV / XIRR data: A1:A6 cash flows, B1:B5 dates, C1:C3 FVSCHEDULE rates (C4 blank, C5 text)
        double[] flows = [-120000, 39000, 30000, 21000, 37000, 46000];
        for (int i = 0; i < flows.Length; i++) s.SetValue(i, 0, CellValue.Number(flows[i]));
        DateTime[] dates = [new(2008, 1, 1), new(2008, 3, 1), new(2008, 10, 30), new(2009, 2, 15), new(2009, 4, 1)];
        for (int i = 0; i < dates.Length; i++) s.SetValue(i, 1, CellValue.Number(DateSerial.ToSerial(dates[i])));
        double[] xflows = [-10000, 2750, 4250, 3250, 2750];
        for (int i = 0; i < xflows.Length; i++) s.SetValue(i, 3, CellValue.Number(xflows[i]));
        s.SetValue(0, 2, CellValue.Number(0.09)); s.SetValue(1, 2, CellValue.Number(0.11)); s.SetValue(2, 2, CellValue.Number(0.1));
        s.SetValue(4, 2, CellValue.Text("x"));
        s.SetValue(5, 3, CellValue.Text("abc"));
        wb.Recalculate();
        return wb;
    }

    private static CellValue Ev(Workbook wb, string f) => wb.Calc.EvaluateFormula(wb.Sheets[0], 19, 9, f);

    public static void Run(TestRunner t)
    {
        NumFmt.UseLocale(false);
        var wb = NewData();
        void N(string f, double expected, double tol = 1e-9)
        {
            var v = Ev(wb, f);
            if (!v.IsNumber || Math.Abs(v.Num - expected) > tol * Math.Max(1, Math.Abs(expected)))
                throw new AssertException($"{f} => {v.Kind}:{v} expected {expected}");
        }
        void E(string f, string code) { var v = Ev(wb, f); if (!v.IsError || v.Str != code) throw new AssertException($"{f} => {v} expected {code}"); }
        double D(int y, int m, int d) => DateSerial.ToSerial(new DateTime(y, m, d));

        t.Test("Financial: coupon dates", () =>
        {
            const string sm = "DATE(2011,1,25),DATE(2011,11,15),2";
            N($"COUPDAYBS({sm},1)", 71); N($"COUPDAYS({sm},1)", 181); N($"COUPDAYSNC({sm},1)", 110);
            N($"COUPNCD({sm},1)", D(2011, 5, 15)); N($"COUPPCD({sm},1)", D(2010, 11, 15));
            N("COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)", 4);
            N($"COUPDAYS({sm},0)", 180); N($"COUPDAYS({sm},3)", 182.5); N($"COUPDAYBS({sm})", 70);
            // end-of-month maturity: quarterly coupons fall on the last day of the month
            N("COUPPCD(DATE(2011,1,25),DATE(2011,8,31),4)", D(2010, 11, 30)); N("COUPNCD(DATE(2011,1,25),DATE(2011,8,31),4)", D(2011, 2, 28));
            N("COUPNUM(DATE(2011,1,25),DATE(2011,8,31),4)", 3); N("COUPDAYBS(DATE(2011,1,25),DATE(2011,8,31),4,0)", 55);
            N("COUPNCD(40568.7,40862.2,2.9)", D(2011, 5, 15)); // dates and frequency are truncated
            E($"COUPDAYS({sm},5)", Errors.Num); E("COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),3)", Errors.Num);
            E("COUPNUM(DATE(2011,11,15),DATE(2011,11,15),2)", Errors.Num); E("COUPNCD(\"x\",DATE(2011,11,15),2)", Errors.Value);
            E("COUPPCD(-1,DATE(2011,11,15),2)", Errors.Num);
        });
        t.Test("Financial: accrued interest", () =>
        {
            N("ACCRINT(39508,39691,39569,0.1,1000,2,0)", 16.666666666666668);
            N("ACCRINT(DATE(2008,3,5),39691,39569,0.1,1000,2,0,FALSE)", 15.555555555555555);
            N("ACCRINT(DATE(2008,4,5),39691,39569,0.1,1000,2,0,TRUE)", 7.222222222222222);
            N("ACCRINT(39508,39691,39569,0.1,1000,2,1)", 16.57608695652174);
            N("ACCRINT(39508,39691,39569,0.1,,2)", 16.666666666666668); // omitted par = 1000
            E("ACCRINT(39508,39691,39569,0,1000,2,0)", Errors.Num); E("ACCRINT(39569,39691,39569,0.1,1000,2,0)", Errors.Num);
            E("ACCRINT(39508,39691,39569,0.1,1000,3,0)", Errors.Num);
            N("ACCRINTM(39539,39614,0.1,1000,3)", 20.54794520547945);
            N("ACCRINTM(39539,39614,0.1)", 1000 * 0.1 * 74 / 360.0); // 30/360: 1 Apr .. 15 Jun = 74 days
            E("ACCRINTM(39614,39539,0.1,1000,3)", Errors.Num); E("ACCRINTM(39539,39614,0.1,-5,3)", Errors.Num);
        });
        t.Test("Financial: bond price, yield, duration", () =>
        {
            N("PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)", 94.63436162132214);
            N("PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,1)", 94.63544920787722);
            N("PRICE(DATE(2008,2,15),DATE(2017,11,30),0.0575,0.065,100,4,4)", 94.59697220394641);
            N("PRICE(DATE(2008,2,15),DATE(2008,5,15),0.0575,0.065,100,2)", 99.792512300123);
            E("PRICE(DATE(2008,2,15),DATE(2017,11,15),-0.01,0.065,100,2,0)", Errors.Num);
            E("PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,0,2,0)", Errors.Num);
            N("YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)", 0.06500000688075455);
            N("YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,1)", 0.0650018206055473);
            N("YIELD(DATE(2008,2,15),DATE(2008,5,15),0.0575,99.5,100,2,0)", 0.0767801857585145);
            E("YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,0,100,2,0)", Errors.Num);
            N("DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", 5.993774955545185);
            N("DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)", 10.919145281591925);
            N("MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)", 5.735669813918838);
            E("DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,-0.09,2,1)", Errors.Num);
            E("MDURATION(DATE(2016,1,1),DATE(2008,1,1),0.08,0.09,2,1)", Errors.Num);
        });
        t.Test("Financial: discounted and at-maturity securities", () =>
        {
            N("DISC(DATE(2007,1,25),DATE(2007,6,15),97.975,100,1)", 0.05242021276595759);
            E("DISC(DATE(2007,1,25),DATE(2007,6,15),0,100,1)", Errors.Num); E("DISC(DATE(2007,6,15),DATE(2007,1,25),97.975,100,1)", Errors.Num);
            N("INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)", 0.05768);
            E("INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,5)", Errors.Num);
            N("RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)", 1014584.6544071021);
            E("RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0,2)", Errors.Num);
            N("PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)", 99.79583333333333);
            E("PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),-0.0525,100,2)", Errors.Num);
            N("YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)", 0.05282257198685834);
            E("YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,0,2)", Errors.Num);
            N("PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)", 99.98449887555694);
            E("PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),-0.061,0.061,0)", Errors.Num);
            N("YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)", 0.06095433369153868);
            E("YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,0,0)", Errors.Num);
        });
        t.Test("Financial: treasury bills", () =>
        {
            N("TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)", 0.09415149356594302);
            N("TBILLEQ(DATE(2008,3,31),DATE(2008,12,31),0.0914)", 0.0980096846650606); // > 182 days: bond-equivalent quadratic
            N("TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)", 98.45);
            N("TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)", 0.09141696292534264);
            E("TBILLEQ(DATE(2008,3,31),DATE(2009,6,1),0.0914)", Errors.Num); E("TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0)", Errors.Num);
            E("TBILLYIELD(DATE(2008,6,1),DATE(2008,3,31),98.45)", Errors.Num);
        });
        t.Test("Financial: odd first / last period", () =>
        {
            N("ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)", 113.59771747407883);
            N("ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)", 113.48952645174103); // long first period
            N("ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)", 113.48894083961049);
            E("ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,12,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)", Errors.Num);
            E("ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-0.0625,100,2,1)", Errors.Num);
            N("ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)", 0.07724554159781737);
            E("ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,0,100,2,0)", Errors.Num);
            N("ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)", 99.87828601472134);
            N("ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,1)", 99.87916768152911);
            E("ODDLPRICE(DATE(2007,10,15),DATE(2008,6,15),DATE(2008,2,7),0.0375,0.0405,100,2,0)", Errors.Num);
            N("ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)", 0.04519223562916916);
            E("ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,3,0)", Errors.Num);
        });
        t.Test("Financial: loans and rates", () =>
        {
            N("CUMIPMT(0.09/12,30*12,125000,13,24,0)", -11135.232130750846); N("CUMIPMT(0.09/12,30*12,125000,1,1,0)", -937.5);
            N("CUMIPMT(0.09/12,360,125000,13,24,1)", -11052.339583871817);
            N("CUMPRINC(0.09/12,30*12,125000,13,24,0)", -934.1071234208741); N("CUMPRINC(0.09/12,30*12,125000,1,1,0)", -68.27827118097673);
            N("CUMPRINC(0.09/12,360,125000,13,24,1)", -927.1534723780305);
            E("CUMIPMT(0.09/12,360,125000,24,13,0)", Errors.Num); E("CUMIPMT(0.09/12,360,125000,13,24,2)", Errors.Num);
            E("CUMPRINC(0.09/12,360,-125000,13,24,0)", Errors.Num); E("CUMPRINC(0.09/12,360,125000,0,24,0)", Errors.Num);
            N("ISPMT(0.1/12,1,3*12,8000000)", -64814.81481481482); N("ISPMT(0.1,1,3,8000000)", -533333.3333333334);
            E("ISPMT(0.1,1,0,8000000)", Errors.Div0);
            N("EFFECT(0.0525,4)", 0.05354266737075819); N("EFFECT(0.0525,4.9)", 0.05354266737075819); E("EFFECT(0.0525,0.5)", Errors.Num); E("EFFECT(0,4)", Errors.Num);
            N("NOMINAL(0.053543,4)", 0.052500319868356016); E("NOMINAL(-0.05,4)", Errors.Num);
            N("FVSCHEDULE(1,{0.09,0.11,0.1})", 1.3308900000000004); N("FVSCHEDULE(1,C1:C4)", 1.3308900000000004);
            E("FVSCHEDULE(1,C1:C5)", Errors.Value);
            N("PDURATION(0.025,2000,2200)", 3.859866162622662); N("PDURATION(0.025/12,1000,1200)", 87.60547641937126); E("PDURATION(0,2000,2200)", Errors.Num);
            N("RRI(96,10000,11000)", 0.0009933073762913303); E("RRI(0,10000,11000)", Errors.Num);
            N("MIRR(A1:A6,0.1,0.12)", 0.1260941303659051); N("MIRR(A1:A4,0.1,0.12)", -0.048044655249980806); N("MIRR(A1:A6,0.1,0.14)", 0.13475911082831482);
            E("MIRR(A2:A6,0.1,0.12)", Errors.Div0);
        });
        t.Test("Financial: XNPV and XIRR", () =>
        {
            N("XNPV(0.09,D1:D5,B1:B5)", 2086.6476020315354);
            N("XIRR(D1:D5,B1:B5)", 0.3733625335188314); N("XIRR(D1:D5,B1:B5,0.5)", 0.3733625335188314);
            N("XIRR({-1000,1100},{39448,39813})", 0.1, 1e-12); // exactly 365 days apart
            E("XNPV(0.09,D1:D5,B1:B4)", Errors.Num); E("XIRR(D2:D5,B2:B5)", Errors.Num); E("XNPV(0.09,D1:D6,B1:B6)", Errors.Value);
            E("XNPV(0.09,{-100,50},{39448,39000})", Errors.Num);
            N("XIRR(D1:D5,B1:B5,-0.9)", 0.3733625335188314); N("XIRR(D1:D5,B1:B5,5)", 0.3733625335188314); // poor guesses still converge
        });
        t.Test("Financial: solver round trips", () =>
        {
            // yield solvers must invert the price formulas
            N("YIELD(DATE(2010,3,10),DATE(2030,9,30),0.04,PRICE(DATE(2010,3,10),DATE(2030,9,30),0.04,0.07,100,4,1),100,4,1)", 0.07);
            N("YIELD(DATE(2010,3,10),DATE(2030,9,30),0,PRICE(DATE(2010,3,10),DATE(2030,9,30),0,0.03,105,1,3),105,1,3)", 0.03);
            N("ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,113.48894083961049,100,2,1)", 0.0625);
            N("ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,1,15),DATE(2009,3,1),0.0785,0.11,100,2,4),100,2,4)", 0.11);
            N("ODDLYIELD(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,99.87916768152911,100,2,1)", 0.0405);
        });
        t.Test("Financial: depreciation", () =>
        {
            double[] db = [186083.33333333334, 259639.41666666666, 176814.44275000002, 120410.63551274998, 81999.64278418274, 55841.75673602846, 15845.098473848071];
            for (int p = 1; p <= 7; p++) N($"DB(1000000,100000,6,{p},7)", db[p - 1]);
            E("DB(1000000,100000,6,7)", Errors.Num); E("DB(1000000,100000,6,1,13)", Errors.Num); E("DB(-1,100000,6,1)", Errors.Num);
            N("VDB(2400,300,10*365,0,1)", 1.3150684931506849); N("VDB(2400,300,10*12,0,1)", 40); N("VDB(2400,300,10,0,1)", 480);
            N("VDB(2400,300,10*12,6,18)", 396.3060532647509); N("VDB(2400,300,10*12,6,18,1.5)", 311.8089366582341); N("VDB(2400,300,10,0,0.875,1.5)", 315);
            N("VDB(10000,0,5,0,5)", 10000); N("VDB(10000,0,5,0,5,2,TRUE)", 9222.4); N("VDB(10000,0,5,3,5)", 2160); N("VDB(10000,0,5,3,5,2,TRUE)", 1382.4);
            N("VDB(10000,1000,5,3,5,1.5)", 2430); N("VDB(10000,1000,5,3,5,1.5,TRUE)", 1749.3);
            E("VDB(2400,300,10,0,11)", Errors.Num); E("VDB(2400,300,10,5,4)", Errors.Num); E("VDB(2400,300,10,0,1,0)", Errors.Num);
            N("AMORDEGRC(2400,39679,39813,300,1,0.15,1)", 776); N("AMORDEGRC(2400,39679,39813,300,0,0.15,1)", 330);
            E("AMORDEGRC(2400,39679,39813,300,1,0.22,1)", Errors.Num); // life 4.5 years
            E("AMORDEGRC(2400,39679,39813,300,1,0.15,2)", Errors.Num);
            N("AMORLINC(2400,39679,39813,300,1,0.15,1)", 360); N("AMORLINC(2400,39679,39813,300,0,0.15,1)", 131.8032786885246);
            N("AMORLINC(2400,39679,39813,300,6,0.15,1)", 168.1967213114754); N("AMORLINC(2400,39679,39813,300,7,0.15,1)", 0);
            E("AMORLINC(2400,39813,39679,300,1,0.15,1)", Errors.Num); E("AMORLINC(2400,39679,39813,300,1,0,1)", Errors.Num);
        });
        t.Test("Financial: fractional dollar prices", () =>
        {
            N("DOLLARDE(1.02,16)", 1.125); N("DOLLARDE(1.1,32)", 1.3125); N("DOLLARDE(-1.02,16)", -1.125); N("DOLLARDE(1.02,16.9)", 1.125);
            N("DOLLARFR(1.125,16)", 1.02); N("DOLLARFR(1.125,32)", 1.04); N("DOLLARFR(1.5,10)", 1.5);
            E("DOLLARDE(1.02,0.5)", Errors.Div0); E("DOLLARDE(1.02,-1)", Errors.Num); E("DOLLARFR(1.125,0)", Errors.Div0);
        });
        t.Test("Financial: catalog entries", () =>
        {
            foreach (var f in FunctionCatalog.All.Where(f => f.Category == FnCategory.Financial))
                if (!Functions.Table.ContainsKey(f.Name)) throw new AssertException($"{f.Name} is in the catalog but not implemented");
            foreach (var n in new[] { "ACCRINT", "ODDFPRICE", "XIRR", "VDB", "YIELDMAT" })
                if (FunctionCatalog.Find(n) is not { Category: FnCategory.Financial }) throw new AssertException($"{n} missing from the catalog");
        });
    }
}
