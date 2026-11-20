using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

/// <summary>Financial functions: bonds and securities, depreciation, cash-flow schedules, T-bills.</summary>
public static partial class Functions
{
    internal static void RegisterFin()
    {
        static CellValue F(double d) => CellValue.Number(FinMath.Check(d));

        // ------------------------------------------------------------ coupon dates
        Reg((c, a) => { Args(a, 3, 4); var b = FinMath.Bond(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0)); var q = FinMath.Coupons(b.S, b.M, b.F); return F(FinMath.Days(q.Pcd, b.S, b.B)); }, "COUPDAYBS");
        Reg((c, a) => { Args(a, 3, 4); var b = FinMath.Bond(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0)); var q = FinMath.Coupons(b.S, b.M, b.F); return F(FinMath.CoupDays(q.Pcd, q.Ncd, b.F, b.B)); }, "COUPDAYS");
        Reg((c, a) =>
        {
            Args(a, 3, 4);
            var b = FinMath.Bond(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0));
            var q = FinMath.Coupons(b.S, b.M, b.F);
            return F(b.B switch
            {
                0 => FinMath.Days360Us(q.Pcd, q.Ncd, true) - FinMath.Days360Us(q.Pcd, b.S, true),
                4 => FinMath.Days360Eu(b.S, q.Ncd),
                _ => FinMath.Act(b.S, q.Ncd)
            });
        }, "COUPDAYSNC");
        Reg((c, a) => { Args(a, 3, 4); var b = FinMath.Bond(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0)); return F(FinMath.Ser(FinMath.Coupons(b.S, b.M, b.F).Ncd)); }, "COUPNCD");
        Reg((c, a) => { Args(a, 3, 4); var b = FinMath.Bond(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0)); return F(FinMath.Ser(FinMath.Coupons(b.S, b.M, b.F).Pcd)); }, "COUPPCD");
        Reg((c, a) => { Args(a, 3, 4); var b = FinMath.Bond(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0)); return F(FinMath.Coupons(b.S, b.M, b.F).Num); }, "COUPNUM");

        // ------------------------------------------------------------ accrued interest
        Reg((c, a) =>
        {
            Args(a, 6, 8);
            double iss = c.Num(a[0]), fi = c.Num(a[1]), st = c.Num(a[2]), rate = c.Num(a[3]), par = c.NumOr(a, 4, 1000), fq = c.Num(a[5]), bs = c.NumOr(a, 6, 0);
            bool calc = c.BoolOr(a, 7, true);
            var issue = FinMath.Date(iss); var first = FinMath.Date(fi); var settle = FinMath.Date(st);
            int f = FinMath.Freq(fq), basis = FinMath.Basis(bs);
            if (rate <= 0 || par <= 0 || issue >= settle) throw Err(Errors.Num);
            var start = !calc && settle > first ? first : issue;
            return F(par * rate / f * FinMath.AccrualPeriods(start, settle, first, f, basis));
        }, "ACCRINT");
        Reg((c, a) =>
        {
            Args(a, 3, 5);
            double iss = c.Num(a[0]), st = c.Num(a[1]), rate = c.Num(a[2]), par = c.NumOr(a, 3, 1000), bs = c.NumOr(a, 4, 0);
            var issue = FinMath.Date(iss); var settle = FinMath.Date(st); int basis = FinMath.Basis(bs);
            if (rate <= 0 || par <= 0 || issue >= settle) throw Err(Errors.Num);
            return F(par * rate * FinMath.YearFrac(issue, settle, basis));
        }, "ACCRINTM");

        // ------------------------------------------------------------ coupon bonds
        Reg((c, a) =>
        {
            Args(a, 6, 7);
            double st = c.Num(a[0]), mt = c.Num(a[1]), rate = c.Num(a[2]), yld = c.Num(a[3]), red = c.Num(a[4]), fq = c.Num(a[5]), bs = c.NumOr(a, 6, 0);
            var b = FinMath.Bond(st, mt, fq, bs);
            if (rate < 0 || yld < 0 || red <= 0) throw Err(Errors.Num);
            return F(FinMath.Price(b.S, b.M, rate, yld, red, b.F, b.B));
        }, "PRICE");
        Reg((c, a) =>
        {
            Args(a, 6, 7);
            double st = c.Num(a[0]), mt = c.Num(a[1]), rate = c.Num(a[2]), pr = c.Num(a[3]), red = c.Num(a[4]), fq = c.Num(a[5]), bs = c.NumOr(a, 6, 0);
            var b = FinMath.Bond(st, mt, fq, bs);
            if (rate < 0 || pr <= 0 || red <= 0) throw Err(Errors.Num);
            return F(FinMath.Yield(b.S, b.M, rate, pr, red, b.F, b.B));
        }, "YIELD");
        FnImpl duration(bool modified) => (c, a) =>
        {
            Args(a, 5, 6);
            double st = c.Num(a[0]), mt = c.Num(a[1]), cpn = c.Num(a[2]), yld = c.Num(a[3]), fq = c.Num(a[4]), bs = c.NumOr(a, 5, 0);
            var b = FinMath.Bond(st, mt, fq, bs);
            if (cpn < 0 || yld < 0) throw Err(Errors.Num);
            double d = FinMath.Duration(b.S, b.M, cpn, yld, b.F, b.B);
            return F(modified ? d / (1 + yld / b.F) : d);
        };
        Reg(duration(false), "DURATION");
        Reg(duration(true), "MDURATION");

        // ------------------------------------------------------------ discounted / at-maturity securities
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double st = c.Num(a[0]), mt = c.Num(a[1]), pr = c.Num(a[2]), red = c.Num(a[3]), bs = c.NumOr(a, 4, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs);
            if (pr <= 0 || red <= 0) throw Err(Errors.Num);
            return F((red - pr) / red / FinMath.YearFrac(s, m, basis));
        }, "DISC");
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double st = c.Num(a[0]), mt = c.Num(a[1]), inv = c.Num(a[2]), red = c.Num(a[3]), bs = c.NumOr(a, 4, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs);
            if (inv <= 0 || red <= 0) throw Err(Errors.Num);
            return F((red - inv) / inv / FinMath.YearFrac(s, m, basis));
        }, "INTRATE");
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double st = c.Num(a[0]), mt = c.Num(a[1]), inv = c.Num(a[2]), disc = c.Num(a[3]), bs = c.NumOr(a, 4, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs);
            if (inv <= 0 || disc <= 0) throw Err(Errors.Num);
            double den = 1 - disc * FinMath.YearFrac(s, m, basis);
            if (den <= 0) throw Err(Errors.Num);
            return F(inv / den);
        }, "RECEIVED");
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double st = c.Num(a[0]), mt = c.Num(a[1]), disc = c.Num(a[2]), red = c.Num(a[3]), bs = c.NumOr(a, 4, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs);
            if (disc <= 0 || red <= 0) throw Err(Errors.Num);
            return F(red * (1 - disc * FinMath.YearFrac(s, m, basis)));
        }, "PRICEDISC");
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double st = c.Num(a[0]), mt = c.Num(a[1]), pr = c.Num(a[2]), red = c.Num(a[3]), bs = c.NumOr(a, 4, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs);
            if (pr <= 0 || red <= 0) throw Err(Errors.Num);
            return F((red - pr) / pr / FinMath.YearFrac(s, m, basis));
        }, "YIELDDISC");
        Reg((c, a) =>
        {
            Args(a, 5, 6);
            double st = c.Num(a[0]), mt = c.Num(a[1]), iss = c.Num(a[2]), rate = c.Num(a[3]), yld = c.Num(a[4]), bs = c.NumOr(a, 5, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs); var issue = FinMath.Date(iss);
            if (rate < 0 || yld < 0) throw Err(Errors.Num);
            var (b, dim, acc, dsm) = FinMath.MatFactors(s, m, issue, basis);
            return F((100 + dim / b * rate * 100) / (1 + dsm / b * yld) - acc / b * rate * 100);
        }, "PRICEMAT");
        Reg((c, a) =>
        {
            Args(a, 5, 6);
            double st = c.Num(a[0]), mt = c.Num(a[1]), iss = c.Num(a[2]), rate = c.Num(a[3]), pr = c.Num(a[4]), bs = c.NumOr(a, 5, 0);
            var (s, m, basis) = FinMath.Span(st, mt, bs); var issue = FinMath.Date(iss);
            if (rate < 0 || pr <= 0) throw Err(Errors.Num);
            var (b, dim, acc, dsm) = FinMath.MatFactors(s, m, issue, basis);
            double p = pr / 100 + acc / b * rate;
            return F((1 + dim / b * rate - p) / p * b / dsm);
        }, "YIELDMAT");

        // ------------------------------------------------------------ treasury bills
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double st = c.Num(a[0]), mt = c.Num(a[1]), disc = c.Num(a[2]);
            double dsm = FinMath.TBillDays(st, mt);
            if (disc <= 0) throw Err(Errors.Num);
            if (dsm <= 182) return F(365 * disc / (360 - disc * dsm));
            double price = 1 - disc * dsm / 360;
            if (price <= 0) throw Err(Errors.Num);
            double y = dsm == 366 ? 366 : 365, t = dsm / y;
            return F(2 * (-t + Math.Sqrt(t * t - (2 * t - 1) * (1 - 1 / price))) / (2 * t - 1));
        }, "TBILLEQ");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double st = c.Num(a[0]), mt = c.Num(a[1]), disc = c.Num(a[2]);
            double dsm = FinMath.TBillDays(st, mt);
            if (disc <= 0) throw Err(Errors.Num);
            double p = 100 * (1 - disc * dsm / 360);
            if (p <= 0) throw Err(Errors.Num);
            return F(p);
        }, "TBILLPRICE");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double st = c.Num(a[0]), mt = c.Num(a[1]), pr = c.Num(a[2]);
            double dsm = FinMath.TBillDays(st, mt);
            if (pr <= 0) throw Err(Errors.Num);
            return F((100 - pr) / pr * 360 / dsm);
        }, "TBILLYIELD");

        // ------------------------------------------------------------ odd first / last period bonds
        Reg((c, a) =>
        {
            Args(a, 8, 9);
            double st = c.Num(a[0]), mt = c.Num(a[1]), iss = c.Num(a[2]), fc = c.Num(a[3]), rate = c.Num(a[4]), yld = c.Num(a[5]), red = c.Num(a[6]), fq = c.Num(a[7]), bs = c.NumOr(a, 8, 0);
            var o = FinMath.OddFirst(st, mt, iss, fc, fq, bs);
            if (rate < 0 || yld < 0 || red <= 0) throw Err(Errors.Num);
            return F(o.Price(rate, yld, red));
        }, "ODDFPRICE");
        Reg((c, a) =>
        {
            Args(a, 8, 9);
            double st = c.Num(a[0]), mt = c.Num(a[1]), iss = c.Num(a[2]), fc = c.Num(a[3]), rate = c.Num(a[4]), pr = c.Num(a[5]), red = c.Num(a[6]), fq = c.Num(a[7]), bs = c.NumOr(a, 8, 0);
            var o = FinMath.OddFirst(st, mt, iss, fc, fq, bs);
            if (rate < 0 || pr <= 0 || red <= 0) throw Err(Errors.Num);
            return F(FinMath.Solve(y => o.Price(rate, y, red) - pr, 0.05, -o.Freq));
        }, "ODDFYIELD");
        Reg((c, a) =>
        {
            Args(a, 7, 8);
            double st = c.Num(a[0]), mt = c.Num(a[1]), li = c.Num(a[2]), rate = c.Num(a[3]), yld = c.Num(a[4]), red = c.Num(a[5]), fq = c.Num(a[6]), bs = c.NumOr(a, 7, 0);
            var o = FinMath.OddLast(st, mt, li, fq, bs);
            if (rate < 0 || yld < 0 || red <= 0) throw Err(Errors.Num);
            double cpn = 100 * rate / o.F;
            return F((red + cpn * o.DC) / (1 + o.DSC * yld / o.F) - cpn * o.A);
        }, "ODDLPRICE");
        Reg((c, a) =>
        {
            Args(a, 7, 8);
            double st = c.Num(a[0]), mt = c.Num(a[1]), li = c.Num(a[2]), rate = c.Num(a[3]), pr = c.Num(a[4]), red = c.Num(a[5]), fq = c.Num(a[6]), bs = c.NumOr(a, 7, 0);
            var o = FinMath.OddLast(st, mt, li, fq, bs);
            if (rate < 0 || pr <= 0 || red <= 0) throw Err(Errors.Num);
            double cpn = 100 * rate / o.F, p = pr + cpn * o.A;
            return F((red + cpn * o.DC - p) / p * o.F / o.DSC);
        }, "ODDLYIELD");

        // ------------------------------------------------------------ loans, rates, cash flows
        FnImpl cumulative(bool interest) => (c, a) =>
        {
            Args(a, 6, 6);
            double r = c.Num(a[0]), n = c.Num(a[1]), pv = c.Num(a[2]), s = Math.Truncate(c.Num(a[3])), e = Math.Truncate(c.Num(a[4])), t = Math.Truncate(c.Num(a[5]));
            if (r <= 0 || n <= 0 || pv <= 0 || s < 1 || e < 1 || s > e || e > n || (t != 0 && t != 1)) throw Err(Errors.Num);
            double pmt = FinMath.Pmt(r, n, pv, t), sum = 0;
            for (double p = s; p <= e; p++)
            {
                double ip = FinMath.Ipmt(r, p, pv, pmt, t);
                sum += interest ? ip : pmt - ip;
            }
            return F(sum);
        };
        Reg(cumulative(true), "CUMIPMT");
        Reg(cumulative(false), "CUMPRINC");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double r = c.Num(a[0]), per = c.Num(a[1]), n = c.Num(a[2]), pv = c.Num(a[3]);
            if (n == 0) throw Err(Errors.Div0);
            return F(pv * r * (per / n - 1));
        }, "ISPMT");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double r = c.Num(a[0]), n = Math.Truncate(c.Num(a[1]));
            if (r <= 0 || n < 1) throw Err(Errors.Num);
            return F(Math.Pow(1 + r / n, n) - 1);
        }, "EFFECT");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double r = c.Num(a[0]), n = Math.Truncate(c.Num(a[1]));
            if (r <= 0 || n < 1) throw Err(Errors.Num);
            return F(n * (Math.Pow(1 + r, 1 / n) - 1));
        }, "NOMINAL");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double p = c.Num(a[0]);
            var v = c.Eval(a[1]);
            if (v is CellValue one) return F(p * (1 + Calc.ToNum(one)));
            foreach (var x in FlatVal(c, v))
            {
                if (x.IsError) throw Err(x.Str!);
                if (x.IsBlank) continue;
                if (!x.IsNumber) throw Err(Errors.Value);
                p *= 1 + x.Num;
            }
            return F(p);
        }, "FVSCHEDULE");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double r = c.Num(a[0]), pv = c.Num(a[1]), fv = c.Num(a[2]);
            if (r <= 0 || pv <= 0 || fv <= 0) throw Err(Errors.Num);
            return F((Math.Log(fv) - Math.Log(pv)) / Math.Log(1 + r));
        }, "PDURATION");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double n = c.Num(a[0]), pv = c.Num(a[1]), fv = c.Num(a[2]);
            if (n <= 0 || pv == 0) throw Err(Errors.Num);
            return F(Math.Pow(fv / pv, 1 / n) - 1);
        }, "RRI");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            var vals = Nums(c, [a[0]]);
            double fr = c.Num(a[1]), rr = c.Num(a[2]);
            int n = vals.Count;
            double pos = 0, neg = 0;
            for (int i = 0; i < n; i++)
            {
                if (vals[i] > 0) pos += vals[i] / Math.Pow(1 + rr, i);
                else if (vals[i] < 0) neg += vals[i] / Math.Pow(1 + fr, i);
            }
            if (pos == 0 || neg == 0 || n < 2) throw Err(Errors.Div0);
            return F(Math.Pow(-pos * Math.Pow(1 + rr, n - 1) / neg, 1.0 / (n - 1)) - 1);
        }, "MIRR");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double r = c.Num(a[0]);
            var (v, d) = FinMath.Schedule(c, a[1], a[2]);
            if (r <= -1) throw Err(Errors.Num);
            return F(FinMath.Xnpv(r, v, d));
        }, "XNPV");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var (v, d) = FinMath.Schedule(c, a[0], a[1]);
            double guess = c.NumOr(a, 2, 0.1);
            if (!v.Any(x => x > 0) || !v.Any(x => x < 0) || guess <= -1) throw Err(Errors.Num);
            return F(FinMath.Solve(r => FinMath.Xnpv(r, v, d), guess, -1));
        }, "XIRR");

        // ------------------------------------------------------------ depreciation
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double cost = c.Num(a[0]), salv = c.Num(a[1]), life = c.Num(a[2]), per = Math.Truncate(c.Num(a[3])), month = Math.Truncate(c.NumOr(a, 4, 12));
            if (cost < 0 || salv < 0 || life <= 0 || per <= 0 || month < 1 || month > 12 || per > life + 1 || (month == 12 && per > life)) throw Err(Errors.Num);
            if (cost == 0) return F(0);
            double rate = Math.Round(1 - Math.Pow(salv / cost, 1 / life), 3, MidpointRounding.AwayFromZero);
            double total = 0, dep = 0;
            for (int p = 1; p <= per; p++)
            {
                dep = p == 1 ? cost * rate * month / 12 : p > life ? (cost - total) * rate * (12 - month) / 12 : (cost - total) * rate;
                total += dep;
            }
            return F(dep);
        }, "DB");
        Reg((c, a) =>
        {
            Args(a, 5, 7);
            double cost = c.Num(a[0]), salv = c.Num(a[1]), life = c.Num(a[2]), s = c.Num(a[3]), e = c.Num(a[4]), factor = c.NumOr(a, 5, 2);
            bool noSwitch = c.BoolOr(a, 6, false);
            if (cost < 0 || salv < 0 || life <= 0 || s < 0 || e < s || e > life || factor <= 0) throw Err(Errors.Num);
            return F(FinMath.Vdb(cost, salv, life, s, e, factor, noSwitch));
        }, "VDB");
        FnImpl amor(bool degressive) => (c, a) =>
        {
            Args(a, 6, 7);
            double cost = c.Num(a[0]), dp = c.Num(a[1]), fp = c.Num(a[2]), salv = c.Num(a[3]), per = c.Num(a[4]), rate = c.Num(a[5]), bs = c.NumOr(a, 6, 0);
            var bought = FinMath.Date(dp); var first = FinMath.Date(fp); int basis = FinMath.Basis(bs, false);
            if (cost < 0 || salv < 0 || salv > cost || per < 0 || rate <= 0 || bought > first) throw Err(Errors.Num);
            double yf = FinMath.YearFrac(bought, first, basis);
            return F(degressive ? FinMath.Amordegrc(cost, yf, salv, (int)Math.Truncate(per), rate) : FinMath.Amorlinc(cost, yf, salv, (int)Math.Truncate(per), rate));
        };
        Reg(amor(true), "AMORDEGRC");
        Reg(amor(false), "AMORLINC");

        // ------------------------------------------------------------ fractional prices
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double x = c.Num(a[0]), f = FinMath.Fraction(c.Num(a[1]));
            double ip = Math.Truncate(x);
            return F(ip + (x - ip) * FinMath.Pow10Ceil(f) / f);
        }, "DOLLARDE");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double x = c.Num(a[0]), f = FinMath.Fraction(c.Num(a[1]));
            double ip = Math.Truncate(x);
            return F(ip + (x - ip) * f / FinMath.Pow10Ceil(f));
        }, "DOLLARFR");
    }
}

/// <summary>Day counts, coupon schedules and solvers for the financial functions (Excel conventions).</summary>
internal static class FinMath
{
    private static FormulaError NumErr() => new(Errors.Num);

    internal static double Check(double v) => double.IsNaN(v) || double.IsInfinity(v) ? throw NumErr() : v;

    // ------------------------------------------------------------------ argument checks

    /// <summary>A date argument: truncated to a whole day serial; negative or out-of-range serials are #NUM!.</summary>
    internal static DateTime Date(double serial)
    {
        if (double.IsNaN(serial) || serial < 0 || serial >= 2958466) throw NumErr();
        return DateSerial.FromSerial(Math.Floor(serial)).Date;
    }

    internal static double Ser(DateTime d) => DateSerial.ToSerial(d.Date);

    internal static int Freq(double f)
    {
        f = Math.Truncate(f);
        if (f != 1 && f != 2 && f != 4) throw NumErr();
        return (int)f;
    }

    internal static int Basis(double b, bool allowActual360 = true)
    {
        b = Math.Truncate(b);
        if (b < 0 || b > 4 || (!allowActual360 && b == 2)) throw NumErr();
        return (int)b;
    }

    /// <summary>settlement, maturity, frequency and basis of a coupon bond (settlement must precede maturity).</summary>
    internal static (DateTime S, DateTime M, int F, int B) Bond(double settlement, double maturity, double frequency, double basis)
    {
        var s = Date(settlement); var m = Date(maturity);
        int f = Freq(frequency), b = Basis(basis);
        if (s >= m) throw NumErr();
        return (s, m, f, b);
    }

    internal static (DateTime S, DateTime M, int B) Span(double settlement, double maturity, double basis)
    {
        var s = Date(settlement); var m = Date(maturity); int b = Basis(basis);
        if (s >= m) throw NumErr();
        return (s, m, b);
    }

    /// <summary>Days from settlement to maturity of a T-bill (at most one year).</summary>
    internal static double TBillDays(double settlement, double maturity)
    {
        var s = Date(settlement); var m = Date(maturity);
        if (s >= m || m > s.AddYears(1)) throw NumErr();
        return Act(s, m);
    }

    internal static double Fraction(double f)
    {
        f = Math.Truncate(f);
        if (f < 0) throw NumErr();
        if (f < 1) throw new FormulaError(Errors.Div0);
        return f;
    }

    /// <summary>10^ceil(log10(f)) for a whole number f ≥ 1.</summary>
    internal static double Pow10Ceil(double f)
    {
        double p = 1;
        while (p < f) p *= 10;
        return p;
    }

    // ------------------------------------------------------------------ day counts

    private static bool LastDay(DateTime d) => d.Day == DateTime.DaysInMonth(d.Year, d.Month);
    private static bool LastFeb(DateTime d) => d.Month == 2 && LastDay(d);

    internal static double Act(DateTime s, DateTime e) => Ser(e) - Ser(s);

    /// <summary>30/360 US (NASD) as Excel uses it for basis 0; <paramref name="both"/> also moves a month-end end date to the 30th.</summary>
    internal static int Days360Us(DateTime s, DateTime e, bool both = false)
    {
        int d1 = s.Day, d2 = e.Day;
        if (LastFeb(e) && (LastFeb(s) || both)) d2 = 30;
        if (d2 == 31 && (d1 >= 30 || both)) d2 = 30;
        if (d1 == 31) d1 = 30;
        if (LastFeb(s)) d1 = 30;
        return (e.Year - s.Year) * 360 + (e.Month - s.Month) * 30 + d2 - d1;
    }

    internal static int Days360Eu(DateTime s, DateTime e)
        => (e.Year - s.Year) * 360 + (e.Month - s.Month) * 30 + Math.Min(e.Day, 30) - Math.Min(s.Day, 30);

    /// <summary>Day count between two dates for a basis (0 = 30/360 US, 1/2/3 = actual, 4 = 30/360 European).</summary>
    internal static double Days(DateTime s, DateTime e, int basis) => basis switch
    {
        0 => Days360Us(s, e),
        4 => Days360Eu(s, e),
        _ => Act(s, e)
    };

    private static bool WithinYear(DateTime s, DateTime e)
        => s.Year == e.Year || (e.Year == s.Year + 1 && (s.Month > e.Month || (s.Month == e.Month && s.Day >= e.Day)));

    /// <summary>Year length used by basis 1 for the span s..e (as YEARFRAC).</summary>
    private static double ActualYear(DateTime s, DateTime e)
    {
        if (s > e) (s, e) = (e, s);
        if (WithinYear(s, e))
        {
            if (s.Year == e.Year) return DateTime.IsLeapYear(s.Year) ? 366 : 365;
            bool leap = (DateTime.IsLeapYear(s.Year) && s <= new DateTime(s.Year, 2, 29)) || (DateTime.IsLeapYear(e.Year) && e >= new DateTime(e.Year, 2, 29));
            return leap ? 366 : 365;
        }
        return (Ser(new DateTime(e.Year, 12, 31)) + 1 - Ser(new DateTime(s.Year, 1, 1))) / (e.Year - s.Year + 1);
    }

    internal static double DaysInYear(DateTime s, DateTime e, int basis) => basis switch
    {
        1 => ActualYear(s, e),
        3 => 365,
        _ => 360
    };

    internal static double YearFrac(DateTime s, DateTime e, int basis)
    {
        if (s > e) (s, e) = (e, s);
        return Days(s, e, basis) / DaysInYear(s, e, basis);
    }

    // ------------------------------------------------------------------ coupon schedules

    /// <summary>Anchor date moved by whole months; with the end-of-month rule the result is the last day of its month.</summary>
    internal static DateTime Shift(DateTime anchor, int months, bool eom)
    {
        DateTime d;
        try { d = anchor.AddMonths(months); }
        catch (ArgumentOutOfRangeException) { throw NumErr(); }
        return eom ? new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month)) : d;
    }

    /// <summary>Largest j with Shift(anchor, j*step) ≤ d.</summary>
    private static int FloorIndex(DateTime anchor, int step, bool eom, DateTime d)
    {
        int months = (d.Year - anchor.Year) * 12 + d.Month - anchor.Month;
        int j = (int)Math.Floor(months / (double)step) + 1;
        while (Shift(anchor, j * step, eom) > d) j--;
        while (Shift(anchor, (j + 1) * step, eom) <= d) j++;
        return j;
    }

    /// <summary>Previous and next coupon dates around settlement, and the number of coupons still payable.</summary>
    internal static (DateTime Pcd, DateTime Ncd, int Num) Coupons(DateTime settle, DateTime mat, int freq)
    {
        int step = 12 / freq; bool eom = LastDay(mat);
        int j = FloorIndex(mat, step, eom, settle);
        return (Shift(mat, j * step, eom), Shift(mat, (j + 1) * step, eom), -j);
    }

    /// <summary>Length in days of the coupon period pcd..ncd for a basis.</summary>
    internal static double CoupDays(DateTime pcd, DateTime ncd, int freq, int basis) => basis switch
    {
        1 => Act(pcd, ncd),
        3 => 365.0 / freq,
        _ => 360.0 / freq
    };

    /// <summary>Accrued fraction of coupon periods between start and end, quasi-coupon periods anchored on <paramref name="anchor"/>.</summary>
    internal static double AccrualPeriods(DateTime start, DateTime end, DateTime anchor, int freq, int basis)
    {
        if (basis != 1) return Days(start, end, basis) / (DaysInYear(start, end, basis) / freq);
        int step = 12 / freq; bool eom = LastDay(anchor);
        int j = FloorIndex(anchor, step, eom, start);
        double sum = 0;
        for (var q = Shift(anchor, j * step, eom); q < end; j++)
        {
            var next = Shift(anchor, (j + 1) * step, eom);
            var from = start > q ? start : q; var to = end < next ? end : next;
            sum += Act(from, to) / Act(q, next);
            q = next;
        }
        return sum;
    }

    // ------------------------------------------------------------------ coupon bond price, yield, duration

    private static (double N, double A, double E, double Dsc) Factors(DateTime settle, DateTime mat, int freq, int basis)
    {
        var (pcd, ncd, n) = Coupons(settle, mat, freq);
        double e = CoupDays(pcd, ncd, freq, basis), a = Days(pcd, settle, basis);
        return (n, a, e, e - a);
    }

    internal static double Price(DateTime settle, DateTime mat, double rate, double yld, double red, int freq, int basis)
    {
        var (n, a, e, dsc) = Factors(settle, mat, freq, basis);
        double cpn = 100 * rate / freq;
        if (n == 1) return (red + cpn) / (1 + dsc / e * yld / freq) - cpn * a / e;
        double x = 1 + yld / freq, t = dsc / e, sum = 0;
        for (int k = 1; k <= n; k++) sum += cpn / Math.Pow(x, k - 1 + t);
        return red / Math.Pow(x, n - 1 + t) + sum - cpn * a / e;
    }

    internal static double Yield(DateTime settle, DateTime mat, double rate, double pr, double red, int freq, int basis)
    {
        var (n, a, e, dsc) = Factors(settle, mat, freq, basis);
        if (n == 1)
        {
            double p = pr / 100 + a / e * rate / freq;
            return (red / 100 + rate / freq - p) / p * freq * e / dsc;
        }
        return Solve(y => Price(settle, mat, rate, y, red, freq, basis) - pr, 0.05, -freq);
    }

    /// <summary>Macaulay duration in years.</summary>
    internal static double Duration(DateTime settle, DateTime mat, double cpn, double yld, int freq, int basis)
    {
        var (n, _, e, dsc) = Factors(settle, mat, freq, basis);
        double x = 1 + yld / freq, t0 = dsc / e, c = 100 * cpn / freq, num = 0, den = 0;
        for (int k = 1; k <= n; k++)
        {
            double t = k - 1 + t0, cf = (k == n ? c + 100 : c) / Math.Pow(x, t);
            num += t * cf; den += cf;
        }
        return num / den / freq;
    }

    /// <summary>Year basis B and day counts issue→maturity, issue→settlement, settlement→maturity (PRICEMAT, YIELDMAT).</summary>
    internal static (double B, double Dim, double A, double Dsm) MatFactors(DateTime settle, DateTime mat, DateTime issue, int basis)
        => (DaysInYear(issue, settle, basis), Days(issue, mat, basis), Days(issue, settle, basis), Days(settle, mat, basis));

    // ------------------------------------------------------------------ odd periods

    internal sealed class OddFirstBond
    {
        public int Freq;
        public bool Long;
        public int N;               // coupons after the first coupon
        public double E, Dfc, A, Dsc, Nq, DcNl, ANl;

        public double Price(double rate, double yld, double red)
        {
            double cpn = 100 * rate / Freq, x = 1 + yld / Freq, t = Dsc / E, sum = 0;
            if (!Long)
            {
                for (int k = 2; k <= N + 1; k++) sum += cpn / Math.Pow(x, k - 1 + t);
                return red / Math.Pow(x, N + t) + cpn * Dfc / E / Math.Pow(x, t) + sum - cpn * A / E;
            }
            for (int k = 1; k <= N; k++) sum += cpn / Math.Pow(x, k + Nq + t);
            return red / Math.Pow(x, N + Nq + t) + cpn * DcNl / Math.Pow(x, Nq + t) + sum - cpn * ANl;
        }
    }

    internal static OddFirstBond OddFirst(double settlement, double maturity, double issue, double firstCoupon, double frequency, double basis)
    {
        var s = Date(settlement); var m = Date(maturity); var iss = Date(issue); var fc = Date(firstCoupon);
        int f = Freq(frequency), b = Basis(basis);
        if (!(m > fc && fc > s && s > iss)) throw NumErr();
        int step = 12 / f; bool eom = LastDay(fc);
        double NormalLength(DateTime q0, DateTime q1) => CoupDays(q0, q1, f, b);
        var o = new OddFirstBond { Freq = f, N = Coupons(fc, m, f).Num };
        int ji = FloorIndex(fc, step, eom, iss);           // quasi-coupon period holding the issue date
        if (ji >= -1)
        {
            var q = Shift(fc, -step, eom);
            o.E = NormalLength(q, fc);
            o.Dfc = Days(iss, fc, b); o.A = Days(iss, s, b); o.Dsc = Days(s, fc, b);
            return o;
        }
        o.Long = true;
        for (int j = ji; j < 0; j++)
        {
            var q0 = Shift(fc, j * step, eom); var q1 = Shift(fc, (j + 1) * step, eom);
            double nl = NormalLength(q0, q1);
            var from = iss > q0 ? iss : q0;
            o.DcNl += Days(from, q1, b) / nl;
            if (s > from) o.ANl += Days(from, s < q1 ? s : q1, b) / nl;
        }
        int js = FloorIndex(fc, step, eom, s);
        var p0 = Shift(fc, js * step, eom); var p1 = Shift(fc, (js + 1) * step, eom);
        o.E = NormalLength(p0, p1);
        o.Nq = -(js + 1);
        o.Dsc = b is 2 or 3 ? Act(s, p1) : o.E - Days(p0, s, b);
        return o;
    }

    internal static (int F, double DC, double A, double DSC) OddLast(double settlement, double maturity, double lastInterest, double frequency, double basis)
    {
        var s = Date(settlement); var m = Date(maturity); var li = Date(lastInterest);
        int f = Freq(frequency), b = Basis(basis);
        if (!(m > s && s > li)) throw NumErr();
        int step = 12 / f; bool eom = LastDay(li);
        double dc = 0, acc = 0, dsc = 0;
        var q0 = li;
        for (int j = 1; q0 < m; j++)
        {
            var q1 = Shift(li, j * step, eom);
            double nl = CoupDays(q0, q1, f, b);
            var end = m < q1 ? m : q1;
            dc += Days(q0, end, b) / nl;
            if (s > q0) acc += Days(q0, s < q1 ? s : q1, b) / nl;
            if (s < q1) dsc += Days(s > q0 ? s : q0, end, b) / nl;
            q0 = q1;
        }
        return (f, dc, acc, dsc);
    }

    // ------------------------------------------------------------------ loans

    internal static double Pmt(double r, double n, double pv, double type)
    {
        double f = Math.Pow(1 + r, n);
        return -r * pv * f / ((1 + r * type) * (f - 1));
    }

    private static double Fv(double r, double n, double pmt, double pv, double type)
    {
        double f = Math.Pow(1 + r, n);
        return -(pv * f + pmt * (1 + r * type) * (f - 1) / r);
    }

    /// <summary>Interest part of payment <paramref name="per"/> (fv = 0).</summary>
    internal static double Ipmt(double r, double per, double pv, double pmt, double type)
    {
        if (per == 1 && type == 1) return 0;
        double ip = Fv(r, per - 1, pmt, pv, type) * r;
        return type == 1 ? ip / (1 + r) : ip;
    }

    // ------------------------------------------------------------------ irregular cash flows

    internal static (double[] Values, double[] Days) Schedule(Ctx c, Expr values, Expr dates)
    {
        static List<double> Read(Ctx c, Expr e)
        {
            var list = new List<double>();
            var v = c.Eval(e);
            if (v is CellValue one) { list.Add(Calc.ToNum(one)); return list; }
            foreach (var x in Functions.FlatVal(c, v))
            {
                if (x.IsError) throw new FormulaError(x.Str!);
                if (!x.IsNumber) throw new FormulaError(Errors.Value);
                list.Add(x.Num);
            }
            return list;
        }
        var vals = Read(c, values);
        var ds = Read(c, dates);
        if (vals.Count != ds.Count || vals.Count == 0) throw NumErr();
        var days = new double[ds.Count];
        for (int i = 0; i < ds.Count; i++)
        {
            if (ds[i] < 0) throw NumErr();
            days[i] = Math.Floor(ds[i]) - Math.Floor(ds[0]);
            if (days[i] < 0) throw NumErr();
        }
        return (vals.ToArray(), days);
    }

    internal static double Xnpv(double r, double[] v, double[] days)
    {
        double s = 0;
        for (int i = 0; i < v.Length; i++) s += v[i] / Math.Pow(1 + r, days[i] / 365);
        return s;
    }

    /// <summary>Root of a function on (lower, ∞): Newton's method from the guess, then a bracketing bisection.</summary>
    internal static double Solve(Func<double, double> f, double guess, double lower)
    {
        double x = guess;
        for (int it = 0; it < 100; it++)
        {
            double fx = f(x);
            if (double.IsNaN(fx) || double.IsInfinity(fx)) break;
            if (fx == 0) return x;
            double h = 1e-6 * Math.Max(1e-3, Math.Abs(x));
            if (x - h <= lower) break;
            double d = (f(x + h) - f(x - h)) / (2 * h);
            if (d == 0 || double.IsNaN(d) || double.IsInfinity(d)) break;
            double nx = x - fx / d;
            if (nx <= lower) nx = (x + lower) / 2;
            if (Math.Abs(nx - x) <= 1e-13 * Math.Max(1, Math.Abs(x))) return nx;
            x = nx;
        }
        // bracket the root around the guess, then bisect
        double eps = 1e-10 * Math.Max(1, Math.Abs(lower));
        double lo = Math.Max(guess, lower + eps), hi = lo, flo = f(lo), fhi = flo, width = 0.05;
        for (int k = 0; k < 200 && !(flo * fhi <= 0 && !double.IsNaN(flo * fhi)); k++)
        {
            lo = Math.Max(lower + eps, lo - width); hi += width; width *= 1.6;
            flo = f(lo); fhi = f(hi);
            if (double.IsInfinity(hi)) break;
        }
        if (!(flo * fhi <= 0)) throw NumErr();
        for (int it = 0; it < 300; it++)
        {
            double mid = (lo + hi) / 2, fm = f(mid);
            if (fm == 0 || hi - lo <= 1e-15 * Math.Max(1, Math.Abs(mid))) return mid;
            if ((fm < 0) == (flo < 0)) { lo = mid; flo = fm; } else hi = mid;
        }
        return (lo + hi) / 2;
    }

    // ------------------------------------------------------------------ depreciation

    /// <summary>Double-declining balance depreciation of one period (never below salvage).</summary>
    private static double Ddb(double cost, double salv, double life, double per, double factor)
    {
        double rate = factor / life, oldValue;
        if (rate >= 1) { rate = 1; oldValue = per == 1 ? cost : 0; }
        else oldValue = cost * Math.Pow(1 - rate, per - 1);
        double newValue = cost * Math.Pow(1 - rate, per);
        double d = newValue < salv ? oldValue - salv : oldValue - newValue;
        return d < 0 ? 0 : d;
    }

    /// <summary>Declining balance switching to straight line, from period 0 to <paramref name="period"/>.</summary>
    private static double InterVdb(double cost, double salv, double life, double life1, double period, double factor)
    {
        double vdb = 0, intEnd = Math.Ceiling(period), remaining = cost - salv, sln = 0;
        bool nowSln = false;
        for (int i = 1; i <= intEnd; i++)
        {
            double term;
            if (!nowSln)
            {
                double ddb = Ddb(cost, salv, life, i, factor);
                sln = remaining / (life1 - (i - 1));
                if (sln > ddb) { term = sln; nowSln = true; }
                else { term = ddb; remaining -= ddb; }
            }
            else term = sln;
            if (i == intEnd) term *= period + 1 - intEnd;
            vdb += term;
        }
        return vdb;
    }

    internal static double Vdb(double cost, double salv, double life, double start, double end, double factor, bool noSwitch)
    {
        double intStart = Math.Floor(start), intEnd = Math.Ceiling(end);
        if (noSwitch)
        {
            double vdb = 0;
            for (double i = intStart + 1; i <= intEnd; i++)
            {
                double term = Ddb(cost, salv, life, i, factor);
                if (i == intStart + 1) term *= Math.Min(end, intStart + 1) - start;
                else if (i == intEnd) term *= end + 1 - intEnd;
                vdb += term;
            }
            return vdb;
        }
        double part = 0;
        if (start != intStart)
        {
            double value = cost - InterVdb(cost, salv, life, life, intStart, factor);
            part += (start - intStart) * InterVdb(value, salv, life, life - intStart, 1, factor);
        }
        if (end != intEnd)
        {
            double tStart = intEnd - 1;
            double value = cost - InterVdb(cost, salv, life, life, tStart, factor);
            part += (intEnd - end) * InterVdb(value, salv, life, life - tStart, 1, factor);
        }
        double rest = cost - InterVdb(cost, salv, life, life, intStart, factor);
        return InterVdb(rest, salv, life, life - intStart, intEnd - intStart, factor) - part;
    }

    private static double Round0(double x) => Math.Round(x, MidpointRounding.AwayFromZero);

    /// <summary>French degressive depreciation (AMORDEGRC); <paramref name="firstFrac"/> is the year fraction of the first period.</summary>
    internal static double Amordegrc(double cost, double firstFrac, double salv, int per, double rate)
    {
        double life = 1 / rate, coeff;
        if ((life > 0 && life < 1) || (life > 1 && life < 2) || (life > 2 && life < 3) || (life > 4 && life < 5)) throw NumErr();
        if (life >= 3 && life <= 4) coeff = 1.5;
        else if (life >= 5 && life <= 6) coeff = 2;
        else if (life > 6) coeff = 2.5;
        else coeff = 1;
        rate *= coeff;
        double dep = Round0(firstFrac * rate * cost);
        cost -= dep;
        double rest = cost - salv;
        for (int n = 0; n < per; n++)
        {
            dep = Round0(rate * cost);
            rest -= dep;
            if (rest < 0) return per - n <= 1 ? Round0(cost * 0.5) : 0;
            cost -= dep;
        }
        return dep;
    }

    /// <summary>French linear depreciation (AMORLINC).</summary>
    internal static double Amorlinc(double cost, double firstFrac, double salv, int per, double rate)
    {
        double full = cost * rate, first = firstFrac * rate * cost;
        double fullPeriods = Math.Floor((cost - salv - first) / full);
        if (per == 0) return first;
        if (per <= fullPeriods) return full;
        if (per == fullPeriods + 1) return cost - salv - full * fullPeriods - first;
        return 0;
    }
}
