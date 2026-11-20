using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

public static partial class Functions
{
    /// <summary>Probability distributions and hypothesis tests, plus their Excel 2007 compatibility names.</summary>
    internal static void RegisterDist()
    {
        // ------------------------------------------------------------------ normal family

        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double x = c.Num(a[0]), m = c.Num(a[1]), s = c.Num(a[2]); bool cum = c.Bool(a[3]);
            if (s <= 0) throw Err(Errors.Num);
            double z = (x - m) / s;
            return DistMath.R(cum ? DistMath.NormCdf(z) : DistMath.NormPdf(z) / s);
        }, "NORM.DIST", "NORMDIST");
        Reg((c, a) => { Args(a, 2, 2); double z = c.Num(a[0]); return DistMath.R(c.Bool(a[1]) ? DistMath.NormCdf(z) : DistMath.NormPdf(z)); }, "NORM.S.DIST");
        Reg((c, a) => { Args(a, 1, 1); return DistMath.R(DistMath.NormCdf(c.Num(a[0]))); }, "NORMSDIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double p = c.Num(a[0]), m = c.Num(a[1]), s = c.Num(a[2]);
            if (p <= 0 || p >= 1 || s <= 0) throw Err(Errors.Num);
            return DistMath.R(m + s * DistMath.NormInv(p));
        }, "NORM.INV", "NORMINV");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            double p = c.Num(a[0]);
            if (p <= 0 || p >= 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.NormInv(p));
        }, "NORM.S.INV", "NORMSINV");
        Reg((c, a) => { Args(a, 1, 1); return DistMath.R(DistMath.NormPdf(c.Num(a[0]))); }, "PHI");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            double z = c.Num(a[0]);
            return DistMath.R(z == 0 ? 0 : Math.Sign(z) * 0.5 * DistMath.GammaReg(0.5, 0.5 * z * z, false));
        }, "GAUSS");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), m = c.Num(a[1]), s = c.Num(a[2]);
            if (s <= 0) throw Err(Errors.Num);
            return DistMath.R((x - m) / s);
        }, "STANDARDIZE");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double x = c.Num(a[0]), m = c.Num(a[1]), s = c.Num(a[2]); bool cum = c.Bool(a[3]);
            if (x <= 0 || s <= 0) throw Err(Errors.Num);
            double z = (Math.Log(x) - m) / s;
            return DistMath.R(cum ? DistMath.NormCdf(z) : DistMath.NormPdf(z) / (s * x));
        }, "LOGNORM.DIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), m = c.Num(a[1]), s = c.Num(a[2]);
            if (x <= 0 || s <= 0) throw Err(Errors.Num);
            return DistMath.R(DistMath.NormCdf((Math.Log(x) - m) / s));
        }, "LOGNORMDIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double p = c.Num(a[0]), m = c.Num(a[1]), s = c.Num(a[2]);
            if (p <= 0 || p >= 1 || s <= 0) throw Err(Errors.Num);
            return DistMath.R(Math.Exp(m + s * DistMath.NormInv(p)));
        }, "LOGNORM.INV", "LOGINV");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double alpha = c.Num(a[0]), sd = c.Num(a[1]), n = DistMath.Tr(c, a[2]);
            if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) throw Err(Errors.Num);
            return DistMath.R(-DistMath.NormInv(alpha / 2) * sd / Math.Sqrt(n));
        }, "CONFIDENCE.NORM", "CONFIDENCE");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double alpha = c.Num(a[0]), sd = c.Num(a[1]), n = DistMath.Tr(c, a[2]);
            if (alpha <= 0 || alpha >= 1 || sd <= 0 || n < 1) throw Err(Errors.Num);
            if (n == 1) throw Err(Errors.Div0);
            return DistMath.R(DistMath.TInvTail(alpha / 2, n - 1) * sd / Math.Sqrt(n));
        }, "CONFIDENCE.T");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var l = Nums(c, [a[0]]);
            double x = c.Num(a[1]);
            int n = l.Count;
            if (n == 0) throw Err(Errors.NA);
            double mean = l.Average(), sigma;
            if (a.Length > 2 && a[2] is not MissingExpr)
            {
                sigma = c.Num(a[2]);
                if (sigma < 0) throw Err(Errors.Num);
            }
            else
            {
                if (n < 2) throw Err(Errors.Div0);
                sigma = Math.Sqrt(l.Sum(v => (v - mean) * (v - mean)) / (n - 1));
            }
            if (sigma == 0) throw Err(Errors.Div0);
            return DistMath.R(DistMath.NormCdf(-(mean - x) / (sigma / Math.Sqrt(n))));
        }, "Z.TEST", "ZTEST");

        // ------------------------------------------------------------------ discrete distributions

        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double k = DistMath.Tr(c, a[0]), n = DistMath.Tr(c, a[1]), p = c.Num(a[2]); bool cum = c.Bool(a[3]);
            if (k < 0 || k > n || p < 0 || p > 1) throw Err(Errors.Num);
            return DistMath.R(cum ? DistMath.BinomCdf(k, n, p, false) : DistMath.BinomPmf(k, n, p));
        }, "BINOM.DIST", "BINOMDIST");
        Reg((c, a) =>
        {
            Args(a, 3, 4);
            double n = DistMath.Tr(c, a[0]), p = c.Num(a[1]), s1 = DistMath.Tr(c, a[2]);
            double s2 = a.Length > 3 && a[3] is not MissingExpr ? DistMath.Tr(c, a[3]) : s1;
            if (n < 0 || p < 0 || p > 1 || s1 < 0 || s1 > n || s2 < s1 || s2 > n) throw Err(Errors.Num);
            return DistMath.R(DistMath.BinomRange(n, p, s1, s2));
        }, "BINOM.DIST.RANGE");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double n = DistMath.Tr(c, a[0]), p = c.Num(a[1]), alpha = c.Num(a[2]);
            if (n < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.BinomInv(n, p, alpha));
        }, "BINOM.INV", "CRITBINOM");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double f = DistMath.Tr(c, a[0]), s = DistMath.Tr(c, a[1]), p = c.Num(a[2]); bool cum = c.Bool(a[3]);
            if (f < 0 || s < 1 || p < 0 || p > 1) throw Err(Errors.Num);
            return DistMath.R(cum ? DistMath.BetaReg(p, 1 - p, s, f + 1, false) : DistMath.NegBinomPmf(f, s, p));
        }, "NEGBINOM.DIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double f = DistMath.Tr(c, a[0]), s = DistMath.Tr(c, a[1]), p = c.Num(a[2]);
            if (f < 0 || s < 1 || p < 0 || p > 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.NegBinomPmf(f, s, p));
        }, "NEGBINOMDIST");
        FnImpl hyp = (c, a) =>
        {
            bool modern = a.Length == 5;
            double x = DistMath.Tr(c, a[0]), n = DistMath.Tr(c, a[1]), k = DistMath.Tr(c, a[2]), pop = DistMath.Tr(c, a[3]);
            bool cum = modern && c.Bool(a[4]);
            if (x < 0 || x > Math.Min(n, k) || x < Math.Max(0, n - pop + k) || n <= 0 || n > pop || k <= 0 || k > pop || pop <= 0)
                throw Err(Errors.Num);
            return DistMath.R(cum ? DistMath.HypGeomCdf(x, n, k, pop) : DistMath.HypGeomPmf(x, n, k, pop));
        };
        Reg((c, a) => { Args(a, 5, 5); return hyp(c, a); }, "HYPGEOM.DIST");
        Reg((c, a) => { Args(a, 4, 4); return hyp(c, a); }, "HYPGEOMDIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = DistMath.Tr(c, a[0]), mean = c.Num(a[1]); bool cum = c.Bool(a[2]);
            if (x < 0 || mean < 0) throw Err(Errors.Num);
            return DistMath.R(cum ? (mean == 0 ? 1 : DistMath.GammaReg(x + 1, mean, true)) : DistMath.DPoisRaw(x, mean));
        }, "POISSON.DIST", "POISSON");

        // ------------------------------------------------------------------ continuous distributions

        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), l = c.Num(a[1]); bool cum = c.Bool(a[2]);
            if (x < 0 || l <= 0) throw Err(Errors.Num);
            return DistMath.R(cum ? -DistMath.Expm1(-l * x) : l * Math.Exp(-l * x));
        }, "EXPON.DIST", "EXPONDIST");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double x = c.Num(a[0]), k = c.Num(a[1]), l = c.Num(a[2]); bool cum = c.Bool(a[3]);
            if (x < 0 || k <= 0 || l <= 0) throw Err(Errors.Num);
            double u = Math.Pow(x / l, k);
            return DistMath.R(cum ? -DistMath.Expm1(-u) : k / l * Math.Pow(x / l, k - 1) * Math.Exp(-u));
        }, "WEIBULL.DIST", "WEIBULL");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double x = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]); bool cum = c.Bool(a[3]);
            if (x < 0 || al <= 0 || be <= 0) throw Err(Errors.Num);
            return DistMath.R(cum ? DistMath.GammaReg(al, x / be, false) : DistMath.GammaPdf(x / be, al) / be);
        }, "GAMMA.DIST", "GAMMADIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double p = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]);
            if (p < 0 || p >= 1 || al <= 0 || be <= 0) throw Err(Errors.Num);
            return DistMath.R(be * DistMath.GammaInv(al, p, 1 - p));
        }, "GAMMA.INV", "GAMMAINV");
        Reg((c, a) =>
        {
            Args(a, 4, 6);
            double x = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]); bool cum = c.Bool(a[3]);
            double lo = c.NumOr(a, 4, 0), hi = c.NumOr(a, 5, 1);
            if (al <= 0 || be <= 0 || x < lo || x > hi || lo == hi) throw Err(Errors.Num);
            double w = hi - lo, z = (x - lo) / w, y = (hi - x) / w;
            return DistMath.R(cum ? DistMath.BetaReg(z, y, al, be, false) : DistMath.BetaPdf(z, y, al, be) / w);
        }, "BETA.DIST");
        Reg((c, a) =>
        {
            Args(a, 3, 5);
            double x = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]);
            double lo = c.NumOr(a, 3, 0), hi = c.NumOr(a, 4, 1);
            if (al <= 0 || be <= 0 || x < lo || x > hi || lo == hi) throw Err(Errors.Num);
            double w = hi - lo;
            return DistMath.R(DistMath.BetaReg((x - lo) / w, (hi - x) / w, al, be, false));
        }, "BETADIST");
        Reg((c, a) =>
        {
            Args(a, 3, 5);
            double p = c.Num(a[0]), al = c.Num(a[1]), be = c.Num(a[2]);
            double lo = c.NumOr(a, 3, 0), hi = c.NumOr(a, 4, 1);
            if (p <= 0 || p > 1 || al <= 0 || be <= 0 || lo >= hi) throw Err(Errors.Num);
            return DistMath.R(lo + (hi - lo) * DistMath.BetaInv(al, be, p, 1 - p));
        }, "BETA.INV", "BETAINV");

        // chi-squared
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), df = DistMath.Tr(c, a[1]); bool cum = c.Bool(a[2]);
            if (x < 0 || df < 1 || df > 1e10) throw Err(Errors.Num);
            return DistMath.R(cum ? DistMath.GammaReg(df / 2, x / 2, false) : DistMath.GammaPdf(x / 2, df / 2) / 2);
        }, "CHISQ.DIST");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double x = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (x < 0 || df < 1 || df > 1e10) throw Err(Errors.Num);
            return DistMath.R(DistMath.GammaReg(df / 2, x / 2, true));
        }, "CHISQ.DIST.RT", "CHIDIST");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double p = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (p < 0 || p >= 1 || df < 1 || df > 1e10) throw Err(Errors.Num);
            return DistMath.R(2 * DistMath.GammaInv(df / 2, p, 1 - p));
        }, "CHISQ.INV");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double q = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (q <= 0 || q > 1 || df < 1 || df > 1e10) throw Err(Errors.Num);
            return DistMath.R(2 * DistMath.GammaInv(df / 2, 1 - q, q));
        }, "CHISQ.INV.RT", "CHIINV");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var act = Grid(c, a[0]); var exp = Grid(c, a[1]);
            int rows = act.GetLength(0), cols = act.GetLength(1);
            if (exp.GetLength(0) != rows || exp.GetLength(1) != cols) throw Err(Errors.NA);
            double chi = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    CellValue o = act[i, j], e = exp[i, j];
                    if (o.IsError) throw Err(o.Str!);
                    if (e.IsError) throw Err(e.Str!);
                    if (!o.IsNumber || !e.IsNumber) continue;
                    if (e.Num == 0) throw Err(Errors.Div0);
                    chi += (o.Num - e.Num) * (o.Num - e.Num) / e.Num;
                }
            double df = rows > 1 && cols > 1 ? (rows - 1.0) * (cols - 1.0) : rows * cols - 1.0;
            if (df < 1) throw Err(Errors.NA);
            if (chi < 0) throw Err(Errors.Num);
            return DistMath.R(DistMath.GammaReg(df / 2, chi / 2, true));
        }, "CHISQ.TEST", "CHITEST");

        // F
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double x = c.Num(a[0]), d1 = DistMath.Tr(c, a[1]), d2 = DistMath.Tr(c, a[2]); bool cum = c.Bool(a[3]);
            if (x < 0 || d1 < 1 || d2 < 1) throw Err(Errors.Num);
            return DistMath.R(cum ? DistMath.FCdf(x, d1, d2, false) : DistMath.FPdf(x, d1, d2));
        }, "F.DIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), d1 = DistMath.Tr(c, a[1]), d2 = DistMath.Tr(c, a[2]);
            if (x < 0 || d1 < 1 || d2 < 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.FCdf(x, d1, d2, true));
        }, "F.DIST.RT");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), d1 = DistMath.Tr(c, a[1]), d2 = DistMath.Tr(c, a[2]);
            if (x < 0 || d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10) throw Err(Errors.Num);
            return DistMath.R(DistMath.FCdf(x, d1, d2, true));
        }, "FDIST");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double p = c.Num(a[0]), d1 = DistMath.Tr(c, a[1]), d2 = DistMath.Tr(c, a[2]);
            if (p < 0 || p >= 1 || d1 < 1 || d2 < 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.FInv(d1, d2, p, 1 - p));
        }, "F.INV");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double q = c.Num(a[0]), d1 = DistMath.Tr(c, a[1]), d2 = DistMath.Tr(c, a[2]);
            if (q <= 0 || q > 1 || d1 < 1 || d2 < 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.FInv(d1, d2, 1 - q, q));
        }, "F.INV.RT");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double q = c.Num(a[0]), d1 = DistMath.Tr(c, a[1]), d2 = DistMath.Tr(c, a[2]);
            if (q <= 0 || q > 1 || d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10) throw Err(Errors.Num);
            return DistMath.R(DistMath.FInv(d1, d2, 1 - q, q));
        }, "FINV");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var x = Nums(c, [a[0]]); var y = Nums(c, [a[1]]);
            if (x.Count < 2 || y.Count < 2) throw Err(Errors.Div0);
            double v1 = DistMath.SampleVar(x), v2 = DistMath.SampleVar(y);
            if (v1 == 0 || v2 == 0) throw Err(Errors.Div0);
            double f = v1 / v2, d1 = x.Count - 1, d2 = y.Count - 1;
            double p = 2 * Math.Min(DistMath.FCdf(f, d1, d2, false), DistMath.FCdf(f, d1, d2, true));
            return DistMath.R(Math.Min(1, p));
        }, "F.TEST", "FTEST");

        // Student t
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), df = DistMath.Tr(c, a[1]); bool cum = c.Bool(a[2]);
            if (df < 1) throw Err(Errors.Num);
            if (!cum) return DistMath.R(DistMath.TPdf(x, df));
            return DistMath.R(x < 0 ? DistMath.TTail(-x, df) : 1 - DistMath.TTail(x, df));
        }, "T.DIST");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double x = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (df < 1) throw Err(Errors.Num);
            return DistMath.R(x < 0 ? 1 - DistMath.TTail(-x, df) : DistMath.TTail(x, df));
        }, "T.DIST.RT");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double x = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (x < 0 || df < 1) throw Err(Errors.Num);
            return DistMath.R(Math.Min(1, 2 * DistMath.TTail(x, df)));
        }, "T.DIST.2T");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x = c.Num(a[0]), df = DistMath.Tr(c, a[1]), tails = DistMath.Tr(c, a[2]);
            if (x < 0 || df < 1 || (tails != 1 && tails != 2)) throw Err(Errors.Num);
            return DistMath.R(Math.Min(1, tails * DistMath.TTail(x, df)));
        }, "TDIST");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double p = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (p <= 0 || p >= 1 || df < 1) throw Err(Errors.Num);
            return DistMath.R(p < 0.5 ? -DistMath.TInvTail(p, df) : DistMath.TInvTail(1 - p, df));
        }, "T.INV");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double p = c.Num(a[0]), df = DistMath.Tr(c, a[1]);
            if (p <= 0 || p > 1 || df < 1) throw Err(Errors.Num);
            return DistMath.R(DistMath.TInvTail(p / 2, df));
        }, "T.INV.2T", "TINV");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double tails = DistMath.Tr(c, a[2]), type = DistMath.Tr(c, a[3]);
            if ((tails != 1 && tails != 2) || type < 1 || type > 3) throw Err(Errors.Num);
            double t, df;
            if (type == 1)
            {
                var g1 = Grid(c, a[0]); var g2 = Grid(c, a[1]);
                if (g1.Length != g2.Length) throw Err(Errors.NA);
                var v1 = g1.Cast<CellValue>().ToArray(); var v2 = g2.Cast<CellValue>().ToArray();
                var d = new List<double>();
                for (int i = 0; i < v1.Length; i++)
                {
                    if (v1[i].IsError) throw Err(v1[i].Str!);
                    if (v2[i].IsError) throw Err(v2[i].Str!);
                    if (v1[i].IsNumber && v2[i].IsNumber) d.Add(v1[i].Num - v2[i].Num);
                }
                if (d.Count < 2) throw Err(Errors.Div0);
                double var = DistMath.SampleVar(d);
                if (var == 0) throw Err(Errors.Div0);
                t = d.Average() / Math.Sqrt(var / d.Count);
                df = d.Count - 1;
            }
            else
            {
                var x = Nums(c, [a[0]]); var y = Nums(c, [a[1]]);
                if (x.Count < 2 || y.Count < 2) throw Err(Errors.Div0);
                double n1 = x.Count, n2 = y.Count, v1 = DistMath.SampleVar(x), v2 = DistMath.SampleVar(y), diff = x.Average() - y.Average();
                if (type == 2)
                {
                    df = n1 + n2 - 2;
                    double pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
                    if (pooled == 0) throw Err(Errors.Div0);
                    t = diff / Math.Sqrt(pooled * (1 / n1 + 1 / n2));
                }
                else
                {
                    double s1 = v1 / n1, s2 = v2 / n2;
                    if (s1 + s2 == 0) throw Err(Errors.Div0);
                    t = diff / Math.Sqrt(s1 + s2);
                    df = (s1 + s2) * (s1 + s2) / (s1 * s1 / (n1 - 1) + s2 * s2 / (n2 - 1));
                }
            }
            return DistMath.R(Math.Min(1, tails * DistMath.TTail(Math.Abs(t), df)));
        }, "T.TEST", "TTEST");

        // ------------------------------------------------------------------ gamma function, Fisher, misc

        Reg((c, a) =>
        {
            Args(a, 1, 1);
            double x = c.Num(a[0]);
            if (x <= 0 && x == Math.Floor(x)) throw Err(Errors.Num);
            return DistMath.R(DistMath.Gamma(x));
        }, "GAMMA");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            double x = c.Num(a[0]);
            if (x <= 0) throw Err(Errors.Num);
            return DistMath.R(DistMath.LogGamma(x));
        }, "GAMMALN", "GAMMALN.PRECISE");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            double x = c.Num(a[0]);
            if (x <= -1 || x >= 1) throw Err(Errors.Num);
            return DistMath.R(Math.Atanh(x));
        }, "FISHER");
        Reg((c, a) => { Args(a, 1, 1); return DistMath.R(Math.Tanh(c.Num(a[0]))); }, "FISHERINV");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double n = DistMath.Tr(c, a[0]), k = DistMath.Tr(c, a[1]);
            if (n < 0 || k < 0 || (n == 0 && k > 0)) throw Err(Errors.Num);
            return DistMath.R(Math.Pow(n, k));
        }, "PERMUTATIONA");
        Reg((c, a) =>
        {
            Args(a, 3, 4);
            var xs = Grid(c, a[0]); var ps = Grid(c, a[1]);
            if (xs.Length != ps.Length) throw Err(Errors.NA);
            double lo = c.Num(a[2]);
            double hi = a.Length > 3 && a[3] is not MissingExpr ? c.Num(a[3]) : lo;
            if (lo > hi) (lo, hi) = (hi, lo);
            var xv = xs.Cast<CellValue>().ToArray(); var pv = ps.Cast<CellValue>().ToArray();
            double sum = 0, res = 0;
            for (int i = 0; i < xv.Length; i++)
            {
                if (xv[i].IsError) throw Err(xv[i].Str!);
                if (pv[i].IsError) throw Err(pv[i].Str!);
                if (!pv[i].IsNumber) continue;
                double p = pv[i].Num;
                if (p <= 0 || p > 1) throw Err(Errors.Num);
                sum += p;
                if (xv[i].IsNumber && xv[i].Num >= lo && xv[i].Num <= hi) res += p;
            }
            if (Math.Abs(sum - 1) > 1e-7) throw Err(Errors.Num);
            return DistMath.R(res);
        }, "PROB");
    }
}

/// <summary>
/// Special functions and distribution kernels for <see cref="Functions.RegisterDist"/>. Density and incomplete-function
/// prefactors use Loader's saddle-point expansion (stirlerr/bd0, as in R's nmath), so results keep near full double
/// precision even for large parameters.
/// </summary>
internal static class DistMath
{
    private const double LnSqrt2Pi = 0.91893853320467274178032973640562; // ln(sqrt(2*pi))
    private const double Ln2Pi = 1.8378770664093454835606594728112;
    private const double Sqrt2Pi = 2.5066282746310005024157652848110;

    // ------------------------------------------------------------------ argument helpers

    internal static CellValue R(double v) => double.IsFinite(v) ? CellValue.Number(v) : throw new FormulaError(Errors.Num);

    /// <summary>Argument truncated toward zero (Excel truncates counts and degrees of freedom).</summary>
    internal static double Tr(Ctx c, Expr e) => Math.Truncate(c.Num(e));

    internal static double SampleVar(List<double> l)
    {
        double m = l.Average(), s = 0;
        foreach (var v in l) s += (v - m) * (v - m);
        return s / (l.Count - 1);
    }

    // ------------------------------------------------------------------ elementary helpers

    internal static double Log1p(double x)
    {
        double u = 1 + x;
        return u == 1 ? x : Math.Log(u) * x / (u - 1);
    }

    internal static double Expm1(double x)
    {
        if (Math.Abs(x) > 0.5) return Math.Exp(x) - 1;
        double u = Math.Exp(x);
        if (u == 1) return x;
        double um1 = u - 1;
        return um1 == -1 ? -1 : um1 * x / Math.Log(u);
    }

    private static double SinPi(double x)
    {
        x -= 2 * Math.Round(x / 2); // reduce to [-1, 1]
        return Math.Sin(Math.PI * x);
    }

    // ------------------------------------------------------------------ gamma function

    private static readonly double[] Lanczos =
    [
        0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
        12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7
    ];

    private static double LanczosSum(double x) // x = z - 1
    {
        double s = Lanczos[0];
        for (int i = 1; i < Lanczos.Length; i++) s += Lanczos[i] / (x + i);
        return s;
    }

    internal static double Gamma(double x)
    {
        if (x == Math.Floor(x))
        {
            if (x <= 0) return double.NaN;
            if (x > 171) return double.PositiveInfinity;
            double r = 1;
            for (int i = 2; i < x; i++) r *= i;
            return r;
        }
        if (x < 0.5) return Math.PI / (SinPi(x) * Gamma(1 - x));
        if (x > 171.7) return double.PositiveInfinity;
        double z = x - 1, t = z + 7.5, tp = Math.Pow(t, 0.5 * (z + 0.5));
        return Sqrt2Pi * tp * (tp * Math.Exp(-t)) * LanczosSum(z);
    }

    /// <summary>ln Γ(x) for x &gt; 0.</summary>
    internal static double LogGamma(double x)
    {
        if (x == 1 || x == 2) return 0;
        if (x < 0.5) return Math.Log(Math.PI / Math.Abs(SinPi(x))) - LogGamma(1 - x);
        if (x > 1e7)
        {
            double r = 1 / (x * x);
            return (x - 0.5) * Math.Log(x) - x + LnSqrt2Pi + (1.0 / 12 - r * (1.0 / 360 - r / 1260)) / x;
        }
        double z = x - 1, t = z + 7.5;
        return LnSqrt2Pi + (z + 0.5) * Math.Log(t) - t + Math.Log(LanczosSum(z));
    }

    private static double LogBeta(double a, double b) => LogGamma(a) + LogGamma(b) - LogGamma(a + b);

    /// <summary>Stirling-formula error ln Γ(n+1) − (n+½)ln n + n − ln√(2π).</summary>
    private static double StirlErr(double n)
    {
        const double S0 = 1.0 / 12, S1 = 1.0 / 360, S2 = 1.0 / 1260, S3 = 1.0 / 1680, S4 = 1.0 / 1188;
        if (n <= 15) return LogGamma(n + 1) - (n + 0.5) * Math.Log(n) + n - LnSqrt2Pi;
        double nn = n * n;
        if (n > 500) return (S0 - S1 / nn) / n;
        if (n > 80) return (S0 - (S1 - S2 / nn) / nn) / n;
        if (n > 35) return (S0 - (S1 - (S2 - S3 / nn) / nn) / nn) / n;
        return (S0 - (S1 - (S2 - (S3 - S4 / nn) / nn) / nn) / nn) / n;
    }

    /// <summary>Deviance term x·ln(x/np) + np − x, computed without cancellation.</summary>
    private static double Bd0(double x, double np)
    {
        if (Math.Abs(x - np) < 0.1 * (x + np))
        {
            double v = (x - np) / (x + np), s = (x - np) * v, ej = 2 * x * v;
            v *= v;
            for (int j = 1; j < 1000; j++)
            {
                ej *= v;
                double s1 = s + ej / (2 * j + 1);
                if (s1 == s) return s1;
                s = s1;
            }
        }
        return x * Math.Log(x / np) + np - x;
    }

    /// <summary>ln of the binomial density C(n,x) p^x q^(n−x) (x, n need not be integers).</summary>
    private static double DBinomRawLog(double x, double n, double p, double q)
    {
        if (p == 0) return x == 0 ? 0 : double.NegativeInfinity;
        if (q == 0) return x == n ? 0 : double.NegativeInfinity;
        if (x == 0)
        {
            if (n == 0) return 0;
            return p < 0.1 ? -Bd0(n, n * q) - n * p : n * Math.Log(q);
        }
        if (x == n) return q < 0.1 ? -Bd0(n, n * p) - n * q : n * Math.Log(p);
        if (x < 0 || x > n) return double.NegativeInfinity;
        double lc = StirlErr(n) - StirlErr(x) - StirlErr(n - x) - Bd0(x, n * p) - Bd0(n - x, n * q);
        double lf = Ln2Pi + Math.Log(x) + Log1p(-x / n);
        return lc - 0.5 * lf;
    }

    /// <summary>Poisson density e^−λ λ^x / Γ(x+1) (x need not be an integer).</summary>
    internal static double DPoisRaw(double x, double lambda)
    {
        if (lambda == 0) return x == 0 ? 1 : 0;
        if (x == 0) return Math.Exp(-lambda);
        if (x < 0) return 0;
        return Math.Exp(-StirlErr(x) - Bd0(x, lambda)) / Math.Sqrt(2 * Math.PI * x);
    }

    // ------------------------------------------------------------------ incomplete gamma / beta

    /// <summary>Regularized incomplete gamma P(a, x), or Q(a, x) = 1 − P when <paramref name="upper"/>.</summary>
    internal static double GammaReg(double a, double x, bool upper)
    {
        if (double.IsNaN(x)) return double.NaN;
        if (x <= 0) return upper ? 1 : 0;
        if (double.IsPositiveInfinity(x)) return upper ? 0 : 1;
        if (x < a + 1)
        {
            double sum = 1, term = 1;
            for (int n = 1; n < 10_000_000; n++)
            {
                term *= x / (a + n);
                sum += term;
                if (term < sum * 1e-17) break;
            }
            double p = DPoisRaw(a, x) * sum;
            return upper ? 1 - p : p;
        }
        const double Tiny = 1e-300;
        double b = x + 1 - a, cc = 1 / Tiny, d = 1 / b, h = d;
        for (int i = 1; i < 10_000_000; i++)
        {
            double an = -i * (i - a);
            b += 2;
            d = an * d + b; if (Math.Abs(d) < Tiny) d = Tiny;
            cc = b + an / cc; if (Math.Abs(cc) < Tiny) cc = Tiny;
            d = 1 / d;
            double del = d * cc;
            h *= del;
            if (Math.Abs(del - 1) < 3e-16) break;
        }
        double q = a * DPoisRaw(a, x) * h; // x^a e^−x / Γ(a) · continued fraction
        return upper ? q : 1 - q;
    }

    /// <summary>Regularized incomplete beta I_x(a, b) (or 1 − I_x when <paramref name="upper"/>); y = 1 − x supplied by the caller.</summary>
    internal static double BetaReg(double x, double y, double a, double b, bool upper)
    {
        if (double.IsNaN(x) || double.IsNaN(y)) return double.NaN;
        if (x <= 0) return upper ? 1 : 0;
        if (y <= 0) return upper ? 0 : 1;
        if (x > (a + 1) / (a + b + 2))
        {
            (x, y) = (y, x);
            (a, b) = (b, a);
            upper = !upper;
        }
        // x^a y^b / (a·B(a,b)) = dbinom(a; a+b, x) · b/(a+b)
        double front = Math.Exp(DBinomRawLog(a, a + b, x, y)) * b / (a + b);
        double v = front * BetaCf(a, b, x);
        return upper ? 1 - v : v;
    }

    private static double BetaCf(double a, double b, double x)
    {
        const double Tiny = 1e-300;
        double qab = a + b, qap = a + 1, qam = a - 1, c = 1, d = 1 - qab * x / qap;
        if (Math.Abs(d) < Tiny) d = Tiny;
        d = 1 / d;
        double h = d;
        for (double m = 1; m < 10_000_000; m++)
        {
            double m2 = 2 * m;
            double aa = m * (b - m) * x / ((qam + m2) * (a + m2));
            d = 1 + aa * d; if (Math.Abs(d) < Tiny) d = Tiny;
            c = 1 + aa / c; if (Math.Abs(c) < Tiny) c = Tiny;
            d = 1 / d;
            h *= d * c;
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
            d = 1 + aa * d; if (Math.Abs(d) < Tiny) d = Tiny;
            c = 1 + aa / c; if (Math.Abs(c) < Tiny) c = Tiny;
            d = 1 / d;
            double del = d * c;
            h *= del;
            if (Math.Abs(del - 1) < 3e-16) break;
        }
        return h;
    }

    // ------------------------------------------------------------------ densities

    internal static double NormPdf(double z) => Math.Exp(-0.5 * z * z) / Sqrt2Pi;

    /// <summary>Standard normal CDF Φ(z) via the incomplete gamma function (accurate in both tails).</summary>
    internal static double NormCdf(double z)
    {
        if (double.IsNaN(z)) return double.NaN;
        if (z == 0) return 0.5;
        double x = 0.5 * z * z;
        if (x < 1.5)
        {
            double p = 0.5 * GammaReg(0.5, x, false);
            return z < 0 ? 0.5 - p : 0.5 + p;
        }
        double q = 0.5 * GammaReg(0.5, x, true);
        return z < 0 ? q : 1 - q;
    }

    /// <summary>Gamma density with shape a and unit scale.</summary>
    internal static double GammaPdf(double x, double a)
    {
        if (x < 0) return 0;
        if (x == 0) return a < 1 ? double.PositiveInfinity : a > 1 ? 0 : 1;
        if (a < 1) return DPoisRaw(a, x) * a / x;
        return DPoisRaw(a - 1, x);
    }

    internal static double BetaPdf(double x, double y, double a, double b)
    {
        if (x < 0 || y < 0) return 0;
        if (x == 0) return a > 1 ? 0 : a < 1 ? double.PositiveInfinity : b;
        if (y == 0) return b > 1 ? 0 : b < 1 ? double.PositiveInfinity : a;
        if (a <= 2 || b <= 2) return Math.Exp((a - 1) * Math.Log(x) + (b - 1) * Math.Log(y) - LogBeta(a, b));
        return (a + b - 1) * Math.Exp(DBinomRawLog(a - 1, a + b - 2, x, y));
    }

    /// <summary>Upper tail P(T &gt; t) of Student's t for t ≥ 0 (df may be fractional).</summary>
    internal static double TTail(double t, double df)
    {
        if (t <= 0) return 0.5;
        double t2 = t * t;
        if (double.IsInfinity(t2)) return 0;
        double s = df + t2;
        return 0.5 * BetaReg(df / s, t2 / s, df / 2, 0.5, false);
    }

    internal static double TPdf(double x, double n)
    {
        double t = -Bd0(n / 2, (n + 1) / 2) + StirlErr((n + 1) / 2) - StirlErr(n / 2), x2n = x * x / n, lx2n, u;
        double ax = Math.Abs(x);
        bool large = x2n > 1 / 2.220446049250313e-16;
        if (large) { lx2n = Math.Log(ax) - Math.Log(n) / 2; u = n * lx2n; }
        else if (x2n > 0.2) { lx2n = Math.Log(1 + x2n) / 2; u = n * lx2n; }
        else { lx2n = Log1p(x2n) / 2; u = -Bd0(n / 2, (n + x * x) / 2) + x * x / 2; }
        double inv = large ? Math.Sqrt(n) / ax : Math.Exp(-lx2n);
        return Math.Exp(t - u) * inv / Sqrt2Pi;
    }

    internal static double FCdf(double x, double d1, double d2, bool upper)
    {
        if (x <= 0) return upper ? 1 : 0;
        if (double.IsPositiveInfinity(x)) return upper ? 0 : 1;
        double z = d1 * x, s = d2 + z;
        return BetaReg(z / s, d2 / s, d1 / 2, d2 / 2, upper);
    }

    internal static double FPdf(double x, double m, double n)
    {
        if (x < 0) return 0;
        if (x == 0) return m > 2 ? 0 : m == 2 ? 1 : double.PositiveInfinity;
        double f = 1 / (n + x * m), q = n * f, p = x * m * f;
        if (m >= 2) return m * q / 2 * Math.Exp(DBinomRawLog((m - 2) / 2, (m + n - 2) / 2, p, q));
        return m * m * q / (2 * p * (m + n)) * Math.Exp(DBinomRawLog(m / 2, (m + n) / 2, p, q));
    }

    // ------------------------------------------------------------------ discrete distributions

    internal static double BinomPmf(double k, double n, double p) => Math.Exp(DBinomRawLog(k, n, p, 1 - p));

    /// <summary>P(X ≤ k) (or P(X &gt; k) when <paramref name="upper"/>) for X ~ Binomial(n, p).</summary>
    internal static double BinomCdf(double k, double n, double p, bool upper)
    {
        double lower;
        if (k < 0) lower = 0;
        else if (k >= n) lower = 1;
        else if (p == 0) lower = 1;
        else if (p == 1) lower = 0;
        else return BetaReg(1 - p, p, n - k, k + 1, upper);
        return upper ? 1 - lower : lower;
    }

    internal static double BinomRange(double n, double p, double s1, double s2)
    {
        if (s2 - s1 < 1000)
        {
            double sum = 0;
            for (double k = s1; k <= s2; k++) sum += BinomPmf(k, n, p);
            return Math.Min(1, sum);
        }
        if (s1 > n * p) return BinomCdf(s1 - 1, n, p, true) - BinomCdf(s2, n, p, true);
        return BinomCdf(s2, n, p, false) - BinomCdf(s1 - 1, n, p, false);
    }

    /// <summary>Smallest k with P(X ≤ k) ≥ alpha (BINOM.INV / CRITBINOM).</summary>
    internal static double BinomInv(double n, double p, double alpha)
    {
        double lo = 0, hi = n;
        double target = alpha * (1 - 1e-14);
        while (lo < hi)
        {
            double mid = Math.Floor((lo + hi) / 2);
            if (BinomCdf(mid, n, p, false) >= target) hi = mid; else lo = mid + 1;
        }
        return lo;
    }

    internal static double NegBinomPmf(double f, double s, double p) => s / (s + f) * Math.Exp(DBinomRawLog(s, s + f, p, 1 - p));

    internal static double HypGeomPmf(double x, double n, double k, double pop)
    {
        double p = n / pop, q = (pop - n) / pop;
        return Math.Exp(DBinomRawLog(x, k, p, q) + DBinomRawLog(n - x, pop - k, p, q) - DBinomRawLog(n, pop, p, q));
    }

    internal static double HypGeomCdf(double x, double n, double k, double pop)
    {
        double lo = Math.Max(0, n - (pop - k)), hi = Math.Min(n, k);
        if (x >= hi) return 1;
        double px = HypGeomPmf(x, n, k, pop), term = px;
        if (x < n * k / pop)
        {
            double sum = px;
            for (double j = x; j > lo; j--)
            {
                term *= j * (pop - k - n + j) / ((k - j + 1) * (n - j + 1));
                sum += term;
                if (term < 1e-17 * sum) break;
            }
            return Math.Min(1, sum);
        }
        double up = 0;
        for (double j = x; j < hi; j++)
        {
            term *= (k - j) * (n - j) / ((j + 1) * (pop - k - n + j + 1));
            up += term;
            if (term < 1e-17 * up) break;
        }
        return Math.Max(0, 1 - up);
    }

    // ------------------------------------------------------------------ inverses

    /// <summary>Standard normal quantile: Wichura's AS 241 (PPND16) followed by one Halley refinement step.</summary>
    internal static double NormInv(double p)
    {
        double q = p - 0.5, r, x;
        if (Math.Abs(q) <= 0.425)
        {
            r = 0.180625 - q * q;
            x = q * (((((((2509.0809287301226727 * r + 33430.575583588128105) * r + 67265.770927008700853) * r
                        + 45921.953931549871457) * r + 13731.693765509461125) * r + 1971.5909503065514427) * r + 133.14166789178437745) * r
                        + 3.387132872796366608)
                / (((((((5226.495278852854561 * r + 28729.085735721942674) * r + 39307.89580009271061) * r
                        + 21213.794301586595867) * r + 5394.1960214247511077) * r + 687.1870074920579083) * r + 42.313330701600911252) * r + 1);
        }
        else
        {
            r = q < 0 ? p : 1 - p;
            r = Math.Sqrt(-Math.Log(r));
            if (r <= 5)
            {
                r -= 1.6;
                x = (((((((7.7454501427834140764e-4 * r + 0.0227238449892691845833) * r + 0.24178072517745061177) * r
                        + 1.27045825245236838258) * r + 3.64784832476320460504) * r + 5.7694972214606914055) * r + 4.6303378461565452959) * r
                        + 1.42343711074968357734)
                    / (((((((1.05075007164441684324e-9 * r + 5.475938084995344946e-4) * r + 0.0151986665636164571966) * r
                        + 0.14810397642748007459) * r + 0.68976733498510000455) * r + 1.6763848301838038494) * r + 2.05319162663775882187) * r + 1);
            }
            else
            {
                r -= 5;
                x = (((((((2.01033439929228813265e-7 * r + 2.71155556874348757815e-5) * r + 0.0012426609473880784386) * r
                        + 0.026532189526576123093) * r + 0.29656057182850489123) * r + 1.7848265399172913358) * r + 5.4637849111641143699) * r
                        + 6.6579046435011037772)
                    / (((((((2.04426310338993978564e-15 * r + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r
                        + 7.868691311456132591e-4) * r + 0.0148753612908506148525) * r + 0.13692988092273580531) * r + 0.59983220655588793769) * r + 1);
            }
            if (q < 0) x = -x;
        }
        // Halley step against the accurate CDF (compare in the smaller tail to keep relative precision)
        double e = p < 0.5 ? NormCdf(x) - p : (1 - p) - NormCdf(-x);
        double u = e * Sqrt2Pi * Math.Exp(0.5 * x * x);
        if (double.IsFinite(u)) x -= u / (1 + 0.5 * x * u);
        return x;
    }

    /// <summary>
    /// Root of a continuous CDF: finds x in (lo, hi) with F(x) = p (equivalently 1 − F(x) = q, used when q is the smaller
    /// tail) by safeguarded Newton iteration on a shrinking bracket. <paramref name="cdf"/>(x, upper) returns F or 1 − F.
    /// </summary>
    private static double Solve(Func<double, bool, double> cdf, Func<double, double> pdf, double p, double q, double lo, double hi, double guess)
    {
        bool useLower = p <= q;
        double G(double x) => useLower ? cdf(x, false) - p : q - cdf(x, true);
        if (double.IsPositiveInfinity(hi))
        {
            double h = guess > lo && double.IsFinite(guess) ? guess : Math.Max(1, 2 * lo);
            int k = 0;
            while (G(h) < 0)
            {
                lo = h; h *= 2;
                if (++k > 3000 || double.IsInfinity(h)) return double.NaN;
            }
            hi = h;
        }
        double x = guess > lo && guess < hi ? guess : 0.5 * (lo + hi);
        for (int it = 0; it < 3000; it++)
        {
            double gx = G(x);
            if (gx == 0) return x;
            if (gx < 0) lo = x; else hi = x;
            double d = pdf(x), xn = x - gx / d;
            if (!(d > 0) || !(xn > lo && xn < hi)) xn = lo > 0 && hi > 4 * lo ? Math.Sqrt(lo * hi) : 0.5 * (lo + hi);
            if (Math.Abs(xn - x) <= 1e-15 * Math.Abs(xn)) return xn;
            x = xn;
        }
        return x;
    }

    /// <summary>Quantile of the unit-scale gamma distribution with shape a (p = lower tail, q = 1 − p).</summary>
    internal static double GammaInv(double a, double p, double q)
    {
        if (p <= 0) return 0;
        if (q <= 0) return double.PositiveInfinity;
        double z = NormInv(Math.Clamp(p, 1e-300, 1 - 1e-16)), c9 = 1 / (9 * a);
        double g = a * Math.Pow(1 - c9 + z * Math.Sqrt(c9), 3);
        if (!(g > 0) || a < 1) g =Math.Exp((Math.Log(p) + LogGamma(a + 1)) / a);
        return Solve((x, up) => GammaReg(a, x, up), x => GammaPdf(x, a), p, q, 0, double.PositiveInfinity, g);
    }

    internal static double BetaInv(double a, double b, double p, double q)
    {
        if (p <= 0) return 0;
        if (q <= 0) return 1;
        double mean = a / (a + b), lb = LogBeta(a, b), g;
        if (p <= q) g = Math.Min(mean, Math.Exp((Math.Log(a * p) + lb) / a));
        else g = Math.Max(mean, 1 - Math.Exp((Math.Log(b * q) + lb) / b));
        if (!(g > 0 && g < 1)) g = 0.5;
        return Solve((x, up) => BetaReg(x, 1 - x, a, b, up), x => BetaPdf(x, 1 - x, a, b), p, q, 0, 1, g);
    }

    internal static double FInv(double d1, double d2, double p, double q)
    {
        if (p <= 0) return 0;
        if (q <= 0) return double.PositiveInfinity;
        double g = d2 > 2 ? d2 / (d2 - 2) : 1;
        return Solve((x, up) => FCdf(x, d1, d2, up), x => FPdf(x, d1, d2), p, q, 0, double.PositiveInfinity, g);
    }

    /// <summary>t ≥ 0 with P(T &gt; t) = r (0 &lt; r ≤ ½).</summary>
    internal static double TInvTail(double r, double df)
    {
        if (r >= 0.5) return 0;
        if (df == 1) return 1 / Math.Tan(Math.PI * r);
        if (df == 2) return (1 - 2 * r) / Math.Sqrt(2 * r * (1 - r));
        double z = -NormInv(r);
        return Solve((x, up) => up ? TTail(x, df) : 1 - TTail(x, df), x => TPdf(x, df), 1 - r, r, 0, double.PositiveInfinity, z);
    }
}
