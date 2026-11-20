using System.Globalization;
using System.Text.RegularExpressions;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

/// <summary>Engineering functions: Bessel, base conversion, bitwise, complex numbers, CONVERT, ERF, DELTA, GESTEP.</summary>
public static partial class Functions
{
    internal static void RegisterEng()
    {
        // ------------------------------------------------------------ Bessel
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.Finite(EngMath.BesselI(EngMath.Num(c, a[0]), EngMath.Order(c, a[1])))); }, "BESSELI");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.Finite(EngMath.BesselJ(EngMath.Num(c, a[0]), EngMath.Order(c, a[1])))); }, "BESSELJ");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.Finite(EngMath.BesselK(EngMath.Num(c, a[0]), EngMath.Order(c, a[1])))); }, "BESSELK");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.Finite(EngMath.BesselY(EngMath.Num(c, a[0]), EngMath.Order(c, a[1])))); }, "BESSELY");

        // ------------------------------------------------------------ base conversion
        Reg((c, a) => { Args(a, 1, 1); return N(EngMath.ParseBase(EngMath.BaseText(c, a[0]), 2)); }, "BIN2DEC");
        Reg((c, a) => { Args(a, 1, 1); return N(EngMath.ParseBase(EngMath.BaseText(c, a[0]), 8)); }, "OCT2DEC");
        Reg((c, a) => { Args(a, 1, 1); return N(EngMath.ParseBase(EngMath.BaseText(c, a[0]), 16)); }, "HEX2DEC");
        foreach (var (name, from, to) in new[]
                 {
                     ("BIN2OCT", 2, 8), ("BIN2HEX", 2, 16), ("OCT2BIN", 8, 2), ("OCT2HEX", 8, 16), ("HEX2BIN", 16, 2), ("HEX2OCT", 16, 8)
                 })
        {
            int f = from, t = to;
            Reg((c, a) => { Args(a, 1, 2); return S(EngMath.ToBase(EngMath.ParseBase(EngMath.BaseText(c, a[0]), f), t, c, a, 1)); }, name);
        }
        foreach (var (name, to) in new[] { ("DEC2BIN", 2), ("DEC2OCT", 8), ("DEC2HEX", 16) })
        {
            int t = to;
            Reg((c, a) =>
            {
                Args(a, 1, 2);
                double d = Math.Truncate(EngMath.Num(c, a[0]));
                if (Math.Abs(d) > 1e15) throw Err(Errors.Num);
                return S(EngMath.ToBase((long)d, t, c, a, 1));
            }, name);
        }

        // ------------------------------------------------------------ bitwise
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.BitArg(c, a[0]) & EngMath.BitArg(c, a[1])); }, "BITAND");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.BitArg(c, a[0]) | EngMath.BitArg(c, a[1])); }, "BITOR");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.BitArg(c, a[0]) ^ EngMath.BitArg(c, a[1])); }, "BITXOR");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.Shift(EngMath.BitArg(c, a[0]), Math.Truncate(c.Num(a[1])))); }, "BITLSHIFT");
        Reg((c, a) => { Args(a, 2, 2); return N(EngMath.Shift(EngMath.BitArg(c, a[0]), -Math.Truncate(c.Num(a[1])))); }, "BITRSHIFT");

        // ------------------------------------------------------------ simple tests and error function
        Reg((c, a) => { Args(a, 1, 2); return N(EngMath.Num(c, a[0]) == EngMath.NumOr(c, a, 1, 0) ? 1 : 0); }, "DELTA");
        Reg((c, a) => { Args(a, 1, 2); return N(EngMath.Num(c, a[0]) >= EngMath.NumOr(c, a, 1, 0) ? 1 : 0); }, "GESTEP");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            double lo = EngMath.Num(c, a[0]);
            if (a.Length < 2 || a[1] is MissingExpr) return N(EngMath.Erf(lo));
            return N(EngMath.Erf(EngMath.Num(c, a[1])) - EngMath.Erf(lo));
        }, "ERF");
        Reg((c, a) => { Args(a, 1, 1); return N(EngMath.Erf(c.Num(a[0]))); }, "ERF.PRECISE");
        Reg((c, a) => { Args(a, 1, 1); return N(EngMath.Erfc(EngMath.Num(c, a[0]))); }, "ERFC");
        Reg((c, a) => { Args(a, 1, 1); return N(EngMath.Erfc(c.Num(a[0]))); }, "ERFC.PRECISE");

        // ------------------------------------------------------------ CONVERT
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double v = EngMath.Num(c, a[0]);
            return N(EngMath.Finite(EngMath.Convert(v, c.Str(a[1]), c.Str(a[2]))));
        }, "CONVERT");

        // ------------------------------------------------------------ complex numbers
        void RegCx1(string name, Func<double, double, (double, double)> f)
        {
            Reg((c, a) =>
            {
                Args(a, 1, 1);
                var z = EngMath.Cx(c, a[0]);
                var (x, y) = f(z.Re, z.Im);
                return S(EngMath.FormatCx(x, y, z.Suf ?? 'i'));
            }, name);
        }
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            double re = EngMath.Num(c, a[0]), im = EngMath.Num(c, a[1]);
            char suf = 'i';
            if (a.Length > 2 && a[2] is not MissingExpr)
            {
                var s = c.Str(a[2]);
                if (s == "j") suf = 'j';
                else if (s != "i" && s.Length != 0) throw Err(Errors.Value);
            }
            return S(EngMath.FormatCx(re, im, suf));
        }, "COMPLEX");
        Reg((c, a) => { Args(a, 1, 1); var z = EngMath.Cx(c, a[0]); return N(z.Re); }, "IMREAL");
        Reg((c, a) => { Args(a, 1, 1); var z = EngMath.Cx(c, a[0]); return N(z.Im); }, "IMAGINARY");
        Reg((c, a) => { Args(a, 1, 1); var z = EngMath.Cx(c, a[0]); return N(EngMath.Hypot(z.Re, z.Im)); }, "IMABS");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var z = EngMath.Cx(c, a[0]);
            if (z.Re == 0 && z.Im == 0) throw Err(Errors.Div0);
            return N(Math.Atan2(z.Im, z.Re));
        }, "IMARGUMENT");
        RegCx1("IMCONJUGATE", (re, im) => (re, -im));
        RegCx1("IMCOS", EngMath.CxCos);
        RegCx1("IMCOSH", EngMath.CxCosh);
        RegCx1("IMSIN", EngMath.CxSin);
        RegCx1("IMSINH", EngMath.CxSinh);
        RegCx1("IMTAN", EngMath.CxTan);
        RegCx1("IMCOT", EngMath.CxCot);
        RegCx1("IMSEC", (re, im) => { var (x, y) = EngMath.CxCos(re, im); return EngMath.CxRecip(x, y); });
        RegCx1("IMCSC", (re, im) => { var (x, y) = EngMath.CxSin(re, im); return EngMath.CxRecip(x, y); });
        RegCx1("IMSECH", (re, im) => { var (x, y) = EngMath.CxCosh(re, im); return EngMath.CxRecip(x, y); });
        RegCx1("IMCSCH", (re, im) => { var (x, y) = EngMath.CxSinh(re, im); return EngMath.CxRecip(x, y); });
        RegCx1("IMEXP", EngMath.CxExp);
        RegCx1("IMLN", EngMath.CxLn);
        RegCx1("IMLOG10", (re, im) => { var (x, y) = EngMath.CxLn(re, im); return (x / Math.Log(10), y / Math.Log(10)); });
        RegCx1("IMLOG2", (re, im) => { var (x, y) = EngMath.CxLn(re, im); return (x / Math.Log(2), y / Math.Log(2)); });
        RegCx1("IMSQRT", (re, im) => EngMath.CxPow(re, im, 0.5));
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var z = EngMath.Cx(c, a[0]);
            double n = EngMath.Num(c, a[1]);
            var (x, y) = EngMath.CxPow(z.Re, z.Im, n);
            return S(EngMath.FormatCx(x, y, z.Suf ?? 'i'));
        }, "IMPOWER");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var p = EngMath.Cx(c, a[0]); var q = EngMath.Cx(c, a[1]);
            return S(EngMath.FormatCx(p.Re - q.Re, p.Im - q.Im, EngMath.Suffix(p.Suf, q.Suf)));
        }, "IMSUB");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var p = EngMath.Cx(c, a[0]); var q = EngMath.Cx(c, a[1]);
            char suf = EngMath.Suffix(p.Suf, q.Suf);
            double d = q.Re * q.Re + q.Im * q.Im;
            if (d == 0) throw Err(Errors.Num);
            return S(EngMath.FormatCx((p.Re * q.Re + p.Im * q.Im) / d, (p.Im * q.Re - p.Re * q.Im) / d, suf));
        }, "IMDIV");
        Reg((c, a) =>
        {
            Args(a, 1);
            double re = 0, im = 0; char? suf = null;
            foreach (var z in EngMath.CxList(c, a)) { re += z.Re; im += z.Im; suf = EngMath.SuffixAcc(suf, z.Suf); }
            return S(EngMath.FormatCx(re, im, suf ?? 'i'));
        }, "IMSUM");
        Reg((c, a) =>
        {
            Args(a, 1);
            double re = 1, im = 0; char? suf = null;
            foreach (var z in EngMath.CxList(c, a))
            {
                (re, im) = (re * z.Re - im * z.Im, re * z.Im + im * z.Re);
                suf = EngMath.SuffixAcc(suf, z.Suf);
            }
            return S(EngMath.FormatCx(re, im, suf ?? 'i'));
        }, "IMPRODUCT");
    }

}

/// <summary>Numerical helpers of the engineering functions.</summary>
internal static class EngMath
{
    private static FormulaError Err(string code) => new(code);

    // ================================================================== argument helpers

    /// <summary>Number argument of an Analysis ToolPak function: logical values are rejected with #VALUE!.</summary>
    internal static double Num(Ctx c, Expr e)
    {
        var v = c.Scalar(e);
        if (v.IsBool) throw Err(Errors.Value);
        return Calc.ToNum(v);
    }

    internal static double NumOr(Ctx c, Expr[] a, int i, double def) => i < a.Length && a[i] is not MissingExpr ? Num(c, a[i]) : def;

    internal static int Order(Ctx c, Expr e)
    {
        double n = Math.Truncate(Num(c, e));
        if (n < 0 || n > int.MaxValue) throw Err(Errors.Num);
        return (int)n;
    }

    internal static double Finite(double d) => double.IsNaN(d) || double.IsInfinity(d) ? throw Err(Errors.Num) : d;

    // ================================================================== base conversion

    internal static string BaseText(Ctx c, Expr e)
    {
        var v = c.Scalar(e);
        if (v.IsError) throw Err(v.Str!);
        if (v.IsBool) throw Err(Errors.Value);
        if (v.IsBlank) return "";
        return Calc.ToText(v);
    }

    private static int Bits(int radix) => radix switch { 2 => 10, 8 => 30, _ => 40 };

    /// <summary>Parses up to 10 digits; a 10-digit value with the top bit set is a two's complement negative.</summary>
    internal static long ParseBase(string s, int radix)
    {
        if (s.Length > 10) throw Err(Errors.Num);
        long v = 0;
        foreach (char ch in s)
        {
            int d = ch >= '0' && ch <= '9' ? ch - '0' : ch >= 'A' && ch <= 'F' ? ch - 'A' + 10 : ch >= 'a' && ch <= 'f' ? ch - 'a' + 10 : 99;
            if (d >= radix) throw Err(Errors.Num);
            v = v * radix + d;
        }
        int bits = Bits(radix);
        if (s.Length == 10 && v >= 1L << (bits - 1)) v -= 1L << bits;
        return v;
    }

    /// <summary>Formats <paramref name="v"/> in <paramref name="radix"/> (10-digit two's complement when negative), honouring places.</summary>
    internal static string ToBase(long v, int radix, Ctx c, Expr[] a, int placesIndex)
    {
        int bits = Bits(radix);
        long lim = 1L << (bits - 1);
        if (v < -lim || v >= lim) throw Err(Errors.Num);
        bool hasPlaces = placesIndex < a.Length && a[placesIndex] is not MissingExpr;
        double places = hasPlaces ? Math.Truncate(Num(c, a[placesIndex])) : 0;
        if (v < 0) return System.Convert.ToString(v + (1L << bits), radix).ToUpperInvariant();
        var s = System.Convert.ToString(v, radix).ToUpperInvariant();
        if (!hasPlaces) return s;
        if (places < 1 || places > 10 || s.Length > places) throw Err(Errors.Num);
        return s.PadLeft((int)places, '0');
    }

    // ================================================================== bitwise

    private const double Max48 = 281474976710655; // 2^48 - 1

    internal static long BitArg(Ctx c, Expr e)
    {
        double d = c.Num(e);
        if (d < 0 || d != Math.Floor(d) || d > Max48) throw Err(Errors.Num);
        return (long)d;
    }

    internal static double Shift(long n, double shift)
    {
        if (Math.Abs(shift) > 53) throw Err(Errors.Num);
        int s = (int)shift;
        if (s <= 0) return n >> -s;
        double r = n * Math.Pow(2, s);
        if (r > Max48) throw Err(Errors.Num);
        return r;
    }

    // ================================================================== error function

    private const double TwoOverSqrtPi = 1.1283791670955126;
    private const double InvSqrtPi = 0.56418958354775628;

    internal static double Erf(double x)
    {
        if (x < 0) return -Erf(-x);
        if (x < 2) return ErfSeries(x);
        return 1 - ErfcCf(x);
    }

    internal static double Erfc(double x)
    {
        if (x < 0) return 2 - Erfc(-x);
        if (x < 2) return 1 - ErfSeries(x);
        return ErfcCf(x);
    }

    /// <summary>erf(x) = 2/sqrt(pi) e^(-x^2) sum 2^n x^(2n+1)/(1*3*...*(2n+1)) — all terms positive.</summary>
    private static double ErfSeries(double x)
    {
        double x2 = 2 * x * x, term = x, sum = x;
        for (int n = 1; n < 500; n++)
        {
            term *= x2 / (2 * n + 1);
            sum += term;
            if (term < 1e-17 * sum) break;
        }
        return TwoOverSqrtPi * Math.Exp(-x * x) * sum;
    }

    /// <summary>erfc(x) for x &gt;= 2 by the continued fraction e^(-x^2)/sqrt(pi) / (x + (1/2)/(x + 1/(x + (3/2)/(x + ...)))).</summary>
    private static double ErfcCf(double x)
    {
        if (x > 27) return 0;
        const double tiny = 1e-300;
        double f = x, C = f, D = 0;
        for (int n = 1; n < 5000; n++)
        {
            double an = n / 2.0;
            D = x + an * D; if (D == 0) D = tiny; D = 1 / D;
            C = x + an / C; if (C == 0) C = tiny;
            double delta = C * D;
            f *= delta;
            if (Math.Abs(delta - 1) < 1e-16) break;
        }
        return InvSqrtPi * Math.Exp(-x * x) / f;
    }

    // ================================================================== Bessel functions

    private const double EulerGamma = 0.57721566490153286;
    private const int MaxOrder = 10_000_000;

    /// <summary>True when (x/2)^n/n! * e^(x^2/(4(n+1))) — an upper bound of |J_n| and I_n — underflows.</summary>
    private static bool Underflows(double x, int n)
    {
        if (n < 100_000) return false;
        double lgam = n * Math.Log(n) - n + 0.5 * Math.Log(2 * Math.PI * n); // ln n! (Stirling, ample for this test)
        return n * Math.Log(x / 2) - lgam + x * x / (4.0 * (n + 1)) < -760;
    }

    /// <summary>Leading series term (x/2)^n/n! for tiny x.</summary>
    private static double SmallTerm(double x, int n)
    {
        double v = 1;
        for (int k = 1; k <= n && v != 0; k++) v *= x / 2 / k;
        return v;
    }

    /// <summary>Hankel asymptotic P and Q for order nu (x &gt;= 25).</summary>
    private static void Hankel(double x, int nu, out double p, out double q)
    {
        double mu = 4.0 * nu * nu, t = 1, prev = double.MaxValue;
        p = 1; q = 0;
        for (int k = 1; k < 300; k++)
        {
            double odd = 2 * k - 1;
            double tk = t * (mu - odd * odd) / (8.0 * k * x);
            if (Math.Abs(tk) >= prev) break; // asymptotic series starts diverging
            switch (k % 4)
            {
                case 1: q += tk; break;
                case 2: p -= tk; break;
                case 3: q -= tk; break;
                default: p += tk; break;
            }
            if (Math.Abs(tk) < 1e-18) break;
            prev = Math.Abs(tk);
            t = tk;
        }
    }

    private static void BesselLarge(double x, out double j0, out double j1, out double y0, out double y1)
    {
        Hankel(x, 0, out var p0, out var q0);
        Hankel(x, 1, out var p1, out var q1);
        double s = Math.Sin(x), c = Math.Cos(x), f = Math.Sqrt(1 / (Math.PI * x));
        j0 = f * (p0 * (c + s) - q0 * (s - c));
        y0 = f * (p0 * (s - c) + q0 * (c + s));
        j1 = f * (p1 * (s - c) + q1 * (s + c));
        y1 = f * (-p1 * (s + c) + q1 * (s - c));
    }

    /// <summary>Miller's backward recurrence for J_n(x), x &gt; 0, normalised with J0 + 2(J2 + J4 + ...) = 1.</summary>
    private static double JMiller(double x, int n, double[]? all)
    {
        int top = Math.Max(n, (int)Math.Ceiling(x));
        int m = top + 30 + (int)(4 * Math.Sqrt(top));
        if (m % 2 == 1) m++;
        if (all != null && all.Length < m + 1) throw new ArgumentException("array too small");
        double next = 0, cur = 1, sum = 0, ans = 0;
        for (int k = m; k >= 1; k--)
        {
            if (k == n) ans = cur;
            if (all != null) all[k] = cur;
            if (k % 2 == 0) sum += 2 * cur;
            double prev = 2.0 * k / x * cur - next;
            next = cur; cur = prev;
            if (Math.Abs(cur) > 1e250)
            {
                cur *= 1e-250; next *= 1e-250; sum *= 1e-250; ans *= 1e-250;
                if (all != null) for (int i = k; i <= m; i++) all[i] *= 1e-250;
            }
        }
        if (n == 0) ans = cur;
        sum += cur;
        if (all != null) { all[0] = cur; for (int i = 0; i <= m; i++) all[i] /= sum; }
        return ans / sum;
    }

    private static int JMillerSize(double x) { int top = (int)Math.Ceiling(x); return top + 32 + (int)(4 * Math.Sqrt(top)); }

    internal static double BesselJ(double x, int n)
    {
        double sign = x < 0 && (n & 1) == 1 ? -1 : 1;
        x = Math.Abs(x);
        if (x == 0) return n == 0 ? 1 : 0;
        if (Underflows(x, n)) return 0;
        if (n > MaxOrder) throw Err(Errors.Num);
        if (x < 1e-20) return sign * SmallTerm(x, n);
        if (x >= 25 && n < x)
        {
            BesselLarge(x, out var j0, out var j1, out _, out _);
            if (n == 0) return sign * j0;
            for (int k = 1; k < n; k++) (j0, j1) = (j1, 2.0 * k / x * j1 - j0);
            return sign * j1;
        }
        return sign * JMiller(x, n, null);
    }

    internal static double BesselY(double x, int n)
    {
        if (x <= 0 || n > MaxOrder) throw Err(Errors.Num);
        double y0, y1;
        if (x >= 25) BesselLarge(x, out _, out _, out y0, out y1);
        else if (x < 1e-20)
        {
            y0 = 2 / Math.PI * (Math.Log(x / 2) + EulerGamma);
            y1 = -2 / (Math.PI * x);
        }
        else
        {
            // Neumann series: Y0 = 2/pi (ln(x/2)+g) J0 - 4/pi sum (-1)^k J2k/k,
            //                 Y1 = -Y0' = 2/pi ((ln(x/2)+g) J1 - J0/x) + 2/pi sum (-1)^k (J(2k-1) - J(2k+1))/k
            var j = new double[JMillerSize(x) + 2];
            JMiller(x, 0, j);
            double L = Math.Log(x / 2) + EulerGamma, s0 = 0, s1 = 0;
            for (int k = 1; 2 * k + 1 < j.Length; k++)
            {
                double sg = (k & 1) == 1 ? -1 : 1;
                s0 += sg * j[2 * k] / k;
                s1 += sg * (j[2 * k - 1] - j[2 * k + 1]) / k;
            }
            y0 = 2 / Math.PI * (L * j[0] - 2 * s0);
            y1 = 2 / Math.PI * (L * j[1] - j[0] / x + s1);
        }
        if (n == 0) return y0;
        for (int k = 1; k < n; k++)
        {
            (y0, y1) = (y1, 2.0 * k / x * y1 - y0);
            if (double.IsInfinity(y1)) throw Err(Errors.Num);
        }
        return y1;
    }

    internal static double BesselI(double x, int n)
    {
        double sign = x < 0 && (n & 1) == 1 ? -1 : 1;
        x = Math.Abs(x);
        if (x == 0) return n == 0 ? 1 : 0;
        if (Underflows(x, n)) return 0;
        if (n > MaxOrder || x > 1e6) throw Err(Errors.Num);
        if (x < 1e-20) return sign * SmallTerm(x, n);
        // Miller's backward recurrence normalised with e^x = I0 + 2(I1 + I2 + ...)
        int m = Math.Max(n, (int)Math.Sqrt(80 * x)) + 40 + (int)(4 * Math.Sqrt(n));
        // ans is kept unscaled; 'drop' counts the 1e-250 rescalings applied to the sum after ans was taken
        double next = 0, cur = 1, sum = 0, ans = 0;
        int drop = 0;
        for (int k = m; k >= 1; k--)
        {
            if (k == n) { ans = cur; drop = 0; }
            sum += 2 * cur;
            double prev = 2.0 * k / x * cur + next;
            next = cur; cur = prev;
            if (cur > 1e250) { cur *= 1e-250; next *= 1e-250; sum *= 1e-250; drop++; }
        }
        if (n == 0) { ans = cur; drop = 0; }
        sum += cur;
        double logI = x + Math.Log(ans) - Math.Log(sum) - drop * 250 * Math.Log(10);
        double r = x < 700 && drop == 0 ? ans / sum * Math.Exp(x) : Math.Exp(logI);
        return sign * Finite(r);
    }

    /// <summary>e^x K_nu(x) for nu = 0, 1 by the trapezoidal rule on K_nu(x) = integral_0^inf e^(-x cosh t) cosh(nu t) dt
    /// (the integrand is analytic in a strip, so the rule converges exponentially).</summary>
    private static double KScaled(double x, int nu)
    {
        double h = Math.Min(0.1, 0.5 / Math.Sqrt(x)), sum = 0.5;
        for (int j = 1; j < 100000; j++)
        {
            double t = j * h, sh = Math.Sinh(t / 2);
            double f = Math.Exp(-2 * x * sh * sh) * (nu == 0 ? 1 : Math.Cosh(t));
            sum += f;
            if (f < 1e-18 * sum) break;
        }
        return h * sum;
    }

    internal static double BesselK(double x, int n)
    {
        if (x <= 0 || n > MaxOrder) throw Err(Errors.Num);
        // forward recurrence on e^x K_n(x), rescaled by 1e-250 whenever it grows large ('lift' counts the rescalings)
        double k0 = KScaled(x, 0), k1 = n == 0 ? 0 : KScaled(x, 1);
        double kn = k0;
        int lift = 0;
        if (n > 0)
        {
            for (int k = 1; k < n; k++)
            {
                (k0, k1) = (k1, k0 + 2.0 * k / x * k1);
                if (k1 > 1e250) { k0 *= 1e-250; k1 *= 1e-250; lift++; }
                if (double.IsInfinity(k1)) throw Err(Errors.Num);
            }
            kn = k1;
        }
        if (x < 700 && lift == 0) return kn * Math.Exp(-x);
        return Finite(Math.Exp(Math.Log(kn) + lift * 250 * Math.Log(10) - x));
    }

    // ================================================================== complex numbers

    internal readonly record struct Complex(double Re, double Im, char? Suf);

    private const string NumRx = @"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?";
    private static readonly Regex CxRx = new(
        $@"^(?:(?<r>[+-]?{NumRx})(?<i>[+-](?:{NumRx})?[ij])?|(?<i>[+-]?(?:{NumRx})?[ij]))$", RegexOptions.CultureInvariant);

    internal static Complex Cx(Ctx c, Expr e) => Cx(c.Scalar(e));

    internal static Complex Cx(CellValue v)
    {
        switch (v.Kind)
        {
            case ValueKind.Error: throw Err(v.Str!);
            case ValueKind.Bool: throw Err(Errors.Value);
            case ValueKind.Blank: return new Complex(0, 0, null);
            case ValueKind.Number: return new Complex(v.Num, 0, null);
        }
        var s = v.Str!;
        if (s.Length == 0) return new Complex(0, 0, null);
        var m = CxRx.Match(s);
        if (!m.Success) throw Err(Errors.Num);
        double re = m.Groups["r"].Success ? ParseD(m.Groups["r"].Value) : 0, im = 0;
        char? suf = null;
        if (m.Groups["i"].Success)
        {
            var g = m.Groups["i"].Value;
            suf = g[^1];
            var coef = g[..^1];
            im = coef switch { "" or "+" => 1, "-" => -1, _ => ParseD(coef) };
        }
        return new Complex(re, im, suf);
    }

    private static double ParseD(string s)
    {
        double d = double.Parse(s, NumberStyles.Float, CultureInfo.InvariantCulture);
        return Finite(d);
    }

    /// <summary>Every complex value of the arguments; ranges and arrays are flattened, blank cells skipped.</summary>
    internal static IEnumerable<Complex> CxList(Ctx c, Expr[] args)
    {
        foreach (var e in args)
        {
            if (e is MissingExpr) continue;
            var v = c.Eval(e);
            if (v is CellValue s) { yield return Cx(s); continue; }
            foreach (var x in Functions.FlatVal(c, v))
                if (!x.IsBlank) yield return Cx(x);
        }
    }

    /// <summary>Common suffix of two operands; mixing i and j gives #VALUE!.</summary>
    internal static char Suffix(char? a, char? b) => a != null && b != null && a != b ? throw Err(Errors.Value) : a ?? b ?? 'i';

    /// <summary>Running suffix of a list of operands (IMSUM, IMPRODUCT).</summary>
    internal static char? SuffixAcc(char? acc, char? next) => next == null ? acc : acc != null && acc != next ? throw Err(Errors.Value) : next;

    private static string Fmt(double d) => CellValue.FormatGeneralInvariant(d == 0 ? 0 : d);

    internal static string FormatCx(double re, double im, char suf)
    {
        Finite(re); Finite(im);
        if (im == 0) return Fmt(re);
        var ims = Fmt(im);
        ims = ims == "1" ? "" : ims == "-1" ? "-" : ims;
        if (re == 0) return ims + suf;
        return Fmt(re) + (im > 0 ? "+" : "") + ims + suf;
    }

    internal static double Hypot(double a, double b)
    {
        a = Math.Abs(a); b = Math.Abs(b);
        if (a < b) (a, b) = (b, a);
        if (a == 0) return 0;
        if (a < 1e150 && b > 1e-150) return Math.Sqrt(a * a + b * b);
        double r = b / a;
        return a * Math.Sqrt(1 + r * r);
    }

    internal static (double, double) CxRecip(double x, double y)
    {
        if (x == 0 && y == 0) throw Err(Errors.Num);
        if (Math.Abs(x) >= Math.Abs(y)) { double r = y / x, d = x + y * r; return (1 / d, -r / d); }
        else { double r = x / y, d = x * r + y; return (r / d, -1 / d); }
    }

    internal static (double, double) CxSin(double a, double b) => (Math.Sin(a) * Math.Cosh(b), Math.Cos(a) * Math.Sinh(b));
    internal static (double, double) CxCos(double a, double b) => (Math.Cos(a) * Math.Cosh(b), -Math.Sin(a) * Math.Sinh(b));
    internal static (double, double) CxSinh(double a, double b) => (Math.Sinh(a) * Math.Cos(b), Math.Cosh(a) * Math.Sin(b));
    internal static (double, double) CxCosh(double a, double b) => (Math.Cosh(a) * Math.Cos(b), Math.Sinh(a) * Math.Sin(b));

    internal static (double, double) CxTan(double a, double b)
    {
        if (Math.Abs(b) > 350) return (0, Math.Sign(b));
        double d = Math.Cos(2 * a) + Math.Cosh(2 * b);
        if (d == 0) throw Err(Errors.Num);
        return (Math.Sin(2 * a) / d, Math.Sinh(2 * b) / d);
    }

    internal static (double, double) CxCot(double a, double b)
    {
        if (Math.Abs(b) > 350) return (0, -Math.Sign(b));
        double d = Math.Cosh(2 * b) - Math.Cos(2 * a);
        if (d == 0) throw Err(Errors.Num);
        return (Math.Sin(2 * a) / d, -Math.Sinh(2 * b) / d);
    }

    internal static (double, double) CxExp(double a, double b)
    {
        double e = Math.Exp(a);
        return (e * Math.Cos(b), e * Math.Sin(b));
    }

    internal static (double, double) CxLn(double a, double b)
    {
        if (a == 0 && b == 0) throw Err(Errors.Num);
        return (Math.Log(Hypot(a, b)), Math.Atan2(b, a));
    }

    /// <summary>z^n in polar form r^n (cos n t + i sin n t), as Excel computes IMPOWER and IMSQRT.</summary>
    internal static (double, double) CxPow(double a, double b, double n)
    {
        if (a == 0 && b == 0)
        {
            if (n > 0) return (0, 0);
            throw Err(Errors.Num);
        }
        double rn = Math.Pow(Hypot(a, b), n), t = n * Math.Atan2(b, a);
        return (rn * Math.Cos(t), rn * Math.Sin(t));
    }

    // ================================================================== CONVERT

    private enum UCat { Mass, Distance, Time, Pressure, Force, Energy, Power, Magnetism, Temperature, Volume, Area, Information, Speed }

    /// <summary>A unit: factor to the category's base unit (g, m, s, Pa, N, J, W, T, -, m3, m2, bit, m/s),
    /// whether metric prefixes apply and the power the prefix is raised to (m2: 2, m3: 3).</summary>
    private sealed record Unit(UCat Cat, double Factor, bool Prefix, int Power);

    private static readonly Dictionary<string, Unit> Units = BuildUnits();

    private static readonly (string P, int Exp)[] MetricPrefixes =
    [
        ("Y", 24), ("Z", 21), ("E", 18), ("P", 15), ("T", 12), ("G", 9), ("M", 6), ("k", 3), ("h", 2), ("da", 1),
        ("d", -1), ("c", -2), ("m", -3), ("u", -6), ("n", -9), ("p", -12), ("f", -15), ("a", -18), ("z", -21), ("y", -24)
    ];

    private static readonly (string P, int Exp)[] BinaryPrefixes =
        [("Yi", 80), ("Zi", 70), ("Ei", 60), ("Pi", 50), ("Ti", 40), ("Gi", 30), ("Mi", 20), ("ki", 10)];

    private static Dictionary<string, Unit> BuildUnits()
    {
        var d = new Dictionary<string, Unit>(StringComparer.Ordinal);
        void Add(UCat cat, double f, bool prefix, params string[] names) { foreach (var n in names) d[n] = new Unit(cat, f, prefix, 1); }
        void AddP(UCat cat, double f, int power, params string[] names) { foreach (var n in names) d[n] = new Unit(cat, f, true, power); }

        const double lbm = 453.59237, inch = 0.0254, ft = 0.3048, yd = 0.9144, mi = 1609.344, nmi = 1852;
        const double ly = 9460730472580800, pica = inch / 72, gn = 9.80665;
        const double lbf = lbm / 1000 * gn, hp = 550 * ft * lbf, usGal = 231 * inch * inch * inch;

        // Weight and mass (gram)
        Add(UCat.Mass, 1, true, "g");
        Add(UCat.Mass, lbf / ft * 1000, false, "sg");
        Add(UCat.Mass, lbm, false, "lbm");
        Add(UCat.Mass, 1.66053100460465E-24, true, "u");
        Add(UCat.Mass, lbm / 16, false, "ozm");
        Add(UCat.Mass, 0.06479891, false, "grain");
        Add(UCat.Mass, 100 * lbm, false, "cwt", "shweight");
        Add(UCat.Mass, 112 * lbm, false, "uk_cwt", "lcwt", "hweight");
        Add(UCat.Mass, 14 * lbm, false, "stone");
        Add(UCat.Mass, 2000 * lbm, false, "ton");
        Add(UCat.Mass, 2240 * lbm, false, "uk_ton", "LTON", "brton");

        // Distance (meter)
        Add(UCat.Distance, 1, true, "m");
        Add(UCat.Distance, mi, false, "mi");
        Add(UCat.Distance, nmi, false, "Nmi");
        Add(UCat.Distance, inch, false, "in");
        Add(UCat.Distance, ft, false, "ft");
        Add(UCat.Distance, yd, false, "yd");
        Add(UCat.Distance, 1e-10, true, "ang");
        Add(UCat.Distance, 45 * inch, false, "ell");
        Add(UCat.Distance, ly, false, "ly");
        Add(UCat.Distance, 3.08567782E16, false, "parsec", "pc");
        Add(UCat.Distance, pica, false, "Picapt", "Pica");
        Add(UCat.Distance, inch / 6, false, "pica");
        Add(UCat.Distance, 6336000.0 / 3937, false, "survey_mi");

        // Time (second)
        Add(UCat.Time, 365.25 * 86400, false, "yr");
        Add(UCat.Time, 86400, false, "day", "d");
        Add(UCat.Time, 3600, false, "hr");
        Add(UCat.Time, 60, false, "mn", "min");
        Add(UCat.Time, 1, true, "sec", "s");

        // Pressure (pascal)
        Add(UCat.Pressure, 1, true, "Pa", "p");
        Add(UCat.Pressure, 101325, true, "atm", "at");
        Add(UCat.Pressure, 101325.0 / 760, true, "mmHg");
        Add(UCat.Pressure, lbf / (inch * inch), false, "psi");
        Add(UCat.Pressure, 101325.0 / 760, false, "Torr");

        // Force (newton)
        Add(UCat.Force, 1, true, "N");
        Add(UCat.Force, 1e-5, true, "dyn", "dy");
        Add(UCat.Force, lbf, false, "lbf");
        Add(UCat.Force, gn / 1000, true, "pond");

        // Energy (joule)
        Add(UCat.Energy, 1, true, "J");
        Add(UCat.Energy, 1e-7, true, "e");
        Add(UCat.Energy, 4.184, true, "c");
        Add(UCat.Energy, 4.1868, true, "cal");
        Add(UCat.Energy, 1.602176487E-19, true, "eV", "ev");
        Add(UCat.Energy, hp * 3600, false, "HPh", "hh");
        Add(UCat.Energy, 3600, true, "Wh", "wh");
        Add(UCat.Energy, ft * lbm / 1000 * ft, false, "flb"); // Excel's "flb" is the foot-poundal
        Add(UCat.Energy, 1055.05585262, false, "BTU", "btu");

        // Power (watt)
        Add(UCat.Power, hp, false, "HP", "h");
        Add(UCat.Power, 75 * gn, false, "PS");
        Add(UCat.Power, 1, true, "W", "w");

        // Magnetism (tesla)
        Add(UCat.Magnetism, 1, true, "T");
        Add(UCat.Magnetism, 1e-4, true, "ga");

        // Temperature (handled separately; only kelvin takes prefixes)
        Add(UCat.Temperature, 0, false, "C", "cel", "F", "fah", "Rank", "Reau");
        Add(UCat.Temperature, 1, true, "K", "kel");

        // Volume (cubic meter)
        Add(UCat.Volume, usGal / 768, false, "tsp");
        Add(UCat.Volume, 5e-6, false, "tspm");
        Add(UCat.Volume, usGal / 256, false, "tbs");
        Add(UCat.Volume, usGal / 128, false, "oz");
        Add(UCat.Volume, usGal / 16, false, "cup");
        Add(UCat.Volume, usGal / 8, false, "pt", "us_pt");
        Add(UCat.Volume, 4.54609e-3 / 8, false, "uk_pt");
        Add(UCat.Volume, usGal / 4, false, "qt");
        Add(UCat.Volume, 4.54609e-3 / 4, false, "uk_qt");
        Add(UCat.Volume, usGal, false, "gal");
        Add(UCat.Volume, 4.54609e-3, false, "uk_gal");
        Add(UCat.Volume, 1e-3, true, "l", "L", "lt");
        AddP(UCat.Volume, 1e-30, 3, "ang3", "ang^3");
        Add(UCat.Volume, 42 * usGal, false, "barrel");
        Add(UCat.Volume, 2150.42 * inch * inch * inch, false, "bushel");
        Add(UCat.Volume, ft * ft * ft, false, "ft3", "ft^3");
        Add(UCat.Volume, inch * inch * inch, false, "in3", "in^3");
        Add(UCat.Volume, ly * ly * ly, false, "ly3", "ly^3");
        AddP(UCat.Volume, 1, 3, "m3", "m^3");
        Add(UCat.Volume, mi * mi * mi, false, "mi3", "mi^3");
        Add(UCat.Volume, yd * yd * yd, false, "yd3", "yd^3");
        Add(UCat.Volume, nmi * nmi * nmi, false, "Nmi3", "Nmi^3");
        Add(UCat.Volume, pica * pica * pica, false, "Picapt3", "Picapt^3", "Pica3", "Pica^3");
        Add(UCat.Volume, 100 * ft * ft * ft, false, "GRT", "regton");
        Add(UCat.Volume, 40 * ft * ft * ft, false, "MTON");

        // Area (square meter)
        Add(UCat.Area, 4046.8564224, false, "uk_acre");
        Add(UCat.Area, 43560 * (1200.0 / 3937) * (1200.0 / 3937), false, "us_acre");
        AddP(UCat.Area, 1e-20, 2, "ang2", "ang^2");
        Add(UCat.Area, 100, true, "ar");
        Add(UCat.Area, ft * ft, false, "ft2", "ft^2");
        Add(UCat.Area, 10000, false, "ha");
        Add(UCat.Area, inch * inch, false, "in2", "in^2");
        Add(UCat.Area, ly * ly, false, "ly2", "ly^2");
        AddP(UCat.Area, 1, 2, "m2", "m^2");
        Add(UCat.Area, 2500, false, "Morgen");
        Add(UCat.Area, mi * mi, false, "mi2", "mi^2");
        Add(UCat.Area, nmi * nmi, false, "Nmi2", "Nmi^2");
        Add(UCat.Area, pica * pica, false, "Picapt2", "Pica2", "Pica^2", "Picapt^2");
        Add(UCat.Area, yd * yd, false, "yd2", "yd^2");

        // Information (bit)
        Add(UCat.Information, 1, true, "bit");
        Add(UCat.Information, 8, true, "byte");

        // Speed (meter per second)
        Add(UCat.Speed, 6080 * ft / 3600, false, "admkn");
        Add(UCat.Speed, nmi / 3600, false, "kn");
        Add(UCat.Speed, 1.0 / 3600, true, "m/h", "m/hr");
        Add(UCat.Speed, 1, true, "m/s", "m/sec");
        Add(UCat.Speed, mi / 3600, false, "mph");
        return d;
    }

    private static double Pow10(int e) => double.Parse("1e" + e.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

    /// <summary>Unit with optional metric/binary prefix; returns the unit and the prefix multiplier (already raised to the unit's power).</summary>
    private static (Unit U, double Mult) FindUnit(string s)
    {
        if (Units.TryGetValue(s, out var u)) return (u, 1);
        foreach (var (p, e) in BinaryPrefixes)
            if (s.Length > p.Length && s.StartsWith(p, StringComparison.Ordinal) && Units.TryGetValue(s[p.Length..], out u) && u.Cat == UCat.Information)
                return (u, Math.Pow(2, e));
        foreach (var (p, e) in MetricPrefixes)
            if (s.Length > p.Length && s.StartsWith(p, StringComparison.Ordinal) && Units.TryGetValue(s[p.Length..], out u) && u.Prefix)
                return (u, Pow10(e * u.Power));
        throw Err(Errors.NA);
    }

    internal static double Convert(double v, string from, string to)
    {
        var (uf, mf) = FindUnit(from);
        var (ut, mt) = FindUnit(to);
        if (uf.Cat != ut.Cat) throw Err(Errors.NA);
        if (uf.Cat == UCat.Temperature)
            return TempBase(from) == TempBase(to) && mf == mt ? v : FromCelsius(ToCelsius(v, from, mf), to, mt);
        double ff = uf.Factor * mf, ft = ut.Factor * mt;
        return ff == ft ? v : v * ff / ft;
    }

    private static string TempBase(string unit) => unit switch
    {
        "cel" => "C",
        "fah" => "F",
        _ when unit.EndsWith("kel", StringComparison.Ordinal) || unit.EndsWith('K') => "K",
        _ => unit
    };

    private static double ToCelsius(double v, string unit, double mult) => TempBase(unit) switch
    {
        "C" => v,
        "F" => (v - 32) / 1.8,
        "Rank" => v / 1.8 - 273.15,
        "Reau" => v * 1.25,
        _ => v * mult - 273.15
    };

    private static double FromCelsius(double c, string unit, double mult) => TempBase(unit) switch
    {
        "C" => c,
        "F" => c * 1.8 + 32,
        "Rank" => (c + 273.15) * 1.8,
        "Reau" => c * 0.8,
        _ => (c + 273.15) / mult
    };
}
