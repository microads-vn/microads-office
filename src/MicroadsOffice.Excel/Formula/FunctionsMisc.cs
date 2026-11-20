using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

public static partial class Functions
{
    /// <summary>Math, statistics, text, date, information, database and web functions added for Excel parity.</summary>
    internal static void RegisterMisc()
    {
        // ---------------------------------------------------------- math & trig
        Reg(MiscMath.Trig(x => { if (x < 1) throw new FormulaError(Errors.Num); return Math.Acosh(x); }), "ACOSH");
        Reg(MiscMath.Trig(Math.Asinh), "ASINH");
        Reg(MiscMath.Trig(x => { if (Math.Abs(x) >= 1) throw new FormulaError(Errors.Num); return Math.Atanh(x); }), "ATANH");
        Reg(MiscMath.Trig(x => Math.PI / 2 - Math.Atan(x)), "ACOT");
        Reg(MiscMath.Trig(x => { if (Math.Abs(x) <= 1) throw new FormulaError(Errors.Num); return 0.5 * Math.Log((x + 1) / (x - 1)); }), "ACOTH");
        Reg(MiscMath.Recip(Math.Tan, true), "COT");
        Reg(MiscMath.Recip(Math.Tanh, true), "COTH");
        Reg(MiscMath.Recip(Math.Sin, true), "CSC");
        Reg(MiscMath.Recip(Math.Sinh, true), "CSCH");
        Reg(MiscMath.Recip(Math.Cos, false), "SEC");
        Reg(MiscMath.Recip(Math.Cosh, false), "SECH");
        Reg(MiscMath.Arabic, "ARABIC");
        Reg(MiscMath.Base, "BASE");
        Reg(MiscMath.Decimal, "DECIMAL");
        Reg(MiscMath.CombinA, "COMBINA");
        Reg(MiscMath.FactDouble, "FACTDOUBLE");
        Reg(MiscMath.MDeterm, "MDETERM");
        Reg(MiscMath.MInverse, "MINVERSE");
        Reg(MiscMath.MMult, "MMULT");
        Reg(MiscMath.MUnit, "MUNIT");
        Reg(MiscMath.Multinomial, "MULTINOMIAL");
        Reg(MiscMath.RandArray, "RANDARRAY");
        Reg(MiscMath.SeriesSum, "SERIESSUM");
        Reg((c, a) => MiscMath.SumPairs(c, a, (x, y) => x * x - y * y), "SUMX2MY2");
        Reg((c, a) => MiscMath.SumPairs(c, a, (x, y) => x * x + y * y), "SUMX2PY2");
        Reg((c, a) => MiscMath.SumPairs(c, a, (x, y) => (x - y) * (x - y)), "SUMXMY2");
        Reg(MiscMath.PercentOf, "PERCENTOF");
        Reg(MiscMath.Aggregate, "AGGREGATE");

        // ---------------------------------------------------------- statistics
        Reg((c, a) => MiscMath.Covariance(c, a, false), "COVARIANCE.P", "COVAR");
        Reg((c, a) => MiscMath.Covariance(c, a, true), "COVARIANCE.S");
        Reg(MiscMath.DevSq, "DEVSQ");
        Reg(MiscMath.Frequency, "FREQUENCY");
        Reg(MiscMath.Kurt, "KURT");
        Reg((c, a) => MiscMath.Skew(c, a, false), "SKEW");
        Reg((c, a) => MiscMath.Skew(c, a, true), "SKEW.P");
        Reg(MiscMath.ModeMult, "MODE.MULT");
        Reg((c, a) => MiscMath.PercentRank(c, a, true), "PERCENTRANK.INC", "PERCENTRANK");
        Reg((c, a) => MiscMath.PercentRank(c, a, false), "PERCENTRANK.EXC");
        Reg(MiscMath.QuartileExc, "QUARTILE.EXC");
        Reg(MiscMath.Rsq, "RSQ");
        Reg(MiscMath.Steyx, "STEYX");
        Reg(MiscMath.TrimMean, "TRIMMEAN");
        Reg((c, a) => CellValue.Number(StDev(Nums(c, a, true), false)), "STDEVPA");
        Reg((c, a) => CellValue.Number(Var(Nums(c, a, true), true)), "VARA");
        Reg((c, a) => CellValue.Number(Var(Nums(c, a, true), false)), "VARPA");
        Reg((c, a) => MiscMath.Linest(c, a, false), "LINEST");
        Reg((c, a) => MiscMath.Linest(c, a, true), "LOGEST");
        Reg((c, a) => MiscMath.Trend(c, a, false), "TREND");
        Reg((c, a) => MiscMath.Trend(c, a, true), "GROWTH");

        // ---------------------------------------------------------- text
        Reg(MiscMath.ArrayToText, "ARRAYTOTEXT");
        Reg(MiscMath.ValueToText, "VALUETOTEXT");
        Reg(MiscMath.TextSplit, "TEXTSPLIT");
        Reg(MiscMath.RegexTest, "REGEXTEST");
        Reg(MiscMath.RegexExtract, "REGEXEXTRACT");
        Reg(MiscMath.RegexReplace, "REGEXREPLACE");
        Reg((c, a) => { MiscMath.Args(a, 1, 1); return CellValue.Text(MiscMath.ToHalfWidth(c.Str(a[0]))); }, "ASC");
        Reg((c, a) => { MiscMath.Args(a, 1, 1); return CellValue.Text(MiscMath.ToFullWidth(c.Str(a[0]))); }, "DBCS", "JIS");
        Reg(MiscMath.BahtText, "BAHTTEXT");
        // single-byte locale: the byte versions behave exactly like the character versions
        foreach (var (b, ch) in new[] { ("FINDB", "FIND"), ("SEARCHB", "SEARCH"), ("LEFTB", "LEFT"), ("RIGHTB", "RIGHT"), ("MIDB", "MID"), ("LENB", "LEN"), ("REPLACEB", "REPLACE") })
            Reg(Table[ch], b);
        Reg(MiscMath.Phonetic, "PHONETIC");

        // ---------------------------------------------------------- date, information
        Reg(MiscMath.NetworkDaysIntl, "NETWORKDAYS.INTL");
        Reg(MiscMath.WorkdayIntl, "WORKDAY.INTL");
        Reg(MiscMath.Info, "INFO");

        // ---------------------------------------------------------- database
        foreach (var name in new[] { "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP" })
        {
            var n = name;
            Reg((c, a) => MiscMath.Database(c, a, n), n);
        }

        // ---------------------------------------------------------- web
        Reg((c, a) => { MiscMath.Args(a, 1, 1); return CellValue.Text(Uri.EscapeDataString(c.Str(a[0]))); }, "ENCODEURL");
        Reg(MiscMath.FilterXml, "FILTERXML");
    }
}

/// <summary>Implementation helpers of <see cref="Functions.RegisterMisc"/> (kept out of the partial class to avoid name clashes).</summary>
internal static class MiscMath
{
    private static CellValue N(double d) => CellValue.Number(d);
    private static CellValue S(string s) => CellValue.Text(s);
    private static FormulaError Err(string code) => new(code);
    private static readonly CellValue NA = CellValue.Error(Errors.NA);
    private const double TwoPow27 = 134217728;
    private const double TwoPow53 = 9007199254740992;

    internal static void Args(Expr[] a, int min, int max = int.MaxValue) { if (a.Length < min || a.Length > max) throw Err(Errors.Value); }
    private static bool Has(Expr[] a, int i) => i < a.Length && a[i] is not MissingExpr;
    private static double Trunc(double x) => Math.Truncate(x);

    // ================================================================== MATH & TRIG

    internal static FnImpl Trig(Func<double, double> f) => (c, a) => { Args(a, 1, 1); return N(f(c.Num(a[0]))); };

    /// <summary>1/f(x) with Excel's limits: |x| &lt; 2^27, #DIV/0! where f(x) = 0 at x = 0.</summary>
    internal static FnImpl Recip(Func<double, double> f, bool zeroIsDiv0) => (c, a) =>
    {
        Args(a, 1, 1);
        double x = c.Num(a[0]);
        if (Math.Abs(x) >= TwoPow27) throw Err(Errors.Num);
        if (zeroIsDiv0 && x == 0) throw Err(Errors.Div0);
        double v = f(x);
        if (v == 0) throw Err(Errors.Div0);
        return N(1 / v);
    };

    internal static object Arabic(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        var s = c.Str(a[0]).Trim();
        if (s.Length > 255) throw Err(Errors.Value);
        bool neg = s.StartsWith('-');
        if (neg) s = s[1..];
        double total = 0; int prev = 0;
        for (int i = s.Length - 1; i >= 0; i--)
        {
            int v = char.ToUpperInvariant(s[i]) switch
            {
                'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50, 'C' => 100, 'D' => 500, 'M' => 1000,
                _ => throw Err(Errors.Value)
            };
            if (v < prev) total -= v; else { total += v; prev = v; }
        }
        return N(neg ? -total : total);
    }

    internal static object Base(Ctx c, Expr[] a)
    {
        Args(a, 2, 3);
        double num = Trunc(c.Num(a[0])), radix = Trunc(c.Num(a[1])), minLen = Trunc(c.NumOr(a, 2, 0));
        if (num < 0 || num >= TwoPow53 || radix < 2 || radix > 36 || minLen < 0 || minLen > 255) throw Err(Errors.Num);
        const string digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        long n = (long)num; int r = (int)radix;
        var sb = new StringBuilder();
        do { sb.Insert(0, digits[(int)(n % r)]); n /= r; } while (n > 0);
        if (sb.Length < minLen) sb.Insert(0, new string('0', (int)minLen - sb.Length));
        return S(sb.ToString());
    }

    internal static object Decimal(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var s = c.Str(a[0]).Trim();
        double radix = Trunc(c.Num(a[1]));
        if (s.Length > 255) throw Err(Errors.Value);
        if (radix < 2 || radix > 36) throw Err(Errors.Num);
        double v = 0;
        foreach (var ch in s)
        {
            int d = ch >= '0' && ch <= '9' ? ch - '0' : char.IsAsciiLetter(ch) ? char.ToUpperInvariant(ch) - 'A' + 10 : 99;
            if (d >= radix) throw Err(Errors.Num);
            v = v * radix + d;
        }
        if (v >= TwoPow53) throw Err(Errors.Num);
        return N(v);
    }

    private static double Binom(double n, double k)
    {
        k = Math.Min(k, n - k);
        double r = 1;
        for (int i = 1; i <= k; i++) r = r * (n - k + i) / i;
        return Math.Round(r);
    }

    internal static object CombinA(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        double n = Trunc(c.Num(a[0])), k = Trunc(c.Num(a[1]));
        if (n < 0 || k < 0 || n < k) throw Err(Errors.Num);
        if (n == 0) return N(1);
        return N(Binom(n + k - 1, k));
    }

    internal static object FactDouble(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        double n = Trunc(c.Num(a[0]));
        if (n < 0) throw Err(Errors.Num);
        double f = 1;
        for (double i = n; i > 1; i -= 2) { f *= i; if (double.IsInfinity(f)) throw Err(Errors.Num); }
        return N(f);
    }

    /// <summary>A numeric matrix: every cell must be a number (text, logicals and blanks give #VALUE!).</summary>
    private static double[,] Matrix(Ctx c, Expr e)
    {
        var g = Functions.Grid(c, e);
        int r = g.GetLength(0), k = g.GetLength(1);
        var m = new double[r, k];
        for (int i = 0; i < r; i++)
            for (int j = 0; j < k; j++)
            {
                var v = g[i, j];
                if (v.IsError) throw Err(v.Str!);
                if (!v.IsNumber) throw Err(Errors.Value);
                m[i, j] = v.Num;
            }
        return m;
    }

    private static CellValue[,] ToCells(double[,] m)
    {
        var res = new CellValue[m.GetLength(0), m.GetLength(1)];
        for (int i = 0; i < m.GetLength(0); i++) for (int j = 0; j < m.GetLength(1); j++) res[i, j] = N(m[i, j]);
        return res;
    }

    internal static object MDeterm(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        var m = Matrix(c, a[0]);
        int n = m.GetLength(0);
        if (n != m.GetLength(1)) throw Err(Errors.Value);
        double det = 1;
        for (int col = 0; col < n; col++)
        {
            int p = col;
            for (int r = col + 1; r < n; r++) if (Math.Abs(m[r, col]) > Math.Abs(m[p, col])) p = r;
            if (m[p, col] == 0) return N(0);
            if (p != col) { for (int k = 0; k < n; k++) (m[p, k], m[col, k]) = (m[col, k], m[p, k]); det = -det; }
            det *= m[col, col];
            for (int r = col + 1; r < n; r++)
            {
                double f = m[r, col] / m[col, col];
                if (f == 0) continue;
                for (int k = col; k < n; k++) m[r, k] -= f * m[col, k];
            }
        }
        return N(det);
    }

    internal static object MInverse(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        var m = Matrix(c, a[0]);
        int n = m.GetLength(0);
        if (n != m.GetLength(1)) throw Err(Errors.Value);
        double scale = 0;
        foreach (var x in m) scale = Math.Max(scale, Math.Abs(x));
        var inv = new double[n, n];
        for (int i = 0; i < n; i++) inv[i, i] = 1;
        for (int col = 0; col < n; col++)
        {
            int p = col;
            for (int r = col + 1; r < n; r++) if (Math.Abs(m[r, col]) > Math.Abs(m[p, col])) p = r;
            if (Math.Abs(m[p, col]) <= 1e-15 * scale) throw Err(Errors.Num);
            if (p != col)
                for (int k = 0; k < n; k++) { (m[p, k], m[col, k]) = (m[col, k], m[p, k]); (inv[p, k], inv[col, k]) = (inv[col, k], inv[p, k]); }
            double d = m[col, col];
            for (int k = 0; k < n; k++) { m[col, k] /= d; inv[col, k] /= d; }
            for (int r = 0; r < n; r++)
            {
                if (r == col) continue;
                double f = m[r, col];
                if (f == 0) continue;
                for (int k = 0; k < n; k++) { m[r, k] -= f * m[col, k]; inv[r, k] -= f * inv[col, k]; }
            }
        }
        return new ArrayVal(ToCells(inv));
    }

    internal static object MMult(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var x = Matrix(c, a[0]); var y = Matrix(c, a[1]);
        int n = x.GetLength(0), m = x.GetLength(1), p = y.GetLength(1);
        if (m != y.GetLength(0)) throw Err(Errors.Value);
        var res = new double[n, p];
        for (int i = 0; i < n; i++)
            for (int j = 0; j < p; j++)
            {
                double s = 0;
                for (int k = 0; k < m; k++) s += x[i, k] * y[k, j];
                res[i, j] = s;
            }
        return new ArrayVal(ToCells(res));
    }

    internal static object MUnit(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        double n = Trunc(c.Num(a[0]));
        if (n < 1) throw Err(Errors.Value);
        if (n * n > 4_000_000) throw Err(Errors.Num);
        var res = new CellValue[(int)n, (int)n];
        for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) res[i, j] = N(i == j ? 1 : 0);
        return new ArrayVal(res);
    }

    internal static object Multinomial(Ctx c, Expr[] a)
    {
        Args(a, 1);
        double result = 1, sum = 0;
        foreach (var raw in Functions.Nums(c, a))
        {
            double v = Trunc(raw);
            if (v < 0) throw Err(Errors.Num);
            for (int i = 1; i <= v; i++)
            {
                sum++;
                result = result * sum / i;
                if (double.IsInfinity(result)) throw Err(Errors.Num);
            }
        }
        return N(result < TwoPow53 ? Math.Round(result) : result);
    }

    internal static object RandArray(Ctx c, Expr[] a)
    {
        Args(a, 0, 5);
        double rows = Trunc(c.NumOr(a, 0, 1)), cols = Trunc(c.NumOr(a, 1, 1));
        double min = c.NumOr(a, 2, 0), max = c.NumOr(a, 3, 1);
        bool whole = c.BoolOr(a, 4, false);
        if (rows < 0 || cols < 0 || min > max) throw Err(Errors.Value);
        if (rows == 0 || cols == 0) throw Err(Errors.Calc);
        if (whole && (min != Math.Floor(min) || max != Math.Floor(max))) throw Err(Errors.Value);
        if (rows * cols > 4_000_000) throw Err(Errors.Num);
        var res = new CellValue[(int)rows, (int)cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
            {
                double r = Random.Shared.NextDouble();
                res[i, j] = N(whole ? Math.Min(max, Math.Floor(min + r * (max - min + 1))) : min + r * (max - min));
            }
        return new ArrayVal(res);
    }

    internal static object SeriesSum(Ctx c, Expr[] a)
    {
        Args(a, 4, 4);
        double x = c.Num(a[0]), n = c.Num(a[1]), m = c.Num(a[2]);
        double s = 0; int i = 0;
        foreach (var v in Functions.Grid(c, a[3]))
        {
            if (v.IsError) throw Err(v.Str!);
            if (!v.IsNumber) throw Err(Errors.Value);
            s += v.Num * Math.Pow(x, n + i * m);
            i++;
        }
        return N(s);
    }

    /// <summary>Numeric pairs of two equally sized arrays (pairs with a non-number are skipped, errors propagate).</summary>
    private static (List<double> X, List<double> Y) Pairs(Ctx c, Expr ea, Expr eb)
    {
        var A = Functions.Grid(c, ea); var B = Functions.Grid(c, eb);
        if (A.Length != B.Length) throw Err(Errors.NA);
        var fa = A.Cast<CellValue>().ToList(); var fb = B.Cast<CellValue>().ToList();
        var x = new List<double>(); var y = new List<double>();
        for (int i = 0; i < fa.Count; i++)
        {
            if (fa[i].IsError) throw Err(fa[i].Str!);
            if (fb[i].IsError) throw Err(fb[i].Str!);
            if (fa[i].IsNumber && fb[i].IsNumber) { x.Add(fa[i].Num); y.Add(fb[i].Num); }
        }
        return (x, y);
    }

    internal static object SumPairs(Ctx c, Expr[] a, Func<double, double, double> f)
    {
        Args(a, 2, 2);
        var (x, y) = Pairs(c, a[0], a[1]);
        if (x.Count == 0) throw Err(Errors.Div0);
        double s = 0;
        for (int i = 0; i < x.Count; i++) s += f(x[i], y[i]);
        return N(s);
    }

    internal static object PercentOf(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        double part = Functions.Nums(c, [a[0]]).Sum(), all = Functions.Nums(c, [a[1]]).Sum();
        if (all == 0) throw Err(Errors.Div0);
        return N(part / all);
    }

    private static readonly Regex NestedAgg = new(@"\b(SUBTOTAL|AGGREGATE)\s*\(", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>Values of an AGGREGATE argument, honouring the hidden-row and nested SUBTOTAL/AGGREGATE options.</summary>
    private static void AggCollect(Ctx c, object v, bool ignoreHidden, bool ignoreNested, List<CellValue> into)
    {
        if (v is RangeVal r)
        {
            var rg = c.Calc.Clip(r);
            for (int row = rg.R1; row <= rg.R2; row++)
            {
                if (ignoreHidden && (r.Sheet.HiddenRows.Contains(row) || r.Sheet.FilteredRows.Contains(row))) continue;
                for (int col = rg.C1; col <= rg.C2; col++)
                {
                    var f = r.Sheet.GetCell(row, col)?.Formula;
                    if (ignoreNested && f != null && NestedAgg.IsMatch(f)) continue;
                    into.Add(c.Calc.ValueOf(r.Sheet, row, col));
                }
            }
            return;
        }
        into.AddRange(Functions.FlatVal(c, v));
    }

    internal static object Aggregate(Ctx c, Expr[] a)
    {
        Args(a, 3);
        int fn = (int)Trunc(c.Num(a[0]));
        int opt = Has(a, 1) ? (int)Trunc(c.Num(a[1])) : 0;
        if (fn < 1 || fn > 19 || opt < 0 || opt > 7) throw Err(Errors.Value);
        bool ignoreNested = opt <= 3, ignoreHidden = opt is 1 or 3 or 5 or 7, ignoreErrors = opt is 2 or 3 or 6 or 7;
        var vals = new List<CellValue>();
        double k = 0;
        if (fn >= 14)
        {
            // array form: AGGREGATE(fn, options, array, k)
            if (a.Length != 4 || a[3] is MissingExpr) throw Err(Errors.Value);
            AggCollect(c, c.Eval(a[2]), ignoreHidden, ignoreNested, vals);
            k = c.Num(a[3]);
        }
        else
        {
            // reference form: AGGREGATE(fn, options, ref1, [ref2], ...) – references only
            for (int i = 2; i < a.Length; i++)
            {
                if (a[i] is MissingExpr) continue;
                if (c.Eval(a[i]) is not RangeVal r) throw Err(Errors.Value);
                AggCollect(c, r, ignoreHidden, ignoreNested, vals);
            }
        }
        var nums = new List<double>();
        int counta = 0;
        foreach (var v in vals)
        {
            if (v.IsError)
            {
                if (ignoreErrors) continue;
                if (fn == 3) { counta++; continue; }
                if (fn == 2) continue;
                throw Err(v.Str!);
            }
            if (v.IsNumber) nums.Add(v.Num);
            if (!v.IsBlank) counta++;
        }
        switch (fn)
        {
            case 1: if (nums.Count == 0) throw Err(Errors.Div0); return N(nums.Average());
            case 2: return N(nums.Count);
            case 3: return N(counta);
            case 4: return N(nums.Count == 0 ? 0 : nums.Max());
            case 5: return N(nums.Count == 0 ? 0 : nums.Min());
            case 6: return N(nums.Count == 0 ? 0 : nums.Aggregate(1.0, (x, y) => x * y));
            case 7: return N(Functions.StDev(nums, true));
            case 8: return N(Functions.StDev(nums, false));
            case 9: return N(nums.Sum());
            case 10: return N(Functions.Var(nums, true));
            case 11: return N(Functions.Var(nums, false));
            case 12:
            {
                if (nums.Count == 0) throw Err(Errors.Num);
                nums.Sort();
                int n = nums.Count;
                return N(n % 2 == 1 ? nums[n / 2] : (nums[n / 2 - 1] + nums[n / 2]) / 2);
            }
            case 13:
            {
                var best = nums.Select((x, i) => (x, i)).GroupBy(t => t.x).Where(g => g.Count() > 1)
                    .OrderByDescending(g => g.Count()).ThenBy(g => g.First().i).FirstOrDefault();
                if (best == null) throw Err(Errors.NA);
                return N(best.Key);
            }
            case 14:
            case 15:
            {
                int kk = (int)Math.Ceiling(k);
                if (kk < 1 || kk > nums.Count) throw Err(Errors.Num);
                nums.Sort();
                return N(fn == 15 ? nums[kk - 1] : nums[nums.Count - kk]);
            }
            case 16: return N(Functions.Percentile(nums, k));
            case 18: return N(Functions.Percentile(nums, k, false));
            case 17:
            {
                int q = (int)Trunc(k);
                if (q < 0 || q > 4) throw Err(Errors.Num);
                return N(Functions.Percentile(nums, q / 4.0));
            }
            default:
            {
                int q = (int)Trunc(k);
                if (q < 1 || q > 3) throw Err(Errors.Num);
                return N(Functions.Percentile(nums, q / 4.0, false));
            }
        }
    }

    // ================================================================== STATISTICS

    internal static object Covariance(Ctx c, Expr[] a, bool sample)
    {
        Args(a, 2, 2);
        var (x, y) = Pairs(c, a[0], a[1]);
        int n = x.Count;
        if (n == 0 || (sample && n < 2)) throw Err(Errors.Div0);
        double mx = x.Average(), my = y.Average(), s = 0;
        for (int i = 0; i < n; i++) s += (x[i] - mx) * (y[i] - my);
        return N(s / (sample ? n - 1 : n));
    }

    internal static object DevSq(Ctx c, Expr[] a)
    {
        var l = Functions.Nums(c, a);
        if (l.Count == 0) throw Err(Errors.Num);
        double m = l.Average();
        return N(l.Sum(x => (x - m) * (x - m)));
    }

    /// <summary>Numbers of one argument: ranges/arrays contribute numbers only, errors propagate.</summary>
    private static List<double> NumsOf(Ctx c, Expr e)
    {
        var res = new List<double>();
        foreach (var v in Functions.FlatVal(c, c.Eval(e)))
        {
            if (v.IsError) throw Err(v.Str!);
            if (v.IsNumber) res.Add(v.Num);
        }
        return res;
    }

    internal static object Frequency(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var data = NumsOf(c, a[0]);
        var bins = NumsOf(c, a[1]);
        if (bins.Count == 0) return new ArrayVal(new[,] { { N(data.Count) } });
        var order = Enumerable.Range(0, bins.Count).OrderBy(i => bins[i]).ToList();
        var counts = new double[bins.Count + 1];
        foreach (var x in data)
        {
            int slot = bins.Count;
            foreach (var i in order) if (x <= bins[i]) { slot = i; break; }
            counts[slot]++;
        }
        var res = new CellValue[bins.Count + 1, 1];
        for (int i = 0; i <= bins.Count; i++) res[i, 0] = N(counts[i]);
        return new ArrayVal(res);
    }

    private static (List<double> L, double Mean, double Sd) Moments(Ctx c, Expr[] a, bool sample)
    {
        var l = Functions.Nums(c, a);
        if (l.Count == 0) throw Err(Errors.Div0);
        double m = l.Average();
        double ss = l.Sum(x => (x - m) * (x - m));
        int den = sample ? l.Count - 1 : l.Count;
        if (den <= 0) throw Err(Errors.Div0);
        double sd = Math.Sqrt(ss / den);
        if (sd == 0) throw Err(Errors.Div0);
        return (l, m, sd);
    }

    internal static object Kurt(Ctx c, Expr[] a)
    {
        var (l, m, s) = Moments(c, a, true);
        double n = l.Count;
        if (n < 4) throw Err(Errors.Div0);
        double s4 = l.Sum(x => Math.Pow((x - m) / s, 4));
        return N(n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * s4 - 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3)));
    }

    internal static object Skew(Ctx c, Expr[] a, bool population)
    {
        var (l, m, s) = Moments(c, a, !population);
        double n = l.Count;
        double s3 = l.Sum(x => Math.Pow((x - m) / s, 3));
        if (population) return N(s3 / n);
        if (n < 3) throw Err(Errors.Div0);
        return N(n / ((n - 1) * (n - 2)) * s3);
    }

    internal static object ModeMult(Ctx c, Expr[] a)
    {
        Args(a, 1);
        var l = Functions.Nums(c, a);
        var groups = l.Select((x, i) => (x, i)).GroupBy(t => t.x).Select(g => (Value: g.Key, Count: g.Count(), First: g.First().i)).ToList();
        int max = groups.Count == 0 ? 0 : groups.Max(g => g.Count);
        if (max < 2) throw Err(Errors.NA);
        var modes = groups.Where(g => g.Count == max).OrderBy(g => g.First).ToList();
        var res = new CellValue[modes.Count, 1];
        for (int i = 0; i < modes.Count; i++) res[i, 0] = N(modes[i].Value);
        return new ArrayVal(res);
    }

    internal static object PercentRank(Ctx c, Expr[] a, bool inclusive)
    {
        Args(a, 2, 3);
        var l = NumsOf(c, a[0]);
        double x = c.Num(a[1]);
        double sig = Trunc(c.NumOr(a, 2, 3));
        if (sig < 1 || l.Count == 0) throw Err(Errors.Num);
        l.Sort();
        int n = l.Count;
        if (x < l[0] || x > l[n - 1]) throw Err(Errors.NA);
        int less = l.Count(v => v < x);
        double rank;
        if (l.Contains(x)) rank = less;
        else
        {
            int i = less - 1;
            rank = i + (x - l[i]) / (l[i + 1] - l[i]);
        }
        double r = inclusive ? (n == 1 ? 1 : rank / (n - 1)) : (rank + 1) / (n + 1);
        double f = Math.Pow(10, sig);
        return N(Math.Floor(Math.Round(r * f, 6)) / f);
    }

    internal static object QuartileExc(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var l = NumsOf(c, a[0]);
        int q = (int)Trunc(c.Num(a[1]));
        if (q < 1 || q > 3) throw Err(Errors.Num);
        return N(Functions.Percentile(l, q / 4.0, false));
    }

    private static (double Sxx, double Syy, double Sxy, int N) Sums(List<double> x, List<double> y)
    {
        double mx = x.Average(), my = y.Average(), sxx = 0, syy = 0, sxy = 0;
        for (int i = 0; i < x.Count; i++)
        {
            double dx = x[i] - mx, dy = y[i] - my;
            sxx += dx * dx; syy += dy * dy; sxy += dx * dy;
        }
        return (sxx, syy, sxy, x.Count);
    }

    internal static object Rsq(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var (y, x) = Pairs(c, a[0], a[1]);
        if (x.Count < 2) throw Err(Errors.Div0);
        var (sxx, syy, sxy, _) = Sums(x, y);
        if (sxx == 0 || syy == 0) throw Err(Errors.Div0);
        return N(sxy * sxy / (sxx * syy));
    }

    internal static object Steyx(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var (y, x) = Pairs(c, a[0], a[1]);
        if (x.Count < 3) throw Err(Errors.Div0);
        var (sxx, syy, sxy, n) = Sums(x, y);
        if (sxx == 0) throw Err(Errors.Div0);
        return N(Math.Sqrt(Math.Max(0, syy - sxy * sxy / sxx) / (n - 2)));
    }

    internal static object TrimMean(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var l = NumsOf(c, a[0]);
        double p = c.Num(a[1]);
        if (p < 0 || p >= 1 || l.Count == 0) throw Err(Errors.Num);
        l.Sort();
        int k = (int)Math.Floor(l.Count * p / 2);
        return N(l.Skip(k).Take(l.Count - 2 * k).Average());
    }

    // ------------------------------------------------------------------ regression (LINEST, LOGEST, TREND, GROWTH)

    private enum Layout { Column, Row, Grid }

    private sealed class RegData
    {
        public double[] Y = [];
        public double[][] X = [];   // X[variable][observation]
        public Layout Layout;
        public int YRows, YCols;
        public CellValue[,]? XGrid;
    }

    private static double[,] NumGrid(CellValue[,] g)
    {
        var m = new double[g.GetLength(0), g.GetLength(1)];
        for (int i = 0; i < g.GetLength(0); i++)
            for (int j = 0; j < g.GetLength(1); j++)
            {
                var v = g[i, j];
                if (v.IsError) throw Err(v.Str!);
                if (!v.IsNumber) throw Err(Errors.Value);
                m[i, j] = v.Num;
            }
        return m;
    }

    private static RegData Prepare(Ctx c, Expr[] a, bool log)
    {
        var yg = Functions.Grid(c, a[0]);
        var ym = NumGrid(yg);
        int ry = ym.GetLength(0), cy = ym.GetLength(1), n = ry * cy;
        var d = new RegData { YRows = ry, YCols = cy, Y = new double[n] };
        for (int i = 0; i < ry; i++) for (int j = 0; j < cy; j++) d.Y[i * cy + j] = ym[i, j];
        if (log)
        {
            if (d.Y.Any(v => v <= 0)) throw Err(Errors.Num);
            for (int i = 0; i < n; i++) d.Y[i] = Math.Log(d.Y[i]);
        }
        if (!Has(a, 1))
        {
            d.Layout = cy == 1 ? Layout.Column : ry == 1 ? Layout.Row : Layout.Grid;
            d.X = [Enumerable.Range(1, n).Select(i => (double)i).ToArray()];
            return d;
        }
        d.XGrid = Functions.Grid(c, a[1]);
        var xm = NumGrid(d.XGrid);
        int rx = xm.GetLength(0), cx = xm.GetLength(1);
        if (cy == 1)
        {
            if (rx != ry) throw Err(Errors.Ref);
            d.Layout = Layout.Column;
            d.X = Enumerable.Range(0, cx).Select(j => Enumerable.Range(0, ry).Select(i => xm[i, j]).ToArray()).ToArray();
        }
        else if (ry == 1)
        {
            if (cx != cy) throw Err(Errors.Ref);
            d.Layout = Layout.Row;
            d.X = Enumerable.Range(0, rx).Select(j => Enumerable.Range(0, cy).Select(i => xm[j, i]).ToArray()).ToArray();
        }
        else
        {
            if (rx != ry || cx != cy) throw Err(Errors.Ref);
            d.Layout = Layout.Grid;
            var flat = new double[n];
            for (int i = 0; i < ry; i++) for (int j = 0; j < cy; j++) flat[i * cy + j] = xm[i, j];
            d.X = [flat];
        }
        return d;
    }

    private sealed class Fit
    {
        public double[] Coef = [], Se = [];
        public double Intercept, SeIntercept = double.NaN, R2, SeY, F, Df, SsReg, SsResid;
    }

    private static double Dot(double[] a, double[] b) { double s = 0; for (int i = 0; i < a.Length; i++) s += a[i] * b[i]; return s; }

    /// <summary>
    /// Least squares by QR (classical Gram-Schmidt with re-orthogonalisation) on centred data when there is a constant.
    /// Collinear (redundant) columns get coefficient 0 and standard error 0, like Excel.
    /// </summary>
    private static Fit Regress(double[][] x, double[] y, bool constant)
    {
        int n = y.Length, k = x.Length;
        double ym = constant ? y.Average() : 0;
        var xm = new double[k];
        var q = new List<double[]>();
        var rcols = new List<(double[] R, double Diag)>();
        var kept = new List<int>();
        for (int j = 0; j < k; j++)
        {
            if (constant) xm[j] = x[j].Average();
            var v = new double[n];
            for (int i = 0; i < n; i++) v[i] = x[j][i] - xm[j];
            double scale = Math.Sqrt(Dot(x[j], x[j]));
            var r = new double[q.Count];
            for (int pass = 0; pass < 2; pass++)
                for (int t = 0; t < q.Count; t++)
                {
                    double dd = Dot(q[t], v);
                    r[t] += dd;
                    for (int i = 0; i < n; i++) v[i] -= dd * q[t][i];
                }
            double nv = Math.Sqrt(Dot(v, v));
            if (scale == 0 || nv <= 1e-10 * scale) continue; // redundant column
            for (int i = 0; i < n; i++) v[i] /= nv;
            q.Add(v); rcols.Add((r, nv)); kept.Add(j);
        }
        int p = kept.Count;
        // Qᵀ y (two passes)
        var res = new double[n];
        for (int i = 0; i < n; i++) res[i] = y[i] - ym;
        var qy = new double[p];
        for (int pass = 0; pass < 2; pass++)
            for (int t = 0; t < p; t++)
            {
                double dd = Dot(q[t], res);
                qy[t] += dd;
                for (int i = 0; i < n; i++) res[i] -= dd * q[t][i];
            }
        // R (p×p upper triangular) and its inverse
        var R = new double[p, p];
        for (int s = 0; s < p; s++) { for (int t = 0; t < s; t++) R[t, s] = rcols[s].R[t]; R[s, s] = rcols[s].Diag; }
        var b = new double[p];
        for (int s = p - 1; s >= 0; s--)
        {
            double v = qy[s];
            for (int t = s + 1; t < p; t++) v -= R[s, t] * b[t];
            b[s] = v / R[s, s];
        }
        var Ri = new double[p, p];
        for (int col = 0; col < p; col++)
            for (int s = col; s >= 0; s--)
            {
                double v = s == col ? 1 : 0;
                for (int t = s + 1; t <= col; t++) v -= R[s, t] * Ri[t, col];
                Ri[s, col] = v / R[s, s];
            }
        var fit = new Fit { Coef = new double[k], Se = new double[k] };
        for (int s = 0; s < p; s++) fit.Coef[kept[s]] = b[s];
        fit.Intercept = constant ? ym - Enumerable.Range(0, k).Sum(j => fit.Coef[j] * xm[j]) : 0;
        double ssRes = 0, ssReg = 0;
        for (int i = 0; i < n; i++)
        {
            double yh = fit.Intercept;
            for (int j = 0; j < k; j++) yh += fit.Coef[j] * x[j][i];
            double e = y[i] - yh;
            ssRes += e * e;
            double dReg = constant ? yh - ym : yh;
            ssReg += dReg * dReg;
        }
        fit.SsResid = ssRes; fit.SsReg = ssReg;
        fit.Df = n - p - (constant ? 1 : 0);
        double sigma2 = fit.Df > 0 ? ssRes / fit.Df : double.NaN;
        fit.SeY = Math.Sqrt(sigma2);
        fit.R2 = ssReg + ssRes == 0 ? double.NaN : ssReg / (ssReg + ssRes);
        fit.F = p == 0 || fit.Df <= 0 || ssRes == 0 ? double.NaN : ssReg / p / (ssRes / fit.Df);
        // covariance of the slopes: sigma² · R⁻¹ R⁻ᵀ
        var cov = new double[p, p];
        for (int s = 0; s < p; s++)
            for (int t = 0; t < p; t++)
            {
                double v = 0;
                for (int u = Math.Max(s, t); u < p; u++) v += Ri[s, u] * Ri[t, u];
                cov[s, t] = v;
            }
        for (int s = 0; s < p; s++) fit.Se[kept[s]] = Math.Sqrt(sigma2 * cov[s, s]);
        if (constant)
        {
            double v = 1.0 / n;
            for (int s = 0; s < p; s++) for (int t = 0; t < p; t++) v += xm[kept[s]] * cov[s, t] * xm[kept[t]];
            fit.SeIntercept = Math.Sqrt(sigma2 * v);
        }
        return fit;
    }

    internal static object Linest(Ctx c, Expr[] a, bool log)
    {
        Args(a, 1, 4);
        var d = Prepare(c, a, log);
        bool constant = c.BoolOr(a, 2, true), stats = c.BoolOr(a, 3, false);
        var fit = Regress(d.X, d.Y, constant);
        int k = d.X.Length, cols = k + 1;
        Func<double, double> tr = log ? Math.Exp : v => v;
        var res = new CellValue[stats ? 5 : 1, cols];
        for (int j = 0; j < k; j++) res[0, j] = N(tr(fit.Coef[k - 1 - j]));
        res[0, k] = N(tr(fit.Intercept));
        if (stats)
        {
            for (int r = 1; r < 5; r++) for (int j = 0; j < cols; j++) res[r, j] = NA;
            for (int j = 0; j < k; j++) res[1, j] = N(fit.Se[k - 1 - j]);
            res[1, k] = constant ? N(fit.SeIntercept) : NA;
            res[2, 0] = N(fit.R2); res[2, 1] = N(fit.SeY);
            res[3, 0] = N(fit.F); res[3, 1] = N(fit.Df);
            res[4, 0] = N(fit.SsReg); res[4, 1] = N(fit.SsResid);
        }
        return new ArrayVal(res);
    }

    internal static object Trend(Ctx c, Expr[] a, bool log)
    {
        Args(a, 1, 4);
        var d = Prepare(c, a, log);
        bool constant = c.BoolOr(a, 3, true);
        var fit = Regress(d.X, d.Y, constant);
        int k = d.X.Length;
        double[,] nx;
        if (Has(a, 2)) nx = NumGrid(Functions.Grid(c, a[2]));
        else if (d.XGrid != null) nx = NumGrid(d.XGrid);
        else
        {
            nx = new double[d.YRows, d.YCols];
            for (int i = 0; i < d.YRows; i++) for (int j = 0; j < d.YCols; j++) nx[i, j] = i * d.YCols + j + 1;
        }
        int nr = nx.GetLength(0), nc = nx.GetLength(1);
        double Pred(Func<int, double> xv)
        {
            double v = fit.Intercept;
            for (int j = 0; j < k; j++) v += fit.Coef[j] * xv(j);
            return log ? Math.Exp(v) : v;
        }
        CellValue[,] res;
        if (k == 1)
        {
            res = new CellValue[nr, nc];
            for (int i = 0; i < nr; i++) for (int j = 0; j < nc; j++) res[i, j] = N(Pred(_ => nx[i, j]));
        }
        else if (d.Layout == Layout.Column)
        {
            if (nc != k) throw Err(Errors.Ref);
            res = new CellValue[nr, 1];
            for (int i = 0; i < nr; i++) res[i, 0] = N(Pred(j => nx[i, j]));
        }
        else
        {
            if (nr != k) throw Err(Errors.Ref);
            res = new CellValue[1, nc];
            for (int i = 0; i < nc; i++) res[0, i] = N(Pred(j => nx[j, i]));
        }
        return new ArrayVal(res);
    }

    // ================================================================== TEXT

    private static string ArrText(CellValue v, bool strict) => v.Kind switch
    {
        ValueKind.Text => strict ? "\"" + v.Str!.Replace("\"", "\"\"") + "\"" : v.Str!,
        ValueKind.Error => v.Str!,
        _ => Calc.ToText(v)
    };

    private static int Format01(Ctx c, Expr[] a, int i)
    {
        double f = Trunc(c.NumOr(a, i, 0));
        if (f != 0 && f != 1) throw Err(Errors.Value);
        return (int)f;
    }

    internal static object ArrayToText(Ctx c, Expr[] a)
    {
        Args(a, 1, 2);
        var g = Functions.Grid(c, a[0]);
        bool strict = Format01(c, a, 1) == 1;
        int rows = g.GetLength(0), cols = g.GetLength(1);
        if (!strict) return S(string.Join(", ", g.Cast<CellValue>().Select(v => ArrText(v, false))));
        var sb = new StringBuilder("{");
        for (int i = 0; i < rows; i++)
        {
            if (i > 0) sb.Append(';');
            for (int j = 0; j < cols; j++) { if (j > 0) sb.Append(','); sb.Append(ArrText(g[i, j], true)); }
        }
        return S(sb.Append('}').ToString());
    }

    /// <summary>Applies a scalar function element-wise when the argument is an array or a multi-cell range.</summary>
    private static object Lift(Ctx c, object v, Func<CellValue, CellValue> f)
    {
        if (v is ArrayVal or RangeVal { Range.IsSingle: false })
        {
            var g = Functions.Grid(c, v);
            var res = new CellValue[g.GetLength(0), g.GetLength(1)];
            for (int i = 0; i < g.GetLength(0); i++)
                for (int j = 0; j < g.GetLength(1); j++)
                {
                    try { res[i, j] = f(g[i, j]); } catch (FormulaError e) { res[i, j] = CellValue.Error(e.Code); }
                }
            return new ArrayVal(res);
        }
        return f(c.Calc.ToScalar(v, c));
    }

    internal static object ValueToText(Ctx c, Expr[] a)
    {
        Args(a, 1, 2);
        bool strict = Format01(c, a, 1) == 1;
        return Lift(c, c.Eval(a[0]), v => S(ArrText(v, strict)));
    }

    private static List<string> Delims(Ctx c, Expr[] a, int i)
    {
        var res = new List<string>();
        if (!Has(a, i)) return res;
        foreach (var v in Functions.FlatVal(c, c.Eval(a[i])))
        {
            var s = Calc.ToText(v);
            if (s.Length == 0) throw Err(Errors.Value);
            res.Add(s);
        }
        return res;
    }

    private static List<string> SplitBy(string s, List<string> delims, StringComparison cmp)
    {
        var parts = new List<string>();
        if (delims.Count == 0) { parts.Add(s); return parts; }
        int start = 0, i = 0;
        while (i < s.Length)
        {
            int len = 0;
            foreach (var d in delims)
                if (d.Length > len && i + d.Length <= s.Length && string.Compare(s, i, d, 0, d.Length, cmp) == 0) len = d.Length;
            if (len > 0) { parts.Add(s[start..i]); i += len; start = i; }
            else i++;
        }
        parts.Add(s[start..]);
        return parts;
    }

    internal static object TextSplit(Ctx c, Expr[] a)
    {
        Args(a, 2, 6);
        var text = c.Str(a[0]);
        var colD = Delims(c, a, 1);
        var rowD = Delims(c, a, 2);
        bool ignoreEmpty = c.BoolOr(a, 3, false);
        double mode = Trunc(c.NumOr(a, 4, 0));
        if (mode != 0 && mode != 1) throw Err(Errors.Value);
        var pad = Has(a, 5) ? c.Scalar(a[5]) : NA;
        if (colD.Count == 0 && rowD.Count == 0) throw Err(Errors.Value);
        var cmp = mode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var rows = SplitBy(text, rowD, cmp);
        if (ignoreEmpty) rows.RemoveAll(r => r.Length == 0);
        var grid = rows.Select(r =>
        {
            var cells = SplitBy(r, colD, cmp);
            if (ignoreEmpty) cells.RemoveAll(x => x.Length == 0);
            return cells;
        }).Where(r => r.Count > 0).ToList();
        if (grid.Count == 0) throw Err(Errors.Calc);
        int cols = grid.Max(r => r.Count);
        var res = new CellValue[grid.Count, cols];
        for (int i = 0; i < grid.Count; i++)
            for (int j = 0; j < cols; j++) res[i, j] = j < grid[i].Count ? S(grid[i][j]) : pad;
        return new ArrayVal(res);
    }

    private static readonly Dictionary<(string, bool), Regex> RegexCache = new();

    private static Regex MakeRegex(string pattern, Ctx c, Expr[] a, int caseArg)
    {
        double cs = Trunc(c.NumOr(a, caseArg, 0));
        if (cs != 0 && cs != 1) throw Err(Errors.Value);
        bool ignoreCase = cs == 1;
        lock (RegexCache)
        {
            if (RegexCache.TryGetValue((pattern, ignoreCase), out var rx)) return rx;
            try
            {
                rx = new Regex(pattern, (ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None) | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
            }
            catch (ArgumentException) { throw Err(Errors.Value); }
            if (RegexCache.Count > 200) RegexCache.Clear();
            RegexCache[(pattern, ignoreCase)] = rx;
            return rx;
        }
    }

    private static T Timed<T>(Func<T> f)
    {
        try { return f(); } catch (RegexMatchTimeoutException) { throw Err(Errors.Value); }
    }

    internal static object RegexTest(Ctx c, Expr[] a)
    {
        Args(a, 2, 3);
        var rx = MakeRegex(c.Str(a[1]), c, a, 2);
        return Lift(c, c.Eval(a[0]), v => CellValue.Bool(Timed(() => rx.IsMatch(Calc.ToText(v)))));
    }

    internal static object RegexExtract(Ctx c, Expr[] a)
    {
        Args(a, 2, 4);
        var text = c.Str(a[0]);
        var rx = MakeRegex(c.Str(a[1]), c, a, 3);
        double mode = Trunc(c.NumOr(a, 2, 0));
        if (mode < 0 || mode > 2) throw Err(Errors.Value);
        if (mode == 1)
        {
            var all = Timed(() => rx.Matches(text).Select(m => m.Value).ToList());
            if (all.Count == 0) throw Err(Errors.NA);
            var res = new CellValue[all.Count, 1];
            for (int i = 0; i < all.Count; i++) res[i, 0] = S(all[i]);
            return new ArrayVal(res);
        }
        var first = Timed(() => rx.Match(text));
        if (!first.Success) throw Err(Errors.NA);
        if (mode == 0) return S(first.Value);
        int groups = first.Groups.Count - 1;
        if (groups == 0) throw Err(Errors.NA);
        var row = new CellValue[1, groups];
        for (int i = 0; i < groups; i++) row[0, i] = S(first.Groups[i + 1].Success ? first.Groups[i + 1].Value : "");
        return new ArrayVal(row);
    }

    internal static object RegexReplace(Ctx c, Expr[] a)
    {
        Args(a, 3, 5);
        var rx = MakeRegex(c.Str(a[1]), c, a, 4);
        var repl = c.Str(a[2]);
        int occ = (int)Trunc(c.NumOr(a, 3, 0));
        return Lift(c, c.Eval(a[0]), v =>
        {
            var text = Calc.ToText(v);
            if (occ == 0) return S(Timed(() => rx.Replace(text, repl)));
            var ms = Timed(() => rx.Matches(text).ToList());
            int idx = occ > 0 ? occ - 1 : ms.Count + occ;
            if (idx < 0 || idx >= ms.Count) return S(text);
            var m = ms[idx];
            return S(text[..m.Index] + m.Result(repl) + text[(m.Index + m.Length)..]);
        });
    }

    // ------------------------------------------------------------------ full-width / half-width (ASC, DBCS, JIS)

    private static readonly Dictionary<char, string> HalfToFull = new();
    private static readonly Dictionary<char, string> FullToHalf = new();

    static MiscMath()
    {
        for (char h = '｡'; h <= 'ﾝ'; h++)
        {
            var full = h.ToString().Normalize(NormalizationForm.FormKC);
            if (full.Length != 1) continue;
            HalfToFull[h] = full;
            FullToHalf[full[0]] = h.ToString();
            foreach (var mark in new[] { 'ﾞ', 'ﾟ' })
            {
                var voiced = (h.ToString() + mark).Normalize(NormalizationForm.FormKC);
                if (voiced.Length == 1) FullToHalf[voiced[0]] = h.ToString() + mark;
            }
        }
        HalfToFull['ﾞ'] = "゛"; HalfToFull['ﾟ'] = "゜";
        FullToHalf['゛'] = "ﾞ"; FullToHalf['゜'] = "ﾟ";
        FullToHalf['゙'] = "ﾞ"; FullToHalf['゚'] = "ﾟ";
    }

    internal static string ToHalfWidth(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            if (ch >= '！' && ch <= '～') sb.Append((char)(ch - 0xFEE0));
            else if (ch == '　') sb.Append(' ');
            else if (FullToHalf.TryGetValue(ch, out var h)) sb.Append(h);
            else sb.Append(ch);
        }
        return sb.ToString();
    }

    internal static string ToFullWidth(string s)
    {
        var sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            char ch = s[i];
            if (ch >= '!' && ch <= '~') { sb.Append((char)(ch + 0xFEE0)); continue; }
            if (ch == ' ') { sb.Append('　'); continue; }
            if (ch >= '｡' && ch <= 'ﾝ' && i + 1 < s.Length && s[i + 1] is 'ﾞ' or 'ﾟ')
            {
                var voiced = s.Substring(i, 2).Normalize(NormalizationForm.FormKC);
                if (voiced.Length == 1) { sb.Append(voiced); i++; continue; }
            }
            sb.Append(HalfToFull.TryGetValue(ch, out var f) ? f : ch.ToString());
        }
        return sb.ToString();
    }

    // ------------------------------------------------------------------ BAHTTEXT

    private static readonly string[] ThaiDigits = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];

    private static void ThaiBlock(StringBuilder sb, int v)
    {
        (int Div, string Name)[] units = [(100000, "แสน"), (10000, "หมื่น"), (1000, "พัน"), (100, "ร้อย")];
        foreach (var (div, name) in units)
            if (v >= div) { sb.Append(ThaiDigits[v / div]).Append(name); v %= div; }
        if (v <= 0) return;
        int ten = v / 10, one = v % 10;
        if (ten >= 1)
        {
            if (ten >= 3) sb.Append(ThaiDigits[ten]);
            else if (ten == 2) sb.Append("ยี่");
            sb.Append("สิบ");
        }
        if (ten > 0 && one == 1) sb.Append("เอ็ด");
        else if (one > 0) sb.Append(ThaiDigits[one]);
    }

    internal static object BahtText(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        double x = c.Num(a[0]);
        if (Math.Abs(x) >= 1e27) throw Err(Errors.Num);
        var dv = Math.Round((decimal)Math.Abs(x), 2, MidpointRounding.AwayFromZero);
        bool neg = x < 0 && dv != 0;
        var baht = Math.Floor(dv);
        int satang = (int)((dv - baht) * 100);
        var sb = new StringBuilder();
        if (baht == 0) { if (satang == 0) sb.Append(ThaiDigits[0]); }
        else
            while (baht > 0)
            {
                var block = new StringBuilder();
                int nb = (int)(baht % 1000000);
                baht = Math.Floor(baht / 1000000);
                if (nb > 0) ThaiBlock(block, nb);
                if (baht > 0) block.Insert(0, "ล้าน");
                sb.Insert(0, block.ToString());
            }
        if (sb.Length > 0) sb.Append("บาท");
        if (satang == 0) sb.Append("ถ้วน");
        else { ThaiBlock(sb, satang); sb.Append("สตางค์"); }
        if (neg) sb.Insert(0, "ลบ");
        return S(sb.ToString());
    }

    internal static object Phonetic(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        var v = c.Eval(a[0]);
        var first = v is RangeVal r ? c.Calc.ValueOf(r.Sheet, r.Range.R1, r.Range.C1) : c.Calc.ToScalar(v, c);
        if (first.IsError) throw Err(first.Str!);
        return S(first.IsText ? first.Str! : "");
    }

    // ================================================================== DATE (NETWORKDAYS.INTL, WORKDAY.INTL)

    private const double MaxSerial = 2958465;

    /// <summary>Weekend mask indexed Monday = 0 … Sunday = 6.</summary>
    private static bool[] WeekendMask(Ctx c, Expr[] a, int i)
    {
        var mask = new bool[7];
        if (!Has(a, i)) { mask[5] = mask[6] = true; return mask; }
        var v = c.Scalar(a[i]);
        if (v.IsError) throw Err(v.Str!);
        if (v.IsText)
        {
            var s = v.Str!;
            if (s.Length != 7 || s.Any(ch => ch != '0' && ch != '1')) throw Err(Errors.Value);
            for (int k = 0; k < 7; k++) mask[k] = s[k] == '1';
            return mask;
        }
        int n = (int)Trunc(Calc.ToNum(v));
        if (n >= 1 && n <= 7) { mask[(n + 4) % 7] = mask[(n + 5) % 7] = true; }
        else if (n >= 11 && n <= 17) mask[(n - 12 + 7) % 7] = true;
        else throw Err(Errors.Num);
        return mask;
    }

    private static int MonIndex(double serial) => (int)((((long)serial + 5) % 7 + 7) % 7);

    private static double DateArg(Ctx c, Expr e)
    {
        var v = c.Scalar(e);
        double d;
        if (v.IsText) { if (!ValueParser.TryCoerceNumber(v.Str!, out d)) throw Err(Errors.Value); }
        else d = Calc.ToNum(v);
        d = Math.Floor(d);
        if (d < 0 || d > MaxSerial) throw Err(Errors.Num);
        return d;
    }

    private static HashSet<double> HolidaySet(Ctx c, Expr[] a, int i)
    {
        var set = new HashSet<double>();
        if (!Has(a, i)) return set;
        foreach (var v in Functions.FlatVal(c, c.Eval(a[i])))
        {
            if (v.IsBlank) continue;
            if (v.IsError) throw Err(v.Str!);
            double d;
            if (v.IsText) { if (!ValueParser.TryCoerceNumber(v.Str!, out d)) throw Err(Errors.Value); }
            else d = v.Num;
            if (d < 0 || d > MaxSerial) throw Err(Errors.Num);
            set.Add(Math.Floor(d));
        }
        return set;
    }

    internal static object NetworkDaysIntl(Ctx c, Expr[] a)
    {
        Args(a, 2, 4);
        double s = DateArg(c, a[0]), e = DateArg(c, a[1]);
        var mask = WeekendMask(c, a, 2);
        var hol = HolidaySet(c, a, 3);
        int sign = 1;
        if (s > e) { (s, e) = (e, s); sign = -1; }
        long days = (long)(e - s) + 1, weeks = days / 7;
        int perWeek = mask.Count(w => !w);
        long count = weeks * perWeek;
        for (long t = weeks * 7; t < days; t++) if (!mask[MonIndex(s + t)]) count++;
        foreach (var h in hol) if (h >= s && h <= e && !mask[MonIndex(h)]) count--;
        return N(sign * count);
    }

    internal static object WorkdayIntl(Ctx c, Expr[] a)
    {
        Args(a, 2, 4);
        double d = DateArg(c, a[0]);
        long days = (long)Trunc(c.Num(a[1]));
        var mask = WeekendMask(c, a, 2);
        var hol = HolidaySet(c, a, 3);
        if (mask.All(w => w)) throw Err(Errors.Value);
        int step = Math.Sign(days);
        while (days != 0)
        {
            d += step;
            if (d < 0 || d > MaxSerial) throw Err(Errors.Num);
            if (!mask[MonIndex(d)] && !hol.Contains(d)) days -= step;
        }
        return N(d);
    }

    // ================================================================== INFORMATION

    internal static object Info(Ctx c, Expr[] a)
    {
        Args(a, 1, 1);
        switch (c.Str(a[0]).Trim().ToLowerInvariant())
        {
            case "directory":
                var dir = Environment.CurrentDirectory;
                return S(dir.EndsWith(Path.DirectorySeparatorChar) ? dir : dir + Path.DirectorySeparatorChar);
            case "numfile": return N(c.Book.Sheets.Count);
            case "origin": return S("$A:$A$1");
            case "osversion":
                var ver = Environment.OSVersion.Version;
                return S($"Windows ({(Environment.Is64BitOperatingSystem ? "64" : "32")}-bit) NT {ver.Major}.{ver.Minor:00}");
            case "recalc": return S(c.Book.ManualCalc ? "Manual" : "Automatic");
            case "release": return S("16.0");
            case "system": return S("pcdos");
            case "memavail": case "memused": case "totmem": throw Err(Errors.NA);
            default: throw Err(Errors.Value);
        }
    }

    // ================================================================== DATABASE

    private static bool SameHeader(CellValue h, CellValue name) =>
        !h.IsBlank && !h.IsError && string.Equals(Calc.ToText(h).Trim(), Calc.ToText(name).Trim(), StringComparison.CurrentCultureIgnoreCase);

    /// <summary>Database criteria: plain text means "begins with" (Advanced Filter semantics); everything else as COUNTIF.</summary>
    private static Functions.Criterion DbCriterion(CellValue v)
    {
        if (v.IsText)
        {
            var s = v.Str!;
            bool op = s.Length > 0 && (s[0] is '=' or '<' or '>');
            bool numeric = ValueParser.TryCoerceNumber(s, out _) && !s.Any(char.IsLetter);
            bool logical = s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || s.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
            if (!op && !numeric && !logical) return Functions.MakeCriterion(CellValue.Text(s + "*"));
        }
        return Functions.MakeCriterion(v);
    }

    private sealed record DbCond(int Col, Functions.Criterion? Crit, string? Formula, bool Never);

    internal static object Database(Ctx c, Expr[] a, string fn)
    {
        Args(a, 3, 3);
        var db = Functions.Grid(c, a[0]);
        int rows = db.GetLength(0), cols = db.GetLength(1);
        if (rows < 1) throw Err(Errors.Value);
        int field = -1;
        if (a[1] is not MissingExpr)
        {
            var fv = c.Scalar(a[1]);
            if (fv.IsError) throw Err(fv.Str!);
            if (fv.IsNumber || fv.IsBool)
            {
                int idx = (int)Trunc(fv.Num);
                if (idx < 1 || idx > cols) throw Err(Errors.Value);
                field = idx - 1;
            }
            else
            {
                for (int j = 0; j < cols && field < 0; j++) if (SameHeader(db[0, j], fv)) field = j;
                if (field < 0) throw Err(Errors.Value);
            }
        }
        else if (fn is not ("DCOUNT" or "DCOUNTA")) throw Err(Errors.Value);

        var crv = c.Eval(a[2]);
        var cr = Functions.Grid(c, crv);
        if (cr.GetLength(0) < 2) throw Err(Errors.Value);
        var critRows = new List<List<DbCond>>();
        for (int i = 1; i < cr.GetLength(0); i++)
        {
            var conds = new List<DbCond>();
            for (int j = 0; j < cr.GetLength(1); j++)
            {
                int col = -1;
                for (int k = 0; k < cols && col < 0; k++) if (SameHeader(db[0, k], cr[0, j])) col = k;
                var v = cr[i, j];
                if (col >= 0)
                {
                    if (v.IsBlank || (v.IsText && v.Str!.Length == 0)) continue;
                    conds.Add(new DbCond(col, DbCriterion(v), null, false));
                    continue;
                }
                // computed criterion: a formula under a header that is not a field name
                string? formula = crv is RangeVal rv ? rv.Sheet.GetCell(rv.Range.R1 + i, rv.Range.C1 + j)?.Formula : null;
                if (formula != null) { conds.Add(new DbCond(-1, null, formula, false)); continue; }
                if (v.IsBlank) continue;
                bool ok;
                try { ok = Calc.ToBool(v); } catch (FormulaError) { ok = false; }
                if (!ok) conds.Add(new DbCond(-1, null, null, true));
            }
            critRows.Add(conds);
        }

        bool Computed(string formula, int record)
        {
            try
            {
                var e = FormulaParser.Parse(RefAdjust.Shift(formula, record, 0));
                var r = c.Calc.ToScalar(c.Eval(e), c);
                return !r.IsError && (r.IsBool || r.IsNumber) && r.Num != 0;
            }
            catch (FormulaError) { return false; }
            catch (FormulaParseException) { return false; }
        }

        var matched = new List<CellValue>();
        int records = 0;
        for (int r = 1; r < rows; r++)
        {
            bool any = false;
            foreach (var conds in critRows)
            {
                bool all = true;
                foreach (var cd in conds)
                {
                    bool ok = cd.Never ? false : cd.Formula != null ? Computed(cd.Formula, r - 1) : cd.Crit!.Match(db[r, cd.Col]);
                    if (!ok) { all = false; break; }
                }
                if (all) { any = true; break; }
            }
            if (!any) continue;
            records++;
            if (field >= 0) matched.Add(db[r, field]);
        }

        if (fn == "DCOUNT" && field < 0) return N(records);
        if (fn == "DCOUNTA") return N(field < 0 ? records : matched.Count(v => !v.IsBlank));
        if (fn == "DCOUNT") return N(matched.Count(v => v.IsNumber));
        if (fn == "DGET")
        {
            if (matched.Count == 0) throw Err(Errors.Value);
            if (matched.Count > 1) throw Err(Errors.Num);
            return matched[0];
        }
        var nums = new List<double>();
        foreach (var v in matched)
        {
            if (v.IsError) throw Err(v.Str!);
            if (v.IsNumber) nums.Add(v.Num);
        }
        return fn switch
        {
            "DAVERAGE" => nums.Count == 0 ? throw Err(Errors.Div0) : N(nums.Average()),
            "DMAX" => N(nums.Count == 0 ? 0 : nums.Max()),
            "DMIN" => N(nums.Count == 0 ? 0 : nums.Min()),
            "DPRODUCT" => N(nums.Count == 0 ? 0 : nums.Aggregate(1.0, (x, y) => x * y)),
            "DSUM" => N(nums.Sum()),
            "DSTDEV" => N(Functions.StDev(nums, true)),
            "DSTDEVP" => N(Functions.StDev(nums, false)),
            "DVAR" => N(Functions.Var(nums, true)),
            _ => N(Functions.Var(nums, false))
        };
    }

    // ================================================================== WEB

    private static CellValue XmlValue(string s) =>
        double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) && s.Trim().Length > 0 ? N(d) : S(s);

    internal static object FilterXml(Ctx c, Expr[] a)
    {
        Args(a, 2, 2);
        var xml = c.Str(a[0]); var xpath = c.Str(a[1]);
        XPathNavigator nav;
        try
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using var reader = XmlReader.Create(new StringReader(xml), settings);
            nav = new XPathDocument(reader).CreateNavigator();
        }
        catch (XmlException) { throw Err(Errors.Value); }
        object result;
        try { result = nav.Evaluate(xpath); }
        catch (XPathException) { throw Err(Errors.Value); }
        switch (result)
        {
            case double d: return N(d);
            case bool b: return CellValue.Bool(b);
            case string s: return XmlValue(s);
            case XPathNodeIterator it:
            {
                var vals = new List<CellValue>();
                while (it.MoveNext()) vals.Add(XmlValue(it.Current!.Value));
                if (vals.Count == 0) throw Err(Errors.Value);
                if (vals.Count == 1) return vals[0];
                var res = new CellValue[vals.Count, 1];
                for (int i = 0; i < vals.Count; i++) res[i, 0] = vals[i];
                return new ArrayVal(res);
            }
        }
        throw Err(Errors.Value);
    }
}
