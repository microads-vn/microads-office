using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

public delegate object FnImpl(Ctx c, Expr[] a);

/// <summary>Worksheet function library.</summary>
public static partial class Functions
{
    public static readonly Dictionary<string, FnImpl> Table = new(StringComparer.OrdinalIgnoreCase);

    private static void Reg(FnImpl f, params string[] names) { foreach (var n in names) Table[n] = f; }
    private static CellValue N(double d) => CellValue.Number(d);
    private static CellValue S(string s) => CellValue.Text(s);
    private static CellValue B(bool b) => CellValue.Bool(b);
    private static FormulaError Err(string code) => new(code);
    private static void Args(Expr[] a, int min, int max = int.MaxValue) { if (a.Length < min || a.Length > max) throw Err(Errors.Value); }

    // ------------------------------------------------------------------ value collection helpers

    /// <summary>Numbers for aggregate functions: ranges/arrays contribute numbers only, direct arguments are coerced.</summary>
    internal static List<double> Nums(Ctx c, IEnumerable<Expr> args, bool countLogicalAndTextInRanges = false)
    {
        var res = new List<double>();
        foreach (var e in args)
        {
            if (e is MissingExpr) continue;
            var v = c.Eval(e);
            switch (v)
            {
                case RangeVal r:
                    foreach (var x in c.Calc.RangeValues(r)) AddFromRange(res, x, countLogicalAndTextInRanges);
                    break;
                case ArrayVal arr:
                    foreach (var x in arr.V) AddFromRange(res, x, countLogicalAndTextInRanges);
                    break;
                case CellValue s:
                    if (s.IsError) throw Err(s.Str!);
                    if (s.IsBlank) break;
                    res.Add(Calc.ToNum(s));
                    break;
            }
        }
        return res;
    }

    private static void AddFromRange(List<double> res, CellValue x, bool all)
    {
        switch (x.Kind)
        {
            case ValueKind.Number: res.Add(x.Num); break;
            case ValueKind.Error: throw Err(x.Str!);
            case ValueKind.Bool: if (all) res.Add(x.Num); break;
            case ValueKind.Text: if (all) res.Add(0); break;
        }
    }

    /// <summary>All values of the arguments (ranges flattened).</summary>
    internal static IEnumerable<CellValue> Flat(Ctx c, IEnumerable<Expr> args)
    {
        foreach (var e in args)
        {
            if (e is MissingExpr) continue;
            var v = c.Eval(e);
            foreach (var x in FlatVal(c, v)) yield return x;
        }
    }

    internal static IEnumerable<CellValue> FlatVal(Ctx c, object v)
    {
        switch (v)
        {
            case RangeVal r: foreach (var x in c.Calc.RangeValues(r)) yield return x; break;
            case ArrayVal a: foreach (var x in a.V) yield return x; break;
            case CellValue s: yield return s; break;
        }
    }

    /// <summary>2-D view of a range/array (clipped to the used area for whole columns).</summary>
    internal static CellValue[,] Grid(Ctx c, Expr e) => c.Calc.ToArray(c.Eval(e));

    internal static CellValue[,] Grid(Ctx c, object v) => c.Calc.ToArray(v);

    private static RangeVal RangeArg(Ctx c, Expr e) => c.Eval(e) as RangeVal ?? throw Err(Errors.Value);

    // ------------------------------------------------------------------ criteria (COUNTIF, SUMIF...)

    internal sealed class Criterion
    {
        private readonly Func<CellValue, bool> _f;
        public Criterion(Func<CellValue, bool> f) { _f = f; }
        public bool Match(CellValue v) => _f(v);
    }

    internal static Criterion MakeCriterion(CellValue crit)
    {
        if (crit.IsError) throw Err(crit.Str!);
        if (crit.IsNumber) { var n = crit.Num; return new Criterion(v => NumEq(v, n)); }
        if (crit.IsBool) { var b = crit.Num; return new Criterion(v => v.IsBool && v.Num == b); }
        if (crit.IsBlank) return new Criterion(v => v.IsBlank || (v.IsText && v.Str!.Length == 0));
        var s = crit.Str!;
        string op = "";
        foreach (var o in new[] { ">=", "<=", "<>", ">", "<", "=" })
            if (s.StartsWith(o)) { op = o; s = s[o.Length..]; break; }
        if (s.Length == 0)
        {
            return op switch
            {
                "<>" => new Criterion(v => !(v.IsBlank || (v.IsText && v.Str!.Length == 0))),
                _ => new Criterion(v => v.IsBlank || (v.IsText && v.Str!.Length == 0))
            };
        }
        bool isNum = ValueParser.TryCoerceNumber(s, out double num) && !s.Any(char.IsLetter);
        if (isNum)
        {
            return op switch
            {
                ">" => new Criterion(v => v.IsNumber && v.Num > num),
                ">=" => new Criterion(v => v.IsNumber && v.Num >= num),
                "<" => new Criterion(v => v.IsNumber && v.Num < num),
                "<=" => new Criterion(v => v.IsNumber && v.Num <= num),
                "<>" => new Criterion(v => !NumEq(v, num)),
                _ => new Criterion(v => NumEq(v, num))
            };
        }
        if (s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || s.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            double b = s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            return op == "<>" ? new Criterion(v => !(v.IsBool && v.Num == b)) : new Criterion(v => v.IsBool && v.Num == b);
        }
        if (op is "" or "=" or "<>")
        {
            var rx = WildcardRegex(s);
            bool eq(CellValue v) => v.IsText ? rx.IsMatch(v.Str!) : (v.IsBlank ? false : rx.IsMatch(Calc.ToText(v)) && !v.IsNumber);
            return op == "<>" ? new Criterion(v => !eq(v)) : new Criterion(eq);
        }
        int Cmp(CellValue v) => string.Compare(v.Str, s, StringComparison.CurrentCultureIgnoreCase);
        return op switch
        {
            ">" => new Criterion(v => v.IsText && Cmp(v) > 0),
            ">=" => new Criterion(v => v.IsText && Cmp(v) >= 0),
            "<" => new Criterion(v => v.IsText && Cmp(v) < 0),
            _ => new Criterion(v => v.IsText && Cmp(v) <= 0)
        };
    }

    private static bool NumEq(CellValue v, double n)
    {
        if (v.IsNumber) return Math.Abs(v.Num - n) < 1e-12 * Math.Max(1, Math.Abs(n));
        if (v.IsText && ValueParser.TryCoerceNumber(v.Str!, out var d) && !v.Str!.Any(char.IsLetter)) return Math.Abs(d - n) < 1e-12 * Math.Max(1, Math.Abs(n));
        return false;
    }

    private static readonly Dictionary<string, Regex> RxCache = new();

    internal static Regex WildcardRegex(string pattern)
    {
        lock (RxCache)
        {
            if (RxCache.TryGetValue(pattern, out var rx)) return rx;
            var sb = new StringBuilder("^");
            for (int i = 0; i < pattern.Length; i++)
            {
                char ch = pattern[i];
                if (ch == '~' && i + 1 < pattern.Length) { sb.Append(Regex.Escape(pattern[++i].ToString())); continue; }
                sb.Append(ch switch { '*' => ".*", '?' => ".", _ => Regex.Escape(ch.ToString()) });
            }
            sb.Append('$');
            rx = new Regex(sb.ToString(), RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);
            if (RxCache.Count > 500) RxCache.Clear();
            RxCache[pattern] = rx;
            return rx;
        }
    }

    /// <summary>Indices (row, col offsets) of cells matching every (range, criterion) pair.</summary>
    private static List<(int R, int C)> MatchIfs(Ctx c, Expr[] a, int start)
    {
        if ((a.Length - start) % 2 != 0 || a.Length - start < 2) throw Err(Errors.Value);
        List<(int, int)>? res = null;
        int rows = -1, cols = -1;
        for (int i = start; i < a.Length; i += 2)
        {
            var g = Grid(c, a[i]);
            var crit = MakeCriterion(c.Scalar(a[i + 1]));
            if (rows < 0) { rows = g.GetLength(0); cols = g.GetLength(1); }
            else if (g.GetLength(0) != rows || g.GetLength(1) != cols)
            {
                // clipping can produce different shapes for whole-column refs: use the smaller area
                rows = Math.Min(rows, g.GetLength(0)); cols = Math.Min(cols, g.GetLength(1));
            }
            var cur = new List<(int, int)>();
            if (res == null)
            {
                for (int r = 0; r < g.GetLength(0); r++)
                    for (int k = 0; k < g.GetLength(1); k++) if (crit.Match(g[r, k])) cur.Add((r, k));
            }
            else foreach (var (r, k) in res) if (r < g.GetLength(0) && k < g.GetLength(1) && crit.Match(g[r, k])) cur.Add((r, k));
            res = cur;
        }
        return res ?? new();
    }

    /// <summary>Value at (r, k) offset of a range argument (used by SUMIF's sum_range).</summary>
    private static CellValue At(Ctx c, object range, int r, int k)
    {
        if (range is RangeVal rv)
        {
            int row = rv.Range.R1 + r, col = rv.Range.C1 + k;
            if (row >= Addr.MaxRows || col >= Addr.MaxCols) return CellValue.Blank;
            return c.Calc.ValueOf(rv.Sheet, row, col);
        }
        if (range is ArrayVal av) return r < av.Rows && k < av.Cols ? av.V[r, k] : CellValue.Blank;
        return range is CellValue s && r == 0 && k == 0 ? s : CellValue.Blank;
    }

    static Functions()
    {
        RegisterMath();
        RegisterStat();
        RegisterLogic();
        RegisterLookup();
        RegisterText();
        RegisterDate();
        RegisterFinance();
        RegisterInfo();
    }

    // ================================================================== MATH

    private static void RegisterMath()
    {
        Reg((c, a) => N(Nums(c, a).Sum()), "SUM");
        Reg((c, a) => { var l = Nums(c, a); return N(l.Count == 0 ? 0 : l.Aggregate(1.0, (x, y) => x * y)); }, "PRODUCT");
        Reg((c, a) => N(Nums(c, a).Sum(x => x * x)), "SUMSQ");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var g = Grid(c, a[0]);
            var crit = MakeCriterion(c.Scalar(a[1]));
            var sumRange = a.Length > 2 && a[2] is not MissingExpr ? c.Eval(a[2]) : c.Eval(a[0]);
            double s = 0;
            for (int r = 0; r < g.GetLength(0); r++)
                for (int k = 0; k < g.GetLength(1); k++)
                    if (crit.Match(g[r, k])) { var v = At(c, sumRange, r, k); if (v.IsNumber) s += v.Num; else if (v.IsError) throw Err(v.Str!); }
            return N(s);
        }, "SUMIF");
        Reg((c, a) =>
        {
            Args(a, 3);
            var sum = c.Eval(a[0]);
            double s = 0;
            foreach (var (r, k) in MatchIfs(c, a, 1)) { var v = At(c, sum, r, k); if (v.IsNumber) s += v.Num; else if (v.IsError) throw Err(v.Str!); }
            return N(s);
        }, "SUMIFS");
        Reg((c, a) =>
        {
            Args(a, 1);
            var grids = a.Select(e => Grid(c, e)).ToList();
            int rows = grids[0].GetLength(0), cols = grids[0].GetLength(1);
            if (grids.Any(g => g.GetLength(0) != rows || g.GetLength(1) != cols)) throw Err(Errors.Value);
            double total = 0;
            for (int r = 0; r < rows; r++)
                for (int k = 0; k < cols; k++)
                {
                    double p = 1;
                    foreach (var g in grids)
                    {
                        var v = g[r, k];
                        if (v.IsError) throw Err(v.Str!);
                        p *= v.IsNumber ? v.Num : v.IsBool && grids.Count == 1 ? 0 : 0;
                    }
                    total += p;
                }
            return N(total);
        }, "SUMPRODUCT");
        Reg((c, a) => { Args(a, 1, 2); return N(RoundAway(c.Num(a[0]), (int)c.NumOr(a, 1, 0))); }, "ROUND");
        Reg((c, a) => { Args(a, 2, 2); double x = c.Num(a[0]); int d = c.Int(a[1]); double f = Math.Pow(10, d); return N(Math.Sign(x) * Math.Ceiling(Math.Round(Math.Abs(x) * f, 9)) / f); }, "ROUNDUP");
        Reg((c, a) => { Args(a, 2, 2); double x = c.Num(a[0]); int d = c.Int(a[1]); double f = Math.Pow(10, d); return N(Math.Sign(x) * Math.Floor(Math.Round(Math.Abs(x) * f, 9)) / f); }, "ROUNDDOWN");
        Reg((c, a) => { Args(a, 1, 2); double x = c.Num(a[0]); int d = (int)c.NumOr(a, 1, 0); double f = Math.Pow(10, d); return N(Math.Truncate(Math.Round(x * f, 9)) / f); }, "TRUNC");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Floor(c.Num(a[0]))); }, "INT");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double n = c.Num(a[0]), d = c.Num(a[1]);
            if (d == 0) throw Err(Errors.Div0);
            return N(n - d * Math.Floor(n / d));
        }, "MOD");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Abs(c.Num(a[0]))); }, "ABS");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); if (x < 0) throw Err(Errors.Num); return N(Math.Sqrt(x)); }, "SQRT");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); if (x < 0) throw Err(Errors.Num); return N(Math.Sqrt(x * Math.PI)); }, "SQRTPI");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double b = c.Num(a[0]), e = c.Num(a[1]);
            if (b == 0 && e < 0) throw Err(Errors.Div0);
            var r = Math.Pow(b, e);
            if (double.IsNaN(r) || double.IsInfinity(r)) throw Err(Errors.Num);
            return N(r);
        }, "POWER");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Exp(c.Num(a[0]))); }, "EXP");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); if (x <= 0) throw Err(Errors.Num); return N(Math.Log(x)); }, "LN");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); if (x <= 0) throw Err(Errors.Num); return N(Math.Log10(x)); }, "LOG10");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            var x = c.Num(a[0]); var b = c.NumOr(a, 1, 10);
            if (x <= 0 || b <= 0) throw Err(Errors.Num);
            if (b == 1) throw Err(Errors.Div0);
            return N(Math.Log(x) / Math.Log(b));
        }, "LOG");
        Reg((c, a) => N(Math.PI), "PI");
        Reg((c, a) => N(Random.Shared.NextDouble()), "RAND");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double lo = Math.Ceiling(c.Num(a[0])), hi = Math.Floor(c.Num(a[1]));
            if (lo > hi) throw Err(Errors.Num);
            return N(Math.Floor(lo + Random.Shared.NextDouble() * (hi - lo + 1)));
        }, "RANDBETWEEN");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            double x = c.Num(a[0]), s = c.NumOr(a, 1, x >= 0 ? 1 : -1);
            if (s == 0) return N(0);
            if (x > 0 && s < 0) throw Err(Errors.Num);
            return N(Math.Ceiling(Math.Round(x / s, 9)) * s);
        }, "CEILING");
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            double x = c.Num(a[0]), s = Math.Abs(c.NumOr(a, 1, 1)); double mode = c.NumOr(a, 2, 0);
            if (s == 0) return N(0);
            if (x < 0 && mode != 0) return N(-Math.Ceiling(Math.Round(-x / s, 9)) * s);
            return N(Math.Ceiling(Math.Round(x / s, 9)) * s);
        }, "CEILING.MATH", "CEILING.PRECISE", "ISO.CEILING");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            double x = c.Num(a[0]), s = c.NumOr(a, 1, x >= 0 ? 1 : -1);
            if (s == 0) throw Err(Errors.Div0);
            if (x > 0 && s < 0) throw Err(Errors.Num);
            return N(Math.Floor(Math.Round(x / s, 9)) * s);
        }, "FLOOR");
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            double x = c.Num(a[0]), s = Math.Abs(c.NumOr(a, 1, 1)); double mode = c.NumOr(a, 2, 0);
            if (s == 0) return N(0);
            if (x < 0 && mode != 0) return N(-Math.Floor(Math.Round(-x / s, 9)) * s);
            return N(Math.Floor(Math.Round(x / s, 9)) * s);
        }, "FLOOR.MATH", "FLOOR.PRECISE");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Sign(c.Num(a[0]))); }, "SIGN");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var n = Math.Floor(c.Num(a[0]));
            if (n < 0 || n > 170) throw Err(Errors.Num);
            double f = 1; for (int i = 2; i <= n; i++) f *= i;
            return N(f);
        }, "FACT");
        Reg((c, a) => { Args(a, 2, 2); double n = Math.Floor(c.Num(a[0])), k = Math.Floor(c.Num(a[1])); if (n < 0 || k < 0 || k > n) throw Err(Errors.Num); return N(Math.Round(Combin(n, k))); }, "COMBIN");
        Reg((c, a) => { Args(a, 2, 2); double n = Math.Floor(c.Num(a[0])), k = Math.Floor(c.Num(a[1])); if (n < 0 || k < 0 || k > n) throw Err(Errors.Num); double p = 1; for (int i = 0; i < k; i++) p *= n - i; return N(p); }, "PERMUT");
        Reg((c, a) =>
        {
            var l = Nums(c, a).Select(x => (long)Math.Floor(x)).ToList();
            if (l.Any(x => x < 0)) throw Err(Errors.Num);
            long g = 0; foreach (var x in l) g = Gcd(g, x);
            return N(g);
        }, "GCD");
        Reg((c, a) =>
        {
            var l = Nums(c, a).Select(x => (long)Math.Floor(x)).ToList();
            if (l.Any(x => x < 0)) throw Err(Errors.Num);
            long m = 1; foreach (var x in l) { if (x == 0) return N(0); m = m / Gcd(m, x) * x; }
            return N(m);
        }, "LCM");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); var r = Math.Ceiling(Math.Abs(x)); if (r % 2 != 0) r++; return N(Math.Sign(x) * r); }, "EVEN");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); var r = Math.Ceiling(Math.Abs(x)); if (r % 2 == 0) r++; return N(x < 0 ? -r : r); }, "ODD");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            double x = c.Num(a[0]), m = c.Num(a[1]);
            if (m == 0) return N(0);
            if (Math.Sign(x) * Math.Sign(m) < 0) throw Err(Errors.Num);
            return N(RoundAway(x / m, 0) * m);
        }, "MROUND");
        Reg((c, a) => { Args(a, 2, 2); var d = c.Num(a[1]); if (d == 0) throw Err(Errors.Div0); return N(Math.Truncate(c.Num(a[0]) / d)); }, "QUOTIENT");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Sin(c.Num(a[0]))); }, "SIN");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Cos(c.Num(a[0]))); }, "COS");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Tan(c.Num(a[0]))); }, "TAN");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); if (Math.Abs(x) > 1) throw Err(Errors.Num); return N(Math.Asin(x)); }, "ASIN");
        Reg((c, a) => { Args(a, 1, 1); var x = c.Num(a[0]); if (Math.Abs(x) > 1) throw Err(Errors.Num); return N(Math.Acos(x)); }, "ACOS");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Atan(c.Num(a[0]))); }, "ATAN");
        Reg((c, a) => { Args(a, 2, 2); double x = c.Num(a[0]), y = c.Num(a[1]); if (x == 0 && y == 0) throw Err(Errors.Div0); return N(Math.Atan2(y, x)); }, "ATAN2");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Sinh(c.Num(a[0]))); }, "SINH");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Cosh(c.Num(a[0]))); }, "COSH");
        Reg((c, a) => { Args(a, 1, 1); return N(Math.Tanh(c.Num(a[0]))); }, "TANH");
        Reg((c, a) => { Args(a, 1, 1); return N(c.Num(a[0]) * 180 / Math.PI); }, "DEGREES");
        Reg((c, a) => { Args(a, 1, 1); return N(c.Num(a[0]) * Math.PI / 180); }, "RADIANS");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            int n = (int)c.Num(a[0]);
            if (n < 0 || n > 3999) throw Err(Errors.Value);
            return S(Roman(n));
        }, "ROMAN");
        Reg((c, a) =>
        {
            Args(a, 2);
            int fn = (int)c.Num(a[0]);
            bool skipHidden = fn > 100;
            int f = fn % 100;
            var vals = new List<double>();
            int counta = 0;
            foreach (var e in a.Skip(1))
            {
                if (c.Eval(e) is not RangeVal r) throw Err(Errors.Value);
                var rg = c.Calc.Clip(r);
                for (int row = rg.R1; row <= rg.R2; row++)
                {
                    if (r.Sheet.FilteredRows.Contains(row) || (skipHidden && r.Sheet.HiddenRows.Contains(row))) continue;
                    for (int col = rg.C1; col <= rg.C2; col++)
                    {
                        var cell = r.Sheet.GetCell(row, col);
                        if (cell?.Formula != null && cell.Formula.TrimStart().StartsWith("SUBTOTAL", StringComparison.OrdinalIgnoreCase)) continue;
                        var v = c.Calc.ValueOf(r.Sheet, row, col);
                        if (v.IsError) throw Err(v.Str!);
                        if (v.IsNumber) vals.Add(v.Num);
                        if (!v.IsBlank) counta++;
                    }
                }
            }
            return f switch
            {
                1 => vals.Count == 0 ? throw Err(Errors.Div0) : N(vals.Average()),
                2 => N(vals.Count),
                3 => N(counta),
                4 => N(vals.Count == 0 ? 0 : vals.Max()),
                5 => N(vals.Count == 0 ? 0 : vals.Min()),
                6 => N(vals.Aggregate(1.0, (x, y) => x * y)),
                7 => N(StDev(vals, true)),
                8 => N(StDev(vals, false)),
                9 => N(vals.Sum()),
                10 => N(Var(vals, true)),
                11 => N(Var(vals, false)),
                _ => throw Err(Errors.Value)
            };
        }, "SUBTOTAL");
        Reg((c, a) =>
        {
            Args(a, 1, 4);
            int rows = (int)c.Num(a[0]), cols = (int)c.NumOr(a, 1, 1);
            double start = c.NumOr(a, 2, 1), step = c.NumOr(a, 3, 1);
            if (rows < 1 || cols < 1) throw Err(Errors.Value);
            var arr = new CellValue[rows, cols];
            for (int r = 0; r < rows; r++) for (int k = 0; k < cols; k++) arr[r, k] = N(start + (r * cols + k) * step);
            return new ArrayVal(arr);
        }, "SEQUENCE");
    }

    internal static double RoundAway(double x, int digits)
    {
        if (digits >= 0)
        {
            var d = Math.Round((decimal)Math.Clamp(x, -7.9e27, 7.9e27), Math.Min(digits, 28), MidpointRounding.AwayFromZero);
            return Math.Abs(x) < 7.9e27 ? (double)d : x;
        }
        double f = Math.Pow(10, -digits);
        return Math.Round(x / f, MidpointRounding.AwayFromZero) * f;
    }

    private static double Combin(double n, double k)
    {
        k = Math.Min(k, n - k);
        double r = 1;
        for (int i = 1; i <= k; i++) r = r * (n - k + i) / i;
        return r;
    }

    private static long Gcd(long a, long b) { while (b != 0) (a, b) = (b, a % b); return Math.Abs(a); }

    private static string Roman(int n)
    {
        var vals = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var syms = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new StringBuilder();
        for (int i = 0; i < vals.Length; i++) while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
        return sb.ToString();
    }

    // ================================================================== STATISTICS

    internal static double Var(List<double> v, bool sample)
    {
        int n = v.Count;
        if (n < (sample ? 2 : 1)) throw Err(Errors.Div0);
        double m = v.Average();
        return v.Sum(x => (x - m) * (x - m)) / (sample ? n - 1 : n);
    }

    internal static double StDev(List<double> v, bool sample) => Math.Sqrt(Var(v, sample));

    internal static double Percentile(List<double> v, double p, bool inclusive = true)
    {
        if (v.Count == 0) throw Err(Errors.Num);
        if (p < 0 || p > 1) throw Err(Errors.Num);
        var s = v.OrderBy(x => x).ToList();
        double rank;
        if (inclusive) rank = p * (s.Count - 1);
        else
        {
            rank = p * (s.Count + 1) - 1;
            if (rank < 0 || rank > s.Count - 1) throw Err(Errors.Num);
        }
        int lo = (int)Math.Floor(rank);
        double frac = rank - lo;
        return lo + 1 < s.Count ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
    }

    private static void RegisterStat()
    {
        Reg((c, a) => { var l = Nums(c, a); if (l.Count == 0) throw Err(Errors.Div0); return N(l.Average()); }, "AVERAGE");
        Reg((c, a) => { var l = Nums(c, a, true); if (l.Count == 0) throw Err(Errors.Div0); return N(l.Average()); }, "AVERAGEA");
        Reg((c, a) =>
        {
            int n = 0;
            foreach (var e in a)
            {
                if (e is MissingExpr) continue;
                var v = c.Eval(e);
                if (v is CellValue s) { if (s.IsNumber || s.IsBool || (s.IsText && ValueParser.TryCoerceNumber(s.Str!, out _))) n++; }
                else n += FlatVal(c, v).Count(x => x.IsNumber);
            }
            return N(n);
        }, "COUNT");
        Reg((c, a) =>
        {
            int n = 0;
            foreach (var e in a)
            {
                if (e is MissingExpr) continue;
                var v = c.Eval(e);
                if (v is CellValue s) n++;
                else n += FlatVal(c, v).Count(x => !x.IsBlank);
            }
            return N(n);
        }, "COUNTA");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var v = c.Eval(a[0]);
            if (v is RangeVal r)
            {
                var rg = r.Range;
                long total = rg.Count;
                long nonBlank = c.Calc.RangeValues(r).Count(x => !(x.IsBlank || (x.IsText && x.Str!.Length == 0)));
                return N(total - nonBlank);
            }
            return N(FlatVal(c, v).Count(x => x.IsBlank || (x.IsText && x.Str!.Length == 0)));
        }, "COUNTBLANK");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var crit = MakeCriterion(c.Scalar(a[1]));
            var v = c.Eval(a[0]);
            if (v is RangeVal r && r.Range.Count > 20000)
            {
                // whole column: blanks that match "" must be counted over the full range
                var vals = c.Calc.RangeValues(r).ToList();
                long n = vals.Count(crit.Match);
                if (crit.Match(CellValue.Blank)) n += r.Range.Count - vals.Count;
                return N(n);
            }
            return N(FlatVal(c, v).Count(crit.Match));
        }, "COUNTIF");
        Reg((c, a) => N(MatchIfs(c, a, 0).Count), "COUNTIFS");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var g = Grid(c, a[0]);
            var crit = MakeCriterion(c.Scalar(a[1]));
            var avg = a.Length > 2 && a[2] is not MissingExpr ? c.Eval(a[2]) : c.Eval(a[0]);
            var l = new List<double>();
            for (int r = 0; r < g.GetLength(0); r++)
                for (int k = 0; k < g.GetLength(1); k++)
                    if (crit.Match(g[r, k])) { var v = At(c, avg, r, k); if (v.IsNumber) l.Add(v.Num); }
            if (l.Count == 0) throw Err(Errors.Div0);
            return N(l.Average());
        }, "AVERAGEIF");
        Reg((c, a) =>
        {
            var avg = c.Eval(a[0]);
            var l = MatchIfs(c, a, 1).Select(p => At(c, avg, p.R, p.C)).Where(v => v.IsNumber).Select(v => v.Num).ToList();
            if (l.Count == 0) throw Err(Errors.Div0);
            return N(l.Average());
        }, "AVERAGEIFS");
        Reg((c, a) => { var l = Nums(c, a); return N(l.Count == 0 ? 0 : l.Max()); }, "MAX");
        Reg((c, a) => { var l = Nums(c, a); return N(l.Count == 0 ? 0 : l.Min()); }, "MIN");
        Reg((c, a) => { var l = Nums(c, a, true); return N(l.Count == 0 ? 0 : l.Max()); }, "MAXA");
        Reg((c, a) => { var l = Nums(c, a, true); return N(l.Count == 0 ? 0 : l.Min()); }, "MINA");
        Reg((c, a) =>
        {
            var m = c.Eval(a[0]);
            var l = MatchIfs(c, a, 1).Select(p => At(c, m, p.R, p.C)).Where(v => v.IsNumber).Select(v => v.Num).ToList();
            return N(l.Count == 0 ? 0 : l.Max());
        }, "MAXIFS");
        Reg((c, a) =>
        {
            var m = c.Eval(a[0]);
            var l = MatchIfs(c, a, 1).Select(p => At(c, m, p.R, p.C)).Where(v => v.IsNumber).Select(v => v.Num).ToList();
            return N(l.Count == 0 ? 0 : l.Min());
        }, "MINIFS");
        Reg((c, a) =>
        {
            var l = Nums(c, a).OrderBy(x => x).ToList();
            if (l.Count == 0) throw Err(Errors.Num);
            int n = l.Count;
            return N(n % 2 == 1 ? l[n / 2] : (l[n / 2 - 1] + l[n / 2]) / 2);
        }, "MEDIAN");
        Reg((c, a) =>
        {
            var l = Nums(c, a);
            var best = l.Select((x, i) => (x, i)).GroupBy(t => t.x).Where(g => g.Count() > 1).OrderByDescending(g => g.Count()).ThenBy(g => g.First().i).FirstOrDefault();
            if (best == null) throw Err(Errors.NA);
            return N(best.Key);
        }, "MODE", "MODE.SNGL");
        Reg((c, a) => N(StDev(Nums(c, a), true)), "STDEV", "STDEV.S");
        Reg((c, a) => N(StDev(Nums(c, a), false)), "STDEVP", "STDEV.P");
        Reg((c, a) => N(StDev(Nums(c, a, true), true)), "STDEVA");
        Reg((c, a) => N(Var(Nums(c, a), true)), "VAR", "VAR.S");
        Reg((c, a) => N(Var(Nums(c, a), false)), "VARP", "VAR.P");
        Reg((c, a) =>
        {
            var l = Nums(c, a);
            if (l.Count == 0) throw Err(Errors.Num);
            var m = l.Average();
            return N(l.Sum(x => Math.Abs(x - m)) / l.Count);
        }, "AVEDEV");
        Reg((c, a) =>
        {
            var l = Nums(c, a);
            if (l.Count == 0 || l.Any(x => x <= 0)) throw Err(Errors.Num);
            return N(Math.Exp(l.Average(Math.Log)));
        }, "GEOMEAN");
        Reg((c, a) =>
        {
            var l = Nums(c, a);
            if (l.Count == 0 || l.Any(x => x <= 0)) throw Err(Errors.Num);
            return N(l.Count / l.Sum(x => 1 / x));
        }, "HARMEAN");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var l = Nums(c, [a[0]]).OrderByDescending(x => x).ToList();
            int k = (int)Math.Ceiling(c.Num(a[1]));
            if (k < 1 || k > l.Count) throw Err(Errors.Num);
            return N(l[k - 1]);
        }, "LARGE");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var l = Nums(c, [a[0]]).OrderBy(x => x).ToList();
            int k = (int)Math.Ceiling(c.Num(a[1]));
            if (k < 1 || k > l.Count) throw Err(Errors.Num);
            return N(l[k - 1]);
        }, "SMALL");
        FnImpl rank = (c, a) =>
        {
            Args(a, 2, 3);
            double x = c.Num(a[0]);
            var l = Nums(c, [a[1]]);
            bool asc = c.NumOr(a, 2, 0) != 0;
            if (!l.Any(v => v == x)) throw Err(Errors.NA);
            return N(asc ? l.Count(v => v < x) + 1 : l.Count(v => v > x) + 1);
        };
        Reg(rank, "RANK", "RANK.EQ");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            double x = c.Num(a[0]);
            var l = Nums(c, [a[1]]);
            bool asc = c.NumOr(a, 2, 0) != 0;
            int eq = l.Count(v => v == x);
            if (eq == 0) throw Err(Errors.NA);
            int better = asc ? l.Count(v => v < x) : l.Count(v => v > x);
            return N(better + (eq + 1) / 2.0);
        }, "RANK.AVG");
        Reg((c, a) => { Args(a, 2, 2); return N(Percentile(Nums(c, [a[0]]), c.Num(a[1]))); }, "PERCENTILE", "PERCENTILE.INC");
        Reg((c, a) => { Args(a, 2, 2); return N(Percentile(Nums(c, [a[0]]), c.Num(a[1]), false)); }, "PERCENTILE.EXC");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            int q = (int)c.Num(a[1]);
            if (q < 0 || q > 4) throw Err(Errors.Num);
            return N(Percentile(Nums(c, [a[0]]), q / 4.0));
        }, "QUARTILE", "QUARTILE.INC");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var (x, y) = Pairs(c, a[0], a[1]);
            if (x.Count < 2) throw Err(Errors.Div0);
            double mx = x.Average(), my = y.Average();
            double sxy = 0, sxx = 0, syy = 0;
            for (int i = 0; i < x.Count; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) * (x[i] - mx); syy += (y[i] - my) * (y[i] - my); }
            if (sxx == 0 || syy == 0) throw Err(Errors.Div0);
            return N(sxy / Math.Sqrt(sxx * syy));
        }, "CORREL", "PEARSON");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var (y, x) = Pairs(c, a[0], a[1]);
            var (slope, _) = Regress(x, y);
            return N(slope);
        }, "SLOPE");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var (y, x) = Pairs(c, a[0], a[1]);
            var (_, icpt) = Regress(x, y);
            return N(icpt);
        }, "INTERCEPT");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double x0 = c.Num(a[0]);
            var (y, x) = Pairs(c, a[1], a[2]);
            var (slope, icpt) = Regress(x, y);
            return N(icpt + slope * x0);
        }, "FORECAST", "FORECAST.LINEAR");
    }

    private static (List<double>, List<double>) Pairs(Ctx c, Expr ea, Expr eb)
    {
        var A = Grid(c, ea); var Bg = Grid(c, eb);
        if (A.Length != Bg.Length) throw Err(Errors.NA);
        var fa = A.Cast<CellValue>().ToList(); var fb = Bg.Cast<CellValue>().ToList();
        var x = new List<double>(); var y = new List<double>();
        for (int i = 0; i < fa.Count; i++)
            if (fa[i].IsNumber && fb[i].IsNumber) { x.Add(fa[i].Num); y.Add(fb[i].Num); }
        return (x, y);
    }

    private static (double Slope, double Intercept) Regress(List<double> x, List<double> y)
    {
        if (x.Count < 2) throw Err(Errors.Div0);
        double mx = x.Average(), my = y.Average(), sxy = 0, sxx = 0;
        for (int i = 0; i < x.Count; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) * (x[i] - mx); }
        if (sxx == 0) throw Err(Errors.Div0);
        var s = sxy / sxx;
        return (s, my - s * mx);
    }

    // ================================================================== LOGICAL

    private static void RegisterLogic()
    {
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            var cond = c.Eval(a[0]);
            if (cond is ArrayVal or RangeVal { Range.IsSingle: false })
            {
                // element-wise IF (array formulas, SUMPRODUCT(IF(...)))
                var g = Grid(c, cond);
                object tv = a.Length > 1 ? c.Eval(a[1]) : CellValue.Bool(true);
                object fv = a.Length > 2 ? c.Eval(a[2]) : CellValue.Bool(false);
                var tg = Grid(c, tv); var fg = Grid(c, fv);
                var res = new CellValue[g.GetLength(0), g.GetLength(1)];
                for (int i = 0; i < g.GetLength(0); i++)
                    for (int j = 0; j < g.GetLength(1); j++)
                    {
                        var cv = g[i, j];
                        if (cv.IsError) { res[i, j] = cv; continue; }
                        bool b;
                        try { b = Calc.ToBool(cv); } catch (FormulaError e) { res[i, j] = CellValue.Error(e.Code); continue; }
                        var src = b ? tg : fg;
                        res[i, j] = src[Math.Min(i, src.GetLength(0) - 1), Math.Min(j, src.GetLength(1) - 1)];
                    }
                return new ArrayVal(res);
            }
            bool test = Calc.ToBool(c.Calc.ToScalar(cond, c));
            if (test) return a.Length > 1 ? (a[1] is MissingExpr ? N(0) : c.Eval(a[1])) : B(true);
            return a.Length > 2 ? (a[2] is MissingExpr ? N(0) : c.Eval(a[2])) : B(false);
        }, "IF");
        Reg((c, a) =>
        {
            if (a.Length < 2 || a.Length % 2 != 0) throw Err(Errors.Value);
            for (int i = 0; i < a.Length; i += 2) if (c.Bool(a[i])) return c.Eval(a[i + 1]);
            throw Err(Errors.NA);
        }, "IFS");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var v = c.Eval(a[0]);
            var s = v is CellValue cv ? cv : c.Calc.ToScalar(v, c);
            return s.IsError ? c.Eval(a[1]) : v;
        }, "IFERROR");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var v = c.Eval(a[0]);
            var s = c.Calc.ToScalar(v, c);
            return s.IsError && s.Str == Errors.NA ? c.Eval(a[1]) : v;
        }, "IFNA");
        Reg((c, a) => B(Logicals(c, a).All(x => x)), "AND");
        Reg((c, a) => B(Logicals(c, a).Any(x => x)), "OR");
        Reg((c, a) => B(Logicals(c, a).Count(x => x) % 2 == 1), "XOR");
        Reg((c, a) => { Args(a, 1, 1); return B(!c.Bool(a[0])); }, "NOT");
        Reg((c, a) => B(true), "TRUE");
        Reg((c, a) => B(false), "FALSE");
        Reg((c, a) =>
        {
            Args(a, 3);
            var v = c.Scalar(a[0]);
            int i = 1;
            for (; i + 1 < a.Length; i += 2)
                if (Calc.Compare(v, c.Scalar(a[i])) == 0) return c.Eval(a[i + 1]);
            if (i < a.Length) return c.Eval(a[i]);
            throw Err(Errors.NA);
        }, "SWITCH");
        Reg((c, a) =>
        {
            Args(a, 2);
            int i = (int)c.Num(a[0]);
            if (i < 1 || i >= a.Length) throw Err(Errors.Value);
            return c.Eval(a[i]);
        }, "CHOOSE");
    }

    private static List<bool> Logicals(Ctx c, Expr[] a)
    {
        var res = new List<bool>();
        foreach (var e in a)
        {
            var v = c.Eval(e);
            if (v is CellValue s) { if (!s.IsBlank) res.Add(Calc.ToBool(s)); continue; }
            foreach (var x in FlatVal(c, v))
            {
                if (x.IsError) throw Err(x.Str!);
                if (x.IsNumber || x.IsBool) res.Add(x.Num != 0);
            }
        }
        if (res.Count == 0) throw Err(Errors.Value);
        return res;
    }
}
