using System.Globalization;
using System.Text;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

/// <summary>LET / LAMBDA and its helpers, dynamic-array functions (FILTER, SORT, UNIQUE, VSTACK…), GROUPBY / PIVOTBY.</summary>
public static partial class Functions
{
    // ================================================================== element-wise evaluation ("lifting")

    /// <summary>
    /// Functions whose arguments are single values: given an array or a multi-cell range for such an argument they
    /// are evaluated once per element and return an array (=LEN(A1:A5), SUMPRODUCT(--(LEFT(A1:A9,2)="AB"))).
    /// Value: which argument positions are lifted.
    /// </summary>
    private static readonly Dictionary<string, Func<int, bool>> Lift = new(StringComparer.OrdinalIgnoreCase);

    private static void RegisterLift(Func<int, bool> which, string names)
    {
        foreach (var n in names.Split(' ', StringSplitOptions.RemoveEmptyEntries)) Lift[n] = which;
    }

    private static void RegisterLifting()
    {
        Func<int, bool> all = _ => true, first = i => i == 0, firstTwo = i => i < 2, second = i => i == 1;
        RegisterLift(all, @"ABS ACOS ACOSH ACOT ACOTH ASIN ASINH ATAN ATAN2 ATANH BASE CEILING CEILING.MATH CEILING.PRECISE COMBIN COMBINA COS COSH COT COTH
            CSC CSCH DECIMAL DEGREES EVEN EXP FACT FACTDOUBLE FLOOR FLOOR.MATH FLOOR.PRECISE INT ISO.CEILING LN LOG LOG10 MOD MROUND ODD POWER QUOTIENT
            RADIANS ROMAN ARABIC ROUND ROUNDDOWN ROUNDUP SEC SECH SIGN SIN SINH SQRT SQRTPI TAN TANH TRUNC PERMUT PERMUTATIONA
            CHAR CLEAN CODE DOLLAR EXACT FIND FINDB FIXED LEFT LEFTB LEN LENB LOWER MID MIDB NUMBERVALUE PROPER REPLACE REPLACEB REPT RIGHT RIGHTB
            SEARCH SEARCHB SUBSTITUTE T TEXT TRIM UNICHAR UNICODE UPPER VALUE ASC DBCS JIS BAHTTEXT VND DOCSO ENCODEURL REGEXTEST REGEXREPLACE PHONETIC
            DATE DATEVALUE DAY DAYS DAYS360 EDATE EOMONTH HOUR ISOWEEKNUM MINUTE MONTH SECOND TIME TIMEVALUE WEEKDAY WEEKNUM YEAR YEARFRAC DATEDIF
            ISBLANK ISERR ISERROR ISEVEN ISODD ISLOGICAL ISNA ISNONTEXT ISNUMBER ISTEXT N ERROR.TYPE NOT
            BESSELI BESSELJ BESSELK BESSELY BIN2DEC BIN2HEX BIN2OCT BITAND BITLSHIFT BITOR BITRSHIFT BITXOR COMPLEX CONVERT DEC2BIN DEC2HEX DEC2OCT DELTA
            ERF ERF.PRECISE ERFC ERFC.PRECISE GESTEP HEX2BIN HEX2DEC HEX2OCT IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV
            IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMREAL IMSEC IMSECH IMSIN IMSINH IMSQRT IMSUB IMTAN OCT2BIN OCT2DEC OCT2HEX
            BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CONFIDENCE.NORM CONFIDENCE.T
            EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT FISHER FISHERINV GAMMA GAMMA.DIST GAMMA.INV GAMMALN GAMMALN.PRECISE GAUSS HYPGEOM.DIST
            LOGNORM.DIST LOGNORM.INV NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PHI POISSON.DIST STANDARDIZE T.DIST T.DIST.2T T.DIST.RT
            T.INV T.INV.2T WEIBULL.DIST BETADIST BETAINV BINOMDIST CHIDIST CHIINV CONFIDENCE CRITBINOM EXPONDIST FDIST FINV GAMMADIST GAMMAINV
            HYPGEOMDIST LOGINV LOGNORMDIST NEGBINOMDIST NORMDIST NORMINV NORMSDIST NORMSINV POISSON TDIST TINV WEIBULL
            PMT FV PV IPMT PPMT NPER RATE SLN SYD DDB DB VDB ISPMT EFFECT NOMINAL PDURATION RRI DOLLARDE DOLLARFR CUMIPMT CUMPRINC
            ACCRINT ACCRINTM COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD DISC DURATION INTRATE MDURATION PRICE PRICEDISC PRICEMAT RECEIVED
            TBILLEQ TBILLPRICE TBILLYIELD YIELD YIELDDISC YIELDMAT ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD AMORDEGRC AMORLINC");
        RegisterLift(first, "VLOOKUP HLOOKUP MATCH XMATCH XLOOKUP");
        RegisterLift(i => i != 1, "TEXTBEFORE TEXTAFTER");
        RegisterLift(firstTwo, "NETWORKDAYS WORKDAY");
        RegisterLift(i => i < 3, "NETWORKDAYS.INTL WORKDAY.INTL");
        RegisterLift(second, "COUNTIF SUMIF AVERAGEIF");
        RegisterLift(i => i % 2 == 1, "COUNTIFS");
        RegisterLift(i => i > 0 && i % 2 == 0, "SUMIFS AVERAGEIFS MAXIFS MINIFS");
        RegisterLift(i => i > 0, "LARGE SMALL PERCENTILE PERCENTILE.INC PERCENTILE.EXC QUARTILE QUARTILE.INC QUARTILE.EXC");
        RegisterLift(first, "RANK RANK.EQ RANK.AVG");
    }

    /// <summary>Calls a worksheet function, evaluating it per element when single-value arguments receive arrays.</summary>
    internal static object Call(string name, FnImpl impl, Ctx c, Expr[] a)
    {
        if (a.Length == 0 || !Lift.TryGetValue(name, out var lifted)) return impl(c, a);
        var vals = new Expr[a.Length];
        bool any = false;
        for (int i = 0; i < a.Length; i++)
        {
            if (a[i] is MissingExpr) { vals[i] = a[i]; continue; }
            object v;
            // an argument that fails (#NAME?, #REF!…) is passed on as an error value: ISERROR(NoSuchName) is TRUE
            try { v = a[i] is ValueExpr ve0 ? ve0.Value : c.Eval(a[i]); }
            catch (FormulaError e) { v = CellValue.Error(e.Code); }
            // files without dynamic arrays: a range given to a single-value argument is intersected with the formula's row/column
            if (c.Implicit && lifted(i) && v is RangeVal { Range.IsSingle: false }) v = c.Calc.ToScalar(v, c);
            if (lifted(i) && IsMulti(v)) any = true;
            vals[i] = a[i] is ValueExpr && ReferenceEquals(v, ((ValueExpr)a[i]).Value) ? a[i] : new ValueExpr(v);
        }
        if (!any) return impl(c, vals);
        var grids = new CellValue[a.Length][,];
        int rows = 1, cols = 1;
        for (int i = 0; i < a.Length; i++)
        {
            if (!lifted(i) || vals[i] is not ValueExpr ve || !IsMulti(ve.Value)) continue;
            grids[i] = c.Calc.ToArray(ve.Value);
            rows = Math.Max(rows, grids[i].GetLength(0));
            cols = Math.Max(cols, grids[i].GetLength(1));
        }
        var res = new CellValue[rows, cols];
        var args = new Expr[a.Length];
        for (int r = 0; r < rows; r++)
            for (int k = 0; k < cols; k++)
            {
                bool na = false;
                for (int i = 0; i < a.Length; i++)
                {
                    var g = grids[i];
                    if (g == null) { args[i] = vals[i]; continue; }
                    int gr = g.GetLength(0) == 1 ? 0 : r, gc = g.GetLength(1) == 1 ? 0 : k;
                    if (gr >= g.GetLength(0) || gc >= g.GetLength(1)) { na = true; break; }
                    args[i] = new ValueExpr(g[gr, gc]);
                }
                if (na) { res[r, k] = CellValue.Error(Errors.NA); continue; }
                try
                {
                    var v = impl(c, args);
                    res[r, k] = v is CellValue cv ? cv : v is LambdaVal ? CellValue.Error(Errors.Calc) : c.Calc.ToArray(v) is { Length: > 0 } arr ? arr[0, 0] : CellValue.Error(Errors.Value);
                }
                catch (FormulaError e) { res[r, k] = CellValue.Error(e.Code); }
                catch (OverflowException) { res[r, k] = CellValue.Error(Errors.Num); }
                catch (ArgumentException) { res[r, k] = CellValue.Error(Errors.Value); }
                catch (IndexOutOfRangeException) { res[r, k] = CellValue.Error(Errors.Value); }
            }
        return new ArrayVal(res);
    }

    private static bool IsMulti(object v) => v is ArrayVal || v is RangeVal { Range.IsSingle: false };

    // ================================================================== helpers

    private static CellValue[,] Arr(Ctx c, Expr e)
    {
        var v = c.Eval(e);
        if (v is LambdaVal) throw Err(Errors.Calc);
        if (v is CellValue { IsError: true } err && e is not RefExpr) throw Err(err.Str!);
        return c.Calc.ToArray(v);
    }

    private static LambdaVal Fn(Ctx c, Expr e) => c.Eval(e) as LambdaVal ?? throw Err(Errors.Value);

    private static ArrayVal Out(CellValue[,] g) => new(g);

    private static CellValue OneValue(Ctx c, object v)
    {
        switch (v)
        {
            case CellValue s: return s;
            case LambdaVal: return CellValue.Error(Errors.Calc);
            default:
                var g = c.Calc.ToArray(v);
                return g.Length == 1 ? g[0, 0] : CellValue.Error(Errors.Calc);
        }
    }

    private static CellValue[,] Row(CellValue[,] g, int r)
    {
        var res = new CellValue[1, g.GetLength(1)];
        for (int k = 0; k < g.GetLength(1); k++) res[0, k] = g[r, k];
        return res;
    }

    private static CellValue[,] Col(CellValue[,] g, int k)
    {
        var res = new CellValue[g.GetLength(0), 1];
        for (int r = 0; r < g.GetLength(0); r++) res[r, 0] = g[r, k];
        return res;
    }

    private static CellValue[,] Transposed(CellValue[,] g)
    {
        var t = new CellValue[g.GetLength(1), g.GetLength(0)];
        for (int i = 0; i < g.GetLength(0); i++) for (int j = 0; j < g.GetLength(1); j++) t[j, i] = g[i, j];
        return t;
    }

    private static CellValue[,] FromRows(List<CellValue[]> rows, int cols)
    {
        var res = new CellValue[rows.Count, cols];
        for (int r = 0; r < rows.Count; r++) for (int k = 0; k < cols; k++) res[r, k] = k < rows[r].Length ? rows[r][k] : CellValue.Error(Errors.NA);
        return res;
    }

    private static CellValue[] RowOf(CellValue[,] g, int r)
    {
        var res = new CellValue[g.GetLength(1)];
        for (int k = 0; k < res.Length; k++) res[k] = g[r, k];
        return res;
    }

    /// <summary>Case-insensitive identity of a value (UNIQUE, GROUPBY).</summary>
    private static string KeyOf(CellValue v) => v.Kind switch
    {
        ValueKind.Number => "n" + v.Num.ToString("R", CultureInfo.InvariantCulture),
        ValueKind.Text => "s" + v.Str!.ToUpperInvariant(),
        ValueKind.Bool => "b" + v.Num,
        ValueKind.Error => "e" + v.Str,
        _ => "_"
    };

    private static string RowKey(CellValue[] row) => string.Join("\u0001", row.Select(KeyOf));

    /// <summary>Ordering used by SORT/SORTBY: numbers &lt; text &lt; logicals &lt; errors, blanks last.</summary>
    private static int SortCompare(CellValue x, CellValue y)
    {
        if (x.IsBlank || y.IsBlank) return x.IsBlank && y.IsBlank ? 0 : x.IsBlank ? 1 : -1;
        if (x.IsError || y.IsError) return x.IsError && y.IsError ? 0 : x.IsError ? 1 : -1;
        return Calc.Compare(x, y);
    }

    private static List<int> IntList(Ctx c, Expr e)
    {
        var res = new List<int>();
        foreach (var v in Arr(c, e))
        {
            if (v.IsError) throw Err(v.Str!);
            res.Add((int)Math.Truncate(Calc.ToNum(v)));
        }
        return res;
    }

    // ================================================================== registration

    private static void RegisterDyn()
    {
        RegisterLifting();

        // ---------------------------------------------------------- LET / LAMBDA
        Reg((c, a) =>
        {
            if (a.Length < 3 || a.Length % 2 == 0) throw Err(Errors.Value);
            var ctx = c;
            for (int i = 0; i + 1 < a.Length; i += 2)
            {
                if (a[i] is not NameExpr n || n.Name.StartsWith('[')) throw Err(Errors.Value);
                ctx = ctx.Bind(n.Name, new ValueExpr(ctx.Eval(a[i + 1])));
            }
            return ctx.Eval(a[^1]);
        }, "LET");
        Reg((c, a) =>
        {
            if (a.Length == 0) throw Err(Errors.Value);
            var names = new string[a.Length - 1];
            var opt = new bool[a.Length - 1];
            bool seenOptional = false;
            for (int i = 0; i < names.Length; i++)
            {
                if (a[i] is not NameExpr n) throw Err(Errors.Value);
                var nm = n.Name;
                if (nm.StartsWith('[') && nm.EndsWith(']')) { nm = nm[1..^1]; opt[i] = true; seenOptional = true; }
                else if (seenOptional) throw Err(Errors.Value); // optional parameters come last
                names[i] = nm;
            }
            return new LambdaVal { Params = names, Optional = opt, Body = a[^1], Closure = c };
        }, "LAMBDA");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            if (a[0] is NameExpr n && c.Scope?.Find(n.Name) is Expr bound) return B(bound is MissingExpr);
            return B(a[0] is MissingExpr);
        }, "ISOMITTED");
        Reg((c, a) =>
        {
            Args(a, 2);
            var f = Fn(c, a[^1]);
            var grids = a[..^1].Select(e => Arr(c, e)).ToArray();
            int rows = grids.Max(g => g.GetLength(0)), cols = grids.Max(g => g.GetLength(1));
            var res = new CellValue[rows, cols];
            var args = new Expr[grids.Length];
            for (int r = 0; r < rows; r++)
                for (int k = 0; k < cols; k++)
                {
                    bool na = false;
                    for (int i = 0; i < grids.Length; i++)
                    {
                        var g = grids[i];
                        int gr = g.GetLength(0) == 1 ? 0 : r, gc = g.GetLength(1) == 1 ? 0 : k;
                        if (gr >= g.GetLength(0) || gc >= g.GetLength(1)) { na = true; break; }
                        args[i] = new ValueExpr(g[gr, gc]);
                    }
                    res[r, k] = na ? CellValue.Error(Errors.NA) : SafeCall(c, f, args);
                }
            return Out(res);
        }, "MAP");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            object acc = a.Length == 3 && a[0] is not MissingExpr ? c.Eval(a[0]) : CellValue.Blank;
            var g = Arr(c, a.Length == 3 ? a[1] : a[0]);
            var f = Fn(c, a[^1]);
            for (int r = 0; r < g.GetLength(0); r++)
                for (int k = 0; k < g.GetLength(1); k++)
                    acc = c.Calc.Invoke(f, [new ValueExpr(acc), new ValueExpr(g[r, k])], c);
            return acc;
        }, "REDUCE");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            object acc = a.Length == 3 && a[0] is not MissingExpr ? c.Eval(a[0]) : CellValue.Blank;
            var g = Arr(c, a.Length == 3 ? a[1] : a[0]);
            var f = Fn(c, a[^1]);
            var res = new CellValue[g.GetLength(0), g.GetLength(1)];
            for (int r = 0; r < g.GetLength(0); r++)
                for (int k = 0; k < g.GetLength(1); k++)
                {
                    acc = c.Calc.Invoke(f, [new ValueExpr(acc), new ValueExpr(g[r, k])], c);
                    res[r, k] = OneValue(c, acc);
                }
            return Out(res);
        }, "SCAN");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var g = Arr(c, a[0]);
            var f = Fn(c, a[1]);
            var res = new CellValue[g.GetLength(0), 1];
            for (int r = 0; r < g.GetLength(0); r++) res[r, 0] = SafeCall(c, f, [new ValueExpr(new ArrayVal(Row(g, r)))]);
            return Out(res);
        }, "BYROW");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var g = Arr(c, a[0]);
            var f = Fn(c, a[1]);
            var res = new CellValue[1, g.GetLength(1)];
            for (int k = 0; k < g.GetLength(1); k++) res[0, k] = SafeCall(c, f, [new ValueExpr(new ArrayVal(Col(g, k)))]);
            return Out(res);
        }, "BYCOL");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            int rows = c.Int(a[0]), cols = c.Int(a[1]);
            if (rows < 1 || cols < 1) throw Err(Errors.Value);
            if ((long)rows * cols > 4_000_000) throw Err(Errors.Num);
            var f = Fn(c, a[2]);
            var res = new CellValue[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int k = 0; k < cols; k++) res[r, k] = SafeCall(c, f, [new ValueExpr(N(r + 1)), new ValueExpr(N(k + 1))]);
            return Out(res);
        }, "MAKEARRAY");

        // ---------------------------------------------------------- FILTER / SORT / UNIQUE
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var g = Arr(c, a[0]);
            var inc = Arr(c, a[1]);
            int rows = g.GetLength(0), cols = g.GetLength(1);
            bool byRow;
            if (inc.GetLength(1) == 1 && inc.GetLength(0) == rows) byRow = true;
            else if (inc.GetLength(0) == 1 && inc.GetLength(1) == cols) byRow = false;
            else throw Err(Errors.Value);
            int n = byRow ? rows : cols;
            var keep = new List<int>();
            for (int i = 0; i < n; i++)
            {
                var v = byRow ? inc[i, 0] : inc[0, i];
                if (v.IsError) throw Err(v.Str!);
                if (v.IsBlank) continue;
                if (Calc.ToBool(v.IsText ? (ValueParser.TryCoerceNumber(v.Str!, out var d) ? N(d) : v) : v)) keep.Add(i);
            }
            if (keep.Count == 0)
            {
                if (a.Length > 2 && a[2] is not MissingExpr) return c.Eval(a[2]);
                throw Err(Errors.Calc);
            }
            var res = byRow ? new CellValue[keep.Count, cols] : new CellValue[rows, keep.Count];
            for (int i = 0; i < keep.Count; i++)
                if (byRow) for (int k = 0; k < cols; k++) res[i, k] = g[keep[i], k];
                else for (int r = 0; r < rows; r++) res[r, i] = g[r, keep[i]];
            return Out(res);
        }, "FILTER");
        Reg((c, a) =>
        {
            Args(a, 1, 4);
            var g = Arr(c, a[0]);
            bool byCol = c.BoolOr(a, 3, false);
            if (byCol) g = Transposed(g);
            var idx = a.Length > 1 && a[1] is not MissingExpr ? IntList(c, a[1]) : [1];
            var ord = a.Length > 2 && a[2] is not MissingExpr ? IntList(c, a[2]) : [1];
            if (ord.Count != 1 && ord.Count != idx.Count) throw Err(Errors.Value);
            foreach (var i in idx) if (i < 1 || i > g.GetLength(1)) throw Err(Errors.Value);
            foreach (var o in ord) if (o != 1 && o != -1) throw Err(Errors.Value);
            var keys = idx.Select((col, i) => (Col: col - 1, Desc: (ord.Count == 1 ? ord[0] : ord[i]) == -1)).ToList();
            var order = Enumerable.Range(0, g.GetLength(0)).ToList();
            var sorted = order.OrderBy(r => r, Comparer<int>.Create((x, y) =>
            {
                foreach (var (col, desc) in keys)
                {
                    int cmp = SortCompare(g[x, col], g[y, col]);
                    if (cmp != 0) return desc && !g[x, col].IsBlank && !g[y, col].IsBlank ? -cmp : cmp;
                }
                return 0;
            })).ToList();
            var res = new CellValue[g.GetLength(0), g.GetLength(1)];
            for (int i = 0; i < sorted.Count; i++) for (int k = 0; k < g.GetLength(1); k++) res[i, k] = g[sorted[i], k];
            return Out(byCol ? Transposed(res) : res);
        }, "SORT");
        Reg((c, a) =>
        {
            Args(a, 2);
            var g = Arr(c, a[0]);
            var keys = new List<(CellValue[] Vals, bool Desc)>();
            bool? horizontal = null;
            for (int i = 1; i < a.Length; i += 2)
            {
                var by = Arr(c, a[i]);
                bool desc = false;
                if (i + 1 < a.Length && a[i + 1] is not MissingExpr)
                {
                    int o = c.Int(a[i + 1]);
                    if (o != 1 && o != -1) throw Err(Errors.Value);
                    desc = o == -1;
                }
                bool h;
                if (by.GetLength(1) == 1 && by.GetLength(0) == g.GetLength(0)) h = false;
                else if (by.GetLength(0) == 1 && by.GetLength(1) == g.GetLength(1)) h = true;
                else throw Err(Errors.Value);
                if (horizontal != null && horizontal != h) throw Err(Errors.Value);
                horizontal = h;
                keys.Add((h ? RowOf(by, 0) : RowOf(Transposed(by), 0), desc));
            }
            bool hz = horizontal == true;
            if (hz) g = Transposed(g);
            var sorted = Enumerable.Range(0, g.GetLength(0)).OrderBy(r => r, Comparer<int>.Create((x, y) =>
            {
                foreach (var (vals, desc) in keys)
                {
                    int cmp = SortCompare(vals[x], vals[y]);
                    if (cmp != 0) return desc && !vals[x].IsBlank && !vals[y].IsBlank ? -cmp : cmp;
                }
                return 0;
            })).ToList();
            var res = new CellValue[g.GetLength(0), g.GetLength(1)];
            for (int i = 0; i < sorted.Count; i++) for (int k = 0; k < g.GetLength(1); k++) res[i, k] = g[sorted[i], k];
            return Out(hz ? Transposed(res) : res);
        }, "SORTBY");
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            var g = Arr(c, a[0]);
            bool byCol = c.BoolOr(a, 1, false), once = c.BoolOr(a, 2, false);
            if (byCol) g = Transposed(g);
            var rows = new List<CellValue[]>();
            var count = new Dictionary<string, int>();
            var firstIndex = new Dictionary<string, int>();
            for (int r = 0; r < g.GetLength(0); r++)
            {
                var row = RowOf(g, r);
                var k = RowKey(row);
                if (count.TryGetValue(k, out var n)) { count[k] = n + 1; continue; }
                count[k] = 1;
                firstIndex[k] = rows.Count;
                rows.Add(row);
            }
            if (once) rows = rows.Where(row => count[RowKey(row)] == 1).ToList();
            if (rows.Count == 0) throw Err(Errors.Calc);
            var res = FromRows(rows, g.GetLength(1));
            return Out(byCol ? Transposed(res) : res);
        }, "UNIQUE");

        // ---------------------------------------------------------- reshaping
        Reg((c, a) => Stack(c, a, vertical: true), "VSTACK");
        Reg((c, a) => Stack(c, a, vertical: false), "HSTACK");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var g = Arr(c, a[0]);
            int R = g.GetLength(0), C = g.GetLength(1);
            int rows = a[1] is MissingExpr ? R : c.Int(a[1]), cols = a.Length > 2 && a[2] is not MissingExpr ? c.Int(a[2]) : C;
            if (rows == 0 || cols == 0) throw Err(Errors.Calc);
            int r0 = rows > 0 ? 0 : Math.Max(0, R + rows), r1 = rows > 0 ? Math.Min(R, rows) : R;
            int c0 = cols > 0 ? 0 : Math.Max(0, C + cols), c1 = cols > 0 ? Math.Min(C, cols) : C;
            return Out(Sub(g, r0, r1, c0, c1));
        }, "TAKE");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var g = Arr(c, a[0]);
            int R = g.GetLength(0), C = g.GetLength(1);
            int rows = a[1] is MissingExpr ? 0 : c.Int(a[1]), cols = a.Length > 2 && a[2] is not MissingExpr ? c.Int(a[2]) : 0;
            int r0 = rows > 0 ? rows : 0, r1 = rows < 0 ? R + rows : R;
            int c0 = cols > 0 ? cols : 0, c1 = cols < 0 ? C + cols : C;
            if (r0 >= r1 || c0 >= c1) throw Err(Errors.Calc);
            return Out(Sub(g, r0, r1, c0, c1));
        }, "DROP");
        Reg((c, a) =>
        {
            Args(a, 2);
            var g = Arr(c, a[0]);
            var idx = a[1..].SelectMany(e => IntList(c, e)).ToList();
            int C = g.GetLength(1);
            var res = new CellValue[g.GetLength(0), idx.Count];
            for (int i = 0; i < idx.Count; i++)
            {
                int k = idx[i] > 0 ? idx[i] - 1 : C + idx[i];
                if (idx[i] == 0 || k < 0 || k >= C) throw Err(Errors.Value);
                for (int r = 0; r < g.GetLength(0); r++) res[r, i] = g[r, k];
            }
            return Out(res);
        }, "CHOOSECOLS");
        Reg((c, a) =>
        {
            Args(a, 2);
            var g = Arr(c, a[0]);
            var idx = a[1..].SelectMany(e => IntList(c, e)).ToList();
            int R = g.GetLength(0);
            var res = new CellValue[idx.Count, g.GetLength(1)];
            for (int i = 0; i < idx.Count; i++)
            {
                int r = idx[i] > 0 ? idx[i] - 1 : R + idx[i];
                if (idx[i] == 0 || r < 0 || r >= R) throw Err(Errors.Value);
                for (int k = 0; k < g.GetLength(1); k++) res[i, k] = g[r, k];
            }
            return Out(res);
        }, "CHOOSEROWS");
        Reg((c, a) => ToLine(c, a, column: true), "TOCOL");
        Reg((c, a) => ToLine(c, a, column: false), "TOROW");
        Reg((c, a) => Wrap(c, a, rowsFirst: true), "WRAPROWS");
        Reg((c, a) => Wrap(c, a, rowsFirst: false), "WRAPCOLS");
        Reg((c, a) =>
        {
            Args(a, 2, 4);
            var g = Arr(c, a[0]);
            int R = g.GetLength(0), C = g.GetLength(1);
            int rows = a[1] is MissingExpr ? R : c.Int(a[1]), cols = a.Length > 2 && a[2] is not MissingExpr ? c.Int(a[2]) : C;
            if (rows < R || cols < C) throw Err(Errors.Value);
            if ((long)rows * cols > 4_000_000) throw Err(Errors.Num);
            var pad = a.Length > 3 && a[3] is not MissingExpr ? c.Scalar(a[3]) : CellValue.Error(Errors.NA);
            var res = new CellValue[rows, cols];
            for (int r = 0; r < rows; r++) for (int k = 0; k < cols; k++) res[r, k] = r < R && k < C ? g[r, k] : pad;
            return Out(res);
        }, "EXPAND");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            return c.Eval(a[0]) is RangeVal ? N(1) : throw Err(Errors.Value);
        }, "AREAS");
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            if (c.Eval(a[0]) is not RangeVal rv) throw Err(Errors.Value);
            int tr = (int)c.NumOr(a, 1, 3), tc = (int)c.NumOr(a, 2, 3);
            if (tr is < 0 or > 3 || tc is < 0 or > 3) throw Err(Errors.Value);
            var rg = c.Calc.Clip(rv);
            int r1 = rg.R1, r2 = rg.R2, c1 = rg.C1, c2 = rg.C2;
            bool RowEmpty(int r) { for (int k = c1; k <= c2; k++) if (!c.Calc.ValueOf(rv.Sheet, r, k).IsBlank) return false; return true; }
            bool ColEmpty(int k) { for (int r = r1; r <= r2; r++) if (!c.Calc.ValueOf(rv.Sheet, r, k).IsBlank) return false; return true; }
            if ((tr & 1) != 0) while (r1 < r2 && RowEmpty(r1)) r1++;
            if ((tr & 2) != 0) while (r2 > r1 && RowEmpty(r2)) r2--;
            if ((tc & 1) != 0) while (c1 < c2 && ColEmpty(c1)) c1++;
            if ((tc & 2) != 0) while (c2 > c1 && ColEmpty(c2)) c2--;
            return new RangeVal(rv.Sheet, new CellRange(r1, c1, r2, c2));
        }, "TRIMRANGE");

        // ---------------------------------------------------------- GROUPBY / PIVOTBY
        Reg(GroupBy, "GROUPBY");
        Reg(PivotBy, "PIVOTBY");
    }

    private static CellValue SafeCall(Ctx c, LambdaVal f, Expr[] args)
    {
        try { return OneValue(c, c.Calc.Invoke(f, args, c)); }
        catch (FormulaError e) { return CellValue.Error(e.Code); }
    }

    private static CellValue[,] Sub(CellValue[,] g, int r0, int r1, int c0, int c1)
    {
        var res = new CellValue[r1 - r0, c1 - c0];
        for (int r = r0; r < r1; r++) for (int k = c0; k < c1; k++) res[r - r0, k - c0] = g[r, k];
        return res;
    }

    private static object Stack(Ctx c, Expr[] a, bool vertical)
    {
        Args(a, 1);
        var parts = a.Select(e => Arr(c, e)).ToList();
        if (!vertical) parts = parts.Select(Transposed).ToList();
        int cols = parts.Max(p => p.GetLength(1)), rows = parts.Sum(p => p.GetLength(0));
        var res = new CellValue[rows, cols];
        int at = 0;
        foreach (var p in parts)
        {
            for (int r = 0; r < p.GetLength(0); r++)
                for (int k = 0; k < cols; k++) res[at + r, k] = k < p.GetLength(1) ? p[r, k] : CellValue.Error(Errors.NA);
            at += p.GetLength(0);
        }
        return Out(vertical ? res : Transposed(res));
    }

    private static object ToLine(Ctx c, Expr[] a, bool column)
    {
        Args(a, 1, 3);
        var g = Arr(c, a[0]);
        int ignore = (int)c.NumOr(a, 1, 0);
        bool byCol = c.BoolOr(a, 2, false);
        if (ignore is < 0 or > 3) throw Err(Errors.Value);
        var list = new List<CellValue>();
        int R = g.GetLength(0), C = g.GetLength(1);
        for (int i = 0; i < (byCol ? C : R); i++)
            for (int j = 0; j < (byCol ? R : C); j++)
            {
                var v = byCol ? g[j, i] : g[i, j];
                if ((ignore & 1) != 0 && v.IsBlank) continue;
                if ((ignore & 2) != 0 && v.IsError) continue;
                list.Add(v);
            }
        if (list.Count == 0) throw Err(Errors.Calc);
        var res = column ? new CellValue[list.Count, 1] : new CellValue[1, list.Count];
        for (int i = 0; i < list.Count; i++) if (column) res[i, 0] = list[i]; else res[0, i] = list[i];
        return Out(res);
    }

    private static object Wrap(Ctx c, Expr[] a, bool rowsFirst)
    {
        Args(a, 2, 3);
        var g = Arr(c, a[0]);
        if (g.GetLength(0) != 1 && g.GetLength(1) != 1) throw Err(Errors.Value);
        var list = g.Cast<CellValue>().ToList();
        int n = c.Int(a[1]);
        if (n < 1) throw Err(Errors.Num);
        var pad = a.Length > 2 && a[2] is not MissingExpr ? c.Scalar(a[2]) : CellValue.Error(Errors.NA);
        int other = (list.Count + n - 1) / n;
        var res = rowsFirst ? new CellValue[other, n] : new CellValue[n, other];
        for (int i = 0; i < other * n; i++)
        {
            var v = i < list.Count ? list[i] : pad;
            if (rowsFirst) res[i / n, i % n] = v; else res[i % n, i / n] = v;
        }
        return Out(res);
    }

    // ---------------------------------------------------------- GROUPBY / PIVOTBY

    private sealed class GroupTree
    {
        public readonly Dictionary<string, (CellValue[] Key, List<int> Rows)> Groups = new();
        public readonly List<string> Order = new();
        public void Add(CellValue[] key, int row)
        {
            var k = RowKey(key);
            if (!Groups.TryGetValue(k, out var g)) { g = (key, new List<int>()); Groups[k] = g; Order.Add(k); }
            g.Rows.Add(row);
        }
    }

    /// <summary>Header detection: 0 none, 1 has headers (hidden), 2 no headers but generate, 3 has headers and show.</summary>
    private static int HeaderMode(Ctx c, Expr[] a, int i, CellValue[,] values)
    {
        if (i < a.Length && a[i] is not MissingExpr) { int m = c.Int(a[i]); if (m is < 0 or > 3) throw Err(Errors.Value); return m; }
        // automatic: the first row of values is text and the second is a number → headers present
        if (values.GetLength(0) >= 2 && values[0, 0].IsText && values[1, 0].IsNumber) return 1;
        return 0;
    }

    private static CellValue Aggregate(Ctx c, LambdaVal f, CellValue[,] values, int col, List<int> rows)
    {
        var arr = new CellValue[rows.Count, 1];
        for (int i = 0; i < rows.Count; i++) arr[i, 0] = values[rows[i], col];
        return SafeCall(c, f, [new ValueExpr(new ArrayVal(arr))]);
    }

    private static string FnLabel(LambdaVal f) => f.Builtin != null ? f.Builtin[..1] + f.Builtin[1..].ToLowerInvariant() : "Value";

    private static object GroupBy(Ctx c, Expr[] a)
    {
        Args(a, 3, 8);
        var keys = Arr(c, a[0]);
        var values = Arr(c, a[1]);
        var f = Fn(c, a[2]);
        if (keys.GetLength(0) != values.GetLength(0)) throw Err(Errors.Value);
        int headers = HeaderMode(c, a, 3, values);
        int depth = (int)c.NumOr(a, 4, 1);
        if (Math.Abs(depth) > 2) throw Err(Errors.Value);
        var sortSpec = a.Length > 5 && a[5] is not MissingExpr ? IntList(c, a[5]) : [];
        var filter = a.Length > 6 && a[6] is not MissingExpr ? Arr(c, a[6]) : null;
        int start = headers is 1 or 3 ? 1 : 0;
        int nk = keys.GetLength(1), nv = values.GetLength(1);
        var tree = new GroupTree();
        for (int r = start; r < keys.GetLength(0); r++)
        {
            if (filter != null)
            {
                var fv = filter[Math.Min(r, filter.GetLength(0) - 1), 0];
                if (fv.IsError) throw Err(fv.Str!);
                if (!Calc.ToBool(fv)) continue;
            }
            tree.Add(RowOf(keys, r), r);
        }
        var groups = tree.Order.Select(k => tree.Groups[k]).ToList();
        // aggregate rows
        var rowsOut = new List<(CellValue[] Key, CellValue[] Vals, int Level)>();
        foreach (var g in groups)
            rowsOut.Add((g.Key, Enumerable.Range(0, nv).Select(k => Aggregate(c, f, values, k, g.Rows)).ToArray(), 0));
        // sort
        Comparison<(CellValue[] Key, CellValue[] Vals, int Level)> cmp = (x, y) =>
        {
            if (sortSpec.Count == 0)
            {
                for (int k = 0; k < nk; k++) { int d = SortCompare(x.Key[k], y.Key[k]); if (d != 0) return d; }
                return 0;
            }
            foreach (var s in sortSpec)
            {
                int col = Math.Abs(s) - 1;
                if (col < 0 || col >= nk + nv) throw Err(Errors.Value);
                var xv = col < nk ? x.Key[col] : x.Vals[col - nk];
                var yv = col < nk ? y.Key[col] : y.Vals[col - nk];
                int d = SortCompare(xv, yv);
                if (d != 0) return s < 0 ? -d : d;
            }
            return 0;
        };
        rowsOut = rowsOut.OrderBy(x => x, Comparer<(CellValue[] Key, CellValue[] Vals, int Level)>.Create(cmp)).ToList();
        var lines = new List<CellValue[]>();
        if (headers is 2 or 3)
        {
            var h = new CellValue[nk + nv];
            for (int k = 0; k < nk; k++) h[k] = headers == 3 ? keys[0, k] : S(nk == 1 ? "Row Field" : "Row Field " + (k + 1));
            for (int k = 0; k < nv; k++) h[nk + k] = headers == 3 ? values[0, k] : S(nv == 1 ? FnLabel(f) : FnLabel(f) + " " + (k + 1));
            lines.Add(h);
        }
        var body = new List<CellValue[]>();
        var allRows = groups.SelectMany(g => g.Rows).ToList();
        if (Math.Abs(depth) == 2 && nk > 1)
        {
            // subtotals by the first row field
            foreach (var grp in rowsOut.GroupBy(x => KeyOf(x.Key[0])))
            {
                var items = grp.ToList();
                var sub = new CellValue[nk + nv];
                sub[0] = items[0].Key[0];
                for (int k = 1; k < nk; k++) sub[k] = CellValue.Blank;
                var subRows = groups.Where(g => KeyOf(g.Key[0]) == grp.Key).SelectMany(g => g.Rows).ToList();
                for (int k = 0; k < nv; k++) sub[nk + k] = Aggregate(c, f, values, k, subRows);
                if (depth < 0) body.Add(sub);
                body.AddRange(items.Select(x => x.Key.Concat(x.Vals).ToArray()));
                if (depth > 0) body.Add(sub);
            }
        }
        else body.AddRange(rowsOut.Select(x => x.Key.Concat(x.Vals).ToArray()));
        CellValue[]? total = null;
        if (depth != 0)
        {
            total = new CellValue[nk + nv];
            total[0] = S("Total");
            for (int k = 1; k < nk; k++) total[k] = CellValue.Blank;
            for (int k = 0; k < nv; k++) total[nk + k] = Aggregate(c, f, values, k, allRows);
        }
        if (total != null && depth < 0) lines.Add(total);
        lines.AddRange(body);
        if (total != null && depth > 0) lines.Add(total);
        if (lines.Count == 0) throw Err(Errors.Calc);
        return Out(FromRows(lines, nk + nv));
    }

    private static object PivotBy(Ctx c, Expr[] a)
    {
        Args(a, 4, 11);
        var rk = Arr(c, a[0]);
        var ck = Arr(c, a[1]);
        var values = Arr(c, a[2]);
        var f = Fn(c, a[3]);
        int n = values.GetLength(0);
        if (rk.GetLength(0) != n || ck.GetLength(0) != n) throw Err(Errors.Value);
        int headers = HeaderMode(c, a, 4, values);
        int rowDepth = (int)c.NumOr(a, 5, 1);
        int rowSort = (int)c.NumOr(a, 6, 1);
        int colDepth = (int)c.NumOr(a, 7, 1);
        int colSort = (int)c.NumOr(a, 8, 1);
        var filter = a.Length > 9 && a[9] is not MissingExpr ? Arr(c, a[9]) : null;
        int start = headers is 1 or 3 ? 1 : 0;
        var rows = new GroupTree();
        var cols = new GroupTree();
        var used = new List<int>();
        for (int r = start; r < n; r++)
        {
            if (filter != null)
            {
                var fv = filter[Math.Min(r, filter.GetLength(0) - 1), 0];
                if (fv.IsError) throw Err(fv.Str!);
                if (!Calc.ToBool(fv)) continue;
            }
            rows.Add(RowOf(rk, r), r);
            cols.Add(RowOf(ck, r), r);
            used.Add(r);
        }
        int nrk = rk.GetLength(1), nck = ck.GetLength(1), nv = values.GetLength(1);
        List<(CellValue[] Key, List<int> Rows)> Sorted(GroupTree t, int sort, int nkeys)
        {
            var list = t.Order.Select(k => t.Groups[k]).ToList();
            int col = Math.Abs(sort) - 1;
            if (col < 0 || col >= nkeys) col = 0;
            var res = list.OrderBy(x => x.Key[col], Comparer<CellValue>.Create(SortCompare)).ToList();
            if (sort < 0) res.Reverse();
            return res;
        }
        var rlist = Sorted(rows, rowSort, nrk);
        var clist = Sorted(cols, colSort, nck);
        int width = nrk + (clist.Count + (colDepth != 0 ? 1 : 0)) * nv;
        var lines = new List<CellValue[]>();
        // header rows: column keys (one line per column field) then value names when several value columns
        for (int lvl = 0; lvl < nck; lvl++)
        {
            var h = Enumerable.Repeat(CellValue.Blank, width).ToArray();
            int at = nrk;
            if (colDepth < 0) { if (lvl == 0) h[at] = S("Total"); at += nv; }
            foreach (var ckey in clist) { h[at] = ckey.Key[lvl]; at += nv; }
            if (colDepth > 0 && lvl == 0) h[at] = S("Total");
            lines.Add(h);
        }
        if (headers is 2 or 3 || nv > 1)
        {
            var h = Enumerable.Repeat(CellValue.Blank, width).ToArray();
            for (int k = 0; k < nrk; k++) h[k] = headers == 3 ? rk[0, k] : headers == 2 ? S("Row Field") : CellValue.Blank;
            for (int at = nrk; at < width; at++) { int vk = (at - nrk) % nv; h[at] = headers == 3 ? values[0, vk] : S(FnLabel(f)); }
            lines.Add(h);
        }
        CellValue[] Line(CellValue[] key, List<int> rset)
        {
            var line = new CellValue[width];
            for (int k = 0; k < nrk; k++) line[k] = k < key.Length ? key[k] : CellValue.Blank;
            int at = nrk;
            var set = new HashSet<int>(rset);
            if (colDepth < 0) { for (int v = 0; v < nv; v++) line[at++] = Aggregate(c, f, values, v, rset); }
            foreach (var ckey in clist)
            {
                var inter = ckey.Rows.Where(set.Contains).ToList();
                for (int v = 0; v < nv; v++) line[at++] = inter.Count == 0 ? CellValue.Blank : Aggregate(c, f, values, v, inter);
            }
            if (colDepth > 0) for (int v = 0; v < nv; v++) line[at++] = Aggregate(c, f, values, v, rset);
            return line;
        }
        var totalKey = new CellValue[nrk];
        totalKey[0] = S("Total");
        for (int k = 1; k < nrk; k++) totalKey[k] = CellValue.Blank;
        if (rowDepth < 0) lines.Add(Line(totalKey, used));
        foreach (var r in rlist) lines.Add(Line(r.Key, r.Rows));
        if (rowDepth > 0) lines.Add(Line(totalKey, used));
        return Out(FromRows(lines, width));
    }
}
