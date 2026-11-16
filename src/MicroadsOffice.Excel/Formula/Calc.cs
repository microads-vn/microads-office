using System.Globalization;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

public sealed class FormulaError(string code) : Exception(code)
{
    public string Code { get; } = code;
}

/// <summary>A reference value (range of cells on a sheet).</summary>
public sealed class RangeVal(Sheet sheet, CellRange range)
{
    public Sheet Sheet { get; } = sheet;
    public CellRange Range { get; } = range;
    public int Rows => Range.Rows;
    public int Cols => Range.Cols;
}

/// <summary>An in-memory array (array constants, element-wise results).</summary>
public sealed class ArrayVal(CellValue[,] v)
{
    public CellValue[,] V { get; } = v;
    public int Rows => V.GetLength(0);
    public int Cols => V.GetLength(1);
}

/// <summary>Evaluation context of one formula.</summary>
public sealed class Ctx
{
    public Calc Calc { get; }
    public Sheet Sheet { get; }
    public int Row { get; }
    public int Col { get; }
    internal Ctx(Calc calc, Sheet sheet, int row, int col) { Calc = calc; Sheet = sheet; Row = row; Col = col; }
    public Workbook Book => Sheet.Book;

    public object Eval(Expr e) => Calc.Eval(e, this);

    public CellValue Scalar(Expr e) => Calc.ToScalar(Eval(e), this);

    public double Num(Expr e)
    {
        var v = Scalar(e);
        return Calc.ToNum(v);
    }

    public double NumOr(Expr[] a, int i, double def) => i < a.Length && a[i] is not MissingExpr ? Num(a[i]) : def;
    public string Str(Expr e) => Calc.ToText(Scalar(e));
    public string StrOr(Expr[] a, int i, string def) => i < a.Length && a[i] is not MissingExpr ? Str(a[i]) : def;
    public bool Bool(Expr e) => Calc.ToBool(Scalar(e));
    public bool BoolOr(Expr[] a, int i, bool def) => i < a.Length && a[i] is not MissingExpr ? Bool(a[i]) : def;
    public int Int(Expr e) => (int)Math.Truncate(Num(e));
}

/// <summary>
/// Workbook calculation: on-demand recursive evaluation with a per-pass stamp (each formula is evaluated once per
/// recalculation), circular reference detection and a large-stack worker thread for long dependency chains.
/// </summary>
public sealed class Calc
{
    private readonly Workbook _wb;
    private int _stamp;
    private readonly HashSet<Cell> _evaluating = new();
    private readonly Dictionary<Sheet, CellRange> _usedCache = new();
    private readonly Dictionary<string, Expr?> _nameCache = new(StringComparer.OrdinalIgnoreCase);
    public List<(Sheet Sheet, int Row, int Col)> Circular { get; } = new();
    public DateTime? FixedNow { get; set; }

    public Calc(Workbook wb) { _wb = wb; }

    // ------------------------------------------------------------------ public entry points

    public void RecalculateAll(bool force = false)
    {
        RunWithBigStack(() =>
        {
            _stamp++;
            _usedCache.Clear();
            _nameCache.Clear();
            Circular.Clear();
            _evaluating.Clear();
            foreach (var sheet in _wb.Sheets.ToList())
            {
                var formulas = sheet.Cells.Where(kv => kv.Value.Formula != null).Select(kv => kv.Key).ToList();
                formulas.Sort();
                foreach (var k in formulas)
                {
                    var (r, c) = Sheet.FromKey(k);
                    if (sheet.Cells.TryGetValue(k, out var cell)) Ensure(sheet, r, c, cell);
                }
            }
        });
    }

    /// <summary>Evaluates a formula text as if it were in cell (row, col) of <paramref name="sheet"/>.</summary>
    public CellValue EvaluateFormula(Sheet sheet, int row, int col, string formula)
    {
        CellValue result = default;
        RunWithBigStack(() =>
        {
            if (_stamp == 0) _stamp++;
            Expr e;
            try { e = FormulaParser.Parse(formula); }
            catch { result = CellValue.Error(Errors.Name); return; }
            var ctx = new Ctx(this, sheet, row, col);
            result = Final(SafeEval(e, ctx), ctx);
        });
        return result;
    }

    /// <summary>Recomputes a single cell (manual calculation mode).</summary>
    public void RecalculateCell(Sheet sheet, int row, int col)
    {
        var cell = sheet.GetCell(row, col);
        if (cell?.Formula == null) return;
        RunWithBigStack(() => { _stamp++; _usedCache.Clear(); _nameCache.Clear(); Ensure(sheet, row, col, cell); });
    }

    private static void RunWithBigStack(Action a)
    {
        Exception? err = null;
        var t = new Thread(() => { try { a(); } catch (Exception ex) { err = ex; } }, 256 * 1024 * 1024) { IsBackground = true, Name = "Calc" };
        t.Start();
        t.Join();
        if (err != null) throw new InvalidOperationException("Calculation failed", err);
    }

    // ------------------------------------------------------------------ cell evaluation

    internal CellValue ValueOf(Sheet s, int r, int c)
    {
        if (!s.Cells.TryGetValue(Sheet.Key(r, c), out var cell)) return CellValue.Blank;
        if (cell.Formula != null) Ensure(s, r, c, cell);
        return cell.Value;
    }

    private void Ensure(Sheet sheet, int r, int c, Cell cell)
    {
        if (cell.CalcStamp == _stamp) return;
        if (_evaluating.Contains(cell))
        {
            Circular.Add((sheet, r, c));
            return;
        }
        _evaluating.Add(cell);
        try
        {
            if (cell.Ast == null)
            {
                try { cell.Ast = FormulaParser.Parse(cell.Formula!); }
                catch { cell.Ast = new ErrorExpr(Errors.Name); }
            }
            var ctx = new Ctx(this, sheet, r, c);
            var v = Final(SafeEval(cell.Ast, ctx), ctx);
            if (Circular.Any(x => x.Sheet == sheet && x.Row == r && x.Col == c)) v = CellValue.Number(0);
            cell.Value = v;
            cell.CalcStamp = _stamp;
        }
        finally { _evaluating.Remove(cell); }
    }

    private object SafeEval(Expr e, Ctx ctx)
    {
        try { return Eval(e, ctx); }
        catch (FormulaError fe) { return CellValue.Error(fe.Code); }
        catch (OverflowException) { return CellValue.Error(Errors.Num); }
        catch (Exception) { return CellValue.Error(Errors.Value); }
    }

    /// <summary>Reduces a formula result to one cell value (implicit intersection / top-left of arrays).</summary>
    internal CellValue Final(object v, Ctx ctx)
    {
        var s = ToScalar(v, ctx);
        return s.IsBlank ? CellValue.Number(0) : s;
    }

    internal CellRange UsedRange(Sheet s)
    {
        if (!_usedCache.TryGetValue(s, out var u)) _usedCache[s] = u = s.UsedRange();
        return u;
    }

    // ------------------------------------------------------------------ expression evaluation

    internal object Eval(Expr e, Ctx ctx)
    {
        switch (e)
        {
            case NumberExpr n: return CellValue.Number(n.Value);
            case StringExpr s: return CellValue.Text(s.Value);
            case BoolExpr b: return CellValue.Bool(b.Value);
            case ErrorExpr er: return CellValue.Error(er.Code);
            case MissingExpr: return CellValue.Blank;
            case ParenExpr p: return Eval(p.Inner, ctx);
            case ArrayExpr a: return new ArrayVal(a.Values);
            case RefExpr r:
            {
                var sheet = r.Sheet == null ? ctx.Sheet : _wb.Sheet(r.Sheet) ?? throw new FormulaError(Errors.Ref);
                return new RangeVal(sheet, r.Range);
            }
            case NameExpr nm: return EvalName(nm.Name, ctx);
            case UnaryExpr u:
            {
                var v = Eval(u.Operand, ctx);
                return u.Op == '-' ? Lift1(v, ctx, x => CellValue.Number(-ToNum(x))) : Lift1(v, ctx, x => CellValue.Number(ToNum(x)));
            }
            case PercentExpr pc: return Lift1(Eval(pc.Operand, ctx), ctx, x => CellValue.Number(ToNum(x) / 100));
            case BinaryExpr b: return EvalBinary(b, ctx);
            case FuncExpr f:
            {
                if (!Functions.Table.TryGetValue(f.Name, out var impl)) throw new FormulaError(Errors.Name);
                try { return impl(ctx, f.Args); }
                catch (FormulaError fe) { return CellValue.Error(fe.Code); }
                catch (OverflowException) { return CellValue.Error(Errors.Num); }
                catch (ArgumentOutOfRangeException) { return CellValue.Error(Errors.Value); }
                catch (IndexOutOfRangeException) { return CellValue.Error(Errors.Value); }
            }
        }
        throw new FormulaError(Errors.Value);
    }

    private object EvalName(string name, Ctx ctx)
    {
        if (!_nameCache.TryGetValue(name, out var expr))
        {
            expr = null;
            if (_wb.Names.TryGetValue(name, out var text))
            {
                try { expr = FormulaParser.Parse(text); } catch { expr = new ErrorExpr(Errors.Name); }
            }
            _nameCache[name] = expr;
        }
        if (expr == null) throw new FormulaError(Errors.Name);
        return Eval(expr, ctx);
    }

    public void InvalidateNames() => _nameCache.Clear();

    private object EvalBinary(BinaryExpr b, Ctx ctx)
    {
        var l = Eval(b.Left, ctx);
        var r = Eval(b.Right, ctx);
        Func<CellValue, CellValue, CellValue> op = b.Op switch
        {
            "+" => (x, y) => CellValue.Number(ToNum(x) + ToNum(y)),
            "-" => (x, y) => CellValue.Number(ToNum(x) - ToNum(y)),
            "*" => (x, y) => CellValue.Number(ToNum(x) * ToNum(y)),
            "/" => (x, y) => { var d = ToNum(y); var n = ToNum(x); if (d == 0) throw new FormulaError(Errors.Div0); return CellValue.Number(n / d); },
            "^" => (x, y) =>
            {
                var bs = ToNum(x); var ex = ToNum(y);
                if (bs == 0 && ex == 0) throw new FormulaError(Errors.Num);
                if (bs == 0 && ex < 0) throw new FormulaError(Errors.Div0);
                var res = Math.Pow(bs, ex);
                if (double.IsNaN(res) || double.IsInfinity(res)) throw new FormulaError(Errors.Num);
                return CellValue.Number(res);
            },
            "&" => (x, y) => CellValue.Text(ToText(x) + ToText(y)),
            "=" => (x, y) => CellValue.Bool(CompareOp(x, y) == 0),
            "<>" => (x, y) => CellValue.Bool(CompareOp(x, y) != 0),
            "<" => (x, y) => CellValue.Bool(CompareOp(x, y) < 0),
            ">" => (x, y) => CellValue.Bool(CompareOp(x, y) > 0),
            "<=" => (x, y) => CellValue.Bool(CompareOp(x, y) <= 0),
            ">=" => (x, y) => CellValue.Bool(CompareOp(x, y) >= 0),
            _ => throw new FormulaError(Errors.Value)
        };
        return Lift2(l, r, ctx, op);
    }

    private static int CompareOp(CellValue x, CellValue y)
    {
        if (x.IsError) throw new FormulaError(x.Str!);
        if (y.IsError) throw new FormulaError(y.Str!);
        return Compare(x, y);
    }

    /// <summary>Excel ordering: numbers &lt; text &lt; logicals; text is case-insensitive; blank equals 0 / "" / FALSE.</summary>
    public static int Compare(CellValue x, CellValue y)
    {
        if (x.IsBlank) x = y.Kind switch { ValueKind.Text => CellValue.Text(""), ValueKind.Bool => CellValue.Bool(false), _ => CellValue.Number(0) };
        if (y.IsBlank) y = x.Kind switch { ValueKind.Text => CellValue.Text(""), ValueKind.Bool => CellValue.Bool(false), _ => CellValue.Number(0) };
        int Rank(CellValue v) => v.Kind switch { ValueKind.Number => 0, ValueKind.Text => 1, ValueKind.Bool => 2, _ => 3 };
        int rx = Rank(x), ry = Rank(y);
        if (rx != ry) return rx.CompareTo(ry);
        return x.Kind switch
        {
            ValueKind.Number or ValueKind.Bool => Math.Abs(x.Num - y.Num) < 1e-12 * Math.Max(1, Math.Abs(x.Num)) ? 0 : x.Num.CompareTo(y.Num),
            ValueKind.Text => string.Compare(x.Str, y.Str, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase),
            _ => 0
        };
    }

    // ------------------------------------------------------------------ element-wise helpers

    internal object Lift1(object v, Ctx ctx, Func<CellValue, CellValue> op)
    {
        if (v is CellValue s) return Wrap(op, s);
        var arr = ToArray(v);
        var res = new CellValue[arr.GetLength(0), arr.GetLength(1)];
        for (int i = 0; i < arr.GetLength(0); i++)
            for (int j = 0; j < arr.GetLength(1); j++) res[i, j] = Wrap(op, arr[i, j]);
        return res.Length == 1 && v is RangeVal ? res[0, 0] : new ArrayVal(res);
    }

    private static CellValue Wrap(Func<CellValue, CellValue> op, CellValue v)
    {
        try { return op(v); } catch (FormulaError e) { return CellValue.Error(e.Code); }
    }

    internal object Lift2(object a, object b, Ctx ctx, Func<CellValue, CellValue, CellValue> op)
    {
        if (a is RangeVal ra && ra.Range.IsSingle) a = ValueOf(ra.Sheet, ra.Range.R1, ra.Range.C1);
        if (b is RangeVal rb && rb.Range.IsSingle) b = ValueOf(rb.Sheet, rb.Range.R1, rb.Range.C1);
        if (a is CellValue x && b is CellValue y)
        {
            try { return op(x, y); } catch (FormulaError e) { return CellValue.Error(e.Code); }
        }
        var A = a is CellValue av ? new[,] { { av } } : ToArray(a);
        var B = b is CellValue bv ? new[,] { { bv } } : ToArray(b);
        int rows = Math.Max(A.GetLength(0), B.GetLength(0)), cols = Math.Max(A.GetLength(1), B.GetLength(1));
        var res = new CellValue[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
            {
                var p = Pick(A, i, j);
                var q = Pick(B, i, j);
                if (p == null || q == null) { res[i, j] = CellValue.Error(Errors.NA); continue; }
                try { res[i, j] = op(p.Value, q.Value); } catch (FormulaError e) { res[i, j] = CellValue.Error(e.Code); }
            }
        return new ArrayVal(res);
    }

    private static CellValue? Pick(CellValue[,] m, int i, int j)
    {
        int r = m.GetLength(0) == 1 ? 0 : i, c = m.GetLength(1) == 1 ? 0 : j;
        if (r >= m.GetLength(0) || c >= m.GetLength(1)) return null;
        return m[r, c];
    }

    /// <summary>Materialises a range/array as a 2-D array (large ranges are clipped to the used area).</summary>
    public CellValue[,] ToArray(object v)
    {
        switch (v)
        {
            case CellValue s: return new[,] { { s } };
            case ArrayVal a: return a.V;
            case RangeVal r:
            {
                var rg = Clip(r);
                if (rg.Rows <= 0 || rg.Cols <= 0) return new CellValue[1, 1];
                var res = new CellValue[rg.Rows, rg.Cols];
                for (int i = 0; i < rg.Rows; i++)
                    for (int j = 0; j < rg.Cols; j++) res[i, j] = ValueOf(r.Sheet, rg.R1 + i, rg.C1 + j);
                return res;
            }
        }
        return new CellValue[1, 1];
    }

    /// <summary>Range clipped to the used area of its sheet (for whole-column references).</summary>
    public CellRange Clip(RangeVal r)
    {
        var rg = r.Range;
        if (rg.Count <= 20000) return rg;
        var u = UsedRange(r.Sheet);
        return new CellRange(rg.R1, rg.C1, Math.Min(rg.R2, Math.Max(u.R2, rg.R1)), Math.Min(rg.C2, Math.Max(u.C2, rg.C1)));
    }

    /// <summary>Every value in a range (cells only; for huge ranges only existing cells are visited).</summary>
    public IEnumerable<CellValue> RangeValues(RangeVal r)
    {
        var rg = r.Range;
        if (rg.Count > 20000)
        {
            var keys = r.Sheet.Cells.Keys.Where(k => { var (row, col) = Sheet.FromKey(k); return rg.Contains(row, col); }).ToList();
            foreach (var k in keys) { var (row, col) = Sheet.FromKey(k); yield return ValueOf(r.Sheet, row, col); }
            yield break;
        }
        for (int i = rg.R1; i <= rg.R2; i++)
            for (int j = rg.C1; j <= rg.C2; j++) yield return ValueOf(r.Sheet, i, j);
    }

    // ------------------------------------------------------------------ coercions

    public CellValue ToScalar(object v, Ctx ctx)
    {
        switch (v)
        {
            case CellValue s: return s;
            case ArrayVal a: return a.V.Length > 0 ? a.V[0, 0] : CellValue.Error(Errors.Value);
            case RangeVal r:
            {
                var rg = r.Range;
                if (rg.IsSingle) return ValueOf(r.Sheet, rg.R1, rg.C1);
                if (rg.C1 == rg.C2 && ctx.Row >= rg.R1 && ctx.Row <= rg.R2) return ValueOf(r.Sheet, ctx.Row, rg.C1);
                if (rg.R1 == rg.R2 && ctx.Col >= rg.C1 && ctx.Col <= rg.C2) return ValueOf(r.Sheet, rg.R1, ctx.Col);
                if (r.Sheet == ctx.Sheet && rg.Contains(ctx.Row, ctx.Col)) return CellValue.Error(Errors.Value);
                return CellValue.Error(Errors.Value);
            }
        }
        return CellValue.Error(Errors.Value);
    }

    public static double ToNum(CellValue v)
    {
        switch (v.Kind)
        {
            case ValueKind.Number: case ValueKind.Bool: return v.Num;
            case ValueKind.Blank: return 0;
            case ValueKind.Error: throw new FormulaError(v.Str!);
            default:
                if (ValueParser.TryCoerceNumber(v.Str!, out var d)) return d;
                throw new FormulaError(Errors.Value);
        }
    }

    public static string ToText(CellValue v) => v.Kind switch
    {
        ValueKind.Number => CellValue.FormatGeneralInvariant(v.Num),
        ValueKind.Bool => v.Num != 0 ? "TRUE" : "FALSE",
        ValueKind.Text => v.Str!,
        ValueKind.Error => throw new FormulaError(v.Str!),
        _ => ""
    };

    public static string ToTextSafe(CellValue v) => v.IsError ? v.Str! : ToText(v);

    public static bool ToBool(CellValue v)
    {
        switch (v.Kind)
        {
            case ValueKind.Number: case ValueKind.Bool: return v.Num != 0;
            case ValueKind.Blank: return false;
            case ValueKind.Error: throw new FormulaError(v.Str!);
            default:
                if (v.Str!.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (v.Str!.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;
                throw new FormulaError(Errors.Value);
        }
    }

    public DateTime Now => FixedNow ?? DateTime.Now;
}
