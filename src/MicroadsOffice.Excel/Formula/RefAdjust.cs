using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

/// <summary>Rewrites references inside formulas (copy/fill, insert/delete rows/columns, sheet rename).</summary>
public static class RefAdjust
{
    public static Expr Map(Expr e, Func<RefExpr, Expr> f) => e switch
    {
        RefExpr r => f(r),
        FuncExpr fn => fn with { Args = fn.Args.Select(a => Map(a, f)).ToArray() },
        UnaryExpr u => u with { Operand = Map(u.Operand, f) },
        BinaryExpr b => b with { Left = Map(b.Left, f), Right = Map(b.Right, f) },
        PercentExpr p => p with { Operand = Map(p.Operand, f) },
        ParenExpr p => p with { Inner = Map(p.Inner, f) },
        _ => e
    };

    public static IEnumerable<RefExpr> Refs(Expr e)
    {
        switch (e)
        {
            case RefExpr r: yield return r; break;
            case FuncExpr fn: foreach (var a in fn.Args) foreach (var x in Refs(a)) yield return x; break;
            case UnaryExpr u: foreach (var x in Refs(u.Operand)) yield return x; break;
            case BinaryExpr b: foreach (var x in Refs(b.Left)) yield return x; foreach (var x in Refs(b.Right)) yield return x; break;
            case PercentExpr p: foreach (var x in Refs(p.Operand)) yield return x; break;
            case ParenExpr p: foreach (var x in Refs(p.Inner)) yield return x; break;
        }
    }

    /// <summary>Copying a formula by (dr, dc): relative parts move, absolute parts stay. Out of sheet → #REF!.</summary>
    public static string Shift(string formula, int dr, int dc)
    {
        if (dr == 0 && dc == 0) return formula;
        Expr e;
        try { e = FormulaParser.Parse(formula); } catch { return formula; }
        var res = Map(e, r =>
        {
            int r1 = r.AR1 || r.Kind == RefKind.Cols ? r.R1 : r.R1 + dr;
            int r2 = r.AR2 || r.Kind == RefKind.Cols ? r.R2 : r.R2 + dr;
            int c1 = r.AC1 || r.Kind == RefKind.Rows ? r.C1 : r.C1 + dc;
            int c2 = r.AC2 || r.Kind == RefKind.Rows ? r.C2 : r.C2 + dc;
            if (r1 < 0 || r2 < 0 || c1 < 0 || c2 < 0 || r1 >= Addr.MaxRows || r2 >= Addr.MaxRows || c1 >= Addr.MaxCols || c2 >= Addr.MaxCols)
                return new ErrorExpr(Errors.Ref);
            return r with { R1 = r1, R2 = r2, C1 = c1, C2 = c2 };
        });
        return FormulaPrinter.Print(res);
    }

    /// <summary>
    /// Adjusts a formula living on <paramref name="formulaSheet"/> after rows (isRow) or columns were inserted
    /// (count &gt; 0) or deleted (count &lt; 0) at index <paramref name="at"/> of <paramref name="changedSheet"/>.
    /// </summary>
    public static string InsertDelete(string formula, string formulaSheet, string changedSheet, bool isRow, int at, int count)
    {
        Expr e;
        try { e = FormulaParser.Parse(formula); } catch { return formula; }
        bool changed = false;
        var res = Map(e, r =>
        {
            var target = r.Sheet ?? formulaSheet;
            if (!string.Equals(target, changedSheet, StringComparison.OrdinalIgnoreCase)) return r;
            if (isRow && r.Kind == RefKind.Cols) return r;
            if (!isRow && r.Kind == RefKind.Rows) return r;
            int a1 = isRow ? r.R1 : r.C1, a2 = isRow ? r.R2 : r.C2;
            bool flip = a1 > a2;
            if (flip) (a1, a2) = (a2, a1);
            int n1 = a1, n2 = a2;
            if (count > 0)
            {
                if (a1 >= at) n1 += count;
                if (a2 >= at) n2 += count;
            }
            else
            {
                int del = -count, dEnd = at + del - 1;
                if (a1 >= at && a2 <= dEnd) { changed = true; return new ErrorExpr(Errors.Ref); }
                n1 = a1 > dEnd ? a1 - del : a1 >= at ? at : a1;
                n2 = a2 > dEnd ? a2 - del : a2 >= at ? at - 1 : a2;
                if (n2 < n1) { changed = true; return new ErrorExpr(Errors.Ref); }
            }
            int max = isRow ? Addr.MaxRows - 1 : Addr.MaxCols - 1;
            if (n1 > max || n2 > max) { changed = true; return new ErrorExpr(Errors.Ref); }
            if (n1 == a1 && n2 == a2) return r;
            changed = true;
            if (flip) (n1, n2) = (n2, n1);
            return isRow ? r with { R1 = n1, R2 = n2 } : r with { C1 = n1, C2 = n2 };
        });
        return changed ? FormulaPrinter.Print(res) : formula;
    }

    /// <summary>Formula text after a sheet was renamed (or deleted → #REF!).</summary>
    public static string RenameSheet(string formula, string oldName, string? newName)
    {
        if (formula.IndexOf(oldName, StringComparison.OrdinalIgnoreCase) < 0) return formula;
        Expr e;
        try { e = FormulaParser.Parse(formula); } catch { return formula; }
        bool changed = false;
        var res = Map(e, r =>
        {
            if (r.Sheet == null || !string.Equals(r.Sheet, oldName, StringComparison.OrdinalIgnoreCase)) return r;
            changed = true;
            return newName == null ? new ErrorExpr(Errors.Ref) : r with { Sheet = newName };
        });
        return changed ? FormulaPrinter.Print(res) : formula;
    }

    /// <summary>Moves references pointing into <paramref name="src"/> by (dr, dc) — used for cut &amp; paste.</summary>
    public static string MoveRefs(string formula, string formulaSheet, string movedSheet, CellRange src, int dr, int dc)
    {
        Expr e;
        try { e = FormulaParser.Parse(formula); } catch { return formula; }
        bool changed = false;
        var res = Map(e, r =>
        {
            var target = r.Sheet ?? formulaSheet;
            if (!string.Equals(target, movedSheet, StringComparison.OrdinalIgnoreCase) || r.Kind is RefKind.Cols or RefKind.Rows) return r;
            if (!src.Contains(r.Range)) return r;
            changed = true;
            return r with { R1 = r.R1 + dr, R2 = r.R2 + dr, C1 = r.C1 + dc, C2 = r.C2 + dc };
        });
        return changed ? FormulaPrinter.Print(res) : formula;
    }
}
