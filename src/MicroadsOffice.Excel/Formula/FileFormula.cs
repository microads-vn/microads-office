namespace MicroadsOffice.Excel.Formula;

/// <summary>
/// Formula text as stored in .xlsx files: functions newer than Excel 2007 carry the "_xlfn." prefix (FILTER/SORT
/// "_xlfn._xlws."), LET/LAMBDA variables "_xlpm.", the spill operator A1# is "_xlfn.ANCHORARRAY(A1)" and "@" is
/// "_xlfn.SINGLE(...)". Without the prefixes Excel shows #NAME? for these functions.
/// </summary>
public static class FileFormula
{
    private static readonly HashSet<string> Future = new(StringComparer.OrdinalIgnoreCase)
    {
        "ACOT", "ACOTH", "AGGREGATE", "ARABIC", "ARRAYTOTEXT", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
        "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL", "BYROW", "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT",
        "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHOOSECOLS", "CHOOSEROWS", "COMBINA", "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH",
        "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "DROP", "ECMA.CEILING", "ENCODEURL", "ERF.PRECISE", "ERFC.PRECISE", "EXPAND",
        "EXPON.DIST", "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FILTERXML", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS",
        "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV",
        "GAMMALN.PRECISE", "GAUSS", "GROUPBY", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH",
        "IMSINH", "IMTAN", "ISFORMULA", "ISOMITTED", "ISOWEEKNUM", "LAMBDA", "LET", "LOGNORM.DIST", "LOGNORM.INV", "MAKEARRAY", "MAP", "MAXIFS",
        "MINIFS", "MODE.MULT", "MODE.SNGL", "MUNIT", "NEGBINOM.DIST", "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
        "NUMBERVALUE", "PDURATION", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTOF", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA", "PHI",
        "PIVOTBY", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "RANDARRAY", "RANK.AVG", "RANK.EQ", "REDUCE", "REGEXEXTRACT", "REGEXREPLACE",
        "REGEXTEST", "RRI", "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS", "SKEW.P", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST",
        "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "TOCOL", "TOROW",
        "TRIMRANGE", "UNICHAR", "UNICODE", "UNIQUE", "VALUETOTEXT", "VAR.P", "VAR.S", "VSTACK", "WEBSERVICE", "WEIBULL.DIST", "WORKDAY.INTL",
        "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "XOR", "Z.TEST"
    };

    private static readonly HashSet<string> Worksheet = new(StringComparer.OrdinalIgnoreCase) { "FILTER", "SORT" };

    /// <summary>Formula text to write into a file (unchanged when it needs no prefixes).</summary>
    public static string ToFile(string formula)
    {
        Expr e;
        try { e = FormulaParser.Parse(formula); } catch { return formula; }
        var changed = new Flag();
        var res = ToFileExpr(e, new HashSet<string>(StringComparer.OrdinalIgnoreCase), changed);
        return changed.Changed ? FormulaPrinter.Print(res) : formula;
    }

    /// <summary>Formula text read from a file → the form shown to the user (prefixes removed).</summary>
    public static string FromFile(string text)
    {
        if (text.IndexOf("_xl", StringComparison.OrdinalIgnoreCase) < 0) return text;
        try { return FormulaPrinter.Print(Normalize(FormulaParser.Parse(text))); } catch { return text; }
    }

    private sealed class Flag { public bool Changed; }

    /// <summary>Functions whose result can be an array: their cells are saved as dynamic-array formulas (cm="1").</summary>
    private static readonly HashSet<string> ArrayResult = new(StringComparer.OrdinalIgnoreCase)
    {
        "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "XLOOKUP", "VSTACK", "HSTACK", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS",
        "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS", "EXPAND", "TEXTSPLIT", "MAP", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "REDUCE", "GROUPBY", "PIVOTBY",
        "LET", "LAMBDA", "REGEXEXTRACT", "TRANSPOSE", "MMULT", "MINVERSE", "MUNIT", "FREQUENCY", "LINEST", "LOGEST", "TREND", "GROWTH", "MODE.MULT",
        "TRIMRANGE"
    };

    /// <summary>The formula may return several values (dynamic-array functions, spill references, user functions).</summary>
    public static bool MayReturnArray(string? formula)
    {
        if (string.IsNullOrEmpty(formula)) return false;
        Expr e;
        try { e = FormulaParser.Parse(formula); } catch { return false; }
        return Any(e);
        static bool Any(Expr x) => x switch
        {
            RefExpr r => r.Spill,
            CallExpr => true,
            FuncExpr f => ArrayResult.Contains(f.Name) && f.Args.Any(a => a is not (NumberExpr or StringExpr or BoolExpr)) || !Functions.Table.ContainsKey(f.Name) || f.Args.Any(Any),
            UnaryExpr u => Any(u.Operand),
            BinaryExpr b => Any(b.Left) || Any(b.Right),
            PercentExpr p => Any(p.Operand),
            ParenExpr p => Any(p.Inner),
            _ => false
        };
    }

    private static Expr Normalize(Expr e) => e switch
    {
        FuncExpr { Name: "ANCHORARRAY", Args: [RefExpr { Kind: RefKind.Cell } r] } => r with { Spill = true },
        FuncExpr { Name: "SINGLE", Args: [var x] } => new UnaryExpr('@', Normalize(x) is var nx && nx is RefExpr or FuncExpr or NameExpr or ParenExpr or CallExpr or NumberExpr or StringExpr ? nx : new ParenExpr(nx)),
        FuncExpr f => f with { Args = f.Args.Select(Normalize).ToArray() },
        CallExpr ce => ce with { Callee = Normalize(ce.Callee), Args = ce.Args.Select(Normalize).ToArray() },
        UnaryExpr u => u with { Operand = Normalize(u.Operand) },
        BinaryExpr b => b with { Left = Normalize(b.Left), Right = Normalize(b.Right) },
        PercentExpr p => p with { Operand = Normalize(p.Operand) },
        ParenExpr p => p with { Inner = Normalize(p.Inner) },
        _ => e
    };

    private static Expr ToFileExpr(Expr e, HashSet<string> bound, Flag changed)
    {
        switch (e)
        {
            case RefExpr { Spill: true } r:
                changed.Changed = true;
                return new FuncExpr("_xlfn.ANCHORARRAY", [r with { Spill = false }]);
            case NameExpr n:
            {
                var name = n.Name;
                bool opt = name.StartsWith('[') && name.EndsWith(']');
                var bare = opt ? name[1..^1] : name;
                if (!bound.Contains(bare)) return e;
                changed.Changed = true;
                return new NameExpr(opt ? "[_xlpm." + bare + "]" : "_xlpm." + bare);
            }
            case UnaryExpr { Op: '@' } u:
                changed.Changed = true;
                return new FuncExpr("_xlfn.SINGLE", [ToFileExpr(u.Operand, bound, changed)]);
            case FuncExpr f when f.Name.Equals("LET", StringComparison.OrdinalIgnoreCase):
            {
                changed.Changed = true;
                var inner = new HashSet<string>(bound, StringComparer.OrdinalIgnoreCase);
                var args = new Expr[f.Args.Length];
                for (int i = 0; i < f.Args.Length; i++)
                {
                    bool isName = i % 2 == 0 && i < f.Args.Length - 1 && f.Args[i] is NameExpr;
                    if (isName) inner.Add(((NameExpr)f.Args[i]).Name);
                    args[i] = isName ? new NameExpr("_xlpm." + ((NameExpr)f.Args[i]).Name) : ToFileExpr(f.Args[i], inner, changed);
                }
                return new FuncExpr("_xlfn.LET", args);
            }
            case FuncExpr f when f.Name.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase):
            {
                changed.Changed = true;
                var inner = new HashSet<string>(bound, StringComparer.OrdinalIgnoreCase);
                var args = new Expr[f.Args.Length];
                for (int i = 0; i < f.Args.Length - 1; i++)
                {
                    if (f.Args[i] is not NameExpr p) { args[i] = f.Args[i]; continue; }
                    bool opt = p.Name.StartsWith('[') && p.Name.EndsWith(']');
                    var bare = opt ? p.Name[1..^1] : p.Name;
                    inner.Add(bare);
                    args[i] = new NameExpr(opt ? "[_xlpm." + bare + "]" : "_xlpm." + bare);
                }
                if (f.Args.Length > 0) args[^1] = ToFileExpr(f.Args[^1], inner, changed);
                return new FuncExpr("_xlfn.LAMBDA", args);
            }
            case FuncExpr f:
            {
                var args = f.Args.Select(a => ToFileExpr(a, bound, changed)).ToArray();
                string name = f.Name;
                if (bound.Contains(name)) { name = "_xlpm." + name; changed.Changed = true; }
                else if (Worksheet.Contains(name)) { name = "_xlfn._xlws." + name; changed.Changed = true; }
                else if (Future.Contains(name)) { name = "_xlfn." + name; changed.Changed = true; }
                return new FuncExpr(name, args);
            }
            case CallExpr ce:
            {
                var callee = ToFileExpr(ce.Callee, bound, changed);
                var args = ce.Args.Select(a => ToFileExpr(a, bound, changed)).ToArray();
                return new CallExpr(callee, args);
            }
            case UnaryExpr u: return u with { Operand = ToFileExpr(u.Operand, bound, changed) };
            case BinaryExpr b:
            {
                var l = ToFileExpr(b.Left, bound, changed);
                var r = ToFileExpr(b.Right, bound, changed);
                return b with { Left = l, Right = r };
            }
            case PercentExpr p: return p with { Operand = ToFileExpr(p.Operand, bound, changed) };
            case ParenExpr p: return p with { Inner = ToFileExpr(p.Inner, bound, changed) };
            default: return e;
        }
    }
}
