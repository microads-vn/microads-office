using System.Globalization;
using System.Text;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

public enum RefKind : byte { Cell, Area, Cols, Rows }

public abstract record Expr;
public sealed record NumberExpr(double Value) : Expr;
public sealed record StringExpr(string Value) : Expr;
public sealed record BoolExpr(bool Value) : Expr;
public sealed record ErrorExpr(string Code) : Expr;
public sealed record NameExpr(string Name) : Expr;
public sealed record FuncExpr(string Name, Expr[] Args) : Expr;
public sealed record UnaryExpr(char Op, Expr Operand) : Expr;
public sealed record BinaryExpr(string Op, Expr Left, Expr Right) : Expr;
public sealed record PercentExpr(Expr Operand) : Expr;
public sealed record ParenExpr(Expr Inner) : Expr;
public sealed record ArrayExpr(CellValue[,] Values) : Expr;
public sealed record MissingExpr : Expr;
/// <summary>Call of a computed function value, e.g. LAMBDA(x,x+1)(5).</summary>
public sealed record CallExpr(Expr Callee, Expr[] Args) : Expr;
/// <summary>Already evaluated value passed as an argument (lambda parameters, lifted arguments).</summary>
public sealed record ValueExpr(object Value) : Expr;

/// <summary>Cell / area / whole-column / whole-row reference. Coordinates are zero-based.</summary>
/// <remarks><see cref="Spill"/>: the spill-range operator (A1#).</remarks>
public sealed record RefExpr(string? Sheet, RefKind Kind, int R1, int C1, int R2, int C2, bool AR1, bool AC1, bool AR2, bool AC2, bool Spill = false) : Expr
{
    public CellRange Range => Kind switch
    {
        RefKind.Cols => new CellRange(0, Math.Min(C1, C2), Addr.MaxRows - 1, Math.Max(C1, C2)),
        RefKind.Rows => new CellRange(Math.Min(R1, R2), 0, Math.Max(R1, R2), Addr.MaxCols - 1),
        _ => new CellRange(R1, C1, R2, C2).Normalized
    };
}

public sealed class FormulaParseException(string message, int pos) : Exception(message)
{
    public int Position { get; } = pos;
}

internal enum TokType { Number, String, Bool, Error, Ref, Name, Func, Op, LParen, RParen, Comma, Semicolon, LBrace, RBrace, End }

internal readonly record struct Token(TokType Type, string Text, int Pos, double Num = 0, RefExpr? Ref = null);

/// <summary>Tokenizer + recursive-descent parser for Excel formulas (en-US syntax).</summary>
public sealed class FormulaParser
{
    private readonly string _s;
    private readonly List<Token> _toks = new();
    private int _i;

    private FormulaParser(string s) { _s = s; }

    public static Expr Parse(string formula)
    {
        if (formula.StartsWith('=')) formula = formula[1..];
        var p = new FormulaParser(formula);
        p.Lex();
        var e = p.ParseCompare();
        if (p.Peek.Type != TokType.End) throw new FormulaParseException("Unexpected '" + p.Peek.Text + "'", p.Peek.Pos);
        return e;
    }

    public static bool TryParse(string formula, out Expr? e, out string? error)
    {
        try { e = Parse(formula); error = null; return true; }
        catch (FormulaParseException ex) { e = null; error = ex.Message; return false; }
    }

    // ------------------------------------------------------------------ lexer

    private void Lex()
    {
        int i = 0;
        while (i < _s.Length)
        {
            char ch = _s[i];
            if (char.IsWhiteSpace(ch)) { i++; continue; }
            int start = i;
            if (ch == '"')
            {
                var sb = new StringBuilder();
                i++;
                while (true)
                {
                    if (i >= _s.Length) throw new FormulaParseException("Unterminated string", start);
                    if (_s[i] == '"') { if (i + 1 < _s.Length && _s[i + 1] == '"') { sb.Append('"'); i += 2; continue; } i++; break; }
                    sb.Append(_s[i++]);
                }
                _toks.Add(new Token(TokType.String, sb.ToString(), start));
                continue;
            }
            if (ch == '#')
            {
                var code = Errors.All.FirstOrDefault(e => string.Compare(_s, i, e, 0, e.Length, StringComparison.OrdinalIgnoreCase) == 0);
                if (code == null) throw new FormulaParseException("Unknown error value", i);
                i += code.Length;
                // "#REF!A1" style leftovers
                _toks.Add(new Token(TokType.Error, code, start));
                continue;
            }
            if (char.IsAsciiDigit(ch) || (ch == '.' && i + 1 < _s.Length && char.IsAsciiDigit(_s[i + 1])))
            {
                // could be a row range "1:3"
                if (TryLexRef(ref i)) continue;
                int j = i;
                while (j < _s.Length && (char.IsAsciiDigit(_s[j]) || _s[j] == '.')) j++;
                if (j < _s.Length && (_s[j] == 'e' || _s[j] == 'E') && j + 1 < _s.Length && (char.IsAsciiDigit(_s[j + 1]) || ((_s[j + 1] == '+' || _s[j + 1] == '-') && j + 2 < _s.Length && char.IsAsciiDigit(_s[j + 2]))))
                {
                    j += 2;
                    while (j < _s.Length && char.IsAsciiDigit(_s[j])) j++;
                }
                if (!double.TryParse(_s[i..j], NumberStyles.Float, CultureInfo.InvariantCulture, out double d)) throw new FormulaParseException("Bad number", i);
                _toks.Add(new Token(TokType.Number, _s[i..j], start, d));
                i = j;
                continue;
            }
            if (ch == '\'' || ch == '$' || char.IsLetter(ch) || ch == '_' || ch == '\\')
            {
                if (TryLexRef(ref i)) continue;
                if (ch == '\'') throw new FormulaParseException("Bad sheet reference", i);
                int j = i;
                while (j < _s.Length && (char.IsLetterOrDigit(_s[j]) || _s[j] == '_' || _s[j] == '.' || _s[j] == '\\')) j++;
                var word = _s[i..j];
                int k = j;
                while (k < _s.Length && _s[k] == ' ') k++;
                if (k < _s.Length && _s[k] == '(')
                {
                    var fn = word.ToUpperInvariant();
                    foreach (var pre in new[] { "_XLFN._XLWS.", "_XLFN.", "_XLWS.", "_XLPM." }) if (fn.StartsWith(pre)) fn = fn[pre.Length..];
                    _toks.Add(new Token(TokType.Func, fn, start));
                    i = k + 1;
                    continue;
                }
                if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) _toks.Add(new Token(TokType.Bool, "TRUE", start, 1));
                else if (word.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) _toks.Add(new Token(TokType.Bool, "FALSE", start, 0));
                else _toks.Add(new Token(TokType.Name, word.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase) ? word[6..] : word, start));
                i = j;
                continue;
            }
            switch (ch)
            {
                case '(': _toks.Add(new Token(TokType.LParen, "(", i)); i++; continue;
                case ')': _toks.Add(new Token(TokType.RParen, ")", i)); i++; continue;
                case ',': _toks.Add(new Token(TokType.Comma, ",", i)); i++; continue;
                case ';': _toks.Add(new Token(TokType.Semicolon, ";", i)); i++; continue;
                case '{': _toks.Add(new Token(TokType.LBrace, "{", i)); i++; continue;
                case '}': _toks.Add(new Token(TokType.RBrace, "}", i)); i++; continue;
                case '<':
                    if (i + 1 < _s.Length && (_s[i + 1] == '=' || _s[i + 1] == '>')) { _toks.Add(new Token(TokType.Op, _s.Substring(i, 2), i)); i += 2; }
                    else { _toks.Add(new Token(TokType.Op, "<", i)); i++; }
                    continue;
                case '>':
                    if (i + 1 < _s.Length && _s[i + 1] == '=') { _toks.Add(new Token(TokType.Op, ">=", i)); i += 2; }
                    else { _toks.Add(new Token(TokType.Op, ">", i)); i++; }
                    continue;
                case '[':
                {
                    // optional LAMBDA parameter: [name]
                    int j = i + 1;
                    while (j < _s.Length && (char.IsLetterOrDigit(_s[j]) || _s[j] == '_' || _s[j] == '.')) j++;
                    if (j > i + 1 && j < _s.Length && _s[j] == ']')
                    {
                        var w = _s[(i + 1)..j];
                        if (w.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase)) w = w[6..];
                        _toks.Add(new Token(TokType.Name, "[" + w + "]", i));
                        i = j + 1;
                        continue;
                    }
                    break;
                }
                case '@': _toks.Add(new Token(TokType.Op, "@", i)); i++; continue;
                case '+': case '-': case '*': case '/': case '^': case '&': case '=': case '%':
                    _toks.Add(new Token(TokType.Op, ch.ToString(), i)); i++; continue;
            }
            throw new FormulaParseException("Unexpected character '" + ch + "'", i);
        }
        _toks.Add(new Token(TokType.End, "", _s.Length));
    }

    /// <summary>Tries to read [sheet!]A1[:B2] | [sheet!]A:B | [sheet!]1:3 at position i.</summary>
    private bool TryLexRef(ref int i)
    {
        int start = i;
        int j = i;
        string? sheet = null;
        // sheet prefix
        if (_s[j] == '\'')
        {
            var sb = new StringBuilder();
            int k = j + 1;
            while (k < _s.Length)
            {
                if (_s[k] == '\'') { if (k + 1 < _s.Length && _s[k + 1] == '\'') { sb.Append('\''); k += 2; continue; } break; }
                sb.Append(_s[k++]);
            }
            if (k + 1 >= _s.Length || _s[k] != '\'' || _s[k + 1] != '!') return false;
            sheet = sb.ToString();
            j = k + 2;
        }
        else
        {
            int k = j;
            while (k < _s.Length && (char.IsLetterOrDigit(_s[k]) || _s[k] == '_' || _s[k] == '.')) k++;
            if (k < _s.Length && _s[k] == '!' && k > j) { sheet = _s[j..k]; j = k + 1; }
        }
        if (!ReadPart(ref j, out var p1)) return false;
        var p2 = p1;
        int afterFirst = j;
        bool range = false;
        if (j < _s.Length && _s[j] == ':')
        {
            int k = j + 1;
            // optional repeated sheet prefix (Sheet1!A1:Sheet1!B2)
            int save = k;
            if (sheet != null)
            {
                var pre = (_s[k] == '\'' ? "'" + sheet.Replace("'", "''") + "'" : sheet) + "!";
                if (string.Compare(_s, k, pre, 0, pre.Length, StringComparison.OrdinalIgnoreCase) == 0) k += pre.Length;
            }
            if (ReadPart(ref k, out p2) && p2.Kind == p1.Kind) { j = k; range = true; }
            else { k = save; j = afterFirst; p2 = p1; }
        }
        if (!range && p1.Kind != 0) return false; // bare "A" or "3" is not a reference
        // must not be followed by an identifier char or '(' (function like LOG10( ) )
        if (j < _s.Length && (char.IsLetterOrDigit(_s[j]) || _s[j] == '_' || _s[j] == '(' || _s[j] == '.')) return false;
        bool spill = false;
        if (!range && p1.Kind == 0 && j < _s.Length && _s[j] == '#' && !Errors.All.Any(e => string.Compare(_s, j, e, 0, e.Length, StringComparison.OrdinalIgnoreCase) == 0))
        {
            spill = true;
            j++;
        }
        RefExpr r;
        if (p1.Kind == 0)
            r = new RefExpr(sheet, range ? RefKind.Area : RefKind.Cell, p1.Row, p1.Col, p2.Row, p2.Col, p1.AbsR, p1.AbsC, p2.AbsR, p2.AbsC, spill);
        else if (p1.Kind == 1)
            r = new RefExpr(sheet, RefKind.Cols, 0, p1.Col, Addr.MaxRows - 1, p2.Col, false, p1.AbsC, false, p2.AbsC);
        else
            r = new RefExpr(sheet, RefKind.Rows, p1.Row, 0, p2.Row, Addr.MaxCols - 1, p1.AbsR, false, p2.AbsR, false);
        _toks.Add(new Token(TokType.Ref, _s[start..j], start, 0, r));
        i = j;
        return true;
    }

    private readonly record struct Part(int Kind, int Row, int Col, bool AbsR, bool AbsC); // Kind 0 cell, 1 col, 2 row

    private bool ReadPart(ref int j, out Part p)
    {
        p = default;
        int k = j;
        bool absC = false, absR = false;
        if (k < _s.Length && _s[k] == '$') { absC = true; k++; }
        int ls = k;
        while (k < _s.Length && char.IsAsciiLetter(_s[k])) k++;
        int letters = k - ls;
        if (letters > 3) return false;
        int col = letters > 0 ? Addr.ColIndex(_s[ls..k]) : -1;
        if (letters > 0 && col < 0) return false;
        if (k < _s.Length && _s[k] == '$') { absR = true; k++; }
        int ds = k;
        while (k < _s.Length && char.IsAsciiDigit(_s[k])) k++;
        int digits = k - ds;
        if (letters > 0 && digits > 0)
        {
            if (!int.TryParse(_s[ds..k], out int row) || row < 1 || row > Addr.MaxRows) return false;
            p = new Part(0, row - 1, col, absR, absC);
        }
        else if (letters > 0 && digits == 0 && !absR)
            p = new Part(1, 0, col, false, absC);
        else if (letters == 0 && digits > 0)
        {
            if (!int.TryParse(_s[ds..k], out int row) || row < 1 || row > Addr.MaxRows) return false;
            p = new Part(2, row - 1, 0, absC, false); // "$3" → the '$' was read as column-abs flag
        }
        else return false;
        j = k;
        return true;
    }

    // ------------------------------------------------------------------ parser

    private Token Peek => _toks[_i];
    private Token Next() => _toks[_i++];
    private bool IsOp(string op) => Peek.Type == TokType.Op && Peek.Text == op;

    private Expr ParseCompare()
    {
        var l = ParseConcat();
        while (Peek.Type == TokType.Op && Peek.Text is "=" or "<>" or "<" or ">" or "<=" or ">=")
        {
            var op = Next().Text;
            l = new BinaryExpr(op, l, ParseConcat());
        }
        return l;
    }

    private Expr ParseConcat()
    {
        var l = ParseAdd();
        while (IsOp("&")) { Next(); l = new BinaryExpr("&", l, ParseAdd()); }
        return l;
    }

    private Expr ParseAdd()
    {
        var l = ParseMul();
        while (IsOp("+") || IsOp("-")) { var op = Next().Text; l = new BinaryExpr(op, l, ParseMul()); }
        return l;
    }

    private Expr ParseMul()
    {
        var l = ParsePow();
        while (IsOp("*") || IsOp("/")) { var op = Next().Text; l = new BinaryExpr(op, l, ParsePow()); }
        return l;
    }

    private Expr ParsePow()
    {
        var l = ParseUnary();
        while (IsOp("^")) { Next(); l = new BinaryExpr("^", l, ParseUnary()); }
        return l;
    }

    private Expr ParseUnary()
    {
        if (IsOp("-")) { Next(); return new UnaryExpr('-', ParseUnary()); }
        if (IsOp("+")) { Next(); return new UnaryExpr('+', ParseUnary()); }
        if (IsOp("@")) { Next(); return new UnaryExpr('@', ParseUnary()); }
        var e = ParsePrimary();
        while (Peek.Type == TokType.LParen && e is FuncExpr or ParenExpr or CallExpr) e = new CallExpr(e, ParseArgs("(call)"));
        while (IsOp("%")) { Next(); e = new PercentExpr(e); }
        return e;
    }

    private Expr ParsePrimary()
    {
        var t = Next();
        switch (t.Type)
        {
            case TokType.Number: return new NumberExpr(t.Num);
            case TokType.String: return new StringExpr(t.Text);
            case TokType.Bool: return new BoolExpr(t.Num != 0);
            case TokType.Error: return new ErrorExpr(t.Text);
            case TokType.Ref: return t.Ref!;
            case TokType.Name: return new NameExpr(t.Text);
            case TokType.LParen:
            {
                var e = ParseCompare();
                if (Next().Type != TokType.RParen) throw new FormulaParseException("Missing ')'", t.Pos);
                return new ParenExpr(e);
            }
            case TokType.Func:
            {
                if (Peek.Type == TokType.RParen) { Next(); return new FuncExpr(t.Text, []); }
                return new FuncExpr(t.Text, ParseArgsAfterParen(t.Text));
            }
            case TokType.LBrace: return ParseArray(t.Pos);
            case TokType.Op when t.Text == "-" || t.Text == "+":
                _i--; return ParseUnary();
        }
        throw new FormulaParseException(t.Type == TokType.End ? "Incomplete formula" : "Unexpected '" + t.Text + "'", t.Pos);
    }

    private Expr[] ParseArgs(string what)
    {
        Next(); // '('
        if (Peek.Type == TokType.RParen) { Next(); return []; }
        return ParseArgsAfterParen(what);
    }

    private Expr[] ParseArgsAfterParen(string what)
    {
        var args = new List<Expr>();
        while (true)
        {
            if (Peek.Type is TokType.Comma or TokType.RParen) args.Add(new MissingExpr());
            else args.Add(ParseCompare());
            var sep = Next();
            if (sep.Type == TokType.RParen) break;
            if (sep.Type != TokType.Comma) throw new FormulaParseException("Expected ',' or ')' in " + what, sep.Pos);
        }
        return args.ToArray();
    }

    private Expr ParseArray(int pos)
    {
        var rows = new List<List<CellValue>> { new() };
        while (true)
        {
            bool neg = false;
            if (IsOp("-")) { Next(); neg = true; }
            var t = Next();
            CellValue v = t.Type switch
            {
                TokType.Number => CellValue.Number(neg ? -t.Num : t.Num),
                TokType.String => CellValue.Text(t.Text),
                TokType.Bool => CellValue.Bool(t.Num != 0),
                TokType.Error => CellValue.Error(t.Text),
                _ => throw new FormulaParseException("Bad array constant", t.Pos)
            };
            rows[^1].Add(v);
            var sep = Next();
            if (sep.Type == TokType.RBrace) break;
            if (sep.Type == TokType.Semicolon) rows.Add(new());
            else if (sep.Type != TokType.Comma) throw new FormulaParseException("Bad array constant", sep.Pos);
        }
        int cols = rows.Max(r => r.Count);
        var arr = new CellValue[rows.Count, cols];
        for (int r = 0; r < rows.Count; r++)
            for (int c = 0; c < cols; c++) arr[r, c] = c < rows[r].Count ? rows[r][c] : CellValue.Error(Errors.NA);
        return new ArrayExpr(arr);
    }
}

/// <summary>Converts an AST back to formula text (no leading '=').</summary>
public static class FormulaPrinter
{
    public static string Print(Expr e)
    {
        var sb = new StringBuilder();
        Write(sb, e);
        return sb.ToString();
    }

    private static void Write(StringBuilder sb, Expr e)
    {
        switch (e)
        {
            case NumberExpr n: sb.Append(CellValue.FormatGeneralInvariant(n.Value)); break;
            case StringExpr s: sb.Append('"').Append(s.Value.Replace("\"", "\"\"")).Append('"'); break;
            case BoolExpr b: sb.Append(b.Value ? "TRUE" : "FALSE"); break;
            case ErrorExpr er: sb.Append(er.Code); break;
            case NameExpr nm: sb.Append(nm.Name); break;
            case MissingExpr: break;
            case RefExpr r: sb.Append(RefText(r)); break;
            case ParenExpr p: sb.Append('('); Write(sb, p.Inner); sb.Append(')'); break;
            case PercentExpr pc: Write(sb, pc.Operand); sb.Append('%'); break;
            case UnaryExpr u: sb.Append(u.Op); Write(sb, u.Operand); break;
            case CallExpr ce:
                Write(sb, ce.Callee);
                sb.Append('(');
                for (int i = 0; i < ce.Args.Length; i++) { if (i > 0) sb.Append(','); Write(sb, ce.Args[i]); }
                sb.Append(')');
                break;
            case BinaryExpr b: Write(sb, b.Left); sb.Append(b.Op); Write(sb, b.Right); break;
            case FuncExpr f:
                sb.Append(f.Name).Append('(');
                for (int i = 0; i < f.Args.Length; i++) { if (i > 0) sb.Append(','); Write(sb, f.Args[i]); }
                sb.Append(')');
                break;
            case ArrayExpr a:
                sb.Append('{');
                for (int r = 0; r < a.Values.GetLength(0); r++)
                {
                    if (r > 0) sb.Append(';');
                    for (int c = 0; c < a.Values.GetLength(1); c++)
                    {
                        if (c > 0) sb.Append(',');
                        var v = a.Values[r, c];
                        sb.Append(v.IsText ? "\"" + v.Str!.Replace("\"", "\"\"") + "\"" : v.ToString());
                    }
                }
                sb.Append('}');
                break;
        }
    }

    public static string RefText(RefExpr r)
    {
        var prefix = r.Sheet != null ? Addr.QuoteSheet(r.Sheet) + "!" : "";
        return r.Kind switch
        {
            RefKind.Cell => prefix + Addr.ToA1(r.R1, r.C1, r.AR1, r.AC1) + (r.Spill ? "#" : ""),
            RefKind.Area => prefix + Addr.ToA1(r.R1, r.C1, r.AR1, r.AC1) + ":" + Addr.ToA1(r.R2, r.C2, r.AR2, r.AC2),
            RefKind.Cols => prefix + (r.AC1 ? "$" : "") + Addr.ColName(r.C1) + ":" + (r.AC2 ? "$" : "") + Addr.ColName(r.C2),
            _ => prefix + (r.AR1 ? "$" : "") + (r.R1 + 1) + ":" + (r.AR2 ? "$" : "") + (r.R2 + 1)
        };
    }
}
