using System.Globalization;
using System.Text;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Formula;

public static partial class Functions
{
    // ================================================================== LOOKUP & REFERENCE

    /// <summary>Position (0-based) of <paramref name="key"/> in a 1-D list. mode: 0 exact (wildcards), 1 largest ≤ (sorted asc), -1 smallest ≥ (sorted desc).</summary>
    internal static int Find(IReadOnlyList<CellValue> list, CellValue key, int mode, bool wildcards = true)
    {
        if (key.IsError) throw Err(key.Str!);
        if (mode == 0)
        {
            if (key.IsText && wildcards && (key.Str!.Contains('*') || key.Str.Contains('?') || key.Str.Contains('~')))
            {
                var rx = WildcardRegex(key.Str!);
                for (int i = 0; i < list.Count; i++) if (list[i].IsText && rx.IsMatch(list[i].Str!)) return i;
                return -1;
            }
            for (int i = 0; i < list.Count; i++)
            {
                var v = list[i];
                if (v.Kind == key.Kind || (key.IsBlank && v.IsBlank)) { if (Calc.Compare(v, key) == 0) return i; }
            }
            return -1;
        }
        // approximate: binary search like Excel on same-type values
        int lo = 0, hi = list.Count - 1, found = -1;
        if (mode == 1)
        {
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                var v = list[mid];
                int k = mid;
                while (k <= hi && (list[k].IsBlank || list[k].Kind != key.Kind && !(list[k].IsNumber && key.IsNumber))) k++;
                if (k > hi) { hi = mid - 1; continue; }
                v = list[k];
                int cmp = Calc.Compare(v, key);
                if (cmp <= 0) { found = k; lo = k + 1; }
                else hi = mid - 1;
            }
            return found;
        }
        // -1: descending, smallest value >= key
        for (int i = 0; i < list.Count; i++)
        {
            var v = list[i];
            if (v.Kind != key.Kind) continue;
            if (Calc.Compare(v, key) >= 0) found = i; else break;
        }
        return found;
    }

    private static void RegisterLookup()
    {
        Reg((c, a) =>
        {
            Args(a, 3, 4);
            var key = c.Scalar(a[0]);
            var g = Grid(c, a[1]);
            int col = (int)c.Num(a[2]);
            bool approx = c.BoolOr(a, 3, true);
            if (col < 1) throw Err(Errors.Value);
            if (col > g.GetLength(1)) throw Err(Errors.Ref);
            var first = Enumerable.Range(0, g.GetLength(0)).Select(r => g[r, 0]).ToList();
            int idx = Find(first, key, approx ? 1 : 0);
            if (idx < 0) throw Err(Errors.NA);
            return g[idx, col - 1];
        }, "VLOOKUP");
        Reg((c, a) =>
        {
            Args(a, 3, 4);
            var key = c.Scalar(a[0]);
            var g = Grid(c, a[1]);
            int row = (int)c.Num(a[2]);
            bool approx = c.BoolOr(a, 3, true);
            if (row < 1) throw Err(Errors.Value);
            if (row > g.GetLength(0)) throw Err(Errors.Ref);
            var first = Enumerable.Range(0, g.GetLength(1)).Select(k => g[0, k]).ToList();
            int idx = Find(first, key, approx ? 1 : 0);
            if (idx < 0) throw Err(Errors.NA);
            return g[row - 1, idx];
        }, "HLOOKUP");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var key = c.Scalar(a[0]);
            var g = Grid(c, a[1]);
            bool vertical = g.GetLength(0) >= g.GetLength(1);
            var look = vertical ? Enumerable.Range(0, g.GetLength(0)).Select(r => g[r, 0]).ToList() : Enumerable.Range(0, g.GetLength(1)).Select(k => g[0, k]).ToList();
            int idx = Find(look, key, 1);
            if (idx < 0) throw Err(Errors.NA);
            if (a.Length == 3)
            {
                var res = Grid(c, a[2]).Cast<CellValue>().ToList();
                return idx < res.Count ? res[idx] : throw Err(Errors.NA);
            }
            return vertical ? g[idx, g.GetLength(1) - 1] : g[g.GetLength(0) - 1, idx];
        }, "LOOKUP");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var key = c.Scalar(a[0]);
            var list = Grid(c, a[1]).Cast<CellValue>().ToList();
            int type = (int)c.NumOr(a, 2, 1);
            int idx = Find(list, key, Math.Sign(type));
            if (idx < 0) throw Err(Errors.NA);
            return N(idx + 1);
        }, "MATCH");
        Reg((c, a) =>
        {
            Args(a, 2, 4);
            var key = c.Scalar(a[0]);
            var list = Grid(c, a[1]).Cast<CellValue>().ToList();
            int mode = (int)c.NumOr(a, 2, 0), search = (int)c.NumOr(a, 3, 1);
            int idx = XFind(list, key, mode, search);
            if (idx < 0) throw Err(Errors.NA);
            return N(idx + 1);
        }, "XMATCH");
        Reg((c, a) =>
        {
            Args(a, 3, 6);
            var key = c.Scalar(a[0]);
            var lg = Grid(c, a[1]);
            var rv = c.Eval(a[2]);
            var rg = Grid(c, rv);
            int mode = (int)c.NumOr(a, 4, 0), search = (int)c.NumOr(a, 5, 1);
            bool vertical = lg.GetLength(1) == 1;
            var list = lg.Cast<CellValue>().ToList();
            int idx = XFind(list, key, mode, search);
            if (idx < 0)
            {
                if (a.Length > 3 && a[3] is not MissingExpr) return c.Eval(a[3]);
                throw Err(Errors.NA);
            }
            if (vertical)
            {
                if (idx >= rg.GetLength(0)) throw Err(Errors.Value);
                if (rg.GetLength(1) == 1) return rg[idx, 0];
                var row = new CellValue[1, rg.GetLength(1)];
                for (int k = 0; k < rg.GetLength(1); k++) row[0, k] = rg[idx, k];
                return new ArrayVal(row);
            }
            if (idx >= rg.GetLength(1)) throw Err(Errors.Value);
            if (rg.GetLength(0) == 1) return rg[0, idx];
            var colv = new CellValue[rg.GetLength(0), 1];
            for (int r = 0; r < rg.GetLength(0); r++) colv[r, 0] = rg[r, idx];
            return new ArrayVal(colv);
        }, "XLOOKUP");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var v = c.Eval(a[0]);
            int row = (int)c.Num(a[1]);
            int col = a.Length > 2 && a[2] is not MissingExpr ? (int)c.Num(a[2]) : 0;
            if (v is RangeVal r)
            {
                var rg = r.Range;
                if (rg.Rows == 1 && a.Length == 2) { col = row; row = 1; }
                if (row < 0 || col < 0 || row > rg.Rows || col > rg.Cols) throw Err(Errors.Ref);
                if (row == 0 && col == 0) return r;
                if (row == 0) return new RangeVal(r.Sheet, new CellRange(rg.R1, rg.C1 + col - 1, rg.R2, rg.C1 + col - 1));
                if (col == 0) { if (rg.Cols == 1) col = 1; else return new RangeVal(r.Sheet, new CellRange(rg.R1 + row - 1, rg.C1, rg.R1 + row - 1, rg.C2)); }
                return new RangeVal(r.Sheet, CellRange.Single(rg.R1 + row - 1, rg.C1 + col - 1));
            }
            var g = Grid(c, v);
            if (g.GetLength(0) == 1 && a.Length == 2) { col = row; row = 1; }
            else if (g.GetLength(1) == 1 && col == 0) col = 1;
            if (g.GetLength(0) == 1 && row == 0) row = 1;
            if (row < 0 || col < 0 || row > g.GetLength(0) || col > g.GetLength(1)) throw Err(Errors.Ref);
            if (row == 0 && col == 0) return v;
            if (row == 0)
            {
                var colv = new CellValue[g.GetLength(0), 1];
                for (int i = 0; i < g.GetLength(0); i++) colv[i, 0] = g[i, col - 1];
                return new ArrayVal(colv);
            }
            if (col == 0)
            {
                var rowv = new CellValue[1, g.GetLength(1)];
                for (int k = 0; k < g.GetLength(1); k++) rowv[0, k] = g[row - 1, k];
                return new ArrayVal(rowv);
            }
            return g[row - 1, col - 1];
        }, "INDEX");
        Reg((c, a) =>
        {
            Args(a, 3, 5);
            if (c.Eval(a[0]) is not RangeVal r) throw Err(Errors.Value);
            int dr = (int)c.Num(a[1]), dc = (int)c.Num(a[2]);
            int h = a.Length > 3 && a[3] is not MissingExpr ? (int)c.Num(a[3]) : r.Rows;
            int w = a.Length > 4 && a[4] is not MissingExpr ? (int)c.Num(a[4]) : r.Cols;
            if (h < 1 || w < 1) throw Err(Errors.Ref);
            int r1 = r.Range.R1 + dr, c1 = r.Range.C1 + dc;
            if (r1 < 0 || c1 < 0 || r1 + h - 1 >= Addr.MaxRows || c1 + w - 1 >= Addr.MaxCols) throw Err(Errors.Ref);
            return new RangeVal(r.Sheet, new CellRange(r1, c1, r1 + h - 1, c1 + w - 1));
        }, "OFFSET");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            var text = c.Str(a[0]).Trim();
            if (text.Length == 0) throw Err(Errors.Ref);
            if (c.Book.Names.ContainsKey(text)) return c.Eval(new NameExpr(text));
            Expr e;
            try { e = FormulaParser.Parse(text); } catch { throw Err(Errors.Ref); }
            if (e is not RefExpr) throw Err(Errors.Ref);
            return c.Eval(e);
        }, "INDIRECT");
        Reg((c, a) =>
        {
            if (a.Length == 0 || a[0] is MissingExpr) return N(c.Row + 1);
            if (c.Eval(a[0]) is not RangeVal r) throw Err(Errors.Value);
            if (r.Range.R1 == r.Range.R2 || r.Range.Rows > Addr.MaxRows / 2) return N(r.Range.R1 + 1);
            var res = new CellValue[r.Range.Rows, 1];
            for (int i = 0; i < r.Range.Rows; i++) res[i, 0] = N(r.Range.R1 + 1 + i);
            return new ArrayVal(res);
        }, "ROW");
        Reg((c, a) =>
        {
            if (a.Length == 0 || a[0] is MissingExpr) return N(c.Col + 1);
            if (c.Eval(a[0]) is not RangeVal r) throw Err(Errors.Value);
            if (r.Range.C1 == r.Range.C2) return N(r.Range.C1 + 1);
            var res = new CellValue[1, r.Range.Cols];
            for (int i = 0; i < r.Range.Cols; i++) res[0, i] = N(r.Range.C1 + 1 + i);
            return new ArrayVal(res);
        }, "COLUMN");
        Reg((c, a) => { Args(a, 1, 1); var v = c.Eval(a[0]); return N(v is RangeVal r ? r.Rows : v is ArrayVal av ? av.Rows : 1); }, "ROWS");
        Reg((c, a) => { Args(a, 1, 1); var v = c.Eval(a[0]); return N(v is RangeVal r ? r.Cols : v is ArrayVal av ? av.Cols : 1); }, "COLUMNS");
        Reg((c, a) =>
        {
            Args(a, 2, 5);
            int row = (int)c.Num(a[0]), col = (int)c.Num(a[1]);
            int abs = (int)c.NumOr(a, 2, 1);
            bool a1 = c.BoolOr(a, 3, true);
            string sheet = c.StrOr(a, 4, "");
            if (row < 1 || col < 1) throw Err(Errors.Value);
            string s = a1
                ? Addr.ToA1(row - 1, col - 1, abs is 1 or 2, abs is 1 or 3)
                : (abs is 1 or 2 ? "R" + row : "R[" + row + "]") + (abs is 1 or 3 ? "C" + col : "C[" + col + "]");
            return S(sheet.Length > 0 ? Addr.QuoteSheet(sheet) + "!" + s : s);
        }, "ADDRESS");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var g = Grid(c, a[0]);
            var t = new CellValue[g.GetLength(1), g.GetLength(0)];
            for (int i = 0; i < g.GetLength(0); i++) for (int j = 0; j < g.GetLength(1); j++) t[j, i] = g[i, j];
            return new ArrayVal(t);
        }, "TRANSPOSE");
        Reg((c, a) => { Args(a, 1, 2); return a.Length > 1 ? c.Scalar(a[1]) : S(c.Str(a[0])); }, "HYPERLINK");
    }

    private static int XFind(List<CellValue> list, CellValue key, int mode, int search)
    {
        IEnumerable<int> order = search < 0 ? Enumerable.Range(0, list.Count).Reverse() : Enumerable.Range(0, list.Count);
        if (mode == 2 || mode == 0)
        {
            var rx = mode == 2 && key.IsText ? WildcardRegex(key.Str!) : null;
            foreach (var i in order)
            {
                var v = list[i];
                if (rx != null ? v.IsText && rx.IsMatch(v.Str!) : v.Kind == key.Kind && Calc.Compare(v, key) == 0) return i;
            }
            return -1;
        }
        int best = -1;
        foreach (var i in order)
        {
            var v = list[i];
            if (v.Kind != key.Kind) continue;
            int cmp = Calc.Compare(v, key);
            if (cmp == 0) return i;
            if (mode == -1 && cmp < 0 && (best < 0 || Calc.Compare(v, list[best]) > 0)) best = i;
            if (mode == 1 && cmp > 0 && (best < 0 || Calc.Compare(v, list[best]) < 0)) best = i;
        }
        return best;
    }

    // ================================================================== TEXT

    private static readonly Encoding Win1252 = CodePagesEncoding();
    private static Encoding CodePagesEncoding()
    {
        try { return Encoding.Latin1; } catch { return Encoding.ASCII; }
    }

    private static void RegisterText()
    {
        Reg((c, a) => { Args(a, 1, 2); var s = c.Str(a[0]); int n = (int)c.NumOr(a, 1, 1); if (n < 0) throw Err(Errors.Value); return S(s[..Math.Min(n, s.Length)]); }, "LEFT");
        Reg((c, a) => { Args(a, 1, 2); var s = c.Str(a[0]); int n = (int)c.NumOr(a, 1, 1); if (n < 0) throw Err(Errors.Value); return S(s[Math.Max(0, s.Length - n)..]); }, "RIGHT");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            var s = c.Str(a[0]); int start = (int)c.Num(a[1]), n = (int)c.Num(a[2]);
            if (start < 1 || n < 0) throw Err(Errors.Value);
            if (start > s.Length) return S("");
            return S(s.Substring(start - 1, Math.Min(n, s.Length - start + 1)));
        }, "MID");
        Reg((c, a) => { Args(a, 1, 1); return N(c.Str(a[0]).Length); }, "LEN");
        Reg((c, a) => { Args(a, 1, 1); return S(c.Str(a[0]).ToLower(CultureInfo.CurrentCulture)); }, "LOWER");
        Reg((c, a) => { Args(a, 1, 1); return S(c.Str(a[0]).ToUpper(CultureInfo.CurrentCulture)); }, "UPPER");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var s = c.Str(a[0]);
            var sb = new StringBuilder(s.Length);
            bool prevLetter = false;
            foreach (var ch in s)
            {
                sb.Append(prevLetter ? char.ToLower(ch) : char.ToUpper(ch));
                prevLetter = char.IsLetter(ch);
            }
            return S(sb.ToString());
        }, "PROPER");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var s = c.Str(a[0]).Trim(' ');
            var sb = new StringBuilder();
            bool sp = false;
            foreach (var ch in s) { if (ch == ' ') { if (!sp) sb.Append(ch); sp = true; } else { sb.Append(ch); sp = false; } }
            return S(sb.ToString());
        }, "TRIM");
        Reg((c, a) => { Args(a, 1, 1); return S(new string(c.Str(a[0]).Where(ch => ch >= 32).ToArray())); }, "CLEAN");
        Reg((c, a) =>
        {
            var sb = new StringBuilder();
            foreach (var e in a) foreach (var v in FlatVal(c, c.Eval(e))) sb.Append(Calc.ToText(v));
            return S(sb.ToString());
        }, "CONCAT");
        Reg((c, a) => S(string.Concat(a.Select(e => c.Str(e)))), "CONCATENATE");
        Reg((c, a) =>
        {
            Args(a, 3);
            var delim = c.Str(a[0]);
            bool ignore = c.Bool(a[1]);
            var parts = new List<string>();
            foreach (var e in a.Skip(2))
                foreach (var v in FlatVal(c, c.Eval(e)))
                {
                    var t = Calc.ToText(v);
                    if (ignore && t.Length == 0) continue;
                    parts.Add(t);
                }
            return S(string.Join(delim, parts));
        }, "TEXTJOIN");
        Reg((c, a) =>
        {
            Args(a, 3, 4);
            var s = c.Str(a[0]); var old = c.Str(a[1]); var nw = c.Str(a[2]);
            if (old.Length == 0) return S(s);
            if (a.Length < 4) return S(s.Replace(old, nw, StringComparison.Ordinal));
            int inst = (int)c.Num(a[3]);
            if (inst < 1) throw Err(Errors.Value);
            int idx = -1;
            for (int i = 0; i < inst; i++) { idx = s.IndexOf(old, idx + 1, StringComparison.Ordinal); if (idx < 0) return S(s); }
            return S(s[..idx] + nw + s[(idx + old.Length)..]);
        }, "SUBSTITUTE");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            var s = c.Str(a[0]); int start = (int)c.Num(a[1]), n = (int)c.Num(a[2]); var nw = c.Str(a[3]);
            if (start < 1 || n < 0) throw Err(Errors.Value);
            start = Math.Min(start - 1, s.Length);
            n = Math.Min(n, s.Length - start);
            return S(s[..start] + nw + s[(start + n)..]);
        }, "REPLACE");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var f = c.Str(a[0]); var s = c.Str(a[1]); int start = (int)c.NumOr(a, 2, 1);
            if (start < 1 || start > s.Length + 1) throw Err(Errors.Value);
            int i = s.IndexOf(f, start - 1, StringComparison.Ordinal);
            if (i < 0) throw Err(Errors.Value);
            return N(i + 1);
        }, "FIND");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var f = c.Str(a[0]); var s = c.Str(a[1]); int start = (int)c.NumOr(a, 2, 1);
            if (start < 1 || start > s.Length + 1) throw Err(Errors.Value);
            if (f.IndexOfAny(['*', '?', '~']) >= 0)
            {
                var pat = WildcardRegex(f).ToString().Trim('^', '$');
                var m = System.Text.RegularExpressions.Regex.Match(s[(start - 1)..], pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!m.Success) throw Err(Errors.Value);
                return N(m.Index + start);
            }
            int i = s.IndexOf(f, start - 1, StringComparison.CurrentCultureIgnoreCase);
            if (i < 0) throw Err(Errors.Value);
            return N(i + 1);
        }, "SEARCH");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var v = c.Scalar(a[0]);
            var fmt = c.Str(a[1]);
            if (v.IsText && ValueParser.TryCoerceNumber(v.Str!, out var d)) v = N(d);
            return S(NumFmt.Format(v, fmt, out _));
        }, "TEXT");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var v = c.Scalar(a[0]);
            if (v.IsNumber) return v;
            if (v.IsBlank) return N(0);
            var s = Calc.ToText(v);
            if (ValueParser.TryCoerceNumber(s, out var d)) return N(d);
            throw Err(Errors.Value);
        }, "VALUE");
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            var s = c.Str(a[0]).Replace(" ", "");
            var dec = c.StrOr(a, 1, ".");
            var grp = c.StrOr(a, 2, ",");
            if (grp.Length > 0) s = s.Replace(grp, "");
            if (dec.Length > 0 && dec != ".") s = s.Replace(dec, ".");
            bool pct = s.EndsWith('%');
            if (pct) s = s.TrimEnd('%');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) throw Err(Errors.Value);
            return N(pct ? d / 100 : d);
        }, "NUMBERVALUE");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var s = c.Str(a[0]); int n = (int)c.Num(a[1]);
            if (n < 0 || (long)s.Length * n > 32767) throw Err(Errors.Value);
            return S(string.Concat(Enumerable.Repeat(s, n)));
        }, "REPT");
        Reg((c, a) => { Args(a, 2, 2); return B(string.Equals(c.Str(a[0]), c.Str(a[1]), StringComparison.Ordinal)); }, "EXACT");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            int n = (int)c.Num(a[0]);
            if (n < 1 || n > 255) throw Err(Errors.Value);
            return S(Encoding.Latin1.GetString([(byte)n]));
        }, "CHAR");
        Reg((c, a) => { Args(a, 1, 1); int n = (int)c.Num(a[0]); if (n < 1 || n > 0x10FFFF) throw Err(Errors.Value); return S(char.ConvertFromUtf32(n)); }, "UNICHAR");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var s = c.Str(a[0]);
            if (s.Length == 0) throw Err(Errors.Value);
            return N(s[0] < 256 ? s[0] : '?');
        }, "CODE");
        Reg((c, a) => { Args(a, 1, 1); var s = c.Str(a[0]); if (s.Length == 0) throw Err(Errors.Value); return N(char.ConvertToUtf32(s, 0)); }, "UNICODE");
        Reg((c, a) =>
        {
            Args(a, 1, 3);
            double x = c.Num(a[0]); int d = (int)c.NumOr(a, 1, 2); bool noCommas = c.BoolOr(a, 2, false);
            x = RoundAway(x, d);
            var code = (noCommas ? "0" : "#,##0") + (d > 0 ? "." + new string('0', d) : "");
            return S(NumFmt.FormatNumber(x, code, out _));
        }, "FIXED");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            double x = c.Num(a[0]); int d = (int)c.NumOr(a, 1, 2);
            x = RoundAway(x, d);
            var num = "#,##0" + (d > 0 ? "." + new string('0', d) : "");
            var code = NumFmt.Vietnamese ? num + " \"₫\";-" + num + " \"₫\"" : "\"$\"" + num + ";(\"$\"" + num + ")";
            return S(NumFmt.FormatNumber(x, code, out _));
        }, "DOLLAR");
        Reg((c, a) => { Args(a, 1, 1); var v = c.Scalar(a[0]); return v.IsText ? v : S(""); }, "T");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var s = c.Str(a[0]); var d = c.Str(a[1]); int inst = (int)c.NumOr(a, 2, 1);
            int idx = NthIndex(s, d, inst);
            if (idx < 0) throw Err(Errors.NA);
            return S(s[..idx]);
        }, "TEXTBEFORE");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var s = c.Str(a[0]); var d = c.Str(a[1]); int inst = (int)c.NumOr(a, 2, 1);
            int idx = NthIndex(s, d, inst);
            if (idx < 0) throw Err(Errors.NA);
            return S(s[(idx + d.Length)..]);
        }, "TEXTAFTER");
        Reg((c, a) => { Args(a, 1, 2); return S(DocSo.Read(c.Num(a[0]), c.StrOr(a, 1, "đồng"))); }, "DOCSO", "VND");
    }

    private static int NthIndex(string s, string d, int n)
    {
        if (n == 0 || d.Length == 0) return -1;
        if (n > 0)
        {
            int idx = -1;
            for (int i = 0; i < n; i++) { idx = s.IndexOf(d, idx + 1, StringComparison.Ordinal); if (idx < 0) return -1; }
            return idx;
        }
        int j = s.Length;
        for (int i = 0; i < -n; i++) { j = j - 1 < 0 ? -1 : s.LastIndexOf(d, j - 1, StringComparison.Ordinal); if (j < 0) return -1; }
        return j;
    }

    // ================================================================== DATE & TIME

    internal static DateTime Dt(Ctx c, Expr e)
    {
        var v = c.Scalar(e);
        double s;
        if (v.IsText)
        {
            if (!ValueParser.TryCoerceNumber(v.Str!, out s)) throw Err(Errors.Value);
        }
        else s = Calc.ToNum(v);
        if (s < 0) throw Err(Errors.Num);
        return DateSerial.FromSerial(s);
    }

    private static double Ser(DateTime d) => DateSerial.ToSerial(d.Date);

    private static void RegisterDate()
    {
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            int y = (int)c.Num(a[0]), m = (int)c.Num(a[1]), d = (int)c.Num(a[2]);
            if (y < 0 || y > 9999) throw Err(Errors.Num);
            if (y < 1900) y += 1900;
            var baseDate = new DateTime(y, 1, 1);
            var dt = baseDate.AddMonths(m - 1).AddDays(d - 1);
            var s = DateSerial.ToSerial(dt);
            if (s < 0) throw Err(Errors.Num);
            return N(s);
        }, "DATE");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            double t = (c.Num(a[0]) * 3600 + c.Num(a[1]) * 60 + c.Num(a[2])) / 86400.0;
            if (t < 0) throw Err(Errors.Num);
            return N(t - Math.Floor(t));
        }, "TIME");
        Reg((c, a) => N(Ser(c.Calc.Now)), "TODAY");
        Reg((c, a) => N(DateSerial.ToSerial(c.Calc.Now)), "NOW");
        Reg((c, a) => { Args(a, 1, 1); return N(Dt(c, a[0]).Year); }, "YEAR");
        Reg((c, a) => { Args(a, 1, 1); return N(Dt(c, a[0]).Month); }, "MONTH");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var v = c.Num(a[0]);
            if (Math.Floor(v) == 60) return N(29);
            if (Math.Floor(v) == 0) return N(0);
            return N(Dt(c, a[0]).Day);
        }, "DAY");
        Reg((c, a) => { Args(a, 1, 1); return N(Dt(c, a[0]).Hour); }, "HOUR");
        Reg((c, a) => { Args(a, 1, 1); return N(Dt(c, a[0]).Minute); }, "MINUTE");
        Reg((c, a) => { Args(a, 1, 1); return N(Dt(c, a[0]).Second); }, "SECOND");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            var d = Dt(c, a[0]);
            int type = (int)c.NumOr(a, 1, 1);
            int dow = (int)d.DayOfWeek; // Sunday = 0
            return type switch
            {
                1 or 17 => N(dow + 1),
                2 or 11 => N(dow == 0 ? 7 : dow),
                3 => N(dow == 0 ? 6 : dow - 1),
                _ when type >= 12 && type <= 16 => N(((dow - (type - 10)) % 7 + 7) % 7 + 1),
                _ => throw Err(Errors.Num)
            };
        }, "WEEKDAY");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            var d = Dt(c, a[0]);
            int type = (int)c.NumOr(a, 1, 1);
            if (type == 21) return N(ISOWeek.GetWeekOfYear(d));
            var startDow = type is 2 or 11 ? DayOfWeek.Monday : DayOfWeek.Sunday;
            var jan1 = new DateTime(d.Year, 1, 1);
            int offset = ((int)jan1.DayOfWeek - (int)startDow + 7) % 7;
            return N((d.DayOfYear + offset - 1) / 7 + 1);
        }, "WEEKNUM");
        Reg((c, a) => { Args(a, 1, 1); return N(ISOWeek.GetWeekOfYear(Dt(c, a[0]))); }, "ISOWEEKNUM");
        Reg((c, a) => { Args(a, 2, 2); return N(Ser(Dt(c, a[0]).AddMonths((int)c.Num(a[1])))); }, "EDATE");
        Reg((c, a) =>
        {
            Args(a, 2, 2);
            var d = Dt(c, a[0]).AddMonths((int)c.Num(a[1]));
            return N(Ser(new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month))));
        }, "EOMONTH");
        Reg((c, a) =>
        {
            Args(a, 3, 3);
            var s = Dt(c, a[0]).Date; var e = Dt(c, a[1]).Date; var unit = c.Str(a[2]).ToUpperInvariant();
            if (s > e) throw Err(Errors.Num);
            int months = (e.Year - s.Year) * 12 + e.Month - s.Month;
            if (e.Day < s.Day) months--;
            return unit switch
            {
                "Y" => N(months / 12),
                "M" => N(months),
                "D" => N((e - s).TotalDays),
                "MD" => N(e.Day >= s.Day ? e.Day - s.Day : (e - e.AddMonths(-1).AddDays(s.Day - e.AddMonths(-1).Day)).Days),
                "YM" => N(months % 12),
                "YD" => N(YdDiff(s, e)),
                _ => throw Err(Errors.Num)
            };
        }, "DATEDIF");
        Reg((c, a) => { Args(a, 2, 2); return N(Math.Floor(DateSerial.ToSerial(Dt(c, a[0]))) - Math.Floor(DateSerial.ToSerial(Dt(c, a[1])))); }, "DAYS");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var s = Dt(c, a[0]); var e = Dt(c, a[1]); bool eu = c.BoolOr(a, 2, false);
            int d1 = s.Day, d2 = e.Day;
            if (eu) { if (d1 == 31) d1 = 30; if (d2 == 31) d2 = 30; }
            else
            {
                if (d1 == 31 || (s.Month == 2 && d1 == DateTime.DaysInMonth(s.Year, 2))) d1 = 30;
                if (d2 == 31 && d1 >= 30) d2 = 30;
            }
            return N((e.Year - s.Year) * 360 + (e.Month - s.Month) * 30 + d2 - d1);
        }, "DAYS360");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var s = Dt(c, a[0]).Date; var e = Dt(c, a[1]).Date;
            var hol = a.Length > 2 ? Holidays(c, a[2]) : new HashSet<DateTime>();
            int sign = 1;
            if (s > e) { (s, e) = (e, s); sign = -1; }
            int n = 0;
            for (var d = s; d <= e; d = d.AddDays(1)) if (d.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday && !hol.Contains(d)) n++;
            return N(sign * n);
        }, "NETWORKDAYS");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var d = Dt(c, a[0]).Date; int days = (int)c.Num(a[1]);
            var hol = a.Length > 2 ? Holidays(c, a[2]) : new HashSet<DateTime>();
            int step = Math.Sign(days);
            while (days != 0)
            {
                d = d.AddDays(step);
                if (d.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday && !hol.Contains(d)) days -= step;
            }
            return N(Ser(d));
        }, "WORKDAY");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var s = c.Str(a[0]);
            if (!ValueParser.TryDateTime(s.Trim(), out var d, out _) && !(DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt) && (d = DateSerial.ToSerial(dt)) >= 0)) throw Err(Errors.Value);
            return N(Math.Floor(d));
        }, "DATEVALUE");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var s = c.Str(a[0]).Trim();
            if (!ValueParser.TryDateTime(s, out var d, out _)) throw Err(Errors.Value);
            return N(d - Math.Floor(d));
        }, "TIMEVALUE");
        Reg((c, a) =>
        {
            Args(a, 2, 3);
            var s = Dt(c, a[0]).Date; var e = Dt(c, a[1]).Date; int basis = (int)c.NumOr(a, 2, 0);
            if (s > e) (s, e) = (e, s);
            return basis switch
            {
                0 => N(Days360Us(s, e) / 360.0),
                1 => N(YearFracActual(s, e)),
                2 => N((e - s).TotalDays / 360.0),
                3 => N((e - s).TotalDays / 365.0),
                4 => N(Days360Eu(s, e) / 360.0),
                _ => throw Err(Errors.Num)
            };
        }, "YEARFRAC");
    }

    private static int YdDiff(DateTime s, DateTime e)
    {
        var adj = new DateTime(e.Year, s.Month, Math.Min(s.Day, DateTime.DaysInMonth(e.Year, s.Month)));
        if (adj > e) adj = new DateTime(e.Year - 1, s.Month, Math.Min(s.Day, DateTime.DaysInMonth(e.Year - 1, s.Month)));
        return (e - adj).Days;
    }

    private static int Days360Us(DateTime s, DateTime e)
    {
        int d1 = s.Day, d2 = e.Day;
        bool lastFebS = s.Month == 2 && d1 == DateTime.DaysInMonth(s.Year, 2);
        bool lastFebE = e.Month == 2 && d2 == DateTime.DaysInMonth(e.Year, 2);
        if (lastFebS && lastFebE) d2 = 30;
        if (lastFebS) d1 = 30;
        if (d2 == 31 && d1 >= 30) d2 = 30;
        if (d1 == 31) d1 = 30;
        return (e.Year - s.Year) * 360 + (e.Month - s.Month) * 30 + d2 - d1;
    }

    private static int Days360Eu(DateTime s, DateTime e)
    {
        int d1 = Math.Min(s.Day, 30), d2 = Math.Min(e.Day, 30);
        return (e.Year - s.Year) * 360 + (e.Month - s.Month) * 30 + d2 - d1;
    }

    private static double YearFracActual(DateTime s, DateTime e)
    {
        if (s.Year == e.Year) return (e - s).TotalDays / (DateTime.IsLeapYear(s.Year) ? 366.0 : 365.0);
        bool withinYear = e <= s.AddYears(1);
        if (withinYear)
        {
            bool leap = (DateTime.IsLeapYear(s.Year) && s <= new DateTime(s.Year, 2, 29)) || (DateTime.IsLeapYear(e.Year) && e >= new DateTime(e.Year, 2, 29));
            return (e - s).TotalDays / (leap ? 366.0 : 365.0);
        }
        int years = e.Year - s.Year + 1;
        double days = (new DateTime(e.Year + 1, 1, 1) - new DateTime(s.Year, 1, 1)).TotalDays;
        return (e - s).TotalDays / (days / years);
    }

    private static HashSet<DateTime> Holidays(Ctx c, Expr e)
    {
        var set = new HashSet<DateTime>();
        foreach (var v in FlatVal(c, c.Eval(e))) if (v.IsNumber) set.Add(DateSerial.FromSerial(v.Num).Date);
        return set;
    }

    // ================================================================== FINANCIAL

    internal static double Fv(double r, double n, double pmt, double pv, double type)
        => r == 0 ? -(pv + pmt * n) : -(pv * Math.Pow(1 + r, n) + pmt * (1 + r * type) * (Math.Pow(1 + r, n) - 1) / r);

    internal static double Pmt(double r, double n, double pv, double fv, double type)
    {
        if (n == 0) throw Err(Errors.Num);
        if (r == 0) return -(pv + fv) / n;
        double f = Math.Pow(1 + r, n);
        return -(r * (fv + pv * f)) / ((1 + r * type) * (f - 1));
    }

    private static void RegisterFinance()
    {
        Reg((c, a) => { Args(a, 3, 5); return N(Pmt(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0), c.NumOr(a, 4, 0))); }, "PMT");
        Reg((c, a) => { Args(a, 3, 5); return N(Fv(c.Num(a[0]), c.Num(a[1]), c.Num(a[2]), c.NumOr(a, 3, 0), c.NumOr(a, 4, 0))); }, "FV");
        Reg((c, a) =>
        {
            Args(a, 3, 5);
            double r = c.Num(a[0]), n = c.Num(a[1]), pmt = c.Num(a[2]), fv = c.NumOr(a, 3, 0), t = c.NumOr(a, 4, 0);
            if (r == 0) return N(-(fv + pmt * n));
            double f = Math.Pow(1 + r, n);
            return N(-(fv + pmt * (1 + r * t) * (f - 1) / r) / f);
        }, "PV");
        Reg((c, a) =>
        {
            Args(a, 3, 5);
            double r = c.Num(a[0]), pmt = c.Num(a[1]), pv = c.Num(a[2]), fv = c.NumOr(a, 3, 0), t = c.NumOr(a, 4, 0);
            if (r == 0) { if (pmt == 0) throw Err(Errors.Num); return N(-(pv + fv) / pmt); }
            double num = pmt * (1 + r * t) - fv * r, den = pmt * (1 + r * t) + pv * r;
            if (num / den <= 0) throw Err(Errors.Num);
            return N(Math.Log(num / den) / Math.Log(1 + r));
        }, "NPER");
        Reg((c, a) =>
        {
            Args(a, 2);
            double r = c.Num(a[0]);
            var vals = Nums(c, a.Skip(1));
            double s = 0;
            for (int i = 0; i < vals.Count; i++) s += vals[i] / Math.Pow(1 + r, i + 1);
            return N(s);
        }, "NPV");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            var vals = Nums(c, [a[0]]);
            double guess = c.NumOr(a, 1, 0.1);
            if (!vals.Any(v => v > 0) || !vals.Any(v => v < 0)) throw Err(Errors.Num);
            double r = guess;
            for (int it = 0; it < 100; it++)
            {
                double f = 0, df = 0;
                for (int i = 0; i < vals.Count; i++) { f += vals[i] / Math.Pow(1 + r, i); df -= i * vals[i] / Math.Pow(1 + r, i + 1); }
                if (df == 0) break;
                double nr = r - f / df;
                if (Math.Abs(nr - r) < 1e-10) return N(nr);
                r = nr;
                if (r <= -1) r = -0.99;
            }
            throw Err(Errors.Num);
        }, "IRR");
        Reg((c, a) =>
        {
            Args(a, 3, 6);
            double n = c.Num(a[0]), pmt = c.Num(a[1]), pv = c.Num(a[2]), fv = c.NumOr(a, 3, 0), t = c.NumOr(a, 4, 0), r = c.NumOr(a, 5, 0.1);
            for (int it = 0; it < 200; it++)
            {
                double f = RateF(r, n, pmt, pv, fv, t);
                double h = 1e-7;
                double df = (RateF(r + h, n, pmt, pv, fv, t) - f) / h;
                if (df == 0) break;
                double nr = r - f / df;
                if (Math.Abs(nr - r) < 1e-10) return N(nr);
                r = nr;
            }
            throw Err(Errors.Num);
        }, "RATE");
        Reg((c, a) =>
        {
            Args(a, 4, 6);
            double r = c.Num(a[0]), per = c.Num(a[1]), n = c.Num(a[2]), pv = c.Num(a[3]), fv = c.NumOr(a, 4, 0), t = c.NumOr(a, 5, 0);
            if (per < 1 || per > n) throw Err(Errors.Num);
            return N(Ipmt(r, per, n, pv, fv, t));
        }, "IPMT");
        Reg((c, a) =>
        {
            Args(a, 4, 6);
            double r = c.Num(a[0]), per = c.Num(a[1]), n = c.Num(a[2]), pv = c.Num(a[3]), fv = c.NumOr(a, 4, 0), t = c.NumOr(a, 5, 0);
            if (per < 1 || per > n) throw Err(Errors.Num);
            return N(Pmt(r, n, pv, fv, t) - Ipmt(r, per, n, pv, fv, t));
        }, "PPMT");
        Reg((c, a) => { Args(a, 3, 3); var life = c.Num(a[2]); if (life == 0) throw Err(Errors.Div0); return N((c.Num(a[0]) - c.Num(a[1])) / life); }, "SLN");
        Reg((c, a) =>
        {
            Args(a, 4, 4);
            double cost = c.Num(a[0]), salv = c.Num(a[1]), life = c.Num(a[2]), per = c.Num(a[3]);
            if (life <= 0 || per < 1 || per > life) throw Err(Errors.Num);
            return N((cost - salv) * (life - per + 1) * 2 / (life * (life + 1)));
        }, "SYD");
        Reg((c, a) =>
        {
            Args(a, 4, 5);
            double cost = c.Num(a[0]), salv = c.Num(a[1]), life = c.Num(a[2]), per = c.Num(a[3]), factor = c.NumOr(a, 4, 2);
            if (life <= 0 || per < 1 || per > life) throw Err(Errors.Num);
            double book = cost, dep = 0;
            for (int i = 1; i <= per; i++) { dep = Math.Min(book * factor / life, Math.Max(0, book - salv)); book -= dep; }
            return N(dep);
        }, "DDB");
    }

    private static double RateF(double r, double n, double pmt, double pv, double fv, double t)
    {
        if (Math.Abs(r) < 1e-12) return pv + pmt * n + fv;
        double f = Math.Pow(1 + r, n);
        return pv * f + pmt * (1 + r * t) * (f - 1) / r + fv;
    }

    private static double Ipmt(double r, double per, double n, double pv, double fv, double t)
    {
        double pmt = Pmt(r, n, pv, fv, t);
        if (per == 1 && t == 1) return 0;
        double ip = Fv(r, per - 1, pmt, pv, t) * r;
        if (t == 1) ip /= 1 + r;
        return ip;
    }

    // ================================================================== INFORMATION

    private static CellValue Probe(Ctx c, Expr e)
    {
        try { return c.Scalar(e); }
        catch (FormulaError fe) { return CellValue.Error(fe.Code); }
    }

    private static void RegisterInfo()
    {
        Reg((c, a) => { Args(a, 1, 1); return B(Probe(c, a[0]).IsBlank); }, "ISBLANK");
        Reg((c, a) => { Args(a, 1, 1); return B(Probe(c, a[0]).IsNumber); }, "ISNUMBER");
        Reg((c, a) => { Args(a, 1, 1); return B(Probe(c, a[0]).IsText); }, "ISTEXT");
        Reg((c, a) => { Args(a, 1, 1); return B(!Probe(c, a[0]).IsText); }, "ISNONTEXT");
        Reg((c, a) => { Args(a, 1, 1); return B(Probe(c, a[0]).IsBool); }, "ISLOGICAL");
        Reg((c, a) => { Args(a, 1, 1); return B(Probe(c, a[0]).IsError); }, "ISERROR");
        Reg((c, a) => { Args(a, 1, 1); var v = Probe(c, a[0]); return B(v.IsError && v.Str != Errors.NA); }, "ISERR");
        Reg((c, a) => { Args(a, 1, 1); var v = Probe(c, a[0]); return B(v.IsError && v.Str == Errors.NA); }, "ISNA");
        Reg((c, a) => { Args(a, 1, 1); return B(((long)Math.Floor(c.Num(a[0]))) % 2 == 0); }, "ISEVEN");
        Reg((c, a) => { Args(a, 1, 1); return B(((long)Math.Floor(c.Num(a[0]))) % 2 != 0); }, "ISODD");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            try { return B(c.Eval(a[0]) is RangeVal); } catch (FormulaError) { return B(false); }
        }, "ISREF");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            if (c.Eval(a[0]) is not RangeVal r) throw Err(Errors.Value);
            return B(r.Sheet.GetCell(r.Range.R1, r.Range.C1)?.Formula != null);
        }, "ISFORMULA");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            if (c.Eval(a[0]) is not RangeVal r) throw Err(Errors.NA);
            var f = r.Sheet.GetCell(r.Range.R1, r.Range.C1)?.Formula;
            return f == null ? throw Err(Errors.NA) : S("=" + f);
        }, "FORMULATEXT");
        Reg((c, a) => throw Err(Errors.NA), "NA");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var v = c.Eval(a[0]);
            if (v is ArrayVal or RangeVal { Range.IsSingle: false }) return N(64);
            var s = c.Calc.ToScalar(v, c);
            return N(s.Kind switch { ValueKind.Number or ValueKind.Blank => 1, ValueKind.Text => 2, ValueKind.Bool => 4, _ => 16 });
        }, "TYPE");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var v = c.Scalar(a[0]);
            return v.Kind switch { ValueKind.Number => v, ValueKind.Bool => N(v.Num), ValueKind.Error => throw Err(v.Str!), _ => N(0) };
        }, "N");
        Reg((c, a) =>
        {
            Args(a, 1, 1);
            var v = Probe(c, a[0]);
            if (!v.IsError) throw Err(Errors.NA);
            return N(v.Str switch { Errors.Null => 1, Errors.Div0 => 2, Errors.Value => 3, Errors.Ref => 4, Errors.Name => 5, Errors.Num => 6, Errors.NA => 7, Errors.Spill => 9, Errors.Calc => 14, _ => 8 });
        }, "ERROR.TYPE");
        Reg((c, a) =>
        {
            Args(a, 1, 2);
            var info = c.Str(a[0]).ToLowerInvariant();
            if (info == "address" && a.Length > 1 && c.Eval(a[1]) is RangeVal r) return S(Addr.ToA1(r.Range.R1, r.Range.C1, true, true));
            if (info == "row" && a.Length > 1 && c.Eval(a[1]) is RangeVal r2) return N(r2.Range.R1 + 1);
            if (info == "col" && a.Length > 1 && c.Eval(a[1]) is RangeVal r3) return N(r3.Range.C1 + 1);
            if (info == "contents" && a.Length > 1) return c.Scalar(a[1]);
            throw Err(Errors.Value);
        }, "CELL");
        Reg((c, a) => N(c.Book.Sheets.IndexOf(c.Sheet) + 1), "SHEET");
        Reg((c, a) => N(c.Book.Sheets.Count), "SHEETS");
    }
}

/// <summary>Reads a number aloud in Vietnamese: 120000 → "Một trăm hai mươi nghìn đồng".</summary>
public static class DocSo
{
    private static readonly string[] Digits = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
    private static readonly string[] Units = ["", "nghìn", "triệu", "tỷ"];

    public static string Read(double value, string unit = "đồng")
    {
        bool neg = value < 0;
        var n = (decimal)Math.Round(Math.Abs(value), 0, MidpointRounding.AwayFromZero);
        string words;
        if (n == 0) words = "không";
        else
        {
            var groups = new List<int>();
            while (n > 0) { groups.Add((int)(n % 1000)); n = Math.Floor(n / 1000); }
            var parts = new List<string>();
            for (int i = groups.Count - 1; i >= 0; i--)
            {
                int g = groups[i];
                if (g == 0) { if (i % 3 == 0 && i > 0 && parts.Count > 0 && groups.Skip(i).Take(3).Any(x => x != 0) && !parts[^1].EndsWith("tỷ")) parts.Add("tỷ"); continue; }
                bool full = i < groups.Count - 1; // not the leading group → read "không trăm", "linh"
                parts.Add(ReadGroup(g, full));
                string u = Scale(i);
                if (u.Length > 0) parts.Add(u);
            }
            words = string.Join(" ", parts).Replace("  ", " ").Trim();
        }
        if (neg) words = "âm " + words;
        if (!string.IsNullOrWhiteSpace(unit)) words += " " + unit.Trim();
        return char.ToUpper(words[0]) + words[1..];
    }

    private static string Scale(int groupIndex)
    {
        if (groupIndex == 0) return "";
        // 1 nghìn, 2 triệu, 3 tỷ, 4 nghìn tỷ, 5 triệu tỷ, 6 tỷ tỷ ...
        int k = groupIndex % 3;
        int billions = groupIndex / 3;
        var s = k == 0 ? "" : Units[k];
        for (int i = 0; i < billions; i++) s += (s.Length > 0 ? " " : "") + "tỷ";
        return s;
    }

    private static string ReadGroup(int g, bool full)
    {
        int h = g / 100, t = g / 10 % 10, u = g % 10;
        var w = new List<string>();
        if (h > 0 || full) { w.Add(Digits[h]); w.Add("trăm"); }
        if (t == 0)
        {
            if (u > 0 && (h > 0 || full)) w.Add("linh");
        }
        else if (t == 1) w.Add("mười");
        else { w.Add(Digits[t]); w.Add("mươi"); }
        if (u > 0)
        {
            if (u == 1 && t >= 2) w.Add("mốt");
            else if (u == 5 && t >= 1) w.Add("lăm");
            else w.Add(Digits[u]);
        }
        return string.Join(" ", w);
    }
}
