using System.Globalization;
using System.Text.RegularExpressions;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;

namespace MicroadsOffice.Excel.Model;

public enum PasteMode { All, Values, Formulas, Formats, ValuesAndFormats }

/// <summary>Content copied with Ctrl+C / Ctrl+X (kept independent of later edits).</summary>
public sealed class ClipData
{
    public string SheetName = "";
    public CellRange Range;
    public Dictionary<(int, int), Cell> Cells = new();
    public Dictionary<(int, int), int> StyleIds = new();
    public List<CellRange> Merges = new();
    public StyleTable Styles = new();
    public bool IsCut;
    public Dictionary<int, double> ColWidths = new();
    public Workbook? SourceBook;
    /// <summary>Offsets of values spilled by a formula that is copied too: pasting formulas leaves them empty so the formula can spill again.</summary>
    public HashSet<(int, int)> SpillChildren = new();
}

/// <summary>Structural and data operations on sheets (all keep formulas consistent).</summary>
public static class SheetOps
{
    // ------------------------------------------------------------------ insert / delete rows & columns

    public static void InsertRows(Sheet s, int at, int count) => Shift(s, true, at, count);
    public static void DeleteRows(Sheet s, int at, int count) => Shift(s, true, at, -count);
    public static void InsertCols(Sheet s, int at, int count) => Shift(s, false, at, count);
    public static void DeleteCols(Sheet s, int at, int count) => Shift(s, false, at, -count);

    private static int? MoveIndex(int i, int at, int count)
    {
        if (count > 0) return i >= at ? i + count : i;
        int del = -count;
        if (i >= at && i < at + del) return null;
        return i >= at + del ? i - del : i;
    }

    public static CellRange? AdjustRange(CellRange r, bool isRow, int at, int count)
    {
        int a1 = isRow ? r.R1 : r.C1, a2 = isRow ? r.R2 : r.C2;
        int max = isRow ? Addr.MaxRows - 1 : Addr.MaxCols - 1;
        int n1, n2;
        if (count > 0)
        {
            n1 = a1 >= at ? a1 + count : a1;
            n2 = a2 >= at ? a2 + count : a2;
            if (a2 == max) n2 = max; // whole column/row stays whole
            if (n1 > max) return null;
            n2 = Math.Min(n2, max);
        }
        else
        {
            int del = -count, end = at + del - 1;
            if (a1 >= at && a2 <= end) return null;
            n1 = a1 > end ? a1 - del : a1 >= at ? at : a1;
            n2 = a2 > end ? a2 - del : a2 >= at ? at - 1 : a2;
            if (n2 < n1) return null;
        }
        return isRow ? r with { R1 = n1, R2 = n2 } : r with { C1 = n1, C2 = n2 };
    }

    private static Dictionary<int, T> ShiftKeys<T>(Dictionary<int, T> d, int at, int count)
    {
        var res = new Dictionary<int, T>();
        foreach (var (k, v) in d) { var nk = MoveIndex(k, at, count); if (nk != null) res[nk.Value] = v; }
        return res;
    }

    private static HashSet<int> ShiftSet(HashSet<int> d, int at, int count)
    {
        var res = new HashSet<int>();
        foreach (var k in d) { var nk = MoveIndex(k, at, count); if (nk != null) res.Add(nk.Value); }
        return res;
    }

    private static void Shift(Sheet s, bool isRow, int at, int count)
    {
        var wb = s.Book;
        // 1. cells
        var cells = new Dictionary<long, Cell>(s.Cells.Count);
        foreach (var (k, cell) in s.Cells)
        {
            var (r, c) = Sheet.FromKey(k);
            var n = MoveIndex(isRow ? r : c, at, count);
            if (n == null) continue;
            int nr = isRow ? n.Value : r, nc = isRow ? c : n.Value;
            if (nr >= Addr.MaxRows || nc >= Addr.MaxCols) continue;
            cells[Sheet.Key(nr, nc)] = cell;
        }
        s.Cells = cells;
        Dictionary<long, string> ShiftPos(Dictionary<long, string> d)
        {
            var res = new Dictionary<long, string>();
            foreach (var (k, v) in d)
            {
                var (r, c) = Sheet.FromKey(k);
                var n = MoveIndex(isRow ? r : c, at, count);
                if (n != null) res[isRow ? Sheet.Key(n.Value, c) : Sheet.Key(r, n.Value)] = v;
            }
            return res;
        }
        s.Notes = ShiftPos(s.Notes);
        s.Links = ShiftPos(s.Links);
        if (isRow)
        {
            s.RowHeights = ShiftKeys(s.RowHeights, at, count);
            s.AutoHeights = ShiftKeys(s.AutoHeights, at, count);
            s.RowStyles = ShiftKeys(s.RowStyles, at, count);
            s.HiddenRows = ShiftSet(s.HiddenRows, at, count);
            s.FilteredRows = ShiftSet(s.FilteredRows, at, count);
        }
        else
        {
            s.ColWidths = ShiftKeys(s.ColWidths, at, count);
            s.ColStyles = ShiftKeys(s.ColStyles, at, count);
            s.HiddenCols = ShiftSet(s.HiddenCols, at, count);
        }
        s.Merges = s.Merges.Select(m => AdjustRange(m, isRow, at, count)).Where(m => m != null && !m.Value.IsSingle).Select(m => m!.Value).ToList();
        foreach (var cf in s.CondFormats.ToList()) { var r = AdjustRange(cf.Range, isRow, at, count); if (r == null) s.CondFormats.Remove(cf); else cf.Range = r.Value; }
        foreach (var dv in s.Validations.ToList()) { var r = AdjustRange(dv.Range, isRow, at, count); if (r == null) s.Validations.Remove(dv); else dv.Range = r.Value; }
        foreach (var ch in s.Charts) { var r = AdjustRange(ch.Data, isRow, at, count); if (r != null) ch.Data = r.Value; }
        if (s.Filter != null)
        {
            var r = AdjustRange(s.Filter.Range, isRow, at, count);
            if (r == null) s.Filter = null;
            else
            {
                s.Filter.Range = r.Value;
                if (!isRow)
                {
                    s.Filter.Allowed = ShiftKeys(s.Filter.Allowed, at, count);
                    s.Filter.Conditions = ShiftKeys(s.Filter.Conditions, at, count);
                }
            }
        }
        if (s.Page.PrintArea is CellRange pa) s.Page.PrintArea = AdjustRange(pa, isRow, at, count);
        // 2. formulas everywhere
        foreach (var sheet in wb.Sheets)
            foreach (var cell in sheet.Cells.Values)
                if (cell.Formula != null)
                {
                    var nf = RefAdjust.InsertDelete(cell.Formula, sheet.Name, s.Name, isRow, at, count);
                    if (!ReferenceEquals(nf, cell.Formula)) { cell.Formula = nf; cell.Ast = null; }
                }
        foreach (var name in wb.Names.Keys.ToList())
            wb.Names[name] = RefAdjust.InsertDelete(wb.Names[name], s.Name, s.Name, isRow, at, count);
        wb.Calc.InvalidateNames();
        wb.MarkDirty();
    }

    // ------------------------------------------------------------------ copy / paste

    public static ClipData Copy(Sheet s, CellRange r, bool cut)
    {
        var d = new ClipData { SheetName = s.Name, Range = r, IsCut = cut, SourceBook = s.Book };
        var used = s.UsedRange();
        var clipR = new CellRange(r.R1, r.C1, Math.Min(r.R2, Math.Max(used.R2, r.R1)), Math.Min(r.C2, Math.Max(used.C2, r.C1)));
        d.Range = r.Count > 2_000_000 ? clipR : r;
        foreach (var (k, cell) in s.Cells)
        {
            var (row, col) = Sheet.FromKey(k);
            if (!d.Range.Contains(row, col)) continue;
            var copy = cell.Clone();
            copy.StyleId = d.Styles.Intern(s.Book.Styles[cell.StyleId]);
            d.Cells[(row - d.Range.R1, col - d.Range.C1)] = copy;
            if (s.SpillSource(row, col) is var (mr, mc, _) && d.Range.Contains(mr, mc)) d.SpillChildren.Add((row - d.Range.R1, col - d.Range.C1));
        }
        d.Merges = s.Merges.Where(m => d.Range.Contains(m)).Select(m => new CellRange(m.R1 - d.Range.R1, m.C1 - d.Range.C1, m.R2 - d.Range.R1, m.C2 - d.Range.C1)).ToList();
        for (int c = d.Range.C1; c <= d.Range.C2 && c < d.Range.C1 + 200; c++) if (s.ColWidths.TryGetValue(c, out var w)) d.ColWidths[c - d.Range.C1] = w;
        return d;
    }

    /// <summary>Pastes at <paramref name="target"/>'s top-left (tiled when the target is a multiple of the source). Returns the pasted area.</summary>
    public static CellRange Paste(Sheet s, ClipData d, CellRange target, PasteMode mode = PasteMode.All, bool transpose = false, bool skipBlanks = false)
    {
        int srcRows = transpose ? d.Range.Cols : d.Range.Rows, srcCols = transpose ? d.Range.Rows : d.Range.Cols;
        int tileR = 1, tileC = 1;
        if (!d.IsCut && target.Rows % srcRows == 0 && target.Cols % srcCols == 0 && (target.Rows > srcRows || target.Cols > srcCols) && target.Count < 1_000_000)
        { tileR = target.Rows / srcRows; tileC = target.Cols / srcCols; }
        var srcSheet = d.SourceBook?.Sheet(d.SheetName);
        if (d.IsCut && srcSheet != null && srcSheet.Book == s.Book)
        {
            // clear source first (contents + formats), remembering we move
            foreach (var (row, col) in d.Range.Cells().Take(5_000_000).ToList()) srcSheet.Cells.Remove(Sheet.Key(row, col));
            srcSheet.Merges.RemoveAll(m => d.Range.Contains(m));
        }
        for (int tr = 0; tr < tileR; tr++)
            for (int tc = 0; tc < tileC; tc++)
            {
                int baseR = target.R1 + tr * srcRows, baseC = target.C1 + tc * srcCols;
                // clear destination for full paste
                if (mode == PasteMode.All && !skipBlanks)
                    for (int r = 0; r < srcRows; r++) for (int c = 0; c < srcCols; c++) s.Cells.Remove(Sheet.Key(baseR + r, baseC + c));
                foreach (var ((dr, dc), src) in d.Cells)
                {
                    int r = transpose ? dc : dr, c = transpose ? dr : dc;
                    int row = baseR + r, col = baseC + c;
                    if (row >= Addr.MaxRows || col >= Addr.MaxCols) continue;
                    if (skipBlanks && src.Value.IsBlank && src.Formula == null) continue;
                    var cell = s.GetOrCreate(row, col);
                    int offR = row - (d.Range.R1 + dr), offC = col - (d.Range.C1 + dc);
                    if (mode is not PasteMode.Formats) s.ForgetSpill(row, col, cell);
                    if (mode is PasteMode.All or PasteMode.Formulas && d.SpillChildren.Contains((dr, dc)))
                    {
                        cell.Formula = null; cell.Ast = null; cell.Value = CellValue.Blank;
                    }
                    else if (mode is PasteMode.All or PasteMode.Formulas)
                    {
                        if (src.Formula != null)
                        {
                            cell.Formula = d.IsCut ? src.Formula : RefAdjust.Shift(src.Formula, offR, offC);
                            cell.Ast = null; cell.Value = src.Value;
                            cell.Implicit = src.Implicit; // a copied formula from an older file keeps its "@" behaviour
                        }
                        else { cell.Formula = null; cell.Ast = null; cell.Value = src.Value; }
                    }
                    else if (mode is PasteMode.Values or PasteMode.ValuesAndFormats)
                    {
                        cell.Formula = null; cell.Ast = null; cell.Value = src.Value;
                    }
                    if (mode is PasteMode.All or PasteMode.Formats or PasteMode.ValuesAndFormats)
                        cell.StyleId = s.Book.Styles.Intern(d.Styles[src.StyleId]);
                    if (cell.IsEmpty) s.Cells.Remove(Sheet.Key(row, col));
                }
                if (mode is PasteMode.All or PasteMode.Formats)
                    foreach (var m in d.Merges)
                    {
                        var mr = transpose ? new CellRange(m.C1, m.R1, m.C2, m.R2) : m;
                        var nm = new CellRange(baseR + mr.R1, baseC + mr.C1, baseR + mr.R2, baseC + mr.C2);
                        s.Merges.RemoveAll(x => x.Intersects(nm));
                        s.Merges.Add(nm);
                    }
            }
        if (d.IsCut && srcSheet != null && srcSheet.Book == s.Book)
        {
            int dr = target.R1 - d.Range.R1, dc = target.C1 - d.Range.C1;
            // references to the moved block follow it (only when moved within the same sheet)
            if (srcSheet == s)
                foreach (var sheet in s.Book.Sheets)
                    foreach (var cell in sheet.Cells.Values)
                        if (cell.Formula != null)
                        {
                            var nf = RefAdjust.MoveRefs(cell.Formula, sheet.Name, s.Name, d.Range, dr, dc);
                            if (!ReferenceEquals(nf, cell.Formula)) { cell.Formula = nf; cell.Ast = null; }
                        }
        }
        s.Book.MarkDirty();
        return new CellRange(target.R1, target.C1, target.R1 + srcRows * tileR - 1, target.C1 + srcCols * tileC - 1);
    }

    /// <summary>Tab-separated display text of a range (for the Windows clipboard).</summary>
    public static string ToTsv(Sheet s, CellRange r)
    {
        var used = s.UsedRange();
        int r2 = Math.Min(r.R2, used.R2), c2 = Math.Min(r.C2, used.C2);
        var sb = new System.Text.StringBuilder();
        for (int row = r.R1; row <= r2; row++)
        {
            for (int col = r.C1; col <= c2; col++)
            {
                if (col > r.C1) sb.Append('\t');
                var cell = s.GetCell(row, col);
                if (cell == null) continue;
                var text = NumFmt.Format(cell.Value, s.Book.Styles[cell.StyleId].NumberFormat, out _);
                if (text.Contains('\t') || text.Contains('\n') || text.Contains('"')) text = "\"" + text.Replace("\"", "\"\"") + "\"";
                sb.Append(text);
            }
            sb.Append("\r\n");
        }
        return sb.ToString();
    }

    /// <summary>Pastes plain text (from other programs) starting at (row, col).</summary>
    public static CellRange PasteText(Sheet s, string text, int row, int col)
    {
        var rows = CsvIO.ParseDelimited(text, text.Contains('\t') ? '\t' : CsvIO.DetectDelimiter(text));
        int maxC = 0;
        for (int i = 0; i < rows.Count; i++)
        {
            for (int j = 0; j < rows[i].Count; j++)
            {
                SetInput(s, row + i, col + j, rows[i][j]);
                maxC = Math.Max(maxC, j);
            }
        }
        s.Book.MarkDirty();
        return new CellRange(row, col, row + Math.Max(0, rows.Count - 1), col + maxC);
    }

    /// <summary>Enters what a user typed: "=..." formula, number/date/percent (auto format), or text.</summary>
    public static void SetInput(Sheet s, int row, int col, string input)
    {
        if (input.StartsWith('=') && input.Length > 1)
        {
            s.SetFormula(row, col, input[1..]);
            return;
        }
        var style = s.StyleAt(row, col);
        if (style.NumberFormat == "@") { s.SetValue(row, col, input.Length == 0 ? CellValue.Blank : CellValue.Text(input)); return; }
        var v = ValueParser.Parse(input, out var fmt);
        s.SetValue(row, col, v);
        if (fmt != null && style.NumberFormat == "General")
        {
            var id = s.Book.Restyle(s.StyleIdAt(row, col), st => st with { NumberFormat = fmt });
            s.SetStyle(row, col, id);
        }
    }

    // ------------------------------------------------------------------ fill

    private static readonly string[][] Lists =
    [
        ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
        ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
        ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"],
        ["T2", "T3", "T4", "T5", "T6", "T7", "CN"],
        ["Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5", "Tháng 6", "Tháng 7", "Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12"],
        ["Quý 1", "Quý 2", "Quý 3", "Quý 4"],
        ["Q1", "Q2", "Q3", "Q4"]
    ];

    /// <summary>
    /// Fills <paramref name="target"/> (which extends <paramref name="source"/> down/right/up/left) with series
    /// detected from the source (numbers, dates, weekday/month names, "Item 1"), or copies (formulas shift).
    /// </summary>
    public static void Fill(Sheet s, CellRange source, CellRange target, bool copyOnly = false)
    {
        bool vertical = target.C1 == source.C1 && target.C2 == source.C2;
        bool forward = vertical ? target.R2 > source.R2 : target.C2 > source.C2;
        int lines = vertical ? source.Cols : source.Rows;
        int srcLen = vertical ? source.Rows : source.Cols;
        int total = vertical ? target.Rows : target.Cols;
        for (int line = 0; line < lines; line++)
        {
            (int R, int C) SrcPos(int i) => vertical ? (forward ? source.R1 + i : source.R2 - i, source.C1 + line) : (source.R1 + line, forward ? source.C1 + i : source.C2 - i);
            (int R, int C) DstPos(int i) => vertical ? (forward ? source.R1 + i : source.R2 - i, source.C1 + line) : (source.R1 + line, forward ? source.C1 + i : source.C2 - i);
            var cells = Enumerable.Range(0, srcLen).Select(i => { var p = SrcPos(i); return s.GetCell(p.R, p.C); }).ToList();
            var series = copyOnly ? null : DetectSeries(s, cells, srcLen);
            for (int i = srcLen; i < total; i++)
            {
                var (dr, dc) = DstPos(i);
                var srcIdx = i % srcLen;
                var (sr, sc) = SrcPos(srcIdx);
                var src = s.GetCell(sr, sc);
                var cell = s.GetOrCreate(dr, dc);
                cell.StyleId = src?.StyleId ?? 0;
                if (series != null)
                {
                    cell.Formula = null; cell.Ast = null;
                    cell.Value = series(i);
                }
                else if (src?.Formula != null)
                {
                    cell.Formula = RefAdjust.Shift(src.Formula, dr - sr, dc - sc);
                    cell.Ast = null;
                }
                else { cell.Formula = null; cell.Ast = null; cell.Value = src?.Value ?? CellValue.Blank; }
                if (cell.IsEmpty) s.Cells.Remove(Sheet.Key(dr, dc));
            }
        }
        s.Book.MarkDirty();
    }

    private static Func<int, CellValue>? DetectSeries(Sheet s, List<Cell?> cells, int n)
    {
        if (cells.Any(c => c?.Formula != null)) return null;
        var vals = cells.Select(c => c?.Value ?? CellValue.Blank).ToList();
        if (vals.All(v => v.IsNumber))
        {
            if (n == 1)
            {
                // a single date or number: dates step by one day, plain numbers are copied (like Excel without Ctrl)
                var fmt = s.Book.Styles[cells[0]!.StyleId].NumberFormat;
                if (NumFmt.IsDateFormat(fmt)) { double d0 = vals[0].Num; return i => CellValue.Number(d0 + i); }
                return null;
            }
            var fmt0 = s.Book.Styles[cells[0]!.StyleId].NumberFormat;
            if (NumFmt.IsDateFormat(fmt0))
            {
                var dates = vals.Select(v => DateSerial.FromSerial(v.Num)).ToList();
                bool monthly = dates.All(d => d.Day == dates[0].Day) && dates.Select(d => d.Month + d.Year * 12).Distinct().Count() == n;
                if (monthly)
                {
                    int step = (dates[1].Year * 12 + dates[1].Month) - (dates[0].Year * 12 + dates[0].Month);
                    var d0 = dates[0];
                    return i => CellValue.Number(DateSerial.ToSerial(d0.AddMonths(step * i)));
                }
            }
            double first = vals[0].Num, stepN = (vals[^1].Num - first) / (n - 1);
            return i => CellValue.Number(Math.Round(first + stepN * i, 10));
        }
        if (vals.All(v => v.IsText))
        {
            foreach (var list in Lists)
            {
                var idx = vals.Select(v => Array.FindIndex(list, x => x.Equals(v.Str, StringComparison.OrdinalIgnoreCase))).ToList();
                if (idx.Any(x => x < 0)) continue;
                int step = n > 1 ? ((idx[1] - idx[0]) % list.Length + list.Length) % list.Length : 1;
                if (step == 0) step = 1;
                bool upper = vals[0].Str!.All(ch => !char.IsLetter(ch) || char.IsUpper(ch)) && vals[0].Str!.Length > 3;
                int i0 = idx[0];
                return i =>
                {
                    var t = list[((i0 + step * i) % list.Length + list.Length) % list.Length];
                    return CellValue.Text(upper ? t.ToUpperInvariant() : t);
                };
            }
            var rx = new Regex(@"^(.*?)(\d+)(\D*)$");
            var ms = vals.Select(v => rx.Match(v.Str!)).ToList();
            if (ms.All(m => m.Success) && ms.All(m => m.Groups[1].Value == ms[0].Groups[1].Value && m.Groups[3].Value == ms[0].Groups[3].Value))
            {
                var nums = ms.Select(m => long.Parse(m.Groups[2].Value)).ToList();
                long step = n > 1 ? nums[1] - nums[0] : 1;
                string pre = ms[0].Groups[1].Value, post = ms[0].Groups[3].Value;
                int width = ms[0].Groups[2].Value.Length;
                bool pad = ms[0].Groups[2].Value.StartsWith('0') && width > 1;
                long start = nums[0];
                return i => { var v = start + step * i; return CellValue.Text(pre + (pad ? v.ToString(new string('0', width)) : v.ToString()) + post); };
            }
        }
        return null;
    }

    // ------------------------------------------------------------------ sort

    public sealed record SortKey(int Col, bool Ascending = true);

    /// <summary>Sorts the rows of <paramref name="r"/> (header row excluded when <paramref name="hasHeader"/>).</summary>
    /// <summary>The range contains a dynamic-array formula or its spilled values (Excel refuses to sort those).</summary>
    public static bool TouchesSpill(Sheet s, CellRange r)
    {
        foreach (var (k, cell) in s.Cells)
        {
            if (cell.SpillRows <= 0 && cell.SpillFrom == null) continue;
            var (row, col) = Sheet.FromKey(k);
            if (r.Contains(row, col)) return true;
        }
        return false;
    }

    public static string SpillSortMessage => MicroadsOffice.Core.Office.Loc.T(
        "Không sắp xếp được: vùng này chứa kết quả tràn của công thức mảng động (FILTER, SORT, UNIQUE…). Hãy sắp xếp dữ liệu nguồn hoặc dùng hàm SORT.",
        "Can't sort: the range contains the spilled result of a dynamic-array formula (FILTER, SORT, UNIQUE…). Sort the source data or use the SORT function.");

    public static void Sort(Sheet s, CellRange r, IReadOnlyList<SortKey> keys, bool hasHeader)
    {
        if (TouchesSpill(s, r)) throw new InvalidOperationException(SpillSortMessage);
        int start = hasHeader ? r.R1 + 1 : r.R1;
        if (start > r.R2) return;
        var used = s.UsedRange();
        int end = Math.Min(r.R2, used.R2);
        var rows = new List<(int Row, Dictionary<int, Cell> Cells)>();
        for (int row = start; row <= end; row++)
        {
            var d = new Dictionary<int, Cell>();
            for (int c = r.C1; c <= Math.Min(r.C2, used.C2); c++) { var cell = s.GetCell(row, c); if (cell != null) d[c] = cell; }
            rows.Add((row, d));
        }
        int Cmp((int Row, Dictionary<int, Cell> Cells) a, (int Row, Dictionary<int, Cell> Cells) b)
        {
            foreach (var k in keys)
            {
                var va = a.Cells.TryGetValue(k.Col, out var ca) ? ca.Value : CellValue.Blank;
                var vb = b.Cells.TryGetValue(k.Col, out var cb) ? cb.Value : CellValue.Blank;
                // blanks always last
                if (va.IsBlank && vb.IsBlank) continue;
                if (va.IsBlank) return 1;
                if (vb.IsBlank) return -1;
                int c = SortCompare(va, vb);
                if (c != 0) return k.Ascending ? c : -c;
            }
            return a.Row.CompareTo(b.Row);
        }
        var sorted = rows.ToList();
        sorted.Sort(Cmp);
        for (int i = 0; i < sorted.Count; i++)
        {
            int dest = start + i;
            for (int c = r.C1; c <= Math.Min(r.C2, used.C2); c++) s.Cells.Remove(Sheet.Key(dest, c));
        }
        for (int i = 0; i < sorted.Count; i++)
        {
            int dest = start + i;
            foreach (var (c, cell) in sorted[i].Cells)
            {
                if (cell.Formula != null && dest != sorted[i].Row) { cell.Formula = RefAdjust.Shift(cell.Formula, dest - sorted[i].Row, 0); cell.Ast = null; }
                s.Cells[Sheet.Key(dest, c)] = cell;
            }
        }
        s.Book.MarkDirty();
    }

    private static int SortCompare(CellValue a, CellValue b)
    {
        // numbers < text < logical < errors
        int Rank(CellValue v) => v.Kind switch { ValueKind.Number => 0, ValueKind.Text => 1, ValueKind.Bool => 2, _ => 3 };
        int ra = Rank(a), rb = Rank(b);
        if (ra != rb) return ra.CompareTo(rb);
        if (a.IsText) return string.Compare(a.Str, b.Str, CultureInfo.GetCultureInfo("vi-VN"), CompareOptions.IgnoreCase);
        if (a.IsError) return string.CompareOrdinal(a.Str, b.Str);
        return a.Num.CompareTo(b.Num);
    }

    /// <summary>Contiguous data block around a cell (like Ctrl+A / automatic sort range).</summary>
    public static CellRange CurrentRegion(Sheet s, int row, int col)
    {
        var r = new CellRange(row, col, row, col);
        bool NonEmpty(int rr, int cc) => rr >= 0 && cc >= 0 && rr < Addr.MaxRows && cc < Addr.MaxCols && s.GetCell(rr, cc) is { } c && (!c.Value.IsBlank || c.Formula != null);
        bool grown = true;
        int guard = 0;
        while (grown && guard++ < 10000)
        {
            grown = false;
            for (int c = r.C1 - 1; c <= r.C2 + 1; c++)
            {
                if (NonEmpty(r.R1 - 1, c)) { r = r with { R1 = r.R1 - 1 }; grown = true; break; }
                if (NonEmpty(r.R2 + 1, c)) { r = r with { R2 = r.R2 + 1 }; grown = true; break; }
            }
            for (int rr = r.R1 - 1; rr <= r.R2 + 1; rr++)
            {
                if (NonEmpty(rr, r.C1 - 1)) { r = r with { C1 = r.C1 - 1 }; grown = true; break; }
                if (NonEmpty(rr, r.C2 + 1)) { r = r with { C2 = r.C2 + 1 }; grown = true; break; }
            }
            r = new CellRange(Math.Max(0, r.R1), Math.Max(0, r.C1), r.R2, r.C2);
        }
        return r;
    }

    // ------------------------------------------------------------------ filter

    public static string DisplayText(Sheet s, int r, int c)
    {
        var cell = s.GetCell(r, c);
        if (cell == null) return "";
        return NumFmt.Format(cell.Value, s.Book.Styles[cell.StyleId].NumberFormat, out _);
    }

    /// <summary>Recomputes the rows hidden by the autofilter.</summary>
    public static void ApplyFilter(Sheet s)
    {
        s.FilteredRows.Clear();
        var f = s.Filter;
        if (f == null) return;
        var used = s.UsedRange();
        for (int row = f.Range.R1 + 1; row <= Math.Max(f.Range.R2, used.R2); row++)
        {
            bool visible = true;
            foreach (var (col, allowed) in f.Allowed)
                if (!allowed.Contains(DisplayText(s, row, col))) { visible = false; break; }
            if (visible)
                foreach (var (col, cond) in f.Conditions)
                    if (!MatchCondition(s.GetValue(row, col), cond)) { visible = false; break; }
            if (!visible) s.FilteredRows.Add(row);
        }
        if (used.R2 > f.Range.R2) f.Range = f.Range with { R2 = used.R2 };
    }

    /// <summary>Condition syntax: "&gt;100", "&lt;=5", "=abc", "contains:x", "begins:x", "ends:x", "top:10", "between:1|9".</summary>
    public static bool MatchCondition(CellValue v, string cond)
    {
        if (cond.StartsWith("contains:")) return Calc.ToTextSafe(v).Contains(cond[9..], StringComparison.CurrentCultureIgnoreCase);
        if (cond.StartsWith("notcontains:")) return !Calc.ToTextSafe(v).Contains(cond[12..], StringComparison.CurrentCultureIgnoreCase);
        if (cond.StartsWith("begins:")) return Calc.ToTextSafe(v).StartsWith(cond[7..], StringComparison.CurrentCultureIgnoreCase);
        if (cond.StartsWith("ends:")) return Calc.ToTextSafe(v).EndsWith(cond[5..], StringComparison.CurrentCultureIgnoreCase);
        if (cond.StartsWith("between:"))
        {
            var p = cond[8..].Split('|');
            if (p.Length == 2 && v.IsNumber && double.TryParse(p[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var lo) && double.TryParse(p[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var hi))
                return v.Num >= Math.Min(lo, hi) && v.Num <= Math.Max(lo, hi);
            return false;
        }
        return Functions.MakeCriterion(CellValue.Text(cond)).Match(v);
    }

    // ------------------------------------------------------------------ remove duplicates / text to columns

    /// <summary>Removes duplicate rows inside <paramref name="r"/> comparing <paramref name="cols"/>. Returns the number removed.</summary>
    public static int RemoveDuplicates(Sheet s, CellRange r, IReadOnlyList<int> cols, bool hasHeader)
    {
        var used = s.UsedRange();
        int start = hasHeader ? r.R1 + 1 : r.R1, end = Math.Min(r.R2, used.R2);
        var seen = new HashSet<string>();
        var keep = new List<int>();
        for (int row = start; row <= end; row++)
        {
            var key = string.Join("\u0001", cols.Select(c => { var v = s.GetValue(row, c); return v.Kind + ":" + v.ToString().ToUpperInvariant(); }));
            if (seen.Add(key)) keep.Add(row);
        }
        int removed = end - start + 1 - keep.Count;
        if (removed == 0) return 0;
        var data = keep.Select(row => Enumerable.Range(r.C1, Math.Min(r.C2, used.C2) - r.C1 + 1).Select(c => (c, s.GetCell(row, c))).ToList()).ToList();
        for (int row = start; row <= end; row++) for (int c = r.C1; c <= Math.Min(r.C2, used.C2); c++) s.Cells.Remove(Sheet.Key(row, c));
        for (int i = 0; i < data.Count; i++)
            foreach (var (c, cell) in data[i])
                if (cell != null) s.Cells[Sheet.Key(start + i, c)] = cell;
        s.Book.MarkDirty();
        return removed;
    }

    /// <summary>Splits the text of column <paramref name="col"/> (rows r1..r2) by <paramref name="delimiter"/> into neighbouring columns.</summary>
    public static int TextToColumns(Sheet s, int col, int r1, int r2, char delimiter, bool treatConsecutiveAsOne)
    {
        int maxParts = 0;
        for (int row = r1; row <= r2; row++)
        {
            var v = s.GetValue(row, col);
            if (!v.IsText) continue;
            var parts = v.Str!.Split(delimiter, treatConsecutiveAsOne ? StringSplitOptions.RemoveEmptyEntries : StringSplitOptions.None);
            for (int i = 0; i < parts.Length; i++) SetInput(s, row, col + i, parts[i].Trim());
            maxParts = Math.Max(maxParts, parts.Length);
        }
        s.Book.MarkDirty();
        return maxParts;
    }

    // ------------------------------------------------------------------ find / replace

    public sealed record FindOptions(string Text, bool MatchCase = false, bool WholeCell = false, bool InFormulas = true, bool WholeWorkbook = false);

    public static IEnumerable<(Sheet Sheet, int Row, int Col)> FindAll(Workbook wb, Sheet current, FindOptions o)
    {
        var sheets = o.WholeWorkbook ? wb.Sheets : [current];
        var cmp = o.MatchCase ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;
        var rx = o.Text.IndexOfAny(['*', '?']) >= 0 ? Functions.WildcardRegex(o.WholeCell ? o.Text : "*" + o.Text + "*") : null;
        foreach (var sh in sheets)
        {
            var keys = sh.Cells.Keys.ToList();
            keys.Sort();
            foreach (var k in keys)
            {
                var cell = sh.Cells[k];
                var text = o.InFormulas && cell.Formula != null ? "=" + cell.Formula : NumFmt.Format(cell.Value, wb.Styles[cell.StyleId].NumberFormat, out _);
                if (text.Length == 0) continue;
                bool hit = rx != null ? rx.IsMatch(text) : o.WholeCell ? string.Equals(text, o.Text, cmp) : text.Contains(o.Text, cmp);
                if (hit) { var (r, c) = Sheet.FromKey(k); yield return (sh, r, c); }
            }
        }
    }

    public static int ReplaceAll(Workbook wb, Sheet current, FindOptions o, string replacement)
    {
        int n = 0;
        foreach (var (sh, r, c) in FindAll(wb, current, o).ToList())
            if (ReplaceIn(sh, r, c, o, replacement)) n++;
        if (n > 0) wb.MarkDirty();
        return n;
    }

    public static bool ReplaceIn(Sheet sh, int r, int c, FindOptions o, string replacement)
    {
        var cell = sh.GetCell(r, c);
        if (cell == null) return false;
        var cmp = o.MatchCase ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;
        string src = cell.Formula != null ? "=" + cell.Formula : cell.Value.ToString();
        string res = o.WholeCell ? replacement : ReplaceStr(src, o.Text, replacement, cmp);
        if (res == src) return false;
        SetInput(sh, r, c, res);
        return true;
    }

    private static string ReplaceStr(string s, string find, string rep, StringComparison cmp)
    {
        if (find.Length == 0) return s;
        var sb = new System.Text.StringBuilder();
        int i = 0;
        while (true)
        {
            int j = s.IndexOf(find, i, cmp);
            if (j < 0) { sb.Append(s, i, s.Length - i); break; }
            sb.Append(s, i, j - i).Append(rep);
            i = j + find.Length;
        }
        return sb.ToString();
    }

    // ------------------------------------------------------------------ sheets

    public static void RenameSheet(Workbook wb, Sheet s, string newName)
    {
        var old = s.Name;
        s.Name = newName;
        foreach (var sheet in wb.Sheets)
            foreach (var cell in sheet.Cells.Values)
                if (cell.Formula != null)
                {
                    var nf = RefAdjust.RenameSheet(cell.Formula, old, newName);
                    if (!ReferenceEquals(nf, cell.Formula)) { cell.Formula = nf; cell.Ast = null; }
                }
        foreach (var k in wb.Names.Keys.ToList()) wb.Names[k] = RefAdjust.RenameSheet(wb.Names[k], old, newName);
        wb.Calc.InvalidateNames();
        wb.MarkDirty();
    }

    public static void DeleteSheet(Workbook wb, Sheet s)
    {
        var name = s.Name;
        wb.Sheets.Remove(s);
        foreach (var sheet in wb.Sheets)
            foreach (var cell in sheet.Cells.Values)
                if (cell.Formula != null)
                {
                    var nf = RefAdjust.RenameSheet(cell.Formula, name, null);
                    if (!ReferenceEquals(nf, cell.Formula)) { cell.Formula = nf; cell.Ast = null; }
                }
        wb.ActiveSheet = Math.Clamp(wb.ActiveSheet, 0, wb.Sheets.Count - 1);
        wb.MarkDirty();
    }

    public static Sheet DuplicateSheet(Workbook wb, Sheet s)
    {
        var copy = s.Clone(wb);
        string baseName = s.Name;
        int n = 2;
        string name;
        do { name = $"{baseName} ({n++})"; } while (wb.Sheet(name) != null);
        copy.Name = name.Length > 31 ? name[..31] : name;
        wb.Sheets.Insert(wb.Sheets.IndexOf(s) + 1, copy);
        foreach (var cell in copy.Cells.Values) cell.Ast = null;
        wb.MarkDirty();
        return copy;
    }

    // ------------------------------------------------------------------ autofit

    public static double AutoFitWidth(Sheet s, int col, Func<string, CellStyle, double> measure)
    {
        double max = 0;
        foreach (var (k, cell) in s.Cells)
        {
            var (r, c) = Sheet.FromKey(k);
            if (c != col || s.IsRowHidden(r) || s.MergeAt(r, c) is { IsSingle: false }) continue;
            var st = s.Book.Styles[cell.StyleId];
            var text = NumFmt.Format(cell.Value, st.NumberFormat, out _);
            if (text.Length == 0) continue;
            max = Math.Max(max, measure(text, st) + 10);
        }
        return max <= 0 ? Sheet.DefaultColWidth : Math.Min(max, 1200);
    }
}
