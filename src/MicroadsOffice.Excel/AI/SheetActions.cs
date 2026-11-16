using System.Drawing;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.AI;

/// <summary>
/// The "microads-sheet" action block: a JSON list of spreadsheet edits written by the assistant (or Claude Desktop).
/// Everything is validated first (addresses, formulas, function names); nothing is changed when any action is invalid.
/// </summary>
public static class SheetActions
{
    public const string Guide =
        """
        Action block for Microads Excel. Put a JSON object inside:
        ```microads-sheet
        {"actions":[ ... ]}
        ```
        Actions run in order on the active sheet (or the sheet chosen with "sheet"):
        {"sheet":"Tên trang"}                       switch the target sheet (must exist, or create it first with addSheet)
        {"addSheet":"Tổng hợp"}                     add a sheet (and make it the target)
        {"cell":"A1","value":"Doanh thu"}            text, number (JSON number) or date ("2026-01-31")
        {"cell":"D2","formula":"=SUM(B2:C2)"}        a formula
        {"range":"A2","values":[["Hà Nội",120,"=B2*2"],["HCM",340,"=B3*2"]]}  a block starting at the top-left cell; strings starting with "=" are formulas; null leaves a cell unchanged
        {"range":"A1:D1","bold":true,"italic":false,"fill":"#D9E1F2","fontColor":"#1F2937","numberFormat":"#,##0","align":"center","wrap":true,"border":true,"fontSize":12}
        {"clear":"A10:D20"}                         clear contents and formats
        {"autofit":"A:D"}                           fit column widths to the content
        {"columnWidth":{"columns":"B:C","width":120}}  width in pixels
        {"sort":{"range":"A1:D20","by":"C","ascending":false,"header":true}}
        {"freeze":"A2"}                             freeze rows above / columns left of this cell ("A1" removes)
        {"chart":{"type":"column|bar|line|pie|doughnut|area|scatter|stackedColumn","range":"A1:B6","title":"…","at":"F2"}}
        Formulas: English function names, comma "," between arguments, "." as decimal point, e.g. =IF(B2>=5,"Đạt","Không đạt"), =VLOOKUP(A2,Sheet2!A:B,2,FALSE). Only use functions this app supports (listed below). Number formats use Excel codes ("#,##0", "0.00%", "dd/mm/yyyy").
        Put headers in bold with a fill, use formulas (not hard-coded results) for totals, and autofit the columns you wrote.
        """;

    public static string FunctionList() => string.Join(", ", Functions.Table.Keys.OrderBy(k => k, StringComparer.Ordinal));

    public sealed record Op(string Kind, JsonObject Node);

    /// <summary>Parses and validates; throws <see cref="AiException"/> listing every problem.</summary>
    public static List<Op> Parse(string json, Workbook wb)
    {
        JsonNode? root;
        try { root = JsonNode.Parse(StripFence(json), documentOptions: new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip }); }
        catch (JsonException e) { throw new AiException(Loc.T("Nội dung bảng tính không phải JSON hợp lệ: ", "The sheet content is not valid JSON: ") + e.Message); }
        var arr = root as JsonArray ?? root?["actions"] as JsonArray ?? throw new AiException(Loc.T("Thiếu danh sách \"actions\".", "Missing the \"actions\" list."));
        var ops = new List<Op>();
        var errors = new List<string>();
        var sheets = new HashSet<string>(wb.Sheets.Select(s => s.Name), StringComparer.OrdinalIgnoreCase);
        int i = 0;
        foreach (var n in arr)
        {
            i++;
            if (n is not JsonObject o) { errors.Add($"#{i}: not an object"); continue; }
            string kind = o.ContainsKey("addSheet") ? "addSheet" : o.ContainsKey("sheet") ? "sheet" : o.ContainsKey("chart") ? "chart" : o.ContainsKey("sort") ? "sort"
                : o.ContainsKey("clear") ? "clear" : o.ContainsKey("autofit") ? "autofit" : o.ContainsKey("columnWidth") ? "columnWidth" : o.ContainsKey("freeze") ? "freeze"
                : o.ContainsKey("cell") ? "cell" : o.ContainsKey("range") ? (o.ContainsKey("values") ? "values" : "format") : "";
            try
            {
                switch (kind)
                {
                    case "addSheet":
                        var nm = Str(o, "addSheet").Trim();
                        if (nm.Length == 0 || nm.Length > 31 || nm.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0) throw new FormatException("invalid sheet name");
                        sheets.Add(nm);
                        break;
                    case "sheet":
                        if (!sheets.Contains(Str(o, "sheet"))) throw new FormatException($"sheet \"{Str(o, "sheet")}\" does not exist");
                        break;
                    case "cell":
                        Cell(Str(o, "cell"));
                        if (o["formula"] is JsonNode f) CheckFormula(f.GetValue<string>());
                        else if (o["value"] is JsonValue v && v.GetValueKind() == JsonValueKind.String && v.GetValue<string>().StartsWith('=')) CheckFormula(v.GetValue<string>());
                        if (o.ContainsKey("bold") || o.ContainsKey("fill") || o.ContainsKey("numberFormat")) CheckFormat(o);
                        break;
                    case "values":
                        var start = Cell(Str(o, "range").Split(':')[0]);
                        if (o["values"] is not JsonArray rows) throw new FormatException("values must be an array of rows");
                        if (start.Row + rows.Count > Addr.MaxRows) throw new FormatException("values go past the last row");
                        foreach (var row in rows)
                        {
                            if (row is not JsonArray cells) throw new FormatException("each row must be an array");
                            if (start.Col + cells.Count > Addr.MaxCols) throw new FormatException("values go past the last column");
                            foreach (var c in cells)
                                if (c is JsonValue cv && cv.GetValueKind() == JsonValueKind.String && cv.GetValue<string>().StartsWith('=')) CheckFormula(cv.GetValue<string>());
                        }
                        break;
                    case "format":
                        CheckArea(Range(Str(o, "range")));
                        CheckFormat(o);
                        break;
                    case "clear": CheckArea(Range(Str(o, "clear"))); break;
                    case "autofit": Columns(Str(o, "autofit")); break;
                    case "columnWidth":
                        Columns(o["columnWidth"]?["columns"]?.GetValue<string>() ?? throw new FormatException("columnWidth.columns missing"));
                        var w = o["columnWidth"]?["width"]?.GetValue<double>() ?? throw new FormatException("columnWidth.width missing");
                        if (w < 0 || w > 2000) throw new FormatException("width out of range");
                        break;
                    case "freeze": Cell(Str(o, "freeze")); break;
                    case "sort":
                        Range(o["sort"]?["range"]?.GetValue<string>() ?? throw new FormatException("sort.range missing"));
                        Addr.ColIndex(o["sort"]?["by"]?.GetValue<string>() ?? throw new FormatException("sort.by missing"));
                        break;
                    case "chart":
                        var ch = o["chart"] as JsonObject ?? throw new FormatException("chart must be an object");
                        Range(ch["range"]?.GetValue<string>() ?? throw new FormatException("chart.range missing"));
                        ChartKind(ch["type"]?.GetValue<string>() ?? "column");
                        if (ch["at"] is JsonNode at) Cell(at.GetValue<string>());
                        break;
                    default: throw new FormatException("unknown action");
                }
                ops.Add(new Op(kind, o));
            }
            catch (Exception e) when (e is FormatException or InvalidOperationException or AiException or ArgumentException or KeyNotFoundException)
            { errors.Add($"#{i} ({kind}): {e.Message}"); }
        }
        if (errors.Count > 0)
            throw new AiException(Loc.T("Không ghi gì vào trang tính vì có thao tác không hợp lệ:\n", "Nothing was written because some actions are invalid:\n") + string.Join("\n", errors.Take(12)));
        if (ops.Count == 0) throw new AiException(Loc.T("Không có thao tác nào.", "There are no actions."));
        return ops;
    }

    static string StripFence(string s)
    {
        s = s.Trim();
        if (s.StartsWith("```")) { int nl = s.IndexOf('\n'); s = nl > 0 ? s[(nl + 1)..] : ""; }
        if (s.EndsWith("```")) s = s[..^3];
        return s;
    }

    static string Str(JsonObject o, string k) => o[k]?.GetValue<string>() ?? throw new FormatException($"{k} missing");

    public static CellPos Cell(string a)
    {
        if (!Addr.TryParseA1(a.Replace("$", "").Trim(), out var p)) throw new FormatException($"invalid cell \"{a}\"");
        return p;
    }

    public static CellRange Range(string a)
    {
        var t = a.Replace("$", "").Trim();
        if (Addr.TryParseRange(t, out var r)) return r.Normalized;
        if (Addr.TryParseA1(t, out var p)) return new CellRange(p.Row, p.Col, p.Row, p.Col);
        throw new FormatException($"invalid range \"{a}\"");
    }

    const long MaxArea = 200_000;
    static bool WholeRows(CellRange r) => r.R1 == 0 && r.R2 >= Addr.MaxRows - 1;
    static bool WholeCols(CellRange r) => r.C1 == 0 && r.C2 >= Addr.MaxCols - 1;

    /// <summary>A whole-column ("B:B") or whole-row ("3:3") range is fine (it is clipped to the data); any other range must stay reasonable.</summary>
    static void CheckArea(CellRange r)
    {
        long rows = WholeRows(r) ? 1 : r.R2 - r.R1 + 1, cols = WholeCols(r) ? 1 : r.C2 - r.C1 + 1;
        if (rows * cols > MaxArea) throw new FormatException($"range {r} is too large ({rows * cols:N0} cells; at most {MaxArea:N0})");
    }

    /// <summary>Limits whole-column / whole-row ranges to the part of the sheet in use, so "B:B" never creates a million cells.</summary>
    static CellRange Clip(Sheet s, CellRange r)
    {
        var used = s.UsedRange();
        int r2 = WholeRows(r) ? Math.Max(r.R1, used.R2) : r.R2;
        int c2 = WholeCols(r) ? Math.Max(r.C1, used.C2) : r.C2;
        return new CellRange(r.R1, r.C1, r2, c2);
    }

    static (int C1, int C2) Columns(string a)
    {
        var parts = a.Replace("$", "").Trim().Split(':');
        int c1 = Addr.ColIndex(new string(parts[0].TakeWhile(char.IsLetter).ToArray()));
        int c2 = parts.Length > 1 ? Addr.ColIndex(new string(parts[1].TakeWhile(char.IsLetter).ToArray())) : c1;
        if (c1 < 0 || c2 < 0) throw new FormatException($"invalid columns \"{a}\"");
        return (Math.Min(c1, c2), Math.Max(c1, c2));
    }

    public static void CheckFormula(string f)
    {
        if (!FormulaParser.TryParse(f, out var e, out var err) || e == null) throw new FormatException($"formula {f}: {err}");
        foreach (var name in FunctionNames(e))
            if (!Functions.Table.ContainsKey(name)) throw new FormatException($"formula {f}: function {name} is not supported");
    }

    static IEnumerable<string> FunctionNames(Expr e)
    {
        switch (e)
        {
            case FuncExpr fe:
                yield return fe.Name;
                foreach (var a in fe.Args) foreach (var n in FunctionNames(a)) yield return n;
                break;
            case UnaryExpr u: foreach (var n in FunctionNames(u.Operand)) yield return n; break;
            case BinaryExpr b: foreach (var n in FunctionNames(b.Left).Concat(FunctionNames(b.Right))) yield return n; break;
            case PercentExpr p: foreach (var n in FunctionNames(p.Operand)) yield return n; break;
            case ParenExpr pe: foreach (var n in FunctionNames(pe.Inner)) yield return n; break;
        }
    }

    static void CheckFormat(JsonObject o)
    {
        if (o["fill"] is JsonNode f) Color(f.GetValue<string>());
        if (o["fontColor"] is JsonNode c) Color(c.GetValue<string>());
        if (o["align"] is JsonNode a && a.GetValue<string>() is not ("left" or "center" or "right" or "justify" or "general")) throw new FormatException("align must be left/center/right/justify");
        if (o["fontSize"] is JsonNode fs && fs.GetValue<double>() is < 4 or > 96) throw new FormatException("fontSize out of range");
    }

    static Color Color(string hex)
    {
        var h = hex.Trim().TrimStart('#');
        if (h.Length != 6 || !int.TryParse(h, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v)) throw new FormatException($"invalid colour \"{hex}\"");
        return System.Drawing.Color.FromArgb(255, (v >> 16) & 255, (v >> 8) & 255, v & 255);
    }

    static ChartType ChartKind(string t) => t.ToLowerInvariant() switch
    {
        "column" => ChartType.Column, "stackedcolumn" => ChartType.StackedColumn, "bar" => ChartType.Bar, "line" => ChartType.Line,
        "pie" => ChartType.Pie, "doughnut" => ChartType.Doughnut, "area" => ChartType.Area, "scatter" => ChartType.Scatter,
        _ => throw new FormatException($"unknown chart type \"{t}\""),
    };

    /// <summary>
    /// Applies validated actions to <paramref name="wb"/>. The caller has already taken ONE undo checkpoint.
    /// <paramref name="measure"/> measures text for autofit (null = estimate from character count). Returns a summary.
    /// </summary>
    public static string Apply(List<Op> ops, Workbook wb, Func<string, CellStyle, double>? measure = null)
    {
        var sheet = wb.Active;
        int cells = 0, formats = 0, charts = 0, other = 0;
        var touchedCols = new HashSet<int>();
        foreach (var op in ops)
        {
            var o = op.Node;
            switch (op.Kind)
            {
                case "addSheet":
                    {
                        var name = Str(o, "addSheet").Trim();
                        sheet = wb.Sheet(name) ?? wb.AddSheet(name);
                        other++;
                        break;
                    }
                case "sheet":
                    sheet = wb.Sheet(Str(o, "sheet"))!;
                    break;
                case "cell":
                    {
                        var p = Cell(Str(o, "cell"));
                        if (o["formula"] is JsonNode f) SheetOps.SetInput(sheet, p.Row, p.Col, "=" + f.GetValue<string>().TrimStart('='));
                        else if (o["value"] is JsonNode v) SetValue(sheet, p.Row, p.Col, v);
                        cells++; touchedCols.Add(p.Col);
                        if (o.ContainsKey("bold") || o.ContainsKey("fill") || o.ContainsKey("numberFormat") || o.ContainsKey("fontColor") || o.ContainsKey("align"))
                        { Format(wb, sheet, new CellRange(p.Row, p.Col, p.Row, p.Col), o); formats++; }
                        break;
                    }
                case "values":
                    {
                        var p = Cell(Str(o, "range").Split(':')[0]);
                        int r = 0;
                        foreach (var row in (JsonArray)o["values"]!)
                        {
                            int c = 0;
                            foreach (var v in (JsonArray)row!)
                            {
                                if (v != null) { SetValue(sheet, p.Row + r, p.Col + c, v); cells++; touchedCols.Add(p.Col + c); }
                                c++;
                            }
                            r++;
                        }
                        break;
                    }
                case "format":
                    Format(wb, sheet, Clip(sheet, Range(Str(o, "range"))), o); formats++;
                    break;
                case "clear":
                    {
                        var rg = Clip(sheet, Range(Str(o, "clear")));
                        for (int r = rg.R1; r <= rg.R2; r++)
                            for (int c = rg.C1; c <= rg.C2; c++) sheet.ClearCell(r, c, true, true);
                        other++;
                        break;
                    }
                case "autofit":
                    {
                        var (c1, c2) = Columns(Str(o, "autofit"));
                        for (int c = c1; c <= c2; c++) AutoFit(sheet, c, measure);
                        other++;
                        break;
                    }
                case "columnWidth":
                    {
                        var (c1, c2) = Columns(o["columnWidth"]!["columns"]!.GetValue<string>());
                        var w = o["columnWidth"]!["width"]!.GetValue<double>();
                        for (int c = c1; c <= c2; c++) { if (w == 0) sheet.HiddenCols.Add(c); else { sheet.HiddenCols.Remove(c); sheet.ColWidths[c] = w; } }
                        other++;
                        break;
                    }
                case "freeze":
                    {
                        var p = Cell(Str(o, "freeze"));
                        sheet.FreezeRows = p.Row; sheet.FreezeCols = p.Col;
                        other++;
                        break;
                    }
                case "sort":
                    {
                        var s = o["sort"]!;
                        var rg = Range(s["range"]!.GetValue<string>());
                        int col = Addr.ColIndex(s["by"]!.GetValue<string>());
                        SheetOps.Sort(sheet, rg, [new SheetOps.SortKey(col, s["ascending"]?.GetValue<bool>() ?? true)], s["header"]?.GetValue<bool>() ?? true);
                        other++;
                        break;
                    }
                case "chart":
                    {
                        var ch = (JsonObject)o["chart"]!;
                        var rg = Range(ch["range"]!.GetValue<string>());
                        float x = 320, y = 40;
                        if (ch["at"] is JsonNode at)
                        {
                            var p = Cell(at.GetValue<string>());
                            x = 0; for (int c = 0; c < p.Col; c++) x += (float)sheet.ColWidth(c);
                            y = 0; for (int r = 0; r < Math.Min(p.Row, 5000); r++) y += (float)sheet.RowHeight(r);
                        }
                        else
                        {
                            x = 0; for (int c = 0; c <= rg.C2 + 1; c++) x += (float)sheet.ColWidth(c);
                            y = 0; for (int r = 0; r < rg.R1; r++) y += (float)sheet.RowHeight(r);
                        }
                        sheet.Charts.Add(new Chart
                        {
                            Type = ChartKind(ch["type"]?.GetValue<string>() ?? "column"), Title = ch["title"]?.GetValue<string>() ?? "", Data = rg,
                            Bounds = new RectangleF(x, y, 480, 290), Name = "Chart " + (sheet.Charts.Count + 1),
                        });
                        charts++;
                        break;
                    }
            }
        }
        wb.ActiveSheet = wb.Sheets.IndexOf(sheet);
        var parts = new List<string>();
        if (cells > 0) parts.Add(Loc.F("{0} ô", "{0} cells", cells));
        if (formats > 0) parts.Add(Loc.F("{0} định dạng", "{0} formats", formats));
        if (charts > 0) parts.Add(Loc.F("{0} biểu đồ", "{0} charts", charts));
        if (other > 0) parts.Add(Loc.F("{0} thao tác khác", "{0} other actions", other));
        return Loc.T("Đã ghi vào trang tính: ", "Written to the sheet: ") + string.Join(", ", parts) + ".";
    }

    static void SetValue(Sheet s, int r, int c, JsonNode v)
    {
        switch (v.GetValueKind())
        {
            case JsonValueKind.Number: s.SetValue(r, c, CellValue.Number(v.GetValue<double>())); break;
            case JsonValueKind.True: case JsonValueKind.False: s.SetValue(r, c, CellValue.Bool(v.GetValue<bool>())); break;
            case JsonValueKind.String:
                var t = v.GetValue<string>();
                if (t.StartsWith('=')) SheetOps.SetInput(s, r, c, t);
                else if (DateTime.TryParseExact(t, ["yyyy-MM-dd", "yyyy-MM-ddTHH:mm", "yyyy-MM-dd HH:mm"], CultureInfo.InvariantCulture, DateTimeStyles.None, out _)) SheetOps.SetInput(s, r, c, t);
                else { s.SetValue(r, c, CellValue.Text(t)); }
                break;
            case JsonValueKind.Null: break;
            default: s.SetValue(r, c, CellValue.Text(v.ToJsonString())); break;
        }
    }

    static void Format(Workbook wb, Sheet s, CellRange rg, JsonObject o)
    {
        Func<CellStyle, CellStyle> f = st =>
        {
            if (o["bold"] is JsonNode b) st = st with { Bold = b.GetValue<bool>() };
            if (o["italic"] is JsonNode it) st = st with { Italic = it.GetValue<bool>() };
            if (o["fill"] is JsonNode fi) st = st with { Fill = Color(fi.GetValue<string>()).ToArgb() };
            if (o["fontColor"] is JsonNode fc) st = st with { FontColor = Color(fc.GetValue<string>()).ToArgb() };
            if (o["numberFormat"] is JsonNode nf) st = st with { NumberFormat = nf.GetValue<string>() };
            if (o["wrap"] is JsonNode w) st = st with { Wrap = w.GetValue<bool>() };
            if (o["fontSize"] is JsonNode fs) st = st with { FontSize = (float)fs.GetValue<double>() };
            if (o["align"] is JsonNode a)
                st = st with { HAlign = a.GetValue<string>() switch { "left" => HAlign.Left, "center" => HAlign.Center, "right" => HAlign.Right, "justify" => HAlign.Justify, _ => HAlign.General } };
            if (o["border"] is JsonNode bd && bd.GetValue<bool>())
                st = st with { Left = BorderLine.Thin(), Right = BorderLine.Thin(), Top = BorderLine.Thin(), Bottom = BorderLine.Thin() };
            return st;
        };
        var map = new Dictionary<int, int>();
        for (int r = rg.R1; r <= rg.R2; r++)
            for (int c = rg.C1; c <= rg.C2; c++)
            {
                int id = s.StyleIdAt(r, c);
                if (!map.TryGetValue(id, out var n)) map[id] = n = wb.Styles.Intern(f(wb.Styles[id]));
                s.SetStyle(r, c, n);
            }
    }

    static void AutoFit(Sheet s, int c, Func<string, CellStyle, double>? measure)
    {
        measure ??= (text, st) => text.Length * 7.0 * (st.FontSize / 11.0) + 4;
        var w = SheetOps.AutoFitWidth(s, c, measure);
        s.ColWidths[c] = Math.Round(w);
        s.HiddenCols.Remove(c);
    }

    // ------------------------------------------------------------------ context for the model

    /// <summary>The workbook as text: sheet list, selection, then the non-empty cells of the active sheet (value and formula).</summary>
    public static string Describe(Workbook wb, CellRange selection, CellPos active, int maxChars)
    {
        var s = wb.Active;
        var sb = new StringBuilder();
        sb.AppendLine("Sheets: " + string.Join(", ", wb.Sheets.Select(x => x == s ? x.Name + " (active)" : x.Name)));
        var u = s.DataRange();
        bool empty = !s.Cells.Values.Any(c => !c.Value.IsBlank || c.Formula != null);
        sb.AppendLine($"Active sheet: {s.Name}; data range: {(empty ? "(empty)" : u.ToString())}; selection: {selection}; active cell: {active}");
        if (s.Charts.Count > 0) sb.AppendLine("Charts: " + string.Join("; ", s.Charts.Select(ch => $"{ch.Name} {ch.Type} {ch.Data} \"{ch.Title}\"")));
        if (empty) return sb.ToString();
        sb.AppendLine("Cells (address: shown value [= formula]):");
        for (int r = u.R1; r <= u.R2; r++)
        {
            for (int c = u.C1; c <= u.C2; c++)
            {
                var cell = s.GetCell(r, c);
                if (cell == null || (cell.Value.IsBlank && cell.Formula == null)) continue;
                var shown = SheetOps.DisplayText(s, r, c);
                sb.Append(Addr.ToA1(r, c, false, false)).Append(": ").Append(shown.Replace('\n', ' '));
                if (cell.Formula != null) sb.Append("  [=").Append(cell.Formula).Append(']');
                sb.AppendLine();
                if (sb.Length > maxChars) { sb.AppendLine($"[… truncated at row {r + 1} …]"); return sb.ToString(); }
            }
        }
        return sb.ToString();
    }
}
