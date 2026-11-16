using System.Net;
using System.Text;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel;

/// <summary>CSV / TSV import-export (UTF-8 with BOM, delimiter detection) and HTML export.</summary>
public static class CsvIO
{
    public static char DetectDelimiter(string text)
    {
        var sample = text.Length > 20000 ? text[..20000] : text;
        char best = ',';
        double bestScore = 0;
        foreach (var d in new[] { '	', ';', ',', '|' })
        {
            var rows = ParseDelimited(sample, d).Take(30).Where(r => r.Count > 0 && !(r.Count == 1 && r[0].Length == 0)).ToList();
            if (rows.Count == 0) continue;
            int fields = rows[0].Count;
            if (fields < 2) continue;
            double consistency = rows.Count(r => r.Count == fields) / (double)rows.Count;
            double score = consistency * 100 + Math.Min(fields, 50);
            if (score > bestScore) { bestScore = score; best = d; }
        }
        return best;
    }

    public static List<List<string>> ParseDelimited(string text, char delim)
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool inQ = false;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (inQ)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else inQ = false;
                }
                else field.Append(ch);
                continue;
            }
            if (ch == '"' && field.Length == 0) { inQ = true; continue; }
            if (ch == delim) { row.Add(field.ToString()); field.Clear(); continue; }
            if (ch == '\r') continue;
            if (ch == '\n') { row.Add(field.ToString()); field.Clear(); rows.Add(row); row = new(); continue; }
            field.Append(ch);
        }
        if (field.Length > 0 || row.Count > 0) { row.Add(field.ToString()); rows.Add(row); }
        return rows;
    }

    public static Workbook Read(string path)
    {
        var bytes = File.ReadAllBytes(path);
        string text = DecodeText(bytes);
        var wb = Workbook.CreateBlank();
        var s = wb.Sheets[0];
        s.Name = SafeSheetName(Path.GetFileNameWithoutExtension(path));
        char delim = Path.GetExtension(path).Equals(".tsv", StringComparison.OrdinalIgnoreCase) ? '\t' : DetectDelimiter(text);
        var rows = ParseDelimited(text, delim);
        for (int r = 0; r < rows.Count && r < Addr.MaxRows; r++)
            for (int c = 0; c < rows[r].Count && c < Addr.MaxCols; c++)
            {
                var f = rows[r][c];
                if (f.Length == 0) continue;
                // values only: a CSV field beginning with '=' stays text unless it parses as a formula Excel would accept
                if (f.StartsWith('=') && f.Length > 1) { s.SetFormula(r, c, f[1..]); continue; }
                SheetOps.SetInput(s, r, c, f);
            }
        wb.Recalculate();
        return wb;
    }

    public static string DecodeText(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        try { return new UTF8Encoding(false, true).GetString(bytes); }
        catch (DecoderFallbackException) { return Encoding.Latin1.GetString(bytes); }
    }

    public static string SafeSheetName(string n)
    {
        var s = new string(n.Where(ch => ch is not (':' or '\\' or '/' or '?' or '*' or '[' or ']')).ToArray()).Trim('\'');
        if (s.Length == 0) s = "Sheet1";
        return s.Length > 31 ? s[..31] : s;
    }

    /// <summary>Writes the active sheet's used range (displayed values) as CSV/TSV with a UTF-8 BOM (Excel opens Vietnamese text correctly).</summary>
    public static void Write(Sheet s, string path, char delim)
    {
        var used = s.DataRange();
        var sb = new StringBuilder();
        if (s.Cells.Count > 0)
            for (int r = 0; r <= used.R2; r++)
            {
                for (int c = 0; c <= used.C2; c++)
                {
                    if (c > 0) sb.Append(delim);
                    var cell = s.GetCell(r, c);
                    if (cell == null) continue;
                    var st = s.Book.Styles[cell.StyleId];
                    string text = cell.Value.IsNumber && !NumFmt.IsDateFormat(st.NumberFormat) && st.NumberFormat == "General"
                        ? cell.Value.ToString() : NumFmt.Format(cell.Value, st.NumberFormat, out _);
                    if (text.IndexOf(delim) >= 0 || text.Contains('"') || text.Contains('\n') || text.Contains('\r'))
                        text = "\"" + text.Replace("\"", "\"\"") + "\"";
                    sb.Append(text);
                }
                sb.Append("\r\n");
            }
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
    }

    private static string Html(string s) => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    /// <summary>Exports every visible sheet as an HTML document with basic formatting.</summary>
    public static void WriteHtml(Workbook wb, string path)
    {
        var sb = new StringBuilder();
        sb.Append("<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>").Append(Html(wb.Title.Length > 0 ? wb.Title : Path.GetFileNameWithoutExtension(path)))
          .Append("</title><style>body{font-family:Calibri,Segoe UI,Arial,sans-serif}table{border-collapse:collapse;margin-bottom:24px}td{padding:2px 6px;border:1px solid #e1e3e6;vertical-align:bottom;white-space:nowrap}h2{font-size:15px}</style></head><body>");
        foreach (var s in wb.Sheets.Where(x => !x.Hidden))
        {
            sb.Append("<h2>").Append(Html(s.Name)).Append("</h2><table>");
            var used = s.UsedRange();
            var covered = new HashSet<(int, int)>();
            for (int r = 0; r <= used.R2 && s.Cells.Count > 0; r++)
            {
                if (s.IsRowHidden(r)) continue;
                sb.Append("<tr>");
                for (int c = 0; c <= used.C2; c++)
                {
                    if (s.HiddenCols.Contains(c) || covered.Contains((r, c))) continue;
                    var cell = s.GetCell(r, c);
                    var st = s.StyleAt(r, c);
                    var css = new StringBuilder();
                    if (st.Bold) css.Append("font-weight:bold;");
                    if (st.Italic) css.Append("font-style:italic;");
                    if (st.Underline) css.Append("text-decoration:underline;");
                    if (st.FontColor != 0) css.Append($"color:#{st.FontColor & 0xFFFFFF:X6};");
                    if (st.Fill != 0) css.Append($"background:#{st.Fill & 0xFFFFFF:X6};");
                    var v = cell?.Value ?? CellValue.Blank;
                    var align = st.HAlign switch { HAlign.Center => "center", HAlign.Right => "right", HAlign.Left => "left", _ => v.IsNumber ? "right" : v.IsBool || v.IsError ? "center" : "left" };
                    css.Append("text-align:").Append(align).Append(';');
                    if (st.Wrap) css.Append("white-space:normal;");
                    css.Append($"width:{s.ColWidth(c):0}px;");
                    var span = "";
                    if (s.MergeAt(r, c) is CellRange m && m.R1 == r && m.C1 == c)
                    {
                        span = $" rowspan=\"{m.Rows}\" colspan=\"{m.Cols}\"";
                        foreach (var p in m.Cells()) covered.Add(p);
                    }
                    var text = NumFmt.Format(v, st.NumberFormat, out _);
                    sb.Append("<td").Append(span).Append(" style=\"").Append(css).Append("\">").Append(Html(text).Replace("\n", "<br>")).Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
        }
        sb.Append("<p style=\"color:#888;font-size:11px\">Microads Office Excel — microads.vn</p></body></html>");
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
    }
}
