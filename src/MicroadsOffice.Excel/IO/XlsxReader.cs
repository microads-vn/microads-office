using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.IO;

/// <summary>Reads .xlsx / .xlsm / .xltx files.</summary>
public static class XlsxReader
{
    private static readonly XNamespace M = XlsxWriter.NsMain;
    private static readonly XNamespace R = XlsxWriter.NsRel;
    private static readonly XNamespace P = XlsxWriter.NsPkgRel;
    private static readonly XNamespace X = XlsxWriter.NsDrawing;
    private static readonly XNamespace A = XlsxWriter.NsA;
    private static readonly XNamespace C = XlsxWriter.NsC;
    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    private static readonly int[] ThemeColors =
    [
        0xFFFFFF, 0x000000, 0xE7E6E6, 0x44546A, 0x4472C4, 0xED7D31, 0xA5A5A5, 0xFFC000, 0x5B9BD5, 0x70AD47, 0x0563C1, 0x954F72
    ];

    private static readonly int[] Indexed =
    [
        0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
        0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
        0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
        0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333
    ];

    public static Workbook Read(string path)
    {
        using var fs = File.OpenRead(path);
        return Read(fs);
    }

    public static Workbook Read(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var wb = new Workbook();
        XDocument? Load(string name)
        {
            var e = zip.GetEntry(name.TrimStart('/')) ?? zip.Entries.FirstOrDefault(x => x.FullName.Equals(name.TrimStart('/'), StringComparison.OrdinalIgnoreCase));
            if (e == null) return null;
            using var s = e.Open();
            return XDocument.Load(s);
        }
        byte[]? LoadBytes(string name)
        {
            var e = zip.GetEntry(name.TrimStart('/')) ?? zip.Entries.FirstOrDefault(x => x.FullName.Equals(name.TrimStart('/'), StringComparison.OrdinalIgnoreCase));
            if (e == null) return null;
            using var s = e.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return ms.ToArray();
        }
        Dictionary<string, (string Target, string Type, bool External)> Rels(string part)
        {
            var dir = Path.GetDirectoryName(part)!.Replace('\\', '/');
            var relPath = (dir.Length > 0 ? dir + "/" : "") + "_rels/" + Path.GetFileName(part) + ".rels";
            var d = new Dictionary<string, (string, string, bool)>();
            var doc = Load(relPath);
            if (doc == null) return d;
            foreach (var r in doc.Root!.Elements(P + "Relationship"))
            {
                var target = (string?)r.Attribute("Target") ?? "";
                bool ext = (string?)r.Attribute("TargetMode") == "External";
                if (!ext) target = Resolve(dir, target);
                d[(string)r.Attribute("Id")!] = (target, (string?)r.Attribute("Type") ?? "", ext);
            }
            return d;
        }

        // workbook part
        var rootRels = Load("_rels/.rels");
        var wbPath = rootRels?.Root?.Elements(P + "Relationship").FirstOrDefault(r => ((string?)r.Attribute("Type") ?? "").EndsWith("/officeDocument"))?.Attribute("Target")?.Value ?? "xl/workbook.xml";
        wbPath = wbPath.TrimStart('/');
        var wdoc = Load(wbPath) ?? throw new InvalidDataException("workbook.xml missing");
        var wrels = Rels(wbPath);
        // shared strings
        var sst = new List<string>();
        var sstPart = wrels.Values.FirstOrDefault(r => r.Type.EndsWith("/sharedStrings")).Target;
        var sdoc = sstPart != null ? Load(sstPart) : null;
        if (sdoc != null) foreach (var si in sdoc.Root!.Elements(M + "si")) sst.Add(RichText(si));
        // styles
        var stylesPart = wrels.Values.FirstOrDefault(r => r.Type.EndsWith("/styles")).Target;
        var (xfMap, dxfs) = ReadStyles(wb, stylesPart != null ? Load(stylesPart) : null);

        var wroot = wdoc.Root!;
        var sheetsEl = wroot.Element(M + "sheets")?.Elements(M + "sheet").ToList() ?? new();
        int active = (int?)wroot.Element(M + "bookViews")?.Element(M + "workbookView")?.Attribute("activeTab") ?? 0;
        var sheetParts = new List<(Sheet Sheet, string Part)>();
        foreach (var se in sheetsEl)
        {
            var name = (string?)se.Attribute("name") ?? wb.NextSheetName();
            var rid = (string?)se.Attribute(R + "id");
            if (rid == null || !wrels.TryGetValue(rid, out var rel)) continue;
            if (!rel.Type.EndsWith("/worksheet")) continue; // chartsheets are skipped
            var sheet = new Sheet(wb, name) { Hidden = (string?)se.Attribute("state") is "hidden" or "veryHidden" };
            wb.Sheets.Add(sheet);
            sheetParts.Add((sheet, rel.Target));
        }
        if (wb.Sheets.Count == 0) wb.AddSheet();
        wb.ActiveSheet = Math.Clamp(active, 0, wb.Sheets.Count - 1);
        if ((string?)wroot.Element(M + "calcPr")?.Attribute("calcMode") == "manual") wb.ManualCalc = true;

        foreach (var (sheet, part) in sheetParts)
        {
            var doc = Load(part);
            if (doc == null) continue;
            ReadSheet(sheet, doc, sst, xfMap, dxfs);
            var rels = Rels(part);
            foreach (var hl in doc.Root!.Element(M + "hyperlinks")?.Elements(M + "hyperlink") ?? [])
            {
                if (!Addr.TryParseA1(((string?)hl.Attribute("ref") ?? "").Split(':')[0], out var p)) continue;
                var rid = (string?)hl.Attribute(R + "id");
                string? url = rid != null && rels.TryGetValue(rid, out var rr) ? rr.Target : null;
                var loc = (string?)hl.Attribute("location");
                if (url == null && loc != null) url = "#" + loc;
                if (url != null) sheet.Links[Sheet.Key(p.Row, p.Col)] = url;
            }
            foreach (var rel in rels.Values)
            {
                if (rel.Type.EndsWith("/comments")) ReadComments(sheet, Load(rel.Target));
                else if (rel.Type.EndsWith("/drawing")) ReadDrawing(sheet, rel.Target, Load, LoadBytes, Rels);
            }
        }
        // defined names
        foreach (var dn in wroot.Element(M + "definedNames")?.Elements(M + "definedName") ?? [])
        {
            var name = (string?)dn.Attribute("name") ?? "";
            var text = dn.Value.Trim();
            int? local = (int?)dn.Attribute("localSheetId");
            if (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
            {
                var s = local is int li && li < sheetsEl.Count ? wb.Sheet((string?)sheetsEl[li].Attribute("name") ?? "") : null;
                if (s == null) continue;
                var refPart = text.Contains('!') ? text[(text.LastIndexOf('!') + 1)..] : text;
                if (name.Equals("_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase) && Addr.TryParseRange(refPart.Split(',')[0], out var pa)) s.Page.PrintArea = pa;
                else if (name.Equals("_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var piece in text.Split(','))
                    {
                        var rp = piece.Contains('!') ? piece[(piece.LastIndexOf('!') + 1)..] : piece;
                        if (Addr.TryParseRange(rp, out var tr) && tr.C1 == 0 && tr.C2 == Addr.MaxCols - 1) s.Page.PrintTitleRows = (tr.R1, tr.R2);
                    }
                }
                continue;
            }
            text = FileFormula.FromFile(text);
            if (local != null || (string?)dn.Attribute("hidden") == "1") { if (local != null && !wb.Names.ContainsKey(name)) wb.Names[name] = text; continue; }
            wb.Names[name] = text;
        }
        var core = Load("docProps/core.xml");
        if (core != null)
        {
            wb.Title = core.Root!.Elements().FirstOrDefault(e => e.Name.LocalName == "title")?.Value ?? "";
            wb.Author = core.Root!.Elements().FirstOrDefault(e => e.Name.LocalName == "creator")?.Value ?? "";
        }
        wb.Recalculate();
        return wb;
    }

    private static string Resolve(string dir, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = (dir.Length > 0 ? dir + "/" + target : target).Split('/').ToList();
        var stack = new List<string>();
        foreach (var p in parts)
        {
            if (p == "..") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); }
            else if (p != "." && p.Length > 0) stack.Add(p);
        }
        return string.Join("/", stack);
    }

    private static string RichText(XElement si)
    {
        var t = si.Element(M + "t");
        if (t != null && !si.Elements(M + "r").Any()) return Unescape(t.Value);
        return Unescape(string.Concat(si.Elements(M + "r").Select(r => r.Element(M + "t")?.Value ?? "")));
    }

    private static string Unescape(string s)
    {
        if (!s.Contains("_x")) return s;
        return System.Text.RegularExpressions.Regex.Replace(s, "_x([0-9A-Fa-f]{4})_", m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());
    }

    internal static int Color(XElement? e)
    {
        if (e == null) return 0;
        var rgb = (string?)e.Attribute("rgb");
        int baseRgb;
        if (rgb != null && rgb.Length >= 6)
        {
            var hex = rgb.Length == 8 ? rgb[2..] : rgb[^6..];
            baseRgb = int.Parse(hex, NumberStyles.HexNumber);
        }
        else if ((int?)e.Attribute("theme") is int th && th >= 0 && th < ThemeColors.Length) baseRgb = ThemeColors[th];
        else if ((int?)e.Attribute("indexed") is int ix)
        {
            if (ix == 64 || ix >= Indexed.Length) return 0;
            baseRgb = Indexed[ix];
        }
        else return 0;
        double tint = (double?)e.Attribute("tint") ?? 0;
        var c = System.Drawing.Color.FromArgb(255, (baseRgb >> 16) & 255, (baseRgb >> 8) & 255, baseRgb & 255);
        if (tint > 0) c = Blend(c, System.Drawing.Color.White, tint);
        else if (tint < 0) c = Blend(c, System.Drawing.Color.Black, -tint);
        return c.ToArgb();
    }

    private static System.Drawing.Color Blend(System.Drawing.Color a, System.Drawing.Color b, double t)
        => System.Drawing.Color.FromArgb(255, (int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));

    // ------------------------------------------------------------------ styles

    private sealed class Dxf { public int FontColor, Fill; public bool Bold; }

    private static (List<int> XfMap, List<Dxf> Dxfs) ReadStyles(Workbook wb, XDocument? doc)
    {
        var map = new List<int>();
        var dxfs = new List<Dxf>();
        if (doc == null) return (map, dxfs);
        var root = doc.Root!;
        var numFmts = new Dictionary<int, string>();
        foreach (var nf in root.Element(M + "numFmts")?.Elements(M + "numFmt") ?? []) numFmts[(int)nf.Attribute("numFmtId")!] = (string?)nf.Attribute("formatCode") ?? "General";
        var fonts = root.Element(M + "fonts")?.Elements(M + "font").ToList() ?? new();
        var fills = root.Element(M + "fills")?.Elements(M + "fill").ToList() ?? new();
        var borders = root.Element(M + "borders")?.Elements(M + "border").ToList() ?? new();
        foreach (var xf in root.Element(M + "cellXfs")?.Elements(M + "xf") ?? [])
        {
            var st = new CellStyle();
            int numId = (int?)xf.Attribute("numFmtId") ?? 0;
            var code = numFmts.TryGetValue(numId, out var c) ? c : NumFmt.BuiltIn(numId) ?? "General";
            if (numId == 14) code = "m/d/yyyy";
            st = st with { NumberFormat = code };
            int fi = (int?)xf.Attribute("fontId") ?? 0;
            if (fi < fonts.Count)
            {
                var f = fonts[fi];
                var u = f.Element(M + "u");
                st = st with
                {
                    FontName = (string?)f.Element(M + "name")?.Attribute("val") ?? CellStyle.DefaultFont,
                    FontSize = (float?)f.Element(M + "sz")?.Attribute("val") ?? CellStyle.DefaultSize,
                    Bold = IsOn(f.Element(M + "b")), Italic = IsOn(f.Element(M + "i")), Strike = IsOn(f.Element(M + "strike")),
                    Underline = u != null && (string?)u.Attribute("val") != "none",
                    FontColor = NormalizeFontColor(Color(f.Element(M + "color")))
                };
            }
            int fl = (int?)xf.Attribute("fillId") ?? 0;
            if (fl < fills.Count)
            {
                var pf = fills[fl].Element(M + "patternFill");
                var pt = (string?)pf?.Attribute("patternType");
                if (pt != null && pt != "none" && pt != "gray125")
                {
                    var col = Color(pf!.Element(M + "fgColor"));
                    if (col == 0) col = Color(pf.Element(M + "bgColor"));
                    st = st with { Fill = col };
                }
            }
            int bi = (int?)xf.Attribute("borderId") ?? 0;
            if (bi < borders.Count)
            {
                var b = borders[bi];
                st = st with { Left = Side(b.Element(M + "left")), Right = Side(b.Element(M + "right")), Top = Side(b.Element(M + "top")), Bottom = Side(b.Element(M + "bottom")) };
            }
            var al = xf.Element(M + "alignment");
            if (al != null)
            {
                st = st with
                {
                    HAlign = (string?)al.Attribute("horizontal") switch { "left" => HAlign.Left, "center" => HAlign.Center, "right" => HAlign.Right, "justify" or "distributed" => HAlign.Justify, "centerContinuous" => HAlign.CenterAcross, _ => HAlign.General },
                    VAlign = (string?)al.Attribute("vertical") switch { "center" => VAlign.Center, "top" => VAlign.Top, _ => VAlign.Bottom },
                    Wrap = IsTrue(al.Attribute("wrapText")),
                    Indent = (int?)al.Attribute("indent") ?? 0
                };
            }
            var pr = xf.Element(M + "protection");
            if (pr != null) st = st with { Locked = (string?)pr.Attribute("locked") != "0", HiddenFormula = IsTrue(pr.Attribute("hidden")) };
            map.Add(wb.Styles.Intern(st));
        }
        foreach (var d in root.Element(M + "dxfs")?.Elements(M + "dxf") ?? [])
        {
            var x = new Dxf();
            var f = d.Element(M + "font");
            if (f != null) { x.Bold = IsOn(f.Element(M + "b")); x.FontColor = Color(f.Element(M + "color")); }
            var pf = d.Element(M + "fill")?.Element(M + "patternFill");
            if (pf != null) { x.Fill = Color(pf.Element(M + "bgColor")); if (x.Fill == 0) x.Fill = Color(pf.Element(M + "fgColor")); }
            dxfs.Add(x);
        }
        return (map, dxfs);
    }

    private static int NormalizeFontColor(int c) => c == System.Drawing.Color.Black.ToArgb() ? 0 : c;

    private static bool IsOn(XElement? e) => e != null && (string?)e.Attribute("val") is null or "1" or "true";
    private static bool IsTrue(XAttribute? a) => a != null && (a.Value == "1" || a.Value == "true");

    private static BorderLine Side(XElement? e)
    {
        var st = (string?)e?.Attribute("style");
        if (st == null) return BorderLine.None;
        var ls = st switch
        {
            "medium" or "mediumDashed" or "mediumDashDot" => LineStyle.Medium, "thick" => LineStyle.Thick, "dashed" or "dashDot" or "dashDotDot" => LineStyle.Dashed,
            "dotted" => LineStyle.Dotted, "double" => LineStyle.Double, "hair" => LineStyle.Hair, _ => LineStyle.Thin
        };
        var c = Color(e!.Element(M + "color"));
        return new BorderLine(ls, c == 0 ? System.Drawing.Color.Black.ToArgb() : c);
    }

    // ------------------------------------------------------------------ sheet

    private static void ReadSheet(Sheet s, XDocument doc, List<string> sst, List<int> xf, List<Dxf> dxfs)
    {
        var root = doc.Root!;
        int StyleOf(int i) => i >= 0 && i < xf.Count ? xf[i] : 0;
        var pr = root.Element(M + "sheetPr");
        if (pr != null)
        {
            s.TabColor = Color(pr.Element(M + "tabColor"));
            if (IsTrue(pr.Element(M + "pageSetUpPr")?.Attribute("fitToPage"))) s.Page.FitToWidth = true;
        }
        var view = root.Element(M + "sheetViews")?.Element(M + "sheetView");
        if (view != null)
        {
            if ((string?)view.Attribute("showGridLines") == "0") s.ShowGridlines = false;
            if ((string?)view.Attribute("showRowColHeaders") == "0") s.ShowHeadings = false;
            if ((int?)view.Attribute("zoomScale") is int z && z > 0) s.Zoom = Math.Clamp(z / 100.0, 0.1, 4);
            var pane = view.Element(M + "pane");
            if (pane != null && (string?)pane.Attribute("state") is "frozen" or "frozenSplit")
            {
                s.FreezeCols = (int)((double?)pane.Attribute("xSplit") ?? 0);
                s.FreezeRows = (int)((double?)pane.Attribute("ySplit") ?? 0);
            }
            var sel = view.Elements(M + "selection").LastOrDefault();
            if (sel != null && Addr.TryParseA1((string?)sel.Attribute("activeCell") ?? "A1", out var ac)) { s.ActiveCell = ac; s.Selection = CellRange.Single(ac.Row, ac.Col); }
        }
        foreach (var col in root.Element(M + "cols")?.Elements(M + "col") ?? [])
        {
            int min = (int)col.Attribute("min")! - 1, max = Math.Min((int)col.Attribute("max")! - 1, Addr.MaxCols - 1);
            double? w = (double?)col.Attribute("width");
            bool hidden = IsTrue(col.Attribute("hidden"));
            int? style = (int?)col.Attribute("style");
            if (max - min > 2000 && !hidden && style is null or 0) continue; // "all remaining columns" ranges with default width
            for (int c = min; c <= max; c++)
            {
                if (w != null && Math.Abs(w.Value * 7 - Sheet.DefaultColWidth) > 1) s.ColWidths[c] = Math.Max(0, Math.Round(w.Value * 7));
                if (hidden) s.HiddenCols.Add(c);
                if (style is int st && st > 0 && max - min < 2000) s.ColStyles[c] = StyleOf(st);
            }
        }
        var shared = new Dictionary<string, (string Formula, int Row, int Col)>();
        var arrays = new List<(int Row, int Col, CellRange Area)>();
        foreach (var row in root.Element(M + "sheetData")?.Elements(M + "row") ?? [])
        {
            int r = ((int?)row.Attribute("r") ?? 1) - 1;
            if ((double?)row.Attribute("ht") is double ht)
            {
                if (IsTrue(row.Attribute("customHeight"))) s.RowHeights[r] = Math.Round(ht / 0.75);
                else if (Math.Abs(ht / 0.75 - Sheet.DefaultRowHeight) > 1) s.AutoHeights[r] = Math.Round(ht / 0.75);
            }
            if (IsTrue(row.Attribute("hidden"))) s.HiddenRows.Add(r);
            if (IsTrue(row.Attribute("customFormat")) && (int?)row.Attribute("s") is int rs) s.RowStyles[r] = StyleOf(rs);
            int nextCol = 0;
            foreach (var c in row.Elements(M + "c"))
            {
                int col;
                var refAttr = (string?)c.Attribute("r");
                if (refAttr != null && Addr.TryParseA1(refAttr, out var p)) { col = p.Col; r = p.Row; } else col = nextCol;
                nextCol = col + 1;
                var t = (string?)c.Attribute("t") ?? "n";
                int style = StyleOf((int?)c.Attribute("s") ?? 0);
                var vEl = c.Element(M + "v");
                var fEl = c.Element(M + "f");
                CellValue v = CellValue.Blank;
                switch (t)
                {
                    case "s":
                        if (vEl != null && int.TryParse(vEl.Value, out var si) && si >= 0 && si < sst.Count) v = CellValue.Text(sst[si]);
                        break;
                    case "inlineStr": v = CellValue.Text(RichText(c.Element(M + "is") ?? new XElement("x"))); break;
                    case "str": v = CellValue.Text(Unescape(vEl?.Value ?? "")); break;
                    case "b": v = CellValue.Bool(vEl?.Value == "1"); break;
                    case "e": v = vEl != null ? CellValue.Error(vEl.Value) : CellValue.Blank; break;
                    case "d":
                        if (vEl != null && DateTime.TryParse(vEl.Value, Inv, DateTimeStyles.RoundtripKind, out var dt)) v = CellValue.Number(DateSerial.ToSerial(dt));
                        break;
                    default:
                        if (vEl != null && double.TryParse(vEl.Value, NumberStyles.Float, Inv, out var d)) v = CellValue.Number(d);
                        break;
                }
                string? formula = null;
                if (fEl != null)
                {
                    var ft = (string?)fEl.Attribute("t");
                    var text = fEl.Value;
                    if (ft == "shared")
                    {
                        var si = (string?)fEl.Attribute("si") ?? "";
                        if (text.Length > 0) { shared[si] = (text, r, col); formula = text; }
                        else if (shared.TryGetValue(si, out var master)) formula = RefAdjust.Shift(master.Formula, r - master.Row, col - master.Col);
                    }
                    else if (text.Length > 0) formula = text;
                }
                if (formula == null && v.IsBlank && style == 0) continue;
                var cell = s.GetOrCreate(r, col);
                cell.Value = v;
                cell.Formula = formula == null ? null : FileFormula.FromFile(formula);
                cell.StyleId = style;
                if (formula != null)
                {
                    // cm = dynamic-array formula; t="array" = array formula over a range; anything else was written by an
                    // Excel without dynamic arrays (or keeps "@"): arrays reduce to one value there
                    bool isArray = (string?)fEl!.Attribute("t") == "array";
                    cell.Implicit = c.Attribute("cm") == null && !isArray;
                    if (isArray && Addr.TryParseRange((string?)fEl.Attribute("ref") ?? "", out var area) && !area.IsSingle && area.R1 == r && area.C1 == col)
                        arrays.Add((r, col, area));
                }
            }
        }
        foreach (var (ar, acol, area) in arrays)
        {
            // cached values of an array result belong to its formula (they are recomputed and re-spilled on load)
            var master = s.GetCell(ar, acol);
            if (master?.Formula == null) continue;
            master.SpillRows = area.Rows;
            master.SpillCols = area.Cols;
            for (int i = 0; i < area.Rows; i++)
                for (int j = 0; j < area.Cols; j++)
                {
                    if (i == 0 && j == 0) continue;
                    var sc = s.GetCell(ar + i, acol + j);
                    if (sc == null || sc.Formula != null) continue;
                    sc.SpillFrom = master;
                    sc.SpillDr = i;
                    sc.SpillDc = j;
                }
        }
        if (root.Element(M + "sheetProtection") is XElement prot && IsTrue(prot.Attribute("sheet")))
        {
            s.Protected = true;
            s.ProtectionPasswordHash = (string?)prot.Attribute("password");
        }
        foreach (var mc in root.Element(M + "mergeCells")?.Elements(M + "mergeCell") ?? [])
            if (Addr.TryParseRange((string?)mc.Attribute("ref") ?? "", out var mr) && !mr.IsSingle) s.Merges.Add(mr);
        var af = root.Element(M + "autoFilter");
        if (af != null && Addr.TryParseRange((string?)af.Attribute("ref") ?? "", out var afr))
        {
            s.Filter = new AutoFilter { Range = afr };
            foreach (var fc in af.Elements(M + "filterColumn"))
            {
                int colId = afr.C1 + ((int?)fc.Attribute("colId") ?? 0);
                var filters = fc.Element(M + "filters");
                if (filters != null)
                {
                    var set = new HashSet<string>(filters.Elements(M + "filter").Select(f => (string?)f.Attribute("val") ?? ""));
                    if (IsTrue(filters.Attribute("blank"))) set.Add("");
                    s.Filter.Allowed[colId] = set;
                }
                var cf = fc.Element(M + "customFilters")?.Element(M + "customFilter");
                if (cf != null)
                {
                    var op = (string?)cf.Attribute("operator") ?? "equal";
                    var val = (string?)cf.Attribute("val") ?? "";
                    s.Filter.Conditions[colId] = op switch
                    {
                        "greaterThan" => ">" + val, "greaterThanOrEqual" => ">=" + val, "lessThan" => "<" + val, "lessThanOrEqual" => "<=" + val, "notEqual" => "<>" + val,
                        _ => val.StartsWith('*') && val.EndsWith('*') && val.Length > 1 ? "contains:" + val.Trim('*') : val.EndsWith('*') ? "begins:" + val.TrimEnd('*') : val.StartsWith('*') ? "ends:" + val.TrimStart('*') : "=" + val
                    };
                }
            }
            SheetOps.ApplyFilter(s);
        }
        foreach (var cfEl in root.Elements(M + "conditionalFormatting"))
        {
            var sqref = ((string?)cfEl.Attribute("sqref") ?? "").Split(' ')[0];
            if (!Addr.TryParseRange(sqref, out var range)) continue;
            foreach (var rule in cfEl.Elements(M + "cfRule"))
            {
                var cf = new ConditionalFormat { Range = range };
                var type = (string?)rule.Attribute("type");
                var formulas = rule.Elements(M + "formula").Select(f => f.Value).ToList();
                string Val(string f) => f.StartsWith('"') && f.EndsWith('"') && f.Length >= 2 ? f[1..^1].Replace("\"\"", "\"") : f;
                switch (type)
                {
                    case "cellIs":
                        cf.Type = (string?)rule.Attribute("operator") switch { "greaterThan" or "greaterThanOrEqual" => CfType.GreaterThan, "lessThan" or "lessThanOrEqual" => CfType.LessThan, "between" => CfType.Between, _ => CfType.EqualTo };
                        if (formulas.Count > 0) cf.Value1 = Val(formulas[0]);
                        if (formulas.Count > 1) cf.Value2 = Val(formulas[1]);
                        break;
                    case "containsText": cf.Type = CfType.TextContains; cf.Value1 = (string?)rule.Attribute("text") ?? ""; break;
                    case "duplicateValues": cf.Type = CfType.Duplicate; break;
                    case "top10": cf.Type = IsTrue(rule.Attribute("bottom")) ? CfType.Bottom10 : CfType.Top10; cf.Rank = (int?)rule.Attribute("rank") ?? 10; break;
                    case "aboveAverage": cf.Type = (string?)rule.Attribute("aboveAverage") == "0" ? CfType.BelowAverage : CfType.AboveAverage; break;
                    case "expression": cf.Type = CfType.Expression; cf.Value1 = formulas.FirstOrDefault() ?? ""; break;
                    case "dataBar":
                        cf.Type = CfType.DataBar;
                        var dc = Color(rule.Element(M + "dataBar")?.Element(M + "color"));
                        if (dc != 0) cf.Color1 = dc;
                        break;
                    case "colorScale":
                    {
                        var cols = rule.Element(M + "colorScale")?.Elements(M + "color").Select(Color).ToList() ?? new();
                        cf.Type = cols.Count >= 3 ? CfType.ColorScale3 : CfType.ColorScale2;
                        if (cols.Count >= 3) { cf.Color1 = cols[0]; cf.Color2 = cols[1]; cf.Color3 = cols[2]; }
                        else if (cols.Count == 2) { cf.Color1 = cols[0]; cf.Color3 = cols[1]; }
                        break;
                    }
                    default: continue;
                }
                if ((int?)rule.Attribute("dxfId") is int dxfId && dxfId < dxfs.Count)
                {
                    cf.FontColor = dxfs[dxfId].FontColor; cf.Fill = dxfs[dxfId].Fill; cf.Bold = dxfs[dxfId].Bold;
                }
                s.CondFormats.Add(cf);
            }
        }
        foreach (var dvEl in root.Element(M + "dataValidations")?.Elements(M + "dataValidation") ?? [])
        {
            var sqref = ((string?)dvEl.Attribute("sqref") ?? "").Split(' ')[0];
            if (!Addr.TryParseRange(sqref, out var range)) continue;
            var dv = new DataValidation
            {
                Range = range,
                Type = (string?)dvEl.Attribute("type") switch { "list" => ValidationType.List, "whole" => ValidationType.WholeNumber, "decimal" => ValidationType.Decimal, "date" => ValidationType.Date, "textLength" => ValidationType.TextLength, _ => ValidationType.Any },
                Op = (string?)dvEl.Attribute("operator") switch
                {
                    "notBetween" => ValidationOp.NotBetween, "equal" => ValidationOp.Equal, "notEqual" => ValidationOp.NotEqual, "greaterThan" => ValidationOp.Greater,
                    "lessThan" => ValidationOp.Less, "greaterThanOrEqual" => ValidationOp.GreaterOrEqual, "lessThanOrEqual" => ValidationOp.LessOrEqual, _ => ValidationOp.Between
                },
                AllowBlank = IsTrue(dvEl.Attribute("allowBlank")),
                ErrorMessage = (string?)dvEl.Attribute("error") ?? ""
            };
            var f1 = dvEl.Element(M + "formula1")?.Value ?? "";
            if (dv.Type == ValidationType.List && f1.StartsWith('"') && f1.EndsWith('"') && f1.Length >= 2) dv.Formula1 = f1[1..^1].Replace("\"\"", "\"");
            else dv.Formula1 = f1.Length > 0 && dv.Type == ValidationType.List ? "=" + f1 : f1;
            dv.Formula2 = dvEl.Element(M + "formula2")?.Value ?? "";
            s.Validations.Add(dv);
        }
        var po = root.Element(M + "printOptions");
        if (po != null) { s.Page.PrintGridlines = IsTrue(po.Attribute("gridLines")); s.Page.PrintHeadings = IsTrue(po.Attribute("headings")); s.Page.CenterHorizontally = IsTrue(po.Attribute("horizontalCentered")); }
        var pm = root.Element(M + "pageMargins");
        if (pm != null)
        {
            s.Page.MarginLeft = (double?)pm.Attribute("left") ?? 0.7; s.Page.MarginRight = (double?)pm.Attribute("right") ?? 0.7;
            s.Page.MarginTop = (double?)pm.Attribute("top") ?? 0.75; s.Page.MarginBottom = (double?)pm.Attribute("bottom") ?? 0.75;
        }
        var ps = root.Element(M + "pageSetup");
        if (ps != null)
        {
            s.Page.PaperSize = (int?)ps.Attribute("paperSize") ?? 9;
            s.Page.Landscape = (string?)ps.Attribute("orientation") == "landscape";
            s.Page.Scale = (int?)ps.Attribute("scale") ?? 100;
            if ((int?)ps.Attribute("fitToWidth") == 1 && s.Page.FitToWidth) s.Page.FitToWidth = true;
        }
    }

    private static void ReadComments(Sheet s, XDocument? doc)
    {
        if (doc == null) return;
        foreach (var c in doc.Root!.Element(M + "commentList")?.Elements(M + "comment") ?? [])
        {
            if (!Addr.TryParseA1((string?)c.Attribute("ref") ?? "", out var p)) continue;
            var t = c.Element(M + "text");
            if (t == null) continue;
            s.Notes[Sheet.Key(p.Row, p.Col)] = RichText(t);
        }
    }

    private static void ReadDrawing(Sheet s, string part, Func<string, XDocument?> load, Func<string, byte[]?> loadBytes,
        Func<string, Dictionary<string, (string Target, string Type, bool External)>> rels)
    {
        var doc = load(part);
        if (doc == null) return;
        var drels = rels(part);
        foreach (var anchor in doc.Root!.Elements().Where(e => e.Name.LocalName is "twoCellAnchor" or "oneCellAnchor"))
        {
            RectangleF bounds = AnchorBounds(s, anchor);
            var gf = anchor.Element(X + "graphicFrame");
            var chartRef = gf?.Descendants(C + "chart").FirstOrDefault();
            if (chartRef != null)
            {
                var rid = (string?)chartRef.Attribute(R + "id");
                if (rid != null && drels.TryGetValue(rid, out var rel))
                {
                    var ch = ReadChart(s, load(rel.Target));
                    if (ch != null)
                    {
                        ch.Bounds = bounds;
                        ch.Name = (string?)gf!.Descendants(X + "cNvPr").FirstOrDefault()?.Attribute("name") ?? "Chart " + (s.Charts.Count + 1);
                        s.Charts.Add(ch);
                    }
                }
                continue;
            }
            var pic = anchor.Element(X + "pic");
            if (pic != null)
            {
                var blip = pic.Descendants(A + "blip").FirstOrDefault();
                var rid = (string?)blip?.Attribute(R + "embed");
                if (rid != null && drels.TryGetValue(rid, out var rel))
                {
                    var data = loadBytes(rel.Target);
                    if (data != null)
                        s.Pictures.Add(new Picture
                        {
                            Data = data, Ext = Path.GetExtension(rel.Target).TrimStart('.'), Bounds = bounds,
                            Name = (string?)pic.Descendants(X + "cNvPr").FirstOrDefault()?.Attribute("name") ?? "Picture " + (s.Pictures.Count + 1)
                        });
                }
            }
        }
    }

    private static RectangleF AnchorBounds(Sheet s, XElement anchor)
    {
        double ColX(int col, long off) { double x = 0; for (int c = 0; c < col; c++) x += s.ColWidth(c); return x + off / 9525.0; }
        double RowY(int row, long off) { double y = 0; for (int r = 0; r < row; r++) y += s.RowHeight(r); return y + off / 9525.0; }
        var from = anchor.Element(X + "from");
        if (from == null) return new RectangleF(100, 100, 480, 290);
        double x1 = ColX((int?)from.Element(X + "col") ?? 0, (long?)from.Element(X + "colOff") ?? 0);
        double y1 = RowY((int?)from.Element(X + "row") ?? 0, (long?)from.Element(X + "rowOff") ?? 0);
        var to = anchor.Element(X + "to");
        double x2, y2;
        if (to != null)
        {
            x2 = ColX((int?)to.Element(X + "col") ?? 0, (long?)to.Element(X + "colOff") ?? 0);
            y2 = RowY((int?)to.Element(X + "row") ?? 0, (long?)to.Element(X + "rowOff") ?? 0);
        }
        else
        {
            var ext = anchor.Element(X + "ext");
            x2 = x1 + ((long?)ext?.Attribute("cx") ?? 4572000) / 9525.0;
            y2 = y1 + ((long?)ext?.Attribute("cy") ?? 2743200) / 9525.0;
        }
        return new RectangleF((float)x1, (float)y1, (float)Math.Max(20, x2 - x1), (float)Math.Max(20, y2 - y1));
    }

    private static Chart? ReadChart(Sheet s, XDocument? doc)
    {
        if (doc == null) return null;
        var plot = doc.Descendants(C + "plotArea").FirstOrDefault();
        if (plot == null) return null;
        var ch = new Chart();
        var typeEl = plot.Elements().FirstOrDefault(e => e.Name.LocalName.EndsWith("Chart"));
        if (typeEl == null) return null;
        switch (typeEl.Name.LocalName)
        {
            case "barChart": case "bar3DChart":
                bool bar = (string?)typeEl.Element(C + "barDir")?.Attribute("val") == "bar";
                bool stacked = (string?)typeEl.Element(C + "grouping")?.Attribute("val") is "stacked" or "percentStacked";
                ch.Type = bar ? ChartType.Bar : stacked ? ChartType.StackedColumn : ChartType.Column;
                break;
            case "lineChart": case "line3DChart": ch.Type = ChartType.Line; break;
            case "pieChart": case "pie3DChart": ch.Type = ChartType.Pie; break;
            case "doughnutChart": ch.Type = ChartType.Doughnut; break;
            case "areaChart": case "area3DChart": ch.Type = ChartType.Area; break;
            case "scatterChart": ch.Type = ChartType.Scatter; break;
            default: ch.Type = ChartType.Column; break;
        }
        var title = doc.Descendants(C + "title").FirstOrDefault();
        if (title != null) ch.Title = string.Concat(title.Descendants(A + "t").Select(t => t.Value));
        ch.ShowLegend = doc.Descendants(C + "legend").Any();
        ch.ShowDataLabels = plot.Descendants(C + "showVal").Any(e => (string?)e.Attribute("val") == "1");
        // data range = union of category + series references on this sheet
        CellRange? union = null;
        foreach (var f in typeEl.Descendants(C + "f"))
        {
            var text = f.Value;
            var refPart = text.Contains('!') ? text[(text.LastIndexOf('!') + 1)..] : text;
            if (!Addr.TryParseRange(refPart, out var r)) continue;
            union = union == null ? r : union.Value.Union(r);
        }
        if (union == null) return null;
        ch.Data = union.Value;
        // series laid out in rows when the value refs are horizontal
        var firstVal = typeEl.Descendants().FirstOrDefault(e => e.Name.LocalName is "val" or "yVal")?.Descendants(C + "f").FirstOrDefault()?.Value;
        if (firstVal != null && Addr.TryParseRange(firstVal[(firstVal.LastIndexOf('!') + 1)..], out var fv) && fv.Rows == 1 && fv.Cols > 1) ch.SeriesInRows = true;
        return ch;
    }
}
