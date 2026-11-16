using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.CompilerServices;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.Layout;

/// <summary>
/// Measures text in points with GDI+ typographic metrics (no grid fitting) so layout is identical
/// on screen at any zoom, on the printer and in PDF.
/// </summary>
public sealed class TextMeasurer : IDisposable
{
    private readonly Bitmap _bmp = new(1, 1);
    private readonly Graphics _g;
    public static readonly StringFormat Format;
    private readonly Dictionary<(string, float, FontStyle), Font> _fonts = new();
    private readonly Dictionary<(Font, string), float> _widths = new();
    private readonly Dictionary<Font, (float Asc, float Desc, float Line)> _metrics = new();

    static TextMeasurer()
    {
        Format = (StringFormat)StringFormat.GenericTypographic.Clone();
        Format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces | StringFormatFlags.NoClip | StringFormatFlags.NoWrap;
        Format.Trimming = StringTrimming.None;
    }

    public TextMeasurer()
    {
        _g = Graphics.FromImage(_bmp);
        _g.PageUnit = GraphicsUnit.Point;
        _g.TextRenderingHint = TextRenderingHint.AntiAlias;
    }

    public static float EffectiveSize(CharProps p) => p.VertAlign == VertAlign.Baseline ? p.Size : p.Size * 0.65f;

    public Font GetFont(CharProps p)
    {
        var style = (p.Bold ? FontStyle.Bold : 0) | (p.Italic ? FontStyle.Italic : 0);
        var key = (p.FontName, EffectiveSize(p), style);
        if (_fonts.TryGetValue(key, out var f)) return f;
        try { f = new Font(p.FontName, Math.Max(1f, key.Item2), style, GraphicsUnit.Point); }
        catch { f = new Font("Times New Roman", Math.Max(1f, key.Item2), style, GraphicsUnit.Point); }
        if (!f.FontFamily.IsStyleAvailable(style))
        {
            try { f = new Font(f.FontFamily, f.Size, FontStyle.Regular, GraphicsUnit.Point); } catch { }
        }
        _fonts[key] = f;
        return f;
    }

    public float Width(string s, CharProps p)
    {
        if (s.Length == 0) return 0;
        if (p.AllCaps) s = s.ToUpperInvariant();
        var f = GetFont(p);
        var key = (f, s);
        if (_widths.TryGetValue(key, out var w)) return w;
        w = _g.MeasureString(s, f, PointF.Empty, Format).Width;
        if (_widths.Count > 200_000) _widths.Clear();
        _widths[key] = w;
        return w;
    }

    public (float Asc, float Desc, float Line) Metrics(CharProps p)
    {
        var f = GetFont(p);
        if (_metrics.TryGetValue(f, out var m)) return m;
        var fam = f.FontFamily;
        var st = f.Style;
        float em = fam.GetEmHeight(st);
        float asc = f.Size * fam.GetCellAscent(st) / em;
        float desc = f.Size * fam.GetCellDescent(st) / em;
        float line = f.Size * fam.GetLineSpacing(st) / em;
        m = (asc, desc, Math.Max(line, asc + desc));
        _metrics[f] = m;
        return m;
    }

    public void Dispose()
    {
        foreach (var f in _fonts.Values) f.Dispose();
        _g.Dispose();
        _bmp.Dispose();
    }
}

public enum SegKind { Text, Space, Tab, Object, Break }

public sealed class Seg
{
    public SegKind Kind;
    public int Start, End;          // paragraph offsets
    public float X, W;              // relative to the line's left (content box of the paragraph)
    public CharProps Props = CharProps.Default;
    public string Text = "";
    public InlineObject? Obj;
    public float Asc, Desc;
    public bool Leader;             // dotted leader (TOC tab)
    internal float[]? CharX;

    /// <summary>X of the boundary before character i (0..Length) relative to the line.</summary>
    public float XAt(TextMeasurer m, int offset)
    {
        int i = Math.Clamp(offset - Start, 0, End - Start);
        if (i == 0) return X;
        if (i >= End - Start) return X + W;
        if (Kind != SegKind.Text && Kind != SegKind.Space) return X + W;
        if (CharX == null)
        {
            CharX = new float[Text.Length + 1];
            float natural = m.Width(Text, Props);
            float scale = natural > 0 ? W / natural : 1;
            for (int k = 1; k <= Text.Length; k++) CharX[k] = m.Width(Text[..k], Props) * scale;
        }
        return X + CharX[i];
    }
}

public sealed class LineBox
{
    public Paragraph Para = null!;
    public int Start, End;
    public float Height, Baseline;  // Baseline from the top of the line
    public List<Seg> Segs = new();
    public string? Label;           // list number / bullet (first line)
    public float LabelX;
    public CharProps LabelProps = CharProps.Default;
    public bool PageBreakAfter;
    public bool IsFirst, IsLast;
    public float Left, Right;       // available box inside paragraph indents (relative to the content area)
}

public sealed class ParaLayout
{
    public Paragraph Para = null!;
    public List<LineBox> Lines = new();
    public float SpaceBefore, SpaceAfter;
    public int Version;
    public float Width;
}

public sealed class PlacedLine
{
    public LineBox Line = null!;
    public int Page;
    public float X, Y;              // absolute page coordinates (points) of the line box' left/top (X = content left)
    public bool Header, Footer;
    public RectangleF Clip;         // empty = no clipping
}

public sealed class PlacedCell
{
    public RectangleF Rect;
    public TableCell Cell = null!;
    public Table Table = null!;
    public bool DrawTop = true, DrawBottom = true;
}

public sealed class PageLayout
{
    public int Index;
    public List<PlacedLine> Lines = new();
    public List<PlacedCell> Cells = new();
    public List<PlacedLine> HeaderLines = new();
    public List<PlacedLine> FooterLines = new();
}

/// <summary>Lays out a whole document into pages. Paragraph layouts are cached by paragraph version and width.</summary>
public sealed class DocLayout
{
    public const float CellPad = 5.4f;
    public Document Doc { get; }
    public List<PageLayout> Pages { get; } = new();
    public List<PlacedLine> AllLines { get; } = new();
    public Dictionary<Paragraph, List<PlacedLine>> ByPara { get; } = new(ReferenceEqualityComparer.Instance as IEqualityComparer<Paragraph> ?? EqualityComparer<Paragraph>.Default);
    public Dictionary<Paragraph, string> ListLabels { get; } = new();
    public TextMeasurer M { get; }

    private static readonly ConditionalWeakTable<Paragraph, ParaLayout> Cache = new();

    public DocLayout(Document d, TextMeasurer m)
    {
        Doc = d; M = m;
        ComputeLabels();
        Paginate();
    }

    // ------------------------------------------------------------------ list numbering

    private void ComputeLabels()
    {
        foreach (var kv in ComputeListLabels(Doc)) ListLabels[kv.Key] = kv.Value;
    }

    /// <summary>List numbers / bullets of every numbered paragraph (body, then header and footer).</summary>
    public static Dictionary<Paragraph, string> ComputeListLabels(Document doc)
    {
        var result = new Dictionary<Paragraph, string>();
        var counters = new Dictionary<int, int[]>();
        foreach (var p in doc.Paragraphs.Concat(doc.Header).Concat(doc.Footer))
        {
            var pp = p.Props;
            if (pp.List == ListType.None) continue;
            if (!counters.TryGetValue(pp.ListId, out var c)) counters[pp.ListId] = c = new int[9];
            int lvl = Math.Clamp(pp.ListLevel, 0, 8);
            c[lvl]++;
            for (int k = lvl + 1; k < 9; k++) c[k] = 0;
            result[p] = Label(pp.List, lvl, c[lvl]);
        }
        return result;
    }

    public static string Label(ListType t, int level, int n) => t switch
    {
        ListType.Bullet => (level % 3) switch { 0 => "•", 1 => "◦", _ => "▪" },
        ListType.Decimal => n + ".",
        ListType.LowerAlpha => Alpha(n).ToLowerInvariant() + ")",
        ListType.UpperAlpha => Alpha(n) + ".",
        ListType.LowerRoman => Roman(n).ToLowerInvariant() + ".",
        ListType.UpperRoman => Roman(n) + ".",
        _ => ""
    };

    private static string Alpha(int n) { var s = ""; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }

    public static string Roman(int n)
    {
        if (n <= 0 || n > 3999) return n.ToString();
        var vals = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var syms = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var s = "";
        for (int i = 0; i < vals.Length; i++) while (n >= vals[i]) { s += syms[i]; n -= vals[i]; }
        return s;
    }

    // ------------------------------------------------------------------ paragraph layout

    public ParaLayout LayoutParagraph(Paragraph p, float width)
    {
        if (Cache.TryGetValue(p, out var cached) && cached.Version == p.Version && Math.Abs(cached.Width - width) < 0.01f && cached.Para == p)
        {
            ApplyLabel(cached);
            return cached;
        }
        var pl = BuildParagraph(p, width);
        Cache.AddOrUpdate(p, pl);
        ApplyLabel(pl);
        return pl;
    }

    private void ApplyLabel(ParaLayout pl)
    {
        if (pl.Lines.Count == 0) return;
        var first = pl.Lines[0];
        if (ListLabels.TryGetValue(pl.Para, out var label))
        {
            first.Label = label;
            var pp = pl.Para.Props;
            first.LabelX = Math.Max(0, pp.IndentLeft + pp.FirstLine);
            first.LabelProps = pl.Para.Runs.Count > 0 ? pl.Para.Runs[0].Props with { Underline = UnderlineStyle.None, Link = null, Highlight = null } : pl.Para.MarkProps;
        }
        else first.Label = null;
    }

    private sealed class Tok
    {
        public SegKind Kind;
        public int Start, End;
        public string Text = "";
        public CharProps Props = CharProps.Default;
        public InlineObject? Obj;
        public float W;
        public bool GlueBefore;   // no break allowed before this token (word continues across runs)
        public bool PageBreak;
    }

    private ParaLayout BuildParagraph(Paragraph p, float width)
    {
        var pp = p.Props;
        var pl = new ParaLayout { Para = p, Version = p.Version, Width = width, SpaceBefore = pp.SpaceBefore, SpaceAfter = pp.SpaceAfter };
        var toks = Tokenize(p);
        bool isList = pp.List != ListType.None;
        float right = width - pp.IndentRight;
        float firstStart = isList ? pp.IndentLeft : pp.IndentLeft + pp.FirstLine;
        float lineStart = pp.IndentLeft;
        int lastTab = pp.RightTabLeader ? toks.FindLastIndex(t => t.Kind == SegKind.Tab) : -1;

        var line = NewLine(p, true, Math.Max(0, firstStart), right);
        float x = line.Left;
        int i = 0;
        while (i < toks.Count)
        {
            var t = toks[i];
            if (t.Kind == SegKind.Break)
            {
                AddSeg(line, t, x, 0);
                line.PageBreakAfter = t.PageBreak;
                Finish(pl, line, pp, false);
                line = NewLine(p, false, Math.Max(0, lineStart), right);
                x = line.Left;
                i++;
                continue;
            }
            if (t.Kind == SegKind.Tab)
            {
                float tx;
                if (i == lastTab)
                {
                    float rest = 0;
                    for (int k = i + 1; k < toks.Count && toks[k].Kind != SegKind.Break; k++) rest += toks[k].W;
                    tx = Math.Max(x + 4, right - rest);
                    var seg = AddSeg(line, t, x, tx - x);
                    seg.Leader = true;
                }
                else
                {
                    tx = line.IsFirst && isList && x <= pp.IndentLeft + 0.5f && pp.IndentLeft > x ? pp.IndentLeft : (MathF.Floor((x + 0.01f) / 36f) + 1) * 36f;
                    if (tx > right && line.Segs.Count > 0)
                    {
                        Finish(pl, line, pp, true);
                        line = NewLine(p, false, Math.Max(0, lineStart), right);
                        x = line.Left;
                        tx = (MathF.Floor((x + 0.01f) / 36f) + 1) * 36f;
                    }
                    AddSeg(line, t, x, Math.Min(tx, right) - x);
                }
                x = line.Segs[^1].X + line.Segs[^1].W;
                i++;
                continue;
            }
            if (t.Kind == SegKind.Space)
            {
                AddSeg(line, t, x, t.W);   // spaces may hang over the right edge
                x += t.W;
                i++;
                continue;
            }
            // word cluster (text/object tokens glued together)
            int j = i + 1;
            float cw = t.W;
            while (j < toks.Count && toks[j].GlueBefore && toks[j].Kind is SegKind.Text) { cw += toks[j].W; j++; }
            if (x + cw > right + 0.01f && HasContent(line))
            {
                Finish(pl, line, pp, true);
                line = NewLine(p, false, Math.Max(0, lineStart), right);
                x = line.Left;
            }
            if (x + cw > right + 0.01f && !HasContent(line) && t.Kind == SegKind.Text)
            {
                // a single word longer than the line: break it by characters
                for (int k = i; k < j; k++)
                {
                    var tk = toks[k];
                    int s = 0;
                    while (s < tk.Text.Length)
                    {
                        int e = s + 1;
                        while (e < tk.Text.Length && x + M.Width(tk.Text[s..(e + 1)], tk.Props) <= right) e++;
                        var piece = new Tok { Kind = SegKind.Text, Start = tk.Start + s, End = tk.Start + e, Text = tk.Text[s..e], Props = tk.Props };
                        piece.W = M.Width(piece.Text, piece.Props);
                        if (x + piece.W > right + 0.01f && HasContent(line))
                        {
                            Finish(pl, line, pp, true);
                            line = NewLine(p, false, Math.Max(0, lineStart), right);
                            x = line.Left;
                        }
                        AddSeg(line, piece, x, piece.W);
                        x += piece.W;
                        s = e;
                    }
                }
                i = j;
                continue;
            }
            for (int k = i; k < j; k++) { AddSeg(line, toks[k], x, toks[k].W); x += toks[k].W; }
            i = j;
        }
        Finish(pl, line, pp, false);
        return pl;
    }

    private static bool HasContent(LineBox l) => l.Segs.Any(s => s.Kind != SegKind.Space || true) && l.Segs.Count > 0;

    private LineBox NewLine(Paragraph p, bool first, float left, float right) => new() { Para = p, IsFirst = first, Left = left, Right = right };

    private Seg AddSeg(LineBox line, Tok t, float x, float w)
    {
        var (asc, desc, _) = M.Metrics(t.Props);
        if (t.Obj is ImageObject img) { asc = img.Height; desc = 0; }
        var s = new Seg { Kind = t.Kind, Start = t.Start, End = t.End, X = x, W = w, Props = t.Props, Text = t.Text, Obj = t.Obj, Asc = asc, Desc = desc };
        if (t.Props.VertAlign == VertAlign.Superscript) { s.Asc += t.Props.Size * 0.33f; }
        else if (t.Props.VertAlign == VertAlign.Subscript) { s.Desc += t.Props.Size * 0.15f; }
        line.Segs.Add(s);
        return s;
    }

    private void Finish(ParaLayout pl, LineBox line, ParaProps pp, bool wrapped)
    {
        var p = pl.Para;
        line.Start = line.Segs.Count > 0 ? line.Segs[0].Start : (pl.Lines.Count > 0 ? pl.Lines[^1].End : 0);
        line.End = line.Segs.Count > 0 ? line.Segs[^1].End : line.Start;
        float asc, desc, natural;
        var mp = M.Metrics(p.MarkProps);
        if (line.Segs.Count == 0 || line.Segs.All(s => s.Kind == SegKind.Break))
        {
            var props = line.Segs.Count > 0 ? line.Segs[0].Props : (p.Runs.Count > 0 && line.Start > 0 ? p.PropsAt(line.Start) : p.MarkProps);
            var m = M.Metrics(props);
            asc = m.Asc; desc = m.Desc; natural = m.Line;
        }
        else
        {
            asc = line.Segs.Max(s => s.Asc);
            desc = line.Segs.Max(s => s.Desc);
            float textLine = line.Segs.Where(s => s.Obj is not ImageObject).Select(s => M.Metrics(s.Props).Line).DefaultIfEmpty(asc + desc).Max();
            natural = Math.Max(asc + desc, textLine);
            if (line.Segs.Any(s => s.Obj is ImageObject)) natural = Math.Max(natural, asc + desc);
        }
        float h = pp.LineRule switch
        {
            LineRule.Exact => pp.LineSpacing,
            LineRule.AtLeast => Math.Max(natural, pp.LineSpacing),
            _ => natural * Math.Max(0.5f, pp.LineSpacing) + (line.Segs.Any(s => s.Obj is ImageObject) ? 0 : 0)
        };
        if (pp.LineRule == LineRule.Multiple && line.Segs.Any(s => s.Obj is ImageObject)) h = Math.Max(h, asc + desc + desc);
        line.Height = h;
        line.Baseline = pp.LineRule == LineRule.Exact ? h * 0.8f : asc + Math.Max(0, (Math.Min(h, natural) - (asc + desc)) / 2);
        if (pp.LineRule == LineRule.Multiple && pp.LineSpacing < 1) line.Baseline = Math.Min(asc, h - desc * 0.5f);

        // alignment
        float contentEnd = line.Segs.Where(s => s.Kind != SegKind.Space && s.Kind != SegKind.Break).Select(s => s.X + s.W).DefaultIfEmpty(line.Left).Max();
        float extra = line.Right - contentEnd;
        if (extra > 0 && line.Segs.Count > 0)
        {
            switch (pp.Align)
            {
                case ParaAlign.Center: Shift(line, extra / 2); break;
                case ParaAlign.Right: Shift(line, extra); break;
                case ParaAlign.Justify when wrapped && !line.Segs.Any(s => s.Kind == SegKind.Tab):
                {
                    int lastNonSpace = line.Segs.FindLastIndex(s => s.Kind != SegKind.Space && s.Kind != SegKind.Break);
                    var spaces = line.Segs.Take(lastNonSpace + 1).Where(s => s.Kind == SegKind.Space).ToList();
                    if (spaces.Count > 0)
                    {
                        float add = extra / spaces.Count, shift = 0;
                        foreach (var s in line.Segs)
                        {
                            s.X += shift;
                            if (s.Kind == SegKind.Space && spaces.Contains(s)) { s.W += add; shift += add; s.CharX = null; }
                        }
                    }
                    break;
                }
            }
        }
        line.IsLast = !wrapped;
        pl.Lines.Add(line);
    }

    private static void Shift(LineBox l, float dx) { foreach (var s in l.Segs) s.X += dx; }

    private List<Tok> Tokenize(Paragraph p)
    {
        var toks = new List<Tok>();
        int pos = 0;
        bool prevWord = false;
        foreach (var r in p.Runs)
        {
            if (r.Object != null)
            {
                var t = new Tok { Kind = SegKind.Object, Start = pos, End = pos + 1, Text = Run.ObjectChar.ToString(), Props = r.Props, Obj = r.Object };
                t.W = r.Object switch
                {
                    ImageObject img => img.Width,
                    FieldObject f => M.Width(f.Kind == FieldKind.Date ? DateTime.Now.ToString("dd/MM/yyyy") : "00", r.Props),
                    _ => 10
                };
                toks.Add(t);
                pos++;
                prevWord = false;
                continue;
            }
            var s = r.Text;
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '\t') { toks.Add(new Tok { Kind = SegKind.Tab, Start = pos + i, End = pos + i + 1, Text = "\t", Props = r.Props }); i++; prevWord = false; continue; }
                if (c == Run.LineBreak || c == Run.PageBreak || c == '\n' || c == '\r')
                {
                    toks.Add(new Tok { Kind = SegKind.Break, Start = pos + i, End = pos + i + 1, Text = c.ToString(), Props = r.Props, PageBreak = c == Run.PageBreak });
                    i++; prevWord = false; continue;
                }
                int j = i;
                if (c == ' ' || c == ' ' && false)
                {
                    while (j < s.Length && s[j] == ' ') j++;
                    var t = new Tok { Kind = SegKind.Space, Start = pos + i, End = pos + j, Text = s[i..j], Props = r.Props };
                    t.W = M.Width(t.Text, r.Props);
                    toks.Add(t);
                    prevWord = false;
                }
                else
                {
                    while (j < s.Length && s[j] != ' ' && s[j] != '\t' && s[j] != Run.LineBreak && s[j] != Run.PageBreak && s[j] != '\n' && s[j] != '\r')
                    {
                        j++;
                        // allow a break after a hyphen / slash inside long words
                        if ((s[j - 1] == '-' || s[j - 1] == '/') && j - i > 1) break;
                    }
                    var t = new Tok { Kind = SegKind.Text, Start = pos + i, End = pos + j, Text = s[i..j], Props = r.Props, GlueBefore = prevWord };
                    t.W = M.Width(t.Text, r.Props);
                    toks.Add(t);
                    prevWord = !(s[j - 1] == '-' || s[j - 1] == '/');
                }
                i = j;
            }
            pos += s.Length;
        }
        return toks;
    }

    // ------------------------------------------------------------------ pagination

    private PageLayout NewPage()
    {
        var pg = new PageLayout { Index = Pages.Count };
        Pages.Add(pg);
        return pg;
    }

    private void Place(PageLayout pg, LineBox l, float x, float y, RectangleF clip = default)
    {
        var pl = new PlacedLine { Line = l, Page = pg.Index, X = x, Y = y, Clip = clip };
        pg.Lines.Add(pl);
        AllLines.Add(pl);
        if (!ByPara.TryGetValue(l.Para, out var list)) ByPara[l.Para] = list = new();
        list.Add(pl);
    }

    private void Paginate()
    {
        var ps = Doc.Page;
        float top = ps.MarginTop, bottom = ps.Height - ps.MarginBottom, left = ps.MarginLeft, width = ps.ContentWidth;
        // keep header/footer clear of the body
        var pg = NewPage();
        float y = top;
        foreach (var block in Doc.Blocks)
        {
            if (block is Paragraph p)
            {
                var pl = LayoutParagraph(p, width);
                if (p.Props.PageBreakBefore && y > top + 0.5f) { pg = NewPage(); y = top; }
                y += pl.SpaceBefore;
                for (int li = 0; li < pl.Lines.Count; li++)
                {
                    var line = pl.Lines[li];
                    if (y + line.Height > bottom + 0.01f && y > top + pl.SpaceBefore + 0.5f)
                    {
                        pg = NewPage(); y = top;
                    }
                    Place(pg, line, left, y);
                    y += line.Height;
                    if (line.PageBreakAfter) { pg = NewPage(); y = top; }
                }
                y += pl.SpaceAfter;
            }
            else if (block is Table t)
            {
                y = PlaceTable(t, ref pg, y, top, bottom, left, width);
            }
        }
        LayoutHeaderFooter();
    }

    private float PlaceTable(Table t, ref PageLayout pg, float y, float top, float bottom, float left, float width)
    {
        var cols = t.ColumnWidths.Count > 0 ? t.ColumnWidths : [width];
        float total = cols.Sum();
        float scale = total > width * 1.5f ? width / total : 1f;
        float tableLeft = left + t.Align switch { ParaAlign.Center => (width - total * scale) / 2, ParaAlign.Right => width - total * scale, _ => 0 };
        foreach (var row in t.Rows)
        {
            row.Table = t;
            var cellInfo = new List<(TableCell Cell, float X, float W, List<ParaLayout> Paras, float H)>();
            int g = 0;
            foreach (var cell in row.Cells)
            {
                cell.Row = row;
                int span = Math.Max(1, cell.GridSpan);
                float cx = tableLeft + cols.Take(g).Sum() * scale;
                float cw = cols.Skip(g).Take(span).Sum() * scale;
                if (cw <= 0) cw = 36;
                g += span;
                var paras = new List<ParaLayout>();
                float h = 0;
                if (cell.VMerge != VMerge.Continue)
                {
                    foreach (var cp in cell.Paragraphs)
                    {
                        var pl = LayoutParagraph(cp, Math.Max(10, cw - 2 * CellPad));
                        paras.Add(pl);
                        h += pl.SpaceBefore + pl.Lines.Sum(l => l.Height) + pl.SpaceAfter;
                    }
                }
                else
                {
                    foreach (var cp in cell.Paragraphs) paras.Add(LayoutParagraph(cp, Math.Max(10, cw - 2 * CellPad)));
                }
                cellInfo.Add((cell, cx, cw, paras, h));
            }
            float rowH = Math.Max(row.MinHeight, cellInfo.Where(c => c.Cell.VMerge != VMerge.Continue).Select(c => c.H).DefaultIfEmpty(0).Max() + 2 * 2.5f);
            rowH = Math.Max(rowH, 14f);
            if (y + rowH > bottom + 0.01f && y > top + 0.5f) { pg = NewPage(); y = top; }
            float clipBottom = Math.Min(y + rowH, pg.Index >= 0 ? bottom + 1000 : bottom);
            foreach (var (cell, cx, cw, paras, h) in cellInfo)
            {
                var rect = new RectangleF(cx, y, cw, rowH);
                pg.Cells.Add(new PlacedCell { Rect = rect, Cell = cell, Table = t, DrawTop = cell.VMerge != VMerge.Continue, DrawBottom = !(NextRowContinues(t, row, cell)) });
                float cy = y + 2.5f;
                if (cell.VAlign != CellVAlign.Top && cell.VMerge != VMerge.Continue)
                {
                    float free = rowH - 5f - h;
                    if (free > 0) cy += cell.VAlign == CellVAlign.Center ? free / 2 : free;
                }
                foreach (var pl in paras)
                {
                    cy += cell.VMerge == VMerge.Continue ? 0 : pl.SpaceBefore;
                    foreach (var line in pl.Lines)
                    {
                        Place(pg, line, cx + CellPad, cy, rect);
                        if (cell.VMerge != VMerge.Continue) cy += line.Height;
                    }
                    if (cell.VMerge != VMerge.Continue) cy += pl.SpaceAfter;
                }
            }
            y += rowH;
        }
        return y + 2;
    }

    private static bool NextRowContinues(Table t, TableRow row, TableCell cell)
    {
        int ri = t.Rows.IndexOf(row);
        if (ri < 0 || ri + 1 >= t.Rows.Count) return false;
        int g = t.GridColumnOf(row, cell);
        var next = t.Rows[ri + 1];
        int col = 0;
        foreach (var c in next.Cells)
        {
            if (col == g) return c.VMerge == VMerge.Continue;
            col += Math.Max(1, c.GridSpan);
        }
        return false;
    }

    private void LayoutHeaderFooter()
    {
        var ps = Doc.Page;
        float width = ps.ContentWidth;
        var hl = Doc.Header.Select(p => LayoutParagraph(p, width)).ToList();
        var fl = Doc.Footer.Select(p => LayoutParagraph(p, width)).ToList();
        float fh = fl.Sum(p => p.SpaceBefore + p.Lines.Sum(l => l.Height) + p.SpaceAfter);
        foreach (var pg in Pages)
        {
            float y = ps.HeaderDistance;
            foreach (var pl in hl)
            {
                y += pl.SpaceBefore;
                foreach (var l in pl.Lines) { pg.HeaderLines.Add(new PlacedLine { Line = l, Page = pg.Index, X = ps.MarginLeft, Y = y, Header = true }); y += l.Height; }
                y += pl.SpaceAfter;
            }
            y = ps.Height - ps.FooterDistance - fh;
            foreach (var pl in fl)
            {
                y += pl.SpaceBefore;
                foreach (var l in pl.Lines) { pg.FooterLines.Add(new PlacedLine { Line = l, Page = pg.Index, X = ps.MarginLeft, Y = y, Footer = true }); y += l.Height; }
                y += pl.SpaceAfter;
            }
        }
    }

    // ------------------------------------------------------------------ geometry queries

    /// <summary>Caret rectangle (page, x, top, height) for a position.</summary>
    public (int Page, float X, float Y, float H) CaretAt(DocPos pos)
    {
        var ps = Doc.Paragraphs;
        pos = Editing.Clamp(Doc, pos);
        var para = ps[pos.Para];
        if (!ByPara.TryGetValue(para, out var lines) || lines.Count == 0) return (0, Doc.Page.MarginLeft, Doc.Page.MarginTop, 14);
        PlacedLine pl = lines[^1];
        foreach (var l in lines)
        {
            if (pos.Offset >= l.Line.Start && (pos.Offset < l.Line.End || l == lines[^1] || (pos.Offset == l.Line.End && l.Line.Segs.Count > 0 && l.Line.Segs[^1].Kind != SegKind.Break && lines.IndexOf(l) == lines.Count - 1)))
            { pl = l; break; }
        }
        float x = XInLine(pl.Line, pos.Offset);
        var props = para.Length == 0 ? para.MarkProps : para.PropsAt(Math.Max(1, pos.Offset));
        var m = M.Metrics(props);
        float h = Math.Min(pl.Line.Height, m.Asc + m.Desc);
        float y = pl.Y + pl.Line.Baseline - m.Asc;
        return (pl.Page, pl.X + x, y, Math.Max(h, 6));
    }

    public float XInLine(LineBox l, int offset)
    {
        if (l.Segs.Count == 0) return LineStartX(l);
        foreach (var s in l.Segs)
        {
            if (offset < s.End || (offset == s.End && s == l.Segs[^1] && s.Kind != SegKind.Break))
                return offset <= s.Start ? s.X : s.XAt(M, offset);
        }
        var last = l.Segs[^1];
        return last.Kind == SegKind.Break ? last.X : last.X + last.W;
    }

    private static float LineStartX(LineBox l)
    {
        var pp = l.Para.Props;
        float x = l.Left;
        if (pp.Align == ParaAlign.Center) x = (l.Left + l.Right) / 2;
        else if (pp.Align == ParaAlign.Right) x = l.Right;
        return x;
    }

    /// <summary>Offset in a line nearest to x (relative to the line).</summary>
    public int OffsetInLine(LineBox l, float x)
    {
        if (l.Segs.Count == 0) return l.Start;
        foreach (var s in l.Segs)
        {
            if (x < s.X + s.W || s == l.Segs[^1])
            {
                if (s.Kind == SegKind.Break) return s.Start;
                if (x <= s.X) return s.Start;
                if (s.Kind is SegKind.Object or SegKind.Tab) return x < s.X + s.W / 2 ? s.Start : s.End;
                int best = s.Start; float bd = float.MaxValue;
                for (int o = s.Start; o <= s.End; o++)
                {
                    float d = Math.Abs(s.XAt(M, o) - x);
                    if (d < bd) { bd = d; best = o; }
                }
                // do not place the caret after trailing spaces of a wrapped line
                if (s == l.Segs[^1] && best == s.End && !l.IsLast && s.Kind == SegKind.Space) best = Math.Max(s.Start, s.End - 1);
                return best;
            }
        }
        return l.End;
    }

    /// <summary>Hit test on a page (points). Returns null outside the body.</summary>
    public DocPos HitTest(int page, PointF pt)
    {
        page = Math.Clamp(page, 0, Pages.Count - 1);
        var lines = Pages[page].Lines;
        if (lines.Count == 0)
        {
            // empty page: nearest line on previous pages
            var prev = AllLines.LastOrDefault(l => l.Page < page) ?? AllLines.FirstOrDefault();
            if (prev == null) return new DocPos(0, 0);
            return PosOf(prev, prev.Line.End);
        }
        // prefer a cell containing the point
        var cell = Pages[page].Cells.FirstOrDefault(c => c.Rect.Contains(pt) && c.Cell.VMerge != VMerge.Continue);
        IEnumerable<PlacedLine> cand = cell != null ? lines.Where(l => l.Line.Para.Cell == cell.Cell) : lines.Where(l => l.Clip.IsEmpty || l.Clip.Contains(pt) || true);
        if (cell != null && !cand.Any()) cand = lines;
        PlacedLine? best = null; float bd = float.MaxValue;
        foreach (var l in cand)
        {
            float dy = pt.Y < l.Y ? l.Y - pt.Y : pt.Y > l.Y + l.Line.Height ? pt.Y - (l.Y + l.Line.Height) : 0;
            float dx = 0;
            if (!l.Clip.IsEmpty) dx = pt.X < l.Clip.Left ? l.Clip.Left - pt.X : pt.X > l.Clip.Right ? pt.X - l.Clip.Right : 0;
            float d = dy * 1000 + dx;
            if (d < bd) { bd = d; best = l; }
        }
        best ??= lines[0];
        int off = OffsetInLine(best.Line, pt.X - best.X);
        return PosOf(best, off);
    }

    public DocPos PosOf(PlacedLine l, int offset)
    {
        int pi = Doc.IndexOf(l.Line.Para);
        return new DocPos(Math.Max(0, pi), offset);
    }

    public PlacedLine? LineOf(DocPos pos)
    {
        var ps = Doc.Paragraphs;
        if (pos.Para < 0 || pos.Para >= ps.Count) return null;
        if (!ByPara.TryGetValue(ps[pos.Para], out var lines) || lines.Count == 0) return null;
        foreach (var l in lines) if (pos.Offset < l.Line.End || l == lines[^1]) return l;
        return lines[^1];
    }

    /// <summary>Position on the visual line above/below, keeping the desired absolute x (page points).</summary>
    public DocPos? Vertical(DocPos pos, float desiredX, int dir)
    {
        var cur = LineOf(pos);
        if (cur == null) return null;
        for (int page = cur.Page; page >= 0 && page < Pages.Count; page += dir)
        {
            var lines = Pages[page].Lines;
            var list = (page == cur.Page ? lines.Where(l => dir > 0 ? l.Y > cur.Y + 0.5f : l.Y < cur.Y - 0.5f) : lines).ToList();
            if (list.Count == 0) continue;
            float targetY = dir > 0 ? list.Min(l => l.Y) : list.Max(l => l.Y);
            var row = list.Where(l => Math.Abs(l.Y - targetY) < 0.5f).ToList();
            var t = row.FirstOrDefault(l => l.Clip.IsEmpty || (l.Clip.Left <= desiredX && desiredX <= l.Clip.Right)) ?? row[0];
            return PosOf(t, OffsetInLine(t.Line, desiredX - t.X));
        }
        return null;
    }

    public int PageOf(DocPos pos) => LineOf(pos)?.Page ?? 0;
}
