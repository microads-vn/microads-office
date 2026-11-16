using System.Drawing;

namespace MicroadsOffice.Word.Model;

public enum UnderlineStyle { None, Single, Double, Dotted, Thick, Wavy }
public enum VertAlign { Baseline, Superscript, Subscript }
public enum ParaAlign { Left, Center, Right, Justify }
public enum LineRule { Multiple, Exact, AtLeast }
public enum ListType { None, Bullet, Decimal, LowerAlpha, UpperAlpha, LowerRoman, UpperRoman }
public enum CellVAlign { Top, Center, Bottom }
public enum VMerge { None, Restart, Continue }

/// <summary>Character formatting (immutable; use <c>with</c>). Value equality lets adjacent runs merge.</summary>
public sealed record CharProps
{
    public string FontName { get; init; } = "Times New Roman";
    public float Size { get; init; } = 13f;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public UnderlineStyle Underline { get; init; }
    public bool Strike { get; init; }
    public VertAlign VertAlign { get; init; }
    /// <summary>Null = automatic (black).</summary>
    public Color? Color { get; init; }
    public Color? Highlight { get; init; }
    public string? Link { get; init; }
    public bool AllCaps { get; init; }

    public static readonly CharProps Default = new();

    public FontStyle FontStyle =>
        (Bold ? FontStyle.Bold : 0) | (Italic ? FontStyle.Italic : 0) |
        (Underline != UnderlineStyle.None ? FontStyle.Underline : 0) | (Strike ? FontStyle.Strikeout : 0);

    public bool Equals(CharProps? o) => o is not null && FontName == o.FontName && Size == o.Size && Bold == o.Bold && Italic == o.Italic &&
        Underline == o.Underline && Strike == o.Strike && VertAlign == o.VertAlign && ArgbEq(Color, o.Color) && ArgbEq(Highlight, o.Highlight) && Link == o.Link && AllCaps == o.AllCaps;

    public override int GetHashCode() => HashCode.Combine(FontName, Size, Bold, Italic, Underline, Strike, VertAlign, Link);

    private static bool ArgbEq(Color? a, Color? b) => a.HasValue == b.HasValue && (!a.HasValue || a.Value.ToArgb() == b!.Value.ToArgb());
}

/// <summary>Paragraph formatting (immutable). Lengths are in points.</summary>
public sealed record ParaProps
{
    public string Style { get; init; } = StyleIds.Normal;
    public ParaAlign Align { get; init; } = ParaAlign.Left;
    public float IndentLeft { get; init; }
    public float IndentRight { get; init; }
    /// <summary>First line offset; negative = hanging indent.</summary>
    public float FirstLine { get; init; }
    public float SpaceBefore { get; init; }
    public float SpaceAfter { get; init; } = 6f;
    public float LineSpacing { get; init; } = 1.15f;
    public LineRule LineRule { get; init; } = LineRule.Multiple;
    public ListType List { get; init; }
    public int ListLevel { get; init; }
    public int ListId { get; init; }
    public bool KeepWithNext { get; init; }
    public bool PageBreakBefore { get; init; }
    public Color? Shading { get; init; }
    public bool Border { get; init; }
    /// <summary>The last tab jumps to the right indent with a dotted leader (table of contents entries).</summary>
    public bool RightTabLeader { get; init; }
    /// <summary>Outline level for headings (1..9), 0 = body text.</summary>
    public int OutlineLevel { get; init; }

    public bool Equals(ParaProps? o) => o is not null && Style == o.Style && Align == o.Align && IndentLeft == o.IndentLeft && IndentRight == o.IndentRight &&
        FirstLine == o.FirstLine && SpaceBefore == o.SpaceBefore && SpaceAfter == o.SpaceAfter && LineSpacing == o.LineSpacing && LineRule == o.LineRule &&
        List == o.List && ListLevel == o.ListLevel && ListId == o.ListId && KeepWithNext == o.KeepWithNext && PageBreakBefore == o.PageBreakBefore &&
        Shading?.ToArgb() == o.Shading?.ToArgb() && Border == o.Border && RightTabLeader == o.RightTabLeader && OutlineLevel == o.OutlineLevel;

    public override int GetHashCode() => HashCode.Combine(Style, Align, IndentLeft, FirstLine, List, ListLevel);
}

/// <summary>Inline object occupying one character (U+FFFC) in a run.</summary>
public abstract class InlineObject
{
    public abstract InlineObject Clone();
}

public sealed class ImageObject : InlineObject
{
    public byte[] Data { get; }
    public string Extension { get; }
    public float Width { get; set; }
    public float Height { get; set; }
    public float NaturalWidth { get; }
    public float NaturalHeight { get; }
    private Image? _image;

    public ImageObject(byte[] data, string extension, float width = 0, float height = 0)
    {
        Data = data;
        Extension = extension.TrimStart('.').ToLowerInvariant();
        try
        {
            var img = Image;
            NaturalWidth = img.Width * 72f / (img.HorizontalResolution > 1 ? img.HorizontalResolution : 96f);
            NaturalHeight = img.Height * 72f / (img.VerticalResolution > 1 ? img.VerticalResolution : 96f);
        }
        catch { NaturalWidth = 100; NaturalHeight = 100; }
        Width = width > 0 ? width : NaturalWidth;
        Height = height > 0 ? height : NaturalHeight;
    }

    /// <summary>Decoded image (shared by clones; never mutated).</summary>
    public Image Image => _image ??= Image.FromStream(new MemoryStream(Data));

    public override InlineObject Clone() => new ImageObject(Data, Extension, Width, Height) { _image = _image };

    public string ContentType => Extension switch { "jpg" or "jpeg" => "image/jpeg", "gif" => "image/gif", "bmp" => "image/bmp", _ => "image/png" };
}

public enum FieldKind { Page, NumPages, Date }

public sealed class FieldObject : InlineObject
{
    public FieldKind Kind { get; }
    public FieldObject(FieldKind kind) { Kind = kind; }
    public override InlineObject Clone() => new FieldObject(Kind);
    public string Instruction => Kind switch { FieldKind.Page => "PAGE", FieldKind.NumPages => "NUMPAGES", _ => "DATE" };
}

/// <summary>
/// A run of identically formatted text. Special characters: '\t' tab, '\v' line break, '\f' page break,
/// U+FFFC inline object (then the run has exactly one character and <see cref="Object"/> set).
/// </summary>
public sealed class Run
{
    public const char ObjectChar = '￼';
    public const char LineBreak = '\v';
    public const char PageBreak = '\f';

    public string Text { get; set; }
    public CharProps Props { get; set; }
    public InlineObject? Object { get; set; }

    public Run(string text, CharProps props, InlineObject? obj = null) { Text = text; Props = props; Object = obj; }

    public Run Clone() => new(Text, Props, Object?.Clone());
    public int Length => Text.Length;
}

public abstract class Block
{
    public abstract Block CloneBlock();
}

public sealed class Paragraph : Block
{
    public List<Run> Runs { get; set; } = new();
    public ParaProps Props { get; set; } = new();
    /// <summary>Formatting used for typing in an empty paragraph / the paragraph mark.</summary>
    public CharProps MarkProps { get; set; } = CharProps.Default;
    /// <summary>Containing table cell (null for body / header / footer paragraphs).</summary>
    public TableCell? Cell { get; internal set; }
    /// <summary>Incremented on every change (layout cache key).</summary>
    public int Version { get; private set; }

    public Paragraph() { }
    public Paragraph(string text, CharProps? props = null, ParaProps? pp = null)
    {
        props ??= CharProps.Default;
        MarkProps = props;
        if (pp != null) Props = pp;
        if (text.Length > 0) Runs.Add(new Run(text, props));
    }

    public void Touch() => Version++;

    public int Length => Runs.Sum(r => r.Text.Length);
    public string Text => string.Concat(Runs.Select(r => r.Text));

    public override Block CloneBlock() => Clone();

    public Paragraph Clone() => new()
    {
        Runs = Runs.Select(r => r.Clone()).ToList(), Props = Props, MarkProps = MarkProps
    };

    /// <summary>Character formatting at an offset (the char before it, like Word's typing behaviour).</summary>
    public CharProps PropsAt(int offset)
    {
        if (Runs.Count == 0) return MarkProps;
        int pos = 0;
        foreach (var r in Runs)
        {
            if (offset > pos && offset <= pos + r.Length) return r.Object != null ? r.Props with { } : r.Props;
            pos += r.Length;
        }
        return offset == 0 ? Runs[0].Props : Runs[^1].Props;
    }

    /// <summary>Returns (run index, offset in run) for a paragraph offset; splits nothing.</summary>
    public (int Run, int Off) Locate(int offset)
    {
        int pos = 0;
        for (int i = 0; i < Runs.Count; i++)
        {
            if (offset <= pos + Runs[i].Length) return (i, offset - pos);
            pos += Runs[i].Length;
        }
        return (Runs.Count, 0);
    }

    /// <summary>Splits runs so that a run boundary exists at <paramref name="offset"/>; returns index of the run starting there.</summary>
    public int SplitAt(int offset)
    {
        int pos = 0;
        for (int i = 0; i < Runs.Count; i++)
        {
            var r = Runs[i];
            if (offset == pos) return i;
            if (offset < pos + r.Length)
            {
                int k = offset - pos;
                var right = new Run(r.Text[k..], r.Props, null);
                r.Text = r.Text[..k];
                Runs.Insert(i + 1, right);
                return i + 1;
            }
            pos += r.Length;
        }
        return Runs.Count;
    }

    /// <summary>Merges adjacent runs with equal formatting and removes empty runs.</summary>
    public void Normalize()
    {
        for (int i = Runs.Count - 1; i >= 0; i--)
            if (Runs[i].Text.Length == 0) Runs.RemoveAt(i);
        for (int i = Runs.Count - 2; i >= 0; i--)
        {
            var a = Runs[i]; var b = Runs[i + 1];
            if (a.Object == null && b.Object == null && a.Props.Equals(b.Props))
            {
                a.Text += b.Text;
                Runs.RemoveAt(i + 1);
            }
        }
        Touch();
    }

    public bool IsHeading => Props.OutlineLevel is >= 1 and <= 9;
}

public sealed class TableCell
{
    public List<Paragraph> Paragraphs { get; set; } = new();
    public Color? Shading { get; set; }
    public int GridSpan { get; set; } = 1;
    public VMerge VMerge { get; set; }
    public CellVAlign VAlign { get; set; }
    public TableRow? Row { get; internal set; }

    public TableCell Clone()
    {
        var c = new TableCell { Shading = Shading, GridSpan = GridSpan, VMerge = VMerge, VAlign = VAlign };
        c.Paragraphs = Paragraphs.Select(p => p.Clone()).ToList();
        return c;
    }
}

public sealed class TableRow
{
    public List<TableCell> Cells { get; set; } = new();
    public float MinHeight { get; set; }
    public bool HeaderRow { get; set; }
    public Table? Table { get; internal set; }

    public TableRow Clone() => new() { Cells = Cells.Select(c => c.Clone()).ToList(), MinHeight = MinHeight, HeaderRow = HeaderRow };
}

public sealed class Table : Block
{
    public List<TableRow> Rows { get; set; } = new();
    /// <summary>Grid column widths in points.</summary>
    public List<float> ColumnWidths { get; set; } = new();
    public bool Borders { get; set; } = true;
    public Color BorderColor { get; set; } = Color.Black;
    public float BorderWidth { get; set; } = 0.5f;
    public ParaAlign Align { get; set; }

    public override Block CloneBlock() => new Table
    {
        Rows = Rows.Select(r => r.Clone()).ToList(), ColumnWidths = ColumnWidths.ToList(), Borders = Borders, BorderColor = BorderColor,
        BorderWidth = BorderWidth, Align = Align
    };

    public int ColumnCount => ColumnWidths.Count;

    public static Table Create(int rows, int cols, float totalWidth, CharProps cp)
    {
        var t = new Table();
        for (int c = 0; c < cols; c++) t.ColumnWidths.Add(totalWidth / cols);
        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                var cell = new TableCell();
                cell.Paragraphs.Add(new Paragraph("", cp, new ParaProps { SpaceAfter = 0, LineSpacing = 1f }));
                row.Cells.Add(cell);
            }
            t.Rows.Add(row);
        }
        return t;
    }

    /// <summary>Grid column index where each cell starts.</summary>
    public int GridColumnOf(TableRow row, TableCell cell)
    {
        int g = 0;
        foreach (var c in row.Cells) { if (c == cell) return g; g += Math.Max(1, c.GridSpan); }
        return g;
    }
}

public sealed class PageSetup
{
    public float Width { get; set; } = 595.3f;   // A4
    public float Height { get; set; } = 841.9f;
    public float MarginTop { get; set; } = 56.7f;  // 20 mm
    public float MarginBottom { get; set; } = 56.7f;
    public float MarginLeft { get; set; } = 85.05f; // 30 mm
    public float MarginRight { get; set; } = 42.5f; // 15 mm
    public float HeaderDistance { get; set; } = 35.4f;
    public float FooterDistance { get; set; } = 35.4f;
    public string PaperName { get; set; } = "A4";

    public bool Landscape => Width > Height;
    public float ContentWidth => Math.Max(36, Width - MarginLeft - MarginRight);
    public float ContentHeight => Math.Max(36, Height - MarginTop - MarginBottom);

    public PageSetup Clone() => (PageSetup)MemberwiseClone();

    public static readonly (string Name, float W, float H)[] Papers =
    [
        ("A4", 595.3f, 841.9f), ("A5", 419.5f, 595.3f), ("A3", 841.9f, 1190.6f), ("Letter", 612f, 792f), ("Legal", 612f, 1008f), ("B5", 498.9f, 708.7f)
    ];

    public static float Mm(float mm) => mm * 72f / 25.4f;
    public static float ToMm(float pt) => pt * 25.4f / 72f;
}

public static class StyleIds
{
    public const string Normal = "Normal", Heading1 = "Heading1", Heading2 = "Heading2", Heading3 = "Heading3", Title = "Title", Subtitle = "Subtitle",
        Quote = "Quote", NoSpacing = "NoSpacing", ListParagraph = "ListParagraph", Toc1 = "TOC1", Toc2 = "TOC2", Toc3 = "TOC3", TocHeading = "TOCHeading";
}

/// <summary>Named paragraph style: applying it resets the paragraph and character formatting to the style's values.</summary>
public sealed record StyleDef(string Id, string NameVi, string NameEn, ParaProps Para, CharProps Char)
{
    public string Name => MicroadsOffice.Core.Office.Loc.T(NameVi, NameEn);
}

public sealed class Document
{
    public List<Block> Blocks { get; set; } = new();
    public PageSetup Page { get; set; } = new();
    public List<Paragraph> Header { get; set; } = new();
    public List<Paragraph> Footer { get; set; } = new();
    public string? Watermark { get; set; }
    public Color? PageColor { get; set; }
    public string Title { get; set; } = "";
    public string Author { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Keywords { get; set; } = "";
    public DateTime Created { get; set; } = DateTime.Now;
    public Dictionary<string, StyleDef> Styles { get; set; } = DefaultStyles("Times New Roman", 13f);

    private List<Paragraph>? _flat;

    /// <summary>All editable body paragraphs in reading order (table cells row by row).</summary>
    public IReadOnlyList<Paragraph> Paragraphs => _flat ??= BuildFlat();

    /// <summary>Call after adding/removing blocks, rows, cells or paragraphs.</summary>
    public void StructureChanged() { _flat = null; _indexCache = null; }

    private Dictionary<Paragraph, int>? _indexCache;

    public int IndexOf(Paragraph p)
    {
        _indexCache ??= Paragraphs.Select((x, i) => (x, i)).ToDictionary(t => t.x, t => t.i, ReferenceEqualityComparer.Instance as IEqualityComparer<Paragraph>);
        return _indexCache.TryGetValue(p, out var i) ? i : -1;
    }

    private List<Paragraph> BuildFlat()
    {
        var list = new List<Paragraph>();
        foreach (var b in Blocks)
        {
            if (b is Paragraph p) { p.Cell = null; list.Add(p); }
            else if (b is Table t)
            {
                foreach (var row in t.Rows)
                {
                    row.Table = t;
                    foreach (var cell in row.Cells)
                    {
                        cell.Row = row;
                        if (cell.Paragraphs.Count == 0) cell.Paragraphs.Add(new Paragraph());
                        foreach (var cp in cell.Paragraphs) { cp.Cell = cell; list.Add(cp); }
                    }
                }
            }
        }
        if (list.Count == 0)
        {
            var p = new Paragraph("", Styles[StyleIds.Normal].Char, Styles[StyleIds.Normal].Para);
            Blocks.Add(p);
            list.Add(p);
        }
        // the document must end with a paragraph so the caret can be placed after a trailing table
        if (Blocks[^1] is Table)
        {
            var p = new Paragraph("", Styles[StyleIds.Normal].Char, Styles[StyleIds.Normal].Para);
            Blocks.Add(p);
            list.Add(p);
        }
        return list;
    }

    public Document Clone()
    {
        var d = new Document
        {
            Blocks = Blocks.Select(b => b.CloneBlock()).ToList(), Page = Page.Clone(), Header = Header.Select(p => p.Clone()).ToList(),
            Footer = Footer.Select(p => p.Clone()).ToList(), Watermark = Watermark, PageColor = PageColor, Title = Title, Author = Author,
            Subject = Subject, Keywords = Keywords, Created = Created, Styles = new Dictionary<string, StyleDef>(Styles)
        };
        return d;
    }

    /// <summary>Index of the top-level block that contains the paragraph.</summary>
    public int BlockIndexOf(Paragraph p)
    {
        if (p.Cell?.Row?.Table is Table t) return Blocks.IndexOf(t);
        return Blocks.IndexOf(p);
    }

    public string PlainText => string.Join("\n", Paragraphs.Select(p => p.Text.Replace(Run.ObjectChar.ToString(), "")));

    public static Dictionary<string, StyleDef> DefaultStyles(string font, float size)
    {
        var baseChar = new CharProps { FontName = font, Size = size };
        var basePara = new ParaProps { SpaceAfter = 6, LineSpacing = 1.15f };
        var heading = Color.FromArgb(31, 56, 100);
        StyleDef S(string id, string vi, string en, ParaProps p, CharProps c) => new(id, vi, en, p with { Style = id }, c);
        var list = new[]
        {
            S(StyleIds.Normal, "Bình thường", "Normal", basePara, baseChar),
            S(StyleIds.NoSpacing, "Không giãn cách", "No Spacing", basePara with { SpaceAfter = 0, LineSpacing = 1f }, baseChar),
            S(StyleIds.Heading1, "Tiêu đề 1", "Heading 1", basePara with { SpaceBefore = 12, SpaceAfter = 6, KeepWithNext = true, OutlineLevel = 1 }, baseChar with { Size = size + 3, Bold = true, Color = heading }),
            S(StyleIds.Heading2, "Tiêu đề 2", "Heading 2", basePara with { SpaceBefore = 10, SpaceAfter = 4, KeepWithNext = true, OutlineLevel = 2 }, baseChar with { Size = size + 1, Bold = true, Color = heading }),
            S(StyleIds.Heading3, "Tiêu đề 3", "Heading 3", basePara with { SpaceBefore = 8, SpaceAfter = 4, KeepWithNext = true, OutlineLevel = 3 }, baseChar with { Bold = true, Italic = true, Color = heading }),
            S(StyleIds.Title, "Tựa đề", "Title", basePara with { Align = ParaAlign.Center, SpaceAfter = 12, LineSpacing = 1f }, baseChar with { Size = size + 13, Bold = true, Color = Color.FromArgb(23, 54, 93) }),
            S(StyleIds.Subtitle, "Tựa đề phụ", "Subtitle", basePara with { Align = ParaAlign.Center, SpaceAfter = 12 }, baseChar with { Size = size + 2, Italic = true, Color = Color.FromArgb(89, 89, 89) }),
            S(StyleIds.Quote, "Trích dẫn", "Quote", basePara with { IndentLeft = 36, IndentRight = 36, SpaceBefore = 6, SpaceAfter = 6, Align = ParaAlign.Center }, baseChar with { Italic = true, Color = Color.FromArgb(64, 64, 64) }),
            S(StyleIds.ListParagraph, "Đoạn danh sách", "List Paragraph", basePara with { IndentLeft = 36 }, baseChar),
            S(StyleIds.TocHeading, "Tiêu đề mục lục", "TOC Heading", basePara with { Align = ParaAlign.Center, SpaceBefore = 6, SpaceAfter = 12 }, baseChar with { Size = size + 3, Bold = true }),
            S(StyleIds.Toc1, "Mục lục 1", "TOC 1", basePara with { SpaceAfter = 3, RightTabLeader = true }, baseChar with { Bold = true }),
            S(StyleIds.Toc2, "Mục lục 2", "TOC 2", basePara with { SpaceAfter = 3, IndentLeft = 14, RightTabLeader = true }, baseChar),
            S(StyleIds.Toc3, "Mục lục 3", "TOC 3", basePara with { SpaceAfter = 3, IndentLeft = 28, RightTabLeader = true }, baseChar with { Italic = true }),
        };
        return list.ToDictionary(s => s.Id);
    }

    /// <summary>Styles shown in the ribbon gallery (in order).</summary>
    public static readonly string[] GalleryStyles = [StyleIds.Normal, StyleIds.NoSpacing, StyleIds.Heading1, StyleIds.Heading2, StyleIds.Heading3, StyleIds.Title, StyleIds.Subtitle, StyleIds.Quote];

    public static Document CreateBlank(string font = "Times New Roman", float size = 13f)
    {
        var d = new Document { Styles = DefaultStyles(font, size) };
        var n = d.Styles[StyleIds.Normal];
        d.Blocks.Add(new Paragraph("", n.Char, n.Para));
        return d;
    }
}

/// <summary>Position in the document: paragraph index in <see cref="Document.Paragraphs"/> and character offset.</summary>
public readonly record struct DocPos(int Para, int Offset) : IComparable<DocPos>
{
    public int CompareTo(DocPos o) => Para != o.Para ? Para.CompareTo(o.Para) : Offset.CompareTo(o.Offset);
    public static bool operator <(DocPos a, DocPos b) => a.CompareTo(b) < 0;
    public static bool operator >(DocPos a, DocPos b) => a.CompareTo(b) > 0;
    public static bool operator <=(DocPos a, DocPos b) => a.CompareTo(b) <= 0;
    public static bool operator >=(DocPos a, DocPos b) => a.CompareTo(b) >= 0;
}
