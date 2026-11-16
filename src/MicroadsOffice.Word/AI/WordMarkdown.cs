using System.Drawing;
using System.Text;
using MicroadsOffice.AI;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.AI;

/// <summary>Converts between the Word model and the Markdown the assistant reads and writes.</summary>
public static class WordMarkdown
{
    // ------------------------------------------------------------------ document → Markdown (context for the model)

    public static string Export(Document d, int maxChars)
    {
        var sb = new StringBuilder();
        Dictionary<Paragraph, string> labels;
        try { labels = DocLayout.ComputeListLabels(d); } catch { labels = []; }
        foreach (var b in d.Blocks)
        {
            if (sb.Length > maxChars) { sb.AppendLine("\n[… " + MicroadsOffice.Core.Office.Loc.T("phần còn lại của văn bản bị lược bớt", "rest of the document truncated") + " …]"); break; }
            switch (b)
            {
                case Paragraph p: AppendParagraph(sb, p, labels, StyleChar(d, p)); break;
                case Table t:
                    sb.AppendLine();
                    for (int r = 0; r < t.Rows.Count; r++)
                    {
                        sb.Append('|');
                        bool header = r == 0;
                        foreach (var c in t.Rows[r].Cells)
                            sb.Append(' ').Append(string.Join("<br>", c.Paragraphs.Select(x => Inline(x, StyleChar(d, x), plainBold: header).Replace("|", "\\|")))).Append(" |");
                        sb.AppendLine();
                        if (r == 0) sb.AppendLine("|" + string.Concat(Enumerable.Repeat(" --- |", t.Rows[0].Cells.Count)));
                    }
                    sb.AppendLine();
                    break;
            }
        }
        var s = sb.ToString().Trim();
        return s.Length > maxChars + 200 ? s[..maxChars] : s;
    }

    /// <summary>Character formatting the paragraph's style already gives (bold of a heading is not "**bold**").</summary>
    static CharProps StyleChar(Document d, Paragraph p) => d.Styles.TryGetValue(p.Props.Style, out var s) ? s.Char : CharProps.Default;

    static void AppendParagraph(StringBuilder sb, Paragraph p, Dictionary<Paragraph, string> labels, CharProps styleChar)
    {
        var text = Inline(p, styleChar);
        if (p.Props.OutlineLevel is >= 1 and <= 6) { sb.Append(new string('#', p.Props.OutlineLevel)).Append(' ').AppendLine(text); return; }
        if (p.Props.Style == StyleIds.Title) { sb.Append("# ").AppendLine(text); return; }
        if (p.Props.List != ListType.None)
        {
            sb.Append(new string(' ', p.Props.ListLevel * 2));
            sb.Append(p.Props.List == ListType.Bullet ? "- " : (labels.TryGetValue(p, out var l) ? l.Trim() + " " : "1. "));
            sb.AppendLine(text);
            return;
        }
        if (text.Length == 0) { sb.AppendLine(); return; }
        if (p.Props.Align == ParaAlign.Center) { sb.AppendLine(":::center").AppendLine(text).AppendLine(":::"); return; }
        if (p.Props.Align == ParaAlign.Right) { sb.AppendLine(":::right").AppendLine(text).AppendLine(":::"); return; }
        sb.AppendLine(text);
        sb.AppendLine();
    }

    /// <summary>
    /// Paragraph text with **bold** / *italic* markers where the run differs from the paragraph style
    /// (<paramref name="plainBold"/>: bold is part of the look, e.g. a table header row).
    /// </summary>
    public static string Inline(Paragraph p, CharProps? styleChar = null, bool plainBold = false)
    {
        var baseChar = styleChar ?? CharProps.Default;
        var sb = new StringBuilder();
        foreach (var r in p.Runs)
        {
            if (r.Object != null) { sb.Append(r.Object is ImageObject ? "[image]" : ""); continue; }
            var t = r.Text.Replace('\v', ' ').Replace('\f', ' ').Replace("*", "\\*");
            if (t.Trim().Length == 0) { sb.Append(t); continue; }
            bool bold = r.Props.Bold && !baseChar.Bold && !plainBold;
            bool italic = r.Props.Italic && !baseChar.Italic;
            if (bold && italic) sb.Append("***").Append(t).Append("***");
            else if (bold) sb.Append("**").Append(t).Append("**");
            else if (italic) sb.Append('*').Append(t).Append('*');
            else sb.Append(t);
        }
        return sb.ToString().Replace("****", "");
    }

    // ------------------------------------------------------------------ Markdown → blocks

    public static List<Block> ToBlocks(Document d, string markdown) => ToBlocks(d, Markdown.Parse(markdown));

    public static List<Block> ToBlocks(Document d, List<MdBlock> md)
    {
        var normal = Style(d, StyleIds.Normal);
        var blocks = new List<Block>();
        int listId = 0; bool? listOrdered = null;
        foreach (var b in md)
        {
            if (b is not MdListItem) { listOrdered = null; }
            switch (b)
            {
                case MdHeading h:
                    {
                        var st = Style(d, h.Level <= 1 ? StyleIds.Heading1 : h.Level == 2 ? StyleIds.Heading2 : StyleIds.Heading3);
                        var p = Para(h.Spans, st.Char, st.Para with { Align = Align(h.Align, st.Para.Align) });
                        blocks.Add(p);
                        break;
                    }
                case MdParagraph p:
                    blocks.Add(Para(p.Spans, normal.Char, normal.Para with { Align = Align(p.Align, normal.Para.Align) }));
                    break;
                case MdListItem li:
                    {
                        if (listOrdered != li.Ordered || listId == 0) { listId = Editing.NewListId(d) + blocks.Count; listOrdered = li.Ordered; }
                        var pp = normal.Para with
                        {
                            List = li.Ordered ? ListType.Decimal : ListType.Bullet, ListId = listId, ListLevel = li.Level,
                            IndentLeft = 36 + 18 * li.Level, FirstLine = -18, SpaceAfter = 2,
                        };
                        blocks.Add(Para(li.Spans, normal.Char, pp));
                        break;
                    }
                case MdQuote q:
                    {
                        var st = d.Styles.ContainsKey(StyleIds.Quote) ? Style(d, StyleIds.Quote) : normal;
                        blocks.Add(Para(q.Spans, st.Char, st.Para));
                        break;
                    }
                case MdCode c:
                    foreach (var line in c.Text.Split('\n'))
                        blocks.Add(new Paragraph(line.Length == 0 ? "" : line, normal.Char with { FontName = "Consolas", Size = Math.Max(9, normal.Char.Size - 2) }, normal.Para with { SpaceAfter = 0 }));
                    break;
                case MdRule:
                    blocks.Add(new Paragraph("———", normal.Char, normal.Para with { Align = ParaAlign.Center }));
                    break;
                case MdPageBreak:
                    blocks.Add(new Paragraph("\f", normal.Char, normal.Para));
                    break;
                case MdTable t:
                    blocks.Add(BuildTable(d, t, normal));
                    break;
            }
        }
        return blocks;
    }

    static StyleDef Style(Document d, string id) =>
        d.Styles.TryGetValue(id, out var s) ? s : Document.CreateBlank().Styles[id];

    static ParaAlign Align(MdAlign a, ParaAlign dflt) => a switch
    {
        MdAlign.Center => ParaAlign.Center, MdAlign.Right => ParaAlign.Right, MdAlign.Justify => ParaAlign.Justify, MdAlign.Left => ParaAlign.Left, _ => dflt,
    };

    static Paragraph Para(IReadOnlyList<MdSpan> spans, CharProps baseChar, ParaProps pp)
    {
        var p = new Paragraph("", baseChar, pp) { MarkProps = baseChar };
        p.Runs.Clear();
        foreach (var s in spans)
        {
            var cp = baseChar with
            {
                Bold = baseChar.Bold || s.Bold, Italic = baseChar.Italic || s.Italic,
                Underline = s.Underline ? UnderlineStyle.Single : baseChar.Underline, Strike = s.Strike,
                FontName = s.Code ? "Consolas" : baseChar.FontName,
            };
            p.Runs.Add(new Run(s.Text, cp));
        }
        p.Normalize();
        return p;
    }

    static Table BuildTable(Document d, MdTable t, StyleDef normal)
    {
        int cols = Math.Max(1, t.Rows.Max(r => r.Count));
        var cp = normal.Char;
        var table = Table.Create(t.Rows.Count, cols, d.Page.ContentWidth, cp);
        // proportional widths from content length (bounded), so a narrow "STT" column does not take a quarter of the page
        var weights = new float[cols];
        for (int c = 0; c < cols; c++)
            weights[c] = Math.Clamp(t.Rows.Max(r => c < r.Count ? Markdown.Plain(r[c]).Length : 0), 3, 40);
        float sum = weights.Sum();
        for (int c = 0; c < cols; c++) table.ColumnWidths[c] = d.Page.ContentWidth * weights[c] / sum;
        for (int r = 0; r < t.Rows.Count; r++)
        {
            bool header = r == 0 && t.HasHeader;
            table.Rows[r].HeaderRow = header;
            for (int c = 0; c < cols; c++)
            {
                var cell = table.Rows[r].Cells[c];
                var spans = c < t.Rows[r].Count ? t.Rows[r][c] : [];
                var pp = normal.Para with { SpaceAfter = 0, SpaceBefore = 0, Align = header ? ParaAlign.Center : ParaAlign.Left };
                var para = Para(spans, header ? cp with { Bold = true } : cp, pp);
                cell.Paragraphs.Clear();
                cell.Paragraphs.Add(para);
                if (header) cell.Shading = Color.FromArgb(242, 242, 242);
            }
        }
        return table;
    }
}
