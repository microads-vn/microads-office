using System.Text;
using System.Text.RegularExpressions;

namespace MicroadsOffice.AI;

/// <summary>Formatting of a piece of inline text.</summary>
public sealed record MdSpan(string Text, bool Bold = false, bool Italic = false, bool Underline = false, bool Code = false, bool Strike = false);

public enum MdAlign { Default, Left, Center, Right, Justify }

public abstract record MdBlock;
public sealed record MdHeading(int Level, IReadOnlyList<MdSpan> Spans, MdAlign Align = MdAlign.Default) : MdBlock;
public sealed record MdParagraph(IReadOnlyList<MdSpan> Spans, MdAlign Align = MdAlign.Default) : MdBlock;
public sealed record MdListItem(bool Ordered, int Level, IReadOnlyList<MdSpan> Spans) : MdBlock;
public sealed record MdTable(IReadOnlyList<IReadOnlyList<IReadOnlyList<MdSpan>>> Rows, bool HasHeader) : MdBlock;
public sealed record MdCode(string Language, string Text) : MdBlock;
public sealed record MdQuote(IReadOnlyList<MdSpan> Spans) : MdBlock;
public sealed record MdRule : MdBlock;
public sealed record MdPageBreak : MdBlock;

/// <summary>
/// Small Markdown reader for assistant output: headings, paragraphs, bullet/numbered lists (nested by indent), pipe tables,
/// fenced code, quotes, rules, **bold** / *italic* / __underline__ / ~~strike~~ / `code`, and ":::center|right|justify" blocks
/// for alignment (used for Vietnamese administrative headers), plus "\pagebreak".
/// </summary>
public static class Markdown
{
    static readonly Regex Ordered = new(@"^(\s*)(\d+)[.)]\s+(.*)$");
    static readonly Regex Bullet = new(@"^(\s*)[-*+•]\s+(.*)$");
    static readonly Regex Heading = new(@"^(#{1,6})\s+(.*?)\s*#*\s*$");
    static readonly Regex TableSep = new(@"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$");

    public static List<MdBlock> Parse(string md)
    {
        var lines = md.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var blocks = new List<MdBlock>();
        var para = new List<string>();
        var align = MdAlign.Default;

        void Flush()
        {
            if (para.Count == 0) return;
            // A hard line break ("two spaces" or a backslash at the end) is kept as \v (line break inside a paragraph).
            var sb = new StringBuilder();
            for (int i = 0; i < para.Count; i++)
            {
                var l = para[i];
                bool hard = l.EndsWith("  ") || l.EndsWith('\\');
                l = l.TrimEnd().TrimEnd('\\');
                sb.Append(l.Trim());
                if (i < para.Count - 1) sb.Append(hard || align != MdAlign.Default ? "\v" : " ");
            }
            blocks.Add(new MdParagraph(Inline(sb.ToString()), align));
            para.Clear();
        }

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var t = line.Trim();

            if (t.StartsWith("```") || t.StartsWith("~~~"))
            {
                Flush();
                var fence = t[..3]; var lang = t[3..].Trim();
                var code = new StringBuilder();
                for (i++; i < lines.Length && !lines[i].Trim().StartsWith(fence); i++) code.Append(lines[i]).Append('\n');
                blocks.Add(new MdCode(lang, code.ToString().TrimEnd('\n')));
                continue;
            }
            if (t.StartsWith(":::"))
            {
                Flush();
                var name = t[3..].Trim().ToLowerInvariant();
                align = name switch { "center" or "giữa" => MdAlign.Center, "right" or "phải" => MdAlign.Right, "justify" => MdAlign.Justify, "left" => MdAlign.Left, _ => MdAlign.Default };
                continue;
            }
            if (t.Length == 0) { Flush(); continue; }
            if (t is "\\pagebreak" or "<!-- pagebreak -->" or "\\newpage") { Flush(); blocks.Add(new MdPageBreak()); continue; }
            if (Regex.IsMatch(t, @"^(\*\s*){3,}$|^(-\s*){3,}$|^(_\s*){3,}$") && para.Count == 0) { blocks.Add(new MdRule()); continue; }

            var h = Heading.Match(t);
            if (h.Success) { Flush(); blocks.Add(new MdHeading(h.Groups[1].Length, Inline(h.Groups[2].Value), align)); continue; }

            if (t.StartsWith('|') && i + 1 < lines.Length && TableSep.IsMatch(lines[i + 1]))
            {
                Flush();
                var rows = new List<IReadOnlyList<IReadOnlyList<MdSpan>>> { Cells(t) };
                for (i += 2; i < lines.Length && lines[i].Trim().StartsWith('|'); i++) rows.Add(Cells(lines[i].Trim()));
                i--;
                blocks.Add(new MdTable(rows, true));
                continue;
            }

            var o = Ordered.Match(line);
            if (o.Success) { Flush(); blocks.Add(new MdListItem(true, Indent(o.Groups[1].Value), Inline(o.Groups[3].Value))); continue; }
            var b = Bullet.Match(line);
            if (b.Success && !(t.StartsWith("**") && !t.StartsWith("** ")) ) { Flush(); blocks.Add(new MdListItem(false, Indent(b.Groups[1].Value), Inline(b.Groups[2].Value))); continue; }

            if (t.StartsWith('>')) { Flush(); blocks.Add(new MdQuote(Inline(t.TrimStart('>').Trim()))); continue; }

            para.Add(line);
        }
        Flush();
        return blocks;
    }

    static int Indent(string ws) { int n = 0; foreach (var c in ws) n += c == '\t' ? 4 : 1; return Math.Min(8, n / 2); }

    static IReadOnlyList<IReadOnlyList<MdSpan>> Cells(string row)
    {
        var r = row.Trim();
        if (r.StartsWith('|')) r = r[1..];
        if (r.EndsWith('|') && !r.EndsWith("\\|")) r = r[..^1];
        var cells = new List<IReadOnlyList<MdSpan>>();
        var cur = new StringBuilder();
        for (int i = 0; i < r.Length; i++)
        {
            if (r[i] == '\\' && i + 1 < r.Length && r[i + 1] == '|') { cur.Append('|'); i++; continue; }
            if (r[i] == '|') { cells.Add(Inline(cur.ToString().Trim().Replace("<br>", "\v").Replace("<br/>", "\v"))); cur.Clear(); continue; }
            cur.Append(r[i]);
        }
        cells.Add(Inline(cur.ToString().Trim().Replace("<br>", "\v").Replace("<br/>", "\v")));
        return cells;
    }

    /// <summary>Parses inline emphasis. Unbalanced markers are kept as literal text.</summary>
    public static List<MdSpan> Inline(string s)
    {
        var spans = new List<MdSpan>();
        Walk(s, false, false, false, false, spans);
        // merge neighbours with identical formatting
        var merged = new List<MdSpan>();
        foreach (var sp in spans)
        {
            if (sp.Text.Length == 0) continue;
            if (merged.Count > 0 && merged[^1] with { Text = "" } == sp with { Text = "" }) merged[^1] = merged[^1] with { Text = merged[^1].Text + sp.Text };
            else merged.Add(sp);
        }
        return merged;
    }

    static void Walk(string s, bool bold, bool italic, bool under, bool strike, List<MdSpan> outp)
    {
        var text = new StringBuilder();
        void Emit() { if (text.Length > 0) { outp.Add(new MdSpan(Unescape(text.ToString()), bold, italic, under, false, strike)); text.Clear(); } }
        int i = 0;
        while (i < s.Length)
        {
            char c = s[i];
            if (c == '\\' && i + 1 < s.Length && "\\`*_~[]#|".Contains(s[i + 1])) { text.Append(s[i + 1]); i += 2; continue; }
            if (c == '`')
            {
                int end = s.IndexOf('`', i + 1);
                if (end > i) { Emit(); outp.Add(new MdSpan(s[(i + 1)..end], bold, italic, under, true, strike)); i = end + 1; continue; }
            }
            string? marker = null;
            if (s.AsSpan(i).StartsWith("**")) marker = "**";
            else if (s.AsSpan(i).StartsWith("__")) marker = "__";
            else if (s.AsSpan(i).StartsWith("~~")) marker = "~~";
            else if (c == '*' || (c == '_' && (i == 0 || !char.IsLetterOrDigit(s[i - 1])))) marker = c.ToString();
            if (marker != null)
            {
                int end = FindClose(s, i + marker.Length, marker);
                if (end > i + marker.Length)
                {
                    Emit();
                    var inner = s[(i + marker.Length)..end];
                    switch (marker)
                    {
                        case "**": Walk(inner, true, italic, under, strike, outp); break;
                        case "__": Walk(inner, bold, italic, true, strike, outp); break;
                        case "~~": Walk(inner, bold, italic, under, true, outp); break;
                        default: Walk(inner, bold, true, under, strike, outp); break;
                    }
                    i = end + marker.Length;
                    continue;
                }
            }
            // [text](url) → text (url kept after it in parentheses only when it differs)
            if (c == '[')
            {
                int close = s.IndexOf("](", i, StringComparison.Ordinal);
                int paren = close > 0 ? s.IndexOf(')', close) : -1;
                if (close > i && paren > close) { text.Append(s[(i + 1)..close]); var url = s[(close + 2)..paren]; if (url != s[(i + 1)..close]) text.Append(" (" + url + ")"); i = paren + 1; continue; }
            }
            text.Append(c); i++;
        }
        Emit();
    }

    static int FindClose(string s, int from, string marker)
    {
        for (int j = from; j <= s.Length - marker.Length; j++)
        {
            if (s[j] == '\\') { j++; continue; }
            if (string.CompareOrdinal(s, j, marker, 0, marker.Length) == 0)
            {
                if (marker.Length == 1 && j + 1 < s.Length && s[j + 1] == marker[0]) { j++; continue; } // skip "**" when looking for "*"
                if (j == from) return -1;
                return j;
            }
        }
        return -1;
    }

    static string Unescape(string s) => s.Replace("&nbsp;", " ").Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">");

    /// <summary>Plain text of spans.</summary>
    public static string Plain(IEnumerable<MdSpan> spans) => string.Concat(spans.Select(s => s.Text));

    /// <summary>Removes Markdown so a chat answer can be copied as plain text.</summary>
    public static string ToPlainText(string md)
    {
        var sb = new StringBuilder();
        foreach (var b in Parse(md))
        {
            switch (b)
            {
                case MdHeading h: sb.AppendLine(Plain(h.Spans)); break;
                case MdParagraph p: sb.AppendLine(Plain(p.Spans).Replace('\v', '\n')); break;
                case MdListItem l: sb.Append(new string(' ', l.Level * 2)).Append(l.Ordered ? "" : "• ").AppendLine(Plain(l.Spans)); break;
                case MdTable t: foreach (var r in t.Rows) sb.AppendLine(string.Join("\t", r.Select(Plain))); break;
                case MdCode c: sb.AppendLine(c.Text); break;
                case MdQuote q: sb.AppendLine(Plain(q.Spans)); break;
                case MdRule: sb.AppendLine("———"); break;
            }
        }
        return sb.ToString().TrimEnd();
    }
}
