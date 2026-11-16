using System.Drawing;
using System.Globalization;
using System.Net;
using System.Text;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.IO;

public static class RtfWriter
{
    public static string Write(Document d)
    {
        var fonts = new List<string>();
        var colors = new List<Color> { };
        int Font(string f) { int i = fonts.IndexOf(f); if (i < 0) { fonts.Add(f); i = fonts.Count - 1; } return i; }
        int Col(Color c) { int i = colors.FindIndex(x => x.ToArgb() == c.ToArgb()); if (i < 0) { colors.Add(c); i = colors.Count - 1; } return i + 1; }
        Font(d.Styles[StyleIds.Normal].Char.FontName);
        var body = new StringBuilder();
        var p = d.Page;
        body.Append($"\\paperw{Tw(p.Width)}\\paperh{Tw(p.Height)}\\margl{Tw(p.MarginLeft)}\\margr{Tw(p.MarginRight)}\\margt{Tw(p.MarginTop)}\\margb{Tw(p.MarginBottom)}{(p.Landscape ? "\\landscape" : "")}\n");
        var labels = DocLayout.ComputeListLabels(d);
        foreach (var b in d.Blocks)
        {
            if (b is Paragraph para) Para(body, para, false, labels, Font, Col);
            else if (b is Table t)
            {
                foreach (var row in t.Rows)
                {
                    body.Append("\\trowd\\trgaph108");
                    float x = 0; int g = 0;
                    foreach (var c in row.Cells)
                    {
                        int span = Math.Max(1, c.GridSpan);
                        x += t.ColumnWidths.Skip(g).Take(span).Sum();
                        g += span;
                        if (t.Borders) body.Append("\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10");
                        if (c.Shading is Color sh) body.Append($"\\clcbpat{Col(sh)}");
                        body.Append($"\\cellx{Tw(x)}");
                    }
                    body.Append('\n');
                    foreach (var c in row.Cells)
                    {
                        for (int i = 0; i < c.Paragraphs.Count; i++)
                        {
                            Para(body, c.Paragraphs[i], true, labels, Font, Col, last: i == c.Paragraphs.Count - 1);
                        }
                        body.Append("\\cell\n");
                    }
                    body.Append("\\row\n");
                }
            }
        }
        var sb = new StringBuilder("{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1{\\fonttbl");
        for (int i = 0; i < fonts.Count; i++) sb.Append($"{{\\f{i}\\fnil\\fcharset0 {fonts[i]};}}");
        sb.Append("}{\\colortbl;");
        foreach (var c in colors) sb.Append($"\\red{c.R}\\green{c.G}\\blue{c.B};");
        sb.Append("}\n{\\*\\generator Microads Office Word;}\n").Append(body).Append('}');
        return sb.ToString();
    }

    private static string Tw(float pt) => ((int)Math.Round(pt * 20)).ToString(CultureInfo.InvariantCulture);

    private static void Para(StringBuilder sb, Paragraph p, bool inTable, Dictionary<Paragraph, string> labels, Func<string, int> font, Func<Color, int> col, bool last = false)
    {
        var pp = p.Props;
        sb.Append("\\pard\\plain");
        if (inTable) sb.Append("\\intbl");
        sb.Append(pp.Align switch { ParaAlign.Center => "\\qc", ParaAlign.Right => "\\qr", ParaAlign.Justify => "\\qj", _ => "\\ql" });
        sb.Append($"\\li{Tw(pp.IndentLeft)}\\ri{Tw(pp.IndentRight)}\\fi{Tw(pp.FirstLine)}\\sb{Tw(pp.SpaceBefore)}\\sa{Tw(pp.SpaceAfter)}");
        if (pp.LineRule == LineRule.Multiple) sb.Append($"\\sl{(int)(pp.LineSpacing * 240)}\\slmult1");
        else sb.Append($"\\sl{(pp.LineRule == LineRule.Exact ? "-" : "")}{Tw(pp.LineSpacing)}\\slmult0");
        if (pp.PageBreakBefore) sb.Append("\\pagebb");
        if (pp.Shading is Color sh) sb.Append($"\\cbpat{col(sh)}");
        sb.Append(' ');
        if (labels.TryGetValue(p, out var label))
        {
            var lp = p.Runs.FirstOrDefault()?.Props ?? p.MarkProps;
            sb.Append($"{{\\f{font(lp.FontName)}\\fs{(int)(lp.Size * 2)} ").Append(Escape(label)).Append("\\tab}");
        }
        foreach (var r in p.Runs)
        {
            var c = r.Props;
            sb.Append($"{{\\f{font(c.FontName)}\\fs{(int)Math.Round(c.Size * 2)}");
            if (c.Bold) sb.Append("\\b");
            if (c.Italic) sb.Append("\\i");
            if (c.Underline == UnderlineStyle.Double) sb.Append("\\uldb"); else if (c.Underline != UnderlineStyle.None) sb.Append("\\ul");
            if (c.Strike) sb.Append("\\strike");
            if (c.AllCaps) sb.Append("\\caps");
            if (c.VertAlign == VertAlign.Superscript) sb.Append("\\super"); else if (c.VertAlign == VertAlign.Subscript) sb.Append("\\sub");
            if (c.Color is Color cc) sb.Append($"\\cf{col(cc)}");
            if (c.Highlight is Color hc) sb.Append($"\\chcbpat{col(hc)}\\highlight{col(hc)}");
            sb.Append(' ');
            if (r.Object is ImageObject img)
            {
                bool jpg = img.Extension is "jpg" or "jpeg";
                byte[] data = img.Data;
                if (!jpg && img.Extension != "png")
                {
                    using var ms = new MemoryStream();
                    img.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    data = ms.ToArray();
                }
                sb.Append($"{{\\pict{(jpg ? "\\jpegblip" : "\\pngblip")}\\picw{img.Image.Width}\\pich{img.Image.Height}\\picwgoal{Tw(img.Width)}\\pichgoal{Tw(img.Height)} ");
                sb.Append(Convert.ToHexString(data)).Append('}');
            }
            else if (r.Object is FieldObject f) sb.Append($"{{\\field{{\\*\\fldinst {f.Instruction}}}{{\\fldrslt 1}}}}");
            else
            {
                if (c.Link != null) sb.Append($"{{\\field{{\\*\\fldinst HYPERLINK \"{Escape(c.Link)}\"}}{{\\fldrslt ");
                sb.Append(Escape(r.Text));
                if (c.Link != null) sb.Append("}}");
            }
            sb.Append('}');
        }
        if (!inTable || !last) sb.Append("\\par\n");
    }

    public static string Escape(string s)
    {
        var sb = new StringBuilder();
        foreach (var ch in s)
        {
            if (ch == '\\' || ch == '{' || ch == '}') sb.Append('\\').Append(ch);
            else if (ch == '\t') sb.Append("\\tab ");
            else if (ch == Run.LineBreak) sb.Append("\\line ");
            else if (ch == Run.PageBreak) sb.Append("\\page ");
            else if (ch == Run.ObjectChar) { }
            else if (ch > 127) sb.Append("\\u").Append(((short)ch).ToString(CultureInfo.InvariantCulture)).Append('?');
            else if (ch >= 32) sb.Append(ch);
        }
        return sb.ToString();
    }
}

/// <summary>RTF import through the Windows rich edit control (text, fonts, colours, alignment, indents, bullets).</summary>
public static class RtfReader
{
    public static Document Read(string rtf)
    {
        using var rtb = new RichTextBox();
        rtb.CreateControl();
        rtb.Rtf = rtf;
        return FromRichTextBox(rtb);
    }

    public static Document FromRichTextBox(RichTextBox rtb)
    {
        var d = new Document();
        var text = rtb.Text;
        int start = 0;
        var paras = new List<Paragraph>();
        while (start <= text.Length)
        {
            int nl = text.IndexOf('\n', start);
            int end = nl < 0 ? text.Length : nl;
            var p = new Paragraph();
            rtb.Select(start, 0);
            var align = rtb.SelectionAlignment;
            float px = 72f / 96f;
            p.Props = new ParaProps
            {
                Align = align == HorizontalAlignment.Center ? ParaAlign.Center : align == HorizontalAlignment.Right ? ParaAlign.Right : ParaAlign.Left,
                IndentLeft = (rtb.SelectionIndent + rtb.SelectionHangingIndent) * px,
                FirstLine = -rtb.SelectionHangingIndent * px,
                IndentRight = rtb.SelectionRightIndent * px,
                SpaceAfter = 0, LineSpacing = 1.15f,
                List = rtb.SelectionBullet ? ListType.Bullet : ListType.None,
                ListId = rtb.SelectionBullet ? 999 : 0
            };
            if (p.Props.List != ListType.None) p.Props = p.Props with { IndentLeft = Math.Max(36, p.Props.IndentLeft), FirstLine = -18 };
            CharProps? cur = null;
            var sb = new StringBuilder();
            for (int i = start; i < end; i++)
            {
                rtb.Select(i, 1);
                var f = rtb.SelectionFont;
                var cp = new CharProps
                {
                    FontName = f?.Name ?? "Times New Roman", Size = f?.SizeInPoints ?? 12, Bold = f?.Bold ?? false, Italic = f?.Italic ?? false,
                    Underline = f?.Underline == true ? UnderlineStyle.Single : UnderlineStyle.None, Strike = f?.Strikeout ?? false,
                    Color = rtb.SelectionColor.ToArgb() == Color.Black.ToArgb() || rtb.SelectionColor.IsEmpty ? null : rtb.SelectionColor,
                    VertAlign = rtb.SelectionCharOffset > 0 ? VertAlign.Superscript : rtb.SelectionCharOffset < 0 ? VertAlign.Subscript : VertAlign.Baseline,
                    Highlight = rtb.SelectionBackColor.ToArgb() is var bc && bc != Color.White.ToArgb() && bc != rtb.BackColor.ToArgb() && rtb.SelectionBackColor.A > 0 ? rtb.SelectionBackColor : null
                };
                if (cur != null && !cp.Equals(cur)) { p.Runs.Add(new Run(sb.ToString(), cur)); sb.Clear(); }
                cur = cp;
                char ch = text[i];
                sb.Append(ch == '\v' ? Run.LineBreak : ch);
            }
            if (cur != null && sb.Length > 0) p.Runs.Add(new Run(sb.ToString(), cur));
            if (p.Runs.Count == 0)
            {
                rtb.Select(start, 0);
                var f = rtb.SelectionFont;
                p.MarkProps = new CharProps { FontName = f?.Name ?? "Times New Roman", Size = f?.SizeInPoints ?? 12 };
            }
            else p.MarkProps = p.Runs[0].Props;
            p.Normalize();
            paras.Add(p);
            if (nl < 0) break;
            start = nl + 1;
        }
        // RichTextBox always ends with an implicit empty paragraph; drop it when the text ended with a newline
        if (paras.Count > 1 && paras[^1].Length == 0 && text.EndsWith('\n')) paras.RemoveAt(paras.Count - 1);
        d.Blocks.AddRange(paras);
        var normal = paras.FirstOrDefault(x => x.Runs.Count > 0)?.Runs[0].Props;
        if (normal != null) d.Styles = Document.DefaultStyles(normal.FontName, normal.Size);
        d.StructureChanged();
        return d;
    }
}

public static class HtmlWriter
{
    /// <summary>Escapes only markup characters so Vietnamese text stays readable in the UTF-8 file.</summary>
    private static string Enc(string? s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    public static string Write(Document d, bool fragment = false)
    {
        var sb = new StringBuilder();
        if (!fragment)
        {
            sb.Append("<!DOCTYPE html>\n<html><head><meta charset=\"utf-8\"><meta name=\"generator\" content=\"Microads Office Word (microads.vn)\">");
            sb.Append("<title>").Append(Enc(string.IsNullOrEmpty(d.Title) ? "Microads Office" : d.Title)).Append("</title>");
            sb.Append($"<style>body{{max-width:{(int)(d.Page.ContentWidth * 96 / 72)}px;margin:40px auto;font-family:'{d.Styles[StyleIds.Normal].Char.FontName}',serif;}} table{{border-collapse:collapse}} td{{vertical-align:top;padding:2px 7px}} p{{margin:0}}</style></head>");
            sb.Append("<body").Append(d.PageColor is Color pc ? $" style=\"background:#{Editing.Hex(pc)}\"" : "").Append(">\n");
        }
        var labels = DocLayout.ComputeListLabels(d);
        foreach (var b in d.Blocks)
        {
            if (b is Paragraph p) Para(sb, p, labels);
            else if (b is Table t)
            {
                sb.Append("<table").Append(t.Borders ? " border=\"1\" style=\"border-collapse:collapse\"" : "").Append(">\n");
                foreach (var row in t.Rows)
                {
                    sb.Append("<tr>");
                    int g = 0;
                    foreach (var c in row.Cells)
                    {
                        if (c.VMerge == VMerge.Continue) { g += Math.Max(1, c.GridSpan); continue; }
                        int span = Math.Max(1, c.GridSpan);
                        int rowspan = 1;
                        if (c.VMerge == VMerge.Restart)
                        {
                            int ri = t.Rows.IndexOf(row);
                            for (int k = ri + 1; k < t.Rows.Count; k++)
                            {
                                var nc = CellAtGrid(t, t.Rows[k], g);
                                if (nc?.VMerge == VMerge.Continue) rowspan++; else break;
                            }
                        }
                        float w = t.ColumnWidths.Skip(g).Take(span).Sum();
                        sb.Append($"<td style=\"width:{(int)(w * 96 / 72)}px{(c.Shading is Color sh ? ";background:#" + Editing.Hex(sh) : "")}{(t.Borders ? ";border:1px solid #" + Editing.Hex(t.BorderColor) : "")}\"");
                        if (span > 1) sb.Append($" colspan=\"{span}\"");
                        if (rowspan > 1) sb.Append($" rowspan=\"{rowspan}\"");
                        sb.Append('>');
                        foreach (var cp in c.Paragraphs) Para(sb, cp, labels);
                        sb.Append("</td>");
                        g += span;
                    }
                    sb.Append("</tr>\n");
                }
                sb.Append("</table>\n");
            }
        }
        if (!fragment) sb.Append("</body></html>\n");
        return sb.ToString();
    }

    private static TableCell? CellAtGrid(Table t, TableRow r, int g)
    {
        int col = 0;
        foreach (var c in r.Cells) { if (col == g) return c; col += Math.Max(1, c.GridSpan); }
        return null;
    }

    private static string F(float v) => v.ToString("0.##", CultureInfo.InvariantCulture);

    private static void Para(StringBuilder sb, Paragraph p, Dictionary<Paragraph, string> labels)
    {
        var pp = p.Props;
        string tag = pp.OutlineLevel is >= 1 and <= 6 ? "h" + pp.OutlineLevel : "p";
        var style = new StringBuilder();
        style.Append($"text-align:{pp.Align switch { ParaAlign.Center => "center", ParaAlign.Right => "right", ParaAlign.Justify => "justify", _ => "left" }};");
        style.Append($"margin:{F(pp.SpaceBefore)}pt {F(pp.IndentRight)}pt {F(pp.SpaceAfter)}pt {F(pp.IndentLeft)}pt;");
        if (pp.FirstLine != 0) style.Append($"text-indent:{F(pp.FirstLine)}pt;");
        if (pp.LineRule == LineRule.Multiple) style.Append($"line-height:{F(pp.LineSpacing * 1.15f)};"); else style.Append($"line-height:{F(pp.LineSpacing)}pt;");
        if (pp.Shading is Color sh) style.Append($"background:#{Editing.Hex(sh)};");
        if (pp.Border) style.Append("border:1px solid #000;padding:2px;");
        sb.Append($"<{tag} style=\"{style}\">");
        if (labels.TryGetValue(p, out var label)) sb.Append(Enc(label)).Append("&nbsp;&nbsp;");
        if (p.Runs.Count == 0) sb.Append("&nbsp;");
        foreach (var r in p.Runs)
        {
            var c = r.Props;
            if (r.Object is ImageObject img)
            {
                sb.Append($"<img src=\"data:{img.ContentType};base64,{Convert.ToBase64String(img.Data)}\" style=\"width:{F(img.Width)}pt;height:{F(img.Height)}pt\">");
                continue;
            }
            if (r.Object is FieldObject) { sb.Append('1'); continue; }
            var st = new StringBuilder($"font-family:'{c.FontName}';font-size:{F(c.Size)}pt;");
            if (c.Bold) st.Append("font-weight:bold;");
            if (c.Italic) st.Append("font-style:italic;");
            var deco = new List<string>();
            if (c.Underline != UnderlineStyle.None) deco.Add("underline");
            if (c.Strike) deco.Add("line-through");
            if (deco.Count > 0) st.Append($"text-decoration:{string.Join(" ", deco)};");
            if (c.Color is Color cc) st.Append($"color:#{Editing.Hex(cc)};");
            if (c.Highlight is Color hc) st.Append($"background:#{Editing.Hex(hc)};");
            if (c.AllCaps) st.Append("text-transform:uppercase;");
            var text = Enc(r.Text).Replace("\t", "&emsp;").Replace(Run.LineBreak.ToString(), "<br>").Replace(Run.PageBreak.ToString(), "<br style=\"page-break-before:always\">");
            text = text.Replace("  ", " &nbsp;");
            if (c.VertAlign == VertAlign.Superscript) text = "<sup>" + text + "</sup>";
            else if (c.VertAlign == VertAlign.Subscript) text = "<sub>" + text + "</sub>";
            if (c.Link != null) sb.Append($"<a href=\"{Enc(c.Link)}\" style=\"{st}\">{text}</a>");
            else sb.Append($"<span style=\"{st}\">{text}</span>");
        }
        sb.Append($"</{tag}>\n");
    }

    /// <summary>CF_HTML clipboard format (with the required header offsets).</summary>
    public static string ClipboardHtml(Document d)
    {
        var frag = Write(d, fragment: true);
        const string header = "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
        var pre = "<html><body><!--StartFragment-->";
        var post = "<!--EndFragment--></body></html>";
        int headerLen = string.Format(header, 0, 0, 0, 0).Length;
        int startHtml = headerLen;
        int startFrag = startHtml + Encoding.UTF8.GetByteCount(pre);
        int endFrag = startFrag + Encoding.UTF8.GetByteCount(frag);
        int endHtml = endFrag + Encoding.UTF8.GetByteCount(post);
        return string.Format(header, startHtml, endHtml, startFrag, endFrag) + pre + frag + post;
    }
}

public static class TextIO
{
    static TextIO() { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); }

    public static Document Read(string path)
    {
        var bytes = File.ReadAllBytes(path);
        string text = Decode(bytes);
        var d = Document.CreateBlank("Times New Roman", 13);
        d.Blocks.Clear();
        var n = d.Styles[StyleIds.Normal];
        foreach (var line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            d.Blocks.Add(new Paragraph(line.Replace("\f", ""), n.Char, n.Para));
        if (d.Blocks.Count > 1 && ((Paragraph)d.Blocks[^1]).Length == 0) d.Blocks.RemoveAt(d.Blocks.Count - 1);
        d.StructureChanged();
        return d;
    }

    public static string Decode(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        try { return new UTF8Encoding(false, true).GetString(bytes); }
        catch (DecoderFallbackException) { return Encoding.GetEncoding(1258).GetString(bytes); }
    }

    public static void Write(Document d, string path)
    {
        var sb = new StringBuilder();
        var labels = DocLayout.ComputeListLabels(d);
        foreach (var b in d.Blocks)
        {
            if (b is Paragraph p) sb.Append(labels.TryGetValue(p, out var l) ? l + " " : "").Append(Clean(p.Text)).Append("\r\n");
            else if (b is Table t)
                foreach (var row in t.Rows)
                    sb.Append(string.Join("\t", row.Cells.Select(c => string.Join(" ", c.Paragraphs.Select(x => Clean(x.Text)))))).Append("\r\n");
        }
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
    }

    private static string Clean(string s) => s.Replace(Run.ObjectChar.ToString(), "").Replace(Run.LineBreak, '\n').Replace(Run.PageBreak, '\f');
}
