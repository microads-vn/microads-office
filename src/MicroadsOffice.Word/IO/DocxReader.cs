using System.Drawing;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.IO;

/// <summary>Reads .docx files (Office Open XML). Styles are resolved into direct formatting; style ids are kept.</summary>
public sealed class DocxReader
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace DC = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    private readonly ZipArchive _zip;
    private readonly Dictionary<string, XElement> _styles = new();
    private string? _defaultParaStyle;
    private XElement? _defRPr, _defPPr;
    private string _minorFont = "Calibri", _majorFont = "Calibri Light";
    private readonly Dictionary<string, Dictionary<int, (ListType Type, float Left, float Hanging)>> _numbering = new();
    private readonly Document _doc = new();
    private Dictionary<string, (string Target, bool External)> _rels = new();
    private string _partDir = "word/";
    private string? _watermark;

    private DocxReader(ZipArchive zip) { _zip = zip; }

    public static Document Read(string path)
    {
        using var fs = File.OpenRead(path);
        return Read(fs);
    }

    public static Document Read(Stream s)
    {
        using var zip = new ZipArchive(s, ZipArchiveMode.Read, leaveOpen: true);
        return new DocxReader(zip).Parse();
    }

    private XDocument? Load(string name)
    {
        var e = _zip.GetEntry(name) ?? _zip.Entries.FirstOrDefault(x => string.Equals(x.FullName, name, StringComparison.OrdinalIgnoreCase));
        if (e == null) return null;
        using var st = e.Open();
        return XDocument.Load(st);
    }

    private byte[]? Bytes(string name)
    {
        var e = _zip.GetEntry(name) ?? _zip.Entries.FirstOrDefault(x => string.Equals(x.FullName, name, StringComparison.OrdinalIgnoreCase));
        if (e == null) return null;
        using var st = e.Open();
        using var ms = new MemoryStream();
        st.CopyTo(ms);
        return ms.ToArray();
    }

    private Dictionary<string, (string, bool)> LoadRels(string part)
    {
        var dir = Path.GetDirectoryName(part)?.Replace('\\', '/') ?? "";
        var relName = (dir.Length > 0 ? dir + "/" : "") + "_rels/" + Path.GetFileName(part) + ".rels";
        var x = Load(relName);
        var map = new Dictionary<string, (string, bool)>();
        if (x == null) return map;
        foreach (var r in x.Root!.Elements(PR + "Relationship"))
        {
            var id = (string?)r.Attribute("Id");
            var target = (string?)r.Attribute("Target") ?? "";
            bool ext = (string?)r.Attribute("TargetMode") == "External";
            if (id != null) map[id] = (target, ext);
        }
        return map;
    }

    private string MainPart()
    {
        var rels = Load("_rels/.rels");
        var main = rels?.Root?.Elements(PR + "Relationship").FirstOrDefault(r => ((string?)r.Attribute("Type") ?? "").EndsWith("/officeDocument"));
        var t = (string?)main?.Attribute("Target") ?? "word/document.xml";
        return t.TrimStart('/');
    }

    private Document Parse()
    {
        var main = MainPart();
        _partDir = (Path.GetDirectoryName(main)?.Replace('\\', '/') ?? "word") + "/";
        var docX = Load(main) ?? throw new InvalidDataException("word/document.xml not found");
        _rels = LoadRels(main);
        foreach (var (id, (target, ext)) in _rels.ToList()) { }
        LoadTheme();
        LoadStyles();
        LoadNumbering();
        BuildDocStyles();

        var body = docX.Root!.Element(W + "body");
        var bg = docX.Root.Element(W + "background");
        if (bg != null) _doc.PageColor = Editing.ParseColor((string?)bg.Attribute(W + "color"));
        if (body != null)
        {
            ReadBlocks(body, _doc.Blocks);
            var sect = body.Element(W + "sectPr") ?? body.Descendants(W + "sectPr").LastOrDefault();
            if (sect != null) ReadSection(sect);
        }
        if (_doc.Blocks.Count == 0) _doc.Blocks.Add(new Paragraph("", _doc.Styles[StyleIds.Normal].Char, _doc.Styles[StyleIds.Normal].Para));
        LoadCore();
        _doc.Watermark = _watermark;
        _doc.StructureChanged();
        return _doc;
    }

    private void LoadCore()
    {
        var c = Load("docProps/core.xml");
        if (c?.Root == null) return;
        _doc.Title = (string?)c.Root.Element(DC + "title") ?? "";
        _doc.Subject = (string?)c.Root.Element(DC + "subject") ?? "";
        _doc.Author = (string?)c.Root.Element(DC + "creator") ?? "";
        _doc.Keywords = (string?)c.Root.Element(CP + "keywords") ?? "";
    }

    private void LoadTheme()
    {
        var themeRel = _rels.FirstOrDefault(r => r.Value.Target.Contains("theme", StringComparison.OrdinalIgnoreCase));
        var name = themeRel.Value.Target != null ? Resolve(themeRel.Value.Target) : "word/theme/theme1.xml";
        var t = Load(name);
        if (t == null) return;
        var minor = t.Descendants(A + "minorFont").FirstOrDefault()?.Element(A + "latin")?.Attribute("typeface")?.Value;
        var major = t.Descendants(A + "majorFont").FirstOrDefault()?.Element(A + "latin")?.Attribute("typeface")?.Value;
        if (!string.IsNullOrEmpty(minor)) _minorFont = minor;
        if (!string.IsNullOrEmpty(major)) _majorFont = major;
    }

    private string Resolve(string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = (_partDir + target).Split('/').ToList();
        for (int i = 0; i < parts.Count; i++)
        {
            if (parts[i] == "..") { parts.RemoveAt(i); if (i > 0) parts.RemoveAt(i - 1); i -= 2; if (i < -1) i = -1; }
            else if (parts[i] == ".") { parts.RemoveAt(i); i--; }
        }
        return string.Join("/", parts);
    }

    private void LoadStyles()
    {
        var x = Load(_partDir + "styles.xml");
        if (x?.Root == null) return;
        var dd = x.Root.Element(W + "docDefaults");
        _defRPr = dd?.Element(W + "rPrDefault")?.Element(W + "rPr");
        _defPPr = dd?.Element(W + "pPrDefault")?.Element(W + "pPr");
        foreach (var s in x.Root.Elements(W + "style"))
        {
            var id = (string?)s.Attribute(W + "styleId");
            if (id == null) continue;
            _styles[id] = s;
            if ((string?)s.Attribute(W + "type") == "paragraph" && Bool(s.Attribute(W + "default"))) _defaultParaStyle = id;
        }
    }

    private void LoadNumbering()
    {
        var x = Load(_partDir + "numbering.xml");
        if (x?.Root == null) return;
        var abs = new Dictionary<string, Dictionary<int, (ListType, float, float)>>();
        foreach (var a in x.Root.Elements(W + "abstractNum"))
        {
            var id = (string?)a.Attribute(W + "abstractNumId");
            if (id == null) continue;
            var levels = new Dictionary<int, (ListType, float, float)>();
            foreach (var l in a.Elements(W + "lvl"))
            {
                int ilvl = (int?)l.Attribute(W + "ilvl") ?? 0;
                var fmt = (string?)l.Element(W + "numFmt")?.Attribute(W + "val") ?? "decimal";
                var type = fmt switch
                {
                    "bullet" => ListType.Bullet, "lowerLetter" => ListType.LowerAlpha, "upperLetter" => ListType.UpperAlpha,
                    "lowerRoman" => ListType.LowerRoman, "upperRoman" => ListType.UpperRoman, "none" => ListType.None, _ => ListType.Decimal
                };
                var ind = l.Element(W + "pPr")?.Element(W + "ind");
                float left = Twip(ind?.Attribute(W + "left") ?? ind?.Attribute(W + "start")) ?? 36 + 18 * ilvl;
                float hang = Twip(ind?.Attribute(W + "hanging")) ?? 18;
                levels[ilvl] = (type, left, hang);
            }
            abs[id] = levels;
        }
        foreach (var n in x.Root.Elements(W + "num"))
        {
            var id = (string?)n.Attribute(W + "numId");
            var aid = (string?)n.Element(W + "abstractNumId")?.Attribute(W + "val");
            if (id != null && aid != null && abs.TryGetValue(aid, out var lv)) _numbering[id] = lv;
        }
    }

    // ------------------------------------------------------------------ style resolution

    private IEnumerable<XElement> Chain(string? styleId)
    {
        var stack = new List<XElement>();
        var seen = new HashSet<string>();
        while (styleId != null && _styles.TryGetValue(styleId, out var s) && seen.Add(styleId))
        {
            stack.Add(s);
            styleId = (string?)s.Element(W + "basedOn")?.Attribute(W + "val");
        }
        stack.Reverse();
        return stack;
    }

    private CharProps BaseChar()
    {
        var c = new CharProps { FontName = "Times New Roman", Size = 10 };
        if (_defRPr != null) c = ApplyRPr(c, _defRPr);
        return c;
    }

    private ParaProps BaseParaProps()
    {
        var p = new ParaProps { SpaceAfter = 0, LineSpacing = 1f };
        if (_defPPr != null) p = ApplyPPr(p, _defPPr);
        return p;
    }

    private (ParaProps P, CharProps C) StyleProps(string? styleId)
    {
        var p = BaseParaProps();
        var c = BaseChar();
        foreach (var s in Chain(styleId))
        {
            var ppr = s.Element(W + "pPr");
            if (ppr != null) p = ApplyPPr(p, ppr);
            var rpr = s.Element(W + "rPr");
            if (rpr != null) c = ApplyRPr(c, rpr);
        }
        return (p, c);
    }

    private void BuildDocStyles()
    {
        var (np, nc) = StyleProps(_defaultParaStyle);
        var styles = Document.DefaultStyles(nc.FontName, nc.Size);
        styles[StyleIds.Normal] = styles[StyleIds.Normal] with { Para = np with { Style = StyleIds.Normal }, Char = nc };
        foreach (var id in styles.Keys.ToList())
        {
            if (id == StyleIds.Normal) continue;
            var sid = FindStyleId(id);
            if (sid == null) continue;
            var (p, c) = StyleProps(sid);
            styles[id] = styles[id] with { Para = p with { Style = id, OutlineLevel = styles[id].Para.OutlineLevel, RightTabLeader = styles[id].Para.RightTabLeader }, Char = c };
        }
        _doc.Styles = styles;
    }

    /// <summary>Maps our style ids to the document's style ids (by id or by name like "heading 1").</summary>
    private string? FindStyleId(string ourId)
    {
        if (_styles.ContainsKey(ourId)) return ourId;
        string name = ourId switch
        {
            StyleIds.Heading1 => "heading 1", StyleIds.Heading2 => "heading 2", StyleIds.Heading3 => "heading 3", StyleIds.Toc1 => "toc 1", StyleIds.Toc2 => "toc 2",
            StyleIds.Toc3 => "toc 3", StyleIds.NoSpacing => "No Spacing", StyleIds.ListParagraph => "List Paragraph", StyleIds.TocHeading => "TOC Heading", _ => ourId
        };
        return _styles.FirstOrDefault(kv => string.Equals((string?)kv.Value.Element(W + "name")?.Attribute(W + "val"), name, StringComparison.OrdinalIgnoreCase)).Key;
    }

    /// <summary>Our style id for a document style id (heading detection by name/outline level).</summary>
    private (string Id, int Outline) OurStyle(string? docStyle, ParaProps resolved)
    {
        if (docStyle == null) return (StyleIds.Normal, 0);
        var name = ((string?)(_styles.TryGetValue(docStyle, out var s) ? s.Element(W + "name")?.Attribute(W + "val") : null) ?? docStyle).ToLowerInvariant();
        if (name.StartsWith("heading ") && int.TryParse(name[8..], out var lvl)) return (lvl switch { 1 => StyleIds.Heading1, 2 => StyleIds.Heading2, 3 => StyleIds.Heading3, _ => docStyle }, lvl);
        if (name == "title") return (StyleIds.Title, 0);
        if (name == "subtitle") return (StyleIds.Subtitle, 0);
        if (name == "quote") return (StyleIds.Quote, 0);
        if (name == "no spacing") return (StyleIds.NoSpacing, 0);
        if (name == "list paragraph") return (StyleIds.ListParagraph, 0);
        if (name.StartsWith("toc ") && int.TryParse(name[4..], out var t)) return (t switch { 1 => StyleIds.Toc1, 2 => StyleIds.Toc2, _ => StyleIds.Toc3 }, 0);
        if (name == "toc heading") return (StyleIds.TocHeading, 0);
        return (docStyle, resolved.OutlineLevel);
    }

    private static bool Bool(XAttribute? a) => a != null && a.Value is "1" or "true" or "on";

    private static bool OnOff(XElement? e)
    {
        if (e == null) return false;
        var v = (string?)e.Attribute(W + "val");
        return v == null || v is "1" or "true" or "on";
    }

    private static float? Twip(XAttribute? a)
    {
        if (a == null) return null;
        if (float.TryParse(a.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) return v / 20f;
        return null;
    }

    private CharProps ApplyRPr(CharProps c, XElement rpr)
    {
        var fonts = rpr.Element(W + "rFonts");
        if (fonts != null)
        {
            var f = (string?)fonts.Attribute(W + "ascii") ?? (string?)fonts.Attribute(W + "hAnsi");
            var theme = (string?)fonts.Attribute(W + "asciiTheme") ?? (string?)fonts.Attribute(W + "hAnsiTheme");
            if (f != null) c = c with { FontName = f };
            else if (theme != null) c = c with { FontName = theme.StartsWith("major") ? _majorFont : _minorFont };
        }
        if (rpr.Element(W + "b") is { } b) c = c with { Bold = OnOff(b) };
        if (rpr.Element(W + "i") is { } i) c = c with { Italic = OnOff(i) };
        if (rpr.Element(W + "caps") is { } caps) c = c with { AllCaps = OnOff(caps) };
        if (rpr.Element(W + "strike") is { } st) c = c with { Strike = OnOff(st) };
        if (rpr.Element(W + "dstrike") is { } ds && OnOff(ds)) c = c with { Strike = true };
        if (rpr.Element(W + "u") is { } u)
        {
            var v = (string?)u.Attribute(W + "val") ?? "single";
            c = c with { Underline = v switch { "none" => UnderlineStyle.None, "double" => UnderlineStyle.Double, "dotted" or "dottedHeavy" or "dash" => UnderlineStyle.Dotted, "thick" => UnderlineStyle.Thick, "wave" or "wavyHeavy" or "wavyDouble" => UnderlineStyle.Wavy, _ => UnderlineStyle.Single } };
        }
        if (rpr.Element(W + "vertAlign") is { } va)
        {
            var v = (string?)va.Attribute(W + "val");
            c = c with { VertAlign = v == "superscript" ? VertAlign.Superscript : v == "subscript" ? VertAlign.Subscript : VertAlign.Baseline };
        }
        if (rpr.Element(W + "color") is { } col) c = c with { Color = Editing.ParseColor((string?)col.Attribute(W + "val")) };
        if (rpr.Element(W + "sz") is { } sz && float.TryParse((string?)sz.Attribute(W + "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out var hs)) c = c with { Size = hs / 2f };
        if (rpr.Element(W + "highlight") is { } hl) c = c with { Highlight = HighlightColor((string?)hl.Attribute(W + "val")) };
        if (rpr.Element(W + "shd") is { } shd && Editing.ParseColor((string?)shd.Attribute(W + "fill")) is Color fill) c = c with { Highlight = fill };
        return c;
    }

    private static Color? HighlightColor(string? v) => v switch
    {
        "yellow" => Color.Yellow, "green" => Color.Lime, "cyan" => Color.Cyan, "magenta" => Color.Magenta, "blue" => Color.Blue, "red" => Color.Red,
        "darkBlue" => Color.DarkBlue, "darkCyan" => Color.DarkCyan, "darkGreen" => Color.DarkGreen, "darkMagenta" => Color.DarkMagenta, "darkRed" => Color.DarkRed,
        "darkYellow" => Color.Olive, "darkGray" => Color.DarkGray, "lightGray" => Color.LightGray, "black" => Color.Black, "white" => Color.White, _ => null
    };

    private ParaProps ApplyPPr(ParaProps p, XElement ppr)
    {
        if (ppr.Element(W + "jc") is { } jc)
        {
            var v = (string?)jc.Attribute(W + "val");
            p = p with { Align = v switch { "center" => ParaAlign.Center, "right" or "end" => ParaAlign.Right, "both" or "distribute" => ParaAlign.Justify, _ => ParaAlign.Left } };
        }
        if (ppr.Element(W + "ind") is { } ind)
        {
            if (Twip(ind.Attribute(W + "left") ?? ind.Attribute(W + "start")) is float l) p = p with { IndentLeft = l };
            if (Twip(ind.Attribute(W + "right") ?? ind.Attribute(W + "end")) is float r) p = p with { IndentRight = r };
            if (Twip(ind.Attribute(W + "firstLine")) is float f) p = p with { FirstLine = f };
            if (Twip(ind.Attribute(W + "hanging")) is float h) p = p with { FirstLine = -h };
        }
        if (ppr.Element(W + "spacing") is { } sp)
        {
            if (Twip(sp.Attribute(W + "before")) is float b) p = p with { SpaceBefore = b };
            if (Twip(sp.Attribute(W + "after")) is float a) p = p with { SpaceAfter = a };
            if (sp.Attribute(W + "line") is { } line && float.TryParse(line.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var lv))
            {
                var rule = (string?)sp.Attribute(W + "lineRule") ?? "auto";
                p = rule switch
                {
                    "exact" => p with { LineRule = LineRule.Exact, LineSpacing = lv / 20f },
                    "atLeast" => p with { LineRule = LineRule.AtLeast, LineSpacing = lv / 20f },
                    _ => p with { LineRule = LineRule.Multiple, LineSpacing = lv / 240f }
                };
            }
        }
        if (ppr.Element(W + "keepNext") is { } kn) p = p with { KeepWithNext = OnOff(kn) };
        if (ppr.Element(W + "pageBreakBefore") is { } pb) p = p with { PageBreakBefore = OnOff(pb) };
        if (ppr.Element(W + "shd") is { } shd) p = p with { Shading = Editing.ParseColor((string?)shd.Attribute(W + "fill")) };
        if (ppr.Element(W + "pBdr") is { } bdr) p = p with { Border = bdr.Elements().Any(e => (string?)e.Attribute(W + "val") is not ("none" or "nil")) };
        if (ppr.Element(W + "outlineLvl") is { } ol && int.TryParse((string?)ol.Attribute(W + "val"), out var olv) && olv < 9) p = p with { OutlineLevel = olv + 1 };
        if (ppr.Element(W + "tabs") is { } tabs && tabs.Elements(W + "tab").Any(t => (string?)t.Attribute(W + "val") == "right" && (string?)t.Attribute(W + "leader") is "dot" or "hyphen" or "underscore"))
            p = p with { RightTabLeader = true };
        if (ppr.Element(W + "numPr") is { } np)
        {
            var numId = (string?)np.Element(W + "numId")?.Attribute(W + "val");
            int ilvl = (int?)np.Element(W + "ilvl")?.Attribute(W + "val") ?? 0;
            if (numId != null && numId != "0" && _numbering.TryGetValue(numId, out var lv) && lv.TryGetValue(ilvl, out var info) && info.Type != ListType.None)
                p = p with { List = info.Type, ListLevel = ilvl, ListId = int.TryParse(numId, out var nid) ? nid : 1, IndentLeft = info.Left, FirstLine = -info.Hanging };
            else if (numId == "0") p = p with { List = ListType.None };
        }
        return p;
    }

    // ------------------------------------------------------------------ body

    private void ReadBlocks(XElement container, List<Block> blocks)
    {
        foreach (var e in container.Elements())
        {
            if (e.Name == W + "p") blocks.Add(ReadParagraph(e));
            else if (e.Name == W + "tbl") blocks.Add(ReadTable(e));
            else if (e.Name == W + "sdt") { var c = e.Element(W + "sdtContent"); if (c != null) ReadBlocks(c, blocks); }
            else if (e.Name == W + "customXml" || e.Name == W + "ins") ReadBlocks(e, blocks);
        }
    }

    private void ReadParagraphsInto(XElement container, List<Paragraph> list)
    {
        var blocks = new List<Block>();
        ReadBlocks(container, blocks);
        foreach (var b in blocks)
        {
            if (b is Paragraph p) list.Add(p);
            else if (b is Table t)
                foreach (var row in t.Rows)
                    list.Add(new Paragraph(string.Join("\t", row.Cells.Select(c => string.Join(" ", c.Paragraphs.Select(x => x.Text)))), row.Cells.FirstOrDefault()?.Paragraphs.FirstOrDefault()?.MarkProps));
        }
    }

    private Paragraph ReadParagraph(XElement pe)
    {
        var ppr = pe.Element(W + "pPr");
        var styleId = (string?)ppr?.Element(W + "pStyle")?.Attribute(W + "val") ?? _defaultParaStyle;
        var (sp, sc) = StyleProps(styleId);
        var pp = ppr != null ? ApplyPPr(sp, ppr) : sp;
        var (ours, outline) = OurStyle(styleId, pp);
        pp = pp with { Style = ours, OutlineLevel = outline > 0 ? outline : pp.OutlineLevel };
        var para = new Paragraph { Props = pp };
        var markRpr = ppr?.Element(W + "rPr");
        para.MarkProps = markRpr != null ? ApplyRPr(sc, markRpr) with { Link = null } : sc;
        var state = new FieldState();
        ReadInline(pe, para, sc, null, state);
        para.Normalize();
        return para;
    }

    private sealed class FieldState
    {
        public int Depth;
        public string Instr = "";
        public bool InResult;
        public bool Replaced;
    }

    private void ReadInline(XElement parent, Paragraph para, CharProps styleChar, string? link, FieldState fs)
    {
        foreach (var e in parent.Elements())
        {
            var n = e.Name;
            if (n == W + "r") ReadRun(e, para, styleChar, link, fs);
            else if (n == W + "hyperlink")
            {
                string? url = null;
                var rid = (string?)e.Attribute(R + "id");
                if (rid != null && _rels.TryGetValue(rid, out var t)) url = t.Target;
                else if ((string?)e.Attribute(W + "anchor") is string anchor) url = null;
                ReadInline(e, para, styleChar, url, fs);
            }
            else if (n == W + "fldSimple")
            {
                var instr = ((string?)e.Attribute(W + "instr") ?? "").Trim().Split(' ')[0].ToUpperInvariant();
                var kind = instr switch { "PAGE" => FieldKind.Page, "NUMPAGES" or "SECTIONPAGES" => FieldKind.NumPages, _ => (FieldKind?)null };
                if (kind != null)
                {
                    var rpr = e.Descendants(W + "rPr").FirstOrDefault();
                    para.Runs.Add(new Run(Run.ObjectChar.ToString(), rpr != null ? ApplyRPr(styleChar, rpr) : styleChar, new FieldObject(kind.Value)));
                }
                else ReadInline(e, para, styleChar, link, fs);
            }
            else if (n == W + "ins" || n == W + "smartTag" || n == W + "customXml" || n == W + "moveTo" || n == W + "dir" || n == W + "bdo") ReadInline(e, para, styleChar, link, fs);
            else if (n == W + "sdt") { var c = e.Element(W + "sdtContent"); if (c != null) ReadInline(c, para, styleChar, link, fs); }
        }
    }

    private void ReadRun(XElement re, Paragraph para, CharProps styleChar, string? link, FieldState fs)
    {
        var rpr = re.Element(W + "rPr");
        var cp = styleChar;
        var rStyle = (string?)rpr?.Element(W + "rStyle")?.Attribute(W + "val");
        if (rStyle != null) foreach (var s in Chain(rStyle)) if (s.Element(W + "rPr") is { } srpr) cp = ApplyRPr(cp, srpr);
        if (rpr != null) cp = ApplyRPr(cp, rpr);
        if (link != null) cp = cp with { Link = link, Color = cp.Color ?? Color.FromArgb(5, 99, 193), Underline = cp.Underline == UnderlineStyle.None ? UnderlineStyle.Single : cp.Underline };
        var text = new System.Text.StringBuilder();
        void Flush()
        {
            if (text.Length == 0) return;
            if (!(fs.InResult && fs.Replaced)) para.Runs.Add(new Run(text.ToString(), cp));
            text.Clear();
        }
        foreach (var c in re.Elements())
        {
            var n = c.Name;
            if (n == W + "fldChar")
            {
                Flush();
                var t = (string?)c.Attribute(W + "fldCharType");
                if (t == "begin") { fs.Depth++; fs.Instr = ""; fs.InResult = false; fs.Replaced = false; }
                else if (t == "separate")
                {
                    fs.InResult = true;
                    var word = fs.Instr.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant();
                    if (word is "PAGE" or "NUMPAGES" or "SECTIONPAGES")
                    {
                        para.Runs.Add(new Run(Run.ObjectChar.ToString(), cp, new FieldObject(word == "PAGE" ? FieldKind.Page : FieldKind.NumPages)));
                        fs.Replaced = true;
                    }
                }
                else if (t == "end")
                {
                    if (!fs.InResult)
                    {
                        // field without result (e.g. PAGE written without separate)
                        var word = fs.Instr.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant();
                        if (word is "PAGE" or "NUMPAGES") para.Runs.Add(new Run(Run.ObjectChar.ToString(), cp, new FieldObject(word == "PAGE" ? FieldKind.Page : FieldKind.NumPages)));
                    }
                    fs.Depth = Math.Max(0, fs.Depth - 1); fs.InResult = false; fs.Replaced = false; fs.Instr = "";
                }
                continue;
            }
            if (n == W + "instrText") { fs.Instr += c.Value; continue; }
            if (fs.Depth > 0 && !fs.InResult) continue; // instruction part
            if (n == W + "t") text.Append(c.Value);
            else if (n == W + "tab") text.Append('\t');
            else if (n == W + "br")
            {
                var type = (string?)c.Attribute(W + "type");
                text.Append(type == "page" ? Run.PageBreak : Run.LineBreak);
            }
            else if (n == W + "cr") text.Append(Run.LineBreak);
            else if (n == W + "noBreakHyphen") text.Append('-');
            else if (n == W + "sym")
            {
                var ch = (string?)c.Attribute(W + "char");
                if (ch != null && int.TryParse(ch, NumberStyles.HexNumber, null, out var code))
                {
                    if (code >= 0xF000) code -= 0xF000;
                    var font = (string?)c.Attribute(W + "font") ?? "";
                    text.Append(font.Contains("Symbol") || font.Contains("Wingdings") ? SymbolMap(code) : (char)code);
                }
            }
            else if (n == W + "drawing")
            {
                Flush();
                var img = ReadDrawing(c);
                if (img != null && !(fs.InResult && fs.Replaced)) para.Runs.Add(new Run(Run.ObjectChar.ToString(), cp with { Link = null }, img));
            }
            else if (n == W + "pict" || n == W + "object")
            {
                Flush();
                var tp = c.Descendants(V + "textpath").FirstOrDefault(x => x.Attribute("string") != null);
                if (tp != null) { _watermark ??= (string?)tp.Attribute("string"); continue; }
                var id = c.Descendants(V + "imagedata").FirstOrDefault()?.Attribute(R + "id")?.Value;
                if (id != null && LoadImage(id) is { } im)
                {
                    var shape = c.Descendants(V + "shape").FirstOrDefault();
                    var (w, h) = VmlSize((string?)shape?.Attribute("style"));
                    if (w > 0) { im.Width = w; im.Height = h; }
                    para.Runs.Add(new Run(Run.ObjectChar.ToString(), cp, im));
                }
            }
        }
        Flush();
    }

    private static char SymbolMap(int code) => code switch
    {
        0xB7 => '•', 0xA7 => '▪', 0xFC => '✓', 0xD8 => '➢', 0x76 => '❖', 0x6F => '□', 0x71 => '◆', 0xE0 => '→', _ => (char)code
    };

    private static (float W, float H) VmlSize(string? style)
    {
        if (style == null) return (0, 0);
        float w = 0, h = 0;
        foreach (var part in style.Split(';'))
        {
            var kv = part.Split(':');
            if (kv.Length != 2) continue;
            float Parse(string v) => v.EndsWith("pt") && float.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var f) ? f
                : v.EndsWith("in") && float.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var g) ? g * 72 : 0;
            if (kv[0].Trim() == "width") w = Parse(kv[1].Trim());
            if (kv[0].Trim() == "height") h = Parse(kv[1].Trim());
        }
        return (w, h);
    }

    private ImageObject? ReadDrawing(XElement drawing)
    {
        var blip = drawing.Descendants(A + "blip").FirstOrDefault();
        var id = (string?)blip?.Attribute(R + "embed");
        if (id == null) return null;
        var img = LoadImage(id);
        if (img == null) return null;
        var ext = drawing.Descendants(WP + "extent").FirstOrDefault();
        if (ext != null && long.TryParse((string?)ext.Attribute("cx"), out var cx) && long.TryParse((string?)ext.Attribute("cy"), out var cy) && cx > 0 && cy > 0)
        {
            img.Width = cx / 12700f;
            img.Height = cy / 12700f;
        }
        return img;
    }

    private ImageObject? LoadImage(string relId)
    {
        if (!_rels.TryGetValue(relId, out var t) || t.External) return null;
        var data = Bytes(Resolve(t.Target));
        if (data == null) return null;
        try { return new ImageObject(data, Path.GetExtension(t.Target)); } catch { return null; }
    }

    private Table ReadTable(XElement te)
    {
        var t = new Table { Borders = false };
        var tblPr = te.Element(W + "tblPr");
        var style = (string?)tblPr?.Element(W + "tblStyle")?.Attribute(W + "val");
        XElement? borders = null;
        foreach (var s in Chain(style)) borders = s.Element(W + "tblPr")?.Element(W + "tblBorders") ?? borders;
        borders = tblPr?.Element(W + "tblBorders") ?? borders;
        if (borders != null)
        {
            var any = borders.Elements().Where(e => (string?)e.Attribute(W + "val") is not (null or "none" or "nil")).ToList();
            t.Borders = any.Count > 0;
            var first = any.FirstOrDefault();
            if (first != null)
            {
                t.BorderColor = Editing.ParseColor((string?)first.Attribute(W + "color")) ?? Color.Black;
                if (int.TryParse((string?)first.Attribute(W + "sz"), out var sz)) t.BorderWidth = Math.Max(0.25f, sz / 8f);
            }
        }
        var jc = (string?)tblPr?.Element(W + "jc")?.Attribute(W + "val");
        t.Align = jc == "center" ? ParaAlign.Center : jc is "right" or "end" ? ParaAlign.Right : ParaAlign.Left;
        foreach (var gc in te.Element(W + "tblGrid")?.Elements(W + "gridCol") ?? [])
            t.ColumnWidths.Add(Twip(gc.Attribute(W + "w")) ?? 72);
        bool cellBorders = false;
        foreach (var tr in te.Elements(W + "tr"))
        {
            var row = new TableRow();
            var trPr = tr.Element(W + "trPr");
            if (Twip(trPr?.Element(W + "trHeight")?.Attribute(W + "val")) is float rh) row.MinHeight = rh;
            row.HeaderRow = trPr?.Element(W + "tblHeader") != null;
            foreach (var tc in tr.Elements().SelectMany(e => e.Name == W + "sdt" ? e.Element(W + "sdtContent")?.Elements(W + "tc") ?? [] : e.Name == W + "tc" ? [e] : Array.Empty<XElement>()))
            {
                var cell = new TableCell();
                var tcPr = tc.Element(W + "tcPr");
                if ((int?)tcPr?.Element(W + "gridSpan")?.Attribute(W + "val") is int span) cell.GridSpan = Math.Max(1, span);
                var vm = tcPr?.Element(W + "vMerge");
                if (vm != null) cell.VMerge = (string?)vm.Attribute(W + "val") == "restart" ? VMerge.Restart : VMerge.Continue;
                cell.Shading = Editing.ParseColor((string?)tcPr?.Element(W + "shd")?.Attribute(W + "fill"));
                var va = (string?)tcPr?.Element(W + "vAlign")?.Attribute(W + "val");
                cell.VAlign = va == "center" ? CellVAlign.Center : va == "bottom" ? CellVAlign.Bottom : CellVAlign.Top;
                if (tcPr?.Element(W + "tcBorders")?.Elements().Any(e => (string?)e.Attribute(W + "val") is not (null or "none" or "nil")) == true) cellBorders = true;
                ReadParagraphsInto(tc, cell.Paragraphs);
                if (cell.Paragraphs.Count == 0) cell.Paragraphs.Add(new Paragraph());
                row.Cells.Add(cell);
            }
            if (row.Cells.Count > 0) t.Rows.Add(row);
        }
        if (cellBorders) t.Borders = true;
        int cols = t.Rows.Select(r => r.Cells.Sum(c => Math.Max(1, c.GridSpan))).DefaultIfEmpty(1).Max();
        while (t.ColumnWidths.Count < cols) t.ColumnWidths.Add(t.ColumnWidths.Count > 0 ? t.ColumnWidths.Average() : 72);
        if (t.ColumnWidths.All(w => w <= 0)) for (int i = 0; i < t.ColumnWidths.Count; i++) t.ColumnWidths[i] = 72;
        if (t.Rows.Count == 0) t.Rows.Add(new TableRow { Cells = [new TableCell { Paragraphs = [new Paragraph()] }] });
        return t;
    }

    private void ReadSection(XElement s)
    {
        var p = _doc.Page;
        var sz = s.Element(W + "pgSz");
        if (sz != null)
        {
            if (Twip(sz.Attribute(W + "w")) is float w) p.Width = w;
            if (Twip(sz.Attribute(W + "h")) is float h) p.Height = h;
        }
        var mar = s.Element(W + "pgMar");
        if (mar != null)
        {
            if (Twip(mar.Attribute(W + "top")) is float t) p.MarginTop = Math.Abs(t);
            if (Twip(mar.Attribute(W + "bottom")) is float b) p.MarginBottom = Math.Abs(b);
            if (Twip(mar.Attribute(W + "left") ?? mar.Attribute(W + "start")) is float l) p.MarginLeft = l + (Twip(mar.Attribute(W + "gutter")) ?? 0);
            if (Twip(mar.Attribute(W + "right") ?? mar.Attribute(W + "end")) is float r) p.MarginRight = r;
            if (Twip(mar.Attribute(W + "header")) is float hd) p.HeaderDistance = hd;
            if (Twip(mar.Attribute(W + "footer")) is float ft) p.FooterDistance = ft;
        }
        var paper = PageSetup.Papers.FirstOrDefault(x => Math.Abs(x.W - Math.Min(p.Width, p.Height)) < 3 && Math.Abs(x.H - Math.Max(p.Width, p.Height)) < 3);
        p.PaperName = paper.Name ?? "Custom";
        _doc.Header = ReadHeaderFooter(s, "headerReference", "hdr");
        _doc.Footer = ReadHeaderFooter(s, "footerReference", "ftr");
    }

    private List<Paragraph> ReadHeaderFooter(XElement sect, string refName, string root)
    {
        var refs = sect.Elements(W + refName).ToList();
        var r = refs.FirstOrDefault(x => (string?)x.Attribute(W + "type") == "default") ?? refs.FirstOrDefault();
        var id = (string?)r?.Attribute(R + "id");
        if (id == null || !_rels.TryGetValue(id, out var t)) return new();
        var part = Resolve(t.Target);
        var x = Load(part);
        if (x?.Root == null) return new();
        var saved = _rels;
        _rels = LoadRels(part);
        var list = new List<Paragraph>();
        try { ReadParagraphsInto(x.Root, list); }
        finally { _rels = saved; }
        // drop empty trailing paragraphs that only carried the watermark
        if (list.Count > 0 && list.All(p => p.Length == 0)) list.Clear();
        return list;
    }
}
