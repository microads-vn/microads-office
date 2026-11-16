using System.Drawing;
using System.Text;

namespace MicroadsOffice.Word.Model;

/// <summary>Editing operations on a <see cref="Document"/>. Positions are always (paragraph index, offset).</summary>
public static class Editing
{
    // ------------------------------------------------------------------ helpers

    public static DocPos Clamp(Document d, DocPos p)
    {
        var ps = d.Paragraphs;
        int i = Math.Clamp(p.Para, 0, ps.Count - 1);
        return new DocPos(i, Math.Clamp(p.Offset, 0, ps[i].Length));
    }

    public static DocPos End(Document d) => new(d.Paragraphs.Count - 1, d.Paragraphs[^1].Length);

    /// <summary>List holding the paragraph (body blocks are handled separately).</summary>
    private static List<Paragraph>? CellList(Paragraph p) => p.Cell?.Paragraphs;

    private static bool SameContainer(Paragraph a, Paragraph b) => a.Cell == b.Cell;

    private static void InsertParagraphAfter(Document d, Paragraph anchor, Paragraph np)
    {
        if (anchor.Cell is TableCell c) c.Paragraphs.Insert(c.Paragraphs.IndexOf(anchor) + 1, np);
        else d.Blocks.Insert(d.Blocks.IndexOf(anchor) + 1, np);
        d.StructureChanged();
    }

    private static void RemoveParagraph(Document d, Paragraph p)
    {
        if (p.Cell is TableCell c) { c.Paragraphs.Remove(p); if (c.Paragraphs.Count == 0) c.Paragraphs.Add(new Paragraph("", p.MarkProps, p.Props)); }
        else d.Blocks.Remove(p);
        d.StructureChanged();
    }

    // ------------------------------------------------------------------ text

    /// <summary>Inserts text ('\n' or "\r\n" = new paragraph) and returns the position after it.</summary>
    public static DocPos InsertText(Document d, DocPos pos, string text, CharProps props)
    {
        pos = Clamp(d, pos);
        text = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var parts = text.Split('\n');
        for (int i = 0; i < parts.Length; i++)
        {
            if (i > 0) pos = SplitParagraph(d, pos);
            if (parts[i].Length > 0) pos = InsertInline(d, pos, parts[i], props, null);
        }
        return pos;
    }

    /// <summary>Inserts text without paragraph breaks (or a single inline object).</summary>
    public static DocPos InsertInline(Document d, DocPos pos, string text, CharProps props, InlineObject? obj)
    {
        var p = d.Paragraphs[pos.Para];
        int idx = p.SplitAt(pos.Offset);
        p.Runs.Insert(idx, new Run(obj != null ? Run.ObjectChar.ToString() : text, props, obj));
        p.Normalize();
        return new DocPos(pos.Para, pos.Offset + (obj != null ? 1 : text.Length));
    }

    public static DocPos InsertObject(Document d, DocPos pos, InlineObject obj, CharProps props) => InsertInline(d, Clamp(d, pos), "", props, obj);

    /// <summary>Splits the paragraph at pos (Enter). The new paragraph keeps the formatting, except headings/titles which continue as Normal.</summary>
    public static DocPos SplitParagraph(Document d, DocPos pos)
    {
        pos = Clamp(d, pos);
        var p = d.Paragraphs[pos.Para];
        int idx = p.SplitAt(pos.Offset);
        var np = new Paragraph { Props = p.Props, MarkProps = p.PropsAt(pos.Offset) };
        np.Runs = p.Runs.Skip(idx).ToList();
        p.Runs = p.Runs.Take(idx).ToList();
        if (pos.Offset == 0 && p.Runs.Count == 0) p.MarkProps = np.MarkProps;
        if (np.Length == 0 && (p.Props.OutlineLevel > 0 || p.Props.Style is StyleIds.Title or StyleIds.Subtitle or StyleIds.TocHeading))
        {
            var normal = d.Styles[StyleIds.Normal];
            np.Props = normal.Para;
            np.MarkProps = normal.Char;
        }
        np.Props = np.Props with { PageBreakBefore = false };
        p.Normalize(); np.Normalize();
        InsertParagraphAfter(d, p, np);
        return new DocPos(pos.Para + 1, 0);
    }

    /// <summary>Deletes [a, b) and returns the resulting caret position.</summary>
    public static DocPos Delete(Document d, DocPos a, DocPos b)
    {
        if (b < a) (a, b) = (b, a);
        a = Clamp(d, a); b = Clamp(d, b);
        if (a == b) return a;
        var ps = d.Paragraphs.ToList();
        var A = ps[a.Para];
        var B = ps[b.Para];
        if (a.Para == b.Para)
        {
            DeleteInPara(A, a.Offset, b.Offset);
            return a;
        }
        if (SameContainer(A, B))
        {
            DeleteInPara(A, a.Offset, A.Length);
            var tail = TailRuns(B, b.Offset);
            // remove everything strictly between A and B
            if (A.Cell is TableCell c)
            {
                int ia = c.Paragraphs.IndexOf(A), ib = c.Paragraphs.IndexOf(B);
                c.Paragraphs.RemoveRange(ia + 1, ib - ia);
            }
            else
            {
                int ia = d.Blocks.IndexOf(A), ib = d.Blocks.IndexOf(B);
                d.Blocks.RemoveRange(ia + 1, ib - ia);
            }
            if (A.Length == 0 && A.Runs.Count == 0 && tail.Count > 0) A.MarkProps = tail[0].Props;
            A.Runs.AddRange(tail);
            A.Normalize();
            d.StructureChanged();
            return new DocPos(d.IndexOf(A), a.Offset);
        }
        // Different containers (selection crosses table borders).
        var removeTables = new HashSet<Table>();
        foreach (var t in d.Blocks.OfType<Table>())
        {
            var tp = t.Rows.SelectMany(r => r.Cells).SelectMany(c => c.Paragraphs).ToList();
            if (tp.Count == 0) continue;
            int first = ps.IndexOf(tp[0]), last = ps.IndexOf(tp[^1]);
            if (first >= 0 && last >= 0 && new DocPos(first, 0) >= a && new DocPos(last, tp[^1].Length) <= b && !(first == a.Para && a.Offset > 0)) removeTables.Add(t);
        }
        for (int i = a.Para; i <= b.Para; i++)
        {
            var p = ps[i];
            int s = i == a.Para ? a.Offset : 0, e = i == b.Para ? b.Offset : p.Length;
            if (p.Cell?.Row?.Table is Table tt && removeTables.Contains(tt)) continue;
            if (p.Cell == null && i != a.Para && i != b.Para) { d.Blocks.Remove(p); continue; }
            DeleteInPara(p, s, e);
        }
        foreach (var t in removeTables) d.Blocks.Remove(t);
        d.StructureChanged();
        if (A.Cell == null && B.Cell == null && d.Blocks.Contains(A) && d.Blocks.Contains(B))
        {
            int ia = d.Blocks.IndexOf(A), ib = d.Blocks.IndexOf(B);
            if (ib == ia + 1)
            {
                A.Runs.AddRange(B.Runs);
                A.Normalize();
                d.Blocks.Remove(B);
                d.StructureChanged();
            }
        }
        int ai = d.IndexOf(A);
        if (ai >= 0) return Clamp(d, new DocPos(ai, a.Offset));
        int bi = d.IndexOf(B);
        return Clamp(d, new DocPos(bi >= 0 ? bi : a.Para, 0));
    }

    private static List<Run> TailRuns(Paragraph p, int offset)
    {
        int idx = p.SplitAt(offset);
        return p.Runs.Skip(idx).ToList();
    }

    public static void DeleteInPara(Paragraph p, int s, int e)
    {
        if (e <= s) return;
        var keepProps = p.PropsAt(s == 0 ? 0 : s);
        int i1 = p.SplitAt(s);
        int i2 = p.SplitAt(e);
        p.Runs.RemoveRange(i1, i2 - i1);
        if (p.Runs.Count == 0) p.MarkProps = keepProps;
        p.Normalize();
    }

    /// <summary>Merges paragraph i+1 into i (Delete at end / Backspace at start). Returns false when not allowed (table boundary).</summary>
    public static bool JoinWithNext(Document d, int i)
    {
        var ps = d.Paragraphs;
        if (i < 0 || i + 1 >= ps.Count) return false;
        var A = ps[i]; var B = ps[i + 1];
        if (!SameContainer(A, B)) return false;
        if (A.Cell == null && d.Blocks.IndexOf(B) != d.Blocks.IndexOf(A) + 1) return false;
        if (A.Length == 0 && A.Runs.Count == 0) { A.Props = A.Props; }
        A.Runs.AddRange(B.Runs);
        A.Normalize();
        RemoveParagraph(d, B);
        return true;
    }

    public static string GetText(Document d, DocPos a, DocPos b)
    {
        if (b < a) (a, b) = (b, a);
        var sb = new StringBuilder();
        var ps = d.Paragraphs;
        for (int i = a.Para; i <= b.Para && i < ps.Count; i++)
        {
            var t = ps[i].Text;
            int s = i == a.Para ? Math.Min(a.Offset, t.Length) : 0, e = i == b.Para ? Math.Min(b.Offset, t.Length) : t.Length;
            if (i > a.Para) sb.Append(ps[i].Cell != null && ps[i - 1].Cell != null && ps[i].Cell != ps[i - 1].Cell ? '\t' : '\n');
            sb.Append(t.AsSpan(s, Math.Max(0, e - s)));
        }
        return sb.ToString().Replace(Run.ObjectChar.ToString(), "").Replace(Run.LineBreak, '\n').Replace(Run.PageBreak, '\n');
    }

    // ------------------------------------------------------------------ formatting

    /// <summary>Applies a character format change to [a, b).</summary>
    public static void ApplyChar(Document d, DocPos a, DocPos b, Func<CharProps, CharProps> f)
    {
        if (b < a) (a, b) = (b, a);
        var ps = d.Paragraphs;
        for (int i = a.Para; i <= b.Para; i++)
        {
            var p = ps[i];
            int s = i == a.Para ? a.Offset : 0, e = i == b.Para ? b.Offset : p.Length;
            if (p.Length == 0 || (s == 0 && e >= p.Length)) p.MarkProps = f(p.MarkProps);
            if (e <= s) continue;
            int i1 = p.SplitAt(s), i2 = p.SplitAt(e);
            for (int k = i1; k < i2; k++) p.Runs[k].Props = f(p.Runs[k].Props);
            p.Normalize();
        }
    }

    public static void ApplyPara(Document d, DocPos a, DocPos b, Func<ParaProps, ParaProps> f)
    {
        if (b < a) (a, b) = (b, a);
        var ps = d.Paragraphs;
        for (int i = a.Para; i <= b.Para; i++) { ps[i].Props = f(ps[i].Props); ps[i].Touch(); }
    }

    /// <summary>Applies a named style: paragraph formatting and the characters' formatting are reset to the style.</summary>
    public static void ApplyStyle(Document d, DocPos a, DocPos b, string styleId)
    {
        if (!d.Styles.TryGetValue(styleId, out var s)) return;
        if (b < a) (a, b) = (b, a);
        var ps = d.Paragraphs;
        for (int i = a.Para; i <= b.Para; i++)
        {
            var p = ps[i];
            var keepList = p.Props;
            p.Props = s.Para with { List = keepList.List, ListLevel = keepList.ListLevel, ListId = keepList.ListId, IndentLeft = keepList.List != ListType.None ? keepList.IndentLeft : s.Para.IndentLeft, FirstLine = keepList.List != ListType.None ? keepList.FirstLine : s.Para.FirstLine };
            foreach (var r in p.Runs) r.Props = StyleChar(r.Props, s.Char);
            p.MarkProps = StyleChar(p.MarkProps, s.Char);
            p.Normalize();
        }
    }

    private static CharProps StyleChar(CharProps old, CharProps style) => style with { Link = old.Link, Highlight = old.Highlight, Underline = old.Link != null ? UnderlineStyle.Single : style.Underline, Color = old.Link != null ? old.Color : style.Color };

    /// <summary>Character formatting of a range (formatting at the start) — used by the ribbon state.</summary>
    public static CharProps CharAt(Document d, DocPos p)
    {
        p = Clamp(d, p);
        var para = d.Paragraphs[p.Para];
        if (para.Runs.Count == 0) return para.MarkProps;
        return p.Offset == 0 ? para.Runs[0].Props : para.PropsAt(p.Offset);
    }

    // ------------------------------------------------------------------ lists

    private static int _nextListId = 1;

    public static int NewListId(Document d)
    {
        int max = d.Paragraphs.Select(p => p.Props.ListId).DefaultIfEmpty(0).Max();
        _nextListId = Math.Max(_nextListId, max + 1);
        return _nextListId++;
    }

    /// <summary>Toggles a list type on the paragraphs; continues the list of the previous paragraph when it uses the same type.</summary>
    public static void ToggleList(Document d, DocPos a, DocPos b, ListType type)
    {
        if (b < a) (a, b) = (b, a);
        var ps = d.Paragraphs;
        bool allSame = Enumerable.Range(a.Para, b.Para - a.Para + 1).All(i => ps[i].Props.List == type);
        if (allSame)
        {
            for (int i = a.Para; i <= b.Para; i++) { ps[i].Props = ps[i].Props with { List = ListType.None, ListLevel = 0, IndentLeft = 0, FirstLine = 0 }; ps[i].Touch(); }
            return;
        }
        int id = a.Para > 0 && ps[a.Para - 1].Props.List == type && ps[a.Para - 1].Cell == ps[a.Para].Cell ? ps[a.Para - 1].Props.ListId : NewListId(d);
        for (int i = a.Para; i <= b.Para; i++)
        {
            int lvl = ps[i].Props.List != ListType.None ? ps[i].Props.ListLevel : 0;
            ps[i].Props = ps[i].Props with { List = type, ListId = id, ListLevel = lvl, IndentLeft = 18 * (lvl + 1) + 18 * lvl * 0, FirstLine = -18 };
            ps[i].Props = ps[i].Props with { IndentLeft = 36 + 18 * lvl, FirstLine = -18 };
            ps[i].Touch();
        }
    }

    public static void ChangeListLevel(Document d, DocPos a, DocPos b, int delta)
    {
        if (b < a) (a, b) = (b, a);
        var ps = d.Paragraphs;
        for (int i = a.Para; i <= b.Para; i++)
        {
            var p = ps[i];
            if (p.Props.List != ListType.None)
            {
                int lvl = Math.Clamp(p.Props.ListLevel + delta, 0, 8);
                var type = p.Props.List;
                if (type != ListType.Bullet && delta != 0)
                    type = (lvl % 3) switch { 0 => ListType.Decimal, 1 => ListType.LowerAlpha, _ => ListType.LowerRoman };
                p.Props = p.Props with { ListLevel = lvl, List = type, IndentLeft = 36 + 18 * lvl };
            }
            else p.Props = p.Props with { IndentLeft = Math.Max(0, p.Props.IndentLeft + delta * 36) };
            p.Touch();
        }
    }

    // ------------------------------------------------------------------ fragments (rich copy / paste)

    /// <summary>Copies [a, b) as a list of paragraphs (tables are flattened to tab-separated rows when partially selected).</summary>
    public static List<Block> CopyRange(Document d, DocPos a, DocPos b)
    {
        if (b < a) (a, b) = (b, a);
        var result = new List<Block>();
        var ps = d.Paragraphs;
        var copiedTables = new HashSet<Table>();
        for (int i = a.Para; i <= b.Para; i++)
        {
            var p = ps[i];
            if (p.Cell?.Row?.Table is Table t)
            {
                var tp = t.Rows.SelectMany(r => r.Cells).SelectMany(c => c.Paragraphs).ToList();
                int first = d.IndexOf(tp[0]), last = d.IndexOf(tp[^1]);
                if (first >= a.Para && last <= b.Para && (first > a.Para || a.Offset == 0))
                {
                    if (copiedTables.Add(t)) result.Add(t.CloneBlock());
                    continue;
                }
            }
            int s = i == a.Para ? a.Offset : 0, e = i == b.Para ? b.Offset : p.Length;
            var c = p.Clone();
            DeleteInPara(c, e, c.Length);
            DeleteInPara(c, 0, s);
            result.Add(c);
        }
        return result;
    }

    /// <summary>Inserts copied blocks at pos (the first paragraph merges with the current one).</summary>
    public static DocPos PasteBlocks(Document d, DocPos pos, List<Block> blocks)
    {
        pos = Clamp(d, pos);
        if (blocks.Count == 0) return pos;
        bool inCell = d.Paragraphs[pos.Para].Cell != null;
        for (int bi = 0; bi < blocks.Count; bi++)
        {
            var blk = blocks[bi].CloneBlock();
            if (blk is Table t)
            {
                if (inCell)
                {
                    // nested tables are not supported: paste the text
                    foreach (var row in t.Rows)
                    {
                        var line = string.Join("\t", row.Cells.Select(c => string.Join(" ", c.Paragraphs.Select(x => x.Text))));
                        pos = InsertText(d, pos, line, d.Paragraphs[pos.Para].PropsAt(pos.Offset));
                        pos = SplitParagraph(d, pos);
                    }
                    continue;
                }
                pos = InsertTableBlock(d, pos, t);
                continue;
            }
            var src = (Paragraph)blk;
            var p = d.Paragraphs[pos.Para];
            if (bi > 0) { pos = SplitParagraph(d, pos); p = d.Paragraphs[pos.Para]; if (p.Length == 0) p.Props = src.Props; }
            else if (p.Length == 0 && blocks.Count > 1) p.Props = src.Props;
            int idx = p.SplitAt(pos.Offset);
            p.Runs.InsertRange(idx, src.Runs.Select(r => r.Clone()));
            int len = src.Length;
            p.Normalize();
            pos = new DocPos(pos.Para, pos.Offset + len);
        }
        d.StructureChanged();
        return pos;
    }

    // ------------------------------------------------------------------ tables

    /// <summary>Inserts a table at the caret (splitting the paragraph if needed); returns the first cell position.</summary>
    public static DocPos InsertTable(Document d, DocPos pos, int rows, int cols)
    {
        var cp = Clamp(d, pos);
        var p = d.Paragraphs[cp.Para];
        var t = Table.Create(rows, cols, d.Page.ContentWidth - (p.Cell != null ? 0 : 0), p.PropsAt(cp.Offset) with { Link = null });
        return InsertTableBlock(d, cp, t);
    }

    private static DocPos InsertTableBlock(Document d, DocPos pos, Table t)
    {
        var p = d.Paragraphs[pos.Para];
        if (p.Cell != null) return pos; // no nested tables
        int bi = d.Blocks.IndexOf(p);
        if (pos.Offset == 0 && p.Length > 0) d.Blocks.Insert(bi, t);
        else if (pos.Offset == 0 && p.Length == 0) d.Blocks.Insert(bi, t);
        else
        {
            SplitParagraph(d, pos);
            bi = d.Blocks.IndexOf(p);
            d.Blocks.Insert(bi + 1, t);
        }
        d.StructureChanged();
        var firstPara = t.Rows[0].Cells[0].Paragraphs[0];
        return new DocPos(d.IndexOf(firstPara), 0);
    }

    public static TableCell? CellOf(Document d, DocPos p) => d.Paragraphs[Clamp(d, p).Para].Cell;

    public static Paragraph NewCellParagraph(CharProps cp) => new("", cp, new ParaProps { SpaceAfter = 0, LineSpacing = 1f });

    public static void InsertRow(Document d, TableCell cell, bool below)
    {
        _ = d.Paragraphs; // make sure row/table back-links are set
        var row = cell.Row!; var t = row.Table!;
        var nr = new TableRow { MinHeight = row.MinHeight };
        foreach (var c in row.Cells)
        {
            var nc = new TableCell { GridSpan = c.GridSpan, Shading = c.Shading, VAlign = c.VAlign };
            nc.Paragraphs.Add(NewCellParagraph(c.Paragraphs.FirstOrDefault()?.MarkProps ?? CharProps.Default));
            nr.Cells.Add(nc);
        }
        t.Rows.Insert(t.Rows.IndexOf(row) + (below ? 1 : 0), nr);
        d.StructureChanged();
    }

    public static void InsertColumn(Document d, TableCell cell, bool right)
    {
        _ = d.Paragraphs; // make sure row/table back-links are set
        var row = cell.Row!; var t = row.Table!;
        int g = t.GridColumnOf(row, cell) + (right ? Math.Max(1, cell.GridSpan) : 0);
        float total = t.ColumnWidths.Sum();
        t.ColumnWidths.Insert(Math.Min(g, t.ColumnWidths.Count), total / (t.ColumnWidths.Count + 1));
        float scale = total / t.ColumnWidths.Sum();
        for (int i = 0; i < t.ColumnWidths.Count; i++) t.ColumnWidths[i] *= scale;
        foreach (var r in t.Rows)
        {
            int col = 0, idx = 0;
            foreach (var c in r.Cells) { if (col >= g) break; col += Math.Max(1, c.GridSpan); idx++; }
            var nc = new TableCell();
            nc.Paragraphs.Add(NewCellParagraph(r.Cells.FirstOrDefault()?.Paragraphs.FirstOrDefault()?.MarkProps ?? CharProps.Default));
            r.Cells.Insert(Math.Min(idx, r.Cells.Count), nc);
        }
        d.StructureChanged();
    }

    public static void DeleteRow(Document d, TableCell cell)
    {
        _ = d.Paragraphs; // make sure row/table back-links are set
        var row = cell.Row!; var t = row.Table!;
        if (t.Rows.Count <= 1) { DeleteTable(d, t); return; }
        t.Rows.Remove(row);
        d.StructureChanged();
    }

    public static void DeleteColumn(Document d, TableCell cell)
    {
        _ = d.Paragraphs; // make sure row/table back-links are set
        var row = cell.Row!; var t = row.Table!;
        int g = t.GridColumnOf(row, cell);
        if (t.ColumnWidths.Count <= 1) { DeleteTable(d, t); return; }
        foreach (var r in t.Rows)
        {
            int col = 0;
            foreach (var c in r.Cells.ToList())
            {
                int span = Math.Max(1, c.GridSpan);
                if (g >= col && g < col + span)
                {
                    if (span > 1) c.GridSpan--; else r.Cells.Remove(c);
                    break;
                }
                col += span;
            }
            if (r.Cells.Count == 0) { var nc = new TableCell(); nc.Paragraphs.Add(new Paragraph()); r.Cells.Add(nc); }
        }
        float w = t.ColumnWidths[g];
        t.ColumnWidths.RemoveAt(g);
        if (t.ColumnWidths.Count > 0) t.ColumnWidths[Math.Min(g, t.ColumnWidths.Count - 1)] += w;
        d.StructureChanged();
    }

    public static void DeleteTable(Document d, Table t)
    {
        int i = d.Blocks.IndexOf(t);
        d.Blocks.Remove(t);
        if (d.Blocks.Count == 0 || i >= d.Blocks.Count || d.Blocks[Math.Min(i, d.Blocks.Count - 1)] is Table)
            d.Blocks.Insert(Math.Min(i, d.Blocks.Count), new Paragraph("", d.Styles[StyleIds.Normal].Char, d.Styles[StyleIds.Normal].Para));
        d.StructureChanged();
    }

    /// <summary>Merges the rectangular cell range spanned by two cells of the same table (contents are concatenated).</summary>
    public static bool MergeCells(Document d, TableCell c1, TableCell c2)
    {
        _ = d.Paragraphs; // make sure row/table back-links are set
        var t = c1.Row?.Table;
        if (t == null || t != c2.Row?.Table) return false;
        int r1 = t.Rows.IndexOf(c1.Row!), r2 = t.Rows.IndexOf(c2.Row!);
        int g1 = t.GridColumnOf(c1.Row!, c1), g2 = t.GridColumnOf(c2.Row!, c2);
        int g1e = g1 + Math.Max(1, c1.GridSpan) - 1, g2e = g2 + Math.Max(1, c2.GridSpan) - 1;
        if (r1 > r2) (r1, r2) = (r2, r1);
        int gs = Math.Min(g1, g2), ge = Math.Max(g1e, g2e);
        if (r1 == r2 && gs == ge) return false;
        TableCell? master = null;
        for (int r = r1; r <= r2; r++)
        {
            var row = t.Rows[r];
            int col = 0;
            var inRange = new List<TableCell>();
            foreach (var c in row.Cells)
            {
                int span = Math.Max(1, c.GridSpan);
                if (col >= gs && col + span - 1 <= ge) inRange.Add(c);
                col += span;
            }
            if (inRange.Count == 0) continue;
            var first = inRange[0];
            for (int k = 1; k < inRange.Count; k++)
            {
                AppendCellContent(first, inRange[k]);
                first.GridSpan += Math.Max(1, inRange[k].GridSpan);
                row.Cells.Remove(inRange[k]);
            }
            first.GridSpan = Math.Max(first.GridSpan, ge - gs + 1);
            if (r == r1) { master = first; first.VMerge = r2 > r1 ? VMerge.Restart : VMerge.None; }
            else
            {
                AppendCellContent(master!, first);
                first.Paragraphs = [NewCellParagraph(first.Paragraphs.FirstOrDefault()?.MarkProps ?? CharProps.Default)];
                first.VMerge = VMerge.Continue;
            }
        }
        d.StructureChanged();
        return true;
    }

    private static void AppendCellContent(TableCell into, TableCell from)
    {
        var real = from.Paragraphs.Where(p => p.Length > 0).ToList();
        if (real.Count == 0) return;
        if (into.Paragraphs.Count == 1 && into.Paragraphs[0].Length == 0) into.Paragraphs.Clear();
        into.Paragraphs.AddRange(real);
    }

    // ------------------------------------------------------------------ search

    public static IEnumerable<(DocPos Start, DocPos End)> FindAll(Document d, string what, bool matchCase, bool wholeWord)
    {
        if (string.IsNullOrEmpty(what)) yield break;
        var cmp = matchCase ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;
        var ps = d.Paragraphs;
        for (int i = 0; i < ps.Count; i++)
        {
            var t = ps[i].Text;
            int idx = 0;
            while ((idx = t.IndexOf(what, idx, cmp)) >= 0)
            {
                bool ok = !wholeWord || ((idx == 0 || !char.IsLetterOrDigit(t[idx - 1])) && (idx + what.Length >= t.Length || !char.IsLetterOrDigit(t[idx + what.Length])));
                if (ok) yield return (new DocPos(i, idx), new DocPos(i, idx + what.Length));
                idx += Math.Max(1, what.Length);
            }
        }
    }

    // ------------------------------------------------------------------ word boundaries

    public static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '\'';

    public static (int Start, int End) WordAt(Paragraph p, int offset)
    {
        var t = p.Text;
        if (t.Length == 0) return (0, 0);
        offset = Math.Clamp(offset, 0, t.Length);
        int s = offset, e = offset;
        if (s < t.Length && IsWordChar(t[s]) || (s > 0 && IsWordChar(t[s - 1])))
        {
            while (s > 0 && IsWordChar(t[s - 1])) s--;
            while (e < t.Length && IsWordChar(t[e])) e++;
            while (e < t.Length && t[e] == ' ') e++;
        }
        else if (s < t.Length) e = s + 1;
        return (s, e);
    }

    public static int NextWord(string t, int offset)
    {
        int i = offset;
        if (i < t.Length && IsWordChar(t[i])) while (i < t.Length && IsWordChar(t[i])) i++;
        else if (i < t.Length && !char.IsWhiteSpace(t[i])) i++;
        while (i < t.Length && char.IsWhiteSpace(t[i])) i++;
        return i;
    }

    public static int PrevWord(string t, int offset)
    {
        int i = offset;
        while (i > 0 && char.IsWhiteSpace(t[i - 1])) i--;
        if (i > 0 && IsWordChar(t[i - 1])) while (i > 0 && IsWordChar(t[i - 1])) i--;
        else if (i > 0) i--;
        return i;
    }

    // ------------------------------------------------------------------ statistics

    public static (int Words, int Chars, int CharsNoSpaces, int Paragraphs) Count(Document d, DocPos? a = null, DocPos? b = null)
    {
        string text = a.HasValue && b.HasValue && a != b ? GetText(d, a.Value, b.Value) : d.PlainText;
        int words = 0; bool inWord = false;
        foreach (var c in text)
        {
            bool w = !char.IsWhiteSpace(c);
            if (w && !inWord) words++;
            inWord = w;
        }
        int paras = text.Split('\n').Count(s => s.Trim().Length > 0);
        return (words, text.Count(c => c != '\n'), text.Count(c => !char.IsWhiteSpace(c)), paras);
    }

    public static string ChangeCase(string s, int mode) => mode switch
    {
        0 => s.ToUpper(System.Globalization.CultureInfo.CurrentCulture),
        1 => s.ToLower(System.Globalization.CultureInfo.CurrentCulture),
        2 => System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower()),
        _ => s.Length > 0 ? char.ToUpper(s[0]) + s[1..].ToLower() : s
    };

    /// <summary>Replaces characters in [a,b) one by one keeping formatting (used by Change Case).</summary>
    public static void TransformText(Document d, DocPos a, DocPos b, Func<string, string> f)
    {
        if (b < a) (a, b) = (b, a);
        var ps = d.Paragraphs;
        for (int i = a.Para; i <= b.Para; i++)
        {
            var p = ps[i];
            int s = i == a.Para ? a.Offset : 0, e = i == b.Para ? b.Offset : p.Length;
            if (e <= s) continue;
            int i1 = p.SplitAt(s), i2 = p.SplitAt(e);
            for (int k = i1; k < i2; k++)
            {
                if (p.Runs[k].Object != null) continue;
                var nt = f(p.Runs[k].Text);
                if (nt.Length == p.Runs[k].Text.Length) p.Runs[k].Text = nt;
            }
            p.Normalize();
        }
    }

    public static Color? ParseColor(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex) || hex == "auto") return null;
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length != 6) return null;
            return Color.FromArgb(Convert.ToInt32(hex[..2], 16), Convert.ToInt32(hex.Substring(2, 2), 16), Convert.ToInt32(hex.Substring(4, 2), 16));
        }
        catch { return null; }
    }

    public static string Hex(Color c) => $"{c.R:X2}{c.G:X2}{c.B:X2}";
}
