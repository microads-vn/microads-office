using System.Drawing;
using MicroadsOffice.Word.Editor;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.Tests;

/// <summary>Document model, editing operations, layout and pagination.</summary>
public static class ModelTests
{
    public static Document Sample(int paras = 3)
    {
        var d = Document.CreateBlank();
        d.Blocks.Clear();
        var n = d.Styles[StyleIds.Normal];
        for (int i = 0; i < paras; i++) d.Blocks.Add(new Paragraph($"Đoạn văn số {i + 1} có tiếng Việt: Cộng hòa xã hội chủ nghĩa Việt Nam.", n.Char, n.Para));
        d.StructureChanged();
        return d;
    }

    public static void Run(Runner r)
    {
        r.Test("M01 insert text and split paragraph", () =>
        {
            var d = Document.CreateBlank();
            var p = Editing.InsertText(d, new DocPos(0, 0), "Xin chào\nthế giới", CharProps.Default);
            T.Eq(2, d.Paragraphs.Count, "paragraphs");
            T.Eq("Xin chào", d.Paragraphs[0].Text);
            T.Eq("thế giới", d.Paragraphs[1].Text);
            T.Eq(new DocPos(1, 8), p, "caret after insert");
        });

        r.Test("M02 delete across paragraphs merges them", () =>
        {
            var d = Sample(3);
            var end = Editing.Delete(d, new DocPos(0, 5), new DocPos(2, 3));
            T.Eq(1, d.Paragraphs.Count);
            T.True(d.Paragraphs[0].Text.StartsWith("Đoạn") && d.Paragraphs[0].Text.Contains("n số 3"), "merged text: " + d.Paragraphs[0].Text);
            T.Eq(new DocPos(0, 5), end);
        });

        r.Test("M03 character formatting splits and merges runs", () =>
        {
            var d = Sample(1);
            Editing.ApplyChar(d, new DocPos(0, 5), new DocPos(0, 10), c => c with { Bold = true });
            var p = d.Paragraphs[0];
            T.Eq(3, p.Runs.Count, "runs after bold");
            T.True(p.Runs[1].Props.Bold && p.Runs[1].Text.Length == 5, "bold run");
            Editing.ApplyChar(d, new DocPos(0, 0), new DocPos(0, p.Length), c => c with { Bold = false });
            T.Eq(1, p.Runs.Count, "runs merged again");
        });

        r.Test("M04 paragraph formatting, styles and headings", () =>
        {
            var d = Sample(2);
            Editing.ApplyStyle(d, new DocPos(0, 0), new DocPos(0, 0), StyleIds.Heading1);
            T.Eq(1, d.Paragraphs[0].Props.OutlineLevel, "heading outline");
            T.True(d.Paragraphs[0].Runs.All(x => x.Props.Bold), "heading bold");
            Editing.ApplyPara(d, new DocPos(1, 0), new DocPos(1, 0), p => p with { Align = ParaAlign.Justify, FirstLine = 28 });
            T.Eq(ParaAlign.Justify, d.Paragraphs[1].Props.Align);
            var np = Editing.SplitParagraph(d, new DocPos(0, d.Paragraphs[0].Length));
            T.Eq(StyleIds.Normal, d.Paragraphs[np.Para].Props.Style, "Enter after heading continues with Normal");
        });

        r.Test("M05 bullets and multi-level numbering labels", () =>
        {
            var d = Sample(5);
            Editing.ToggleList(d, new DocPos(0, 0), new DocPos(3, 0), ListType.Decimal);
            Editing.ChangeListLevel(d, new DocPos(1, 0), new DocPos(2, 0), 1);
            var labels = DocLayout.ComputeListLabels(d);
            T.Eq("1.", labels[d.Paragraphs[0]]);
            T.Eq("a)", labels[d.Paragraphs[1]]);
            T.Eq("b)", labels[d.Paragraphs[2]]);
            T.Eq("2.", labels[d.Paragraphs[3]]);
            Editing.ToggleList(d, new DocPos(4, 0), new DocPos(4, 0), ListType.Bullet);
            T.Eq("•", DocLayout.ComputeListLabels(d)[d.Paragraphs[4]]);
        });

        r.Test("M06 tables: insert, rows, columns, merge, delete", () =>
        {
            var d = Sample(1);
            var p = Editing.InsertTable(d, new DocPos(0, d.Paragraphs[0].Length), 2, 3);
            var t = d.Blocks.OfType<Table>().Single();
            T.Eq(2, t.Rows.Count); T.Eq(3, t.ColumnCount);
            var cell = Editing.CellOf(d, p)!;
            Editing.InsertRow(d, cell, true);
            Editing.InsertColumn(d, cell, true);
            T.Eq(3, t.Rows.Count); T.Eq(4, t.ColumnWidths.Count);
            T.True(t.Rows.All(row => row.Cells.Count == 4), "each row has 4 cells");
            T.True(Editing.MergeCells(d, t.Rows[0].Cells[0], t.Rows[1].Cells[1]), "merge 2x2");
            T.Eq(2, t.Rows[0].Cells[0].GridSpan, "gridspan");
            T.Eq(VMerge.Continue, t.Rows[1].Cells[0].VMerge, "vmerge continue");
            Editing.DeleteRow(d, t.Rows[2].Cells[0]);
            T.Eq(2, t.Rows.Count);
            Editing.DeleteTable(d, t);
            T.Eq(0, d.Blocks.OfType<Table>().Count());
            T.True(d.Blocks.Count >= 1, "paragraph remains");
        });

        r.Test("M07 find / replace / word count / change case", () =>
        {
            var d = Sample(3);
            var hits = Editing.FindAll(d, "việt nam", false, true).ToList();
            T.Eq(3, hits.Count, "case-insensitive whole word");
            T.Eq(0, Editing.FindAll(d, "việt nam", true, true).Count(), "match case");
            var (w, _, _, ps) = Editing.Count(d);
            T.True(w > 30 && ps == 3, $"words {w}, paragraphs {ps}");
            Editing.TransformText(d, new DocPos(0, 0), new DocPos(0, 4), s => Editing.ChangeCase(s, 0));
            T.True(d.Paragraphs[0].Text.StartsWith("ĐOẠN"), "uppercase vietnamese: " + d.Paragraphs[0].Text[..6]);
        });

        r.Test("M08 copy range / paste blocks keeps formatting", () =>
        {
            var d = Sample(2);
            Editing.ApplyChar(d, new DocPos(0, 0), new DocPos(0, 4), c => c with { Italic = true, Color = Color.Red });
            var frag = Editing.CopyRange(d, new DocPos(0, 0), new DocPos(1, 4));
            T.Eq(2, frag.Count);
            var d2 = Document.CreateBlank();
            var end = Editing.PasteBlocks(d2, new DocPos(0, 0), frag);
            T.Eq(2, d2.Paragraphs.Count);
            T.True(d2.Paragraphs[0].Runs[0].Props.Italic && d2.Paragraphs[0].Runs[0].Props.Color?.ToArgb() == Color.Red.ToArgb(), "formatting kept");
            T.Eq(new DocPos(1, 4), end);
        });

        r.Test("M09 layout: wrapping stays inside margins, justify fills the line", () =>
        {
            var d = Sample(1);
            var p = d.Paragraphs[0];
            p.Runs[0].Text = string.Concat(Enumerable.Repeat("Microads Office soạn thảo văn bản chuyên nghiệp. ", 12));
            p.Props = p.Props with { Align = ParaAlign.Justify };
            p.Touch();
            using var m = new TextMeasurer();
            var L = new DocLayout(d, m);
            var lines = L.ByPara[p];
            T.True(lines.Count >= 3, "wrapped into several lines: " + lines.Count);
            foreach (var l in lines)
            {
                var right = l.Line.Segs.Where(s => s.Kind != SegKind.Space).Max(s => s.X + s.W);
                T.True(right <= d.Page.ContentWidth + 0.5f, $"line inside margin ({right} <= {d.Page.ContentWidth})");
            }
            var first = lines[0].Line;
            var lastText = first.Segs.Last(s => s.Kind == SegKind.Text);
            T.Near(d.Page.ContentWidth, lastText.X + lastText.W, 1.0f, "justified first line reaches the right margin");
        });

        r.Test("M10 pagination: long document, page break, table rows across pages", () =>
        {
            var d = Sample(120);
            using var m = new TextMeasurer();
            var L = new DocLayout(d, m);
            T.True(L.Pages.Count >= 4, "pages: " + L.Pages.Count);
            foreach (var pg in L.Pages)
                foreach (var l in pg.Lines)
                    T.True(l.Y + l.Line.Height <= d.Page.Height - d.Page.MarginBottom + 0.5f, "line above bottom margin");
            int before = L.Pages.Count;
            Editing.InsertText(d, new DocPos(0, 3), "\f", CharProps.Default);
            var L2 = new DocLayout(d, m);
            T.True(L2.Pages.Count >= before, "page break adds a page");
            T.Eq(1, L2.PageOf(new DocPos(1, 0)), "paragraph after the page break starts on page 2");
            var d3 = Sample(1);
            Editing.InsertTable(d3, new DocPos(0, d3.Paragraphs[0].Length), 80, 3);
            var L3 = new DocLayout(d3, m);
            T.True(L3.Pages.Count >= 2, "80-row table spans pages: " + L3.Pages.Count);
            T.True(L3.Pages[1].Cells.Count > 0, "cells continue on page 2");
        });

        r.Test("M11 caret geometry and hit testing round-trip", () =>
        {
            var d = Sample(4);
            using var m = new TextMeasurer();
            var L = new DocLayout(d, m);
            foreach (var pos in new[] { new DocPos(0, 0), new DocPos(1, 10), new DocPos(3, 20) })
            {
                var (page, x, y, h) = L.CaretAt(pos);
                var hit = L.HitTest(page, new PointF(x + 0.1f, y + h / 2));
                T.Eq(pos, hit, "hit test at caret");
            }
            var down = L.Vertical(new DocPos(0, 5), L.CaretAt(new DocPos(0, 5)).X, 1);
            T.True(down is { } dp && dp.Para >= 0 && dp > new DocPos(0, 5), "arrow down moves forward");
        });

        r.Test("M12 editor: typing, undo coalescing, redo, backspace, enter", () =>
        {
            using var ed = new WordEditor();
            ed.CreateControl();
            foreach (var ch in "Gõ tiếng Việt") ed.InsertText(ch.ToString());
            T.Eq("Gõ tiếng Việt", ed.Document.Paragraphs[0].Text);
            ed.InsertParagraphBreak();
            ed.InsertText("dòng 2");
            ed.Backspace(false);
            T.Eq("dòng ", ed.Document.Paragraphs[1].Text);
            ed.Undo();
            T.Eq("dòng 2", ed.Document.Paragraphs[1].Text, "undo restores");
            ed.Undo(); ed.Undo();
            T.Eq(1, ed.Document.Paragraphs.Count, "undo enter");
            T.Eq("Gõ tiếng Việt", ed.Document.Paragraphs[0].Text, "typing burst is one undo step after enter is undone");
            ed.Redo();
            T.Eq(2, ed.Document.Paragraphs.Count, "redo");
            T.True(ed.Modified, "modified flag");
            ed.MarkSaved();
            T.True(!ed.Modified, "saved");
        });

        r.Test("M13 editor: bold applies to word at caret, pending format for typing", () =>
        {
            using var ed = new WordEditor();
            ed.CreateControl();
            ed.InsertText("hello world");
            ed.SetCaret(new DocPos(0, 2), false);
            ed.ApplyChar(p => p with { Bold = true });
            var p0 = ed.Document.Paragraphs[0];
            T.True(p0.Runs[0].Props.Bold && p0.Runs[0].Text == "hello", "word under caret is bold");
            ed.SetCaret(new DocPos(0, p0.Length), false);
            ed.ApplyChar(p => p with { Italic = true });
            ed.InsertText("!");
            T.True(ed.Document.Paragraphs[0].Runs[^1].Props.Italic, "typed text uses pending italic");
        });

        r.Test("M14 editor: Tab moves between cells and adds a row at the end", () =>
        {
            using var ed = new WordEditor();
            ed.CreateControl();
            ed.InsertTable(1, 2);
            var t = ed.Document.Blocks.OfType<Table>().First();
            ed.InsertText("A");
            var tab = typeof(WordEditor).GetMethod("HandleTab", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;
            tab.Invoke(ed, [false]);
            ed.InsertText("B");
            T.Eq("B", t.Rows[0].Cells[1].Paragraphs[0].Text, "second cell");
            tab.Invoke(ed, [false]);
            T.Eq(2, t.Rows.Count, "row added by Tab in last cell");
            ed.Undo();
            T.Eq(1, ed.Document.Blocks.OfType<Table>().First().Rows.Count, "undo row");
        });
    }
}
