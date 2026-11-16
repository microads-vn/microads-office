using System.Drawing.Printing;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>
/// Prints worksheets: page breaks by rows/columns, print area, rows repeated at top, fit all columns
/// on one page, scale, gridlines/headings, centering, charts and pictures, page numbers in the footer.
/// </summary>
public sealed class SheetPrintDocument : PrintDocument
{
    private readonly List<Sheet> _sheets;
    private readonly List<(Sheet Sheet, List<int> Rows, List<int> Cols, float Scale)> _pages = new();
    private int _page;
    private readonly CellRange? _selection;

    public int PageCount => _pages.Count;

    /// <param name="selection">Print only this range of the first sheet (Print Selection).</param>
    public SheetPrintDocument(IEnumerable<Sheet> sheets, CellRange? selection = null)
    {
        _sheets = sheets.ToList();
        _selection = selection;
        DocumentName = "Microads Office Excel";
        var first = _sheets.FirstOrDefault();
        if (first != null)
        {
            var p = first.Page;
            DefaultPageSettings.Landscape = p.Landscape;
            var ps = PaperFor(p.PaperSize);
            DefaultPageSettings.PaperSize = ps;
            DefaultPageSettings.Margins = new Margins((int)(p.MarginLeft * 100), (int)(p.MarginRight * 100), (int)(p.MarginTop * 100), (int)(p.MarginBottom * 100));
        }
        OriginAtMargins = false;
    }

    private static PaperSize PaperFor(int id) => id switch
    {
        1 => new PaperSize("Letter", 850, 1100),
        5 => new PaperSize("Legal", 850, 1400),
        8 => new PaperSize("A3", 1169, 1654),
        11 => new PaperSize("A5", 583, 827),
        _ => new PaperSize("A4", 827, 1169)
    };

    protected override void OnBeginPrint(PrintEventArgs e)
    {
        base.OnBeginPrint(e);
        _page = 0;
        Paginate();
    }

    /// <summary>Splits all sheets into pages (also used to count pages for the preview).</summary>
    public void Paginate()
    {
        _pages.Clear();
        foreach (var s in _sheets)
        {
            var area = _selection ?? s.Page.PrintArea ?? s.UsedRange();
            if (_selection == null && s.Page.PrintArea == null)
            {
                // include charts and pictures
                foreach (var b in s.Charts.Select(c => c.Bounds).Concat(s.Pictures.Select(p => p.Bounds)))
                {
                    var (c2, _) = IO.XlsxWriter.PxToCol(s, b.Right);
                    var (r2, _) = IO.XlsxWriter.PxToRow(s, b.Bottom);
                    area = area.Union(new CellRange(0, 0, r2, c2));
                }
                if (s.CellCount == 0 && s.Charts.Count == 0 && s.Pictures.Count == 0) continue;
            }
            var p = s.Page;
            var pageSz = DefaultPageSettings.PaperSize;
            float pw = (p.Landscape ? pageSz.Height : pageSz.Width) - (float)(p.MarginLeft + p.MarginRight) * 100;
            float ph = (p.Landscape ? pageSz.Width : pageSz.Height) - (float)(p.MarginTop + p.MarginBottom) * 100 - 30; // footer
            // sheet pixels (96 dpi) → 1/100 inch
            float px = 100f / 96f;
            float headW = p.PrintHeadings ? 30 * px : 0, headH = p.PrintHeadings ? 20 * px : 0;
            var cols = Enumerable.Range(area.C1, area.Cols).Where(c => s.ColWidth(c) > 0).ToList();
            var rows = Enumerable.Range(area.R1, Math.Min(area.Rows, 200000)).Where(r => s.RowHeight(r) > 0).ToList();
            var titles = p.PrintTitleRows is (int t1, int t2) ? Enumerable.Range(t1, t2 - t1 + 1).Where(r => s.RowHeight(r) > 0).ToList() : new List<int>();
            float scale = p.Scale / 100f;
            if (p.FitToWidth)
            {
                float totalW = cols.Sum(c => (float)s.ColWidth(c)) * px + headW;
                scale = Math.Min(1f, pw / Math.Max(1, totalW));
            }
            float titleH = titles.Sum(r => (float)s.RowHeight(r)) * px * scale;
            // column bands
            var colBands = new List<List<int>>();
            var band = new List<int>();
            float x = headW * scale;
            foreach (var c in cols)
            {
                float w = (float)s.ColWidth(c) * px * scale;
                if (band.Count > 0 && x + w > pw) { colBands.Add(band); band = new(); x = headW * scale; }
                band.Add(c); x += w;
            }
            if (band.Count > 0) colBands.Add(band);
            // row bands
            var rowBands = new List<List<int>>();
            var rb = new List<int>();
            float y = headH * scale;
            foreach (var r in rows)
            {
                float h = (float)s.RowHeight(r) * px * scale;
                float avail = ph - (rowBands.Count > 0 || rb.Count > 0 ? (titles.Contains(r) ? 0 : titleH) : 0);
                if (rb.Count > 0 && y + h > avail) { rowBands.Add(rb); rb = new(); y = headH * scale; }
                rb.Add(r); y += h;
            }
            if (rb.Count > 0) rowBands.Add(rb);
            foreach (var cb in colBands)
                for (int i = 0; i < rowBands.Count; i++)
                {
                    var rowsOnPage = rowBands[i];
                    if (i > 0 && titles.Count > 0) rowsOnPage = titles.Concat(rowsOnPage.Where(r => !titles.Contains(r))).ToList();
                    _pages.Add((s, rowsOnPage, cb, scale));
                }
        }
        if (_pages.Count == 0 && _sheets.Count > 0) _pages.Add((_sheets[0], new List<int> { 0 }, new List<int> { 0 }, 1f));
    }

    protected override void OnPrintPage(PrintPageEventArgs e)
    {
        base.OnPrintPage(e);
        if (_page >= _pages.Count) { e.HasMorePages = false; return; }
        var (s, rows, cols, scale) = _pages[_page];
        var g = e.Graphics!;
        g.PageUnit = GraphicsUnit.Display; // 1/100 inch
        var p = s.Page;
        float px = 100f / 96f * scale;
        float left = e.MarginBounds.Left, top = e.MarginBounds.Top;
        bool heads = p.PrintHeadings;
        float headW = heads ? 30 * px : 0, headH = heads ? 20 * px : 0;
        float totalW = cols.Sum(c => (float)s.ColWidth(c)) * px + headW;
        if (p.CenterHorizontally) left += Math.Max(0, (e.MarginBounds.Width - totalW) / 2);
        var rl = new List<CellPainter.Line>();
        float y = top + headH;
        foreach (var r in rows) { float h = (float)s.RowHeight(r) * px; rl.Add(new CellPainter.Line(r, y, h)); y += h; }
        var cl = new List<CellPainter.Line>();
        float x = left + headW;
        foreach (var c in cols) { float w = (float)s.ColWidth(c) * px; cl.Add(new CellPainter.Line(c, x, w)); x += w; }
        var clip = new RectangleF(left, top, x - left, y - top);
        var st = g.Save();
        g.SetClip(clip);
        CellPainter.Paint(g, s, rl, cl, px, scale * 0.98f, p.PrintGridlines, false, false, new CfEvaluator(s), clip);
        g.Restore(st);
        if (heads)
        {
            using var hf = new Font("Segoe UI", 7f * scale);
            using var pen = new Pen(Color.Silver, 0.5f);
            using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            foreach (var c in cl) { g.DrawRectangle(pen, c.Pos, top, c.Size, headH); g.DrawString(Addr.ColName(c.Index), hf, Brushes.DimGray, new RectangleF(c.Pos, top, c.Size, headH), sf); }
            foreach (var r in rl) { g.DrawRectangle(pen, left, r.Pos, headW, r.Size); g.DrawString((r.Index + 1).ToString(), hf, Brushes.DimGray, new RectangleF(left, r.Pos, headW, r.Size), sf); }
        }
        // objects inside this page
        if (cl.Count > 0 && rl.Count > 0)
        {
            double sx0 = 0; for (int c = 0; c < cl[0].Index; c++) sx0 += s.ColWidth(c);
            double sy0 = 0; for (int r = 0; r < rl.Count && r < 1; r++) { } for (int r = 0; r < rl[0].Index; r++) sy0 += s.RowHeight(r);
            var st2 = g.Save();
            g.SetClip(clip);
            foreach (var pic in s.Pictures)
            {
                var b = new RectangleF(cl[0].Pos + (float)(pic.Bounds.X - sx0) * px, rl[0].Pos + (float)(pic.Bounds.Y - sy0) * px, pic.Bounds.Width * px, pic.Bounds.Height * px);
                if (!b.IntersectsWith(clip)) continue;
                try { using var img = Image.FromStream(new MemoryStream(pic.Data)); g.DrawImage(img, b); } catch { }
            }
            foreach (var ch in s.Charts)
            {
                var b = new RectangleF(cl[0].Pos + (float)(ch.Bounds.X - sx0) * px, rl[0].Pos + (float)(ch.Bounds.Y - sy0) * px, ch.Bounds.Width * px, ch.Bounds.Height * px);
                if (b.IntersectsWith(clip)) ChartRenderer.Draw(g, b, s, ch, px);
            }
            g.Restore(st2);
        }
        using (var ff = new Font("Segoe UI", 8f))
        {
            var footer = Loc.F("Trang {0} / {1}", "Page {0} of {1}", _page + 1, _pages.Count);
            var sz = g.MeasureString(footer, ff);
            g.DrawString(footer, ff, Brushes.Gray, e.MarginBounds.Left + (e.MarginBounds.Width - sz.Width) / 2, e.MarginBounds.Bottom - sz.Height + 20);
        }
        _page++;
        e.HasMorePages = _page < _pages.Count;
    }
}

/// <summary>Backstage "Print" page: printer, copies, what to print, page setup, live preview.</summary>
public sealed class PrintPage : UserControl
{
    private readonly ComboBox _printer = new() { Name = "cboPrinter", DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly NumericUpDown _copies = new() { Name = "numCopies", Minimum = 1, Maximum = 999, Value = 1, Width = 70 };
    private readonly ComboBox _what = new() { Name = "cboPrintWhat", DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly PrintPreviewControl _preview = new() { Name = "printPreview", Dock = DockStyle.Fill, UseAntiAlias = true, BackColor = Color.FromArgb(230, 231, 233), AutoZoom = true };
    private readonly Label _pages = new() { Name = "lblPrintPages", AutoSize = true };
    private readonly Func<ExcelForm> _form;

    public PrintPage(Func<ExcelForm> form)
    {
        _form = form;
        Name = "printPage";
        BackColor = Color.White;
        var left = new FlowLayoutPanel { Dock = DockStyle.Left, Width = Dpi.Scale(300), FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(0, 0, Dpi.Scale(10), 0) };
        left.Controls.Add(new Label { Text = Loc.T("In", "Print"), Font = new Font("Segoe UI Light", 20f), AutoSize = true });
        var print = new Button
        {
            Name = "btnPrintNow", Text = "  " + Loc.T("In", "Print"), Image = Icons.Get(Icons.Print, Dpi.Scale(28), Color.White), TextImageRelation = TextImageRelation.ImageAboveText,
            Size = new Size(Dpi.Scale(96), Dpi.Scale(80)), BackColor = Theme.Accent, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = Theme.UiFontBold
        };
        print.FlatAppearance.BorderSize = 0;
        print.Click += (_, _) => _form().DoPrint(_printer.SelectedItem as string, (short)_copies.Value, _what.SelectedIndex);
        var copiesRow = new FlowLayoutPanel { AutoSize = true };
        copiesRow.Controls.Add(print);
        var cp = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(Dpi.Scale(10), Dpi.Scale(20), 0, 0) };
        cp.Controls.Add(new Label { Text = Loc.T("Số bản:", "Copies:"), AutoSize = true });
        cp.Controls.Add(_copies);
        copiesRow.Controls.Add(cp);
        left.Controls.Add(copiesRow);
        left.Controls.Add(new Label { Text = Loc.T("Máy in", "Printer"), Font = Theme.UiFontBold, AutoSize = true, Margin = new Padding(0, Dpi.Scale(12), 0, 2) });
        try { foreach (string p in PrinterSettings.InstalledPrinters) _printer.Items.Add(p); } catch { }
        try { var def = new PrinterSettings().PrinterName; if (_printer.Items.Contains(def)) _printer.SelectedItem = def; else if (_printer.Items.Count > 0) _printer.SelectedIndex = 0; } catch { }
        left.Controls.Add(_printer);
        left.Controls.Add(new Label { Text = Loc.T("Cài đặt", "Settings"), Font = Theme.UiFontBold, AutoSize = true, Margin = new Padding(0, Dpi.Scale(12), 0, 2) });
        _what.Items.AddRange([Loc.T("In trang tính hiện hoạt", "Print Active Sheet"), Loc.T("In toàn bộ sổ làm việc", "Print Entire Workbook"), Loc.T("In vùng chọn", "Print Selection")]);
        _what.SelectedIndex = 0;
        _what.SelectedIndexChanged += (_, _) => RefreshPreview();
        left.Controls.Add(_what);
        var setup = new Button { Name = "btnPrintPageSetup", Text = Loc.T("Thiết lập trang...", "Page Setup..."), AutoSize = true, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, Dpi.Scale(10), 0, 0) };
        setup.Click += (_, _) => { _form().CmdPageSetup(); RefreshPreview(); };
        left.Controls.Add(setup);
        var pdf = new Button { Name = "btnPrintExportPdf", Text = Loc.T("Xuất PDF...", "Export to PDF..."), AutoSize = true, FlatStyle = FlatStyle.Flat };
        pdf.Click += (_, _) => _form().CmdExportPdf();
        left.Controls.Add(pdf);
        left.Controls.Add(_pages);
        var nav = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = Dpi.Scale(36), FlowDirection = FlowDirection.LeftToRight };
        var prev = new Button { Name = "btnPreviewPrev", Text = "◀", Width = Dpi.Scale(36), FlatStyle = FlatStyle.Flat };
        var next = new Button { Name = "btnPreviewNext", Text = "▶", Width = Dpi.Scale(36), FlatStyle = FlatStyle.Flat };
        prev.Click += (_, _) => { if (_preview.StartPage > 0) _preview.StartPage--; };
        next.Click += (_, _) => _preview.StartPage++;
        nav.Controls.AddRange([prev, next]);
        var right = new Panel { Dock = DockStyle.Fill };
        right.Controls.Add(_preview);
        right.Controls.Add(nav);
        Controls.Add(right);
        Controls.Add(left);
    }

    public int WhatIndex { get => _what.SelectedIndex; set => _what.SelectedIndex = value; }

    public void RefreshPreview()
    {
        try
        {
            var doc = _form().BuildPrintDocument(_what.SelectedIndex);
            doc.Paginate();
            _pages.Text = Loc.F("{0} trang", "{0} page(s)", doc.PageCount);
            _preview.Document = doc;
            _preview.InvalidatePreview();
        }
        catch (Exception ex) { _pages.Text = ex.Message; }
    }
}
