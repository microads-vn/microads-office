using System.Drawing.Printing;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.IO;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.UI;

public sealed partial class WordMainForm
{
    private Panel _infoPage = null!;
    private Label _infoStats = null!, _infoPath = null!;
    private TextBox _infoTitle = null!, _infoAuthor = null!, _infoSubject = null!, _infoKeywords = null!;
    private ComboBox _printer = null!;
    private NumericUpDown _copies = null!;
    private TextBox _printPages = null!;
    private PrintPreviewPanel _preview = null!;

    private void BuildBackstage()
    {
        _backstage.BackRequested += (_, _) => HideBackstage();
        var templates = Templates.All;
        void FromTemplate(DocTemplate t) => NewDocument(Templates.Create(t.Id), newWindow: !IsPristine);
        var home = new StartScreen(AppKind.Word, templates, FromTemplate, p => OpenFiles([p], false), CmdOpen) { Name = "startHome" };
        _backstage.AddPage("home", Loc.T("Trang chủ", "Home"), WI.Home, home);
        var newPage = new StartScreen(AppKind.Word, templates, FromTemplate, p => OpenFiles([p], false), CmdOpen, compactTemplates: false) { Name = "startNew" };
        _backstage.AddPage("new", Loc.T("Mới", "New"), WI.NewDoc, newPage);
        var openPage = new StartScreen(AppKind.Word, [], FromTemplate, p => OpenFiles([p], false), CmdOpen) { Name = "startOpen" };
        _backstage.AddPage("open", Loc.T("Mở", "Open"), WI.Open, openPage);
        _backstage.AddSeparator();
        _backstage.AddPage("info", Loc.T("Thông tin", "Info"), WI.Info, BuildInfoPage());
        _backstage.AddCommand("Save", Loc.T("Lưu", "Save"), WI.Save, () => { if (CmdSave()) HideBackstage(); });
        _backstage.AddCommand("SaveAs", Loc.T("Lưu thành", "Save As"), WI.Save, () => { if (CmdSaveAs()) HideBackstage(); });
        _backstage.AddPage("print", Loc.T("In", "Print"), WI.Print, BuildPrintPage());
        _backstage.AddPage("export", Loc.T("Xuất", "Export"), WI.Export, BuildExportPage());
        _backstage.AddCommand("Close", Loc.T("Đóng", "Close"), WI.Close, () => Close());
        _backstage.AddCommand("About", Loc.T("Giới thiệu", "About"), WI.Info, () => OfficeShell.ShowAbout(this), bottom: true);
        _backstage.AddCommand("Options", Loc.T("Cài đặt", "Options"), WI.Settings, () => OfficeShell.ShowSettings(_primary ?? this), bottom: true);
        _backstage.PageShown += (_, page) =>
        {
            if (page == "info") RefreshInfo();
            if (page == "print") RefreshPrint();
            if (page is "home" or "new" or "open") foreach (var s in new[] { home, newPage, openPage }) s.RefreshRecent();
        };
    }

    // ------------------------------------------------------------------ Info

    private Control BuildInfoPage()
    {
        _infoPage = new Panel { Name = "pageInfo", AutoScroll = true, BackColor = Color.White };
        var title = new Label { Text = Loc.T("Thông tin", "Info"), Font = new Font("Segoe UI Light", 22f), AutoSize = true, Location = new Point(0, 0) };
        _infoPath = new Label { Name = "lblInfoPath", AutoSize = true, ForeColor = Theme.SubtleInk, Location = new Point(Dpi.Scale(2), Dpi.Scale(48)) };
        var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Location = new Point(0, Dpi.Scale(84)), Name = "gridInfo" };
        TextBox Tb(string name) { var t = new TextBox { Name = name, Width = Dpi.Scale(360) }; t.Leave += (_, _) => ApplyInfo(); return t; }
        _infoTitle = Tb("txtInfoTitle"); _infoAuthor = Tb("txtInfoAuthor"); _infoSubject = Tb("txtInfoSubject"); _infoKeywords = Tb("txtInfoKeywords");
        void Row(string text, Control c) { grid.Controls.Add(new Label { Text = text, AutoSize = true, Margin = new Padding(0, Dpi.Scale(8), Dpi.Scale(16), 0) }); grid.Controls.Add(c); }
        Row(Loc.T("Tiêu đề", "Title"), _infoTitle);
        Row(Loc.T("Tác giả", "Author"), _infoAuthor);
        Row(Loc.T("Chủ đề", "Subject"), _infoSubject);
        Row(Loc.T("Từ khóa", "Keywords"), _infoKeywords);
        _infoStats = new Label { Name = "lblInfoStats", AutoSize = true, Location = new Point(0, Dpi.Scale(250)), Font = new Font("Segoe UI", 10f), ForeColor = Theme.Ink };
        var apply = new Button { Name = "btnInfoApply", Text = Loc.T("Áp dụng thuộc tính", "Apply properties"), AutoSize = true, Location = new Point(0, Dpi.Scale(214)), FlatStyle = FlatStyle.Flat };
        apply.FlatAppearance.BorderColor = Theme.Border;
        apply.Click += (_, _) => ApplyInfo();
        var openFolder = new LinkLabel { Name = "lnkInfoFolder", Text = Loc.T("Mở thư mục chứa tệp", "Open file location"), AutoSize = true, Location = new Point(0, Dpi.Scale(370)), LinkColor = Theme.AccentDark };
        openFolder.LinkClicked += (_, _) => { if (_path != null) try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + _path + "\""); } catch { } };
        _infoPage.Controls.AddRange([title, _infoPath, grid, apply, _infoStats, openFolder]);
        return _infoPage;
    }

    private void RefreshInfo()
    {
        var d = _ed.Document;
        _infoPath.Text = (_path ?? Loc.T("(chưa lưu)", "(not saved yet)")) + (_ed.Modified ? Loc.T("  — có thay đổi chưa lưu", "  — unsaved changes") : "");
        _infoTitle.Text = d.Title; _infoAuthor.Text = d.Author; _infoSubject.Text = d.Subject; _infoKeywords.Text = d.Keywords;
        var (w, c, cn, p) = Editing.Count(d);
        _infoStats.Text = Loc.F("Số trang: {0}\nSố từ: {1:N0}\nKý tự: {2:N0} ({3:N0} không tính khoảng trắng)\nĐoạn văn: {4:N0}\nKhổ giấy: {5} {6}", "Pages: {0}\nWords: {1:N0}\nCharacters: {2:N0} ({3:N0} without spaces)\nParagraphs: {4:N0}\nPaper: {5} {6}",
            _ed.PageCount, w, c, cn, p, d.Page.PaperName, d.Page.Landscape ? Loc.T("ngang", "landscape") : Loc.T("dọc", "portrait"));
    }

    private void ApplyInfo()
    {
        var d = _ed.Document;
        if (d.Title == _infoTitle.Text && d.Author == _infoAuthor.Text && d.Subject == _infoSubject.Text && d.Keywords == _infoKeywords.Text) return;
        string t = _infoTitle.Text, a = _infoAuthor.Text, s = _infoSubject.Text, k = _infoKeywords.Text;
        _ed.DocumentEdit(doc => { doc.Title = t; doc.Author = a; doc.Subject = s; doc.Keywords = k; });
    }

    // ------------------------------------------------------------------ Print

    private Control BuildPrintPage()
    {
        var page = new Panel { Name = "pagePrint", BackColor = Color.White };
        var left = new Panel { Dock = DockStyle.Left, Width = Dpi.Scale(300), Name = "printOptions" };
        var title = new Label { Text = Loc.T("In", "Print"), Font = new Font("Segoe UI Light", 22f), AutoSize = true, Location = new Point(0, 0) };
        var print = new Button
        {
            Name = "btnDoPrint", Text = "  " + Loc.T("In", "Print"), Image = Icons.Get(WI.Print, Dpi.Scale(22), Color.White), TextImageRelation = TextImageRelation.ImageBeforeText,
            Size = new Size(Dpi.Scale(120), Dpi.Scale(56)), Location = new Point(0, Dpi.Scale(54)), FlatStyle = FlatStyle.Flat, BackColor = Theme.Accent, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 11f)
        };
        print.FlatAppearance.BorderSize = 0;
        print.Click += (_, _) => DoPrint();
        _copies = new NumericUpDown { Name = "numCopies", Minimum = 1, Maximum = 999, Value = 1, Width = Dpi.Scale(70), Location = new Point(Dpi.Scale(200), Dpi.Scale(70)) };
        var copiesLbl = new Label { Text = Loc.T("Số bản:", "Copies:"), AutoSize = true, Location = new Point(Dpi.Scale(136), Dpi.Scale(73)) };
        var prLbl = new Label { Text = Loc.T("Máy in", "Printer"), AutoSize = true, Font = Theme.UiFontBold, Location = new Point(0, Dpi.Scale(128)) };
        _printer = new ComboBox { Name = "cboPrinter", DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(280), Location = new Point(0, Dpi.Scale(150)) };
        var props = new LinkLabel { Name = "lnkPrinterProps", Text = Loc.T("Thuộc tính máy in", "Printer properties"), AutoSize = true, Location = new Point(0, Dpi.Scale(180)), LinkColor = Theme.AccentDark };
        props.LinkClicked += (_, _) =>
        {
            using var pd = new PrintDialog { UseEXDialog = true, AllowSomePages = false };
            pd.PrinterSettings.PrinterName = _printer.Text;
            if (pd.ShowDialog(this) == DialogResult.OK && pd.PrinterSettings.IsValid) _printer.Text = pd.PrinterSettings.PrinterName;
        };
        var rangeLbl = new Label { Text = Loc.T("Trang cần in", "Pages"), AutoSize = true, Font = Theme.UiFontBold, Location = new Point(0, Dpi.Scale(216)) };
        var all = new RadioButton { Name = "rbPrintAll", Text = Loc.T("Tất cả các trang", "Print all pages"), AutoSize = true, Checked = true, Location = new Point(0, Dpi.Scale(238)) };
        var cur = new RadioButton { Name = "rbPrintCurrent", Text = Loc.T("Trang hiện tại", "Current page"), AutoSize = true, Location = new Point(0, Dpi.Scale(262)) };
        var some = new RadioButton { Name = "rbPrintSome", Text = Loc.T("Các trang:", "Pages:"), AutoSize = true, Location = new Point(0, Dpi.Scale(286)) };
        _printPages = new TextBox { Name = "txtPrintPages", Width = Dpi.Scale(150), Location = new Point(Dpi.Scale(100), Dpi.Scale(285)), PlaceholderText = "1-3, 5" };
        _printPages.TextChanged += (_, _) => some.Checked = _printPages.Text.Length > 0;
        var hint = new Label { Text = Loc.T("Khổ giấy và lề theo thiết lập của tài liệu (thẻ Bố cục).", "Paper size and margins follow the document (Layout tab)."), Location = new Point(0, Dpi.Scale(320)), Size = new Size(Dpi.Scale(280), Dpi.Scale(40)), ForeColor = Theme.SubtleInk };
        var setup = new LinkLabel { Name = "lnkPrintPageSetup", Text = Loc.T("Thiết lập trang...", "Page setup..."), AutoSize = true, Location = new Point(0, Dpi.Scale(362)), LinkColor = Theme.AccentDark };
        setup.LinkClicked += (_, _) => { CmdPageSetup(); RefreshPrint(); };
        var pdf = new LinkLabel { Name = "lnkPrintPdf", Text = Loc.T("Xuất ra tệp PDF...", "Export to PDF..."), AutoSize = true, Location = new Point(0, Dpi.Scale(388)), LinkColor = Theme.AccentDark };
        pdf.LinkClicked += (_, _) => CmdExportPdf();
        left.Controls.AddRange([title, print, copiesLbl, _copies, prLbl, _printer, props, rangeLbl, all, cur, some, _printPages, hint, setup, pdf]);
        _cmd["btnDoPrint"] = print;
        _preview = new PrintPreviewPanel { Dock = DockStyle.Fill, Name = "printPreview" };
        page.Controls.Add(_preview);
        page.Controls.Add(left);
        all.CheckedChanged += (_, _) => { if (all.Checked) _printPages.Text = ""; };
        cur.CheckedChanged += (_, _) => { if (cur.Checked) _printPages.Text = ""; };
        page.Tag = (all, cur, some);
        return page;
    }

    private void RefreshPrint()
    {
        if (_printer.Items.Count == 0)
        {
            try
            {
                foreach (string p in PrinterSettings.InstalledPrinters) _printer.Items.Add(p);
                var def = new PrinterSettings().PrinterName;
                _printer.SelectedItem = _printer.Items.Contains(def) ? def : _printer.Items.Count > 0 ? _printer.Items[0] : null;
            }
            catch { }
        }
        _preview.SetLayout(new DocLayout(_ed.Document, _ed.Measurer), _ed.CurrentPage);
    }

    /// <summary>Pages selected on the print page (zero-based), null = all. Empty list = invalid.</summary>
    public List<int>? SelectedPrintPages()
    {
        var page = (Panel)_backstage.Controls.Find("pagePrint", true)[0];
        var (all, cur, some) = ((RadioButton, RadioButton, RadioButton))page.Tag!;
        if (cur.Checked) return [_ed.CurrentPage];
        if (some.Checked) return WordPrintDocument.ParseRange(_printPages.Text, _ed.PageCount) ?? [];
        return null;
    }

    private void DoPrint()
    {
        var pages = SelectedPrintPages();
        if (pages is { Count: 0 }) { Msg.Warn(this, Loc.F("Phạm vi trang không hợp lệ. Tài liệu có {0} trang.", "Invalid page range. The document has {0} pages.", _ed.PageCount)); return; }
        try
        {
            var layout = new DocLayout(_ed.Document, _ed.Measurer);
            using var pd = new WordPrintDocument(layout) { PageList = pages, DocumentName = DisplayName };
            if (_printer.SelectedItem is string name) pd.PrinterSettings.PrinterName = name;
            pd.PrinterSettings.Copies = (short)_copies.Value;
            pd.PrintController = new StandardPrintController();
            pd.Print();
            HideBackstage();
        }
        catch (Exception ex) { Msg.Error(this, Loc.T("Không in được: ", "Printing failed: ") + ex.Message); }
    }

    // ------------------------------------------------------------------ Export

    private Control BuildExportPage()
    {
        var page = new FlowLayoutPanel { Name = "pageExport", FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, BackColor = Color.White };
        page.Controls.Add(new Label { Text = Loc.T("Xuất", "Export"), Font = new Font("Segoe UI Light", 22f), AutoSize = true, Margin = new Padding(0, 0, 0, Dpi.Scale(16)) });
        Button Big(string name, string title, string sub, string glyph, Action a)
        {
            var b = new Button
            {
                Name = name, Size = new Size(Dpi.Scale(460), Dpi.Scale(64)), FlatStyle = FlatStyle.Flat, TextAlign = ContentAlignment.MiddleLeft, Cursor = Cursors.Hand,
                Image = Icons.Get(glyph, Dpi.Scale(32), Theme.AccentDark), ImageAlign = ContentAlignment.MiddleLeft, TextImageRelation = TextImageRelation.ImageBeforeText,
                Text = "  " + title + "\n  " + sub, Font = new Font("Segoe UI", 10f), Margin = new Padding(0, 0, 0, Dpi.Scale(8)), Padding = new Padding(Dpi.Scale(10), 0, 0, 0)
            };
            b.FlatAppearance.BorderColor = Theme.Border;
            b.FlatAppearance.MouseOverBackColor = Theme.Hover;
            b.Click += (_, _) => a();
            _cmd[name] = b;
            return b;
        }
        page.Controls.Add(Big("btnExportPdf", Loc.T("Tạo tài liệu PDF", "Create PDF document"), Loc.T("Giữ nguyên bố cục, phông chữ; chữ trong PDF chọn/tìm được", "Keeps layout and fonts; text stays searchable"), WI.Pdf, CmdExportPdf));
        page.Controls.Add(Big("btnExportHtml", Loc.T("Lưu thành trang web (.html)", "Save as web page (.html)"), Loc.T("Mở được bằng mọi trình duyệt", "Opens in any browser"), WI.Globe, () => SaveCopyAs("html")));
        page.Controls.Add(Big("btnExportRtf", Loc.T("Lưu định dạng RTF (.rtf)", "Save as Rich Text (.rtf)"), Loc.T("Tương thích WordPad và các trình soạn thảo cũ", "Compatible with WordPad and older editors"), WI.Doc, () => SaveCopyAs("rtf")));
        page.Controls.Add(Big("btnExportTxt", Loc.T("Lưu văn bản thuần (.txt)", "Save as plain text (.txt)"), Loc.T("Chỉ giữ chữ, mã hóa UTF-8", "Text only, UTF-8"), WI.Doc, () => SaveCopyAs("txt")));
        return page;
    }

    private void SaveCopyAs(string ext)
    {
        using var d = new SaveFileDialog
        {
            Filter = ext switch { "html" => Loc.T("Trang web", "Web page") + " (*.html)|*.html", "rtf" => "RTF (*.rtf)|*.rtf", _ => Loc.T("Văn bản", "Text") + " (*.txt)|*.txt" },
            FileName = Path.GetFileNameWithoutExtension(_path ?? SuggestName()) + "." + ext, AddExtension = true
        };
        if (d.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            WordFiles.Save(_ed.Document, d.FileName);
            Msg.Info(this, Loc.T("Đã lưu: ", "Saved: ") + d.FileName);
        }
        catch (Exception ex) { Msg.Error(this, ex.Message); }
    }
}

/// <summary>Print preview inside the backstage: renders laid-out pages with page navigation.</summary>
public sealed class PrintPreviewPanel : Panel
{
    private DocLayout? _layout;
    private int _page;
    private readonly Label _lbl;
    private readonly Panel _canvas;

    public PrintPreviewPanel()
    {
        BackColor = Color.FromArgb(236, 237, 240);
        _canvas = new Panel { Dock = DockStyle.Fill, BackColor = BackColor };
        typeof(Panel).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!.SetValue(_canvas, true);
        _canvas.Paint += PaintPage;
        _canvas.Resize += (_, _) => _canvas.Invalidate();
        var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = Dpi.Scale(36), FlowDirection = FlowDirection.LeftToRight, Padding = new Padding(Dpi.Scale(12), Dpi.Scale(4), 0, 0) };
        var prev = new Button { Name = "btnPreviewPrev", Text = "◀", Width = Dpi.Scale(36), FlatStyle = FlatStyle.Flat };
        var next = new Button { Name = "btnPreviewNext", Text = "▶", Width = Dpi.Scale(36), FlatStyle = FlatStyle.Flat };
        _lbl = new Label { Name = "lblPreviewPage", AutoSize = true, Padding = new Padding(Dpi.Scale(8), Dpi.Scale(6), Dpi.Scale(8), 0) };
        prev.Click += (_, _) => Show(_page - 1);
        next.Click += (_, _) => Show(_page + 1);
        bar.Controls.AddRange([prev, _lbl, next]);
        Controls.Add(_canvas);
        Controls.Add(bar);
    }

    public int PageIndex => _page;

    public void SetLayout(DocLayout layout, int page) { _layout = layout; Show(page); }

    public void Show(int page)
    {
        if (_layout == null) return;
        _page = Math.Clamp(page, 0, _layout.Pages.Count - 1);
        _lbl.Text = Loc.F("Trang {0} / {1}", "Page {0} of {1}", _page + 1, _layout.Pages.Count);
        _canvas.Invalidate();
    }

    private void PaintPage(object? sender, PaintEventArgs e)
    {
        if (_layout == null) return;
        var ps = _layout.Doc.Page;
        float pad = Dpi.Scale(20);
        float s = Math.Min((_canvas.Width - 2 * pad) / ps.Width, (_canvas.Height - 2 * pad) / ps.Height);
        if (s <= 0) return;
        float w = ps.Width * s, h = ps.Height * s;
        float x = (_canvas.Width - w) / 2, y = (_canvas.Height - h) / 2;
        var g = e.Graphics;
        using (var sh = new SolidBrush(Color.FromArgb(50, 0, 0, 0))) g.FillRectangle(sh, x + 3, y + 4, w, h);
        var st = g.Save();
        PageRenderer.Setup(g, x, y, s);
        g.SetClip(new RectangleF(0, 0, ps.Width, ps.Height));
        PageRenderer.DrawPage(g, _layout, _page, new RenderOptions { Print = true });
        g.Restore(st);
    }
}
