using MicroadsOffice.Core.Office;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Dialogs;
using MicroadsOffice.UI;
using MicroadsOffice.Viewer;
using MicroadsOffice.Views;

namespace MicroadsOffice;

public sealed partial class MainForm : Form, IOfficeAppWindow
{
    private readonly AppSettings _settings;
    private readonly Ribbon _ribbon;
    private readonly TabControl _tabs;
    private readonly Panel _startPage;
    private readonly ToolStrip _status;
    private readonly ToolStripLabel _statusText;
    private readonly ToolStripTextBox _pageBox;
    private readonly ToolStripLabel _pageCount;
    private readonly ToolStripComboBox _zoomBox;
    private readonly ToolStripComboBox _ribbonZoom;
    private readonly FlowLayoutPanel _recentPanel;
    private readonly Dictionary<string, ToolStripItem> _cmd = new();
    private readonly ToolStripMenuItem _recentMenu;
    private bool _fullScreen;
    private FormWindowState _prevState;
    private FormBorderStyle _prevBorder;
    private bool _updatingZoom;

    public DocumentView? Current => _tabs.SelectedTab?.Controls.OfType<DocumentView>().FirstOrDefault();

    public MainForm(string[] files)
    {
        _settings = AppSettings.Load();
        Text = "Microads Office PDF";
        Name = "MainForm";
        Font = Theme.UiFont;
        AutoScaleMode = AutoScaleMode.None;
        Dpi.Factor = DeviceDpi / 96f;
        MinimumSize = new Size(Dpi.Scale(760), Dpi.Scale(520));
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1366, 728);
        Size = new Size(Math.Min(Dpi.Scale(1320), (int)(wa.Width * 0.92)), Math.Min(Dpi.Scale(860), (int)(wa.Height * 0.92)));
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Theme.ViewerBack;
        AllowDrop = true;
        KeyPreview = true;
        try { Icon = OfficeShell.AppIcon ?? BrandArt.CreateIcon(AppKind.Pdf); DialogBase.AppIcon = Icon; } catch { }

        // ---------------- ribbon
        _ribbon = new Ribbon { Name = "ribbon" };
        BuildFileMenu(out _recentMenu);
        var home = _ribbon.AddTab("Home", Loc.T("Trang chủ", "Home"));
        Cmd(home.Button("btnOpen", Loc.T("Mở", "Open"), Icons.Open, (_, _) => CmdOpen(), Loc.T("Mở tệp PDF (Ctrl+O)", "Open a PDF file (Ctrl+O)")));
        Cmd(home.Button("btnSave", Loc.T("Lưu", "Save"), Icons.Save, (_, _) => CmdSave(), Loc.T("Lưu thay đổi (Ctrl+S)", "Save changes (Ctrl+S)")));
        Cmd(home.Button("btnSaveAs", Loc.T("Lưu bản sao", "Save a copy"), Icons.SaveAs, (_, _) => CmdSaveAs(), Loc.T("Lưu thành tệp khác", "Save as another file")));
        Cmd(home.Button("btnPrint", Loc.T("In", "Print"), Icons.Print, (_, _) => CmdPrint(), Loc.T("In tài liệu (Ctrl+P)", "Print the document (Ctrl+P)")));
        home.Separator();
        Cmd(home.Button("btnHand", Loc.T("Công cụ\nTay", "Hand\ntool"), Icons.Hand, (_, _) => CmdTool(ViewerTool.Hand), Loc.T("Kéo để cuộn trang", "Drag to scroll the page")));
        Cmd(home.Button("btnSelect", Loc.T("Chọn\nvăn bản", "Select\ntext"), Icons.SelectText, (_, _) => CmdTool(ViewerTool.SelectText), Loc.T("Bôi đen văn bản để sao chép", "Highlight text to copy it")));
        Cmd(home.Button("btnCopy", Loc.T("Sao chép", "Copy"), Icons.Copy, (_, _) => CmdCopy(), Loc.T("Sao chép văn bản đã chọn (Ctrl+C)", "Copy the selected text (Ctrl+C)")));
        home.Separator();
        Cmd(home.Button("btnZoomOut", Loc.T("Thu nhỏ", "Zoom out"), Icons.ZoomOut, (_, _) => Current?.Viewer.ZoomStep(false), Loc.T("Thu nhỏ (Ctrl -)", "Zoom out (Ctrl -)")));
        Cmd(home.Button("btnZoomIn", Loc.T("Phóng to", "Zoom in"), Icons.ZoomIn, (_, _) => Current?.Viewer.ZoomStep(true), Loc.T("Phóng to (Ctrl +)", "Zoom in (Ctrl +)")));
        Cmd(home.Button("btnFitPage", Loc.T("Vừa trang", "Fit page"), Icons.FitPage, (_, _) => Current?.Viewer.SetZoomMode(ZoomMode.FitPage), Loc.T("Vừa cả trang (Ctrl+0)", "Fit the whole page (Ctrl+0)")));
        Cmd(home.Button("btnFitWidth", Loc.T("Vừa\nchiều rộng", "Fit\nwidth"), Icons.FitWidth, (_, _) => Current?.Viewer.SetZoomMode(ZoomMode.FitWidth), Loc.T("Vừa chiều rộng (Ctrl+2)", "Fit width (Ctrl+2)")));
        Cmd(home.Button("btnRotateLeft", Loc.T("Xoay trái", "Rotate left"), Icons.Rotate, (_, _) => Current?.Viewer.Rotate(false), Loc.T("Xoay trái 90°", "Rotate left 90°"), mirror: true));
        Cmd(home.Button("btnRotateRight", Loc.T("Xoay phải", "Rotate right"), Icons.Rotate, (_, _) => Current?.Viewer.Rotate(true), Loc.T("Xoay phải 90°", "Rotate right 90°")));
        home.Separator();
        Cmd(home.Button("btnSearch", Loc.T("Tìm kiếm", "Search"), Icons.Search, (_, _) => CmdSearch(), Loc.T("Tìm văn bản (Ctrl+F)", "Find text (Ctrl+F)")));
        Cmd(home.Button("btnSignHome", Loc.T("Ký số", "Sign"), Icons.Sign, (_, _) => CmdSign(), Loc.T("Đặt chữ ký số (Ctrl+Shift+S)", "Place a digital signature (Ctrl+Shift+S)")));
        Cmd(home.Button("btnStampHome", Loc.T("Chèn dấu\nảnh", "Image\nstamp"), Icons.Picture, (_, _) => CmdStamp(), Loc.T("Chèn hình con dấu / chữ ký ảnh", "Insert a seal / signature image")));

        var view = _ribbon.AddTab("View", Loc.T("Xem", "View"));
        Cmd(view.Button("btnFirstPage", Loc.T("Trang đầu", "First page"), Icons.First, (_, _) => Current?.Viewer.GoToPage(0)));
        Cmd(view.Button("btnPrevPage", Loc.T("Trang trước", "Previous page"), Icons.Prev, (_, _) => Current?.Viewer.PrevPage()));
        Cmd(view.Button("btnNextPage", Loc.T("Trang sau", "Next page"), Icons.Next, (_, _) => Current?.Viewer.NextPage()));
        Cmd(view.Button("btnLastPage", Loc.T("Trang cuối", "Last page"), Icons.Last, (_, _) => Current?.Viewer.GoToPage(int.MaxValue)));
        Cmd(view.Button("btnGoToPage", Loc.T("Đi tới\ntrang...", "Go to\npage..."), Icons.GoTo, (_, _) => CmdGoTo(), Loc.T("Đi tới trang (Ctrl+G)", "Go to page (Ctrl+G)")));
        view.Separator();
        Cmd(view.Button("btnActualSize", Loc.T("Kích thước\nthật", "Actual\nsize"), "1:1", (_, _) => Current?.Viewer.SetZoom(1f), Loc.T("Kích thước thật 100% (Ctrl+1)", "Actual size 100% (Ctrl+1)")));
        _ribbonZoom = new ToolStripComboBox { Name = "cboRibbonZoom", AutoSize = false, Width = Dpi.Scale(110), AccessibleName = Loc.T("Tỉ lệ thu phóng", "Zoom level"), Margin = new Padding(Dpi.Scale(4), Dpi.Scale(22), Dpi.Scale(4), 0) };
        FillZoomItems(_ribbonZoom);
        view.Add(_ribbonZoom);
        Cmd(_ribbonZoom);
        view.Separator();
        Cmd(view.Button("btnSinglePage", Loc.T("Một trang", "Single page"), Icons.SinglePage, (_, _) => Current?.Viewer.SetLayoutMode(PageLayoutMode.SinglePage), Loc.T("Hiển thị từng trang", "Show one page at a time")));
        Cmd(view.Button("btnContinuous", Loc.T("Cuộn\nliên tục", "Continuous\nscroll"), Icons.Continuous, (_, _) => Current?.Viewer.SetLayoutMode(PageLayoutMode.Continuous), Loc.T("Cuộn liên tục các trang", "Scroll pages continuously")));
        view.Separator();
        Cmd(view.Button("btnNavPane", Loc.T("Ngăn\nđiều hướng", "Navigation\npane"), Icons.NavPane, (_, _) => CmdToggleNav(), Loc.T("Ẩn/hiện ngăn điều hướng (F4)", "Show/hide the navigation pane (F4)")));
        Cmd(view.Button("btnThumbnails", Loc.T("Hình\nthu nhỏ", "Thumb-\nnails"), Icons.Thumbnails, (_, _) => Current?.ShowNav(NavKind.Thumbnails)));
        Cmd(view.Button("btnBookmarks", Loc.T("Dấu trang", "Bookmarks"), Icons.Bookmarks, (_, _) => Current?.ShowNav(NavKind.Bookmarks)));
        Cmd(view.Button("btnFullScreen", Loc.T("Toàn\nmàn hình", "Full\nscreen"), Icons.FullScreen, (_, _) => CmdFullScreen(), Loc.T("Chế độ đọc toàn màn hình (F11)", "Full-screen reading mode (F11)")));
        Cmd(view.Button("btnDocProperties", Loc.T("Thuộc tính\ntài liệu", "Document\nproperties"), Icons.Properties, (_, _) => CmdProperties(), Loc.T("Thông tin tài liệu (Ctrl+D)", "Document information (Ctrl+D)")));

        var sign = _ribbon.AddTab("Protect", Loc.T("Ký && Bảo vệ", "Sign && Protect"));
        Cmd(sign.Button("btnSign", Loc.T("Ký số", "Sign"), Icons.Sign, (_, _) => CmdSign(), Loc.T("Đặt chữ ký số lên tài liệu (Ctrl+Shift+S)", "Place a digital signature on the document (Ctrl+Shift+S)")));
        Cmd(sign.Button("btnValidateAll", Loc.T("Xác thực\nchữ ký", "Validate\nsignatures"), Icons.Shield, (_, _) => CmdValidate(), Loc.T("Xác thực lại tất cả chữ ký", "Re-validate all signatures")));
        Cmd(sign.Button("btnSignaturePanel", Loc.T("Bảng\nchữ ký", "Signature\npanel"), Icons.Signatures, (_, _) => Current?.ShowNav(NavKind.Signatures), Loc.T("Mở bảng chữ ký số", "Open the signature panel")));
        sign.Separator();
        Cmd(sign.Button("btnStyles", Loc.T("Mẫu\nchữ ký", "Signature\nstyles"), Icons.Style, (_, _) => CmdStyles(), Loc.T("Tạo / quản lý mẫu hiển thị chữ ký (con dấu, chữ ký tay)", "Create / manage signature appearances (seal, handwritten signature)")));
        Cmd(sign.Button("btnCreateId", Loc.T("Tạo\nID số", "Create\ndigital ID"), Icons.Key, (_, _) => CmdCreateId(), Loc.T("Tạo chứng thư tự ký (.pfx)", "Create a self-signed certificate (.pfx)")));
        Cmd(sign.Button("btnCertificates", Loc.T("Chứng thư\ncủa tôi", "My\ncertificates"), Icons.Certificate, (_, _) => CmdCertificates(), Loc.T("Xem các chứng thư số / USB Token khả dụng", "View available certificates / USB tokens")));
        Cmd(sign.Button("btnStamp", Loc.T("Chèn\ndấu ảnh", "Image\nstamp"), Icons.Picture, (_, _) => CmdStamp(), Loc.T("Chèn hình con dấu / chữ ký ảnh (không phải chữ ký số)", "Insert a seal / signature image (not a digital signature)")));
        sign.Separator();
        Cmd(sign.Button("btnSignSettings", Loc.T("Thiết lập\nký số", "Signing\nsettings"), Icons.Settings, (_, _) => CmdSignSettings(), Loc.T("Chuẩn ký, dấu thời gian, kiểm tra thu hồi", "Signature standard, timestamp, revocation checking")));

        var help = _ribbon.AddTab("Help", Loc.T("Trợ giúp", "Help"));
        Cmd(help.Button("btnHelp", Loc.T("Hướng dẫn", "User guide"), Icons.Help, (_, _) => { using var d = new HelpDialog(); d.ShowDialog(this); }, Loc.T("Hướng dẫn sử dụng & phím tắt (F1)", "User guide & keyboard shortcuts (F1)")));
        Cmd(help.Button("btnCheckUpdate", Loc.T("Kiểm tra\ncập nhật", "Check for\nupdates"), Icons.Sync, (_, _) => OfficeShell.CheckForUpdatesInteractive(this), Loc.T("Kiểm tra phiên bản mới tại microads.vn/office", "Check microads.vn/office for a new version")));
        Cmd(help.Button("btnSettings", Loc.T("Cài đặt &&\nngôn ngữ", "Settings &&\nlanguage"), Icons.Settings, (_, _) => OfficeShell.ShowSettings(this), Loc.T("Ngôn ngữ (Tiếng Việt / English), tự động cập nhật", "Language (Tiếng Việt / English), automatic updates")));
        help.Separator();
        Cmd(help.Button("btnFileAssoc", Loc.T("Mở PDF bằng\nMicroads", "Open PDFs with\nMicroads"), Icons.Document, (_, _) => CmdFileAssociation(), Loc.T("Thêm Microads Office vào danh sách 'Mở bằng' cho tệp PDF", "Add Microads Office PDF to the 'Open with' list for PDF files")));
        Cmd(help.Button("btnAbout", Loc.T("Giới thiệu", "About"), Icons.Info, (_, _) => OfficeShell.ShowAbout(this), Loc.T("Giới thiệu Microads Office PDF", "About Microads Office PDF")));
        Cmd(help.Button("btnVisitWebsite", "microads.vn", Icons.Globe, (_, _) => OfficeShell.OpenWebsite(), Loc.T("Truy cập microads.vn — giải pháp phần mềm doanh nghiệp, bảo mật, CRM, ERP, website", "Visit microads.vn — enterprise software, security, CRM, ERP, websites")));
        _ribbon.FileButton.Text = Loc.T("Tệp", "File");
        _ribbon.FileButton.AccessibleName = Loc.T("Tệp", "File");
        foreach (var item in OfficeShell.CreateHeaderItems(this)) _ribbon.AddHeaderRight(item);
        _ribbon.Select(home);

        // ---------------- document tabs
        _tabs = new DocTabControl { Dock = DockStyle.Fill, Name = "tabDocuments" };
        _tabs.SelectedIndexChanged += (_, _) => OnCurrentChanged();
        ((DocTabControl)_tabs).CloseRequested += page => CloseTab(page);

        // ---------------- start page
        _startPage = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Name = "startPage", AutoScroll = true };
        _recentPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Name = "recentList" };
        BuildStartPage();

        // ---------------- status bar
        _status = new ToolStrip
        {
            Dock = DockStyle.Bottom, GripStyle = ToolStripGripStyle.Hidden, Renderer = new RibbonRenderer(false), BackColor = Theme.TabStripBack,
            Name = "statusBar", Padding = new Padding(Dpi.Scale(6), Dpi.Scale(2), Dpi.Scale(6), Dpi.Scale(2)), ImageScalingSize = new Size(Dpi.Scale(16), Dpi.Scale(16))
        };
        _statusText = new ToolStripLabel(Loc.T("Sẵn sàng", "Ready")) { Name = "lblStatus", ForeColor = Theme.SubtleInk };
        _status.Items.Add(_statusText);
        ToolStripButton SB(string name, string glyph, string tip, Action a, bool mirror = false)
        {
            var b = new ToolStripButton(Icons.Small(glyph, null, mirror)) { Name = name, ToolTipText = tip, AccessibleName = tip, Alignment = ToolStripItemAlignment.Right, DisplayStyle = ToolStripItemDisplayStyle.Image };
            b.Click += (_, _) => a();
            return b;
        }
        // right-aligned items are laid out right-to-left
        var sbFitW = SB("sbFitWidth", Icons.FitWidth, Loc.T("Vừa chiều rộng", "Fit width"), () => Current?.Viewer.SetZoomMode(ZoomMode.FitWidth));
        var sbFitP = SB("sbFitPage", Icons.FitPage, Loc.T("Vừa trang", "Fit page"), () => Current?.Viewer.SetZoomMode(ZoomMode.FitPage));
        var sbZin = SB("sbZoomIn", Icons.ZoomIn, Loc.T("Phóng to", "Zoom in"), () => Current?.Viewer.ZoomStep(true));
        _zoomBox = new ToolStripComboBox { Name = "cboZoom", AutoSize = false, Width = Dpi.Scale(96), Alignment = ToolStripItemAlignment.Right, AccessibleName = Loc.T("Tỉ lệ thu phóng (thanh trạng thái)", "Zoom level (status bar)") };
        FillZoomItems(_zoomBox);
        var sbZout = SB("sbZoomOut", Icons.ZoomOut, Loc.T("Thu nhỏ", "Zoom out"), () => Current?.Viewer.ZoomStep(false));
        var sep = new ToolStripSeparator { Alignment = ToolStripItemAlignment.Right };
        var sbLast = SB("sbLastPage", Icons.Last, Loc.T("Trang cuối", "Last page"), () => Current?.Viewer.GoToPage(int.MaxValue));
        var sbNext = SB("sbNextPage", Icons.Next, Loc.T("Trang sau", "Next page"), () => Current?.Viewer.NextPage());
        _pageCount = new ToolStripLabel("/ 0") { Name = "lblPageCount", Alignment = ToolStripItemAlignment.Right };
        _pageBox = new ToolStripTextBox { Name = "txtPage", AutoSize = false, Width = Dpi.Scale(46), Alignment = ToolStripItemAlignment.Right, TextBoxTextAlign = HorizontalAlignment.Center, AccessibleName = Loc.T("Số trang", "Page number") };
        _pageBox.KeyDown += (_, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            e.SuppressKeyPress = true;
            if (int.TryParse(_pageBox.Text, out int p)) Current?.Viewer.GoToPage(p - 1);
            Current?.Viewer.Focus();
        };
        var sbPrev = SB("sbPrevPage", Icons.Prev, Loc.T("Trang trước", "Previous page"), () => Current?.Viewer.PrevPage());
        var sbFirst = SB("sbFirstPage", Icons.First, Loc.T("Trang đầu", "First page"), () => Current?.Viewer.GoToPage(0));
        _status.Items.AddRange([sbFitW, sbFitP, sbZin, _zoomBox, sbZout, sep, sbLast, sbNext, _pageCount, _pageBox, sbPrev, sbFirst]);
        foreach (var it in new ToolStripItem[] { sbFitW, sbFitP, sbZin, _zoomBox, sbZout, sbLast, sbNext, _pageBox, sbPrev, sbFirst }) Cmd(it);

        foreach (var zb in new[] { _zoomBox, _ribbonZoom })
        {
            var box = zb;
            box.SelectedIndexChanged += (_, _) => { if (!_updatingZoom) ApplyZoomText(box.Text); };
            box.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplyZoomText(box.Text); Current?.Viewer.Focus(); } };
            box.ComboBox.PreviewKeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) e.IsInputKey = true; };
            box.ComboBox.Validated += (_, _) => { if (!_updatingZoom && Current != null && box.Text != $"{Math.Round(Current.Viewer.Zoom * 100)}%") ApplyZoomText(box.Text); };
        }

        var host = new Panel { Dock = DockStyle.Fill, Name = "contentHost" };
        host.Controls.Add(_tabs);
        host.Controls.Add(_startPage);
        InitAi(host, home);
        Controls.Add(host);
        Controls.Add(_status);
        Controls.Add(_ribbon);

        // PDFs open; a picture (file, screenshot, image dragged from a browser or chat app) goes onto the open page as an image stamp
        DragEnter += (_, e) => e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true || (Current != null && DataTransfer.HasContent(e.Data)) ? DragDropEffects.Copy : DragDropEffects.None;
        DragOver += (_, e) => e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true || (Current != null && DataTransfer.HasContent(e.Data)) ? DragDropEffects.Copy : DragDropEffects.None;
        DragDrop += (_, e) =>
        {
            var files = DataTransfer.Files(e.Data);
            var pdfs = files.Where(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) && File.Exists(f)).ToList();
            foreach (var f in pdfs) OpenFile(f);
            if (pdfs.Count > 0) return;
            if (!PlaceImageFrom(e.Data) && (files.Length > 0 || DataTransfer.Picture(e.Data) == null))
                SetStatus(Loc.T("Chỉ mở được tệp PDF; ảnh được chèn vào trang PDF đang mở.", "Only PDF files can be opened; pictures are inserted into the open PDF page."));
        };
        Load += (_, _) =>
        {
            RestoreWindow();
            foreach (var f in files) OpenFile(f);
            OnCurrentChanged();
        };
        FormClosing += OnClosingForm;
    }

    private T Cmd<T>(T item) where T : ToolStripItem
    {
        _cmd[item.Name!] = item;
        if (item is ToolStripButton) item.Click += (_, _) => { if (!IsDisposed) BeginInvoke(UpdateStatus); };
        return item;
    }

    private static void FillZoomItems(ToolStripComboBox box)
    {
        box.Items.AddRange(["25%", "50%", "75%", "100%", "125%", "150%", "200%", "300%", "400%", "800%", Loc.T("Vừa trang", "Fit page"), Loc.T("Vừa chiều rộng", "Fit width")]);
    }

    private void ApplyZoomText(string text)
    {
        var v = Current?.Viewer;
        if (v == null) return;
        text = text.Trim();
        if (text == Loc.T("Vừa trang", "Fit page")) { v.SetZoomMode(ZoomMode.FitPage); return; }
        if (text == Loc.T("Vừa chiều rộng", "Fit width")) { v.SetZoomMode(ZoomMode.FitWidth); return; }
        if (float.TryParse(text.TrimEnd('%').Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float pct) && pct > 0)
            v.SetZoom(pct / 100f);
        else UpdateStatus();
    }

    private void BuildFileMenu(out ToolStripMenuItem recent)
    {
        var m = _ribbon.FileButton;
        m.DropDown = new ContextMenuStrip { Renderer = new RibbonRenderer(false), ImageScalingSize = new Size(Dpi.Scale(18), Dpi.Scale(18)) };
        ToolStripMenuItem Item(string name, string text, string glyph, Action a, Keys keys = Keys.None)
        {
            var it = new ToolStripMenuItem(text, Icons.Get(glyph, Dpi.Scale(18), Theme.Ink), (_, _) => a()) { Name = name, ShortcutKeys = keys, ShowShortcutKeys = keys != Keys.None };
            if (keys == Keys.None) it.ShowShortcutKeys = false;
            m.DropDownItems.Add(it);
            _cmd[name] = it;
            return it;
        }
        Item("miOpen", Loc.T("Mở...", "Open..."), Icons.Open, CmdOpen);
        recent = new ToolStripMenuItem(Loc.T("Mở gần đây", "Open recent"), Icons.Get(Icons.Recent, Dpi.Scale(18), Theme.Ink)) { Name = "miRecent" };
        m.DropDownItems.Add(recent);
        m.DropDownItems.Add(new ToolStripSeparator());
        Item("miSave", Loc.T("Lưu", "Save"), Icons.Save, CmdSave);
        Item("miSaveAs", Loc.T("Lưu bản sao...", "Save a copy..."), Icons.SaveAs, CmdSaveAs);
        Item("miPrint", Loc.T("In...", "Print..."), Icons.Print, CmdPrint);
        m.DropDownItems.Add(new ToolStripSeparator());
        Item("miProperties", Loc.T("Thuộc tính tài liệu", "Document properties"), Icons.Properties, CmdProperties);
        Item("miSignSettings", Loc.T("Thiết lập ký số...", "Signing settings..."), Icons.Settings, CmdSignSettings);
        m.DropDownItems.Add(new ToolStripSeparator());
        Item("miCloseTab", Loc.T("Đóng tài liệu", "Close document"), Icons.Close, () => { if (_tabs.SelectedTab != null) CloseTab(_tabs.SelectedTab); });
        m.DropDownItems.Add(new ToolStripSeparator());
        Item("miCheckUpdate", Loc.T("Kiểm tra cập nhật...", "Check for updates..."), Icons.Sync, () => OfficeShell.CheckForUpdatesInteractive(this));
        Item("miSettings", Loc.T("Cài đặt && ngôn ngữ...", "Settings && language..."), Icons.Settings, () => OfficeShell.ShowSettings(this));
        Item("miAbout", Loc.T("Giới thiệu", "About"), Icons.Info, () => OfficeShell.ShowAbout(this));
        m.DropDownItems.Add(new ToolStripSeparator());
        Item("miExit", Loc.T("Thoát", "Exit"), Icons.Exit, Close);
        m.DropDownOpening += (_, _) => RefreshRecentMenu();
    }

    private void RefreshRecentMenu()
    {
        _recentMenu.DropDownItems.Clear();
        foreach (var f in _settings.RecentFiles)
        {
            var path = f;
            _recentMenu.DropDownItems.Add(new ToolStripMenuItem(Path.GetFileName(path) + "   —   " + Path.GetDirectoryName(path), null, (_, _) => OpenFile(path)) { ToolTipText = path });
        }
        if (_recentMenu.DropDownItems.Count == 0) _recentMenu.DropDownItems.Add(new ToolStripMenuItem(Loc.T("(trống)", "(empty)")) { Enabled = false });
        else
        {
            _recentMenu.DropDownItems.Add(new ToolStripSeparator());
            _recentMenu.DropDownItems.Add(new ToolStripMenuItem(Loc.T("Xóa danh sách", "Clear list"), null, (_, _) => { _settings.RecentFiles.Clear(); _settings.Save(); BuildRecentList(); }) { Name = "miClearRecent" });
        }
    }

    private void BuildStartPage()
    {
        var center = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Name = "startCenter", BackColor = Color.White };
        var logo = new PictureBox { Image = BrandArt.RenderWordmark(Dpi.Scale(110), false, AppKind.Pdf), Name = "picStartLogo", SizeMode = PictureBoxSizeMode.AutoSize, Margin = new Padding(0, 0, 0, Dpi.Scale(10)) };
        var slogan = new Label { Text = Loc.T("Đọc • Ký số • In ấn tài liệu PDF — nhanh, chuẩn, an toàn", "Read • Sign • Print PDF documents — fast, standard, secure"), AutoSize = true, Font = new Font("Segoe UI", 12f), ForeColor = Theme.SubtleInk, Margin = new Padding(Dpi.Scale(8), 0, 0, Dpi.Scale(20)) };
        var open = new Button
        {
            Text = Loc.T("   Mở tệp PDF...", "   Open a PDF file..."), Name = "btnStartOpen", Image = Icons.Get(Icons.Open, Dpi.Scale(22), Color.White), ImageAlign = ContentAlignment.MiddleLeft,
            TextImageRelation = TextImageRelation.ImageBeforeText, AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Theme.Accent, ForeColor = Color.White,
            Font = new Font("Segoe UI Semibold", 12f), Padding = new Padding(Dpi.Scale(14), Dpi.Scale(6), Dpi.Scale(14), Dpi.Scale(6)), Cursor = Cursors.Hand, Margin = new Padding(Dpi.Scale(8), 0, 0, Dpi.Scale(8))
        };
        open.FlatAppearance.BorderSize = 0;
        open.Click += (_, _) => CmdOpen();
        var hint = new Label { Text = Loc.T("hoặc kéo-thả tệp PDF vào cửa sổ này", "or drag && drop PDF files into this window"), AutoSize = true, ForeColor = Theme.SubtleInk, Margin = new Padding(Dpi.Scale(10), 0, 0, Dpi.Scale(24)) };
        var recentTitle = new Label { Text = Loc.T("Tài liệu gần đây", "Recent documents"), AutoSize = true, Font = Theme.TitleFont, ForeColor = Theme.Ink, Margin = new Padding(Dpi.Scale(8), 0, 0, Dpi.Scale(6)) };
        center.Controls.AddRange([logo, slogan, open, hint, recentTitle, _recentPanel]);
        _startPage.Controls.Add(center);
        var promo = new PromoStrip { Name = "promoStrip" };
        _startPage.Controls.Add(promo);
        void LayoutStart()
        {
            var cs = _startPage.ClientSize;
            int pw = Math.Min(Dpi.Scale(900), cs.Width - Dpi.Scale(40));
            center.Location = new Point(Math.Max(Dpi.Scale(20), (cs.Width - center.Width) / 2), Math.Max(Dpi.Scale(20), (cs.Height - center.Height - promo.Height) / 3));
            promo.Width = Math.Max(Dpi.Scale(420), pw);
            promo.Location = new Point(Math.Max(Dpi.Scale(20), (cs.Width - promo.Width) / 2),
                Math.Max(center.Bottom + Dpi.Scale(24), cs.Height - promo.Height - Dpi.Scale(20)));
        }
        _startPage.Resize += (_, _) => LayoutStart();
        center.SizeChanged += (_, _) => LayoutStart();
        BuildRecentList();
    }

    private void BuildRecentList()
    {
        _recentPanel.Controls.Clear();
        int n = 0;
        foreach (var f in _settings.RecentFiles.Take(10))
        {
            var path = f;
            var link = new LinkLabel
            {
                Text = $"{Path.GetFileName(path)}   ({Path.GetDirectoryName(path)})", AutoSize = true, LinkColor = Theme.AccentDark,
                Font = new Font("Segoe UI", 10f), Margin = new Padding(Dpi.Scale(10), Dpi.Scale(3), 0, Dpi.Scale(3)), Name = "recent" + n++, LinkBehavior = LinkBehavior.HoverUnderline
            };
            link.LinkClicked += (_, _) => OpenFile(path);
            link.MouseUp += (_, e) => { if (e.Button == MouseButtons.Left) BeginInvoke(() => OpenFile(path)); }; // whole row clickable
            _recentPanel.Controls.Add(link);
        }
        if (n == 0) _recentPanel.Controls.Add(new Label { Text = Loc.T("(chưa có)", "(none yet)"), AutoSize = true, ForeColor = Theme.SubtleInk, Margin = new Padding(Dpi.Scale(10), 0, 0, 0) });
    }

    private void RestoreWindow()
    {
        if (_settings.WindowBounds is { Length: 4 } b)
        {
            var r = new Rectangle(b[0], b[1], b[2], b[3]);
            if (Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(r)) && r.Width > 300 && r.Height > 200) { StartPosition = FormStartPosition.Manual; Bounds = r; }
        }
        if (_settings.WindowState == (int)FormWindowState.Maximized) WindowState = FormWindowState.Maximized;
    }

    private void OnClosingForm(object? sender, FormClosingEventArgs e)
    {
        foreach (TabPage p in _tabs.TabPages.Cast<TabPage>().ToList())
        {
            var v = p.Controls.OfType<DocumentView>().FirstOrDefault();
            if (v is { Modified: true })
            {
                _tabs.SelectedTab = p;
                var r = Msg.AskSave(this, v.DisplayName);
                if (r == SaveChoice.Cancel) { e.Cancel = true; return; }
                if (r == SaveChoice.Save && !SaveView(v, false)) { e.Cancel = true; return; }
            }
        }
        if (_fullScreen) CmdFullScreen();
        _settings.WindowState = (int)WindowState;
        var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        _settings.WindowBounds = [bounds.X, bounds.Y, bounds.Width, bounds.Height];
        _settings.Save();
    }
}

/// <summary>Document tab strip with a close button on each tab and middle-click to close.</summary>
public sealed class DocTabControl : TabControl
{
    public event Action<TabPage>? CloseRequested;

    public DocTabControl()
    {
        DrawMode = TabDrawMode.OwnerDrawFixed;
        SizeMode = TabSizeMode.Fixed;
        ItemSize = new Size(Dpi.Scale(200), Dpi.Scale(28));
        Padding = new Point(Dpi.Scale(12), Dpi.Scale(4));
        Font = Theme.UiFont;
        var menu = new ContextMenuStrip { Renderer = new RibbonRenderer(false) };
        menu.Items.Add(Loc.T("Đóng", "Close"), null, (_, _) => { if (_menuTab != null) CloseRequested?.Invoke(_menuTab); });
        menu.Items.Add(Loc.T("Đóng các thẻ khác", "Close other tabs"), null, (_, _) =>
        {
            foreach (var p in TabPages.Cast<TabPage>().Where(p => p != _menuTab).ToList()) CloseRequested?.Invoke(p);
        });
        _menu = menu;
    }

    private readonly ContextMenuStrip _menu;
    private TabPage? _menuTab;

    private Rectangle CloseRect(int i)
    {
        var r = GetTabRect(i);
        int s = Dpi.Scale(16);
        return new Rectangle(r.Right - s - Dpi.Scale(6), r.Top + (r.Height - s) / 2, s, s);
    }

    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        var g = e.Graphics;
        var r = GetTabRect(e.Index);
        bool sel = e.Index == SelectedIndex;
        using (var back = new SolidBrush(sel ? Color.White : Theme.TabStripBack)) g.FillRectangle(back, r);
        if (sel) using (var acc = new SolidBrush(Theme.Accent)) g.FillRectangle(acc, r.X, r.Y, r.Width, Dpi.Scale(3));
        var page = TabPages[e.Index];
        var tr = new Rectangle(r.X + Dpi.Scale(8), r.Y + 2, r.Width - Dpi.Scale(34), r.Height - 2);
        TextRenderer.DrawText(g, page.Text, sel ? Theme.UiFontBold : Theme.UiFont, tr, Theme.Ink, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.Left);
        var cr = CloseRect(e.Index);
        g.DrawImage(Icons.Get(Icons.Close, cr.Width, Theme.SubtleInk), cr);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        for (int i = 0; i < TabCount; i++)
        {
            if (!GetTabRect(i).Contains(e.Location)) continue;
            if (e.Button == MouseButtons.Middle || (e.Button == MouseButtons.Left && CloseRect(i).Contains(e.Location)))
            {
                CloseRequested?.Invoke(TabPages[i]);
                return;
            }
            if (e.Button == MouseButtons.Right)
            {
                _menuTab = TabPages[i];
                _menu.Show(this, e.Location);
                return;
            }
        }
        base.OnMouseDown(e);
    }
}
