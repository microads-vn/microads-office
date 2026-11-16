using MicroadsOffice.Core.Office;
using System.Drawing.Printing;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Dialogs;
using MicroadsOffice.UI;

namespace MicroadsOffice.Printing;

/// <summary>Foxit-like print dialog with live preview.</summary>
public sealed class PrintDialogForm : DialogBase
{
    private readonly PdfiumDocument _doc;
    private readonly int _currentPage;
    private readonly ComboBox _printer;
    private readonly NumericUpDown _copies;
    private readonly CheckBox _collate, _autoRotate, _gray, _annots, _toFile;
    private readonly RadioButton _all, _current, _range, _fit, _actual, _shrink, _portrait, _landscape;
    private readonly TextBox _rangeText, _filePath;
    private readonly PictureBox _preview;
    private readonly Label _previewLabel;
    private int _previewIndex;
    private PageSettings _pageSettings;
    private PrinterSettings? _driverSettings;

    public PrintOptions? Options { get; private set; }

    public PrintDialogForm(PdfiumDocument doc, int currentPage) : base(Loc.T("In", "Print"), 860, 560)
    {
        _doc = doc;
        _currentPage = currentPage;
        var ps = new PrinterSettings();
        _pageSettings = ps.DefaultPageSettings;

        var left = new TableLayoutPanel { Dock = DockStyle.Left, Width = Dpi.Scale(450), ColumnCount = 2, AutoScroll = true };
        left.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _printer = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(240), Name = "cboPrinter" };
        foreach (string p in PrinterSettings.InstalledPrinters) _printer.Items.Add(p);
        _printer.SelectedItem = ps.PrinterName;
        if (_printer.SelectedIndex < 0 && _printer.Items.Count > 0) _printer.SelectedIndex = 0;
        var prRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        var props = new Button { Text = Loc.T("Thuộc tính...", "Properties..."), AutoSize = true, Name = "btnPrinterProperties" };
        props.Click += (_, _) => ShowPrinterProperties();
        prRow.Controls.AddRange([_printer, props]);
        left.Controls.Add(L(Loc.T("Máy in:", "Printer:"), true)); left.Controls.Add(prRow);

        var cpRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _copies = new NumericUpDown { Minimum = 1, Maximum = 999, Value = 1, Width = Dpi.Scale(60), Name = "numCopies" };
        _collate = new CheckBox { Text = Loc.T("Sắp xếp theo bộ", "Collate"), Checked = true, AutoSize = true, Name = "chkCollate" };
        cpRow.Controls.AddRange([_copies, _collate]);
        left.Controls.Add(L(Loc.T("Số bản:", "Copies:"))); left.Controls.Add(cpRow);

        var rangeBox = new GroupBox { Text = Loc.T("Phạm vi in", "Print range"), Width = Dpi.Scale(420), Height = Dpi.Scale(112) };
        var rf = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        _all = new RadioButton { Text = Loc.T($"Tất cả ({doc.PageCount} trang)", $"All ({doc.PageCount} pages)"), Checked = true, AutoSize = true, Name = "rdoAllPages" };
        _current = new RadioButton { Text = Loc.T($"Trang hiện tại ({currentPage + 1})", $"Current page ({currentPage + 1})"), AutoSize = true, Name = "rdoCurrentPage" };
        var rr = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _range = new RadioButton { Text = Loc.T("Các trang:", "Pages:"), AutoSize = true, Name = "rdoPageRange" };
        _rangeText = new TextBox { Width = Dpi.Scale(160), PlaceholderText = Loc.T("VD: 1-3, 5", "e.g. 1-3, 5"), Name = "txtPageRange" };
        _rangeText.TextChanged += (_, _) => { if (_rangeText.Focused) _range.Checked = true; UpdatePreview(); };
        rr.Controls.AddRange([_range, _rangeText]);
        rf.Controls.AddRange([_all, _current, rr]);
        rangeBox.Controls.Add(rf);
        left.Controls.Add(rangeBox); left.SetColumnSpan(rangeBox, 2);

        var sizeBox = new GroupBox { Text = Loc.T("Kích thước và hướng trang", "Page sizing && orientation"), Width = Dpi.Scale(420), Height = Dpi.Scale(150) };
        var sf = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        _fit = new RadioButton { Text = Loc.T("Vừa khổ giấy", "Fit to paper"), Checked = true, AutoSize = true, Name = "rdoFit" };
        _actual = new RadioButton { Text = Loc.T("Kích thước thực (100%)", "Actual size (100%)"), AutoSize = true, Name = "rdoActualSize" };
        _shrink = new RadioButton { Text = Loc.T("Chỉ thu nhỏ trang quá khổ", "Shrink oversized pages"), AutoSize = true, Name = "rdoShrink" };
        var orRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        var orPanel = new Panel { AutoSize = true };
        _portrait = new RadioButton { Text = Loc.T("Giấy dọc", "Portrait"), AutoSize = true, Name = "rdoPortrait", Checked = !_pageSettings.Landscape };
        _landscape = new RadioButton { Text = Loc.T("Giấy ngang", "Landscape"), AutoSize = true, Name = "rdoLandscape", Checked = _pageSettings.Landscape };
        orRow.Controls.AddRange([_portrait, _landscape]);
        _autoRotate = new CheckBox { Text = Loc.T("Tự động xoay trang cho phù hợp", "Auto-rotate pages to fit"), Checked = true, AutoSize = true, Name = "chkAutoRotate" };
        sf.Controls.AddRange([_fit, _actual, _shrink, orRow, _autoRotate]);
        sizeBox.Controls.Add(sf);
        left.Controls.Add(sizeBox); left.SetColumnSpan(sizeBox, 2);

        _gray = new CheckBox { Text = Loc.T("In đen trắng (thang xám)", "Print in black && white (grayscale)"), AutoSize = true, Name = "chkGrayscale" };
        _annots = new CheckBox { Text = Loc.T("In chú thích, con dấu và chữ ký", "Print annotations, stamps and signatures"), Checked = true, AutoSize = true, Name = "chkPrintAnnotations" };
        var fileRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _toFile = new CheckBox { Text = Loc.T("In ra tệp:", "Print to file:"), AutoSize = true, Name = "chkPrintToFile" };
        _filePath = new TextBox { Width = Dpi.Scale(250), Name = "txtPrintFile", PlaceholderText = Loc.T("Đường dẫn tệp đầu ra", "Output file path") };
        fileRow.Controls.AddRange([_toFile, _filePath]);
        var opts = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = Padding.Empty };
        opts.Controls.AddRange([_gray, _annots, fileRow]);
        left.Controls.Add(L(Loc.T("Tùy chọn:", "Options:"))); left.Controls.Add(opts);

        var right = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Dpi.Scale(12), 0, 0, 0) };
        _preview = new PictureBox { Dock = DockStyle.Fill, BackColor = Theme.ViewerBack, SizeMode = PictureBoxSizeMode.CenterImage, Name = "picPrintPreview" };
        var nav = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = Dpi.Scale(36), FlowDirection = FlowDirection.LeftToRight };
        var prev = new Button { Text = "◀", Width = Dpi.Scale(34), Name = "btnPreviewPrev", AccessibleName = Loc.T("Trang xem trước trước", "Previous preview page") };
        var next = new Button { Text = "▶", Width = Dpi.Scale(34), Name = "btnPreviewNext", AccessibleName = Loc.T("Trang xem trước sau", "Next preview page") };
        _previewLabel = new Label { AutoSize = true, Margin = new Padding(8, 9, 0, 0), Name = "lblPreviewPage" };
        prev.Click += (_, _) => { _previewIndex--; UpdatePreview(); };
        next.Click += (_, _) => { _previewIndex++; UpdatePreview(); };
        nav.Controls.AddRange([prev, next, _previewLabel]);
        right.Controls.Add(_preview);
        right.Controls.Add(nav);
        Body.Controls.Add(right);
        Body.Controls.Add(left);

        // "Các trang" lives in its own row container: keep the page-range options mutually exclusive manually.
        var rangeGroup = new[] { _all, _current, _range };
        foreach (var rb in rangeGroup)
            rb.CheckedChanged += (s, _) =>
            {
                if (s is RadioButton { Checked: true } on)
                    foreach (var other in rangeGroup) if (other != on && other.Checked) other.Checked = false;
            };
        foreach (var rb in new[] { _all, _current, _range, _fit, _actual, _shrink, _portrait, _landscape })
            rb.CheckedChanged += (_, _) => UpdatePreview();
        foreach (var cb in new[] { _autoRotate, _gray, _annots }) cb.CheckedChanged += (_, _) => UpdatePreview();
        _printer.SelectedIndexChanged += (_, _) =>
        {
            try
            {
                var s = new PrinterSettings { PrinterName = _printer.SelectedItem as string ?? "" };
                _pageSettings = s.DefaultPageSettings;
                UpdatePreview();
            }
            catch { }
        };
        Shown += (_, _) => UpdatePreview();

        AddButton(Loc.T("Hủy", "Cancel"), "btnCancel", DialogResult.Cancel);
        var print = AddButton(Loc.T("In", "Print"), "btnPrint", DialogResult.None, true, (_, _) => Accept());
        AddButton(Loc.T("Xem trước khi in...", "Print preview..."), "btnPrintPreview", DialogResult.None, false, (_, _) => ShowPreviewDialog());
        AcceptButton = print;
    }

    private List<int>? SelectedPages()
    {
        if (_all.Checked) return Enumerable.Range(0, _doc.PageCount).ToList();
        if (_current.Checked) return [_currentPage];
        return PdfPrinter.ParseRange(_rangeText.Text, _doc.PageCount);
    }

    private PrintOptions? Build(bool validate)
    {
        var pages = SelectedPages();
        if (pages == null)
        {
            if (validate) Msg.Error(this, Loc.T($"Phạm vi trang không hợp lệ. Nhập theo dạng 1-3, 5 (từ 1 đến {_doc.PageCount}).", $"Invalid page range. Use a form like 1-3, 5 (from 1 to {_doc.PageCount})."));
            return null;
        }
        var ps = (PageSettings)_pageSettings.Clone();
        ps.Landscape = _landscape.Checked;
        return new PrintOptions
        {
            PrinterName = _printer.SelectedItem as string ?? new PrinterSettings().PrinterName,
            Copies = (short)_copies.Value,
            Collate = _collate.Checked,
            Pages = pages,
            Scaling = _fit.Checked ? PrintScaling.Fit : _actual.Checked ? PrintScaling.ActualSize : PrintScaling.ShrinkOversized,
            AutoRotate = _autoRotate.Checked,
            Grayscale = _gray.Checked,
            Annotations = _annots.Checked,
            PrintToFile = _toFile.Checked,
            OutputFile = _filePath.Text.Trim(),
            PageSettings = ps,
            DriverSettings = _driverSettings
        };
    }

    private void Accept()
    {
        var o = Build(true);
        if (o == null) return;
        if (o.PrintToFile && string.IsNullOrWhiteSpace(o.OutputFile)) { Msg.Error(this, Loc.T("Vui lòng nhập đường dẫn tệp đầu ra.", "Please enter the output file path.")); return; }
        if (_printer.Items.Count == 0) { Msg.Error(this, Loc.T("Máy tính chưa cài máy in nào.", "No printer is installed on this computer.")); return; }
        Options = o;
        DialogResult = DialogResult.OK;
    }

    [System.Runtime.InteropServices.DllImport("winspool.drv", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
    private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);
    [System.Runtime.InteropServices.DllImport("winspool.drv", SetLastError = true)]
    private static extern bool ClosePrinter(IntPtr hPrinter);
    [System.Runtime.InteropServices.DllImport("winspool.drv", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
    private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter, string pDeviceName, IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")] private static extern IntPtr GlobalLock(IntPtr hMem);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")] private static extern bool GlobalUnlock(IntPtr hMem);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")] private static extern IntPtr GlobalFree(IntPtr hMem);

    /// <summary>Opens the printer driver's own properties sheet (paper, quality, duplex, tray...) like Foxit.</summary>
    private void ShowPrinterProperties()
    {
        var name = _printer.SelectedItem as string;
        if (string.IsNullOrEmpty(name)) return;
        var ps = _driverSettings != null && _driverSettings.PrinterName == name ? (PrinterSettings)_driverSettings.Clone() : new PrinterSettings { PrinterName = name };
        if (!OpenPrinter(name, out var hPrinter, IntPtr.Zero)) { Msg.Error(this, Loc.T("Không mở được máy in: ", "Could not open the printer: ") + name); return; }
        IntPtr hDevMode = IntPtr.Zero, outBuf = IntPtr.Zero;
        try
        {
            hDevMode = ps.GetHdevmode(ps.DefaultPageSettings);
            var pIn = GlobalLock(hDevMode);
            int size = DocumentProperties(Handle, hPrinter, name, IntPtr.Zero, IntPtr.Zero, 0);
            if (size <= 0) { GlobalUnlock(hDevMode); Msg.Error(this, Loc.T("Máy in không cung cấp bảng thuộc tính.", "The printer has no properties sheet.")); return; }
            outBuf = System.Runtime.InteropServices.Marshal.AllocHGlobal(size);
            const int DM_OUT_BUFFER = 2, DM_IN_PROMPT = 4, DM_IN_BUFFER = 8;
            int r = DocumentProperties(Handle, hPrinter, name, outBuf, pIn, DM_IN_BUFFER | DM_IN_PROMPT | DM_OUT_BUFFER);
            GlobalUnlock(hDevMode);
            if (r != 1) return; // IDOK only
            ps.SetHdevmode(outBuf);
            ps.DefaultPageSettings.SetHdevmode(outBuf);
            _driverSettings = ps;
            _pageSettings = ps.DefaultPageSettings;
            _copies.Value = Math.Clamp(ps.Copies, (short)1, (short)999);
            _collate.Checked = ps.Collate;
            _landscape.Checked = _pageSettings.Landscape;
            _portrait.Checked = !_pageSettings.Landscape;
            UpdatePreview();
        }
        finally
        {
            if (outBuf != IntPtr.Zero) System.Runtime.InteropServices.Marshal.FreeHGlobal(outBuf);
            if (hDevMode != IntPtr.Zero) GlobalFree(hDevMode);
            ClosePrinter(hPrinter);
        }
    }

    private void ShowPreviewDialog()
    {
        var o = Build(true);
        if (o == null) return;
        o.PrintToFile = false;
        using var doc = PdfPrinter.CreateDocument(_doc, o, Loc.T("Xem trước", "Preview"));
        using var dlg = new PrintPreviewDialog { Document = doc, Width = Dpi.Scale(900), Height = Dpi.Scale(700), UseAntiAlias = true, Text = Loc.T("Xem trước khi in", "Print preview") };
        dlg.ShowDialog(this);
    }

    private void UpdatePreview()
    {
        var o = Build(false);
        var pages = o?.Pages ?? [];
        if (pages.Count == 0) { _previewLabel.Text = Loc.T("Phạm vi trang chưa hợp lệ", "Invalid page range"); _preview.Image = null; return; }
        _previewIndex = Math.Clamp(_previewIndex, 0, pages.Count - 1);
        int page = pages[_previewIndex];
        _previewLabel.Text = Loc.T($"Trang in {_previewIndex + 1}/{pages.Count} (trang PDF {page + 1})", $"Printed page {_previewIndex + 1}/{pages.Count} (PDF page {page + 1})");
        var paperSize = _pageSettings.PaperSize;
        float pw = paperSize.Width, ph = paperSize.Height; // 1/100 inch
        if (_landscape.Checked) (pw, ph) = (ph, pw);
        int availW = Math.Max(50, _preview.Width - 20), availH = Math.Max(50, _preview.Height - 20);
        float k = Math.Min(availW / pw, availH / ph);
        var bmp = new Bitmap(Math.Max(1, (int)(pw * k)) + 6, Math.Max(1, (int)(ph * k)) + 6);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Transparent);
            using (var sh = new SolidBrush(Color.FromArgb(50, 0, 0, 0))) g.FillRectangle(sh, 4, 4, pw * k, ph * k);
            g.FillRectangle(Brushes.White, 0, 0, pw * k, ph * k);
            g.DrawRectangle(Pens.Gray, 0, 0, pw * k, ph * k);
            float margin = 25; // ~6mm hard margin estimate
            var printable = new RectangleF(margin, margin, pw - 2 * margin, ph - 2 * margin);
            var dest = PdfPrinter.Layout(_doc.GetPageSize(page), printable, o!, pw > ph, out bool rotate);
            var dr = new RectangleF(dest.X * k, dest.Y * k, dest.Width * k, dest.Height * k);
            try
            {
                using var pageBmp = _doc.Render(page, Math.Max(1, (int)dr.Width * 2), Math.Max(1, (int)dr.Height * 2), rotate ? 1 : 0, null, o!.Annotations, true, o.Grayscale);
                g.DrawImage(pageBmp, dr);
            }
            catch { }
            using var dash = new Pen(Color.LightGray) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot };
            g.DrawRectangle(dash, printable.X * k, printable.Y * k, printable.Width * k, printable.Height * k);
        }
        var old = _preview.Image;
        _preview.Image = bmp;
        old?.Dispose();
    }
}
