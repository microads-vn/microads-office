using System.Drawing.Text;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Single-line input (row height, sheet name, go to...).</summary>
public sealed class AskDialog : DialogBase
{
    public readonly TextBox Box;
    public AskDialog(string title, string prompt, string initial) : base(title, 400, 150)
    {
        Name = "askDialog";
        Box = new TextBox { Name = "txtAsk", Text = initial, Dock = DockStyle.Top, Font = new Font("Segoe UI", 10f) };
        Body.Controls.Add(Box);
        Body.Controls.Add(new Label { Text = prompt, Dock = DockStyle.Top, Height = Dpi.Scale(28), Name = "lblAsk" });
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
        Shown += (_, _) => { Box.Focus(); Box.SelectAll(); };
    }

    public static string? Ask(IWin32Window owner, string title, string prompt, string initial = "")
    {
        using var d = new AskDialog(title, prompt, initial);
        return Ux.Show(d, owner) == DialogResult.OK ? d.Box.Text : null;
    }
}

// ====================================================================== Format Cells

public sealed class FormatCellsDialog : DialogBase
{
    private readonly CellStyle _orig;
    private readonly TabControl _tabs;
    // number
    private readonly ListBox _cat = new() { Name = "lstNumberCategory", IntegralHeight = false };
    private readonly NumericUpDown _dec = new() { Name = "numDecimals", Minimum = 0, Maximum = 15, Value = 2, Width = 60 };
    private readonly CheckBox _thousands = new() { Name = "chkThousands", AutoSize = true };
    private readonly ComboBox _symbol = new() { Name = "cboSymbol", DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
    private readonly ComboBox _dateFmt = new() { Name = "cboDateFormat", DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
    private readonly TextBox _code = new() { Name = "txtFormatCode", Width = 300 };
    private readonly Label _preview = new() { Name = "lblPreview", AutoSize = false, Height = 26, BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleLeft };
    private readonly double _sample;
    // alignment
    private readonly ComboBox _h = new() { Name = "cboHAlign", DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly ComboBox _v = new() { Name = "cboVAlign", DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
    private readonly NumericUpDown _indent = new() { Name = "numIndent", Minimum = 0, Maximum = 15, Width = 60 };
    private readonly CheckBox _wrap = new() { Name = "chkWrap", AutoSize = true };
    public readonly CheckBox MergeBox = new() { Name = "chkMerge", AutoSize = true };
    // font
    private readonly ComboBox _font = new() { Name = "cboFontName", Width = 220 };
    private readonly ComboBox _size = new() { Name = "cboFontSize", Width = 70 };
    private readonly CheckBox _b = new() { Name = "chkBold", AutoSize = true }, _i = new() { Name = "chkItalic", AutoSize = true }, _u = new() { Name = "chkUnderline", AutoSize = true }, _s = new() { Name = "chkStrike", AutoSize = true };
    private readonly Button _fontColor = new() { Name = "btnFontColor", Width = 120, FlatStyle = FlatStyle.Flat };
    private int _fontColorValue;
    // border
    private readonly ComboBox _lineStyle = new() { Name = "cboLineStyle", DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
    private string _borderPreset = "";
    // fill
    private readonly Button _fill = new() { Name = "btnFillColor", Width = 140, Height = 30, FlatStyle = FlatStyle.Flat };
    private int _fillValue;
    // protection
    private readonly CheckBox _locked = new() { Name = "chkLocked", AutoSize = true }, _hidden = new() { Name = "chkHidden", AutoSize = true };
    private readonly HashSet<string> _touched = new();

    public static readonly (string Vi, string En)[] Categories =
    [
        ("Chung", "General"), ("Số", "Number"), ("Tiền tệ", "Currency"), ("Kế toán", "Accounting"), ("Ngày", "Date"), ("Giờ", "Time"),
        ("Phần trăm", "Percentage"), ("Phân số", "Fraction"), ("Khoa học", "Scientific"), ("Văn bản", "Text"), ("Tùy chỉnh", "Custom")
    ];

    public FormatCellsDialog(CellStyle current, double sample, bool merged, int startTab = 0) : base(Loc.T("Định dạng ô", "Format Cells"), 600, 470)
    {
        Name = "formatCellsDialog";
        _orig = current;
        _sample = sample;
        _tabs = new TabControl { Dock = DockStyle.Fill, Name = "tabsFormat" };
        _tabs.TabPages.Add(NumberPage());
        _tabs.TabPages.Add(AlignPage(merged));
        _tabs.TabPages.Add(FontPage());
        _tabs.TabPages.Add(BorderPage());
        _tabs.TabPages.Add(FillPage());
        _tabs.TabPages.Add(ProtectionPage());
        _tabs.SelectedIndex = Math.Clamp(startTab, 0, 5);
        Body.Controls.Add(_tabs);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
        LoadNumber(current.NumberFormat);
    }

    private static TabPage Page(string vi, string en, string name) => new(Loc.T(vi, en)) { Name = name, Padding = new Padding(Dpi.Scale(10)), BackColor = Color.White };

    private TabPage NumberPage()
    {
        var p = Page("Số", "Number", "tabNumber");
        foreach (var c in Categories) _cat.Items.Add(Loc.T(c.Vi, c.En));
        _cat.SetBounds(Dpi.Scale(10), Dpi.Scale(10), Dpi.Scale(150), Dpi.Scale(300));
        var right = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, Location = new Point(Dpi.Scale(175), Dpi.Scale(6)), Size = new Size(Dpi.Scale(380), Dpi.Scale(320)) };
        _thousands.Text = Loc.T("Dùng dấu phân cách hàng nghìn", "Use 1000 separator");
        _symbol.Items.AddRange(["₫", "VND", "$", "€", "¥", Loc.T("(không)", "(none)")]);
        _symbol.SelectedIndex = 0;
        foreach (var f in new[] { "dd/mm/yyyy", "d/m/yyyy", "dd-mm-yyyy", "yyyy-mm-dd", "dd/mm/yy", "d-mmm-yy", "mmmm yyyy", "dddd, dd/mm/yyyy", "m/d/yyyy", "dd/mm/yyyy hh:mm" }) _dateFmt.Items.Add(f);
        _dateFmt.SelectedIndex = 0;
        right.Controls.Add(new Label { Text = Loc.T("Mẫu:", "Sample:"), AutoSize = true });
        _preview.Width = Dpi.Scale(360);
        right.Controls.Add(_preview);
        var decRow = new FlowLayoutPanel { AutoSize = true };
        decRow.Controls.Add(new Label { Text = Loc.T("Số chữ số thập phân:", "Decimal places:"), AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        decRow.Controls.Add(_dec);
        right.Controls.Add(decRow);
        right.Controls.Add(_thousands);
        var symRow = new FlowLayoutPanel { AutoSize = true };
        symRow.Controls.Add(new Label { Text = Loc.T("Ký hiệu:", "Symbol:"), AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
        symRow.Controls.Add(_symbol);
        right.Controls.Add(symRow);
        right.Controls.Add(new Label { Text = Loc.T("Kiểu ngày/giờ:", "Date/time type:"), AutoSize = true });
        right.Controls.Add(_dateFmt);
        right.Controls.Add(new Label { Text = Loc.T("Mã định dạng:", "Format code:"), AutoSize = true, Margin = new Padding(0, 8, 0, 0) });
        right.Controls.Add(_code);
        right.Controls.Add(new Label
        {
            Text = Loc.T("Mã gồm: 0 # , . % ; \"chữ\" [Red] d m y h s @", "Codes: 0 # , . % ; \"text\" [Red] d m y h s @"),
            AutoSize = true, ForeColor = Theme.SubtleInk
        });
        p.Controls.Add(_cat);
        p.Controls.Add(right);
        EventHandler rebuild = (_, _) => { _touched.Add("num"); RebuildCode(); };
        _cat.SelectedIndexChanged += rebuild;
        _dec.ValueChanged += rebuild;
        _thousands.CheckedChanged += rebuild;
        _symbol.SelectedIndexChanged += rebuild;
        _dateFmt.SelectedIndexChanged += rebuild;
        _code.TextChanged += (_, _) => { _touched.Add("num"); UpdatePreview(); };
        return p;
    }

    private bool _loading;

    private void LoadNumber(string code)
    {
        _loading = true;
        int cat = 10;
        if (code == "General") cat = 0;
        else if (code == "@") cat = 9;
        else if (NumFmt.IsDateFormat(code)) cat = code.Contains('h') && !code.Contains('d') ? 5 : 4;
        else if (code.Contains('%')) cat = 6;
        else if (code.Contains('E')) cat = 8;
        else if (code.Contains('/')) cat = 7;
        else if (code.Contains("₫") || code.Contains('$') || code.Contains("VND") || code.Contains('€')) cat = code.Contains('_') || code.Contains('*') ? 3 : 2;
        else if (code.Contains('0')) cat = 1;
        _cat.SelectedIndex = cat;
        int dot = code.IndexOf('.');
        if (dot >= 0) { int n = 0; while (dot + 1 + n < code.Length && code[dot + 1 + n] == '0') n++; _dec.Value = n; }
        else _dec.Value = 0;
        _thousands.Checked = code.Contains("#,##");
        if (cat == 4 && _dateFmt.Items.Contains(code)) _dateFmt.SelectedItem = code;
        _code.Text = code;
        _loading = false;
        _touched.Remove("num");
        UpdatePreview();
    }

    private void RebuildCode()
    {
        if (_loading) return;
        int d = (int)_dec.Value;
        string dec = d > 0 ? "." + new string('0', d) : "";
        string sym = _symbol.SelectedIndex is >= 0 and < 5 ? (string)_symbol.SelectedItem! : "";
        string code = _cat.SelectedIndex switch
        {
            0 => "General",
            1 => (_thousands.Checked ? "#,##0" : "0") + dec,
            2 => sym is "$" or "€" or "¥" ? "\"" + sym + "\"#,##0" + dec : sym.Length > 0 ? "#,##0" + dec + " \"" + sym + "\"" : "#,##0" + dec,
            3 => sym is "$" or "€" or "¥" ? "_(\"" + sym + "\"* #,##0" + dec + "_);_(\"" + sym + "\"* (#,##0" + dec + ");_(\"" + sym + "\"* \"-\"??_);_(@_)" : "_-* #,##0" + dec + " \"" + sym + "\"_-;-* #,##0" + dec + " \"" + sym + "\"_-;_-* \"-\"??_-;_-@_-",
            4 => (string)_dateFmt.SelectedItem!,
            5 => "hh:mm:ss",
            6 => "0" + dec + "%",
            7 => "# ?/?",
            8 => "0" + dec + "E+00",
            9 => "@",
            _ => _code.Text
        };
        _loading = true;
        _code.Text = code;
        _loading = false;
        UpdatePreview();
    }

    private void UpdatePreview()
    {
        try { _preview.Text = "  " + NumFmt.Format(CellValue.Number(_sample), _code.Text, out var c); _preview.ForeColor = c ?? Theme.Ink; }
        catch { _preview.Text = "  ?"; }
    }

    private TabPage AlignPage(bool merged)
    {
        var p = Page("Căn lề", "Alignment", "tabAlignment");
        _h.Items.AddRange([Loc.T("Chung", "General"), Loc.T("Trái", "Left"), Loc.T("Giữa", "Center"), Loc.T("Phải", "Right"), Loc.T("Đều hai bên", "Justify"), Loc.T("Giữa vùng chọn", "Center Across Selection")]);
        _v.Items.AddRange([Loc.T("Dưới", "Bottom"), Loc.T("Giữa", "Center"), Loc.T("Trên", "Top")]);
        _h.SelectedIndex = (int)_orig.HAlign; _v.SelectedIndex = (int)_orig.VAlign;
        _indent.Value = _orig.Indent;
        _wrap.Text = Loc.T("Xuống dòng tự động", "Wrap text"); _wrap.Checked = _orig.Wrap;
        MergeBox.Text = Loc.T("Gộp ô", "Merge cells"); MergeBox.Checked = merged;
        var t = Grid();
        t.Controls.Add(L(Loc.T("Căn ngang:", "Horizontal:")), 0, 0); t.Controls.Add(_h, 1, 0);
        t.Controls.Add(L(Loc.T("Căn dọc:", "Vertical:")), 0, 1); t.Controls.Add(_v, 1, 1);
        t.Controls.Add(L(Loc.T("Thụt lề:", "Indent:")), 0, 2); t.Controls.Add(_indent, 1, 2);
        t.Controls.Add(_wrap, 1, 3); t.Controls.Add(MergeBox, 1, 4);
        p.Controls.Add(t);
        _h.SelectedIndexChanged += (_, _) => _touched.Add("align");
        _v.SelectedIndexChanged += (_, _) => _touched.Add("align");
        _indent.ValueChanged += (_, _) => _touched.Add("align");
        _wrap.CheckedChanged += (_, _) => _touched.Add("wrap");
        MergeBox.CheckedChanged += (_, _) => _touched.Add("merge");
        return p;
    }

    public bool MergeTouched => _touched.Contains("merge");

    private TabPage FontPage()
    {
        var p = Page("Phông chữ", "Font", "tabFont");
        using (var fonts = new InstalledFontCollection()) foreach (var f in fonts.Families) _font.Items.Add(f.Name);
        _font.Text = _orig.FontName;
        foreach (var s in new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72 }) _size.Items.Add(s.ToString());
        _size.Text = _orig.FontSize.ToString("0.#");
        _b.Text = Loc.T("Đậm", "Bold"); _b.Checked = _orig.Bold;
        _i.Text = Loc.T("Nghiêng", "Italic"); _i.Checked = _orig.Italic;
        _u.Text = Loc.T("Gạch chân", "Underline"); _u.Checked = _orig.Underline;
        _s.Text = Loc.T("Gạch ngang", "Strikethrough"); _s.Checked = _orig.Strike;
        _fontColorValue = _orig.FontColor;
        PaintColorButton(_fontColor, _fontColorValue, Loc.T("Tự động", "Automatic"));
        _fontColor.Click += (_, _) => new ColorPickerDropDown(c => { _fontColorValue = c?.ToArgb() ?? 0; _touched.Add("fontcolor"); PaintColorButton(_fontColor, _fontColorValue, Loc.T("Tự động", "Automatic")); }, Loc.T("Tự động", "Automatic")).Show(_fontColor, 0, _fontColor.Height);
        var t = Grid();
        t.Controls.Add(L(Loc.T("Phông:", "Font:")), 0, 0); t.Controls.Add(_font, 1, 0);
        t.Controls.Add(L(Loc.T("Cỡ:", "Size:")), 0, 1); t.Controls.Add(_size, 1, 1);
        t.Controls.Add(L(Loc.T("Kiểu:", "Style:")), 0, 2);
        var flow = new FlowLayoutPanel { AutoSize = true };
        flow.Controls.AddRange([_b, _i, _u, _s]);
        t.Controls.Add(flow, 1, 2);
        t.Controls.Add(L(Loc.T("Màu:", "Colour:")), 0, 3); t.Controls.Add(_fontColor, 1, 3);
        p.Controls.Add(t);
        _font.TextChanged += (_, _) => _touched.Add("font");
        _size.TextChanged += (_, _) => _touched.Add("font");
        foreach (var cb in new[] { _b, _i, _u, _s }) cb.CheckedChanged += (_, _) => _touched.Add("fontstyle");
        return p;
    }

    private static void PaintColorButton(Button b, int argb, string emptyText)
    {
        b.BackColor = argb == 0 ? Color.White : Color.FromArgb(255, Color.FromArgb(argb));
        b.ForeColor = argb != 0 && Color.FromArgb(argb).GetBrightness() < 0.5f ? Color.White : Theme.Ink;
        b.Text = argb == 0 ? emptyText : "";
    }

    private TabPage BorderPage()
    {
        var p = Page("Viền", "Border", "tabBorder");
        _lineStyle.Items.AddRange([Loc.T("Mảnh", "Thin"), Loc.T("Vừa", "Medium"), Loc.T("Dày", "Thick"), Loc.T("Nét đứt", "Dashed"), Loc.T("Chấm", "Dotted"), Loc.T("Đôi", "Double")]);
        _lineStyle.SelectedIndex = 0;
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        flow.Controls.Add(new Label { Text = Loc.T("Kiểu đường:", "Line style:"), AutoSize = true });
        flow.Controls.Add(_lineStyle);
        flow.Controls.Add(new Label { Text = Loc.T("Áp dụng:", "Presets:"), AutoSize = true, Margin = new Padding(0, 10, 0, 0) });
        var row = new FlowLayoutPanel { AutoSize = true };
        foreach (var (key, vi, en, icon) in new[] { ("none", "Không viền", "None", "borderNone"), ("outline", "Viền ngoài", "Outline", "borderOutside"), ("all", "Tất cả", "All", "borderAll"), ("bottom", "Dưới", "Bottom", "borderBottom"), ("top", "Trên", "Top", "borderTop") })
        {
            var k = key;
            var b = new Button { Name = "btnBorder_" + key, Text = Loc.T(vi, en), Image = XlIcons.Get(icon, Dpi.Scale(24)), TextImageRelation = TextImageRelation.ImageAboveText, Size = new Size(Dpi.Scale(84), Dpi.Scale(62)), FlatStyle = FlatStyle.Flat };
            b.Click += (_, _) => { _borderPreset = k; _touched.Add("border"); foreach (Control c in row.Controls) c.BackColor = c == b ? Theme.Hover : Color.White; };
            row.Controls.Add(b);
        }
        flow.Controls.Add(row);
        p.Controls.Add(flow);
        return p;
    }

    public string BorderPreset => _borderPreset;
    public LineStyle BorderLineStyle => _lineStyle.SelectedIndex switch { 1 => LineStyle.Medium, 2 => LineStyle.Thick, 3 => LineStyle.Dashed, 4 => LineStyle.Dotted, 5 => LineStyle.Double, _ => LineStyle.Thin };

    private TabPage FillPage()
    {
        var p = Page("Tô nền", "Fill", "tabFill");
        _fillValue = _orig.Fill;
        PaintColorButton(_fill, _fillValue, Loc.T("Không tô", "No fill"));
        _fill.Click += (_, _) => new ColorPickerDropDown(c => { _fillValue = c?.ToArgb() ?? 0; _touched.Add("fill"); PaintColorButton(_fill, _fillValue, Loc.T("Không tô", "No fill")); }, Loc.T("Không tô", "No fill")).Show(_fill, 0, _fill.Height);
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        flow.Controls.Add(new Label { Text = Loc.T("Màu nền:", "Background colour:"), AutoSize = true });
        flow.Controls.Add(_fill);
        p.Controls.Add(flow);
        return p;
    }

    private TabPage ProtectionPage()
    {
        var p = Page("Bảo vệ", "Protection", "tabProtection");
        _locked.Text = Loc.T("Khóa ô", "Locked"); _locked.Checked = _orig.Locked;
        _hidden.Text = Loc.T("Ẩn công thức", "Hidden"); _hidden.Checked = _orig.HiddenFormula;
        _locked.CheckedChanged += (_, _) => _touched.Add("prot");
        _hidden.CheckedChanged += (_, _) => _touched.Add("prot");
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        flow.Controls.AddRange([_locked, _hidden, new Label { AutoSize = true, ForeColor = Theme.SubtleInk, MaximumSize = new Size(Dpi.Scale(500), 0),
            Text = Loc.T("Khóa ô hoặc ẩn công thức chỉ có tác dụng khi trang tính được bảo vệ (Xem lại ▸ Bảo vệ trang tính).", "Locking cells or hiding formulas has no effect until you protect the worksheet (Review ▸ Protect Sheet).") }]);
        p.Controls.Add(flow);
        return p;
    }

    /// <summary>Applies only the properties the user changed.</summary>
    public CellStyle Apply(CellStyle st)
    {
        if (_touched.Contains("num")) st = st with { NumberFormat = string.IsNullOrWhiteSpace(_code.Text) ? "General" : _code.Text };
        if (_touched.Contains("align")) st = st with { HAlign = (HAlign)Math.Max(0, _h.SelectedIndex), VAlign = (VAlign)Math.Max(0, _v.SelectedIndex), Indent = (int)_indent.Value };
        if (_touched.Contains("wrap")) st = st with { Wrap = _wrap.Checked };
        if (_touched.Contains("font"))
        {
            st = st with { FontName = string.IsNullOrWhiteSpace(_font.Text) ? st.FontName : _font.Text };
            if (float.TryParse(_size.Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var sz) && sz >= 1 && sz <= 409) st = st with { FontSize = sz };
        }
        if (_touched.Contains("fontstyle")) st = st with { Bold = _b.Checked, Italic = _i.Checked, Underline = _u.Checked, Strike = _s.Checked };
        if (_touched.Contains("fontcolor")) st = st with { FontColor = _fontColorValue };
        if (_touched.Contains("fill")) st = st with { Fill = _fillValue };
        if (_touched.Contains("prot")) st = st with { Locked = _locked.Checked, HiddenFormula = _hidden.Checked };
        return st;
    }

    public bool BorderTouched => _touched.Contains("border");
    public void SetNumberFormat(string code) { _code.Text = code; _touched.Add("num"); }
    public void SetBold(bool b) { _b.Checked = b; }
}

// ====================================================================== Insert Function

public sealed class InsertFunctionDialog : DialogBase
{
    private readonly TextBox _search = new() { Name = "txtFnSearch", Dock = DockStyle.Top };
    private readonly ComboBox _cat = new() { Name = "cboFnCategory", Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ListBox _list = new() { Name = "lstFunctions", Dock = DockStyle.Fill, IntegralHeight = false };
    private readonly Label _desc = new() { Name = "lblFnDesc", Dock = DockStyle.Bottom, Height = Dpi.Scale(64) };
    public string? Selected => (_list.SelectedItem as FunctionInfo)?.Name;

    public InsertFunctionDialog() : base(Loc.T("Chèn hàm", "Insert Function"), 520, 480)
    {
        Name = "insertFunctionDialog";
        _cat.Items.Add(Loc.T("Tất cả", "All"));
        foreach (FnCategory c in Enum.GetValues<FnCategory>()) _cat.Items.Add(FunctionCatalog.CategoryName(c));
        _cat.SelectedIndex = 0;
        _search.PlaceholderText = Loc.T("Tìm hàm: gõ tên hoặc mô tả (vd. tổng, tra cứu, ngày)...", "Search: type a name or description (e.g. sum, lookup, date)...");
        _desc.Font = Theme.UiFont;
        Body.Controls.Add(_list);
        Body.Controls.Add(new Label { Text = Loc.T("Chọn hàm:", "Select a function:"), Dock = DockStyle.Top, Height = Dpi.Scale(22) });
        Body.Controls.Add(_cat);
        Body.Controls.Add(new Label { Text = Loc.T("Nhóm:", "Category:"), Dock = DockStyle.Top, Height = Dpi.Scale(22) });
        Body.Controls.Add(_search);
        Body.Controls.Add(_desc);
        _list.DisplayMember = "Name";
        _search.TextChanged += (_, _) => Fill();
        _cat.SelectedIndexChanged += (_, _) => Fill();
        _list.SelectedIndexChanged += (_, _) =>
        {
            if (_list.SelectedItem is FunctionInfo f) _desc.Text = f.Signature + "\n" + f.Description + "  (" + FunctionCatalog.CategoryName(f.Category) + ")";
        };
        _list.DoubleClick += (_, _) => { if (Selected != null) { DialogResult = DialogResult.OK; Close(); } };
        AddOkCancel(Loc.T("Chèn", "Insert"), Loc.T("Hủy", "Cancel"));
        Fill();
    }

    public void SetCategory(FnCategory c) => _cat.SelectedIndex = (int)c + 1;
    public void SetSearch(string s) => _search.Text = s;

    private void Fill()
    {
        var q = _search.Text.Trim();
        IEnumerable<FunctionInfo> items = FunctionCatalog.All;
        if (_cat.SelectedIndex > 0) items = items.Where(f => (int)f.Category == _cat.SelectedIndex - 1);
        if (q.Length > 0) items = items.Where(f => f.Name.Contains(q, StringComparison.OrdinalIgnoreCase) || f.DescVi.Contains(q, StringComparison.CurrentCultureIgnoreCase) || f.DescEn.Contains(q, StringComparison.OrdinalIgnoreCase));
        _list.BeginUpdate();
        _list.Items.Clear();
        foreach (var f in items.OrderBy(f => f.Name)) _list.Items.Add(f);
        _list.EndUpdate();
        if (_list.Items.Count > 0) _list.SelectedIndex = 0;
    }
}

// ====================================================================== Name Manager

public sealed class NameManagerDialog : DialogBase
{
    private readonly ListView _lv = new() { Name = "lvNames", Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false };
    private readonly Workbook _wb;
    private readonly Func<string> _defaultRef;

    public NameManagerDialog(Workbook wb, Func<string> defaultRef) : base(Loc.T("Trình quản lý tên", "Name Manager"), 560, 400)
    {
        Name = "nameManagerDialog";
        _wb = wb;
        _defaultRef = defaultRef;
        _lv.Columns.Add(Loc.T("Tên", "Name"), Dpi.Scale(160));
        _lv.Columns.Add(Loc.T("Tham chiếu tới", "Refers To"), Dpi.Scale(320));
        Body.Controls.Add(_lv);
        var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(38) };
        Button B(string name, string text, Action a) { var b = new Button { Name = name, Text = text, AutoSize = true, FlatStyle = FlatStyle.Flat }; b.Click += (_, _) => a(); bar.Controls.Add(b); return b; }
        B("btnNameNew", Loc.T("Mới...", "New..."), () => Edit(null));
        B("btnNameEdit", Loc.T("Sửa...", "Edit..."), () => { if (_lv.SelectedItems.Count > 0) Edit(_lv.SelectedItems[0].Text); });
        B("btnNameDelete", Loc.T("Xóa", "Delete"), () =>
        {
            if (_lv.SelectedItems.Count == 0) return;
            _wb.Names.Remove(_lv.SelectedItems[0].Text);
            Changed = true;
            Fill();
        });
        Body.Controls.Add(bar);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
        Fill();
    }

    public bool Changed { get; private set; }

    private void Fill()
    {
        _lv.Items.Clear();
        foreach (var (k, v) in _wb.Names.OrderBy(x => x.Key)) _lv.Items.Add(new ListViewItem([k, "=" + v]));
    }

    public void Edit(string? name)
    {
        var n = AskDialog.Ask(this, Loc.T("Tên", "Name"), Loc.T("Tên (không dấu cách, bắt đầu bằng chữ):", "Name (no spaces, starts with a letter):"), name ?? "");
        if (string.IsNullOrWhiteSpace(n)) return;
        n = n.Trim();
        if (!System.Text.RegularExpressions.Regex.IsMatch(n, @"^[\p{L}_][\p{L}\p{N}_.]*$") || Addr.TryParseA1(n, out _))
        { Ux.Warn(this, Loc.T("Tên không hợp lệ.", "The name is not valid.")); return; }
        var current = name != null && _wb.Names.TryGetValue(name, out var v) ? "=" + v : "=" + _defaultRef();
        var refers = AskDialog.Ask(this, Loc.T("Tham chiếu tới", "Refers to"), Loc.T("Tham chiếu hoặc công thức:", "Reference or formula:"), current);
        if (refers == null) return;
        if (name != null && name != n) _wb.Names.Remove(name);
        _wb.Names[n] = refers.TrimStart('=');
        Changed = true;
        Fill();
    }
}

// ====================================================================== Sort

public sealed class SortDialog : DialogBase
{
    private readonly CheckBox _header = new() { Name = "chkHeader", AutoSize = true, Dock = DockStyle.Top };
    private readonly FlowLayoutPanel _levels = new() { Name = "sortLevels", Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoScroll = true, WrapContents = false };
    private readonly CellRange _range;
    private readonly Sheet _sheet;

    public SortDialog(Sheet sheet, CellRange range, bool hasHeader) : base(Loc.T("Sắp xếp", "Sort"), 520, 360)
    {
        Name = "sortDialog";
        _sheet = sheet; _range = range;
        _header.Text = Loc.T("Dữ liệu có hàng tiêu đề", "My data has headers");
        _header.Checked = hasHeader;
        _header.CheckedChanged += (_, _) => { foreach (var l in _levels.Controls.OfType<FlowLayoutPanel>()) RefillColumns((ComboBox)l.Controls[1]); };
        var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(38) };
        var add = new Button { Name = "btnAddLevel", Text = Loc.T("+ Thêm mức", "+ Add Level"), AutoSize = true, FlatStyle = FlatStyle.Flat };
        add.Click += (_, _) => AddLevel();
        var del = new Button { Name = "btnDeleteLevel", Text = Loc.T("− Xóa mức", "− Delete Level"), AutoSize = true, FlatStyle = FlatStyle.Flat };
        del.Click += (_, _) => { if (_levels.Controls.Count > 1) _levels.Controls.RemoveAt(_levels.Controls.Count - 1); };
        bar.Controls.AddRange([add, del]);
        Body.Controls.Add(_levels);
        Body.Controls.Add(bar);
        Body.Controls.Add(_header);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
        AddLevel();
    }

    private void RefillColumns(ComboBox cb)
    {
        int sel = Math.Max(0, cb.SelectedIndex);
        cb.Items.Clear();
        for (int c = _range.C1; c <= _range.C2 && c < _range.C1 + 300; c++)
        {
            var name = _header.Checked ? SheetOps.DisplayText(_sheet, _range.R1, c) : "";
            cb.Items.Add(Loc.T("Cột ", "Column ") + Addr.ColName(c) + (name.Length > 0 ? " — " + name : ""));
        }
        cb.SelectedIndex = Math.Min(sel, cb.Items.Count - 1);
    }

    public void AddLevel()
    {
        var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        var col = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(250), Name = "cboSortCol" + _levels.Controls.Count };
        var ord = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(140), Name = "cboSortOrder" + _levels.Controls.Count };
        ord.Items.AddRange([Loc.T("A → Z (tăng dần)", "A to Z (ascending)"), Loc.T("Z → A (giảm dần)", "Z to A (descending)")]);
        ord.SelectedIndex = 0;
        RefillColumns(col);
        if (col.Items.Count > _levels.Controls.Count) col.SelectedIndex = _levels.Controls.Count;
        row.Controls.Add(new Label { Text = _levels.Controls.Count == 0 ? Loc.T("Sắp xếp theo", "Sort by") : Loc.T("Rồi theo", "Then by"), Width = Dpi.Scale(80), Margin = new Padding(0, 6, 0, 0) });
        row.Controls.Add(col);
        row.Controls.Add(ord);
        _levels.Controls.Add(row);
    }

    public bool HasHeader => _header.Checked;
    public List<SheetOps.SortKey> Keys => _levels.Controls.OfType<FlowLayoutPanel>()
        .Select(p => new SheetOps.SortKey(_range.C1 + Math.Max(0, ((ComboBox)p.Controls[1]).SelectedIndex), ((ComboBox)p.Controls[2]).SelectedIndex == 0)).ToList();
    public void SetLevel(int i, int colOffset, bool asc)
    {
        while (_levels.Controls.Count <= i) AddLevel();
        var p = (FlowLayoutPanel)_levels.Controls[i];
        ((ComboBox)p.Controls[1]).SelectedIndex = colOffset;
        ((ComboBox)p.Controls[2]).SelectedIndex = asc ? 0 : 1;
    }
}

// ====================================================================== Paste Special

public sealed class PasteSpecialDialog : DialogBase
{
    private readonly RadioButton[] _modes;
    private readonly CheckBox _transpose = new() { Name = "chkTranspose", AutoSize = true }, _skip = new() { Name = "chkSkipBlanks", AutoSize = true };

    public PasteSpecialDialog() : base(Loc.T("Dán đặc biệt", "Paste Special"), 360, 330)
    {
        Name = "pasteSpecialDialog";
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        _modes =
        [
            new RadioButton { Name = "rbAll", Text = Loc.T("Tất cả", "All"), Checked = true, AutoSize = true },
            new RadioButton { Name = "rbValues", Text = Loc.T("Giá trị", "Values"), AutoSize = true },
            new RadioButton { Name = "rbFormulas", Text = Loc.T("Công thức", "Formulas"), AutoSize = true },
            new RadioButton { Name = "rbFormats", Text = Loc.T("Định dạng", "Formats"), AutoSize = true },
            new RadioButton { Name = "rbValuesFormats", Text = Loc.T("Giá trị và định dạng số", "Values and number formats"), AutoSize = true }
        ];
        flow.Controls.Add(new Label { Text = Loc.T("Dán:", "Paste:"), AutoSize = true, Font = Theme.UiFontBold });
        flow.Controls.AddRange(_modes);
        _skip.Text = Loc.T("Bỏ qua ô trống", "Skip blanks");
        _transpose.Text = Loc.T("Chuyển vị (đổi hàng ↔ cột)", "Transpose (rows ↔ columns)");
        flow.Controls.Add(new Label { Height = 8 });
        flow.Controls.Add(_skip);
        flow.Controls.Add(_transpose);
        Body.Controls.Add(flow);
        AddOkCancel(Loc.T("Dán", "Paste"), Loc.T("Hủy", "Cancel"));
    }

    public PasteMode Mode => (PasteMode)Array.FindIndex(_modes, r => r.Checked);
    public bool Transpose { get => _transpose.Checked; set => _transpose.Checked = value; }
    public bool SkipBlanks => _skip.Checked;
    public void SetMode(PasteMode m) => _modes[(int)m].Checked = true;
}

// ====================================================================== Conditional formatting rule

public sealed class CfRuleDialog : DialogBase
{
    private readonly TextBox _v1 = new() { Name = "txtCfValue1", Width = 160 }, _v2 = new() { Name = "txtCfValue2", Width = 160 };
    private readonly ComboBox _preset = new() { Name = "cboCfPreset", DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    public readonly CfType Type;
    public static readonly (int Fill, int Font, string Vi, string En)[] Presets =
    [
        (Color.FromArgb(255, 199, 206).ToArgb(), Color.FromArgb(156, 0, 6).ToArgb(), "Nền đỏ nhạt, chữ đỏ đậm", "Light red fill with dark red text"),
        (Color.FromArgb(255, 235, 156).ToArgb(), Color.FromArgb(156, 87, 0).ToArgb(), "Nền vàng, chữ vàng đậm", "Yellow fill with dark yellow text"),
        (Color.FromArgb(198, 239, 206).ToArgb(), Color.FromArgb(0, 97, 0).ToArgb(), "Nền xanh lá, chữ xanh đậm", "Green fill with dark green text"),
        (Color.FromArgb(255, 199, 206).ToArgb(), 0, "Nền đỏ nhạt", "Light red fill"),
        (0, Color.FromArgb(192, 0, 0).ToArgb(), "Chữ đỏ", "Red text")
    ];

    public CfRuleDialog(CfType type, string suggestion) : base(Loc.T("Định dạng có điều kiện", "Conditional Formatting"), 470, 230)
    {
        Name = "cfRuleDialog";
        Type = type;
        string prompt = type switch
        {
            CfType.GreaterThan => Loc.T("Tô các ô có giá trị LỚN HƠN:", "Format cells that are GREATER THAN:"),
            CfType.LessThan => Loc.T("Tô các ô có giá trị NHỎ HƠN:", "Format cells that are LESS THAN:"),
            CfType.Between => Loc.T("Tô các ô có giá trị NẰM GIỮA:", "Format cells that are BETWEEN:"),
            CfType.EqualTo => Loc.T("Tô các ô có giá trị BẰNG:", "Format cells that are EQUAL TO:"),
            CfType.TextContains => Loc.T("Tô các ô có chứa chữ:", "Format cells that contain the text:"),
            CfType.Top10 => Loc.T("Tô N ô có giá trị LỚN NHẤT:", "Format cells that rank in the TOP:"),
            CfType.Bottom10 => Loc.T("Tô N ô có giá trị NHỎ NHẤT:", "Format cells that rank in the BOTTOM:"),
            CfType.Expression => Loc.T("Công thức xác định ô cần tô (vd. =$B2>100):", "Formula to determine which cells to format (e.g. =$B2>100):"),
            _ => Loc.T("Tô các ô:", "Format cells:")
        };
        foreach (var p in Presets) _preset.Items.Add(Loc.T(p.Vi, p.En));
        _preset.SelectedIndex = 0;
        _v1.Text = type is CfType.Top10 or CfType.Bottom10 ? "10" : suggestion;
        var t = Grid();
        t.Controls.Add(new Label { Text = prompt, AutoSize = true }, 0, 0);
        t.SetColumnSpan(t.GetControlFromPosition(0, 0)!, 2);
        var row = new FlowLayoutPanel { AutoSize = true };
        row.Controls.Add(_v1);
        if (type == CfType.Between) { row.Controls.Add(new Label { Text = Loc.T("và", "and"), AutoSize = true, Margin = new Padding(4, 6, 4, 0) }); row.Controls.Add(_v2); }
        t.Controls.Add(row, 0, 1);
        t.Controls.Add(L(Loc.T("với định dạng:", "with:")), 0, 2);
        t.Controls.Add(_preset, 1, 2);
        Body.Controls.Add(t);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }

    public ConditionalFormat Build(CellRange range)
    {
        var p = Presets[Math.Max(0, _preset.SelectedIndex)];
        var cf = new ConditionalFormat { Range = range, Type = Type, Value1 = _v1.Text.Trim(), Value2 = _v2.Text.Trim(), Fill = p.Fill, FontColor = p.Font };
        if (Type is CfType.Top10 or CfType.Bottom10 && int.TryParse(_v1.Text, out var n)) cf.Rank = Math.Max(1, n);
        return cf;
    }

    public void SetValues(string v1, string v2 = "") { _v1.Text = v1; _v2.Text = v2; }
}

// ====================================================================== Data validation

public sealed class ValidationDialog : DialogBase
{
    private readonly ComboBox _type = new() { Name = "cboDvType", DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
    private readonly ComboBox _op = new() { Name = "cboDvOp", DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
    private readonly TextBox _f1 = new() { Name = "txtDvValue1", Width = 260 }, _f2 = new() { Name = "txtDvValue2", Width = 260 }, _msg = new() { Name = "txtDvMessage", Width = 260 };
    private readonly CheckBox _blank = new() { Name = "chkDvBlank", AutoSize = true, Checked = true };
    public bool ClearAll { get; private set; }

    public ValidationDialog(DataValidation? existing) : base(Loc.T("Xác thực dữ liệu", "Data Validation"), 480, 340)
    {
        Name = "validationDialog";
        _type.Items.AddRange([Loc.T("Mọi giá trị", "Any value"), Loc.T("Danh sách", "List"), Loc.T("Số nguyên", "Whole number"), Loc.T("Số thập phân", "Decimal"), Loc.T("Ngày", "Date"), Loc.T("Độ dài văn bản", "Text length")]);
        _op.Items.AddRange([Loc.T("nằm giữa", "between"), Loc.T("không nằm giữa", "not between"), Loc.T("bằng", "equal to"), Loc.T("khác", "not equal to"), Loc.T("lớn hơn", "greater than"), Loc.T("nhỏ hơn", "less than"), Loc.T("lớn hơn hoặc bằng", "greater than or equal to"), Loc.T("nhỏ hơn hoặc bằng", "less than or equal to")]);
        _type.SelectedIndex = existing != null ? (int)existing.Type : 1;
        _op.SelectedIndex = existing != null ? (int)existing.Op : 0;
        _f1.Text = existing?.Formula1 ?? "";
        _f2.Text = existing?.Formula2 ?? "";
        _msg.Text = existing?.ErrorMessage ?? "";
        _blank.Text = Loc.T("Bỏ qua ô trống", "Ignore blank");
        _blank.Checked = existing?.AllowBlank ?? true;
        var t = Grid();
        t.Controls.Add(L(Loc.T("Cho phép:", "Allow:")), 0, 0); t.Controls.Add(_type, 1, 0);
        t.Controls.Add(L(Loc.T("Điều kiện:", "Data:")), 0, 1); t.Controls.Add(_op, 1, 1);
        t.Controls.Add(L(Loc.T("Nguồn / Giá trị:", "Source / Value:")), 0, 2); t.Controls.Add(_f1, 1, 2);
        t.Controls.Add(L(Loc.T("Đến:", "Maximum:")), 0, 3); t.Controls.Add(_f2, 1, 3);
        t.Controls.Add(L(Loc.T("Thông báo lỗi:", "Error message:")), 0, 4); t.Controls.Add(_msg, 1, 4);
        t.Controls.Add(_blank, 1, 5);
        t.Controls.Add(new Label { AutoSize = true, ForeColor = Theme.SubtleInk, Text = Loc.T("Danh sách: các mục cách nhau bằng dấu phẩy (Có,Không) hoặc vùng =$A$1:$A$10", "List: items separated by commas (Yes,No) or a range =$A$1:$A$10") }, 1, 6);
        Body.Controls.Add(t);
        AddButton(Loc.T("Xóa tất cả", "Clear All"), "btnDvClear", DialogResult.None, false, (_, _) => { ClearAll = true; DialogResult = DialogResult.OK; Close(); });
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
        _type.SelectedIndexChanged += (_, _) => Sync();
        _op.SelectedIndexChanged += (_, _) => Sync();
        Sync();
    }

    private void Sync()
    {
        bool list = _type.SelectedIndex == 1, any = _type.SelectedIndex == 0;
        _op.Enabled = !list && !any;
        _f1.Enabled = !any;
        _f2.Enabled = !list && !any && _op.SelectedIndex <= 1;
    }

    public void Set(ValidationType t, string f1, string f2 = "", ValidationOp op = ValidationOp.Between) { _type.SelectedIndex = (int)t; _op.SelectedIndex = (int)op; _f1.Text = f1; _f2.Text = f2; }

    public DataValidation Build(CellRange r) => new()
    {
        Range = r, Type = (ValidationType)_type.SelectedIndex, Op = (ValidationOp)_op.SelectedIndex, Formula1 = _f1.Text.Trim(), Formula2 = _f2.Text.Trim(),
        AllowBlank = _blank.Checked, ErrorMessage = _msg.Text.Trim()
    };
}

// ====================================================================== Chart

public sealed class ChartDialog : DialogBase
{
    private readonly ListView _types = new() { Name = "lvChartTypes", View = View.LargeIcon, MultiSelect = false, HideSelection = false, Height = Dpi.Scale(110), Dock = DockStyle.Top };
    private readonly TextBox _title = new() { Name = "txtChartTitle", Width = 300 }, _range = new() { Name = "txtChartRange", Width = 200 };
    private readonly CheckBox _legend = new() { Name = "chkLegend", AutoSize = true }, _labels = new() { Name = "chkDataLabels", AutoSize = true }, _rows = new() { Name = "chkSeriesInRows", AutoSize = true };
    private readonly PictureBox _preview = new() { Name = "picChartPreview", Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
    private readonly Sheet _sheet;
    public static readonly (ChartType Type, string Icon, string Vi, string En)[] Types =
    [
        (ChartType.Column, "colChart", "Cột", "Column"), (ChartType.StackedColumn, "stackChart", "Cột chồng", "Stacked"), (ChartType.Bar, "barChart", "Thanh", "Bar"),
        (ChartType.Line, "lineChart", "Đường", "Line"), (ChartType.Area, "areaChart", "Vùng", "Area"), (ChartType.Pie, "pieChart", "Tròn", "Pie"),
        (ChartType.Doughnut, "doughnutChart", "Vành khuyên", "Doughnut"), (ChartType.Scatter, "scatterChart", "Phân tán", "Scatter")
    ];

    public ChartDialog(Sheet sheet, Chart ch) : base(Loc.T("Biểu đồ", "Chart"), 620, 560)
    {
        Name = "chartDialog";
        _sheet = sheet;
        var il = new ImageList { ImageSize = new Size(Dpi.Scale(32), Dpi.Scale(32)), ColorDepth = ColorDepth.Depth32Bit };
        foreach (var t in Types) { il.Images.Add(XlIcons.Get(t.Icon, Dpi.Scale(32))); _types.Items.Add(new ListViewItem(Loc.T(t.Vi, t.En), il.Images.Count - 1) { Tag = t.Type }); }
        _types.LargeImageList = il;
        var idx = Array.FindIndex(Types, t => t.Type == ch.Type);
        _types.Items[Math.Max(0, idx)].Selected = true;
        _title.Text = ch.Title;
        _range.Text = ch.Data.ToString();
        _legend.Text = Loc.T("Chú giải", "Legend"); _legend.Checked = ch.ShowLegend;
        _labels.Text = Loc.T("Nhãn dữ liệu", "Data labels"); _labels.Checked = ch.ShowDataLabels;
        _rows.Text = Loc.T("Chuỗi theo hàng", "Series in rows"); _rows.Checked = ch.SeriesInRows;
        var t2 = Grid();
        t2.Dock = DockStyle.Top; t2.Height = Dpi.Scale(110);
        t2.Controls.Add(L(Loc.T("Tiêu đề:", "Title:")), 0, 0); t2.Controls.Add(_title, 1, 0);
        t2.Controls.Add(L(Loc.T("Vùng dữ liệu:", "Data range:")), 0, 1); t2.Controls.Add(_range, 1, 1);
        var opts = new FlowLayoutPanel { AutoSize = true };
        opts.Controls.AddRange([_legend, _labels, _rows]);
        t2.Controls.Add(opts, 1, 2);
        Body.Controls.Add(_preview);
        Body.Controls.Add(t2);
        Body.Controls.Add(_types);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
        _types.SelectedIndexChanged += (_, _) => UpdatePreview();
        _title.TextChanged += (_, _) => UpdatePreview();
        _range.TextChanged += (_, _) => UpdatePreview();
        foreach (var c in new[] { _legend, _labels, _rows }) c.CheckedChanged += (_, _) => UpdatePreview();
        _preview.Resize += (_, _) => UpdatePreview();
        Shown += (_, _) => UpdatePreview();
    }

    public void SetType(ChartType t) { foreach (ListViewItem it in _types.Items) it.Selected = (ChartType)it.Tag! == t; }
    public void SetTitle(string t) => _title.Text = t;

    public bool Apply(Chart ch)
    {
        if (!Addr.TryParseRange(_range.Text.Contains('!') ? _range.Text[(_range.Text.LastIndexOf('!') + 1)..] : _range.Text, out var r)) return false;
        ch.Data = r;
        ch.Type = _types.SelectedItems.Count > 0 ? (ChartType)_types.SelectedItems[0].Tag! : ChartType.Column;
        ch.Title = _title.Text;
        ch.ShowLegend = _legend.Checked;
        ch.ShowDataLabels = _labels.Checked;
        ch.SeriesInRows = _rows.Checked;
        return true;
    }

    private void UpdatePreview()
    {
        if (_preview.Width < 20 || _preview.Height < 20) return;
        var tmp = new Chart();
        if (!Apply(tmp)) return;
        var bmp = new Bitmap(_preview.Width, _preview.Height);
        using (var g = Graphics.FromImage(bmp)) ChartRenderer.Draw(g, new RectangleF(0, 0, bmp.Width, bmp.Height), _sheet, tmp, Dpi.Factor);
        var old = _preview.Image;
        _preview.Image = bmp;
        old?.Dispose();
    }
}

// ====================================================================== Page setup

public sealed class PageSetupDialog : DialogBase
{
    private readonly RadioButton _portrait = new() { Name = "rbPortrait", AutoSize = true }, _landscape = new() { Name = "rbLandscape", AutoSize = true };
    private readonly ComboBox _paper = new() { Name = "cboPaper", DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
    private readonly NumericUpDown _scale = new() { Name = "numScale", Minimum = 10, Maximum = 400, Value = 100, Width = 70 };
    private readonly CheckBox _fit = new() { Name = "chkFitWidth", AutoSize = true }, _grid = new() { Name = "chkPrintGrid", AutoSize = true }, _head = new() { Name = "chkPrintHeadings", AutoSize = true }, _center = new() { Name = "chkCenter", AutoSize = true };
    private readonly NumericUpDown[] _m = Enumerable.Range(0, 4).Select(i => new NumericUpDown { Minimum = 0, Maximum = 10, DecimalPlaces = 2, Increment = 0.1m, Width = 70, Name = "numMargin" + i }).ToArray();
    private readonly TextBox _area = new() { Name = "txtPrintArea", Width = 160 }, _titles = new() { Name = "txtPrintTitles", Width = 160 };
    public static readonly (int Id, string Name)[] Papers = [(9, "A4"), (11, "A5"), (8, "A3"), (1, "Letter"), (5, "Legal")];

    public PageSetupDialog(PageSetup p) : base(Loc.T("Thiết lập trang", "Page Setup"), 520, 420)
    {
        Name = "pageSetupDialog";
        _portrait.Text = Loc.T("Dọc", "Portrait"); _landscape.Text = Loc.T("Ngang", "Landscape");
        (p.Landscape ? _landscape : _portrait).Checked = true;
        foreach (var x in Papers) _paper.Items.Add(x.Name);
        _paper.SelectedIndex = Math.Max(0, Array.FindIndex(Papers, x => x.Id == p.PaperSize));
        _scale.Value = Math.Clamp(p.Scale, 10, 400);
        _fit.Text = Loc.T("Co vừa 1 trang theo chiều rộng", "Fit all columns on one page"); _fit.Checked = p.FitToWidth;
        _grid.Text = Loc.T("In đường lưới", "Print gridlines"); _grid.Checked = p.PrintGridlines;
        _head.Text = Loc.T("In tiêu đề hàng/cột", "Print headings"); _head.Checked = p.PrintHeadings;
        _center.Text = Loc.T("Căn giữa trang theo chiều ngang", "Center horizontally"); _center.Checked = p.CenterHorizontally;
        var cm = new[] { p.MarginTop, p.MarginBottom, p.MarginLeft, p.MarginRight };
        for (int i = 0; i < 4; i++) _m[i].Value = (decimal)Math.Round(cm[i] * 2.54, 2);
        _area.Text = p.PrintArea?.ToString() ?? "";
        _titles.Text = p.PrintTitleRows is (int a, int b) ? $"${a + 1}:${b + 1}" : "";
        var t = Grid(4);
        t.Controls.Add(L(Loc.T("Hướng:", "Orientation:")), 0, 0);
        var or = new FlowLayoutPanel { AutoSize = true }; or.Controls.AddRange([_portrait, _landscape]);
        t.Controls.Add(or, 1, 0); t.SetColumnSpan(or, 3);
        t.Controls.Add(L(Loc.T("Khổ giấy:", "Paper:")), 0, 1); t.Controls.Add(_paper, 1, 1);
        t.Controls.Add(L(Loc.T("Tỉ lệ %:", "Scale %:")), 2, 1); t.Controls.Add(_scale, 3, 1);
        t.Controls.Add(_fit, 1, 2); t.SetColumnSpan(_fit, 3);
        t.Controls.Add(L(Loc.T("Lề trên (cm):", "Top (cm):")), 0, 3); t.Controls.Add(_m[0], 1, 3);
        t.Controls.Add(L(Loc.T("Lề dưới:", "Bottom:")), 2, 3); t.Controls.Add(_m[1], 3, 3);
        t.Controls.Add(L(Loc.T("Lề trái:", "Left:")), 0, 4); t.Controls.Add(_m[2], 1, 4);
        t.Controls.Add(L(Loc.T("Lề phải:", "Right:")), 2, 4); t.Controls.Add(_m[3], 3, 4);
        t.Controls.Add(L(Loc.T("Vùng in:", "Print area:")), 0, 5); t.Controls.Add(_area, 1, 5);
        t.Controls.Add(L(Loc.T("Hàng lặp lại:", "Rows to repeat:")), 2, 5); t.Controls.Add(_titles, 3, 5);
        t.Controls.Add(_grid, 1, 6); t.SetColumnSpan(_grid, 3);
        t.Controls.Add(_head, 1, 7); t.SetColumnSpan(_head, 3);
        t.Controls.Add(_center, 1, 8); t.SetColumnSpan(_center, 3);
        Body.Controls.Add(t);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }

    public void Apply(PageSetup p)
    {
        p.Landscape = _landscape.Checked;
        p.PaperSize = Papers[Math.Max(0, _paper.SelectedIndex)].Id;
        p.Scale = (int)_scale.Value;
        p.FitToWidth = _fit.Checked;
        p.PrintGridlines = _grid.Checked; p.PrintHeadings = _head.Checked; p.CenterHorizontally = _center.Checked;
        p.MarginTop = (double)_m[0].Value / 2.54; p.MarginBottom = (double)_m[1].Value / 2.54; p.MarginLeft = (double)_m[2].Value / 2.54; p.MarginRight = (double)_m[3].Value / 2.54;
        p.PrintArea = Addr.TryParseRange(_area.Text, out var a) ? a : null;
        p.PrintTitleRows = Addr.TryParseRange(_titles.Text, out var tr) ? (tr.R1, tr.R2) : null;
    }

    public void SetLandscape(bool v) { _landscape.Checked = v; _portrait.Checked = !v; }
}

// ====================================================================== Text to columns / remove duplicates / protect / options

public sealed class TextToColumnsDialog : DialogBase
{
    private readonly RadioButton[] _d;
    private readonly TextBox _other = new() { Name = "txtOtherDelim", Width = 40, MaxLength = 1 };
    private readonly CheckBox _consec = new() { Name = "chkConsecutive", AutoSize = true };

    public TextToColumnsDialog() : base(Loc.T("Chuyển văn bản thành cột", "Convert Text to Columns"), 380, 300)
    {
        Name = "textToColumnsDialog";
        _d =
        [
            new RadioButton { Name = "rbTab", Text = "Tab", AutoSize = true }, new RadioButton { Name = "rbComma", Text = Loc.T("Dấu phẩy (,)", "Comma (,)"), AutoSize = true, Checked = true },
            new RadioButton { Name = "rbSemicolon", Text = Loc.T("Dấu chấm phẩy (;)", "Semicolon (;)"), AutoSize = true }, new RadioButton { Name = "rbSpace", Text = Loc.T("Dấu cách", "Space"), AutoSize = true },
            new RadioButton { Name = "rbOther", Text = Loc.T("Khác:", "Other:"), AutoSize = true }
        ];
        _consec.Text = Loc.T("Coi các dấu phân cách liền nhau là một", "Treat consecutive delimiters as one");
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        flow.Controls.Add(new Label { Text = Loc.T("Dấu phân cách:", "Delimiter:"), AutoSize = true, Font = Theme.UiFontBold });
        flow.Controls.AddRange(_d);
        flow.Controls.Add(_other);
        flow.Controls.Add(_consec);
        Body.Controls.Add(flow);
        AddOkCancel(Loc.T("Hoàn tất", "Finish"), Loc.T("Hủy", "Cancel"));
    }

    public char Delimiter => _d[0].Checked ? '\t' : _d[1].Checked ? ',' : _d[2].Checked ? ';' : _d[3].Checked ? ' ' : _other.Text.Length > 0 ? _other.Text[0] : ',';
    public bool Consecutive => _consec.Checked;
    public void SetDelimiter(int i) => _d[i].Checked = true;
}

public sealed class RemoveDuplicatesDialog : DialogBase
{
    private readonly CheckedListBox _cols = new() { Name = "clbDupColumns", Dock = DockStyle.Fill, CheckOnClick = true };
    private readonly CheckBox _header = new() { Name = "chkDupHeader", AutoSize = true, Dock = DockStyle.Top };
    private readonly CellRange _range;

    public RemoveDuplicatesDialog(Sheet s, CellRange range, bool hasHeader) : base(Loc.T("Loại bỏ trùng lặp", "Remove Duplicates"), 400, 360)
    {
        Name = "removeDuplicatesDialog";
        _range = range;
        _header.Text = Loc.T("Dữ liệu có hàng tiêu đề", "My data has headers");
        _header.Checked = hasHeader;
        for (int c = range.C1; c <= range.C2 && c < range.C1 + 300; c++)
        {
            var h = SheetOps.DisplayText(s, range.R1, c);
            _cols.Items.Add(Loc.T("Cột ", "Column ") + Addr.ColName(c) + (hasHeader && h.Length > 0 ? " — " + h : ""), true);
        }
        Body.Controls.Add(_cols);
        Body.Controls.Add(new Label { Text = Loc.T("Chọn các cột dùng để so sánh trùng lặp:", "Select the columns that contain duplicates:"), Dock = DockStyle.Top, Height = Dpi.Scale(24) });
        Body.Controls.Add(_header);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }

    public bool HasHeader => _header.Checked;
    public List<int> Columns => _cols.CheckedIndices.Cast<int>().Select(i => _range.C1 + i).ToList();
}

public sealed class ProtectDialog : DialogBase
{
    public readonly TextBox Password = new() { Name = "txtProtectPassword", UseSystemPasswordChar = true, Dock = DockStyle.Top };
    public ProtectDialog(bool unprotect) : base(unprotect ? Loc.T("Bỏ bảo vệ trang tính", "Unprotect Sheet") : Loc.T("Bảo vệ trang tính", "Protect Sheet"), 380, 170)
    {
        Name = "protectDialog";
        Body.Controls.Add(Password);
        Body.Controls.Add(new Label
        {
            Dock = DockStyle.Top, Height = Dpi.Scale(44),
            Text = unprotect ? Loc.T("Mật khẩu:", "Password:") : Loc.T("Mật khẩu để bỏ bảo vệ (tùy chọn). Các ô có thuộc tính \"Khóa ô\" sẽ không sửa được.", "Password to unprotect (optional). Cells marked \"Locked\" can't be changed.")
        });
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }

    /// <summary>Legacy Excel sheet-protection password hash (16-bit), as stored in sheetProtection/@password.</summary>
    public static string Hash(string pw)
    {
        if (string.IsNullOrEmpty(pw)) return "";
        int h = 0;
        var bytes = System.Text.Encoding.ASCII.GetBytes(pw);
        for (int i = bytes.Length - 1; i >= 0; i--)
        {
            h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
            h ^= bytes[i];
        }
        h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
        h ^= bytes.Length;
        h ^= 0xCE4B;
        return h.ToString("X4");
    }
}

public sealed class ExcelOptionsDialog : DialogBase
{
    private readonly ComboBox _numbers = new() { Name = "cboNumberStyle", DropDownStyle = ComboBoxStyle.DropDownList, Width = 300 };
    private readonly CheckBox _manual = new() { Name = "chkManualCalc", AutoSize = true };
    private readonly NumericUpDown _autosave = new() { Name = "numAutoRecover", Minimum = 1, Maximum = 60, Width = 60 };

    public ExcelOptionsDialog(string numberStyle, bool manual, int autosaveMinutes) : base(Loc.T("Tùy chọn Excel", "Excel Options"), 480, 260)
    {
        Name = "excelOptionsDialog";
        _numbers.Items.AddRange([Loc.T("Theo ngôn ngữ ứng dụng", "Follow the app language"), Loc.T("Kiểu Việt Nam: 1.234.567,89", "Vietnamese: 1.234.567,89"), Loc.T("Kiểu quốc tế: 1,234,567.89", "International: 1,234,567.89")]);
        _numbers.SelectedIndex = numberStyle switch { "vi" => 1, "en" => 2, _ => 0 };
        _manual.Text = Loc.T("Tính toán thủ công (nhấn F9 để tính)", "Manual calculation (press F9 to calculate)");
        _manual.Checked = manual;
        _autosave.Value = Math.Clamp(autosaveMinutes, 1, 60);
        var t = Grid();
        t.Controls.Add(L(Loc.T("Hiển thị số:", "Number display:")), 0, 0); t.Controls.Add(_numbers, 1, 0);
        t.Controls.Add(_manual, 1, 1);
        t.Controls.Add(L(Loc.T("Tự lưu bản phục hồi mỗi (phút):", "Save AutoRecover every (minutes):")), 0, 2); t.Controls.Add(_autosave, 1, 2);
        t.Controls.Add(new Label { AutoSize = true, ForeColor = Theme.SubtleInk, Text = Loc.T("Công thức luôn dùng dấu phẩy ngăn cách đối số và dấu chấm thập phân, vd. =ROUND(A1*1.1,2).", "Formulas always use commas between arguments and a dot as decimal point, e.g. =ROUND(A1*1.1,2).") }, 1, 3);
        Body.Controls.Add(t);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }

    public string NumberStyle => _numbers.SelectedIndex switch { 1 => "vi", 2 => "en", _ => "auto" };
    public bool ManualCalc => _manual.Checked;
    public int AutoRecoverMinutes => (int)_autosave.Value;
}
