using System.Drawing.Text;
using System.Globalization;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Editor;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.UI;

internal static class DlgUtil
{
    public static NumericUpDown Num(string name, decimal value, decimal min, decimal max, int decimals = 1, decimal inc = 0.5m) => new()
    {
        Name = name, Minimum = min, Maximum = max, DecimalPlaces = decimals, Increment = inc, Value = Math.Clamp(value, min, max), Width = Dpi.Scale(90), Anchor = AnchorStyles.Left
    };

    public static ComboBox Combo(string name, IEnumerable<string> items, int selected, int width = 160)
    {
        var c = new ComboBox { Name = name, DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(width), Anchor = AnchorStyles.Left };
        c.Items.AddRange(items.Cast<object>().ToArray());
        if (c.Items.Count > 0) c.SelectedIndex = Math.Clamp(selected, 0, c.Items.Count - 1);
        return c;
    }

    public static Label L(string text) => new() { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, Dpi.Scale(7), Dpi.Scale(8), Dpi.Scale(3)) };

    public static TableLayoutPanel Grid(int cols = 2)
    {
        var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = cols, AutoScroll = true };
        for (int i = 0; i < cols; i++) t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        return t;
    }

    public static string[] FontNames()
    {
        try
        {
            using var f = new InstalledFontCollection();
            return f.Families.Select(x => x.Name).Where(n => !n.StartsWith('@')).OrderBy(n => n).ToArray();
        }
        catch { return ["Arial", "Calibri", "Times New Roman", "Tahoma", "Verdana"]; }
    }
}

/// <summary>Find / Replace / Go To (modeless, like Word).</summary>
public sealed class FindReplaceDialog : DialogBase
{
    private readonly WordEditor _ed;
    private readonly TabControl _tabs;
    private readonly TextBox _find, _replace, _find2, _page;
    private readonly CheckBox _case, _whole;
    private readonly Label _status;

    public FindReplaceDialog(WordEditor ed, int tab) : base(Loc.T("Tìm và thay thế", "Find and Replace"), 520, 300)
    {
        _ed = ed;
        Name = "findReplaceDialog";
        ShowInTaskbar = false;
        _tabs = new TabControl { Dock = DockStyle.Fill, Name = "tabsFind" };
        var pFind = new TabPage(Loc.T("Tìm", "Find")) { Name = "tabFind" };
        var pRep = new TabPage(Loc.T("Thay thế", "Replace")) { Name = "tabReplace" };
        var pGo = new TabPage(Loc.T("Đi tới", "Go To")) { Name = "tabGoTo" };
        _find = new TextBox { Name = "txtFind", Width = Dpi.Scale(330) };
        _find2 = new TextBox { Name = "txtFind2", Width = Dpi.Scale(330) };
        _replace = new TextBox { Name = "txtReplace", Width = Dpi.Scale(330) };
        _page = new TextBox { Name = "txtGoToPage", Width = Dpi.Scale(120) };
        _find.TextChanged += (_, _) => { if (_find2.Text != _find.Text) _find2.Text = _find.Text; };
        _find2.TextChanged += (_, _) => { if (_find.Text != _find2.Text) _find.Text = _find2.Text; };
        _case = new CheckBox { Name = "chkMatchCase", Text = Loc.T("Phân biệt chữ hoa/thường", "Match case"), AutoSize = true };
        _whole = new CheckBox { Name = "chkWholeWord", Text = Loc.T("Chỉ tìm nguyên từ", "Find whole words only"), AutoSize = true };
        var g1 = DlgUtil.Grid(); g1.Controls.Add(DlgUtil.L(Loc.T("Tìm:", "Find what:"))); g1.Controls.Add(_find);
        pFind.Controls.Add(g1);
        var g2 = DlgUtil.Grid();
        g2.Controls.Add(DlgUtil.L(Loc.T("Tìm:", "Find what:"))); g2.Controls.Add(_find2);
        g2.Controls.Add(DlgUtil.L(Loc.T("Thay bằng:", "Replace with:"))); g2.Controls.Add(_replace);
        pRep.Controls.Add(g2);
        var g3 = DlgUtil.Grid();
        g3.Controls.Add(DlgUtil.L(Loc.F("Số trang (1 - {0}):", "Page number (1 - {0}):", ed.PageCount))); g3.Controls.Add(_page);
        pGo.Controls.Add(g3);
        _tabs.TabPages.AddRange([pFind, pRep, pGo]);
        var opts = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = Dpi.Scale(54), FlowDirection = FlowDirection.TopDown };
        _status = new Label { Name = "lblFindStatus", AutoSize = true, ForeColor = Theme.SubtleInk };
        opts.Controls.AddRange([_case, _whole, _status]);
        Body.Controls.Add(_tabs);
        Body.Controls.Add(opts);
        AddButton(Loc.T("Đóng", "Close"), "btnFindClose", DialogResult.Cancel);
        AddButton(Loc.T("Thay tất cả", "Replace All"), "btnReplaceAll", DialogResult.None, false, (_, _) => ReplaceAll());
        AddButton(Loc.T("Thay thế", "Replace"), "btnReplaceOne", DialogResult.None, false, (_, _) => ReplaceOne());
        var next = AddButton(Loc.T("Tìm tiếp", "Find Next"), "btnFindNext", DialogResult.None, true, (_, _) => { if (_tabs.SelectedTab == pGo) GoTo(); else FindNext(); });
        AcceptButton = next;
        CancelButton = ButtonBar.Controls.OfType<Button>().First(b => b.Name == "btnFindClose");
        _tabs.SelectedIndex = Math.Clamp(tab, 0, 2);
        if (ed.HasSelection && !ed.SelectedText.Contains('\n')) _find.Text = ed.SelectedText;
        Shown += (_, _) => (tab == 2 ? _page : tab == 1 ? _find2 : _find).Focus();
    }

    public int Count => Editing.FindAll(_ed.Document, _find.Text, _case.Checked, _whole.Checked).Count();

    public string FindText { get => _find.Text; set => _find.Text = value; }
    public string ReplaceText { get => _replace.Text; set => _replace.Text = value; }

    public bool FindNext()
    {
        var all = Editing.FindAll(_ed.Document, _find.Text, _case.Checked, _whole.Checked).ToList();
        if (all.Count == 0) { _status.Text = Loc.T("Không tìm thấy.", "No matches found."); return false; }
        var after = _ed.HasSelection ? _ed.SelEnd : _ed.Caret;
        var hit = all.FirstOrDefault(x => x.Start >= after);
        if (hit == default) hit = all[0];
        _ed.SelectRange(hit.Start, hit.End);
        _status.Text = Loc.F("Kết quả {0}/{1}", "Match {0} of {1}", all.IndexOf(hit) + 1, all.Count);
        return true;
    }

    public void ReplaceOne()
    {
        var cmp = _case.Checked ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;
        if (_ed.HasSelection && string.Equals(_ed.SelectedText, _find.Text, cmp))
            _ed.ReplaceRange(_ed.SelStart, _ed.SelEnd, _replace.Text);
        FindNext();
    }

    public int ReplaceAll()
    {
        var all = Editing.FindAll(_ed.Document, _find.Text, _case.Checked, _whole.Checked).ToList();
        if (all.Count == 0) { _status.Text = Loc.T("Không tìm thấy.", "No matches found."); return 0; }
        var rep = _replace.Text;
        _ed.Edit("replace", () =>
        {
            for (int i = all.Count - 1; i >= 0; i--)
            {
                var (a, b) = all[i];
                var props = Editing.CharAt(_ed.Document, new DocPos(a.Para, a.Offset + 1));
                var p = Editing.Delete(_ed.Document, a, b);
                if (rep.Length > 0) Editing.InsertInline(_ed.Document, p, rep, props, null);
            }
        });
        _status.Text = Loc.F("Đã thay {0} chỗ.", "Replaced {0} occurrences.", all.Count);
        return all.Count;
    }

    private void GoTo()
    {
        if (int.TryParse(_page.Text, out var n) && n >= 1 && n <= _ed.PageCount) { _ed.GoToPage(n - 1); _status.Text = ""; }
        else _status.Text = Loc.T("Số trang không hợp lệ.", "Invalid page number.");
    }
}

public sealed class FontDialogEx : DialogBase
{
    private readonly ComboBox _font, _style, _under;
    private readonly NumericUpDown _size;
    private readonly CheckBox _strike, _sup, _sub, _caps;
    private readonly Panel _preview;
    private Color? _color;
    public CharProps Result { get; private set; }

    public FontDialogEx(CharProps cp) : base(Loc.T("Phông chữ", "Font"), 480, 400)
    {
        Name = "fontDialog";
        Result = cp;
        _color = cp.Color;
        var g = DlgUtil.Grid();
        _font = new ComboBox { Name = "cboDlgFont", Width = Dpi.Scale(220), DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems };
        _font.Items.AddRange(DlgUtil.FontNames());
        _font.Text = cp.FontName;
        _style = DlgUtil.Combo("cboDlgStyle", [Loc.T("Thường", "Regular"), Loc.T("Nghiêng", "Italic"), Loc.T("Đậm", "Bold"), Loc.T("Đậm nghiêng", "Bold Italic")], (cp.Bold ? 2 : 0) + (cp.Italic ? 1 : 0));
        _size = DlgUtil.Num("numDlgSize", (decimal)cp.Size, 1, 1638, 1, 0.5m);
        _under = DlgUtil.Combo("cboDlgUnderline", [Loc.T("(không)", "(none)"), Loc.T("Đơn", "Single"), Loc.T("Đôi", "Double"), Loc.T("Chấm", "Dotted"), Loc.T("Đậm", "Thick"), Loc.T("Lượn sóng", "Wave")], (int)cp.Underline);
        var color = new Button { Name = "btnDlgColor", Text = Loc.T("Màu chữ...", "Font colour..."), AutoSize = true };
        color.Click += (_, _) =>
        {
            using var d = new ColorDialog { Color = _color ?? Color.Black, FullOpen = true };
            if (d.ShowDialog(this) == DialogResult.OK) { _color = d.Color; _preview!.Invalidate(); }
        };
        _strike = new CheckBox { Name = "chkDlgStrike", Text = Loc.T("Gạch ngang", "Strikethrough"), Checked = cp.Strike, AutoSize = true };
        _sup = new CheckBox { Name = "chkDlgSuper", Text = Loc.T("Chỉ số trên", "Superscript"), Checked = cp.VertAlign == VertAlign.Superscript, AutoSize = true };
        _sub = new CheckBox { Name = "chkDlgSub", Text = Loc.T("Chỉ số dưới", "Subscript"), Checked = cp.VertAlign == VertAlign.Subscript, AutoSize = true };
        _caps = new CheckBox { Name = "chkDlgCaps", Text = Loc.T("Viết hoa tất cả", "All caps"), Checked = cp.AllCaps, AutoSize = true };
        _sup.CheckedChanged += (_, _) => { if (_sup.Checked) _sub.Checked = false; _preview!.Invalidate(); };
        _sub.CheckedChanged += (_, _) => { if (_sub.Checked) _sup.Checked = false; _preview!.Invalidate(); };
        g.Controls.Add(DlgUtil.L(Loc.T("Phông chữ:", "Font:"))); g.Controls.Add(_font);
        g.Controls.Add(DlgUtil.L(Loc.T("Kiểu:", "Style:"))); g.Controls.Add(_style);
        g.Controls.Add(DlgUtil.L(Loc.T("Cỡ:", "Size:"))); g.Controls.Add(_size);
        g.Controls.Add(DlgUtil.L(Loc.T("Gạch chân:", "Underline:"))); g.Controls.Add(_under);
        g.Controls.Add(DlgUtil.L(Loc.T("Màu:", "Colour:"))); g.Controls.Add(color);
        var fx = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
        fx.Controls.AddRange([_strike, _sup, _sub, _caps]);
        g.Controls.Add(DlgUtil.L(Loc.T("Hiệu ứng:", "Effects:"))); g.Controls.Add(fx);
        _preview = new Panel { Dock = DockStyle.Bottom, Height = Dpi.Scale(70), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Name = "pnlFontPreview" };
        _preview.Paint += (_, e) =>
        {
            var r = Build();
            using var f = new Font(r.FontName, Math.Min(r.Size, 36), r.FontStyle, GraphicsUnit.Point);
            TextRenderer.DrawText(e.Graphics, r.AllCaps ? "MICROADS OFFICE" : "Microads Office", f, _preview.ClientRectangle, r.Color ?? Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        };
        foreach (Control c in new Control[] { _font, _style, _size, _under, _strike, _caps }) c.TextChanged += (_, _) => _preview.Invalidate();
        _style.SelectedIndexChanged += (_, _) => _preview.Invalidate();
        _size.ValueChanged += (_, _) => _preview.Invalidate();
        Body.Controls.Add(g);
        Body.Controls.Add(_preview);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"), (_, _) => { Result = Build(); DialogResult = DialogResult.OK; Close(); });
    }

    private CharProps Build() => Result with
    {
        FontName = string.IsNullOrWhiteSpace(_font.Text) ? Result.FontName : _font.Text.Trim(),
        Bold = _style.SelectedIndex >= 2, Italic = _style.SelectedIndex % 2 == 1, Size = (float)_size.Value,
        Underline = (UnderlineStyle)Math.Max(0, _under.SelectedIndex), Strike = _strike.Checked, AllCaps = _caps.Checked,
        VertAlign = _sup.Checked ? VertAlign.Superscript : _sub.Checked ? VertAlign.Subscript : VertAlign.Baseline, Color = _color
    };
}

public sealed class ParagraphDialog : DialogBase
{
    private readonly ComboBox _align, _special, _rule;
    private readonly NumericUpDown _left, _right, _by, _before, _after, _at;
    private readonly CheckBox _pageBreak, _keep;
    public ParaProps Result { get; private set; }

    public ParagraphDialog(ParaProps pp) : base(Loc.T("Đoạn văn", "Paragraph"), 470, 420)
    {
        Name = "paragraphDialog";
        Result = pp;
        var g = DlgUtil.Grid(4);
        _align = DlgUtil.Combo("cboParaAlign", [Loc.T("Trái", "Left"), Loc.T("Giữa", "Centered"), Loc.T("Phải", "Right"), Loc.T("Đều hai bên", "Justified")], (int)pp.Align, 120);
        _left = DlgUtil.Num("numIndLeft", (decimal)PageSetup.ToMm(pp.IndentLeft) / 10, -10, 30, 2, 0.25m);
        _right = DlgUtil.Num("numIndRight", (decimal)PageSetup.ToMm(pp.IndentRight) / 10, -10, 30, 2, 0.25m);
        _special = DlgUtil.Combo("cboIndSpecial", [Loc.T("(không)", "(none)"), Loc.T("Dòng đầu", "First line"), Loc.T("Treo", "Hanging")], pp.FirstLine > 0 ? 1 : pp.FirstLine < 0 ? 2 : 0, 120);
        _by = DlgUtil.Num("numIndBy", (decimal)Math.Abs(PageSetup.ToMm(pp.FirstLine)) / 10, 0, 30, 2, 0.25m);
        _before = DlgUtil.Num("numSpBefore", (decimal)pp.SpaceBefore, 0, 1584, 0, 6);
        _after = DlgUtil.Num("numSpAfter", (decimal)pp.SpaceAfter, 0, 1584, 0, 6);
        int ruleIdx = pp.LineRule == LineRule.Exact ? 4 : pp.LineRule == LineRule.AtLeast ? 3 : Math.Abs(pp.LineSpacing - 1) < 0.01 ? 0 : Math.Abs(pp.LineSpacing - 1.5) < 0.01 ? 1 : Math.Abs(pp.LineSpacing - 2) < 0.01 ? 2 : 5;
        _rule = DlgUtil.Combo("cboLineRule", [Loc.T("Đơn", "Single"), "1.5", Loc.T("Đôi", "Double"), Loc.T("Tối thiểu", "At least"), Loc.T("Chính xác", "Exactly"), Loc.T("Bội số", "Multiple")], ruleIdx, 120);
        _at = DlgUtil.Num("numLineAt", (decimal)pp.LineSpacing, 0, 1584, 2, 0.05m);
        _pageBreak = new CheckBox { Name = "chkPageBreakBefore", Text = Loc.T("Ngắt trang trước đoạn", "Page break before"), Checked = pp.PageBreakBefore, AutoSize = true };
        _keep = new CheckBox { Name = "chkKeepNext", Text = Loc.T("Giữ với đoạn sau", "Keep with next"), Checked = pp.KeepWithNext, AutoSize = true };
        g.Controls.Add(DlgUtil.L(Loc.T("Căn lề:", "Alignment:"))); g.Controls.Add(_align); g.Controls.Add(new Label()); g.Controls.Add(new Label());
        g.Controls.Add(DlgUtil.L(Loc.T("Thụt trái (cm):", "Left (cm):"))); g.Controls.Add(_left);
        g.Controls.Add(DlgUtil.L(Loc.T("Thụt phải (cm):", "Right (cm):"))); g.Controls.Add(_right);
        g.Controls.Add(DlgUtil.L(Loc.T("Đặc biệt:", "Special:"))); g.Controls.Add(_special);
        g.Controls.Add(DlgUtil.L(Loc.T("Bằng (cm):", "By (cm):"))); g.Controls.Add(_by);
        g.Controls.Add(DlgUtil.L(Loc.T("Trước (pt):", "Before (pt):"))); g.Controls.Add(_before);
        g.Controls.Add(DlgUtil.L(Loc.T("Sau (pt):", "After (pt):"))); g.Controls.Add(_after);
        g.Controls.Add(DlgUtil.L(Loc.T("Giãn dòng:", "Line spacing:"))); g.Controls.Add(_rule);
        g.Controls.Add(DlgUtil.L(Loc.T("Giá trị:", "At:"))); g.Controls.Add(_at);
        g.Controls.Add(_pageBreak); g.SetColumnSpan(_pageBreak, 2);
        g.Controls.Add(_keep); g.SetColumnSpan(_keep, 2);
        _rule.SelectedIndexChanged += (_, _) =>
        {
            _at.Value = _rule.SelectedIndex switch { 0 => 1, 1 => 1.5m, 2 => 2, 3 or 4 => Math.Max(12, _at.Value), _ => _at.Value };
        };
        Body.Controls.Add(g);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"), (_, _) => { Result = Build(); DialogResult = DialogResult.OK; Close(); });
    }

    private ParaProps Build()
    {
        float cm(NumericUpDown n) => PageSetup.Mm((float)n.Value * 10);
        float first = _special.SelectedIndex == 1 ? cm(_by) : _special.SelectedIndex == 2 ? -cm(_by) : 0;
        var rule = _rule.SelectedIndex == 3 ? LineRule.AtLeast : _rule.SelectedIndex == 4 ? LineRule.Exact : LineRule.Multiple;
        return Result with
        {
            Align = (ParaAlign)_align.SelectedIndex, IndentLeft = cm(_left), IndentRight = cm(_right), FirstLine = first, SpaceBefore = (float)_before.Value,
            SpaceAfter = (float)_after.Value, LineRule = rule, LineSpacing = Math.Max(rule == LineRule.Multiple ? 0.5f : 1f, (float)_at.Value), PageBreakBefore = _pageBreak.Checked, KeepWithNext = _keep.Checked
        };
    }
}

public sealed class InsertTableDialog : DialogBase
{
    public int Rows => (int)_rows.Value;
    public int Cols => (int)_cols.Value;
    private readonly NumericUpDown _rows, _cols;

    public InsertTableDialog() : base(Loc.T("Chèn bảng", "Insert Table"), 300, 170)
    {
        Name = "insertTableDialog";
        var g = DlgUtil.Grid();
        _cols = DlgUtil.Num("numTableCols", 3, 1, 63, 0, 1);
        _rows = DlgUtil.Num("numTableRows", 3, 1, 500, 0, 1);
        g.Controls.Add(DlgUtil.L(Loc.T("Số cột:", "Number of columns:"))); g.Controls.Add(_cols);
        g.Controls.Add(DlgUtil.L(Loc.T("Số hàng:", "Number of rows:"))); g.Controls.Add(_rows);
        Body.Controls.Add(g);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }
}

public sealed class PageSetupDialogEx : DialogBase
{
    private readonly NumericUpDown _top, _bottom, _left, _right, _header, _footer, _w, _h;
    private readonly ComboBox _paper, _orient;
    public PageSetup Result { get; }

    public PageSetupDialogEx(PageSetup ps) : base(Loc.T("Thiết lập trang", "Page Setup"), 460, 400)
    {
        Name = "pageSetupDialog";
        Result = ps.Clone();
        var g = DlgUtil.Grid(4);
        NumericUpDown Mm(string n, float pt) => DlgUtil.Num(n, (decimal)Math.Round(PageSetup.ToMm(pt), 1), 0, 600, 1, 1);
        _top = Mm("numMarTop", ps.MarginTop); _bottom = Mm("numMarBottom", ps.MarginBottom); _left = Mm("numMarLeft", ps.MarginLeft); _right = Mm("numMarRight", ps.MarginRight);
        _header = Mm("numHeaderDist", ps.HeaderDistance); _footer = Mm("numFooterDist", ps.FooterDistance);
        _w = Mm("numPaperW", ps.Width); _h = Mm("numPaperH", ps.Height);
        var names = PageSetup.Papers.Select(p => p.Name).Append(Loc.T("Tùy chỉnh", "Custom")).ToArray();
        int pi = Array.FindIndex(PageSetup.Papers, p => p.Name == ps.PaperName);
        _paper = DlgUtil.Combo("cboPaper", names, pi < 0 ? names.Length - 1 : pi, 120);
        _orient = DlgUtil.Combo("cboOrientation", [Loc.T("Dọc", "Portrait"), Loc.T("Ngang", "Landscape")], ps.Landscape ? 1 : 0, 120);
        _paper.SelectedIndexChanged += (_, _) => ApplyPaper();
        _orient.SelectedIndexChanged += (_, _) => ApplyPaper();
        g.Controls.Add(DlgUtil.L(Loc.T("Lề trên (mm):", "Top (mm):"))); g.Controls.Add(_top);
        g.Controls.Add(DlgUtil.L(Loc.T("Lề dưới (mm):", "Bottom (mm):"))); g.Controls.Add(_bottom);
        g.Controls.Add(DlgUtil.L(Loc.T("Lề trái (mm):", "Left (mm):"))); g.Controls.Add(_left);
        g.Controls.Add(DlgUtil.L(Loc.T("Lề phải (mm):", "Right (mm):"))); g.Controls.Add(_right);
        g.Controls.Add(DlgUtil.L(Loc.T("Đầu trang (mm):", "Header (mm):"))); g.Controls.Add(_header);
        g.Controls.Add(DlgUtil.L(Loc.T("Chân trang (mm):", "Footer (mm):"))); g.Controls.Add(_footer);
        g.Controls.Add(DlgUtil.L(Loc.T("Khổ giấy:", "Paper:"))); g.Controls.Add(_paper);
        g.Controls.Add(DlgUtil.L(Loc.T("Hướng:", "Orientation:"))); g.Controls.Add(_orient);
        g.Controls.Add(DlgUtil.L(Loc.T("Rộng (mm):", "Width (mm):"))); g.Controls.Add(_w);
        g.Controls.Add(DlgUtil.L(Loc.T("Cao (mm):", "Height (mm):"))); g.Controls.Add(_h);
        var preset = new Button { Name = "btnMarginVN", Text = Loc.T("Lề văn bản hành chính VN", "Vietnamese admin margins"), AutoSize = true };
        preset.Click += (_, _) => { _top.Value = 20; _bottom.Value = 20; _left.Value = 30; _right.Value = 15; };
        g.Controls.Add(preset); g.SetColumnSpan(preset, 4);
        Body.Controls.Add(g);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            Result.MarginTop = PageSetup.Mm((float)_top.Value); Result.MarginBottom = PageSetup.Mm((float)_bottom.Value);
            Result.MarginLeft = PageSetup.Mm((float)_left.Value); Result.MarginRight = PageSetup.Mm((float)_right.Value);
            Result.HeaderDistance = PageSetup.Mm((float)_header.Value); Result.FooterDistance = PageSetup.Mm((float)_footer.Value);
            Result.Width = PageSetup.Mm((float)_w.Value); Result.Height = PageSetup.Mm((float)_h.Value);
            Result.PaperName = _paper.SelectedIndex < PageSetup.Papers.Length ? PageSetup.Papers[_paper.SelectedIndex].Name : "Custom";
            if (Result.ContentWidth < 72 || Result.ContentHeight < 72) { Msg.Warn(this, Loc.T("Lề quá lớn so với khổ giấy.", "The margins are too large for this paper size.")); return; }
            DialogResult = DialogResult.OK; Close();
        });
    }

    private void ApplyPaper()
    {
        if (_paper.SelectedIndex < 0 || _paper.SelectedIndex >= PageSetup.Papers.Length) return;
        var p = PageSetup.Papers[_paper.SelectedIndex];
        bool land = _orient.SelectedIndex == 1;
        _w.Value = (decimal)Math.Round(PageSetup.ToMm(land ? p.H : p.W), 1);
        _h.Value = (decimal)Math.Round(PageSetup.ToMm(land ? p.W : p.H), 1);
    }
}

public sealed class HyperlinkDialog : DialogBase
{
    public string LinkText => _text.Text;
    public string Url => NormalizeUrl(_url.Text.Trim());
    private readonly TextBox _text, _url;

    public HyperlinkDialog(string text, string? url) : base(Loc.T("Chèn liên kết", "Insert Hyperlink"), 460, 180)
    {
        Name = "hyperlinkDialog";
        var g = DlgUtil.Grid();
        _text = new TextBox { Name = "txtLinkText", Text = text, Width = Dpi.Scale(320) };
        _url = new TextBox { Name = "txtLinkUrl", Text = url ?? "https://", Width = Dpi.Scale(320) };
        g.Controls.Add(DlgUtil.L(Loc.T("Chữ hiển thị:", "Text to display:"))); g.Controls.Add(_text);
        g.Controls.Add(DlgUtil.L(Loc.T("Địa chỉ:", "Address:"))); g.Controls.Add(_url);
        Body.Controls.Add(g);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
        Shown += (_, _) => { _url.Focus(); _url.SelectAll(); };
    }

    public static string NormalizeUrl(string u)
    {
        if (u.Length == 0) return u;
        if (u.Contains('@') && !u.Contains("://") && !u.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return "mailto:" + u;
        if (!u.Contains("://") && !u.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) return "https://" + u;
        return u;
    }
}

public sealed class HeaderFooterDialog : DialogBase
{
    private readonly TextBox _text;
    private readonly ComboBox _align;
    private readonly CheckBox _page;
    public string Text2 => _text.Text;
    public ParaAlign Align => (ParaAlign)_align.SelectedIndex;
    public bool PageNumber => _page.Checked;

    public HeaderFooterDialog(bool header, string current, ParaAlign align, bool hasPage) : base(header ? Loc.T("Đầu trang", "Header") : Loc.T("Chân trang", "Footer"), 480, 300)
    {
        Name = header ? "headerDialog" : "footerDialog";
        _text = new TextBox { Name = "txtHeaderFooter", Multiline = true, Dock = DockStyle.Fill, Text = current.Replace("\n", Environment.NewLine), Font = new Font("Times New Roman", 11f), ScrollBars = ScrollBars.Vertical };
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = Dpi.Scale(40) };
        _align = DlgUtil.Combo("cboHFAlign", [Loc.T("Trái", "Left"), Loc.T("Giữa", "Center"), Loc.T("Phải", "Right")], Math.Min(2, (int)align), 100);
        _page = new CheckBox { Name = "chkHFPageNumber", Text = Loc.T("Thêm số trang", "Add page number"), Checked = hasPage, AutoSize = true, Margin = new Padding(Dpi.Scale(12), Dpi.Scale(6), 0, 0) };
        bottom.Controls.AddRange([DlgUtil.L(Loc.T("Căn lề:", "Alignment:")), _align, _page]);
        Body.Controls.Add(_text);
        Body.Controls.Add(bottom);
        var remove = AddButton(Loc.T("Xóa", "Remove"), "btnHFRemove", DialogResult.None, false, (_, _) => { _text.Text = ""; _page.Checked = false; DialogResult = DialogResult.OK; Close(); });
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }
}

public sealed class SymbolDialog : DialogBase
{
    public string? Selected { get; private set; }
    public static readonly string Symbols = "©®™§¶†‡•…‰′″€£¥₫$¢°±×÷≠≈≤≥∞√∑∏∫∂∆∇πΩαβγδεθλμσφψω←↑→↓↔⇒⇔✓✔✗✘★☆●○■□▲△▼▽◆◇♠♣♥♦☎✉☺☻♪♫¼½¾¹²³ªº«»¿¡ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝàáâãèéêìíòóôõùúýĂăĐđĨĩŨũƠơƯưẠạẢảẤấẦầẨẩẪẫẬậẮắẰằẲẳẴẵẶặẸẹẺẻẼẽẾếỀềỂểỄễỆệỈỉỊịỌọỎỏỐốỒồỔổỖỗỘộỚớỜờỞởỠỡỢợỤụỦủỨứỪừỬửỮữỰựỲỳỴỵỶỷỸỹ";

    public SymbolDialog() : base(Loc.T("Ký hiệu", "Symbol"), 520, 400)
    {
        Name = "symbolDialog";
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, Name = "symbolGrid" };
        int i = 0;
        foreach (var ch in Symbols)
        {
            var s = ch.ToString();
            var b = new Button { Name = "sym" + i++, Text = s, Size = new Size(Dpi.Scale(34), Dpi.Scale(34)), Font = new Font("Segoe UI Symbol", 13f), FlatStyle = FlatStyle.Flat, Margin = new Padding(1) };
            b.FlatAppearance.BorderColor = Theme.Border;
            b.Click += (_, _) => { Selected = s; DialogResult = DialogResult.OK; Close(); };
            flow.Controls.Add(b);
        }
        Body.Controls.Add(flow);
        AddButton(Loc.T("Đóng", "Close"), "btnSymClose", DialogResult.Cancel);
    }
}

public sealed class DateTimeDialog : DialogBase
{
    public string? Selected => _list.SelectedItem as string;
    private readonly ListBox _list;

    public DateTimeDialog() : base(Loc.T("Ngày và giờ", "Date and Time"), 360, 340)
    {
        Name = "dateTimeDialog";
        var now = DateTime.Now;
        var vi = new CultureInfo("vi-VN");
        _list = new ListBox { Dock = DockStyle.Fill, Name = "lstDateFormats", Font = new Font("Segoe UI", 10f) };
        _list.Items.AddRange(new object[]
        {
            now.ToString("dd/MM/yyyy"), now.ToString("dd/MM/yyyy HH:mm"), $"Ngày {now:dd} tháng {now:MM} năm {now:yyyy}", now.ToString("dddd, dd/MM/yyyy", vi),
            now.ToString("d MMMM yyyy", CultureInfo.InvariantCulture), now.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture), now.ToString("yyyy-MM-dd"), now.ToString("HH:mm"), now.ToString("HH:mm:ss")
        });
        _list.SelectedIndex = 0;
        _list.DoubleClick += (_, _) => { DialogResult = DialogResult.OK; Close(); };
        Body.Controls.Add(_list);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }
}

public sealed class WordCountDialog : DialogBase
{
    public WordCountDialog(Document d, DocPos? a, DocPos? b, int pages) : base(Loc.T("Đếm từ", "Word Count"), 340, 250)
    {
        Name = "wordCountDialog";
        var (w, c, cn, p) = Editing.Count(d, a, b);
        var g = DlgUtil.Grid();
        void Row(string label, int v, string name) { g.Controls.Add(DlgUtil.L(label)); g.Controls.Add(new Label { Name = name, Text = v.ToString("N0"), AutoSize = true, Font = Theme.UiFontBold, Margin = new Padding(0, Dpi.Scale(7), 0, 0) }); }
        Row(Loc.T("Trang", "Pages"), pages, "lblWcPages");
        Row(Loc.T("Từ", "Words"), w, "lblWcWords");
        Row(Loc.T("Ký tự (không tính khoảng trắng)", "Characters (no spaces)"), cn, "lblWcChars");
        Row(Loc.T("Ký tự (có khoảng trắng)", "Characters (with spaces)"), c, "lblWcCharsSpaces");
        Row(Loc.T("Đoạn văn", "Paragraphs"), p, "lblWcParas");
        Body.Controls.Add(g);
        AddButton(Loc.T("Đóng", "Close"), "btnWcClose", DialogResult.OK, true);
    }
}

public sealed class ZoomDialog : DialogBase
{
    public float Value => (float)_n.Value / 100f;
    private readonly NumericUpDown _n;

    public ZoomDialog(float current) : base(Loc.T("Thu phóng", "Zoom"), 300, 230)
    {
        Name = "zoomDialog";
        _n = DlgUtil.Num("numZoom", (decimal)Math.Round(current * 100), 10, 500, 0, 10);
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
        foreach (var z in new[] { 200, 100, 75 })
        {
            var r = new RadioButton { Text = z + "%", AutoSize = true, Name = "rbZoom" + z, Checked = Math.Abs(current * 100 - z) < 1 };
            r.CheckedChanged += (_, _) => { if (r.Checked) _n.Value = z; };
            flow.Controls.Add(r);
        }
        var row = new FlowLayoutPanel { AutoSize = true };
        row.Controls.Add(DlgUtil.L(Loc.T("Phần trăm:", "Percent:")));
        row.Controls.Add(_n);
        flow.Controls.Add(row);
        Body.Controls.Add(flow);
        AddOkCancel(Loc.T("OK", "OK"), Loc.T("Hủy", "Cancel"));
    }
}

/// <summary>Spelling review: walks the misspelled words of the document with suggestions.</summary>
public sealed class SpellingDialog : DialogBase
{
    private readonly WordEditor _ed;
    private readonly SpellService _sp;
    private readonly Label _word;
    private readonly ListBox _sugg;
    private (DocPos A, DocPos B)? _cur;

    public SpellingDialog(WordEditor ed, SpellService sp) : base(Loc.T("Chính tả", "Spelling"), 420, 330)
    {
        Name = "spellingDialog";
        _ed = ed; _sp = sp;
        _word = new Label { Name = "lblSpellWord", Dock = DockStyle.Top, Height = Dpi.Scale(34), Font = new Font("Segoe UI Semibold", 12f), ForeColor = Theme.ErrRed };
        _sugg = new ListBox { Name = "lstSuggestions", Dock = DockStyle.Fill };
        var lang = new Label { Dock = DockStyle.Bottom, Height = Dpi.Scale(24), ForeColor = Theme.SubtleInk, Text = Loc.T("Từ điển: ", "Dictionary: ") + (sp.Available ? sp.Language : Loc.T("không có", "not available")) };
        Body.Controls.Add(_sugg);
        Body.Controls.Add(_word);
        Body.Controls.Add(lang);
        AddButton(Loc.T("Đóng", "Close"), "btnSpellClose", DialogResult.Cancel);
        AddButton(Loc.T("Thêm vào từ điển", "Add to dictionary"), "btnSpellAdd", DialogResult.None, false, (_, _) => { if (_cur != null) { _sp.AddToDictionary(Word); Next(); } });
        AddButton(Loc.T("Bỏ qua", "Ignore"), "btnSpellIgnore", DialogResult.None, false, (_, _) => { if (_cur != null) { _sp.Ignore(Word); Next(); } });
        AddButton(Loc.T("Thay đổi", "Change"), "btnSpellChange", DialogResult.None, true, (_, _) =>
        {
            if (_cur is { } c && _sugg.SelectedItem is string s) { _ed.ReplaceRange(c.A, c.B, s); _ed.SetCaret(new DocPos(c.A.Para, c.A.Offset + s.Length), false); }
            Next();
        });
        Shown += (_, _) => Next();
    }

    private string Word => _cur is { } c ? Editing.GetText(_ed.Document, c.A, c.B) : "";

    private void Next()
    {
        var ps = _ed.Document.Paragraphs;
        var from = _ed.HasSelection ? _ed.SelEnd : _ed.Caret;
        for (int i = from.Para; i < ps.Count; i++)
        {
            foreach (var (s, l) in _sp.Errors(ps[i]))
            {
                if (i == from.Para && s < from.Offset) continue;
                _cur = (new DocPos(i, s), new DocPos(i, s + l));
                _ed.SelectRange(_cur.Value.A, _cur.Value.B);
                _word.Text = Word;
                _sugg.Items.Clear();
                foreach (var x in _sp.Suggest(Word)) _sugg.Items.Add(x);
                if (_sugg.Items.Count > 0) _sugg.SelectedIndex = 0;
                return;
            }
        }
        _cur = null;
        _word.ForeColor = Theme.OkGreen;
        _word.Text = Loc.T("Đã kiểm tra xong chính tả. ✓", "Spelling check complete. ✓");
        _sugg.Items.Clear();
    }
}
