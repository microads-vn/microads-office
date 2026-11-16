using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>AutoFilter drop-down: sort, value checklist with search, custom condition.</summary>
public sealed class FilterPopup : ToolStripDropDown
{
    private readonly CheckedListBox _list = new() { Name = "clbFilterValues", CheckOnClick = true, IntegralHeight = false, BorderStyle = BorderStyle.FixedSingle };
    private readonly TextBox _search = new() { Name = "txtFilterSearch" };
    private readonly TextBox _cond = new() { Name = "txtFilterCondition" };
    private readonly List<string> _values;
    private readonly HashSet<string> _checked;
    private bool _syncing;

    public event Action<bool>? SortRequested;
    /// <summary>(allowed values or null = all, condition or null).</summary>
    public event Action<HashSet<string>?, string?>? Applied;

    public FilterPopup(List<string> values, HashSet<string>? allowed, string? condition)
    {
        Name = "filterPopup";
        _values = values;
        _checked = allowed != null ? new HashSet<string>(allowed) : new HashSet<string>(values);
        var panel = new Panel { Size = new Size(Dpi.Scale(260), Dpi.Scale(390)), BackColor = Color.White };
        int y = Dpi.Scale(6), w = Dpi.Scale(248), x = Dpi.Scale(6);
        Button Link(string name, string text, Image icon, Action a)
        {
            var b = new Button { Name = name, Text = "  " + text, Image = icon, ImageAlign = ContentAlignment.MiddleLeft, TextImageRelation = TextImageRelation.ImageBeforeText, TextAlign = ContentAlignment.MiddleLeft, FlatStyle = FlatStyle.Flat, Bounds = new Rectangle(x, y, w, Dpi.Scale(26)) };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Theme.Hover;
            b.Click += (_, _) => { Close(); a(); };
            panel.Controls.Add(b);
            y += Dpi.Scale(28);
            return b;
        }
        Link("btnFilterSortAsc", Loc.T("Sắp xếp A → Z", "Sort A to Z"), XlIcons.Small("sortAsc"), () => SortRequested?.Invoke(true));
        Link("btnFilterSortDesc", Loc.T("Sắp xếp Z → A", "Sort Z to A"), XlIcons.Small("sortDesc"), () => SortRequested?.Invoke(false));
        Link("btnFilterClear", Loc.T("Xóa bộ lọc của cột", "Clear filter from column"), XlIcons.Small("clearFilter"), () => Applied?.Invoke(null, null));
        y += Dpi.Scale(4);
        panel.Controls.Add(new Label { Text = Loc.T("Lọc theo điều kiện (vd. >100, <=5, contains:abc):", "Condition (e.g. >100, <=5, contains:abc):"), Bounds = new Rectangle(x, y, w, Dpi.Scale(30)), ForeColor = Theme.SubtleInk, Font = Theme.SmallFont });
        y += Dpi.Scale(30);
        _cond.Bounds = new Rectangle(x, y, w, Dpi.Scale(24));
        _cond.Text = condition ?? "";
        panel.Controls.Add(_cond);
        y += Dpi.Scale(30);
        _search.Bounds = new Rectangle(x, y, w, Dpi.Scale(24));
        _search.PlaceholderText = Loc.T("Tìm kiếm", "Search");
        panel.Controls.Add(_search);
        y += Dpi.Scale(28);
        _list.Bounds = new Rectangle(x, y, w, Dpi.Scale(150));
        panel.Controls.Add(_list);
        y += Dpi.Scale(156);
        var ok = new Button { Name = "btnFilterOk", Text = Loc.T("OK", "OK"), Bounds = new Rectangle(x + w - Dpi.Scale(170), y, Dpi.Scale(80), Dpi.Scale(28)), BackColor = Theme.Accent, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        ok.FlatAppearance.BorderSize = 0;
        var cancel = new Button { Name = "btnFilterCancel", Text = Loc.T("Hủy", "Cancel"), Bounds = new Rectangle(x + w - Dpi.Scale(84), y, Dpi.Scale(80), Dpi.Scale(28)), FlatStyle = FlatStyle.Flat };
        ok.Click += (_, _) => { Close(); Apply(); };
        cancel.Click += (_, _) => Close();
        panel.Controls.AddRange([ok, cancel]);
        panel.Height = y + Dpi.Scale(36);
        _search.TextChanged += (_, _) => Fill();
        _list.ItemCheck += (_, e) =>
        {
            if (_syncing) return;
            BeginInvoke(() =>
            {
                if (e.Index == 0)
                {
                    bool all = _list.GetItemChecked(0);
                    foreach (var v in Visible()) if (all) _checked.Add(v); else _checked.Remove(v);
                    Fill();
                    return;
                }
                var val = (string)_list.Items[e.Index];
                var key = val == Blank ? "" : val;
                if (_list.GetItemChecked(e.Index)) _checked.Add(key); else _checked.Remove(key);
                _syncing = true;
                _list.SetItemChecked(0, Visible().All(_checked.Contains));
                _syncing = false;
            });
        };
        Fill();
        var host = new ToolStripControlHost(panel) { AutoSize = false, Size = panel.Size, Margin = Padding.Empty, Padding = Padding.Empty };
        Items.Add(host);
        Padding = new Padding(1);
    }

    private static string Blank => Loc.T("(Trống)", "(Blanks)");

    private IEnumerable<string> Visible() => _values.Where(v => _search.Text.Length == 0 || v.Contains(_search.Text, StringComparison.CurrentCultureIgnoreCase));

    private void Fill()
    {
        _syncing = true;
        _list.BeginUpdate();
        _list.Items.Clear();
        var vis = Visible().ToList();
        _list.Items.Add(Loc.T("(Chọn tất cả)", "(Select All)"), vis.All(_checked.Contains));
        foreach (var v in vis) _list.Items.Add(v.Length == 0 ? Blank : v, _checked.Contains(v));
        _list.EndUpdate();
        _syncing = false;
    }

    public void Apply()
    {
        var cond = _cond.Text.Trim();
        var all = _values.All(_checked.Contains);
        Applied?.Invoke(all ? null : new HashSet<string>(_checked), cond.Length > 0 ? cond : null);
    }

    /// <summary>Test helper: keep only these values.</summary>
    public void CheckOnly(params string[] values) { _checked.Clear(); foreach (var v in values) _checked.Add(v); Fill(); }
    public void SetCondition(string c) => _cond.Text = c;
}

/// <summary>Find and Replace (modeless).</summary>
public sealed class FindReplaceDialog : DialogBase
{
    private readonly TextBox _find = new() { Name = "txtFind", Width = 300 }, _replace = new() { Name = "txtReplace", Width = 300 };
    private readonly CheckBox _case = new() { Name = "chkMatchCase", AutoSize = true }, _whole = new() { Name = "chkWholeCell", AutoSize = true }, _book = new() { Name = "chkWholeWorkbook", AutoSize = true }, _formulas = new() { Name = "chkInFormulas", AutoSize = true, Checked = true };
    private readonly ListView _results = new() { Name = "lvFindResults", View = View.Details, FullRowSelect = true, Dock = DockStyle.Fill, Height = 120 };
    private readonly Func<(Workbook, Sheet)> _ctx;
    private readonly Action<Sheet, int, int> _goTo;
    private readonly Action _changed;
    private List<(Sheet Sheet, int Row, int Col)> _hits = new();
    private int _pos = -1;
    private readonly Func<WorkbookDoc> _doc;

    public FindReplaceDialog(Func<WorkbookDoc> doc, Func<(Workbook, Sheet)> ctx, Action<Sheet, int, int> goTo, Action changed, bool replace) : base(Loc.T("Tìm và Thay thế", "Find and Replace"), 520, 420)
    {
        Name = "findReplaceDialog";
        _doc = doc; _ctx = ctx; _goTo = goTo; _changed = changed;
        ShowInTaskbar = false;
        _case.Text = Loc.T("Phân biệt hoa/thường", "Match case");
        _whole.Text = Loc.T("Khớp toàn bộ nội dung ô", "Match entire cell contents");
        _book.Text = Loc.T("Trong cả sổ làm việc", "Within: Workbook");
        _formulas.Text = Loc.T("Tìm trong công thức", "Look in formulas");
        _results.Columns.Add(Loc.T("Trang tính", "Sheet"), Dpi.Scale(110));
        _results.Columns.Add(Loc.T("Ô", "Cell"), Dpi.Scale(60));
        _results.Columns.Add(Loc.T("Nội dung", "Value"), Dpi.Scale(260));
        _results.SelectedIndexChanged += (_, _) => { if (_results.SelectedIndices.Count > 0) { _pos = _results.SelectedIndices[0]; var h = _hits[_pos]; _goTo(h.Sheet, h.Row, h.Col); } };
        var t = Grid();
        t.Dock = DockStyle.Top; t.Height = Dpi.Scale(190);
        t.Controls.Add(L(Loc.T("Tìm:", "Find what:")), 0, 0); t.Controls.Add(_find, 1, 0);
        t.Controls.Add(L(Loc.T("Thay bằng:", "Replace with:")), 0, 1); t.Controls.Add(_replace, 1, 1);
        var opts = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };
        opts.Controls.AddRange([_case, _whole, _book, _formulas]);
        t.Controls.Add(opts, 1, 2);
        Body.Controls.Add(_results);
        Body.Controls.Add(t);
        AddButton(Loc.T("Đóng", "Close"), "btnFindClose", DialogResult.Cancel, false, (_, _) => Close());
        AddButton(Loc.T("Thay tất cả", "Replace All"), "btnReplaceAll", DialogResult.None, false, (_, _) => ReplaceAll());
        AddButton(Loc.T("Thay", "Replace"), "btnReplace", DialogResult.None, false, (_, _) => ReplaceOne());
        AddButton(Loc.T("Tìm tất cả", "Find All"), "btnFindAll", DialogResult.None, false, (_, _) => FindAll());
        AddButton(Loc.T("Tìm tiếp", "Find Next"), "btnFindNext", DialogResult.None, true, (_, _) => FindNext());
        AcceptButton = null;
        _find.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { FindNext(); e.SuppressKeyPress = true; } };
        Shown += (_, _) => (replace ? _replace : _find).Focus();
        _find.TextChanged += (_, _) => _pos = -1;
    }

    public string FindText { get => _find.Text; set => _find.Text = value; }
    public string ReplaceText { get => _replace.Text; set => _replace.Text = value; }
    public bool WholeWorkbook { get => _book.Checked; set => _book.Checked = value; }

    private SheetOps.FindOptions Opts() => new(_find.Text, _case.Checked, _whole.Checked, _formulas.Checked, _book.Checked);

    public int FindAll()
    {
        var (wb, s) = _ctx();
        if (_find.Text.Length == 0) return 0;
        _hits = SheetOps.FindAll(wb, s, Opts()).ToList();
        _results.BeginUpdate();
        _results.Items.Clear();
        foreach (var h in _hits.Take(5000)) _results.Items.Add(new ListViewItem([h.Sheet.Name, Addr.ToA1(h.Row, h.Col), SheetOps.DisplayText(h.Sheet, h.Row, h.Col)]));
        _results.EndUpdate();
        Text = Loc.T("Tìm và Thay thế", "Find and Replace") + " — " + Loc.F("tìm thấy {0} ô", "{0} cell(s) found", _hits.Count);
        return _hits.Count;
    }

    public bool FindNext()
    {
        if (_pos < 0 || _hits.Count == 0) FindAll();
        if (_hits.Count == 0) { Ux.Info(this, Loc.T("Không tìm thấy dữ liệu cần tìm.", "We couldn't find what you were looking for.")); return false; }
        _pos = (_pos + 1) % _hits.Count;
        var h = _hits[_pos];
        _goTo(h.Sheet, h.Row, h.Col);
        return true;
    }

    public void ReplaceOne()
    {
        if (_pos < 0 || _pos >= _hits.Count) { FindNext(); return; }
        var h = _hits[_pos];
        _doc().Checkpoint(Loc.T("Thay thế", "Replace"));
        SheetOps.ReplaceIn(h.Sheet, h.Row, h.Col, Opts(), _replace.Text);
        _doc().Recalc();
        _changed();
        FindAll();
        if (_hits.Count > 0) { _pos = Math.Min(_pos, _hits.Count - 1) - 1; FindNext(); }
    }

    public int ReplaceAll()
    {
        var (wb, s) = _ctx();
        if (_find.Text.Length == 0) return 0;
        _doc().Checkpoint(Loc.T("Thay thế tất cả", "Replace all"));
        int n = SheetOps.ReplaceAll(wb, s, Opts(), _replace.Text);
        _doc().Recalc();
        _changed();
        Ux.Info(this, Loc.F("Đã thay thế {0} ô.", "All done. We made {0} replacements.", n));
        FindAll();
        return n;
    }
}
