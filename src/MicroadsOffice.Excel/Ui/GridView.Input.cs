using System.Text.RegularExpressions;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

public sealed partial class GridView
{
    private enum Drag { None, Select, ColHeader, RowHeader, ResizeCol, ResizeRow, Fill, MoveObj, ResizeObj, Point }

    private TextBox _editor = null!;
    private string _editText = "";
    private bool _editing, _enterMode, _fromBar, _suppressEditorEvents;
    private int _editRow, _editCol;
    private Drag _drag;
    private int _dragIndex, _dragStart;
    private double _dragOrigSize;
    private Point _dragPoint;
    private RectangleF _dragBounds;
    private int _dragHandle;
    private CellRange? _fillTarget;
    private int _pointStart = -1, _pointLen;
    private CellPos _pointAnchor;
    private readonly System.Windows.Forms.Timer _autoScroll = new() { Interval = 60 };
    private Point _lastMouse;
    private readonly ToolTip _tip = new() { InitialDelay = 300, ReshowDelay = 100, AutoPopDelay = 15000 };
    private string? _tipText;
    private ListBox? _acList;
    private Label? _sigTip;

    public bool IsEditing => _editing;
    public string EditText => _editText;
    public CellPos EditCell => new(_editRow, _editCol);
    public event EventHandler? EditTextChanged;
    public event EventHandler<int>? FilterButtonClicked;
    public event EventHandler<Point>? CellContextMenu;
    public event EventHandler<HitArea>? HeaderContextMenu;
    public event EventHandler? ObjectDoubleClick;
    public event EventHandler<string>? StatusMessage;
    /// <summary>Clipboard shared by every open workbook in this window.</summary>
    public static ClipData? Clip;
    private static string? _clipText;

    private void InitEditor()
    {
        _editor = new TextBox
        {
            Name = "cellEditor", BorderStyle = BorderStyle.None, Multiline = true, Visible = false, WordWrap = false,
            AcceptsTab = false, AcceptsReturn = false, AccessibleName = Loc.T("Ô đang sửa", "Cell editor")
        };
        _editor.TextChanged += (_, _) =>
        {
            if (_suppressEditorEvents) return;
            _editText = _editor.Text;
            ResizeEditor();
            UpdateAutocomplete();
            Invalidate();
            EditTextChanged?.Invoke(this, EventArgs.Empty);
        };
        _editor.KeyDown += EditorKeyDown;
        _editor.KeyPress += (_, e) => { if (e.KeyChar == '\r' || e.KeyChar == '\t' || e.KeyChar == 27) e.Handled = true; };
        _editor.LostFocus += (_, _) => HideAutocomplete();
        Controls.Add(_editor);
        _autoScroll.Tick += (_, _) => AutoScrollTick();
        ImeMode = ImeMode.NoControl;
    }

    // ------------------------------------------------------------------ editing

    public bool CanEditActive(bool warn = true)
    {
        if (!Sheet.Protected || !Sheet.StyleAt(Active.Row, Active.Col).Locked) return true;
        if (warn) Ux.Warn(FindForm(), Loc.T("Ô này đang được bảo vệ. Hãy bỏ bảo vệ trang tính (Xem lại ▸ Bỏ bảo vệ trang tính) để sửa.",
            "This cell is protected. Unprotect the sheet (Review ▸ Unprotect Sheet) to change it."));
        return false;
    }

    /// <summary>Starts editing the active cell. <paramref name="initial"/> null = current content (F2).</summary>
    public void BeginEdit(string? initial, bool enterMode, bool fromBar = false)
    {
        if (_editing) return;
        if (!CanEditActive()) return;
        SelectObject(null);
        _editing = true;
        _enterMode = enterMode;
        _fromBar = fromBar;
        _editRow = Active.Row; _editCol = Active.Col;
        _pointStart = -1;
        var cell = Sheet.GetCell(_editRow, _editCol);
        string text = initial ?? EditableText(cell, Sheet.StyleAt(_editRow, _editCol));
        _suppressEditorEvents = true;
        _editor.Text = text;
        _suppressEditorEvents = false;
        _editText = text;
        var st = Sheet.StyleAt(_editRow, _editCol);
        _editor.Font = CellPainter.GetFont(st.FontName, st.FontSize * (float)Sheet.Zoom, CellPainter.StyleOf(st));
        _editor.ForeColor = st.FontColorValue;
        _editor.BackColor = st.Fill != 0 ? Color.FromArgb(255, Color.FromArgb(st.Fill)) : Color.White;
        _editor.TextAlign = HorizontalAlignment.Left;
        EnsureVisible(_editRow, _editCol);
        PositionEditor();
        _editor.Visible = true;
        if (!fromBar)
        {
            _editor.Focus();
            _editor.SelectionStart = _editor.Text.Length;
            _editor.SelectionLength = 0;
        }
        Invalidate();
        EditStateChanged?.Invoke(this, EventArgs.Empty);
        EditTextChanged?.Invoke(this, EventArgs.Empty);
    }

    public static string EditableText(Cell? cell, CellStyle st)
    {
        if (cell == null) return "";
        if (cell.Formula != null) return "=" + cell.Formula;
        var v = cell.Value;
        if (v.IsNumber)
        {
            if (NumFmt.IsDateFormat(st.NumberFormat))
            {
                var d = DateSerial.FromSerial(v.Num);
                var date = NumFmt.DayFirst ? d.ToString("dd/MM/yyyy") : d.ToString("M/d/yyyy");
                if (Math.Abs(v.Num - Math.Floor(v.Num)) > 1e-9) return (v.Num < 1 ? "" : date + " ") + d.ToString("H:mm:ss");
                return date;
            }
            var g = CellValue.FormatGeneralInvariant(v.Num);
            if (st.NumberFormat.Contains('%')) g = CellValue.FormatGeneralInvariant(Math.Round(v.Num * 100, 10)) + "%";
            return NumFmt.DecimalSep == '.' ? g : g.Replace('.', NumFmt.DecimalSep);
        }
        if (v.IsText && (ValueParser.TryNumber(v.Str!, out _, out _) || v.Str!.StartsWith('=')) && st.NumberFormat != "@") return "'" + v.Str;
        return v.ToString();
    }

    /// <summary>Called by the formula bar when the user types there.</summary>
    public void SetEditTextFromBar(string text, int caret)
    {
        if (!_editing) BeginEdit(text, false, fromBar: true);
        if (!_editing) return;
        _fromBar = true;
        _suppressEditorEvents = true;
        _editor.Text = text;
        _editor.SelectionStart = Math.Min(caret, text.Length);
        _suppressEditorEvents = false;
        _editText = text;
        ResizeEditor();
        Invalidate();
    }

    public void InsertIntoEditor(string text)
    {
        if (!_editing) BeginEdit("", true);
        if (!_editing) return;
        int pos = _editor.SelectionStart;
        _editor.Text = _editor.Text.Insert(pos, text);
        _editor.SelectionStart = pos + text.Length;
        if (!_fromBar) _editor.Focus();
    }

    internal void PositionEditor()
    {
        if (!_editing || _editor == null) return;
        var r = CellRect(_editRow, _editCol);
        if (r == null) { _editor.Visible = false; return; }
        _editor.Visible = !_fromBar || true;
        var rr = r.Value;
        _editor.SetBounds(rr.X + 2, rr.Y + 2, Math.Max(rr.Width - 3, 10), Math.Max(rr.Height - 3, 10));
        ResizeEditor();
    }

    private void ResizeEditor()
    {
        if (!_editing) return;
        var r = CellRect(_editRow, _editCol);
        if (r == null) return;
        var lines = _editor.Text.Split('\n');
        var sz = TextRenderer.MeasureText(lines.OrderByDescending(l => l.Length).FirstOrDefault() ?? "", _editor.Font);
        int w = Math.Max(r.Value.Width - 3, sz.Width + Dpi.Scale(14));
        w = Math.Min(w, Width - BarW - r.Value.X - 2);
        int lineH = TextRenderer.MeasureText("Ag", _editor.Font).Height;
        int h = Math.Max(r.Value.Height - 3, lineH * lines.Length + 2);
        _editor.Size = new Size(Math.Max(10, w), h);
    }

    public void CancelEdit()
    {
        if (!_editing) return;
        _editing = false;
        _editor.Visible = false;
        HideAutocomplete();
        Focus();
        Invalidate();
        EditStateChanged?.Invoke(this, EventArgs.Empty);
        EditTextChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Commits the edit and moves (dr, dc). Returns false when validation rejected the value.</summary>
    public bool CommitEdit(int dr = 1, int dc = 0)
    {
        if (!_editing) return true;
        var text = _editText;
        if (text.StartsWith('=') && text.Length > 1)
        {
            // auto-close parentheses like Excel
            int open = text.Count(ch => ch == '('), close = text.Count(ch => ch == ')');
            if (open > close && !InString(text)) text += new string(')', open - close);
            if (!FormulaParser.TryParse(text, out _, out var err))
            {
                Ux.Warn(FindForm(), Loc.T("Công thức có lỗi: ", "There is a problem with this formula: ") + err +
                    Loc.T("\n\nMẹo: dùng dấu phẩy (,) để ngăn cách đối số, ví dụ =SUM(A1,B1).", "\n\nTip: separate arguments with commas, e.g. =SUM(A1,B1)."));
                if (!_fromBar) _editor.Focus();
                return false;
            }
        }
        if (!ValidateInput(text)) { if (!_fromBar) _editor.Focus(); return false; }
        _editing = false;
        _editor.Visible = false;
        HideAutocomplete();
        var cell = Sheet.GetCell(_editRow, _editCol);
        string old = EditableText(cell, Sheet.StyleAt(_editRow, _editCol));
        if (text != old)
        {
            _doc.Checkpoint(Loc.T("Nhập dữ liệu", "Typing"));
            var range = Selection;
            SheetOps.SetInput(Sheet, _editRow, _editCol, text);
            _doc.Recalc();
            RaiseContentChanged();
        }
        Focus();
        EditStateChanged?.Invoke(this, EventArgs.Empty);
        MoveAfterEnter(dr, dc);
        EditTextChanged?.Invoke(this, EventArgs.Empty);
        return true;
    }

    /// <summary>Ctrl+Enter: puts the typed value into every selected cell.</summary>
    public void CommitEditToSelection()
    {
        if (!_editing) return;
        var text = _editText;
        if (!ValidateInput(text)) return;
        _editing = false;
        _editor.Visible = false;
        HideAutocomplete();
        _doc.Checkpoint(Loc.T("Nhập dữ liệu", "Typing"));
        var sel = Selection;
        if (sel.Count > 200_000) sel = new CellRange(sel.R1, sel.C1, Math.Min(sel.R2, sel.R1 + 9999), Math.Min(sel.C2, sel.C1 + 199));
        foreach (var (r, c) in sel.Cells())
        {
            var t = text.StartsWith('=') ? "=" + RefAdjust.Shift(text[1..], r - _editRow, c - _editCol) : text;
            SheetOps.SetInput(Sheet, r, c, t);
        }
        _doc.Recalc();
        RaiseContentChanged();
        Focus();
        Invalidate();
        EditStateChanged?.Invoke(this, EventArgs.Empty);
    }

    private static bool InString(string s) => s.Count(c => c == '"') % 2 == 1;

    private bool ValidateInput(string text)
    {
        var dv = ValidationAt(_editRow, _editCol);
        if (dv == null || dv.Type == ValidationType.Any || text.StartsWith('=')) return true;
        if (text.Length == 0) return dv.AllowBlank || true;
        var v = ValueParser.Parse(text, out _);
        bool ok = dv.Type switch
        {
            ValidationType.List => ListItems(dv).Any(x => string.Equals(x, text.Trim(), StringComparison.CurrentCultureIgnoreCase)),
            ValidationType.WholeNumber => v.IsNumber && v.Num == Math.Floor(v.Num) && InRange(dv, v.Num),
            ValidationType.Decimal => v.IsNumber && InRange(dv, v.Num),
            ValidationType.Date => v.IsNumber && InRange(dv, v.Num),
            ValidationType.TextLength => InRange(dv, text.Length),
            _ => true
        };
        if (ok) return true;
        Ux.Warn(FindForm(), dv.ErrorMessage.Length > 0 ? dv.ErrorMessage : Loc.T("Giá trị này không khớp với quy tắc xác thực dữ liệu của ô.", "This value doesn't match the data validation restrictions defined for this cell."));
        return false;
    }

    private bool InRange(DataValidation dv, double x)
    {
        double a = EvalNum(dv.Formula1), b = EvalNum(dv.Formula2);
        return dv.Op switch
        {
            ValidationOp.Between => x >= Math.Min(a, b) && x <= Math.Max(a, b),
            ValidationOp.NotBetween => x < Math.Min(a, b) || x > Math.Max(a, b),
            ValidationOp.Equal => x == a, ValidationOp.NotEqual => x != a, ValidationOp.Greater => x > a, ValidationOp.Less => x < a,
            ValidationOp.GreaterOrEqual => x >= a, _ => x <= a
        };
    }

    private double EvalNum(string f)
    {
        if (string.IsNullOrWhiteSpace(f)) return 0;
        if (ValueParser.TryCoerceNumber(f, out var d)) return d;
        var v = Book.Calc.EvaluateFormula(Sheet, _editRow, _editCol, f.TrimStart('='));
        return v.IsNumber ? v.Num : 0;
    }

    public List<string> ListItems(DataValidation dv)
    {
        var f = dv.Formula1.Trim();
        if (f.StartsWith('='))
        {
            try
            {
                var e = FormulaParser.Parse(f[1..]);
                var ctx = Book.Calc.EvaluateFormula(Sheet, Active.Row, Active.Col, "0");
                if (e is RefExpr re)
                {
                    var sh = re.Sheet != null ? Book.Sheet(re.Sheet) ?? Sheet : Sheet;
                    return re.Range.Cells().Take(5000).Select(p => SheetOps.DisplayText(sh, p.R, p.C)).Where(x => x.Length > 0).ToList();
                }
                if (e is NameExpr ne && Book.Names.TryGetValue(ne.Name, out var refText) && FormulaParser.Parse(refText) is RefExpr nr)
                {
                    var sh = nr.Sheet != null ? Book.Sheet(nr.Sheet) ?? Sheet : Sheet;
                    return nr.Range.Cells().Take(5000).Select(p => SheetOps.DisplayText(sh, p.R, p.C)).Where(x => x.Length > 0).ToList();
                }
            }
            catch { }
            return new();
        }
        return f.Split(f.Contains(';') && !f.Contains(',') ? ';' : ',').Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
    }

    private void MoveAfterEnter(int dr, int dc)
    {
        if (dr == 0 && dc == 0) { Invalidate(); return; }
        var sel = Selection;
        if (!sel.IsSingle && sel.Contains(Active.Row, Active.Col) && Sheet.MergeAt(sel.R1, sel.C1) != sel)
        {
            // move inside the selection
            int r = Active.Row, c = Active.Col;
            if (dr != 0) { r += dr; if (r > sel.R2) { r = sel.R1; c = c + 1 > sel.C2 ? sel.C1 : c + 1; } else if (r < sel.R1) { r = sel.R2; c = c - 1 < sel.C1 ? sel.C2 : c - 1; } }
            else { c += dc; if (c > sel.C2) { c = sel.C1; r = r + 1 > sel.R2 ? sel.R1 : r + 1; } else if (c < sel.C1) { c = sel.C2; r = r - 1 < sel.R1 ? sel.R2 : r - 1; } }
            Sheet.ActiveCell = new CellPos(r, c);
            EnsureVisible(r, c);
            Invalidate();
            RaiseSelectionChanged();
            return;
        }
        var m = Sheet.MergeAt(Active.Row, Active.Col);
        int nr = Active.Row + dr, nc = Active.Col + dc;
        if (m is CellRange mm) { if (dr > 0) nr = mm.R2 + 1; if (dc > 0) nc = mm.C2 + 1; }
        while (nr >= 0 && nr < Addr.MaxRows && Sheet.IsRowHidden(nr) && dr != 0) nr += dr;
        while (nc >= 0 && nc < Addr.MaxCols && Sheet.HiddenCols.Contains(nc) && dc != 0) nc += dc;
        Select(Math.Max(0, nr), Math.Max(0, nc));
    }

    private void EditorKeyDown(object? sender, KeyEventArgs e)
    {
        if (_acList != null && _acList.Visible)
        {
            if (e.KeyCode == Keys.Down) { _acList.SelectedIndex = Math.Min(_acList.Items.Count - 1, _acList.SelectedIndex + 1); e.Handled = e.SuppressKeyPress = true; return; }
            if (e.KeyCode == Keys.Up) { _acList.SelectedIndex = Math.Max(0, _acList.SelectedIndex - 1); e.Handled = e.SuppressKeyPress = true; return; }
            if (e.KeyCode == Keys.Tab) { AcceptAutocomplete(); e.Handled = e.SuppressKeyPress = true; return; }
            if (e.KeyCode == Keys.Escape) { HideAutocomplete(); e.Handled = e.SuppressKeyPress = true; return; }
        }
        HandleEditKey(e, _editor);
    }

    /// <summary>Shared key handling for the in-cell editor and the formula bar.</summary>
    public void HandleEditKey(KeyEventArgs e, TextBox box)
    {
        switch (e.KeyCode)
        {
            case Keys.Enter when e.Alt:
            {
                int p = box.SelectionStart;
                box.Text = box.Text.Insert(p, "\r\n");
                box.SelectionStart = p + 2;
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            case Keys.Enter when e.Control: CommitEditToSelection(); e.Handled = e.SuppressKeyPress = true; return;
            case Keys.Enter: CommitEdit(e.Shift ? -1 : 1, 0); e.Handled = e.SuppressKeyPress = true; return;
            case Keys.Tab: CommitEdit(0, e.Shift ? -1 : 1); e.Handled = e.SuppressKeyPress = true; return;
            case Keys.Escape: CancelEdit(); e.Handled = e.SuppressKeyPress = true; return;
            case Keys.F2: _enterMode = !_enterMode; StatusMessage?.Invoke(this, _enterMode ? Loc.T("Nhập", "Enter") : Loc.T("Sửa", "Edit")); e.Handled = true; return;
            case Keys.F4:
                ToggleAbsolute(box);
                e.Handled = e.SuppressKeyPress = true;
                return;
        }
        if (_enterMode && e.KeyCode is Keys.Up or Keys.Down or Keys.Left or Keys.Right && !e.Control)
        {
            int dr = e.KeyCode == Keys.Up ? -1 : e.KeyCode == Keys.Down ? 1 : 0, dc = e.KeyCode == Keys.Left ? -1 : e.KeyCode == Keys.Right ? 1 : 0;
            if (CanInsertRef())
            {
                PointMove(dr, dc, e.Shift);
                e.Handled = e.SuppressKeyPress = true;
                return;
            }
            CommitEdit(dr, dc);
            e.Handled = e.SuppressKeyPress = true;
            return;
        }
        if (e.KeyCode is not (Keys.ShiftKey or Keys.ControlKey)) _pointStart = -1;
    }

    /// <summary>F4 cycles $A$1 → A$1 → $A1 → A1 for the reference at the caret.</summary>
    private void ToggleAbsolute(TextBox box)
    {
        var text = box.Text;
        int caret = box.SelectionStart;
        foreach (Match m in Regex.Matches(text, @"\$?[A-Za-z]{1,3}\$?\d+"))
        {
            if (caret < m.Index || caret > m.Index + m.Length) continue;
            var t = m.Value;
            bool ac = t.StartsWith('$'), ar = t.IndexOf('$', 1) > 0;
            var bare = t.Replace("$", "");
            int split = bare.IndexOfAny("0123456789".ToCharArray());
            string col = bare[..split], row = bare[split..];
            string next = (ac, ar) switch { (false, false) => "$" + col + "$" + row, (true, true) => col + "$" + row, (false, true) => "$" + col + row, _ => col + row };
            box.Text = text[..m.Index] + next + text[(m.Index + m.Length)..];
            box.SelectionStart = m.Index + next.Length;
            return;
        }
    }

    // ------------------------------------------------------------------ point mode (click / arrows insert references)

    public bool CanInsertRef()
    {
        if (!_editing || !_editText.StartsWith('=')) return false;
        if (_pointStart >= 0) return true;
        int caret = _fromBar ? _editText.Length : _editor.SelectionStart;
        if (caret <= 0) return false;
        int i = caret - 1;
        while (i > 0 && _editText[i] == ' ') i--;
        return "=+-*/^&(,:<>;%".Contains(_editText[i]);
    }

    private void InsertPointRef(CellRange r)
    {
        var refText = (r.IsSingle ? Addr.ToA1(r.R1, r.C1) : Addr.ToA1(r.R1, r.C1) + ":" + Addr.ToA1(r.R2, r.C2));
        if (Book.Active != Book.Sheets.FirstOrDefault(x => x == Sheet)) refText = Addr.QuoteSheet(Sheet.Name) + "!" + refText;
        var text = _editText;
        int start, len;
        if (_pointStart >= 0 && _pointStart + _pointLen <= text.Length) { start = _pointStart; len = _pointLen; }
        else { start = _fromBar ? text.Length : _editor.SelectionStart; len = 0; }
        text = text[..start] + refText + text[(start + len)..];
        _pointStart = start; _pointLen = refText.Length;
        _suppressEditorEvents = true;
        _editor.Text = text;
        _editor.SelectionStart = start + refText.Length;
        _suppressEditorEvents = false;
        _editText = text;
        ResizeEditor();
        Invalidate();
        EditTextChanged?.Invoke(this, EventArgs.Empty);
    }

    private void PointMove(int dr, int dc, bool extend)
    {
        if (_pointStart < 0) _pointAnchor = new CellPos(_editRow, _editCol);
        if (!extend)
        {
            _pointAnchor = new CellPos(Math.Clamp(_pointAnchor.Row + dr, 0, Addr.MaxRows - 1), Math.Clamp(_pointAnchor.Col + dc, 0, Addr.MaxCols - 1));
            _pointEnd = _pointAnchor;
        }
        else _pointEnd = new CellPos(Math.Clamp(_pointEnd.Row + dr, 0, Addr.MaxRows - 1), Math.Clamp(_pointEnd.Col + dc, 0, Addr.MaxCols - 1));
        InsertPointRef(new CellRange(_pointAnchor.Row, _pointAnchor.Col, _pointEnd.Row, _pointEnd.Col).Normalized);
        EnsureVisible(_pointEnd.Row, _pointEnd.Col);
    }

    private CellPos _pointEnd;

    /// <summary>References in the text being edited (for coloured highlights).</summary>
    public IEnumerable<(CellRange Range, string? Sheet)> EditRefs()
    {
        Expr? e = null;
        try { e = FormulaParser.Parse(_editText.Length > 1 ? _editText[1..] : "0"); } catch { }
        if (e != null)
        {
            foreach (var r in RefAdjust.Refs(e)) yield return (r.Range, r.Sheet);
            yield break;
        }
        foreach (Match m in Regex.Matches(_editText, @"(?<![A-Za-z_])\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?"))
            if (Addr.TryParseRange(m.Value, out var rg)) yield return (rg, null);
    }

    // ------------------------------------------------------------------ autocomplete

    private void UpdateAutocomplete()
    {
        if (!_editing || !_editText.StartsWith('=') || _fromBar) { HideAutocomplete(); UpdateSignature(); return; }
        int caret = _editor.SelectionStart;
        int i = caret;
        while (i > 0 && (char.IsLetterOrDigit(_editText[i - 1]) || _editText[i - 1] == '.')) i--;
        var word = _editText[i..caret];
        bool afterOp = i > 0 && "=+-*/^&(,:<>; ".Contains(_editText[i - 1]);
        if (word.Length == 0 || !afterOp || char.IsDigit(word[0]))
        {
            HideAutocomplete();
            UpdateSignature();
            return;
        }
        var matches = FunctionCatalog.All.Where(f => f.Name.StartsWith(word, StringComparison.OrdinalIgnoreCase)).Take(12).ToList();
        if (matches.Count == 0 || (matches.Count == 1 && matches[0].Name.Equals(word, StringComparison.OrdinalIgnoreCase) && caret < _editText.Length && _editText[caret] == '('))
        {
            HideAutocomplete();
            UpdateSignature();
            return;
        }
        _acList ??= CreateAcList();
        _acList.BeginUpdate();
        _acList.Items.Clear();
        foreach (var m in matches) _acList.Items.Add(m);
        _acList.SelectedIndex = 0;
        _acList.EndUpdate();
        var host = FindForm()!;
        var p = host.PointToClient(_editor.Parent!.PointToScreen(new Point(_editor.Left, _editor.Bottom + 2)));
        _acList.SetBounds(p.X, p.Y, Dpi.Scale(330), Math.Min(matches.Count, 8) * _acList.ItemHeight + 4);
        if (_acList.Parent != host) host.Controls.Add(_acList);
        _acList.Visible = true;
        _acList.BringToFront();
        _tip.Show(matches[0].Description, host, p.X + _acList.Width + 4, p.Y, 6000);
    }

    private ListBox CreateAcList()
    {
        var lb = new ListBox { Name = "fnAutocomplete", IntegralHeight = false, Font = Theme.UiFont, BorderStyle = BorderStyle.FixedSingle, TabStop = false, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = Dpi.Scale(20) };
        lb.DrawItem += (_, e) =>
        {
            if (e.Index < 0) return;
            var f = (FunctionInfo)lb.Items[e.Index];
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using var b = new SolidBrush(sel ? Theme.Hover : Color.White);
            e.Graphics.FillRectangle(b, e.Bounds);
            using var icon = XlIcons.Get("function", Dpi.Scale(14));
            e.Graphics.DrawImage(icon, e.Bounds.X + 3, e.Bounds.Y + 3);
            TextRenderer.DrawText(e.Graphics, f.Name, Theme.UiFontBold, new Point(e.Bounds.X + Dpi.Scale(22), e.Bounds.Y + 2), Theme.Ink);
        };
        lb.DoubleClick += (_, _) => AcceptAutocomplete();
        lb.SelectedIndexChanged += (_, _) =>
        {
            if (lb.SelectedItem is FunctionInfo f && lb.Visible) _tip.Show(f.Description, lb.Parent!, lb.Right + 4, lb.Top + (lb.SelectedIndex - lb.TopIndex) * lb.ItemHeight, 6000);
        };
        return lb;
    }

    private void AcceptAutocomplete()
    {
        if (_acList?.SelectedItem is not FunctionInfo f) return;
        int caret = _editor.SelectionStart, i = caret;
        while (i > 0 && (char.IsLetterOrDigit(_editText[i - 1]) || _editText[i - 1] == '.')) i--;
        var text = _editText[..i] + f.Name + "(" + _editText[caret..];
        HideAutocomplete();
        _editor.Text = text;
        _editor.SelectionStart = i + f.Name.Length + 1;
        _editor.Focus();
    }

    private void HideAutocomplete()
    {
        if (_acList != null) _acList.Visible = false;
        _tip.Hide(FindForm() ?? (IWin32Window)this);
    }

    private void UpdateSignature()
    {
        var host = FindForm();
        if (host == null) return;
        string? sig = null;
        if (_editing && _editText.StartsWith('='))
        {
            int caret = _fromBar ? _editText.Length : _editor.SelectionStart;
            int depth = 0;
            for (int i = Math.Min(caret, _editText.Length) - 1; i >= 0; i--)
            {
                char ch = _editText[i];
                if (ch == ')') depth++;
                else if (ch == '(')
                {
                    if (depth > 0) { depth--; continue; }
                    int j = i;
                    while (j > 0 && (char.IsLetterOrDigit(_editText[j - 1]) || _editText[j - 1] == '.')) j--;
                    var name = _editText[j..i];
                    var fi = FunctionCatalog.Find(name);
                    if (fi != null) sig = fi.Signature;
                    break;
                }
            }
        }
        if (sig == null || _fromBar) { if (_sigTip != null) _sigTip.Visible = false; return; }
        _sigTip ??= new Label { Name = "fnSignature", AutoSize = true, BackColor = Color.FromArgb(255, 255, 240), BorderStyle = BorderStyle.FixedSingle, Font = Theme.SmallFont, Padding = new Padding(3) };
        _sigTip.Text = sig;
        var p = host.PointToClient(_editor.Parent!.PointToScreen(new Point(_editor.Left, _editor.Bottom + 2)));
        _sigTip.Location = p;
        if (_sigTip.Parent != host) host.Controls.Add(_sigTip);
        _sigTip.Visible = true;
        _sigTip.BringToFront();
    }

    // ------------------------------------------------------------------ keyboard (grid focused)

    protected override bool IsInputKey(Keys keyData)
    {
        var k = keyData & Keys.KeyCode;
        if (k is Keys.Up or Keys.Down or Keys.Left or Keys.Right or Keys.Tab or Keys.Enter or Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown) return true;
        return base.IsInputKey(keyData);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (_doc == null || e.Handled) return;
        if (SelectedObject != null)
        {
            if (e.KeyCode == Keys.Delete) { DeleteSelectedObject(); e.Handled = true; }
            else if (e.KeyCode == Keys.Escape) SelectObject(null);
            return;
        }
        var s = Sheet;
        int r = Active.Row, c = Active.Col;
        bool shift = e.Shift, ctrl = e.Control;
        // extension end for shift-selection
        var sel = Selection;
        int er = sel.R1 == r ? sel.R2 : sel.R1, ec = sel.C1 == c ? sel.C2 : sel.C1;
        void Go(int nr, int nc)
        {
            if (shift) { ExtendTo(nr, nc); }
            else Select(nr, nc);
        }
        switch (e.KeyCode)
        {
            case Keys.Up: Go(ctrl ? Jump(shift ? er : r, shift ? ec : c, -1, 0) : NextVisibleRow((shift ? er : r), -1), shift ? ec : c); break;
            case Keys.Down: Go(ctrl ? Jump(shift ? er : r, shift ? ec : c, 1, 0) : NextVisibleRow((shift ? er : r), 1), shift ? ec : c); break;
            case Keys.Left: Go(shift ? er : r, ctrl ? JumpCol(shift ? er : r, shift ? ec : c, -1) : NextVisibleCol(shift ? ec : c, -1)); break;
            case Keys.Right: Go(shift ? er : r, ctrl ? JumpCol(shift ? er : r, shift ? ec : c, 1) : NextVisibleCol(shift ? ec : c, 1)); break;
            case Keys.Home: Go(ctrl ? 0 : r, 0); break;
            case Keys.End when ctrl: { var u = s.UsedRange(); Go(u.R2, u.C2); break; }
            case Keys.PageDown: { int n = Math.Max(1, _rows.Count(x => x.Index >= s.FreezeRows) - 1); ScrollTo(s.ScrollRow + n, s.ScrollCol); Go(Math.Min(Addr.MaxRows - 1, r + n), c); break; }
            case Keys.PageUp: { int n = Math.Max(1, _rows.Count(x => x.Index >= s.FreezeRows) - 1); ScrollTo(Math.Max(s.FreezeRows, s.ScrollRow - n), s.ScrollCol); Go(Math.Max(0, r - n), c); break; }
            case Keys.Enter: MoveAfterEnter(shift ? -1 : 1, 0); break;
            case Keys.Tab: MoveAfterEnter(0, shift ? -1 : 1); break;
            case Keys.F2: BeginEdit(null, false); break;
            case Keys.Delete: ClearSelection(contents: true, formats: false); break;
            case Keys.Back: if (CanEditActive()) BeginEdit("", true); break;
            case Keys.Escape: if (CopyMarquee != null) { CopyMarquee = null; Invalidate(); } break;
            case Keys.Space when ctrl && !shift: SelectRange(new CellRange(0, sel.C1, Addr.MaxRows - 1, sel.C2), Active); break;
            case Keys.Space when shift && !ctrl: SelectRange(new CellRange(sel.R1, 0, sel.R2, Addr.MaxCols - 1), Active); break;
            default: return;
        }
        e.Handled = true;
    }

    private void ExtendTo(int r, int c)
    {
        var a = Active;
        Sheet.Selection = ExpandMerges(new CellRange(a.Row, a.Col, Math.Clamp(r, 0, Addr.MaxRows - 1), Math.Clamp(c, 0, Addr.MaxCols - 1)).Normalized);
        EnsureVisible(r, c);
        Invalidate();
        RaiseSelectionChanged();
    }

    private int NextVisibleRow(int r, int d) { r += d; while (r > 0 && r < Addr.MaxRows - 1 && Sheet.IsRowHidden(r)) r += d; return Math.Clamp(r, 0, Addr.MaxRows - 1); }
    private int NextVisibleCol(int c, int d) { c += d; while (c > 0 && c < Addr.MaxCols - 1 && Sheet.HiddenCols.Contains(c)) c += d; return Math.Clamp(c, 0, Addr.MaxCols - 1); }

    private bool Filled(int r, int c) { var cell = Sheet.GetCell(r, c); return cell != null && (!cell.Value.IsBlank || cell.Formula != null); }

    /// <summary>Ctrl+arrow: jump to the edge of the data region.</summary>
    private int Jump(int r, int c, int d, int _)
    {
        var used = Sheet.UsedRange();
        int limit = d > 0 ? Math.Max(used.R2, r) : 0;
        int n = r + d;
        if (n < 0 || n >= Addr.MaxRows) return r;
        if (Filled(r, c) && Filled(n, c)) { while (n + d >= 0 && n + d <= limit && Filled(n + d, c)) n += d; return n; }
        while (n >= 0 && n <= limit && !Filled(n, c)) n += d;
        if (n < 0) return 0;
        if (n > limit) return d > 0 ? Addr.MaxRows - 1 : 0;
        return n;
    }

    private int JumpCol(int r, int c, int d)
    {
        var used = Sheet.UsedRange();
        int limit = d > 0 ? Math.Max(used.C2, c) : 0;
        int n = c + d;
        if (n < 0 || n >= Addr.MaxCols) return c;
        if (Filled(r, c) && Filled(r, n)) { while (n + d >= 0 && n + d <= limit && Filled(r, n + d)) n += d; return n; }
        while (n >= 0 && n <= limit && !Filled(r, n)) n += d;
        if (n < 0) return 0;
        if (n > limit) return d > 0 ? Addr.MaxCols - 1 : 0;
        return n;
    }

    protected override void OnKeyPress(KeyPressEventArgs e)
    {
        base.OnKeyPress(e);
        if (_doc == null || _editing || SelectedObject != null) return;
        if (char.IsControl(e.KeyChar)) return;
        if (!CanEditActive()) { e.Handled = true; return; }
        BeginEdit(e.KeyChar.ToString(), true);
        e.Handled = true;
    }

    // ------------------------------------------------------------------ mouse

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (_doc == null) return;
        var hit = HitTest(e.Location);
        _dragPoint = e.Location;
        if (e.Button == MouseButtons.Right)
        {
            if (_editing) CommitEdit(0, 0);
            if (hit.Area == HitArea.Cell && !AllSelections.Any(x => x.Contains(hit.Row, hit.Col))) Select(hit.Row, hit.Col);
            if (hit.Area == HitArea.ColHeader && !(Selection.R1 == 0 && Selection.R2 == Addr.MaxRows - 1 && hit.Col >= Selection.C1 && hit.Col <= Selection.C2))
                SelectRange(new CellRange(0, hit.Col, Addr.MaxRows - 1, hit.Col), new CellPos(Math.Max(Sheet.ScrollRow, 0), hit.Col));
            if (hit.Area == HitArea.RowHeader && !(Selection.C1 == 0 && Selection.C2 == Addr.MaxCols - 1 && hit.Row >= Selection.R1 && hit.Row <= Selection.R2))
                SelectRange(new CellRange(hit.Row, 0, hit.Row, Addr.MaxCols - 1), new CellPos(hit.Row, Sheet.ScrollCol));
            if (hit.Area is HitArea.Object or HitArea.ObjectHandle) SelectObject(hit.Obj);
            if (hit.Area is HitArea.ColHeader or HitArea.RowHeader) HeaderContextMenu?.Invoke(this, hit.Area);
            else CellContextMenu?.Invoke(this, e.Location);
            return;
        }
        if (e.Button != MouseButtons.Left) return;
        Focus();
        if (_editing)
        {
            if (hit.Area == HitArea.Cell && CanInsertRef())
            {
                _pointAnchor = new CellPos(hit.Row, hit.Col); _pointEnd = _pointAnchor;
                if (_pointStart < 0) { }
                InsertPointRef(CellRange.Single(hit.Row, hit.Col));
                _drag = Drag.Point;
                Capture = true;
                if (!_fromBar) _editor.Focus();
                return;
            }
            if (!CommitEdit(0, 0)) return;
        }
        switch (hit.Area)
        {
            case HitArea.Corner: SelectRange(new CellRange(0, 0, Addr.MaxRows - 1, Addr.MaxCols - 1), new CellPos(Sheet.ScrollRow, Sheet.ScrollCol)); break;
            case HitArea.ColHeader:
                if (ModifierKeys.HasFlag(Keys.Shift)) SelectRange(new CellRange(0, Math.Min(Active.Col, hit.Col), Addr.MaxRows - 1, Math.Max(Active.Col, hit.Col)), Active);
                else { SelectRange(new CellRange(0, hit.Col, Addr.MaxRows - 1, hit.Col), new CellPos(Math.Max(Sheet.FreezeRows, Sheet.ScrollRow) == 0 ? 0 : Sheet.ScrollRow, hit.Col)); Sheet.ActiveCell = new CellPos(0, hit.Col); }
                _drag = Drag.ColHeader; _dragStart = hit.Col; Capture = true;
                break;
            case HitArea.RowHeader:
                if (ModifierKeys.HasFlag(Keys.Shift)) SelectRange(new CellRange(Math.Min(Active.Row, hit.Row), 0, Math.Max(Active.Row, hit.Row), Addr.MaxCols - 1), Active);
                else { SelectRange(new CellRange(hit.Row, 0, hit.Row, Addr.MaxCols - 1), new CellPos(hit.Row, 0)); }
                _drag = Drag.RowHeader; _dragStart = hit.Row; Capture = true;
                break;
            case HitArea.ColBorder:
                _drag = Drag.ResizeCol; _dragIndex = hit.Col; _dragOrigSize = Sheet.RawColWidth(hit.Col); Capture = true;
                break;
            case HitArea.RowBorder:
                _drag = Drag.ResizeRow; _dragIndex = hit.Row; _dragOrigSize = Sheet.RawRowHeight(hit.Row); Capture = true;
                break;
            case HitArea.FillHandle:
                _drag = Drag.Fill; _fillTarget = null; Capture = true;
                break;
            case HitArea.Object:
                SelectObject(hit.Obj);
                _drag = Drag.MoveObj; _dragBounds = BoundsOf(hit.Obj!); Capture = true;
                break;
            case HitArea.ObjectHandle:
                _drag = Drag.ResizeObj; _dragHandle = hit.Handle; _dragBounds = BoundsOf(hit.Obj!); Capture = true;
                break;
            case HitArea.FilterButton:
                Select(hit.Row, hit.Col);
                FilterButtonClicked?.Invoke(this, hit.Col);
                break;
            case HitArea.ValidationArrow:
                ShowValidationList();
                break;
            case HitArea.Cell:
                if (ModifierKeys.HasFlag(Keys.Control) && Sheet.Links.TryGetValue(Sheet.Key(hit.Row, hit.Col), out var url)) { FollowLink(url); return; }
                if (ModifierKeys.HasFlag(Keys.Shift)) ExtendTo(hit.Row, hit.Col);
                else Select(hit.Row, hit.Col, add: ModifierKeys.HasFlag(Keys.Control));
                _drag = Drag.Select; Capture = true;
                break;
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (_doc == null) return;
        _lastMouse = e.Location;
        if (_drag == Drag.None)
        {
            var hit = HitTest(e.Location);
            Cursor = hit.Area switch
            {
                HitArea.ColBorder => Cursors.SizeWE,
                HitArea.RowBorder => Cursors.SizeNS,
                HitArea.FillHandle => Cursors.Cross,
                HitArea.ColHeader or HitArea.RowHeader => Cursors.Hand,
                HitArea.Object => Cursors.SizeAll,
                HitArea.ObjectHandle => hit.Handle is 0 or 4 ? Cursors.SizeNWSE : hit.Handle is 2 or 6 ? Cursors.SizeNESW : hit.Handle is 1 or 5 ? Cursors.SizeNS : Cursors.SizeWE,
                HitArea.FilterButton or HitArea.ValidationArrow => Cursors.Default,
                HitArea.Cell => Cursors.Cross,
                _ => Cursors.Default
            };
            UpdateHoverTip(hit);
            return;
        }
        bool outside = e.X < RowHeaderW || e.Y < ColHeaderH || e.X > Width - BarW || e.Y > Height - BarH;
        if (outside && _drag is Drag.Select or Drag.Fill or Drag.Point) _autoScroll.Start(); else _autoScroll.Stop();
        DragTo(e.Location);
    }

    private void DragTo(Point p)
    {
        int row = RowAtY(Math.Max(p.Y, ColHeaderH + 1)), col = ColAtX(Math.Max(p.X, RowHeaderW + 1));
        switch (_drag)
        {
            case Drag.Select:
                if (row >= 0 && col >= 0) ExtendTo(row, col);
                break;
            case Drag.Point:
                if (row >= 0 && col >= 0) { _pointEnd = new CellPos(row, col); InsertPointRef(new CellRange(_pointAnchor.Row, _pointAnchor.Col, row, col).Normalized); }
                break;
            case Drag.ColHeader:
                if (col >= 0) { Sheet.Selection = new CellRange(0, Math.Min(_dragStart, col), Addr.MaxRows - 1, Math.Max(_dragStart, col)); Invalidate(); RaiseSelectionChanged(); }
                break;
            case Drag.RowHeader:
                if (row >= 0) { Sheet.Selection = new CellRange(Math.Min(_dragStart, row), 0, Math.Max(_dragStart, row), Addr.MaxCols - 1); Invalidate(); RaiseSelectionChanged(); }
                break;
            case Drag.ResizeCol:
            {
                var cr = _cols.FirstOrDefault(c => c.Index == _dragIndex);
                double w = Math.Max(0, (p.X - cr.Pos) / Scale);
                Sheet.ColWidths[_dragIndex] = Math.Round(w);
                ComputeLayout(); Invalidate();
                StatusMessage?.Invoke(this, Loc.F("Độ rộng: {0} điểm ảnh", "Width: {0} pixels", Math.Round(w)));
                break;
            }
            case Drag.ResizeRow:
            {
                var rr = _rows.FirstOrDefault(r => r.Index == _dragIndex);
                double h = Math.Max(0, (p.Y - rr.Pos) / Scale);
                Sheet.RowHeights[_dragIndex] = Math.Round(h);
                ComputeLayout(); Invalidate();
                StatusMessage?.Invoke(this, Loc.F("Chiều cao: {0} điểm ảnh", "Height: {0} pixels", Math.Round(h)));
                break;
            }
            case Drag.Fill:
            {
                var s = Selection;
                if (row < 0 || col < 0) break;
                int dDown = row - s.R2, dUp = s.R1 - row, dRight = col - s.C2, dLeft = s.C1 - col;
                int best = new[] { dDown, dUp, dRight, dLeft }.Max();
                if (best <= 0) _fillTarget = null;
                else if (best == dDown) _fillTarget = s with { R2 = row };
                else if (best == dUp) _fillTarget = s with { R1 = row };
                else if (best == dRight) _fillTarget = s with { C2 = col };
                else _fillTarget = s with { C1 = col };
                Invalidate();
                break;
            }
            case Drag.MoveObj when SelectedObject != null:
            {
                float dx = (p.X - _dragPoint.X) / Scale, dy = (p.Y - _dragPoint.Y) / Scale;
                SetBounds(SelectedObject, new RectangleF(Math.Max(0, _dragBounds.X + dx), Math.Max(0, _dragBounds.Y + dy), _dragBounds.Width, _dragBounds.Height));
                Invalidate();
                break;
            }
            case Drag.ResizeObj when SelectedObject != null:
            {
                float dx = (p.X - _dragPoint.X) / Scale, dy = (p.Y - _dragPoint.Y) / Scale;
                var b = _dragBounds;
                float l = b.Left, t = b.Top, r = b.Right, bt = b.Bottom;
                if (_dragHandle is 0 or 6 or 7) l += dx;
                if (_dragHandle is 2 or 3 or 4) r += dx;
                if (_dragHandle is 0 or 1 or 2) t += dy;
                if (_dragHandle is 4 or 5 or 6) bt += dy;
                if (r - l < 40) { if (_dragHandle is 0 or 6 or 7) l = r - 40; else r = l + 40; }
                if (bt - t < 30) { if (_dragHandle is 0 or 1 or 2) t = bt - 30; else bt = t + 30; }
                SetBounds(SelectedObject, new RectangleF(l, t, r - l, bt - t));
                Invalidate();
                break;
            }
        }
    }

    private void AutoScrollTick()
    {
        if (_drag == Drag.None || !Capture) { _autoScroll.Stop(); return; }
        var p = _lastMouse;
        int dr = p.Y > Height - BarH ? 1 : p.Y < ColHeaderH ? -1 : 0;
        int dc = p.X > Width - BarW ? 1 : p.X < RowHeaderW ? -1 : 0;
        if (dr == 0 && dc == 0) { _autoScroll.Stop(); return; }
        ScrollTo(Sheet.ScrollRow + dr, Sheet.ScrollCol + dc);
        DragTo(new Point(Math.Clamp(p.X, RowHeaderW + 1, Width - BarW - 1), Math.Clamp(p.Y, ColHeaderH + 1, Height - BarH - 1)));
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        _autoScroll.Stop();
        var d = _drag;
        _drag = Drag.None;
        Capture = false;
        switch (d)
        {
            case Drag.ResizeCol:
            {
                var w = Sheet.ColWidths.GetValueOrDefault(_dragIndex, Sheet.DefaultColWidth);
                Sheet.ColWidths[_dragIndex] = _dragOrigSize;
                if (Math.Abs(w - _dragOrigSize) > 0.5)
                {
                    _doc.Checkpoint(Loc.T("Độ rộng cột", "Column width"));
                    var cols = Selection.R1 == 0 && Selection.R2 == Addr.MaxRows - 1 && _dragIndex >= Selection.C1 && _dragIndex <= Selection.C2 ? Enumerable.Range(Selection.C1, Selection.Cols) : [_dragIndex];
                    foreach (var c in cols) { if (w <= 0) { Sheet.HiddenCols.Add(c); } else { Sheet.HiddenCols.Remove(c); Sheet.ColWidths[c] = w; } }
                    RaiseContentChanged();
                }
                RefreshAll();
                break;
            }
            case Drag.ResizeRow:
            {
                var h = Sheet.RowHeights.GetValueOrDefault(_dragIndex, Sheet.DefaultRowHeight);
                Sheet.RowHeights[_dragIndex] = _dragOrigSize;
                if (Math.Abs(_dragOrigSize - Sheet.DefaultRowHeight) < 0.01 && !Sheet.RowHeights.ContainsKey(_dragIndex)) { }
                if (Math.Abs(h - _dragOrigSize) > 0.5)
                {
                    _doc.Checkpoint(Loc.T("Chiều cao hàng", "Row height"));
                    var rows = Selection.C1 == 0 && Selection.C2 == Addr.MaxCols - 1 && _dragIndex >= Selection.R1 && _dragIndex <= Selection.R2 ? Enumerable.Range(Selection.R1, Selection.Rows) : [_dragIndex];
                    foreach (var r in rows) { if (h <= 0) Sheet.HiddenRows.Add(r); else { Sheet.HiddenRows.Remove(r); Sheet.RowHeights[r] = h; } }
                    RaiseContentChanged();
                }
                RefreshAll();
                break;
            }
            case Drag.Fill:
                if (_fillTarget is CellRange t)
                {
                    _doc.Checkpoint(Loc.T("Điền", "Fill"));
                    SheetOps.Fill(Sheet, Selection, t, copyOnly: ModifierKeys.HasFlag(Keys.Control));
                    Sheet.Selection = t;
                    _doc.Recalc();
                    RaiseContentChanged();
                    RaiseSelectionChanged();
                }
                _fillTarget = null;
                Invalidate();
                break;
            case Drag.MoveObj: case Drag.ResizeObj:
                if (SelectedObject != null)
                {
                    var nb = BoundsOf(SelectedObject);
                    if (nb != _dragBounds)
                    {
                        SetBounds(SelectedObject, _dragBounds);
                        _doc.Checkpoint(Loc.T("Di chuyển đối tượng", "Move object"));
                        // the checkpoint cloned the book: the selected object is now the live one again
                        SetBounds(SelectedObject, nb);
                        RaiseContentChanged();
                    }
                }
                break;
        }
    }

    protected override void OnMouseDoubleClick(MouseEventArgs e)
    {
        base.OnMouseDoubleClick(e);
        if (_doc == null || e.Button != MouseButtons.Left) return;
        var hit = HitTest(e.Location);
        switch (hit.Area)
        {
            case HitArea.ColBorder: AutoFitColumns(SelectedColsOr(hit.Col)); break;
            case HitArea.RowBorder: AutoFitRows(SelectedRowsOr(hit.Row)); break;
            case HitArea.Object: ObjectDoubleClick?.Invoke(this, EventArgs.Empty); break;
            case HitArea.FillHandle:
            {
                // double-click fill handle: fill down as far as the neighbouring column has data
                var s = Selection;
                int nb = s.C1 > 0 && Filled(s.R2 + 1, s.C1 - 1) ? s.C1 - 1 : Filled(s.R2 + 1, s.C2 + 1) ? s.C2 + 1 : -1;
                if (nb < 0) break;
                int end = s.R2;
                while (Filled(end + 1, nb)) end++;
                _doc.Checkpoint(Loc.T("Điền", "Fill"));
                SheetOps.Fill(Sheet, s, s with { R2 = end });
                Sheet.Selection = s with { R2 = end };
                _doc.Recalc();
                RaiseContentChanged();
                Invalidate();
                break;
            }
            case HitArea.Cell: if (!_editing) BeginEdit(null, false); break;
        }
    }

    private IEnumerable<int> SelectedColsOr(int c) => Selection.R1 == 0 && Selection.R2 == Addr.MaxRows - 1 && c >= Selection.C1 && c <= Selection.C2 ? Enumerable.Range(Selection.C1, Math.Min(Selection.Cols, 500)) : [c];
    private IEnumerable<int> SelectedRowsOr(int r) => Selection.C1 == 0 && Selection.C2 == Addr.MaxCols - 1 && r >= Selection.R1 && r <= Selection.R2 ? Enumerable.Range(Selection.R1, Math.Min(Selection.Rows, 5000)) : [r];

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        base.OnMouseWheel(e);
        if (_doc == null) return;
        if (ModifierKeys.HasFlag(Keys.Control))
        {
            SetZoom(Math.Round((Sheet.Zoom + (e.Delta > 0 ? 0.1 : -0.1)) * 10) / 10);
            return;
        }
        int lines = Math.Max(1, SystemInformation.MouseWheelScrollLines);
        if (ModifierKeys.HasFlag(Keys.Shift)) ScrollTo(Sheet.ScrollRow, Sheet.ScrollCol + (e.Delta > 0 ? -lines : lines));
        else ScrollTo(Sheet.ScrollRow + (e.Delta > 0 ? -lines : lines), Sheet.ScrollCol);
    }

    private void UpdateHoverTip(Hit hit)
    {
        string? text = null;
        if (hit.Area == HitArea.Cell)
        {
            var k = Sheet.Key(hit.Row, hit.Col);
            if (Sheet.Notes.TryGetValue(k, out var note)) text = note;
            else if (Sheet.Links.TryGetValue(k, out var url)) text = url + "\n" + Loc.T("Ctrl + bấm để mở liên kết", "Ctrl + click to follow the link");
        }
        if (text == _tipText) return;
        _tipText = text;
        if (text == null) _tip.Hide(this);
        else _tip.Show(text, this, _lastMouse.X + 14, _lastMouse.Y + 16, 15000);
    }

    // ------------------------------------------------------------------ objects

    internal static RectangleF BoundsOf(object o) => o is Chart c ? c.Bounds : o is Picture p ? p.Bounds : RectangleF.Empty;
    private static void SetBounds(object o, RectangleF b) { if (o is Chart c) c.Bounds = b; else if (o is Picture p) p.Bounds = b; }

    public void DeleteSelectedObject()
    {
        if (SelectedObject == null) return;
        var obj = SelectedObject;
        int ci = Sheet.Charts.IndexOf(obj as Chart ?? new Chart()), pi = Sheet.Pictures.IndexOf(obj as Picture ?? new Picture());
        _doc.Checkpoint(Loc.T("Xóa đối tượng", "Delete object"));
        if (ci >= 0) Sheet.Charts.RemoveAt(ci);
        if (pi >= 0) Sheet.Pictures.RemoveAt(pi);
        SelectedObject = null;
        ObjectSelectionChanged?.Invoke(this, EventArgs.Empty);
        RaiseContentChanged();
        Invalidate();
    }

    /// <summary>After undo/redo the workbook objects are new instances: re-select by index.</summary>
    internal void RebindSelection() { SelectedObject = null; ObjectSelectionChanged?.Invoke(this, EventArgs.Empty); }

    private void FollowLink(string url)
    {
        if (url.StartsWith('#'))
        {
            var target = url[1..];
            var sheetName = target.Contains('!') ? target[..target.LastIndexOf('!')].Trim('\'') : null;
            var refPart = target.Contains('!') ? target[(target.LastIndexOf('!') + 1)..] : target;
            if (sheetName != null && Book.Sheet(sheetName) is Sheet sh) { Book.ActiveSheet = Book.Sheets.IndexOf(sh); RefreshAll(); }
            if (Addr.TryParseRange(refPart, out var r)) SelectRange(r);
            RaiseSelectionChanged();
            return;
        }
        OfficeShell.OpenWebsite(url.Contains("://") ? url : "https://" + url);
    }

    // ------------------------------------------------------------------ validation list

    public void ShowValidationList()
    {
        var dv = ValidationAt(Active.Row, Active.Col);
        if (dv == null || dv.Type != ValidationType.List) return;
        var items = ListItems(dv);
        if (items.Count == 0) return;
        var lb = new ListBox { Name = "validationList", IntegralHeight = false, Font = Theme.UiFont, BorderStyle = BorderStyle.None };
        lb.Items.AddRange(items.Cast<object>().ToArray());
        var cur = SheetOps.DisplayText(Sheet, Active.Row, Active.Col);
        lb.SelectedItem = items.FirstOrDefault(x => x == cur);
        var host = new ToolStripControlHost(lb) { AutoSize = false, Size = new Size(Math.Max(CellRect(Active.Row, Active.Col)?.Width ?? 120, Dpi.Scale(120)), Math.Min(items.Count, 10) * lb.ItemHeight + 6), Margin = Padding.Empty, Padding = Padding.Empty };
        var dd = new ToolStripDropDown { Padding = new Padding(1), Name = "validationDropDown" };
        dd.Items.Add(host);
        void Pick()
        {
            if (lb.SelectedItem is string v)
            {
                dd.Close();
                _doc.Checkpoint(Loc.T("Chọn từ danh sách", "Pick from list"));
                SheetOps.SetInput(Sheet, Active.Row, Active.Col, v);
                _doc.Recalc();
                RaiseContentChanged();
                Invalidate();
            }
        }
        lb.Click += (_, _) => Pick();
        lb.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) Pick(); };
        var r = CellRect(Active.Row, Active.Col) ?? new Rectangle(0, 0, 100, 20);
        dd.Show(this, new Point(r.X, r.Bottom));
        lb.Focus();
    }

    // ------------------------------------------------------------------ commands

    private IEnumerable<CellRange> ClippedSelections()
    {
        var used = Sheet.UsedRange();
        foreach (var r in AllSelections)
            yield return r.Count > 1_000_000 ? new CellRange(r.R1, r.C1, Math.Min(r.R2, Math.Max(used.R2, r.R1)), Math.Min(r.C2, Math.Max(used.C2, r.C1))) : r;
    }

    public void ClearSelection(bool contents, bool formats)
    {
        if (Sheet.Protected) { CanEditActive(); return; }
        _doc.Checkpoint(contents && formats ? Loc.T("Xóa tất cả", "Clear all") : contents ? Loc.T("Xóa nội dung", "Clear contents") : Loc.T("Xóa định dạng", "Clear formats"));
        foreach (var r in ClippedSelections())
        {
            var keys = Sheet.Cells.Keys.Where(k => { var (row, col) = Sheet.FromKey(k); return r.Contains(row, col); }).ToList();
            foreach (var k in keys) { var (row, col) = Sheet.FromKey(k); Sheet.ClearCell(row, col, contents, formats); }
            if (formats)
            {
                Sheet.Merges.RemoveAll(m => r.Contains(m));
                if (r.R1 == 0 && r.R2 == Addr.MaxRows - 1) for (int c = r.C1; c <= r.C2; c++) Sheet.ColStyles.Remove(c);
                if (r.C1 == 0 && r.C2 == Addr.MaxCols - 1) for (int rr = r.R1; rr <= r.R2; rr++) Sheet.RowStyles.Remove(rr);
            }
            if (contents && formats)
                foreach (var k in Sheet.Notes.Keys.Concat(Sheet.Links.Keys).ToList()) { var (row, col) = Sheet.FromKey(k); if (r.Contains(row, col)) { Sheet.Notes.Remove(k); Sheet.Links.Remove(k); } }
        }
        _doc.Recalc();
        RaiseContentChanged();
        Invalidate();
    }

    private const string PictureClipFormat = "MicroadsExcelPicture";
    private static Picture? _clipPicture;
    private static string? _clipPictureToken;

    public void CopySelection(bool cut)
    {
        if (_editing) { (_fromBar ? null : _editor)?.Copy(); return; }
        if (SelectedObject is Picture pic && PictureImage(pic) is Image img)
        {
            // a picture goes out as a picture (bitmap + PNG) for every program, plus a token so Ctrl+V here keeps its size
            var data = new DataObject();
            DataTransfer.PutImage(data, img);
            _clipPicture = pic.Clone();
            _clipPictureToken = Guid.NewGuid().ToString("N");
            data.SetData(PictureClipFormat, _clipPictureToken);
            DataTransfer.WriteClipboard(data);
            // copied cells no longer apply: Ctrl+V must paste this picture
            Clip = null; CopyMarquee = null;
            if (cut && !Sheet.Protected) DeleteSelectedObject();
            Invalidate();
            return;
        }
        if (ExtraRanges.Count > 0) { Ux.Warn(FindForm(), Loc.T("Không thể sao chép nhiều vùng chọn rời nhau.", "This action won't work on multiple selections.")); return; }
        if (cut && Sheet.Protected) { CanEditActive(); return; }
        var sel = Selection;
        Clip = SheetOps.Copy(Sheet, sel, cut);
        _clipText = SheetOps.ToTsv(Sheet, Clip.Range);
        _clipPicture = null;
        // Other programs (chat apps, clipboard managers) may hold the clipboard for a moment: retry instead of failing silently.
        DataTransfer.WriteClipboard(new DataObject(DataFormats.UnicodeText, _clipText.Length > 0 ? _clipText : " "));
        CopyMarquee = sel;
        Invalidate();
        StatusMessage?.Invoke(this, Loc.T("Chọn ô đích rồi nhấn Enter hoặc Ctrl+V để dán", "Select destination and press Enter or Ctrl+V"));
    }

    public void Paste(PasteMode mode = PasteMode.All, bool transpose = false, bool skipBlanks = false)
    {
        if (_editing) { if (!_fromBar) _editor.Paste(); return; }
        if (Sheet.Protected) { CanEditActive(); return; }
        var data = DataTransfer.ReadClipboard();
        var sysText = DataTransfer.Text(data);
        var files = DataTransfer.Files(data);
        // our own copied picture: pasted with the same size
        if (files.Length == 0 && _clipPicture != null && data != null && data.GetDataPresent(PictureClipFormat) && data.GetData(PictureClipFormat) as string == _clipPictureToken)
        {
            AddPicture(_clipPicture.Data, _clipPicture.Ext, _clipPicture.Bounds.Size);
            return;
        }
        var target = Selection;
        // copied cells, unless something else was copied since (a screenshot, files, text from another program)
        bool somethingElse = files.Length > 0 || (sysText == null && DataTransfer.Picture(data) != null);
        if (Clip != null && !somethingElse && (sysText == null || sysText == _clipText || (sysText == " " && _clipText == "")))
        {
            _doc.Checkpoint(Loc.T("Dán", "Paste"));
            var pasted = SheetOps.Paste(Sheet, Clip, target, mode, transpose, skipBlanks);
            if (Clip.IsCut) { Clip = null; CopyMarquee = null; }
            Sheet.Selection = pasted;
            _doc.Recalc();
            RaiseContentChanged();
            RaiseSelectionChanged();
            RefreshAll();
            return;
        }
        InsertData(data, sysText, files);
    }

    /// <summary>
    /// Data from other programs (Ctrl+V or drag &amp; drop) at the active cell: files and folders (pictures become pictures,
    /// other files and folders become links in the cells below), then text, then a picture (screenshot, copied image).
    /// Text wins over a picture because spreadsheet programs put both on the clipboard when cells are copied.
    /// </summary>
    public bool InsertData(IDataObject? data) => InsertData(data, DataTransfer.Text(data), DataTransfer.Files(data));

    private bool InsertData(IDataObject? data, string? text, string[] files)
    {
        if (data == null) return false;
        if (files.Length > 0) { InsertFiles(files); return true; }
        if (text == null && DataTransfer.Picture(data) is { } pic) { AddPicture(pic.Bytes, pic.Ext, null); return true; }
        if (text == null) return false;
        _doc.Checkpoint(Loc.T("Dán", "Paste"));
        var r = SheetOps.PasteText(Sheet, text.TrimEnd('\r', '\n'), Selection.R1, Selection.C1);
        Sheet.Selection = r;
        _doc.Recalc();
        RaiseContentChanged();
        RaiseSelectionChanged();
        RefreshAll();
        return true;
    }

    /// <summary>Files and folders as ONE undo step: pictures are placed at the active cell, other files / folders are written as links down the column.</summary>
    public void InsertFiles(IReadOnlyList<string> files)
    {
        var images = DataTransfer.ImageFiles(files.Where(DataTransfer.IsImageFile));
        var others = files.Where(f => !DataTransfer.IsImageFile(f)).ToList();
        if (images.Count == 0 && others.Count == 0) return;
        _doc.Checkpoint(Loc.T("Dán", "Paste"));
        var a = Active;
        for (int i = 0; i < images.Count; i++) AddPicture(images[i].Bytes, images[i].Ext, null, i * 24, checkpoint: false);
        if (others.Count > 0)
        {
            for (int k = 0; k < others.Count; k++)
            {
                Sheet.SetValue(a.Row + k, a.Col, CellValue.Text(others[k]));
                Sheet.Links[Sheet.Key(a.Row + k, a.Col)] = new Uri(others[k]).AbsoluteUri;
            }
            if (images.Count == 0) Sheet.Selection = new CellRange(a.Row, a.Col, a.Row + others.Count - 1, a.Col);
            _doc.Recalc();
            RaiseSelectionChanged();
        }
        RaiseContentChanged();
        RefreshAll();
    }

    /// <summary>Drops dragged data on the cell under the mouse.</summary>
    public bool DropData(IDataObject? data, Point client)
    {
        if (_editing) CommitEdit();
        var hit = HitTest(client);
        if (hit.Area == HitArea.Cell) Select(hit.Row, hit.Col);
        if (Sheet.Protected) { CanEditActive(); return false; }
        Focus();
        return InsertData(data);
    }

    public void InsertPicture(Image img)
    {
        using var ms = new MemoryStream();
        img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        AddPicture(ms.ToArray(), "png", null);
    }

    /// <summary>Places a picture at the active cell (natural size, at most 600 wide) and selects it.</summary>
    private void AddPicture(byte[] data, string ext, SizeF? size, float offset = 0, bool checkpoint = true)
    {
        float w, h;
        if (size is { } sz) { w = sz.Width; h = sz.Height; }
        else
        {
            using var img = Image.FromStream(new MemoryStream(data));
            w = img.Width; h = img.Height;
            if (w > 600) { h = h * 600 / w; w = 600; }
        }
        var x = (float)SheetX(Active.Col) + offset; var y = (float)SheetY(Active.Row) + offset;
        if (checkpoint) _doc.Checkpoint(Loc.T("Chèn ảnh", "Insert picture"));
        var pic = new Picture { Data = data, Ext = ext, Bounds = new RectangleF(x, y, w, h), Name = "Picture " + (Sheet.Pictures.Count + 1) };
        Sheet.Pictures.Add(pic);
        SelectObject(pic);
        RaiseContentChanged();
    }

    public void AutoFitColumns(IEnumerable<int> cols)
    {
        _doc.Checkpoint(Loc.T("Tự vừa độ rộng cột", "AutoFit column width"));
        using var g = CreateGraphics();
        foreach (var c in cols)
        {
            var w = SheetOps.AutoFitWidth(Sheet, c, (text, st) => CellPainter.Measure(g, text.Split('\n').OrderByDescending(x => x.Length).First(), CellPainter.GetFont(st.FontName, st.FontSize, CellPainter.StyleOf(st)), true, 0).Width / Dpi.Factor);
            Sheet.ColWidths[c] = Math.Round(w);
            Sheet.HiddenCols.Remove(c);
        }
        RaiseContentChanged();
        RefreshAll();
    }

    public void AutoFitRows(IEnumerable<int> rows)
    {
        _doc.Checkpoint(Loc.T("Tự vừa chiều cao hàng", "AutoFit row height"));
        using var g = CreateGraphics();
        var used = Sheet.UsedRange();
        foreach (var r in rows)
        {
            double h = Sheet.DefaultRowHeight;
            for (int c = 0; c <= used.C2; c++)
            {
                var cell = Sheet.GetCell(r, c);
                if (cell == null || cell.Value.IsBlank) continue;
                var st = Sheet.Book.Styles[cell.StyleId];
                var text = NumFmt.Format(cell.Value, st.NumberFormat, out _);
                var f = CellPainter.GetFont(st.FontName, st.FontSize, CellPainter.StyleOf(st));
                var sz = CellPainter.Measure(g, text, f, true, st.Wrap ? (float)(Sheet.ColWidth(c) * Dpi.Factor - 6) : 0);
                h = Math.Max(h, sz.Height / Dpi.Factor + 5);
            }
            if (Math.Abs(h - Sheet.DefaultRowHeight) < 0.5) Sheet.RowHeights.Remove(r); else Sheet.RowHeights[r] = Math.Ceiling(h);
            Sheet.HiddenRows.Remove(r);
        }
        RaiseContentChanged();
        RefreshAll();
    }
}
