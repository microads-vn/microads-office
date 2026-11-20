using System.Drawing.Text;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Main window of Microads Office Excel.</summary>
public sealed partial class ExcelForm : Form, IOfficeAppWindow
{
    public readonly RibbonBar Ribbon = new();
    public readonly GridView Grid = new();
    public readonly SheetTabs Tabs = new();
    public readonly Backstage BackstageView = new(AppKind.Excel);
    public readonly List<WorkbookDoc> Docs = new();
    private WorkbookDoc _doc = null!;
    private readonly ComboBox _nameBox = new() { Name = "cboNameBox", DropDownStyle = ComboBoxStyle.DropDown };
    private readonly TextBox _formula = new() { Name = "txtFormulaBar", BorderStyle = BorderStyle.None, Multiline = true, WordWrap = false, AcceptsReturn = false };
    private readonly Panel _formulaPanel = new() { Name = "formulaBar", Dock = DockStyle.Top };
    private readonly ToolStrip _docTabs = new() { Name = "workbookTabs", Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, Visible = false };
    private readonly Panel _status = new() { Name = "statusBar", Dock = DockStyle.Bottom };
    private readonly Label _statusMode = new() { Name = "lblMode", AutoSize = true };
    private readonly Label _statusStats = new() { Name = "lblStats", AutoSize = true };
    private readonly TrackBar _zoom = new() { Name = "trkZoom", Minimum = 10, Maximum = 400, TickStyle = TickStyle.None, AutoSize = false, SmallChange = 10, LargeChange = 25 };
    private readonly Label _zoomLabel = new() { Name = "lblZoom", AutoSize = false, TextAlign = ContentAlignment.MiddleRight };
    private readonly Panel _main = new() { Dock = DockStyle.Fill, Name = "mainArea" };
    private StartScreen? _start;
    private PrintPage? _printPage;
    private readonly Dictionary<string, Control> _controls = new();
    private readonly List<ToolStripMenuItem> _menuItems = new();
    private RibbonPage _chartPage = null!, _picturePage = null!;
    private readonly System.Windows.Forms.Timer _autoRecover = new() { Interval = 60_000 };
    private bool _updating;
    private FindReplaceDialog? _find;
    private int _untitled = 1;
    private ComboBox _fontBox = null!, _sizeBox = null!, _numFmtBox = null!;
    private FormatPainterState? _painter;

    public WorkbookDoc Doc => _doc;
    public IReadOnlyDictionary<string, Control> CommandControls => _controls;
    public IReadOnlyList<ToolStripMenuItem> MenuItems => _menuItems;
    public TextBox FormulaBar => _formula;
    public ComboBox NameBox => _nameBox;
    public string StatusStats => _statusStats.Text;

    public ExcelForm(string[] files)
    {
        Name = "excelForm";
        Font = Theme.UiFont;
        AutoScaleMode = AutoScaleMode.None;
        KeyPreview = true;
        AllowDrop = true;
        Icon = OfficeShell.AppIcon;
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1366, 728);
        Size = new Size(Math.Min(Dpi.Scale(1360), (int)(wa.Width * 0.94)), Math.Min(Dpi.Scale(860), (int)(wa.Height * 0.94)));
        MinimumSize = new Size(Dpi.Scale(720), Dpi.Scale(480));
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.White;
        ApplyNumberStyle();

        BuildRibbon();
        BuildFormulaBar();
        BuildStatusBar();
        BuildBackstage();
        InitAi();

        Tabs.SheetSelected += (_, i) => SwitchSheet(i);
        Tabs.AddClicked += (_, _) => CmdInsertSheet();
        Tabs.RenameRequested += (_, i) => { SwitchSheet(i); CmdRenameSheet(); };
        Tabs.MenuRequested += (_, a) => SheetMenu().Show(Tabs, a.Location);
        Tabs.MoveRequested += (_, m) => MoveSheet(m.From, m.To);

        Grid.SelectionChanged += (_, _) => { UpdateCellInfo(); UpdateStats(); };
        Grid.ContentChanged += (_, _) => OnContentChanged();
        Grid.EditStateChanged += (_, _) => { UpdateMode(); SyncFormulaBar(); };
        Grid.EditTextChanged += (_, _) => SyncFormulaBar();
        Grid.ObjectSelectionChanged += (_, _) => UpdateContextualTabs();
        Grid.ZoomChanged += (_, z) => SyncZoom();
        Grid.FilterButtonClicked += (_, col) => ShowFilterPopup(col);
        Grid.CellContextMenu += (_, p) => CellMenu().Show(Grid, p);
        Grid.HeaderContextMenu += (_, a) => HeaderMenu(a == HitArea.ColHeader).Show(Grid, Grid.PointToClient(Cursor.Position));
        Grid.ObjectDoubleClick += (_, _) => { if (Grid.SelectedObject is Chart) CmdEditChart(); };
        Grid.StatusMessage += (_, m) => _statusMode.Text = m;
        Grid.MouseUp += (_, _) => ApplyPainterIfActive();

        Grid.Dock = DockStyle.Fill;
        _main.Controls.Add(Grid);
        _main.Controls.Add(Tabs);
        Controls.Add(_main);
        Controls.Add(BackstageView);
        Controls.Add(_formulaPanel);
        Controls.Add(_docTabs);
        Controls.Add(Ribbon);
        Controls.Add(_status);
        BackstageView.BringToFront();

        DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        DragDrop += (_, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] f) OpenFiles(f, false); };
        // the sheet accepts what other programs drag: text, pictures (browser, chat app, Paint), files and folders
        Grid.AllowDrop = true;
        Grid.DragEnter += (_, e) => e.Effect = DataTransfer.HasContent(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
        Grid.DragOver += (_, e) => e.Effect = DataTransfer.HasContent(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
        Grid.DragDrop += (_, e) =>
        {
            if (e.Data == null) return;
            var pt = Grid.PointToClient(new Point(e.X, e.Y));
            var files = DataTransfer.Files(e.Data);
            if (files.Length == 0) { Grid.DropData(e.Data, pt); return; }
            // workbooks / documents open in their app; pictures and any other file or folder go into the sheet
            var docs = files.Where(f => File.Exists(f) && !DataTransfer.IsImageFile(f) && Product.KindForFile(f) != null).ToList();
            var rest = files.Except(docs).ToArray();
            if (rest.Length > 0) Grid.DropData(new DataObject(DataFormats.FileDrop, rest), pt);
            if (docs.Count > 0) OpenFiles(docs, false);
        };
        FormClosing += OnFormClosing;
        _autoRecover.Tick += (_, _) => AutoRecoverTick();
        _autoRecover.Start();

        NewDocument(Workbook.CreateBlank(), null);
        Load += (_, _) =>
        {
            RestoreWindow();
            bool opened = false;
            foreach (var f in files) opened |= OpenFile(f);
            if (!opened)
            {
                if (!OfferRecovery()) ShowBackstage("home");
            }
            UpdateAll();
        };
    }

    // ------------------------------------------------------------------ layout pieces

    private void BuildFormulaBar()
    {
        _formulaPanel.Height = Dpi.Scale(30);
        _formulaPanel.BackColor = Color.White;
        _formulaPanel.Paint += (_, e) =>
        {
            using var p = new Pen(Theme.Border);
            e.Graphics.DrawLine(p, 0, _formulaPanel.Height - 1, _formulaPanel.Width, _formulaPanel.Height - 1);
            e.Graphics.DrawLine(p, _nameBox.Right + Dpi.Scale(4), Dpi.Scale(5), _nameBox.Right + Dpi.Scale(4), _formulaPanel.Height - Dpi.Scale(6));
        };
        _nameBox.Font = Theme.UiFont;
        _nameBox.AccessibleName = Loc.T("Hộp tên", "Name Box");
        _nameBox.KeyDown += (_, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            e.SuppressKeyPress = true;
            GoToReference(_nameBox.Text);
            Grid.Focus();
        };
        _nameBox.DropDown += (_, _) => { _nameBox.Items.Clear(); foreach (var n in _doc.Book.Names.Keys.OrderBy(x => x)) _nameBox.Items.Add(n); };
        _nameBox.SelectionChangeCommitted += (_, _) => BeginInvoke(() => { GoToReference(_nameBox.Text); Grid.Focus(); });
        var cancel = MkBarButton("btnFormulaCancel", "✕", Loc.T("Hủy (Esc)", "Cancel (Esc)"), () => Grid.CancelEdit());
        var enter = MkBarButton("btnFormulaEnter", "✓", Loc.T("Nhập (Enter)", "Enter"), () => Grid.CommitEdit(0, 0));
        var fx = MkBarButton("btnFormulaFx", "fx", Loc.T("Chèn hàm (Shift+F3)", "Insert Function (Shift+F3)"), CmdInsertFunction);
        fx.Font = new Font("Cambria", 10f, FontStyle.Italic);
        _formula.Font = new Font("Segoe UI", 10f);
        _formula.AccessibleName = Loc.T("Thanh công thức", "Formula bar");
        _formula.TextChanged += (_, _) =>
        {
            if (_updating || !_formula.Focused) return;
            Grid.SetEditTextFromBar(_formula.Text, _formula.SelectionStart);
        };
        _formula.KeyDown += (_, e) =>
        {
            if (e.KeyCode is Keys.Enter or Keys.Tab or Keys.Escape or Keys.F4 || (e.KeyCode == Keys.Enter && e.Alt))
            {
                if (!Grid.IsEditing && e.KeyCode == Keys.Escape) { Grid.Focus(); e.SuppressKeyPress = true; return; }
                if (!Grid.IsEditing) Grid.SetEditTextFromBar(_formula.Text, _formula.SelectionStart);
                Grid.HandleEditKey(e, _formula);
                if (!Grid.IsEditing) Grid.Focus();
            }
        };
        _formula.Enter += (_, _) => { if (!Grid.IsEditing && !Grid.CanEditActive(false)) { } };
        _formulaPanel.Controls.AddRange([_nameBox, cancel, enter, fx, _formula]);
        _formulaPanel.Resize += (_, _) => LayoutFormulaBar(cancel, enter, fx);
        LayoutFormulaBar(cancel, enter, fx);
    }

    private void LayoutFormulaBar(Control cancel, Control enter, Control fx)
    {
        int h = _formulaPanel.Height, y = Dpi.Scale(3);
        _nameBox.SetBounds(Dpi.Scale(6), y + 1, Dpi.Scale(110), h - 8);
        int x = _nameBox.Right + Dpi.Scale(10);
        foreach (var c in new[] { cancel, enter, fx }) { c.SetBounds(x, y, Dpi.Scale(26), h - Dpi.Scale(7)); x += Dpi.Scale(27); }
        _formula.SetBounds(x + Dpi.Scale(6), Dpi.Scale(6), Math.Max(50, _formulaPanel.Width - x - Dpi.Scale(14)), h - Dpi.Scale(10));
    }

    private Button MkBarButton(string name, string text, string tip, Action a)
    {
        var b = new Button { Name = name, Text = text, FlatStyle = FlatStyle.Flat, TabStop = false, ForeColor = Theme.SubtleInk, AccessibleName = tip };
        b.FlatAppearance.BorderSize = 0;
        b.FlatAppearance.MouseOverBackColor = Theme.Hover;
        b.Click += (_, _) => a();
        Ribbon.ToolTip.SetToolTip(b, tip);
        _controls[name] = b;
        return b;
    }

    private void BuildStatusBar()
    {
        _status.Height = Dpi.Scale(26);
        _status.BackColor = Theme.TabStripBack;
        _statusMode.Text = Loc.T("Sẵn sàng", "Ready");
        _statusMode.ForeColor = Theme.SubtleInk;
        _statusStats.ForeColor = Theme.Ink;
        _zoom.Value = 100;
        _zoom.Width = Dpi.Scale(130);
        _zoom.Height = Dpi.Scale(24);
        _zoom.AccessibleName = Loc.T("Thu phóng", "Zoom");
        _zoom.ValueChanged += (_, _) => { if (!_updating) Grid.SetZoom(_zoom.Value / 100.0); };
        _zoomLabel.Width = Dpi.Scale(48);
        _zoomLabel.Cursor = Cursors.Hand;
        _zoomLabel.Click += (_, _) => CmdZoomDialog();
        _status.Controls.AddRange([_statusMode, _statusStats, _zoom, _zoomLabel]);
        _status.Resize += (_, _) => LayoutStatus();
        LayoutStatus();
    }

    private void LayoutStatus()
    {
        int h = _status.Height;
        _statusMode.Location = new Point(Dpi.Scale(10), (h - _statusMode.Height) / 2);
        _zoomLabel.SetBounds(_status.Width - _zoomLabel.Width - Dpi.Scale(10), 0, _zoomLabel.Width, h);
        _zoom.Location = new Point(_zoomLabel.Left - _zoom.Width - Dpi.Scale(4), (h - _zoom.Height) / 2 + 1);
        _statusStats.Location = new Point(_zoom.Left - _statusStats.Width - Dpi.Scale(24), (h - _statusStats.Height) / 2);
    }

    // ------------------------------------------------------------------ documents

    public void NewDocument(Workbook wb, string? path)
    {
        var doc = new WorkbookDoc(wb, path);
        if (path == null) doc.UntitledName = Loc.T("Sổ làm việc", "Book") + _untitled++;
        Docs.Add(doc);
        Activate(doc);
    }

    public void Activate(WorkbookDoc doc)
    {
        if (Grid.IsEditing) Grid.CommitEdit(0, 0);
        _doc = doc;
        Grid.Doc = doc;
        Tabs.Book = doc.Book;
        RebuildDocTabs();
        UpdateAll();
    }

    private void RebuildDocTabs()
    {
        _docTabs.Items.Clear();
        _docTabs.Visible = Docs.Count > 1;
        _docTabs.Renderer = new RibbonRenderer(header: true);
        _docTabs.BackColor = Theme.TabStripBack;
        foreach (var d in Docs)
        {
            var doc = d;
            var b = new ToolStripButton(d.DisplayName + (d.Modified ? " *" : "")) { Checked = d == _doc, Name = "docTab" + Docs.IndexOf(d), ToolTipText = d.FilePath ?? d.DisplayName };
            b.Click += (_, _) => Activate(doc);
            _docTabs.Items.Add(b);
            var x = new ToolStripButton("✕") { Name = "docTabClose" + Docs.IndexOf(d), ToolTipText = Loc.T("Đóng", "Close"), ForeColor = Theme.SubtleInk };
            x.Click += (_, _) => CloseDocument(doc);
            _docTabs.Items.Add(x);
        }
    }

    public IReadOnlyList<string> OpenDocumentPaths => Docs.Where(d => d.FilePath != null).Select(d => d.FilePath!).ToList();

    public void OpenFiles(IReadOnlyList<string> paths, bool fromShell)
    {
        foreach (var p in paths) OpenFile(p);
        if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
    }

    // ------------------------------------------------------------------ UI state sync

    public void UpdateAll()
    {
        if (_doc == null) return;
        Tabs.Book = _doc.Book;
        Tabs.Reveal();
        Grid.RefreshAll();
        UpdateCellInfo();
        UpdateStats();
        UpdateTitle();
        UpdateContextualTabs();
        SyncZoom();
        UpdateMode();
        _printPage?.RefreshPreview();
    }

    private void OnContentChanged()
    {
        _doc.Modified = true;
        Tabs.Invalidate();
        UpdateCellInfo();
        UpdateStats();
        UpdateTitle();
        UpdateContextualTabs();
        Grid.Invalidate();
        if (_docTabs.Visible) RebuildDocTabs();
        if (_doc.Book.Calc.Circular.Count > 0)
        {
            var c = _doc.Book.Calc.Circular[0];
            _statusMode.Text = Loc.T("Tham chiếu vòng: ", "Circular references: ") + Addr.ToA1(c.Row, c.Col);
        }
    }

    private void UpdateTitle()
    {
        Text = _doc.DisplayName + (_doc.Modified ? " •" : "") + " - " + Product.AppName(AppKind.Excel);
        if (Ribbon.Controls.Count > 0) { }
        if (_controls.TryGetValue("qatUndo", out _)) { }
    }

    private void UpdateMode()
    {
        _statusMode.Text = Grid.IsEditing ? (Grid.CanInsertRef() ? Loc.T("Trỏ", "Point") : Loc.T("Nhập", "Enter")) : Loc.T("Sẵn sàng", "Ready");
        if (_doc.Book.ManualCalc && !Grid.IsEditing) _statusMode.Text += "   " + Loc.T("Tính toán thủ công", "Calculate (manual)");
    }

    private void SyncZoom()
    {
        _updating = true;
        int z = (int)Math.Round(_doc.Sheet.Zoom * 100);
        _zoom.Value = Math.Clamp(z, _zoom.Minimum, _zoom.Maximum);
        _zoomLabel.Text = z + "%";
        _updating = false;
    }

    private void SyncFormulaBar()
    {
        if (_formula.Focused && Grid.IsEditing) return;
        _updating = true;
        if (Grid.IsEditing) _formula.Text = Grid.EditText.Replace("\n", "\r\n").Replace("\r\r\n", "\r\n");
        else
        {
            var a = Grid.Active;
            var cell = _doc.Sheet.GetCell(a.Row, a.Col);
            var st = _doc.Sheet.StyleAt(a.Row, a.Col);
            // a value spilled by a dynamic-array formula shows that formula greyed out, like Excel
            var src = _doc.Sheet.SpillSource(a.Row, a.Col);
            _formula.ForeColor = src != null ? SystemColors.GrayText : SystemColors.WindowText;
            if (src is var (_, _, master) && !(_doc.Sheet.Protected && st.HiddenFormula)) _formula.Text = "=" + master.Formula;
            else _formula.Text = _doc.Sheet.Protected && st.HiddenFormula ? "" : GridView.EditableText(cell, st);
        }
        _updating = false;
    }

    public void UpdateCellInfo()
    {
        if (_doc == null) return;
        _updating = true;
        try
        {
            var s = _doc.Sheet;
            var sel = Grid.Selection;
            var a = Grid.Active;
            var name = _doc.Book.Names.FirstOrDefault(n => Addr.TryParseRange(n.Value.Contains('!') ? n.Value[(n.Value.LastIndexOf('!') + 1)..] : n.Value, out var nr) && nr == sel && (!n.Value.Contains('!') || n.Value.Contains(s.Name))).Key;
            _nameBox.Text = name ?? (sel.IsSingle || s.MergeAt(a.Row, a.Col) == sel ? Addr.ToA1(a.Row, a.Col) : Addr.RangeToString(sel).Length > 20 ? Addr.ToA1(a.Row, a.Col) : Addr.ToA1(a.Row, a.Col));
            SyncFormulaBar();
            var st = s.StyleAt(a.Row, a.Col);
            Check("btnBold", st.Bold); Check("btnItalic", st.Italic); Check("btnUnderline", st.Underline); Check("btnStrike", st.Strike);
            Check("btnWrap", st.Wrap);
            Check("btnAlignLeft", st.HAlign == HAlign.Left); Check("btnAlignCenter", st.HAlign == HAlign.Center); Check("btnAlignRight", st.HAlign == HAlign.Right);
            Check("btnAlignTop", st.VAlign == VAlign.Top); Check("btnAlignMiddle", st.VAlign == VAlign.Center); Check("btnAlignBottom", st.VAlign == VAlign.Bottom);
            Check("btnMerge", s.MergeAt(a.Row, a.Col) != null);
            Check("btnFilter", s.Filter != null);
            Check("btnShowFormulas", Grid.ShowFormulas);
            Check("btnViewGridlines", s.ShowGridlines); Check("btnViewHeadings", s.ShowHeadings);
            Check("btnPrintGridlines", s.Page.PrintGridlines); Check("btnPrintHeadings", s.Page.PrintHeadings);
            Check("btnFitWidth", s.Page.FitToWidth);
            Check("btnViewFormulaBar", _formulaPanel.Visible);
            Check("btnProtectSheet", s.Protected);
            Check("btnFormatPainter", _painter != null);
            if (!_fontBox.Focused) _fontBox.Text = st.FontName;
            if (!_sizeBox.Focused) _sizeBox.Text = st.FontSize.ToString("0.#");
            if (!_numFmtBox.Focused) _numFmtBox.Text = NumberFormatName(st.NumberFormat);
            if (_controls.TryGetValue("btnFontColor", out var fc) && fc is RibbonButton fcb) { }
            SetEnabled("qatUndo", _doc.CanUndo); SetEnabled("qatRedo", _doc.CanRedo);
            SetEnabled("btnEditNote", s.Notes.ContainsKey(Sheet.Key(a.Row, a.Col)));
            SetEnabled("btnDeleteNote", s.Notes.ContainsKey(Sheet.Key(a.Row, a.Col)));
            SetEnabled("btnClearFilter", s.Filter != null && (s.Filter.Allowed.Count > 0 || s.Filter.Conditions.Count > 0));
            SetEnabled("btnReapplyFilter", s.Filter != null);
        }
        finally { _updating = false; }
    }

    private void Check(string name, bool v) { if (_controls.TryGetValue(name, out var c) && c is RibbonButton b) b.Checked = v; }
    private void SetEnabled(string name, bool v)
    {
        if (_controls.TryGetValue(name, out var c)) c.Enabled = v;
        if (_qat.TryGetValue(name, out var q)) q.Enabled = v;
    }

    private readonly Dictionary<string, ToolStripButton> _qat = new();

    public static string NumberFormatName(string code)
    {
        if (code == "General") return Loc.T("Chung", "General");
        if (code == "@") return Loc.T("Văn bản", "Text");
        if (NumFmt.IsDateFormat(code)) return code.Contains('h') && !code.Contains('d') ? Loc.T("Giờ", "Time") : code.Contains("mmmm") || code.Contains("dddd") ? Loc.T("Ngày dài", "Long Date") : Loc.T("Ngày ngắn", "Short Date");
        if (code.Contains('%')) return Loc.T("Phần trăm", "Percentage");
        if (code.Contains("E+")) return Loc.T("Khoa học", "Scientific");
        if (code.Contains('/')) return Loc.T("Phân số", "Fraction");
        if (code.Contains("₫") || code.Contains('$')) return code.Contains('*') ? Loc.T("Kế toán", "Accounting") : Loc.T("Tiền tệ", "Currency");
        if (code.Contains('0')) return Loc.T("Số", "Number");
        return Loc.T("Tùy chỉnh", "Custom");
    }

    public void UpdateStats()
    {
        var s = _doc.Sheet;
        double sum = 0; int count = 0, numCount = 0;
        int scanned = 0;
        foreach (var r in Grid.AllSelections)
        {
            if (r.IsSingle && Grid.ExtraRanges.Count == 0) break;
            foreach (var (k, cell) in s.AllCells)
            {
                if (++scanned > 2_000_000) break;
                var (row, col) = Sheet.FromKey(k);
                if (!r.Contains(row, col) || s.IsRowHidden(row)) continue;
                var v = cell.Value;
                if (v.IsBlank) continue;
                count++;
                if (v.IsNumber) { sum += v.Num; numCount++; }
            }
        }
        if (count <= 1 && Grid.ExtraRanges.Count == 0) _statusStats.Text = "";
        else
        {
            var sumText = NumFmt.FormatNumber(sum, "#,##0.##########", out _).TrimEnd(NumFmt.DecimalSep);
            _statusStats.Text = numCount > 0
                ? Loc.F("Trung bình: {0}    Số ô có dữ liệu: {1}    Tổng: {2}", "Average: {0}    Count: {1}    Sum: {2}", NumFmt.General(sum / numCount), count, sumText)
                : Loc.F("Số ô có dữ liệu: {0}", "Count: {0}", count);
        }
        LayoutStatus();
    }

    private void UpdateContextualTabs()
    {
        Ribbon.SetContextual(_chartPage, Grid.SelectedObject is Chart, Grid.SelectedObject is Chart);
        Ribbon.SetContextual(_picturePage, Grid.SelectedObject is Picture, Grid.SelectedObject is Picture);
        if (Grid.SelectedObject is Chart ch) { Check("btnChartLegend", ch.ShowLegend); Check("btnChartDataLabels", ch.ShowDataLabels); }
    }

    // ------------------------------------------------------------------ sheets

    public void SwitchSheet(int i)
    {
        if (i < 0 || i >= _doc.Book.Sheets.Count) return;
        if (Grid.IsEditing && !Grid.CanInsertRef()) Grid.CommitEdit(0, 0);
        _doc.Book.ActiveSheet = i;
        Grid.ExtraRanges.Clear();
        Grid.SelectObject(null);
        Grid.RefreshAll();
        Tabs.Reveal();
        UpdateCellInfo();
        UpdateStats();
        SyncZoom();
    }

    private void MoveSheet(int from, int to)
    {
        _doc.Checkpoint(Loc.T("Di chuyển trang tính", "Move sheet"));
        var s = _doc.Book.Sheets[from];
        _doc.Book.Sheets.RemoveAt(from);
        _doc.Book.Sheets.Insert(to, s);
        _doc.Book.ActiveSheet = to;
        _doc.Book.MarkDirty();
        OnContentChanged();
        Grid.RefreshAll();
    }

    // ------------------------------------------------------------------ backstage

    private void BuildBackstage()
    {
        Ribbon.FileClicked += (_, _) => ShowBackstage("home");
        Ribbon.CommandInvoking += (_, _) => HideBackstage();
        BackstageView.BackRequested += (_, _) => HideBackstage();
        _start = new StartScreen(AppKind.Excel, ExcelTemplates.All, t => { HideBackstage(); CmdNewFromTemplate(t.Id); }, p => { HideBackstage(); OpenFile(p); }, () => { CmdOpen(); }, compactTemplates: false);
        BackstageView.AddPage("home", Loc.T("Trang chủ", "Home"), Icons.Home, _start);
        var newPage = new StartScreen(AppKind.Excel, ExcelTemplates.All, t => { HideBackstage(); CmdNewFromTemplate(t.Id); }, p => { HideBackstage(); OpenFile(p); }, () => CmdOpen(), compactTemplates: false) { Name = "newScreen" };
        BackstageView.AddPage("new", Loc.T("Mới", "New"), Icons.NewDoc, newPage);
        BackstageView.AddCommand("open", Loc.T("Mở", "Open"), Icons.Open, CmdOpen);
        BackstageView.AddPage("info", Loc.T("Thông tin", "Info"), Icons.Info, BuildInfoPage());
        BackstageView.AddCommand("save", Loc.T("Lưu", "Save"), Icons.Save, () => { if (CmdSave()) HideBackstage(); });
        BackstageView.AddCommand("saveAs", Loc.T("Lưu thành", "Save As"), Icons.SaveAs, () => { if (CmdSaveAs()) HideBackstage(); });
        _printPage = new PrintPage(() => this);
        BackstageView.AddPage("print", Loc.T("In", "Print"), Icons.Print, _printPage);
        BackstageView.AddPage("export", Loc.T("Xuất", "Export"), Icons.Export, BuildExportPage());
        BackstageView.AddCommand("close", Loc.T("Đóng", "Close"), Icons.Close, () => { HideBackstage(); CloseDocument(_doc); });
        BackstageView.AddCommand("about", Loc.T("Giới thiệu", "About"), Icons.Info, () => OfficeShell.ShowAbout(this), bottom: true);
        BackstageView.AddCommand("options", Loc.T("Tùy chọn", "Options"), Icons.Settings, CmdOptions, bottom: true);
        BackstageView.PageShown += (_, p) =>
        {
            if (p == "print") _printPage?.RefreshPreview();
            if (p == "info") RefreshInfoPage();
            if (p is "home" or "new") _start?.RefreshRecent();
        };
    }

    private readonly Label _infoText = new() { Name = "lblInfo", AutoSize = true, Font = new Font("Segoe UI", 10f) };
    private readonly TextBox _infoTitle = new() { Name = "txtDocTitle", Width = 300 }, _infoAuthor = new() { Name = "txtDocAuthor", Width = 300 };

    private Control BuildInfoPage()
    {
        var p = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Name = "infoPage" };
        p.Controls.Add(new Label { Text = Loc.T("Thông tin", "Info"), Font = new Font("Segoe UI Light", 20f), AutoSize = true });
        p.Controls.Add(_infoText);
        p.Controls.Add(new Label { Text = Loc.T("Tiêu đề", "Title"), AutoSize = true, Font = Theme.UiFontBold, Margin = new Padding(0, Dpi.Scale(14), 0, 0) });
        p.Controls.Add(_infoTitle);
        p.Controls.Add(new Label { Text = Loc.T("Tác giả", "Author"), AutoSize = true, Font = Theme.UiFontBold });
        p.Controls.Add(_infoAuthor);
        var apply = new Button { Name = "btnInfoApply", Text = Loc.T("Cập nhật thuộc tính", "Update properties"), AutoSize = true, FlatStyle = FlatStyle.Flat };
        apply.Click += (_, _) => { _doc.Checkpoint(Loc.T("Thuộc tính", "Properties")); _doc.Book.Title = _infoTitle.Text; _doc.Book.Author = _infoAuthor.Text; OnContentChanged(); };
        p.Controls.Add(apply);
        var protect = new Button { Name = "btnInfoProtect", Text = Loc.T("Bảo vệ trang tính...", "Protect sheet..."), AutoSize = true, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, Dpi.Scale(14), 0, 0) };
        protect.Click += (_, _) => { HideBackstage(); CmdProtectSheet(); };
        p.Controls.Add(protect);
        return p;
    }

    private void RefreshInfoPage()
    {
        var b = _doc.Book;
        int cells = b.Sheets.Sum(s => s.CellCount), formulas = b.Sheets.Sum(s => s.AllCells.Count(c => c.Value.HasFormula));
        _infoText.Text = (_doc.FilePath ?? Loc.T("(chưa lưu)", "(not saved yet)")) + "\n" +
            Loc.F("Trang tính: {0}    Ô có dữ liệu: {1}    Công thức: {2}    Biểu đồ: {3}", "Sheets: {0}    Cells: {1}    Formulas: {2}    Charts: {3}", b.Sheets.Count, cells, formulas, b.Sheets.Sum(s => s.Charts.Count)) +
            (_doc.FilePath != null && File.Exists(_doc.FilePath) ? "\n" + Loc.T("Sửa lần cuối: ", "Last modified: ") + File.GetLastWriteTime(_doc.FilePath).ToString("dd/MM/yyyy HH:mm") + "    " + (new FileInfo(_doc.FilePath).Length / 1024.0).ToString("0.0") + " KB" : "");
        _infoTitle.Text = b.Title;
        _infoAuthor.Text = b.Author;
    }

    private Control BuildExportPage()
    {
        var p = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, Name = "exportPage" };
        p.Controls.Add(new Label { Text = Loc.T("Xuất", "Export"), Font = new Font("Segoe UI Light", 20f), AutoSize = true });
        Button E(string name, string text, string desc, string icon, Action a)
        {
            var b = new Button
            {
                Name = name, Text = "   " + text + "\n   " + desc, Image = XlIcons.Get(icon, Dpi.Scale(32)), TextImageRelation = TextImageRelation.ImageBeforeText, ImageAlign = ContentAlignment.MiddleLeft,
                TextAlign = ContentAlignment.MiddleLeft, Size = new Size(Dpi.Scale(460), Dpi.Scale(62)), FlatStyle = FlatStyle.Flat, Margin = new Padding(0, Dpi.Scale(6), 0, 0)
            };
            b.FlatAppearance.BorderColor = Theme.Border;
            b.FlatAppearance.MouseOverBackColor = Theme.Hover;
            b.Click += (_, _) => a();
            p.Controls.Add(b);
            return b;
        }
        E("btnExportPdf", Loc.T("Tạo tài liệu PDF", "Create PDF document"), Loc.T("Giữ nguyên bố cục, dễ chia sẻ và in", "Preserves layout, easy to share and print"), "printArea", CmdExportPdf);
        E("btnExportCsv", Loc.T("Lưu CSV (UTF-8)", "Save as CSV (UTF-8)"), Loc.T("Trang tính hiện hoạt, mở được bằng mọi phần mềm", "Active sheet, readable by any program"), "textToCols", () => CmdExportDelimited(','));
        E("btnExportTsv", Loc.T("Lưu văn bản (Tab)", "Save as Text (Tab delimited)"), Loc.T("Trang tính hiện hoạt, phân cách bằng Tab", "Active sheet, tab separated"), "textToCols", () => CmdExportDelimited('\t'));
        E("btnExportHtml", Loc.T("Lưu trang web (HTML)", "Save as Web Page (HTML)"), Loc.T("Toàn bộ sổ làm việc thành bảng HTML", "The whole workbook as HTML tables"), "link", CmdExportHtml);
        return p;
    }

    public void ShowBackstage(string page)
    {
        if (Grid.IsEditing) Grid.CommitEdit(0, 0);
        BackstageView.Visible = true;
        BackstageView.BringToFront();
        BackstageView.ShowPage(page);
    }

    public void HideBackstage()
    {
        BackstageView.Visible = false;
        Grid.Focus();
    }

    public bool BackstageVisible => BackstageView.Visible;

    // ------------------------------------------------------------------ settings

    public static void ApplyNumberStyle()
    {
        var style = OfficeConfig.GetAppData("excel.numberStyle") ?? "auto";
        NumFmt.UseLocale(style == "vi" || (style == "auto" && Loc.IsVi));
    }

    private void RestoreWindow()
    {
        var b = OfficeConfig.GetAppData("excel.window");
        if (b == null) return;
        var p = b.Split(',').Select(x => int.TryParse(x, out var v) ? v : 0).ToArray();
        if (p.Length != 5) return;
        var r = new Rectangle(p[0], p[1], p[2], p[3]);
        if (Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(r)) && r.Width > 400 && r.Height > 300) { StartPosition = FormStartPosition.Manual; Bounds = r; }
        if (p[4] == (int)FormWindowState.Maximized) WindowState = FormWindowState.Maximized;
    }

    private void SaveWindow()
    {
        var r = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        try { OfficeConfig.SetAppData("excel.window", $"{r.X},{r.Y},{r.Width},{r.Height},{(int)WindowState}"); } catch { }
    }

    private sealed class FormatPainterState
    {
        public ClipData Clip = null!;
        public bool Sticky;
    }
}
