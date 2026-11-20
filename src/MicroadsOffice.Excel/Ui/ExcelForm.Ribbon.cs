using System.Drawing.Text;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

public sealed partial class ExcelForm
{
    private RibbonButton B(string name, string vi, string en, Image? img, string tipVi, string tipEn, Action a, RibbonButtonSize size = RibbonButtonSize.Auto)
    {
        var b = RibbonButton.Create(name, Loc.T(vi, en), img, Loc.T(tipVi, tipEn), () => RunCommand(name, a), size);
        _controls[name] = b;
        return b;
    }

    private RibbonButton IconB(string name, Image img, string tipVi, string tipEn, Action a) => B(name, "", "", img, tipVi, tipEn, a, RibbonButtonSize.Icon);

    private ToolStripMenuItem Mi(string name, string vi, string en, Image? img, Action a, Keys keys = Keys.None)
    {
        var it = new ToolStripMenuItem(Loc.T(vi, en), img, (_, _) => RunCommand(name, a)) { Name = name, ShortcutKeys = keys, ShowShortcutKeys = keys != Keys.None };
        _menuItems.Add(it);
        return it;
    }

    private ContextMenuStrip Menu(params ToolStripItem[] items)
    {
        var m = new ContextMenuStrip { Renderer = new RibbonRenderer(false), ImageScalingSize = new Size(Dpi.Scale(16), Dpi.Scale(16)) };
        m.Items.AddRange(items);
        return m;
    }

    private void RunCommand(string name, Action a)
    {
        try
        {
            if (Grid.IsEditing && name is not ("btnFormulaFx" or "btnInsertFunction" or "btnInsertFn") && !name.StartsWith("miFn") && !name.StartsWith("miName") && !name.StartsWith("miSymbol"))
                if (!Grid.CommitEdit(0, 0)) return;
            a();
        }
        catch (Exception ex)
        {
            Ux.Error(this, Loc.T("Không thực hiện được lệnh: ", "The command could not be completed: ") + ex.Message);
        }
        finally
        {
            UpdateCellInfo();
            UpdateTitle();
        }
    }

    private void BuildRibbon()
    {
        Ribbon.AddHeaderRight(OfficeShell.CreateHeaderItems(this));
        _qat["qatSave"] = Ribbon.AddQuickAccess("qatSave", Icons.Small(Icons.Save), Loc.T("Lưu (Ctrl+S)", "Save (Ctrl+S)"), () => RunCommand("qatSave", () => CmdSave()));
        _qat["qatUndo"] = Ribbon.AddQuickAccess("qatUndo", Icons.Small(Icons.Undo), Loc.T("Hoàn tác (Ctrl+Z)", "Undo (Ctrl+Z)"), () => RunCommand("qatUndo", CmdUndo));
        _qat["qatRedo"] = Ribbon.AddQuickAccess("qatRedo", Icons.Small(Icons.Redo), Loc.T("Làm lại (Ctrl+Y)", "Redo (Ctrl+Y)"), () => RunCommand("qatRedo", CmdRedo));

        // =================================================================== HOME
        var home = Ribbon.AddPage("Home", Loc.T("Trang chủ", "Home"));
        var g = home.AddGroup("Clipboard", Loc.T("Bộ nhớ tạm", "Clipboard"));
        var paste = B("btnPaste", "Dán", "Paste", Icons.Large(Icons.Paste), "Dán (Ctrl+V)", "Paste (Ctrl+V)", () => Grid.Paste());
        paste.Split = true;
        paste.DropDown = Menu(
            Mi("miPasteAll", "Dán", "Paste", Icons.Small(Icons.Paste), () => Grid.Paste()),
            Mi("miPasteValues", "Dán giá trị", "Paste Values", null, () => Grid.Paste(PasteMode.Values)),
            Mi("miPasteFormulas", "Dán công thức", "Paste Formulas", null, () => Grid.Paste(PasteMode.Formulas)),
            Mi("miPasteFormats", "Dán định dạng", "Paste Formatting", XlIcons.Small("formatPainter"), () => Grid.Paste(PasteMode.Formats)),
            Mi("miPasteTranspose", "Dán chuyển vị", "Transpose", XlIcons.Small("switchRowCol"), () => Grid.Paste(PasteMode.All, transpose: true)),
            new ToolStripSeparator(),
            Mi("miPasteSpecial", "Dán đặc biệt...", "Paste Special...", null, CmdPasteSpecial, Keys.Control | Keys.Alt | Keys.V));
        g.Large(paste);
        g.Stack(B("btnCut", "Cắt", "Cut", Icons.Small(Icons.Cut), "Cắt (Ctrl+X)", "Cut (Ctrl+X)", () => Grid.CopySelection(true)),
                B("btnCopy", "Sao chép", "Copy", Icons.Small(Icons.Copy), "Sao chép (Ctrl+C)", "Copy (Ctrl+C)", () => Grid.CopySelection(false)),
                B("btnFormatPainter", "Sao định dạng", "Format Painter", XlIcons.Small("formatPainter"), "Sao định dạng: chọn ô nguồn, bấm nút rồi chọn ô đích", "Format Painter: copy formatting to other cells", CmdFormatPainter));

        g = home.AddGroup("Font", Loc.T("Phông chữ", "Font"));
        _fontBox = new ComboBox { Name = "cboFontName", Width = Dpi.Scale(150), AccessibleName = Loc.T("Phông chữ", "Font"), AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems };
        using (var fonts = new InstalledFontCollection()) foreach (var f in fonts.Families) _fontBox.Items.Add(f.Name);
        _sizeBox = new ComboBox { Name = "cboFontSize", Width = Dpi.Scale(50), AccessibleName = Loc.T("Cỡ chữ", "Font size") };
        foreach (var s in new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }) _sizeBox.Items.Add(s.ToString());
        _fontBox.SelectionChangeCommitted += (_, _) => BeginInvoke(() => SetFontName(_fontBox.Text));
        _fontBox.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SetFontName(_fontBox.Text); } };
        _sizeBox.SelectionChangeCommitted += (_, _) => BeginInvoke(() => SetFontSize(_sizeBox.Text));
        _sizeBox.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SetFontSize(_sizeBox.Text); } };
        _controls["cboFontName"] = _fontBox; _controls["cboFontSize"] = _sizeBox;
        g.Row(_fontBox, _sizeBox,
            IconB("btnGrowFont", XlIcons.Small("growFont"), "Tăng cỡ chữ", "Increase Font Size", () => Restyle(s => s with { FontSize = NextSize(s.FontSize, 1) }, "Font size")),
            IconB("btnShrinkFont", XlIcons.Small("shrinkFont"), "Giảm cỡ chữ", "Decrease Font Size", () => Restyle(s => s with { FontSize = NextSize(s.FontSize, -1) }, "Font size")));
        var borders = IconB("btnBorders", XlIcons.Small("borders"), "Viền ô", "Borders", () => ApplyBorders("bottom"));
        borders.Split = true;
        borders.DropDown = Menu(
            Mi("miBorderBottom", "Viền dưới", "Bottom Border", XlIcons.Small("borderBottom"), () => ApplyBorders("bottom")),
            Mi("miBorderTop", "Viền trên", "Top Border", XlIcons.Small("borderTop"), () => ApplyBorders("top")),
            Mi("miBorderLeft", "Viền trái", "Left Border", XlIcons.Small("borderLeft"), () => ApplyBorders("left")),
            Mi("miBorderRight", "Viền phải", "Right Border", XlIcons.Small("borderRight"), () => ApplyBorders("right")),
            new ToolStripSeparator(),
            Mi("miBorderNone", "Không viền", "No Border", XlIcons.Small("borderNone"), () => ApplyBorders("none")),
            Mi("miBorderAll", "Tất cả viền", "All Borders", XlIcons.Small("borderAll"), () => ApplyBorders("all")),
            Mi("miBorderOutside", "Viền ngoài", "Outside Borders", XlIcons.Small("borderOutside"), () => ApplyBorders("outline")),
            Mi("miBorderThick", "Viền ngoài dày", "Thick Outside Borders", XlIcons.Small("borderThick"), () => ApplyBorders("thick")),
            Mi("miBorderBottomDouble", "Viền dưới kép", "Bottom Double Border", XlIcons.Small("borderBottom"), () => ApplyBorders("double")),
            new ToolStripSeparator(),
            Mi("miBorderMore", "Viền khác...", "More Borders...", null, () => CmdFormatCells(3)));
        var fill = IconB("btnFillColor", XlIcons.Small("fill"), "Màu nền ô", "Fill Color", () => ApplyFill(_lastFill));
        fill.Split = true;
        fill.SwatchColor = Color.FromArgb(_lastFill == 0 ? Color.Yellow.ToArgb() : _lastFill);
        fill.DropDownFactory = () => new ColorPickerDropDown(c => { _lastFill = c?.ToArgb() ?? 0; fill.SwatchColor = c ?? Color.Transparent; ApplyFill(_lastFill); }, Loc.T("Không tô", "No Fill"));
        var fcol = IconB("btnFontColor", XlIcons.Small("fontColor"), "Màu chữ", "Font Color", () => Restyle(s => s with { FontColor = _lastFontColor }, "Font colour"));
        fcol.Split = true;
        fcol.SwatchColor = Color.FromArgb(_lastFontColor);
        fcol.DropDownFactory = () => new ColorPickerDropDown(c => { _lastFontColor = c?.ToArgb() ?? 0; fcol.SwatchColor = c ?? Color.Black; Restyle(s => s with { FontColor = _lastFontColor }, "Font colour"); }, Loc.T("Tự động", "Automatic"));
        g.Row(IconB("btnBold", Icons.Small(Icons.Bold), "Đậm (Ctrl+B)", "Bold (Ctrl+B)", () => ToggleStyle(s => s.Bold, (s, v) => s with { Bold = v })),
              IconB("btnItalic", Icons.Small(Icons.Italic), "Nghiêng (Ctrl+I)", "Italic (Ctrl+I)", () => ToggleStyle(s => s.Italic, (s, v) => s with { Italic = v })),
              IconB("btnUnderline", Icons.Small(Icons.Underline), "Gạch chân (Ctrl+U)", "Underline (Ctrl+U)", () => ToggleStyle(s => s.Underline, (s, v) => s with { Underline = v })),
              IconB("btnStrike", XlIcons.Small("strike"), "Gạch ngang (Ctrl+5)", "Strikethrough (Ctrl+5)", () => ToggleStyle(s => s.Strike, (s, v) => s with { Strike = v })),
              borders, fill, fcol);
        g.SetDialogLauncher(Loc.T("Định dạng ô: Phông chữ", "Format Cells: Font"), () => CmdFormatCells(2));

        g = home.AddGroup("Alignment", Loc.T("Căn lề", "Alignment"));
        var merge = IconB("btnMerge", XlIcons.Small("merge"), "Gộp và căn giữa", "Merge & Center", () => CmdMerge(true, false));
        merge.Split = true;
        merge.DropDown = Menu(
            Mi("miMergeCenter", "Gộp và căn giữa", "Merge & Center", XlIcons.Small("merge"), () => CmdMerge(true, false)),
            Mi("miMergeAcross", "Gộp theo hàng", "Merge Across", null, () => CmdMerge(false, true)),
            Mi("miMergeCells", "Gộp ô", "Merge Cells", null, () => CmdMerge(false, false)),
            Mi("miUnmerge", "Bỏ gộp ô", "Unmerge Cells", null, CmdUnmerge));
        g.Row(IconB("btnAlignTop", XlIcons.Small("alignTop"), "Căn trên", "Top Align", () => Restyle(s => s with { VAlign = VAlign.Top }, "Align")),
              IconB("btnAlignMiddle", XlIcons.Small("alignMiddle"), "Căn giữa theo chiều dọc", "Middle Align", () => Restyle(s => s with { VAlign = VAlign.Center }, "Align")),
              IconB("btnAlignBottom", XlIcons.Small("alignBottom"), "Căn dưới", "Bottom Align", () => Restyle(s => s with { VAlign = VAlign.Bottom }, "Align")),
              B("btnWrap", "Xuống dòng", "Wrap Text", XlIcons.Small("wrap"), "Tự động xuống dòng trong ô", "Wrap Text", () => ToggleStyle(s => s.Wrap, (s, v) => s with { Wrap = v }), RibbonButtonSize.Small));
        g.Row(IconB("btnAlignLeft", XlIcons.Small("alignLeft"), "Căn trái", "Align Left", () => ToggleAlign(HAlign.Left)),
              IconB("btnAlignCenter", XlIcons.Small("alignCenter"), "Căn giữa", "Center", () => ToggleAlign(HAlign.Center)),
              IconB("btnAlignRight", XlIcons.Small("alignRight"), "Căn phải", "Align Right", () => ToggleAlign(HAlign.Right)),
              IconB("btnIndentDec", XlIcons.Small("indentDec"), "Giảm thụt lề", "Decrease Indent", () => Restyle(s => s with { Indent = Math.Max(0, s.Indent - 1) }, "Indent")),
              IconB("btnIndentInc", XlIcons.Small("indentInc"), "Tăng thụt lề", "Increase Indent", () => Restyle(s => s with { Indent = Math.Min(15, s.Indent + 1), HAlign = s.HAlign is HAlign.General or HAlign.Center ? HAlign.Left : s.HAlign }, "Indent")),
              merge);
        g.SetDialogLauncher(Loc.T("Định dạng ô: Căn lề", "Format Cells: Alignment"), () => CmdFormatCells(1));

        g = home.AddGroup("Number", Loc.T("Số", "Number"));
        _numFmtBox = new ComboBox { Name = "cboNumberFormat", Width = Dpi.Scale(130), DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = Loc.T("Định dạng số", "Number Format") };
        var fmtChoices = new (string Vi, string En, string Code)[]
        {
            ("Chung", "General", "General"), ("Số", "Number", "#,##0.00"), ("Tiền tệ", "Currency", "#,##0 \"₫\""), ("Kế toán", "Accounting", "_-* #,##0 \"₫\"_-;-* #,##0 \"₫\"_-;_-* \"-\"??_-;_-@_-"),
            ("Ngày ngắn", "Short Date", "dd/mm/yyyy"), ("Ngày dài", "Long Date", "dddd, dd mmmm yyyy"), ("Giờ", "Time", "hh:mm:ss"), ("Phần trăm", "Percentage", "0.00%"),
            ("Phân số", "Fraction", "# ?/?"), ("Khoa học", "Scientific", "0.00E+00"), ("Văn bản", "Text", "@")
        };
        foreach (var f in fmtChoices) _numFmtBox.Items.Add(Loc.T(f.Vi, f.En));
        _numFmtBox.Items.Add(Loc.T("Tùy chỉnh", "Custom"));
        _numFmtBox.SelectionChangeCommitted += (_, _) =>
        {
            int i = _numFmtBox.SelectedIndex;
            BeginInvoke(() =>
            {
                if (i >= 0 && i < fmtChoices.Length) SetNumberFormat(i == 4 ? Format.NumFmt.ShortDateCode : fmtChoices[i].Code);
                else CmdFormatCells(0);
                Grid.Focus();
            });
        };
        _controls["cboNumberFormat"] = _numFmtBox;
        g.Row(_numFmtBox);
        var cur = IconB("btnCurrency", XlIcons.Small("currency"), "Định dạng tiền tệ", "Accounting Number Format", () => SetNumberFormat("#,##0 \"₫\""));
        cur.Split = true;
        cur.DropDown = Menu(
            Mi("miCurVnd", "₫ Việt Nam đồng", "₫ Vietnamese dong", null, () => SetNumberFormat("#,##0 \"₫\"")),
            Mi("miCurUsd", "$ Đô la Mỹ", "$ US dollar", null, () => SetNumberFormat("\"$\"#,##0.00")),
            Mi("miCurEur", "€ Euro", "€ Euro", null, () => SetNumberFormat("#,##0.00 \"€\"")),
            Mi("miCurMore", "Định dạng khác...", "More formats...", null, () => CmdFormatCells(0)));
        g.Row(cur,
              IconB("btnPercent", XlIcons.Small("percent"), "Kiểu phần trăm (Ctrl+Shift+%)", "Percent Style (Ctrl+Shift+%)", () => SetNumberFormat("0%")),
              IconB("btnComma", XlIcons.Small("comma"), "Dấu phân cách hàng nghìn", "Comma Style", () => SetNumberFormat("#,##0")),
              IconB("btnDecInc", XlIcons.Small("decInc"), "Tăng chữ số thập phân", "Increase Decimal", () => AdjustDecimals(1)),
              IconB("btnDecDec", XlIcons.Small("decDec"), "Giảm chữ số thập phân", "Decrease Decimal", () => AdjustDecimals(-1)));
        g.SetDialogLauncher(Loc.T("Định dạng ô: Số", "Format Cells: Number"), () => CmdFormatCells(0));

        g = home.AddGroup("Styles", Loc.T("Kiểu", "Styles"));
        var cf = B("btnCondFormat", "Định dạng\ncó điều kiện", "Conditional\nFormatting", XlIcons.Large("condFormat"), "Tô màu ô theo điều kiện", "Highlight cells based on rules", () => Grid.Focus(), RibbonButtonSize.Large);
        cf.DropDown = Menu(
            Mi("miCfGreater", "Lớn hơn...", "Greater Than...", null, () => CmdCfRule(CfType.GreaterThan)),
            Mi("miCfLess", "Nhỏ hơn...", "Less Than...", null, () => CmdCfRule(CfType.LessThan)),
            Mi("miCfBetween", "Nằm giữa...", "Between...", null, () => CmdCfRule(CfType.Between)),
            Mi("miCfEqual", "Bằng...", "Equal To...", null, () => CmdCfRule(CfType.EqualTo)),
            Mi("miCfText", "Có chứa chữ...", "Text that Contains...", null, () => CmdCfRule(CfType.TextContains)),
            Mi("miCfDuplicate", "Giá trị trùng lặp", "Duplicate Values", null, () => AddCf(new ConditionalFormat { Type = CfType.Duplicate, Fill = CfRuleDialog.Presets[0].Fill, FontColor = CfRuleDialog.Presets[0].Font })),
            new ToolStripSeparator(),
            Mi("miCfTop10", "10 giá trị lớn nhất...", "Top 10 Items...", null, () => CmdCfRule(CfType.Top10)),
            Mi("miCfBottom10", "10 giá trị nhỏ nhất...", "Bottom 10 Items...", null, () => CmdCfRule(CfType.Bottom10)),
            Mi("miCfAboveAvg", "Trên trung bình", "Above Average", null, () => AddCf(new ConditionalFormat { Type = CfType.AboveAverage, Fill = CfRuleDialog.Presets[2].Fill, FontColor = CfRuleDialog.Presets[2].Font })),
            Mi("miCfBelowAvg", "Dưới trung bình", "Below Average", null, () => AddCf(new ConditionalFormat { Type = CfType.BelowAverage, Fill = CfRuleDialog.Presets[0].Fill, FontColor = CfRuleDialog.Presets[0].Font })),
            new ToolStripSeparator(),
            Mi("miCfDataBar", "Thanh dữ liệu", "Data Bars", null, () => AddCf(new ConditionalFormat { Type = CfType.DataBar })),
            Mi("miCfScale3", "Thang màu (xanh-vàng-đỏ)", "Color Scale (green-yellow-red)", null, () => AddCf(new ConditionalFormat { Type = CfType.ColorScale3, Color1 = Color.FromArgb(248, 105, 107).ToArgb(), Color2 = Color.FromArgb(255, 235, 132).ToArgb(), Color3 = Color.FromArgb(99, 190, 123).ToArgb() })),
            Mi("miCfScale2", "Thang màu (trắng-xanh)", "Color Scale (white-green)", null, () => AddCf(new ConditionalFormat { Type = CfType.ColorScale2, Color1 = Color.White.ToArgb(), Color3 = Color.FromArgb(99, 190, 123).ToArgb() })),
            Mi("miCfFormula", "Dùng công thức...", "Use a Formula...", null, () => CmdCfRule(CfType.Expression)),
            new ToolStripSeparator(),
            Mi("miCfClearSel", "Xóa quy tắc khỏi ô đã chọn", "Clear Rules from Selected Cells", null, () => ClearCf(false)),
            Mi("miCfClearSheet", "Xóa quy tắc khỏi trang tính", "Clear Rules from Entire Sheet", null, () => ClearCf(true)));
        g.Large(cf);
        var tbl = B("btnFormatAsTable", "Định dạng\nnhư bảng", "Format as\nTable", XlIcons.Large("table"), "Định dạng vùng dữ liệu thành bảng có sọc và bộ lọc", "Format as a banded table with filter buttons", () => CmdFormatAsTable(0), RibbonButtonSize.Large);
        tbl.DropDown = Menu(
            Mi("miTableGreen", "Xanh lá", "Green", null, () => CmdFormatAsTable(0)),
            Mi("miTableBlue", "Xanh dương", "Blue", null, () => CmdFormatAsTable(1)),
            Mi("miTableOrange", "Cam", "Orange", null, () => CmdFormatAsTable(2)),
            Mi("miTableGray", "Xám", "Gray", null, () => CmdFormatAsTable(3)));
        g.Large(tbl);
        var cs = B("btnCellStyles", "Kiểu ô", "Cell\nStyles", XlIcons.Large("cellStyles"), "Kiểu ô có sẵn", "Cell styles", () => { }, RibbonButtonSize.Large);
        cs.DropDown = Menu(
            Mi("miStyleNormal", "Bình thường", "Normal", null, () => Restyle(_ => CellStyle.Default, "Cell style")),
            Mi("miStyleGood", "Tốt", "Good", null, () => Restyle(s => s with { Fill = Color.FromArgb(198, 239, 206).ToArgb(), FontColor = Color.FromArgb(0, 97, 0).ToArgb() }, "Cell style")),
            Mi("miStyleBad", "Xấu", "Bad", null, () => Restyle(s => s with { Fill = Color.FromArgb(255, 199, 206).ToArgb(), FontColor = Color.FromArgb(156, 0, 6).ToArgb() }, "Cell style")),
            Mi("miStyleNeutral", "Trung tính", "Neutral", null, () => Restyle(s => s with { Fill = Color.FromArgb(255, 235, 156).ToArgb(), FontColor = Color.FromArgb(156, 87, 0).ToArgb() }, "Cell style")),
            Mi("miStyleHeading", "Tiêu đề", "Heading 1", null, () => Restyle(s => s with { Bold = true, FontSize = 15, FontColor = Color.FromArgb(31, 78, 121).ToArgb(), Bottom = new BorderLine(LineStyle.Thick, Color.FromArgb(68, 114, 196).ToArgb()) }, "Cell style")),
            Mi("miStyleTitle", "Tựa đề", "Title", null, () => Restyle(s => s with { Bold = true, FontSize = 18, FontColor = Color.FromArgb(31, 78, 121).ToArgb() }, "Cell style")),
            Mi("miStyleTotal", "Tổng", "Total", null, () => Restyle(s => s with { Bold = true, Top = BorderLine.Thin(Color.FromArgb(68, 114, 196)), Bottom = new BorderLine(LineStyle.Double, Color.FromArgb(68, 114, 196).ToArgb()) }, "Cell style")),
            Mi("miStyleNote", "Ghi chú", "Note", null, () => Restyle(s => s with { Fill = Color.FromArgb(255, 255, 204).ToArgb(), Left = BorderLine.Thin(Color.Gray), Right = BorderLine.Thin(Color.Gray), Top = BorderLine.Thin(Color.Gray), Bottom = BorderLine.Thin(Color.Gray) }, "Cell style")));
        g.Large(cs);

        g = home.AddGroup("Cells", Loc.T("Ô", "Cells"));
        var ins = B("btnInsertCells", "Chèn", "Insert", XlIcons.Large("insertCells"), "Chèn hàng, cột, ô hoặc trang tính", "Insert rows, columns, cells or sheets", () => CmdInsertRows(), RibbonButtonSize.Large);
        ins.Split = true;
        ins.DropDown = Menu(
            Mi("miInsertRows", "Chèn hàng trang tính", "Insert Sheet Rows", XlIcons.Small("insertRow"), () => CmdInsertRows()),
            Mi("miInsertCols", "Chèn cột trang tính", "Insert Sheet Columns", XlIcons.Small("insertCol"), () => CmdInsertCols()),
            Mi("miInsertSheet", "Chèn trang tính", "Insert Sheet", XlIcons.Small("newSheetTab"), CmdInsertSheet, Keys.Shift | Keys.F11));
        g.Large(ins);
        var del = B("btnDeleteCells", "Xóa", "Delete", XlIcons.Large("deleteCells"), "Xóa hàng, cột hoặc trang tính", "Delete rows, columns or sheets", () => CmdDeleteRows(), RibbonButtonSize.Large);
        del.Split = true;
        del.DropDown = Menu(
            Mi("miDeleteRows", "Xóa hàng trang tính", "Delete Sheet Rows", XlIcons.Small("deleteRow"), () => CmdDeleteRows()),
            Mi("miDeleteCols", "Xóa cột trang tính", "Delete Sheet Columns", XlIcons.Small("deleteCol"), () => CmdDeleteCols()),
            Mi("miDeleteSheet", "Xóa trang tính", "Delete Sheet", null, CmdDeleteSheet));
        g.Large(del);
        var fmt = B("btnFormatCellsMenu", "Định dạng", "Format", XlIcons.Large("format"), "Chiều cao hàng, độ rộng cột, ẩn/hiện, trang tính", "Row height, column width, hide/unhide, sheets", () => { }, RibbonButtonSize.Large);
        fmt.DropDown = Menu(
            Mi("miRowHeight", "Chiều cao hàng...", "Row Height...", XlIcons.Small("rowHeight"), CmdRowHeight),
            Mi("miAutoFitRow", "Tự vừa chiều cao hàng", "AutoFit Row Height", null, () => Grid.AutoFitRows(SelRows())),
            Mi("miColWidth", "Độ rộng cột...", "Column Width...", XlIcons.Small("colWidth"), CmdColWidth),
            Mi("miAutoFitCol", "Tự vừa độ rộng cột", "AutoFit Column Width", null, () => Grid.AutoFitColumns(SelCols())),
            new ToolStripSeparator(),
            Mi("miHideRows", "Ẩn hàng", "Hide Rows", null, () => SetHidden(true, true)),
            Mi("miHideCols", "Ẩn cột", "Hide Columns", null, () => SetHidden(false, true)),
            Mi("miUnhideRows", "Hiện hàng", "Unhide Rows", null, () => SetHidden(true, false)),
            Mi("miUnhideCols", "Hiện cột", "Unhide Columns", null, () => SetHidden(false, false)),
            new ToolStripSeparator(),
            Mi("miRenameSheet", "Đổi tên trang tính", "Rename Sheet", null, CmdRenameSheet),
            Mi("miTabColor", "Màu thẻ...", "Tab Color...", null, CmdTabColor),
            Mi("miProtectSheet2", "Bảo vệ trang tính...", "Protect Sheet...", XlIcons.Small("protect"), CmdProtectSheet),
            new ToolStripSeparator(),
            Mi("miFormatCells", "Định dạng ô...", "Format Cells...", null, () => CmdFormatCells(0), Keys.Control | Keys.D1));
        g.Large(fmt);

        g = home.AddGroup("Editing", Loc.T("Chỉnh sửa", "Editing"));
        var sum = B("btnAutoSum", "Tự động tính", "AutoSum", XlIcons.Small("sum"), "Tự động tính tổng (Alt+=)", "AutoSum (Alt+=)", () => CmdAutoSum("SUM"), RibbonButtonSize.Small);
        sum.Split = true;
        sum.DropDown = Menu(
            Mi("miSum", "Tổng", "Sum", null, () => CmdAutoSum("SUM")), Mi("miAverage", "Trung bình", "Average", null, () => CmdAutoSum("AVERAGE")),
            Mi("miCount", "Đếm số", "Count Numbers", null, () => CmdAutoSum("COUNT")), Mi("miMax", "Lớn nhất", "Max", null, () => CmdAutoSum("MAX")),
            Mi("miMin", "Nhỏ nhất", "Min", null, () => CmdAutoSum("MIN")), new ToolStripSeparator(), Mi("miMoreFunctions", "Hàm khác...", "More Functions...", null, CmdInsertFunction));
        var fillBtn = B("btnFill", "Điền", "Fill", XlIcons.Small("fillDown"), "Điền xuống, sang phải...", "Fill down, right...", () => CmdFillDown(), RibbonButtonSize.Small);
        fillBtn.DropDown = Menu(
            Mi("miFillDown", "Xuống", "Down", null, () => CmdFillDown(), Keys.Control | Keys.D),
            Mi("miFillRight", "Sang phải", "Right", null, CmdFillRight, Keys.Control | Keys.R),
            Mi("miFillUp", "Lên", "Up", null, CmdFillUp),
            Mi("miFillLeft", "Sang trái", "Left", null, CmdFillLeft),
            Mi("miFillSeries", "Chuỗi (tự nhận dạng)...", "Series (auto)...", null, CmdFillSeries));
        var clear = B("btnClear", "Xóa", "Clear", XlIcons.Small("clear"), "Xóa nội dung, định dạng...", "Clear contents, formats...", () => Grid.ClearSelection(true, true), RibbonButtonSize.Small);
        clear.DropDown = Menu(
            Mi("miClearAll", "Xóa tất cả", "Clear All", null, () => Grid.ClearSelection(true, true)),
            Mi("miClearFormats", "Xóa định dạng", "Clear Formats", null, () => Grid.ClearSelection(false, true)),
            Mi("miClearContents", "Xóa nội dung", "Clear Contents", null, () => Grid.ClearSelection(true, false)),
            Mi("miClearNotes", "Xóa ghi chú", "Clear Notes", null, CmdDeleteNote),
            Mi("miClearLinks", "Xóa siêu liên kết", "Clear Hyperlinks", null, CmdRemoveHyperlink));
        g.Stack(sum, fillBtn, clear);
        var sf = B("btnSortFilter", "Sắp xếp\nvà lọc", "Sort &&\nFilter", XlIcons.Large("sortFilter"), "Sắp xếp và lọc dữ liệu", "Sort and filter", () => { }, RibbonButtonSize.Large);
        sf.DropDown = Menu(
            Mi("miSortAsc", "Sắp xếp A → Z", "Sort A to Z", XlIcons.Small("sortAsc"), () => CmdSort(true)),
            Mi("miSortDesc", "Sắp xếp Z → A", "Sort Z to A", XlIcons.Small("sortDesc"), () => CmdSort(false)),
            Mi("miSortCustom", "Sắp xếp tùy chỉnh...", "Custom Sort...", XlIcons.Small("sortCustom"), CmdCustomSort),
            new ToolStripSeparator(),
            Mi("miFilter", "Bộ lọc", "Filter", XlIcons.Small("filter"), CmdToggleFilter, Keys.Control | Keys.Shift | Keys.L),
            Mi("miClearFilter", "Xóa lọc", "Clear", XlIcons.Small("clearFilter"), CmdClearFilter),
            Mi("miReapply", "Áp dụng lại", "Reapply", null, CmdReapplyFilter));
        g.Large(sf);
        var fs = B("btnFindSelect", "Tìm và\nchọn", "Find &&\nSelect", XlIcons.Large("find"), "Tìm, thay thế, đi tới", "Find, replace, go to", () => CmdFind(false), RibbonButtonSize.Large);
        fs.DropDown = Menu(
            Mi("miFind", "Tìm...", "Find...", XlIcons.Small("find"), () => CmdFind(false), Keys.Control | Keys.F),
            Mi("miReplace", "Thay thế...", "Replace...", null, () => CmdFind(true), Keys.Control | Keys.H),
            Mi("miGoTo", "Đi tới...", "Go To...", null, CmdGoTo, Keys.Control | Keys.G),
            new ToolStripSeparator(),
            Mi("miSelectFormulas", "Chọn ô có công thức", "Formulas", null, () => SelectWhere(c => c.Formula != null)),
            Mi("miSelectNotes", "Chọn ô có ghi chú", "Notes", null, () => SelectWhereKeys(_doc.Sheet.Notes.Keys)),
            Mi("miSelectBlanks", "Chọn ô trống trong vùng", "Blanks in selection", null, CmdSelectBlanks));
        g.Large(fs);

        g = home.AddGroup("Ai", Loc.T("Trợ lý AI", "AI"));
        g.Large(B("btnAiAssistant", "Trợ lý\nAI", "AI\nAssistant", MicroadsOffice.AI.UI.AiIcons.Sparkle(true),
            "Mở Trợ lý AI (ChatGPT, Claude…): viết công thức, phân tích, tạo bảng, vẽ biểu đồ ngay trong trang tính",
            "Open the AI assistant (ChatGPT, Claude…): write formulas, analyze, build tables and charts right in the sheet", ToggleAi, RibbonButtonSize.Large));
        g.CollapsedIcon = MicroadsOffice.AI.UI.AiIcons.Sparkle(true);

        // =================================================================== INSERT
        var insert = Ribbon.AddPage("Insert", Loc.T("Chèn", "Insert"));
        g = insert.AddGroup("Tables", Loc.T("Bảng", "Tables"));
        g.Large(B("btnInsertTable", "Bảng", "Table", XlIcons.Large("table"), "Tạo bảng có sọc màu và bộ lọc", "Create a table with banded rows and filters", () => CmdFormatAsTable(0), RibbonButtonSize.Large));
        g = insert.AddGroup("Illustrations", Loc.T("Minh họa", "Illustrations"));
        g.Large(B("btnInsertPicture", "Ảnh", "Pictures", XlIcons.Large("picture"), "Chèn ảnh từ máy tính", "Insert a picture from this device", CmdInsertPicture, RibbonButtonSize.Large));
        g = insert.AddGroup("Charts", Loc.T("Biểu đồ", "Charts"));
        g.Large(B("btnInsertChart", "Biểu đồ\nđề xuất", "Recommended\nCharts", XlIcons.Large("colChart"), "Chọn loại biểu đồ phù hợp với dữ liệu", "Pick a chart for your data", () => CmdInsertChart(null), RibbonButtonSize.Large));
        g.Row(IconB("btnChartColumn", XlIcons.Small("colChart"), "Biểu đồ cột", "Column chart", () => CmdInsertChart(ChartType.Column)),
              IconB("btnChartLine", XlIcons.Small("lineChart"), "Biểu đồ đường", "Line chart", () => CmdInsertChart(ChartType.Line)),
              IconB("btnChartPie", XlIcons.Small("pieChart"), "Biểu đồ tròn", "Pie chart", () => CmdInsertChart(ChartType.Pie)));
        g.Row(IconB("btnChartBar", XlIcons.Small("barChart"), "Biểu đồ thanh", "Bar chart", () => CmdInsertChart(ChartType.Bar)),
              IconB("btnChartArea", XlIcons.Small("areaChart"), "Biểu đồ vùng", "Area chart", () => CmdInsertChart(ChartType.Area)),
              IconB("btnChartScatter", XlIcons.Small("scatterChart"), "Biểu đồ phân tán", "Scatter chart", () => CmdInsertChart(ChartType.Scatter)));
        g.Row(IconB("btnChartStacked", XlIcons.Small("stackChart"), "Biểu đồ cột chồng", "Stacked column chart", () => CmdInsertChart(ChartType.StackedColumn)),
              IconB("btnChartDoughnut", XlIcons.Small("doughnutChart"), "Biểu đồ vành khuyên", "Doughnut chart", () => CmdInsertChart(ChartType.Doughnut)));
        g = insert.AddGroup("Links", Loc.T("Liên kết", "Links"));
        g.Large(B("btnHyperlink", "Liên kết", "Link", XlIcons.Large("link"), "Chèn siêu liên kết (Ctrl+K)", "Insert a hyperlink (Ctrl+K)", CmdHyperlink, RibbonButtonSize.Large));
        g = insert.AddGroup("Comments", Loc.T("Ghi chú", "Notes"));
        g.Large(B("btnInsertNote", "Ghi chú", "Note", XlIcons.Large("note"), "Thêm ghi chú cho ô (Shift+F2)", "Add a note to the cell (Shift+F2)", CmdNewNote, RibbonButtonSize.Large));
        g = insert.AddGroup("Symbols", Loc.T("Ký hiệu", "Symbols"));
        var sym = B("btnSymbol", "Ký hiệu", "Symbol", XlIcons.Large("symbol"), "Chèn ký hiệu đặc biệt", "Insert a special character", () => { }, RibbonButtonSize.Large);
        var symItems = new List<ToolStripItem>();
        foreach (var ch in new[] { "₫", "€", "$", "©", "®", "™", "°", "±", "×", "÷", "≤", "≥", "≠", "√", "∑", "π", "Ω", "µ", "→", "←", "↑", "↓", "✓", "✗", "•", "…" })
        {
            var c = ch;
            symItems.Add(Mi("miSymbol_" + ((int)ch[0]).ToString("X"), c, c, null, () => InsertText(c)));
        }
        sym.DropDown = Menu(symItems.ToArray());
        g.Large(sym);
        g.Large(B("btnInsertFunction", "Hàm", "Function", XlIcons.Large("function"), "Chèn hàm (Shift+F3)", "Insert Function (Shift+F3)", CmdInsertFunction, RibbonButtonSize.Large));

        // =================================================================== PAGE LAYOUT
        var layout = Ribbon.AddPage("PageLayout", Loc.T("Bố trí trang", "Page Layout"));
        g = layout.AddGroup("PageSetup", Loc.T("Thiết lập trang", "Page Setup"));
        var margins = B("btnMargins", "Lề", "Margins", XlIcons.Large("margins"), "Lề trang", "Page margins", () => { }, RibbonButtonSize.Large);
        margins.DropDown = Menu(
            Mi("miMarginNormal", "Thường (1,9 / 1,8 cm)", "Normal", null, () => SetMargins(0.75, 0.75, 0.7, 0.7)),
            Mi("miMarginWide", "Rộng (2,5 cm)", "Wide", null, () => SetMargins(1, 1, 1, 1)),
            Mi("miMarginNarrow", "Hẹp (1,9 / 0,6 cm)", "Narrow", null, () => SetMargins(0.75, 0.75, 0.25, 0.25)),
            Mi("miMarginCustom", "Tùy chỉnh...", "Custom Margins...", null, CmdPageSetup));
        g.Large(margins);
        var orient = B("btnOrientation", "Hướng", "Orientation", XlIcons.Large("orientation"), "Hướng giấy", "Page orientation", () => { }, RibbonButtonSize.Large);
        orient.DropDown = Menu(Mi("miPortrait", "Dọc", "Portrait", null, () => SetOrientation(false)), Mi("miLandscape", "Ngang", "Landscape", null, () => SetOrientation(true)));
        g.Large(orient);
        var size = B("btnPaperSize", "Khổ giấy", "Size", XlIcons.Large("paperSize"), "Khổ giấy", "Paper size", () => { }, RibbonButtonSize.Large);
        size.DropDown = Menu(PageSetupDialog.Papers.Select(p => (ToolStripItem)Mi("miPaper" + p.Name, p.Name, p.Name, null, () => SetPaper(p.Id))).ToArray());
        g.Large(size);
        var area = B("btnPrintArea", "Vùng in", "Print\nArea", XlIcons.Large("printArea"), "Chỉ in vùng đã chọn", "Print only the selected area", CmdSetPrintArea, RibbonButtonSize.Large);
        area.DropDown = Menu(Mi("miSetPrintArea", "Đặt vùng in", "Set Print Area", null, CmdSetPrintArea), Mi("miClearPrintArea", "Xóa vùng in", "Clear Print Area", null, CmdClearPrintArea));
        g.Large(area);
        g.Large(B("btnPrintTitles", "In tiêu đề", "Print\nTitles", XlIcons.Large("printTitles"), "Lặp lại các hàng đã chọn ở đầu mỗi trang in", "Repeat the selected rows at the top of every page", CmdPrintTitles, RibbonButtonSize.Large));
        g.SetDialogLauncher(Loc.T("Thiết lập trang", "Page Setup"), CmdPageSetup);
        g = layout.AddGroup("ScaleToFit", Loc.T("Co giãn vừa trang", "Scale to Fit"));
        g.Stack(B("btnFitWidth", "Vừa 1 trang ngang", "Fit to 1 page wide", XlIcons.Small("fitWidth"), "Co tất cả các cột vừa một trang", "Fit all columns on one page", CmdToggleFitWidth, RibbonButtonSize.Small),
                B("btnScale", "Tỉ lệ in...", "Scale...", XlIcons.Small("scale"), "Tỉ lệ in", "Print scale", CmdPrintScale, RibbonButtonSize.Small));
        g = layout.AddGroup("SheetOptions", Loc.T("Tùy chọn trang tính", "Sheet Options"));
        g.Stack(B("btnViewGridlines", "Xem đường lưới", "View Gridlines", XlIcons.Small("gridlines"), "Hiện đường lưới trên màn hình", "Show gridlines on screen", CmdToggleGridlines, RibbonButtonSize.Small),
                B("btnPrintGridlines", "In đường lưới", "Print Gridlines", XlIcons.Small("gridlines"), "In đường lưới", "Print gridlines", () => { _doc.Checkpoint("Print gridlines"); _doc.Sheet.Page.PrintGridlines = !_doc.Sheet.Page.PrintGridlines; OnContentChanged(); }, RibbonButtonSize.Small));
        g.Stack(B("btnViewHeadings", "Xem tiêu đề", "View Headings", XlIcons.Small("headings"), "Hiện tiêu đề hàng/cột", "Show row and column headings", CmdToggleHeadings, RibbonButtonSize.Small),
                B("btnPrintHeadings", "In tiêu đề", "Print Headings", XlIcons.Small("headings"), "In tiêu đề hàng/cột", "Print row and column headings", () => { _doc.Checkpoint("Print headings"); _doc.Sheet.Page.PrintHeadings = !_doc.Sheet.Page.PrintHeadings; OnContentChanged(); }, RibbonButtonSize.Small));

        // =================================================================== FORMULAS
        var formulas = Ribbon.AddPage("Formulas", Loc.T("Công thức", "Formulas"));
        g = formulas.AddGroup("FunctionLibrary", Loc.T("Thư viện hàm", "Function Library"));
        g.Large(B("btnInsertFn", "Chèn\nhàm", "Insert\nFunction", XlIcons.Large("function"), "Chèn hàm (Shift+F3)", "Insert Function (Shift+F3)", CmdInsertFunction, RibbonButtonSize.Large));
        var sum2 = B("btnAutoSum2", "Tự động\ntính", "AutoSum", XlIcons.Large("sum"), "Tự động tính tổng (Alt+=)", "AutoSum (Alt+=)", () => CmdAutoSum("SUM"), RibbonButtonSize.Large);
        g.Large(sum2);
        RibbonButton FnMenu(string name, FnCategory cat, string icon, string vi, string en)
        {
            var b = B(name, vi, en, XlIcons.Small(icon), vi, en, () => { }, RibbonButtonSize.Small);
            b.DropDown = Menu(FunctionCatalog.All.Where(f => f.Category == cat).OrderBy(f => f.Name).Select(f => (ToolStripItem)new ToolStripMenuItem(f.Name, null, (_, _) => InsertFunctionCall(f.Name)) { Name = "miFn" + f.Name, ToolTipText = f.Signature + "\n" + f.Description }).ToArray());
            foreach (ToolStripMenuItem it in b.DropDown.Items) _menuItems.Add(it);
            return b;
        }
        g.Stack(FnMenu("btnFnFinancial", FnCategory.Financial, "currency", "Tài chính", "Financial"), FnMenu("btnFnLogical", FnCategory.Logical, "validation", "Logic", "Logical"), FnMenu("btnFnText", FnCategory.Text, "fontColor", "Văn bản", "Text"));
        g.Stack(FnMenu("btnFnDate", FnCategory.DateTime, "calc", "Ngày & giờ", "Date & Time"), FnMenu("btnFnLookup", FnCategory.Lookup, "find", "Tra cứu", "Lookup & Reference"), FnMenu("btnFnMath", FnCategory.Math, "sum", "Toán học", "Math & Trig"));
        RibbonButton MoreFnMenu()
        {
            var b = B("btnFnMore", "Hàm khác", "More Functions", XlIcons.Small("calc"), "Hàm khác: Kỹ thuật, Cơ sở dữ liệu, Tương thích, Web", "More functions: Engineering, Database, Compatibility, Web", () => { }, RibbonButtonSize.Small);
            var subs = new List<ToolStripItem>();
            foreach (var (cat, vi, en) in new[] { (FnCategory.Engineering, "Kỹ thuật", "Engineering"), (FnCategory.Database, "Cơ sở dữ liệu", "Database"), (FnCategory.Compatibility, "Tương thích", "Compatibility"), (FnCategory.Web, "Web", "Web") })
            {
                var fns = FunctionCatalog.All.Where(f => f.Category == cat).OrderBy(f => f.Name).ToList();
                if (fns.Count == 0) continue;
                var sub = new ToolStripMenuItem(Loc.T(vi, en)) { Name = "miFnCat" + cat };
                foreach (var f in fns)
                {
                    var it = new ToolStripMenuItem(f.Name, null, (_, _) => InsertFunctionCall(f.Name)) { Name = "miFn" + f.Name, ToolTipText = f.Signature + "\n" + f.Description };
                    sub.DropDownItems.Add(it);
                    _menuItems.Add(it);
                }
                subs.Add(sub);
            }
            b.DropDown = Menu(subs.ToArray());
            return b;
        }
        g.Stack(FnMenu("btnFnStat", FnCategory.Statistical, "colChart", "Thống kê", "Statistical"), FnMenu("btnFnInfo", FnCategory.Information, "note", "Thông tin", "Information"), MoreFnMenu());
        g = formulas.AddGroup("DefinedNames", Loc.T("Tên đã định nghĩa", "Defined Names"));
        g.Large(B("btnNameManager", "Quản lý\ntên", "Name\nManager", XlIcons.Large("names"), "Tạo, sửa, xóa tên", "Create, edit and delete names", CmdNameManager, RibbonButtonSize.Large));
        var useName = B("btnUseInFormula", "Dùng trong công thức", "Use in Formula", XlIcons.Small("names"), "Chèn tên vào công thức", "Insert a name into the formula", () => { }, RibbonButtonSize.Small);
        useName.DropDownOpening += (_, _) =>
        {
            useName.DropDown!.Items.Clear();
            foreach (var n in _doc.Book.Names.Keys.OrderBy(x => x)) { var nn = n; useName.DropDown.Items.Add(new ToolStripMenuItem(nn, null, (_, _) => InsertText(nn)) { Name = "miName" + nn }); }
            if (useName.DropDown.Items.Count == 0) useName.DropDown.Items.Add(new ToolStripMenuItem(Loc.T("(chưa có tên)", "(no names)")) { Enabled = false });
        };
        useName.DropDown = Menu();
        g.Stack(B("btnDefineName", "Định nghĩa tên", "Define Name", XlIcons.Small("names"), "Đặt tên cho vùng đã chọn", "Name the selected range", CmdDefineName, RibbonButtonSize.Small), useName);
        g = formulas.AddGroup("Auditing", Loc.T("Kiểm tra công thức", "Formula Auditing"));
        g.Large(B("btnShowFormulas", "Hiện\ncông thức", "Show\nFormulas", XlIcons.Large("showFormulas"), "Hiện công thức thay vì kết quả (Ctrl+`)", "Show formulas instead of results (Ctrl+`)", CmdToggleShowFormulas, RibbonButtonSize.Large));
        g = formulas.AddGroup("Calculation", Loc.T("Tính toán", "Calculation"));
        var calcOpt = B("btnCalcOptions", "Tùy chọn\ntính toán", "Calculation\nOptions", XlIcons.Large("calc"), "Tự động / thủ công", "Automatic / manual", () => { }, RibbonButtonSize.Large);
        calcOpt.DropDown = Menu(Mi("miCalcAuto", "Tự động", "Automatic", null, () => SetManualCalc(false)), Mi("miCalcManual", "Thủ công", "Manual", null, () => SetManualCalc(true)));
        g.Large(calcOpt);
        g.Stack(B("btnCalcNow", "Tính ngay", "Calculate Now", XlIcons.Small("calc"), "Tính lại toàn bộ (F9)", "Calculate the workbook (F9)", CmdCalcNow, RibbonButtonSize.Small));

        // =================================================================== DATA
        var data = Ribbon.AddPage("Data", Loc.T("Dữ liệu", "Data"));
        g = data.AddGroup("GetData", Loc.T("Lấy dữ liệu", "Get Data"));
        g.Large(B("btnImportCsv", "Từ tệp\nText/CSV", "From\nText/CSV", XlIcons.Large("textToCols"), "Nhập dữ liệu CSV vào trang tính mới", "Import a CSV file into a new sheet", CmdImportCsv, RibbonButtonSize.Large));
        g = data.AddGroup("SortFilter", Loc.T("Sắp xếp và lọc", "Sort & Filter"));
        g.Stack(IconB("btnSortAsc", XlIcons.Small("sortAsc"), "Sắp xếp A → Z", "Sort A to Z", () => CmdSort(true)), IconB("btnSortDesc", XlIcons.Small("sortDesc"), "Sắp xếp Z → A", "Sort Z to A", () => CmdSort(false)));
        g.Large(B("btnSortCustom", "Sắp xếp", "Sort", XlIcons.Large("sortCustom"), "Sắp xếp theo nhiều cột", "Sort by several columns", CmdCustomSort, RibbonButtonSize.Large));
        g.Large(B("btnFilter", "Lọc", "Filter", XlIcons.Large("filter"), "Bật/tắt bộ lọc (Ctrl+Shift+L)", "Turn filtering on/off (Ctrl+Shift+L)", CmdToggleFilter, RibbonButtonSize.Large));
        g.Stack(B("btnClearFilter", "Xóa lọc", "Clear", XlIcons.Small("clearFilter"), "Xóa điều kiện lọc", "Clear the filter", CmdClearFilter, RibbonButtonSize.Small),
                B("btnReapplyFilter", "Áp dụng lại", "Reapply", XlIcons.Small("filter"), "Áp dụng lại bộ lọc", "Reapply the filter", CmdReapplyFilter, RibbonButtonSize.Small));
        g = data.AddGroup("DataTools", Loc.T("Công cụ dữ liệu", "Data Tools"));
        g.Large(B("btnTextToColumns", "Văn bản\nthành cột", "Text to\nColumns", XlIcons.Large("textToCols"), "Tách nội dung một cột thành nhiều cột", "Split one column into several", CmdTextToColumns, RibbonButtonSize.Large));
        g.Large(B("btnRemoveDuplicates", "Loại bỏ\ntrùng lặp", "Remove\nDuplicates", XlIcons.Large("removeDup"), "Xóa các hàng trùng lặp", "Delete duplicate rows", CmdRemoveDuplicates, RibbonButtonSize.Large));
        var dv = B("btnDataValidation", "Xác thực\ndữ liệu", "Data\nValidation", XlIcons.Large("validation"), "Giới hạn dữ liệu nhập vào ô (danh sách thả xuống...)", "Restrict what can be entered (drop-down lists...)", CmdDataValidation, RibbonButtonSize.Large);
        dv.Split = true;
        dv.DropDown = Menu(Mi("miDataValidation", "Xác thực dữ liệu...", "Data Validation...", null, CmdDataValidation), Mi("miClearValidation", "Xóa xác thực", "Clear Validation", null, CmdClearValidation));
        g.Large(dv);

        // =================================================================== REVIEW
        var review = Ribbon.AddPage("Review", Loc.T("Xem lại", "Review"));
        g = review.AddGroup("Notes", Loc.T("Ghi chú", "Notes"));
        g.Large(B("btnNewNote", "Ghi chú\nmới", "New\nNote", XlIcons.Large("note"), "Thêm ghi chú (Shift+F2)", "Add a note (Shift+F2)", CmdNewNote, RibbonButtonSize.Large));
        g.Stack(B("btnEditNote", "Sửa ghi chú", "Edit Note", XlIcons.Small("note"), "Sửa ghi chú của ô", "Edit the note", CmdNewNote, RibbonButtonSize.Small),
                B("btnDeleteNote", "Xóa ghi chú", "Delete Note", XlIcons.Small("deleteObj"), "Xóa ghi chú của ô đã chọn", "Delete notes in the selection", CmdDeleteNote, RibbonButtonSize.Small),
                B("btnNextNote", "Ghi chú tiếp theo", "Next Note", XlIcons.Small("note"), "Đi tới ghi chú tiếp theo", "Go to the next note", CmdNextNote, RibbonButtonSize.Small));
        g = review.AddGroup("Protect", Loc.T("Bảo vệ", "Protect"));
        g.Large(B("btnProtectSheet", "Bảo vệ\ntrang tính", "Protect\nSheet", XlIcons.Large("protect"), "Khóa các ô để không bị sửa", "Prevent changes to locked cells", CmdProtectSheet, RibbonButtonSize.Large));
        g.Large(B("btnLockCell", "Khóa ô", "Lock\nCell", XlIcons.Large("protect"), "Bật/tắt thuộc tính khóa của ô đã chọn", "Toggle the Locked property of the selection", () => ToggleStyle(s => s.Locked, (s, v) => s with { Locked = v }), RibbonButtonSize.Large));

        // =================================================================== VIEW
        var view = Ribbon.AddPage("View", Loc.T("Xem", "View"));
        g = view.AddGroup("Show", Loc.T("Hiện", "Show"));
        g.Stack(B("btnViewFormulaBar", "Thanh công thức", "Formula Bar", XlIcons.Small("function"), "Hiện/ẩn thanh công thức", "Show/hide the formula bar", () => { _formulaPanel.Visible = !_formulaPanel.Visible; }, RibbonButtonSize.Small),
                B("btnViewGridlines2", "Đường lưới", "Gridlines", XlIcons.Small("gridlines"), "Hiện/ẩn đường lưới", "Show/hide gridlines", CmdToggleGridlines, RibbonButtonSize.Small),
                B("btnViewHeadings2", "Tiêu đề", "Headings", XlIcons.Small("headings"), "Hiện/ẩn tiêu đề hàng, cột", "Show/hide headings", CmdToggleHeadings, RibbonButtonSize.Small));
        g = view.AddGroup("Zoom", Loc.T("Thu phóng", "Zoom"));
        g.Large(B("btnZoom", "Thu phóng", "Zoom", Icons.Large(Icons.ZoomIn), "Chọn tỉ lệ thu phóng", "Choose a zoom level", CmdZoomDialog, RibbonButtonSize.Large));
        g.Large(B("btnZoom100", "100%", "100%", Icons.Large("1:1"), "Thu phóng 100%", "Zoom to 100%", () => Grid.SetZoom(1), RibbonButtonSize.Large));
        g.Large(B("btnZoomSelection", "Vừa vùng\nchọn", "Zoom to\nSelection", XlIcons.Large("zoomSel"), "Phóng to vừa vùng đang chọn", "Zoom so the selection fills the window", CmdZoomToSelection, RibbonButtonSize.Large));
        g = view.AddGroup("Window", Loc.T("Cửa sổ", "Window"));
        var freeze = B("btnFreeze", "Cố định\nô", "Freeze\nPanes", XlIcons.Large("freeze"), "Giữ hàng/cột luôn hiện khi cuộn", "Keep rows/columns visible while scrolling", CmdFreezePanes, RibbonButtonSize.Large);
        freeze.DropDown = Menu(
            Mi("miFreezePanes", "Cố định tại ô đang chọn", "Freeze Panes", null, CmdFreezePanes),
            Mi("miFreezeTopRow", "Cố định hàng đầu", "Freeze Top Row", null, () => SetFreeze(1, 0)),
            Mi("miFreezeFirstCol", "Cố định cột đầu", "Freeze First Column", null, () => SetFreeze(0, 1)),
            Mi("miUnfreeze", "Bỏ cố định", "Unfreeze Panes", null, () => SetFreeze(0, 0)));
        g.Large(freeze);
        g.Large(B("btnNewWorkbook", "Sổ mới", "New\nWorkbook", Icons.Large(Icons.NewDoc), "Tạo sổ làm việc mới (Ctrl+N)", "New workbook (Ctrl+N)", () => CmdNewFromTemplate("blank"), RibbonButtonSize.Large));

        // =================================================================== HELP
        var help = Ribbon.AddPage("Help", Loc.T("Trợ giúp", "Help"));
        g = help.AddGroup("Help", Loc.T("Trợ giúp", "Help"));
        g.Large(B("btnHelp", "Hướng dẫn", "Help", Icons.Large(Icons.Help), "Hướng dẫn sử dụng và phím tắt (F1)", "How to use and keyboard shortcuts (F1)", CmdHelp, RibbonButtonSize.Large));
        g.Large(B("btnCheckUpdate", "Kiểm tra\ncập nhật", "Check for\nUpdates", Icons.Large(Icons.Sync), "Kiểm tra phiên bản mới tại microads.vn/office", "Check microads.vn/office for a new version", () => OfficeShell.CheckForUpdatesInteractive(this), RibbonButtonSize.Large));
        g.Large(B("btnSettings", "Cài đặt và\nngôn ngữ", "Settings &&\nLanguage", Icons.Large(Icons.Settings), "Ngôn ngữ, cập nhật tự động", "Language, automatic updates", () => OfficeShell.ShowSettings(this), RibbonButtonSize.Large));
        g.Large(B("btnExcelOptions", "Tùy chọn\nExcel", "Excel\nOptions", XlIcons.Large("options"), "Kiểu hiển thị số, tính toán, tự lưu", "Number display, calculation, AutoRecover", CmdOptions, RibbonButtonSize.Large));
        g = help.AddGroup("About", "Microads");
        g.Large(B("btnAbout", "Giới thiệu", "About", Icons.Large(Icons.Info), "Giới thiệu Microads Office", "About Microads Office", () => OfficeShell.ShowAbout(this), RibbonButtonSize.Large));
        g.Large(B("btnVisitWebsite", "microads.vn", "microads.vn", Icons.Large(Icons.Globe), "Giải pháp phần mềm doanh nghiệp, bảo mật, CRM, ERP, thiết kế website", "Enterprise software, security, CRM, ERP, web design", () => OfficeShell.OpenWebsite(), RibbonButtonSize.Large));

        // =================================================================== CONTEXTUAL: CHART / PICTURE
        _chartPage = Ribbon.AddPage("ChartDesign", Loc.T("Biểu đồ", "Chart Design"), contextual: true);
        g = _chartPage.AddGroup("ChartType", Loc.T("Loại", "Type"));
        var changeType = B("btnChartChangeType", "Đổi loại\nbiểu đồ", "Change\nChart Type", XlIcons.Large("colChart"), "Đổi loại biểu đồ", "Change the chart type", () => { }, RibbonButtonSize.Large);
        changeType.DropDown = Menu(ChartDialog.Types.Select(t => (ToolStripItem)Mi("miChartType" + t.Type, t.Vi, t.En, XlIcons.Small(t.Icon), () => SetChartType(t.Type))).ToArray());
        g.Large(changeType);
        g.Large(B("btnChartEditData", "Chọn\ndữ liệu", "Select\nData", XlIcons.Large("editData"), "Sửa vùng dữ liệu, tiêu đề, kiểu biểu đồ", "Edit data range, title and type", CmdEditChart, RibbonButtonSize.Large));
        g.Large(B("btnChartSwitch", "Đổi hàng/\ncột", "Switch\nRow/Column", XlIcons.Large("switchRowCol"), "Đổi chuỗi dữ liệu giữa hàng và cột", "Swap the data over the axis", () => ChartChange(c => c.SeriesInRows = !c.SeriesInRows), RibbonButtonSize.Large));
        g = _chartPage.AddGroup("ChartElements", Loc.T("Thành phần", "Elements"));
        g.Stack(B("btnChartTitle", "Tiêu đề", "Chart Title", XlIcons.Small("chartTitle"), "Sửa tiêu đề biểu đồ", "Edit the chart title", CmdChartTitle, RibbonButtonSize.Small),
                B("btnChartLegend", "Chú giải", "Legend", XlIcons.Small("legend"), "Hiện/ẩn chú giải", "Show/hide the legend", () => ChartChange(c => c.ShowLegend = !c.ShowLegend), RibbonButtonSize.Small),
                B("btnChartDataLabels", "Nhãn dữ liệu", "Data Labels", XlIcons.Small("dataLabels"), "Hiện/ẩn nhãn dữ liệu", "Show/hide data labels", () => ChartChange(c => c.ShowDataLabels = !c.ShowDataLabels), RibbonButtonSize.Small));
        g = _chartPage.AddGroup("ChartArrange", Loc.T("Sắp xếp", "Arrange"));
        g.Large(B("btnChartDelete", "Xóa\nbiểu đồ", "Delete\nChart", XlIcons.Large("deleteObj"), "Xóa biểu đồ đang chọn (Delete)", "Delete the selected chart (Delete)", () => Grid.DeleteSelectedObject(), RibbonButtonSize.Large));
        _picturePage = Ribbon.AddPage("PictureFormat", Loc.T("Ảnh", "Picture Format"), contextual: true);
        g = _picturePage.AddGroup("PictureSize", Loc.T("Kích thước", "Size"));
        g.Large(B("btnPictureReset", "Kích thước\ngốc", "Reset\nSize", XlIcons.Large("picture"), "Đặt lại kích thước gốc của ảnh", "Reset the picture to its original size", CmdPictureResetSize, RibbonButtonSize.Large));
        g.Large(B("btnPictureDelete", "Xóa ảnh", "Delete\nPicture", XlIcons.Large("deleteObj"), "Xóa ảnh đang chọn", "Delete the selected picture", () => Grid.DeleteSelectedObject(), RibbonButtonSize.Large));
        Ribbon.Select(home);
    }

    private int _lastFill = Color.FromArgb(255, 255, 0).ToArgb();
    private int _lastFontColor = Color.FromArgb(192, 0, 0).ToArgb();

    private static float NextSize(float s, int dir)
    {
        float[] sizes = [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
        if (dir > 0) return sizes.FirstOrDefault(x => x > s + 0.01f, Math.Min(409, s + 10));
        var lower = sizes.Where(x => x < s - 0.01f).ToList();
        return lower.Count > 0 ? lower[^1] : Math.Max(1, s - 1);
    }
}
