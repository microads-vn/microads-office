using System.Globalization;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Editor;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.UI;

public sealed partial class WordMainForm
{
    private readonly Dictionary<string, Control> _cmd = new();
    private readonly Dictionary<string, ToolStripItem> _items = new();
    private ToolStrip _statusBar = null!;
    private ToolStripLabel _lblPage = null!, _lblWords = null!, _lblLang = null!, _lblZoom = null!;
    private ZoomSlider _zoomSlider = null!;
    private ComboBox _cboFont = null!, _cboSize = null!;
    private StyleGallery _gallery = null!;
    private RibbonPage _pageTable = null!, _pagePicture = null!;
    private NumericUpDown _numIndL = null!, _numIndR = null!, _numBefore = null!, _numAfter = null!, _numPicW = null!, _numPicH = null!;
    private bool _updating;
    private Color _lastFontColor = Color.FromArgb(192, 0, 0), _lastHighlight = Color.Yellow, _lastShading = Color.FromArgb(217, 217, 217);
    private (CharProps C, ParaProps P, bool Sticky)? _painter;

    /// <summary>All ribbon controls by name (used by tests and keyboard).</summary>
    public IReadOnlyDictionary<string, Control> Commands => _cmd;
    /// <summary>All menu items of ribbon drop-downs by name.</summary>
    public IReadOnlyDictionary<string, ToolStripItem> MenuItems => _items;

    private RibbonButton B(string name, string text, Image? img, string tip, Action a, RibbonButtonSize size = RibbonButtonSize.Auto)
    {
        var b = RibbonButton.Create(name, text, img, tip, () => { if (!_backstage.Visible || name.StartsWith("bs")) a(); }, size);
        _cmd[name] = b;
        return b;
    }

    private ContextMenuStrip MakeMenu(string name, params (string Name, string Text, Action A)[] items)
    {
        var m = new ContextMenuStrip { Renderer = new RibbonRenderer(false), Name = name, ShowImageMargin = false };
        foreach (var (n, t, a) in items)
        {
            if (t == "-") { m.Items.Add(new ToolStripSeparator()); continue; }
            var it = new ToolStripMenuItem(t, null, (_, _) => { a(); _ed.Focus(); }) { Name = n };
            _items[n] = it;
            m.Items.Add(it);
        }
        return m;
    }

    private static Label RL(string text) => new() { Text = text, AutoSize = true, Font = Theme.UiFont, ForeColor = Theme.Ink, Padding = new Padding(0, Dpi.Scale(3), Dpi.Scale(2), 0) };

    private NumericUpDown RNum(string name, decimal min, decimal max, int dec, decimal inc, Action<decimal> apply)
    {
        var n = new NumericUpDown { Name = name, Minimum = min, Maximum = max, DecimalPlaces = dec, Increment = inc, Width = Dpi.Scale(64), Font = Theme.UiFont, TabStop = false };
        n.ValueChanged += (_, _) => { if (!_updating) apply(n.Value); };
        _cmd[name] = n;
        return n;
    }

    private void BuildRibbon()
    {
        var sv = _ribbon.AddQuickAccess("qatSave", Icons.Small(WI.Save), Loc.T("Lưu (Ctrl+S)", "Save (Ctrl+S)"), () => CmdSave());
        var u = _ribbon.AddQuickAccess("qatUndo", Icons.Small(WI.Undo), Loc.T("Hoàn tác (Ctrl+Z)", "Undo (Ctrl+Z)"), () => _ed.Undo());
        var r = _ribbon.AddQuickAccess("qatRedo", Icons.Small(WI.Redo), Loc.T("Làm lại (Ctrl+Y)", "Redo (Ctrl+Y)"), () => _ed.Redo());
        _items["qatSave"] = sv;
        _items["qatUndo"] = u; _items["qatRedo"] = r;
        _ribbon.FileClicked += (_, _) => ShowBackstage("info");
        _ribbon.CommandInvoking += (_, _) => HideBackstage();
        _ribbon.AddHeaderRight(OfficeShell.CreateHeaderItems(this));
        BuildHome();
        BuildInsert();
        BuildLayout();
        BuildReferences();
        BuildReview();
        BuildView();
        BuildHelp();
        BuildTableTab();
        BuildPictureTab();
        _ribbon.Select(_ribbon.Pages[0]);
        // groups measure themselves when first laid out; re-measure now that all items exist
        foreach (var page in _ribbon.Pages) foreach (var g in page.Groups) g.Remeasure();
    }

    // ------------------------------------------------------------------ Home

    private void BuildHome()
    {
        var home = _ribbon.AddPage("Home", Loc.T("Trang chủ", "Home"));
        var clip = home.AddGroup("Clipboard", Loc.T("Bộ nhớ tạm", "Clipboard"));
        var paste = B("btnPaste", Loc.T("Dán", "Paste"), WI.G(WI.Paste, true), Loc.T("Dán (Ctrl+V)", "Paste (Ctrl+V)"), () => _ed.Paste(), RibbonButtonSize.Large);
        paste.Split = true;
        paste.DropDown = MakeMenu("menuPaste", ("miPasteKeep", Loc.T("Giữ nguyên định dạng", "Keep source formatting"), () => _ed.Paste()),
            ("miPasteText", Loc.T("Chỉ giữ văn bản (Ctrl+Shift+V)", "Keep text only (Ctrl+Shift+V)"), () => _ed.Paste(true)));
        clip.Large(paste);
        clip.Stack(B("btnCut", Loc.T("Cắt", "Cut"), WI.G(WI.Cut, false), Loc.T("Cắt (Ctrl+X)", "Cut (Ctrl+X)"), () => _ed.Cut()),
            B("btnCopy", Loc.T("Sao chép", "Copy"), WI.G(WI.Copy, false), Loc.T("Sao chép (Ctrl+C)", "Copy (Ctrl+C)"), () => _ed.Copy()),
            B("btnFormatPainter", Loc.T("Sao định dạng", "Format Painter"), WI.G(WI.Brush, false), Loc.T("Sao chép định dạng sang vùng khác (bấm đúp để dùng nhiều lần)", "Copy formatting to another place (double-click to keep it on)"), () => StartFormatPainter(false)));
        _cmd["btnFormatPainter"].DoubleClick += (_, _) => StartFormatPainter(true);

        var font = home.AddGroup("Font", Loc.T("Phông chữ", "Font"));
        _cboFont = new ComboBox { Name = "cboFontName", Width = Dpi.Scale(150), DropDownStyle = ComboBoxStyle.DropDown, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = Dpi.Scale(20), MaxDropDownItems = 18, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems, AccessibleName = Loc.T("Phông chữ", "Font"), TabStop = false };
        _cboFont.Items.AddRange(DlgUtil.FontNames());
        var fontCache = new Dictionary<string, Font>();
        _cboFont.DrawItem += (_, e) =>
        {
            if (e.Index < 0) return;
            e.DrawBackground();
            var name = (string)_cboFont.Items[e.Index]!;
            if (!fontCache.TryGetValue(name, out var f)) { try { f = new Font(name, 10f); } catch { f = Theme.UiFont; } fontCache[name] = f; }
            TextRenderer.DrawText(e.Graphics, name, (e.State & DrawItemState.ComboBoxEdit) != 0 ? Theme.UiFont : f, e.Bounds, (e.State & DrawItemState.Selected) != 0 ? SystemColors.HighlightText : Theme.Ink, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        };
        _cboFont.SelectionChangeCommitted += (_, _) => BeginInvoke(() => ApplyFontName(_cboFont.Text));
        _cboFont.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplyFontName(_cboFont.Text); } };
        _cmd["cboFontName"] = _cboFont;
        _cboSize = new ComboBox { Name = "cboFontSize", Width = Dpi.Scale(52), DropDownStyle = ComboBoxStyle.DropDown, AccessibleName = Loc.T("Cỡ chữ", "Font size"), TabStop = false };
        _cboSize.Items.AddRange(FontSizes.Standard.Select(s => (object)s.ToString(CultureInfo.InvariantCulture)).ToArray());
        _cboSize.SelectionChangeCommitted += (_, _) => BeginInvoke(() => ApplyFontSize(_cboSize.Text));
        _cboSize.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplyFontSize(_cboSize.Text); } };
        _cmd["cboFontSize"] = _cboSize;
        font.Row(_cboFont, _cboSize,
            B("btnGrowFont", "", WI.G(WI.GrowFont, false), Loc.T("Tăng cỡ chữ (Ctrl+Shift+>)", "Increase font size (Ctrl+Shift+>)"), () => _ed.ApplyChar(p => p with { Size = FontSizes.Next(p.Size, true) })),
            B("btnShrinkFont", "", WI.G(WI.ShrinkFont, false), Loc.T("Giảm cỡ chữ (Ctrl+Shift+<)", "Decrease font size (Ctrl+Shift+<)"), () => _ed.ApplyChar(p => p with { Size = FontSizes.Next(p.Size, false) })));
        var caseBtn = B("btnChangeCase", "", WI.ChangeCase(false), Loc.T("Đổi kiểu chữ hoa/thường", "Change case"), () => { });
        caseBtn.DropDown = MakeMenu("menuCase", ("miCaseSentence", Loc.T("Viết hoa đầu câu.", "Sentence case."), () => _ed.TransformSelection(s => Editing.ChangeCase(s, 3))),
            ("miCaseLower", Loc.T("chữ thường", "lowercase"), () => _ed.TransformSelection(s => Editing.ChangeCase(s, 1))),
            ("miCaseUpper", Loc.T("CHỮ HOA", "UPPERCASE"), () => _ed.TransformSelection(s => Editing.ChangeCase(s, 0))),
            ("miCaseTitle", Loc.T("Viết Hoa Mỗi Từ", "Capitalize Each Word"), () => _ed.TransformSelection(s => Editing.ChangeCase(s, 2))));
        var clear = B("btnClearFormat", "", WI.ClearFormat(false), Loc.T("Xóa định dạng (Ctrl+Space)", "Clear formatting (Ctrl+Space)"), ClearFormatting);
        var under = B("btnUnderline", "", WI.G(WI.Underline, false), Loc.T("Gạch chân (Ctrl+U)", "Underline (Ctrl+U)"), () => { var v = _ed.CurrentCharProps.Underline == UnderlineStyle.None ? UnderlineStyle.Single : UnderlineStyle.None; _ed.ApplyChar(p => p with { Underline = v }); });
        under.Split = true;
        under.DropDown = MakeMenu("menuUnderline", ("miUlSingle", Loc.T("Gạch đơn", "Single"), () => _ed.ApplyChar(p => p with { Underline = UnderlineStyle.Single })),
            ("miUlDouble", Loc.T("Gạch đôi", "Double"), () => _ed.ApplyChar(p => p with { Underline = UnderlineStyle.Double })),
            ("miUlDotted", Loc.T("Gạch chấm", "Dotted"), () => _ed.ApplyChar(p => p with { Underline = UnderlineStyle.Dotted })),
            ("miUlThick", Loc.T("Gạch đậm", "Thick"), () => _ed.ApplyChar(p => p with { Underline = UnderlineStyle.Thick })),
            ("miUlWave", Loc.T("Lượn sóng", "Wave"), () => _ed.ApplyChar(p => p with { Underline = UnderlineStyle.Wavy })));
        var fc = B("btnFontColor", "", WI.G(WI.FontColor, false), Loc.T("Màu chữ", "Font colour"), () => { var c = _lastFontColor; _ed.ApplyChar(p => p with { Color = c }); });
        fc.Split = true; fc.SwatchColor = _lastFontColor;
        fc.DropDownFactory = () => new ColorPickerDropDown(c => { if (c is Color cc) { _lastFontColor = cc; fc.SwatchColor = cc; } _ed.ApplyChar(p => p with { Color = c }); _ed.Focus(); }, Loc.T("Tự động", "Automatic"), _lastFontColor);
        var hl = B("btnHighlight", "", WI.G(WI.Highlight, false), Loc.T("Tô sáng văn bản", "Text highlight colour"), () => { var c = _lastHighlight; _ed.ApplyChar(p => p with { Highlight = c }); });
        hl.Split = true; hl.SwatchColor = _lastHighlight;
        hl.DropDownFactory = () => new ColorPickerDropDown(c => { if (c is Color cc) { _lastHighlight = cc; hl.SwatchColor = cc; } _ed.ApplyChar(p => p with { Highlight = c }); _ed.Focus(); }, Loc.T("Không màu", "No colour"), _lastHighlight);
        font.Row(caseBtn, clear);
        font.Row(B("btnBold", "", WI.G(WI.Bold, false), Loc.T("Đậm (Ctrl+B)", "Bold (Ctrl+B)"), () => { var v = !_ed.CurrentCharProps.Bold; _ed.ApplyChar(p => p with { Bold = v }); }),
            B("btnItalic", "", WI.G(WI.Italic, false), Loc.T("Nghiêng (Ctrl+I)", "Italic (Ctrl+I)"), () => { var v = !_ed.CurrentCharProps.Italic; _ed.ApplyChar(p => p with { Italic = v }); }),
            under,
            B("btnStrike", "", WI.Strike(false), Loc.T("Gạch ngang", "Strikethrough"), () => { var v = !_ed.CurrentCharProps.Strike; _ed.ApplyChar(p => p with { Strike = v }); }),
            B("btnSubscript", "", WI.Subscript(false), Loc.T("Chỉ số dưới (Ctrl+=)", "Subscript (Ctrl+=)"), () => { var v = _ed.CurrentCharProps.VertAlign == VertAlign.Subscript ? VertAlign.Baseline : VertAlign.Subscript; _ed.ApplyChar(p => p with { VertAlign = v }); }),
            B("btnSuperscript", "", WI.Superscript(false), Loc.T("Chỉ số trên (Ctrl+Shift+=)", "Superscript (Ctrl+Shift+=)"), () => { var v = _ed.CurrentCharProps.VertAlign == VertAlign.Superscript ? VertAlign.Baseline : VertAlign.Superscript; _ed.ApplyChar(p => p with { VertAlign = v }); }),
            fc, hl);
        font.SetDialogLauncher(Loc.T("Hộp thoại Phông chữ (Ctrl+D)", "Font dialog (Ctrl+D)"), CmdFontDialog);

        var para = home.AddGroup("Paragraph", Loc.T("Đoạn văn", "Paragraph"));
        var bul = B("btnBullets", "", WI.G(WI.Bullets, false), Loc.T("Dấu đầu dòng", "Bullets"), () => _ed.ToggleList(ListType.Bullet));
        bul.Split = true;
        bul.DropDown = MakeMenu("menuBullets", ("miBulletOn", Loc.T("Dấu đầu dòng •", "Bullets •"), () => { if (_ed.CurrentParaProps.List != ListType.Bullet) _ed.ToggleList(ListType.Bullet); }),
            ("miBulletOff", Loc.T("Bỏ dấu đầu dòng", "Remove bullets"), () => { if (_ed.CurrentParaProps.List != ListType.None) _ed.ToggleList(_ed.CurrentParaProps.List); }));
        var num = B("btnNumbering", "", WI.Numbering(false), Loc.T("Đánh số", "Numbering"), () => _ed.ToggleList(ListType.Decimal));
        num.Split = true;
        num.DropDown = MakeMenu("menuNumbering",
            ("miNum1", "1.  2.  3.", () => SetList(ListType.Decimal)), ("miNumA", "a)  b)  c)", () => SetList(ListType.LowerAlpha)),
            ("miNumAU", "A.  B.  C.", () => SetList(ListType.UpperAlpha)), ("miNumi", "i.  ii.  iii.", () => SetList(ListType.LowerRoman)),
            ("miNumI", "I.  II.  III.", () => SetList(ListType.UpperRoman)), ("sep", "-", () => { }),
            ("miNumOff", Loc.T("Bỏ đánh số", "Remove numbering"), () => { if (_ed.CurrentParaProps.List != ListType.None) _ed.ToggleList(_ed.CurrentParaProps.List); }));
        var multi = B("btnMultilevel", "", WI.Multilevel(false), Loc.T("Danh sách nhiều cấp (Tab / Shift+Tab đổi cấp)", "Multilevel list (Tab / Shift+Tab changes level)"), () => { if (_ed.CurrentParaProps.List == ListType.None) _ed.ToggleList(ListType.Decimal); else _ed.ChangeIndent(1); });
        var spacing = B("btnLineSpacing", "", WI.LineSpacing(false), Loc.T("Giãn dòng và đoạn", "Line and paragraph spacing"), () => { });
        spacing.DropDown = MakeMenu("menuLineSpacing",
            ("miLs10", "1,0", () => LineSpacing(1f)), ("miLs115", "1,15", () => LineSpacing(1.15f)), ("miLs15", "1,5", () => LineSpacing(1.5f)),
            ("miLs20", "2,0", () => LineSpacing(2f)), ("miLs25", "2,5", () => LineSpacing(2.5f)), ("miLs30", "3,0", () => LineSpacing(3f)), ("sep", "-", () => { }),
            ("miSpaceBefore", Loc.T("Thêm/bỏ khoảng cách trước đoạn", "Add/remove space before paragraph"), () => { var v = _ed.CurrentParaProps.SpaceBefore > 0 ? 0f : 12f; _ed.ApplyPara(p => p with { SpaceBefore = v }); }),
            ("miSpaceAfter", Loc.T("Thêm/bỏ khoảng cách sau đoạn", "Add/remove space after paragraph"), () => { var v = _ed.CurrentParaProps.SpaceAfter > 0 ? 0f : 8f; _ed.ApplyPara(p => p with { SpaceAfter = v }); }),
            ("miLsOptions", Loc.T("Tùy chọn giãn dòng...", "Line spacing options..."), CmdParagraphDialog));
        var shading = B("btnShading", "", WI.Shading(false), Loc.T("Tô nền", "Shading"), () => ApplyShading(_lastShading));
        shading.Split = true; shading.SwatchColor = _lastShading;
        shading.DropDownFactory = () => new ColorPickerDropDown(c => { if (c is Color cc) { _lastShading = cc; shading.SwatchColor = cc; } ApplyShading(c); _ed.Focus(); }, Loc.T("Không màu", "No colour"), _lastShading);
        var borders = B("btnBorders", "", WI.Borders(false), Loc.T("Viền", "Borders"), () => { var v = !_ed.CurrentParaProps.Border; _ed.ApplyPara(p => p with { Border = v }); });
        borders.Split = true;
        borders.DropDown = MakeMenu("menuBorders",
            ("miBorderBox", Loc.T("Khung quanh đoạn", "Box around paragraph"), () => _ed.ApplyPara(p => p with { Border = true })),
            ("miBorderNone", Loc.T("Không viền", "No border"), () => _ed.ApplyPara(p => p with { Border = false })),
            ("sep", "-", () => { }),
            ("miTableBordersAll", Loc.T("Tất cả viền bảng", "All table borders"), () => SetTableBorders(true)),
            ("miTableBordersNone", Loc.T("Bỏ viền bảng", "No table borders"), () => SetTableBorders(false)));
        para.Row(bul, num, multi,
            B("btnDecreaseIndent", "", WI.IndentLess(false), Loc.T("Giảm thụt lề", "Decrease indent"), () => _ed.ChangeIndent(-1)),
            B("btnIncreaseIndent", "", WI.IndentMore(false), Loc.T("Tăng thụt lề", "Increase indent"), () => _ed.ChangeIndent(1)),
            B("btnSort", "", WI.G(WI.Sort, false), Loc.T("Sắp xếp đoạn A→Z", "Sort paragraphs A→Z"), CmdSort),
            B("btnShowMarks", "", WI.Pilcrow(false), Loc.T("Hiện/ẩn ký tự định dạng (Ctrl+Shift+8)", "Show/hide formatting marks (Ctrl+Shift+8)"), ToggleMarks));
        para.Row(B("btnAlignLeft", "", WI.G(WI.AlignLeft, false), Loc.T("Căn trái (Ctrl+L)", "Align left (Ctrl+L)"), () => _ed.ApplyPara(p => p with { Align = ParaAlign.Left })),
            B("btnAlignCenter", "", WI.G(WI.AlignCenter, false), Loc.T("Căn giữa (Ctrl+E)", "Center (Ctrl+E)"), () => _ed.ApplyPara(p => p with { Align = ParaAlign.Center })),
            B("btnAlignRight", "", WI.G(WI.AlignRight, false), Loc.T("Căn phải (Ctrl+R)", "Align right (Ctrl+R)"), () => _ed.ApplyPara(p => p with { Align = ParaAlign.Right })),
            B("btnJustify", "", WI.Justify(false), Loc.T("Căn đều hai bên (Ctrl+J)", "Justify (Ctrl+J)"), () => _ed.ApplyPara(p => p with { Align = ParaAlign.Justify })),
            spacing, shading, borders);
        para.SetDialogLauncher(Loc.T("Hộp thoại Đoạn văn", "Paragraph dialog"), CmdParagraphDialog);

        var styles = home.AddGroup("Styles", Loc.T("Kiểu", "Styles"));
        _gallery = new StyleGallery(() => _ed.Document);
        _gallery.StylePicked += id => { _ed.ApplyStyle(id); _ed.Focus(); };
        foreach (var c in _gallery.Controls.OfType<Control>()) _cmd[c.Name] = c;
        styles.Large(_gallery);
        styles.CollapsedIcon = WI.StyleGlyph(true);

        var edit = home.AddGroup("Editing", Loc.T("Chỉnh sửa", "Editing"));
        edit.Stack(B("btnFind", Loc.T("Tìm", "Find"), WI.G(WI.Find, false), Loc.T("Tìm (Ctrl+F)", "Find (Ctrl+F)"), () => ShowFind(0)),
            B("btnReplace", Loc.T("Thay thế", "Replace"), WI.G(WI.Replace, false), Loc.T("Thay thế (Ctrl+H)", "Replace (Ctrl+H)"), () => ShowFind(1)),
            B("btnSelectAll", Loc.T("Chọn tất cả", "Select All"), WI.G(WI.Select, false), Loc.T("Chọn tất cả (Ctrl+A)", "Select all (Ctrl+A)"), () => { _ed.SelectAll(); _ed.Focus(); }));

        var ai = home.AddGroup("Ai", Loc.T("Trợ lý AI", "AI"));
        ai.Large(B("btnAiAssistant", Loc.T("Trợ lý\nAI", "AI\nAssistant"), MicroadsOffice.AI.UI.AiIcons.Sparkle(true),
            Loc.T("Mở Trợ lý AI (ChatGPT, Claude…): hỏi đáp, soạn, viết lại, dịch, sửa chính tả ngay trong văn bản", "Open the AI assistant (ChatGPT, Claude…): ask, draft, rewrite, translate and proofread right in the document"), ToggleAi));
        ai.CollapsedIcon = MicroadsOffice.AI.UI.AiIcons.Sparkle(true);
    }

    // ------------------------------------------------------------------ Insert

    private void BuildInsert()
    {
        var ins = _ribbon.AddPage("Insert", Loc.T("Chèn", "Insert"));
        var pages = ins.AddGroup("Pages", Loc.T("Trang", "Pages"));
        pages.Large(B("btnPageBreak", Loc.T("Ngắt\ntrang", "Page\nBreak"), WI.PageBreak(true), Loc.T("Ngắt trang (Ctrl+Enter)", "Page break (Ctrl+Enter)"), () => _ed.InsertText(Run.PageBreak.ToString(), "break")));
        var tables = ins.AddGroup("Tables", Loc.T("Bảng", "Tables"));
        var tbl = B("btnTable", Loc.T("Bảng", "Table"), WI.G(WI.Table, true), Loc.T("Chèn bảng", "Insert table"), () => { }, RibbonButtonSize.Large);
        tbl.DropDownFactory = () => new TableGridDropDown((r, c) => { _ed.InsertTable(r, c); _ed.Focus(); }, CmdInsertTableDialog);
        tables.Large(tbl);
        _cmd["btnInsertTableDialog"] = B("btnInsertTableDialog", "", null, "", CmdInsertTableDialog);
        var ill = ins.AddGroup("Illustrations", Loc.T("Minh họa", "Illustrations"));
        ill.Large(B("btnPicture", Loc.T("Ảnh", "Picture"), WI.G(WI.Picture, true), Loc.T("Chèn ảnh từ tệp", "Insert a picture from a file"), CmdInsertPicture));
        var links = ins.AddGroup("Links", Loc.T("Liên kết", "Links"));
        links.Large(B("btnHyperlink", Loc.T("Liên kết", "Link"), WI.G(WI.Link, true), Loc.T("Chèn liên kết (Ctrl+K)", "Insert hyperlink (Ctrl+K)"), CmdHyperlink));
        var hf = ins.AddGroup("HeaderFooter", Loc.T("Đầu trang và chân trang", "Header and Footer"));
        hf.Large(B("btnHeader", Loc.T("Đầu\ntrang", "Header"), WI.Header(true), Loc.T("Sửa đầu trang", "Edit header"), () => CmdHeaderFooter(true)));
        hf.Large(B("btnFooter", Loc.T("Chân\ntrang", "Footer"), WI.Footer(true), Loc.T("Sửa chân trang", "Edit footer"), () => CmdHeaderFooter(false)));
        var pn = B("btnPageNumber", Loc.T("Số\ntrang", "Page\nNumber"), WI.PageNumber(true), Loc.T("Chèn số trang", "Insert page numbers"), () => { }, RibbonButtonSize.Large);
        pn.DropDown = MakeMenu("menuPageNumber",
            ("miPnBottomCenter", Loc.T("Cuối trang – giữa", "Bottom of page – center"), () => SetPageNumber(false, ParaAlign.Center, false)),
            ("miPnBottomRight", Loc.T("Cuối trang – phải", "Bottom of page – right"), () => SetPageNumber(false, ParaAlign.Right, false)),
            ("miPnTopRight", Loc.T("Đầu trang – phải", "Top of page – right"), () => SetPageNumber(true, ParaAlign.Right, false)),
            ("miPnXofY", Loc.T("Trang X / Y (cuối trang)", "Page X of Y (bottom)"), () => SetPageNumber(false, ParaAlign.Center, true)),
            ("sep", "-", () => { }),
            ("miPnRemove", Loc.T("Xóa số trang", "Remove page numbers"), RemovePageNumbers));
        hf.Large(pn);
        var text = ins.AddGroup("Text", Loc.T("Văn bản", "Text"));
        text.Stack(B("btnDateTime", Loc.T("Ngày giờ", "Date and Time"), WI.G(WI.Calendar, false), Loc.T("Chèn ngày giờ", "Insert date and time"), CmdDateTime),
            B("btnSymbol", Loc.T("Ký hiệu", "Symbol"), WI.Symbol(false), Loc.T("Chèn ký hiệu đặc biệt", "Insert a special symbol"), CmdSymbol),
            B("btnHorizontalLine", Loc.T("Đường kẻ ngang", "Horizontal line"), WI.HLine(false), Loc.T("Chèn đường kẻ ngang", "Insert a horizontal line"), CmdHorizontalLine));
    }

    // ------------------------------------------------------------------ Layout

    private void BuildLayout()
    {
        var lay = _ribbon.AddPage("Layout", Loc.T("Bố cục", "Layout"));
        var setup = lay.AddGroup("PageSetup", Loc.T("Thiết lập trang", "Page Setup"));
        var margins = B("btnMargins", Loc.T("Lề", "Margins"), WI.Margins(true), Loc.T("Chọn lề trang", "Select page margins"), () => { }, RibbonButtonSize.Large);
        margins.DropDown = MakeMenu("menuMargins",
            ("miMarVN", Loc.T("Văn bản hành chính VN (T20 D20 Tr30 P15 mm)", "Vietnamese admin (T20 B20 L30 R15 mm)"), () => SetMargins(20, 20, 30, 15)),
            ("miMarNormal", Loc.T("Thường (2,54 cm)", "Normal (2.54 cm)"), () => SetMargins(25.4f, 25.4f, 25.4f, 25.4f)),
            ("miMarNarrow", Loc.T("Hẹp (1,27 cm)", "Narrow (1.27 cm)"), () => SetMargins(12.7f, 12.7f, 12.7f, 12.7f)),
            ("miMarModerate", Loc.T("Vừa (2,54 × 1,91 cm)", "Moderate (2.54 × 1.91 cm)"), () => SetMargins(25.4f, 25.4f, 19.1f, 19.1f)),
            ("miMarWide", Loc.T("Rộng (2,54 × 5,08 cm)", "Wide (2.54 × 5.08 cm)"), () => SetMargins(25.4f, 25.4f, 50.8f, 50.8f)),
            ("sep", "-", () => { }),
            ("miMarCustom", Loc.T("Lề tùy chỉnh...", "Custom margins..."), CmdPageSetup));
        var orient = B("btnOrientation", Loc.T("Hướng", "Orientation"), WI.Orientation(true), Loc.T("Hướng giấy", "Page orientation"), () => { }, RibbonButtonSize.Large);
        orient.DropDown = MakeMenu("menuOrientation", ("miPortrait", Loc.T("Dọc", "Portrait"), () => SetOrientation(false)), ("miLandscape", Loc.T("Ngang", "Landscape"), () => SetOrientation(true)));
        var size = B("btnSize", Loc.T("Khổ giấy", "Size"), WI.PaperSize(true), Loc.T("Khổ giấy", "Paper size"), () => { }, RibbonButtonSize.Large);
        var sizeItems = PageSetup.Papers.Select(p => ("miPaper" + p.Name, $"{p.Name} ({PageSetup.ToMm(p.W):0} × {PageSetup.ToMm(p.H):0} mm)", (Action)(() => SetPaper(p.Name, p.W, p.H)))).ToList();
        sizeItems.Add(("sep", "-", () => { }));
        sizeItems.Add(("miPaperMore", Loc.T("Khổ giấy khác...", "More paper sizes..."), CmdPageSetup));
        size.DropDown = MakeMenu("menuSize", sizeItems.ToArray());
        var breaks = B("btnBreaks", Loc.T("Ngắt", "Breaks"), WI.PageBreak(true), Loc.T("Ngắt trang / ngắt dòng", "Page / line breaks"), () => { }, RibbonButtonSize.Large);
        breaks.DropDown = MakeMenu("menuBreaks", ("miBreakPage", Loc.T("Ngắt trang (Ctrl+Enter)", "Page break (Ctrl+Enter)"), () => _ed.InsertText(Run.PageBreak.ToString(), "break")),
            ("miBreakLine", Loc.T("Ngắt dòng (Shift+Enter)", "Line break (Shift+Enter)"), () => _ed.InsertText(Run.LineBreak.ToString(), "break")));
        setup.Large(margins); setup.Large(orient); setup.Large(size); setup.Large(breaks);
        setup.SetDialogLauncher(Loc.T("Thiết lập trang", "Page setup"), CmdPageSetup);

        var para = lay.AddGroup("ParaLayout", Loc.T("Đoạn văn", "Paragraph"));
        _numIndL = RNum("numIndentLeft", 0, 50, 2, 0.25m, v => _ed.ApplyPara(p => p with { IndentLeft = PageSetup.Mm((float)v * 10) }));
        _numIndR = RNum("numIndentRight", 0, 50, 2, 0.25m, v => _ed.ApplyPara(p => p with { IndentRight = PageSetup.Mm((float)v * 10) }));
        _numBefore = RNum("numSpaceBefore", 0, 600, 0, 6, v => _ed.ApplyPara(p => p with { SpaceBefore = (float)v }));
        _numAfter = RNum("numSpaceAfter", 0, 600, 0, 6, v => _ed.ApplyPara(p => p with { SpaceAfter = (float)v }));
        para.Row(RL(Loc.T("Thụt trái (cm)", "Indent left (cm)")), _numIndL, RL(Loc.T("Trước (pt)", "Before (pt)")), _numBefore);
        para.Row(RL(Loc.T("Thụt phải (cm)", "Indent right (cm)")), _numIndR, RL(Loc.T("Sau (pt)  ", "After (pt)  ")), _numAfter);
        para.SetDialogLauncher(Loc.T("Hộp thoại Đoạn văn", "Paragraph dialog"), CmdParagraphDialog);

        var bg = lay.AddGroup("PageBackground", Loc.T("Nền trang", "Page Background"));
        var wm = B("btnWatermark", Loc.T("Hình mờ", "Watermark"), WI.Watermark(true), Loc.T("Hình mờ chữ trên mọi trang", "Text watermark on every page"), () => { }, RibbonButtonSize.Large);
        wm.DropDown = MakeMenu("menuWatermark",
            ("miWmDraft", Loc.T("BẢN NHÁP", "DRAFT"), () => SetWatermark(Loc.T("BẢN NHÁP", "DRAFT"))),
            ("miWmConfidential", Loc.T("MẬT", "CONFIDENTIAL"), () => SetWatermark(Loc.T("MẬT", "CONFIDENTIAL"))),
            ("miWmSample", Loc.T("MẪU", "SAMPLE"), () => SetWatermark(Loc.T("MẪU", "SAMPLE"))),
            ("miWmUrgent", Loc.T("KHẨN", "URGENT"), () => SetWatermark(Loc.T("KHẨN", "URGENT"))),
            ("miWmCustom", Loc.T("Hình mờ tùy chỉnh...", "Custom watermark..."), () =>
            {
                var t = InputDialog.Ask(this, Loc.T("Hình mờ", "Watermark"), Loc.T("Nội dung hình mờ:", "Watermark text:"), _ed.Document.Watermark ?? "");
                if (!string.IsNullOrWhiteSpace(t)) SetWatermark(t.Trim());
            }),
            ("sep", "-", () => { }),
            ("miWmRemove", Loc.T("Xóa hình mờ", "Remove watermark"), () => SetWatermark(null)));
        var pc = B("btnPageColor", Loc.T("Màu\ntrang", "Page\nColor"), WI.PageColor(true), Loc.T("Màu nền trang", "Page background colour"), () => { }, RibbonButtonSize.Large);
        pc.DropDownFactory = () => new ColorPickerDropDown(c => { _ed.DocumentEdit(d => d.PageColor = c); _ed.Focus(); }, Loc.T("Không màu", "No colour"), _ed.Document.PageColor);
        bg.Large(wm); bg.Large(pc);
    }

    // ------------------------------------------------------------------ References / Review / View / Help

    private void BuildReferences()
    {
        var refs = _ribbon.AddPage("References", Loc.T("Tham chiếu", "References"));
        var toc = refs.AddGroup("Toc", Loc.T("Mục lục", "Table of Contents"));
        var tb = B("btnToc", Loc.T("Mục lục", "Table of\nContents"), WI.Toc(true), Loc.T("Chèn mục lục tự động từ các Tiêu đề", "Insert an automatic table of contents from headings"), () => InsertToc(), RibbonButtonSize.Large);
        tb.Split = true;
        tb.DropDown = MakeMenu("menuToc", ("miTocInsert", Loc.T("Mục lục tự động", "Automatic table"), () => InsertToc()),
            ("miTocUpdate", Loc.T("Cập nhật mục lục", "Update table"), () => InsertToc(updateOnly: true)),
            ("miTocRemove", Loc.T("Xóa mục lục", "Remove table of contents"), RemoveToc));
        toc.Large(tb);
        toc.Stack(B("btnUpdateToc", Loc.T("Cập nhật mục lục", "Update Table"), WI.UpdateToc(false), Loc.T("Cập nhật số trang và tiêu đề trong mục lục", "Update headings and page numbers"), () => InsertToc(updateOnly: true)),
            B("btnAddHeading1", Loc.T("Đặt làm Tiêu đề 1", "Make Heading 1"), WI.StyleGlyph(false), Loc.T("Đưa đoạn hiện tại vào mục lục (cấp 1)", "Add the paragraph to the TOC (level 1)"), () => _ed.ApplyStyle(StyleIds.Heading1)),
            B("btnAddHeading2", Loc.T("Đặt làm Tiêu đề 2", "Make Heading 2"), WI.StyleGlyph(false), Loc.T("Đưa đoạn hiện tại vào mục lục (cấp 2)", "Add the paragraph to the TOC (level 2)"), () => _ed.ApplyStyle(StyleIds.Heading2)));
    }

    private void BuildReview()
    {
        var rev = _ribbon.AddPage("Review", Loc.T("Xem lại", "Review"));
        var proof = rev.AddGroup("Proofing", Loc.T("Soát lỗi", "Proofing"));
        proof.Large(B("btnSpelling", Loc.T("Chính tả", "Spelling"), WI.G(WI.Spell, true), Loc.T("Kiểm tra chính tả (F7)", "Check spelling (F7)"), CmdSpelling));
        proof.Large(B("btnWordCount", Loc.T("Đếm từ", "Word\nCount"), WI.WordCount(true), Loc.T("Đếm số từ, ký tự, đoạn, trang", "Count words, characters, paragraphs and pages"), CmdWordCount));
        proof.Large(B("btnSpellAsYouType", Loc.T("Gạch lỗi\nchính tả", "Check as\nyou type"), WI.G(WI.Spell, true), Loc.T("Bật/tắt gạch đỏ dưới từ sai chính tả", "Show/hide red squiggles under misspellings"), () => { _spell.Enabled = !_spell.Enabled; _ed.Invalidate(); UpdateUi(); }));
    }

    private void BuildView()
    {
        var view = _ribbon.AddPage("View", Loc.T("Xem", "View"));
        var views = view.AddGroup("Views", Loc.T("Dạng xem", "Views"));
        views.Large(B("btnPrintLayout", Loc.T("Bố cục\nin", "Print\nLayout"), WI.PrintLayout(true), Loc.T("Xem tài liệu theo trang in", "See the document as printed pages"), () => { }));
        views.Large(B("btnFullScreen", Loc.T("Toàn\nmàn hình", "Full\nScreen"), WI.G(WI.Fullscreen, true), Loc.T("Chế độ đọc toàn màn hình (F11, Esc để thoát)", "Full screen reading (F11, Esc to exit)"), ToggleFullScreen));
        var show = view.AddGroup("Show", Loc.T("Hiển thị", "Show"));
        show.Stack(B("btnRuler", Loc.T("Thước kẻ", "Ruler"), WI.G(WI.Ruler, false), Loc.T("Hiện/ẩn thước kẻ", "Show/hide the ruler"), () => { _ruler.Visible = !_ruler.Visible; UpdateUi(); }),
            B("btnNavPane", Loc.T("Ngăn điều hướng", "Navigation Pane"), WI.NavPane(false), Loc.T("Danh sách tiêu đề để di chuyển nhanh", "Headings list for quick navigation"), () => { _navPane.Visible = !_navPane.Visible; RefreshNav(); UpdateUi(); }),
            B("btnGridlines", Loc.T("Đường lưới bảng", "Table gridlines"), WI.Borders(false), Loc.T("Hiện đường lưới của bảng không viền", "Show gridlines of borderless tables"), () => { _ed.ShowTableGridlines = !_ed.ShowTableGridlines; UpdateUi(); }));
        var zoom = view.AddGroup("Zoom", Loc.T("Thu phóng", "Zoom"));
        zoom.Large(B("btnZoom", Loc.T("Thu phóng", "Zoom"), WI.G(WI.Zoom, true), Loc.T("Chọn tỉ lệ thu phóng", "Choose a zoom level"), CmdZoomDialog));
        zoom.Large(B("btnZoom100", "100%", WI.Zoom100(true), Loc.T("Thu phóng 100%", "Zoom to 100%"), () => _ed.Zoom = 1f));
        zoom.Stack(B("btnOnePage", Loc.T("Một trang", "One Page"), WI.OnePage(false), Loc.T("Vừa một trang trong cửa sổ", "Fit a whole page"), () => _ed.ZoomMode = ZoomMode.OnePage),
            B("btnPageWidth", Loc.T("Vừa chiều rộng", "Page Width"), WI.PageWidth(false), Loc.T("Vừa chiều rộng trang", "Fit the page width"), () => _ed.ZoomMode = ZoomMode.PageWidth));
        var win = view.AddGroup("Window", Loc.T("Cửa sổ", "Window"));
        win.Large(B("btnNewWindow", Loc.T("Tài liệu\nmới", "New\nDocument"), WI.G(WI.NewDoc, true), Loc.T("Mở tài liệu mới trong cửa sổ khác (Ctrl+N)", "New document in a new window (Ctrl+N)"), () => NewDocument(null, newWindow: true)));
    }

    private void BuildHelp()
    {
        var help = _ribbon.AddPage("Help", Loc.T("Trợ giúp", "Help"));
        var g = help.AddGroup("HelpGroup", Loc.T("Trợ giúp", "Help"));
        g.Large(B("btnHelp", Loc.T("Hướng dẫn", "Help"), WI.G(WI.Help, true), Loc.T("Hướng dẫn sử dụng & phím tắt (F1)", "User guide & shortcuts (F1)"), ShowHelp));
        g.Large(B("btnCheckUpdate", Loc.T("Kiểm tra\ncập nhật", "Check for\nUpdates"), WI.G(WI.Sync, true), Loc.T("Kiểm tra phiên bản mới tại microads.vn/office", "Check microads.vn/office for a new version"), () => OfficeShell.CheckForUpdatesInteractive(this)));
        g.Large(B("btnSettings", Loc.T("Cài đặt &&\nngôn ngữ", "Settings &&\nLanguage"), WI.G(WI.Settings, true), Loc.T("Ngôn ngữ (Tiếng Việt / English), cập nhật tự động", "Language (Vietnamese / English), automatic updates"), () => OfficeShell.ShowSettings(_primary ?? this)));
        g.Large(B("btnAbout", Loc.T("Giới thiệu", "About"), WI.G(WI.Info, true), Loc.T("Thông tin Microads Office", "About Microads Office"), () => OfficeShell.ShowAbout(this)));
        g.Large(B("btnVisitWebsite", "microads.vn", WI.G(WI.Globe, true), Loc.T("Truy cập microads.vn — giải pháp phần mềm doanh nghiệp", "Visit microads.vn — business software solutions"), () => OfficeShell.OpenWebsite()));
    }

    private void BuildTableTab()
    {
        _pageTable = _ribbon.AddPage("TableTools", Loc.T("Bảng", "Table"), contextual: true);
        var rc = _pageTable.AddGroup("RowsCols", Loc.T("Hàng và cột", "Rows and Columns"));
        rc.Large(B("btnRowAbove", Loc.T("Chèn\nphía trên", "Insert\nAbove"), WI.RowAbove(true), Loc.T("Chèn hàng phía trên", "Insert a row above"), () => _ed.TableEdit(c => Editing.InsertRow(_ed.Document, c, false))));
        rc.Large(B("btnRowBelow", Loc.T("Chèn\nphía dưới", "Insert\nBelow"), WI.RowBelow(true), Loc.T("Chèn hàng phía dưới", "Insert a row below"), () => _ed.TableEdit(c => Editing.InsertRow(_ed.Document, c, true))));
        rc.Large(B("btnColLeft", Loc.T("Chèn\nbên trái", "Insert\nLeft"), WI.ColLeft(true), Loc.T("Chèn cột bên trái", "Insert a column to the left"), () => _ed.TableEdit(c => Editing.InsertColumn(_ed.Document, c, false))));
        rc.Large(B("btnColRight", Loc.T("Chèn\nbên phải", "Insert\nRight"), WI.ColRight(true), Loc.T("Chèn cột bên phải", "Insert a column to the right"), () => _ed.TableEdit(c => Editing.InsertColumn(_ed.Document, c, true))));
        var del = _pageTable.AddGroup("Delete", Loc.T("Xóa", "Delete"));
        del.Stack(B("btnDeleteRow", Loc.T("Xóa hàng", "Delete Rows"), WI.DeleteRow(false), Loc.T("Xóa hàng hiện tại", "Delete the current row"), () => _ed.TableEdit(c => Editing.DeleteRow(_ed.Document, c))),
            B("btnDeleteCol", Loc.T("Xóa cột", "Delete Columns"), WI.DeleteCol(false), Loc.T("Xóa cột hiện tại", "Delete the current column"), () => _ed.TableEdit(c => Editing.DeleteColumn(_ed.Document, c))),
            B("btnDeleteTable", Loc.T("Xóa bảng", "Delete Table"), WI.DeleteTable(false), Loc.T("Xóa toàn bộ bảng", "Delete the whole table"), () => _ed.TableEdit(c => Editing.DeleteTable(_ed.Document, c.Row!.Table!))));
        var merge = _pageTable.AddGroup("Merge", Loc.T("Gộp", "Merge"));
        merge.Large(B("btnMergeCells", Loc.T("Gộp ô", "Merge\nCells"), WI.Merge(true), Loc.T("Gộp các ô đã chọn", "Merge the selected cells"), CmdMergeCells));
        var style = _pageTable.AddGroup("TableStyle", Loc.T("Kiểu bảng", "Table Style"));
        var cs = B("btnCellShading", Loc.T("Tô nền ô", "Shading"), WI.Shading(true), Loc.T("Màu nền của ô", "Cell background colour"), () => { }, RibbonButtonSize.Large);
        cs.DropDownFactory = () => new ColorPickerDropDown(c => { ApplyShading(c); _ed.Focus(); }, Loc.T("Không màu", "No colour"), _lastShading);
        style.Large(cs);
        style.Stack(B("btnTableBordersAll", Loc.T("Tất cả viền", "All Borders"), WI.Borders(false), Loc.T("Kẻ viền toàn bảng", "Borders on the whole table"), () => SetTableBorders(true)),
            B("btnTableBordersNone", Loc.T("Bỏ viền", "No Borders"), WI.Borders(false), Loc.T("Bỏ viền bảng", "Remove table borders"), () => SetTableBorders(false)),
            B("btnHeaderRow", Loc.T("Hàng tiêu đề", "Header Row"), WI.StyleGlyph(false), Loc.T("Tô đậm và tô nền hàng đầu", "Bold and shade the first row"), CmdHeaderRow));
        var align = _pageTable.AddGroup("CellAlign", Loc.T("Căn trong ô", "Alignment"));
        align.Row(B("btnCellTop", "", WI.CellAlign(false, 0), Loc.T("Căn trên", "Align top"), () => SetCellVAlign(CellVAlign.Top)),
            B("btnCellMiddle", "", WI.CellAlign(false, 1), Loc.T("Căn giữa theo chiều dọc", "Align middle"), () => SetCellVAlign(CellVAlign.Center)),
            B("btnCellBottom", "", WI.CellAlign(false, 2), Loc.T("Căn dưới", "Align bottom"), () => SetCellVAlign(CellVAlign.Bottom)));
        align.Row(B("btnTableCenter", "", WI.G(WI.AlignCenter, false), Loc.T("Căn bảng giữa trang", "Center the table on the page"), () => SetTableAlign(ParaAlign.Center)),
            B("btnTableLeft", "", WI.G(WI.AlignLeft, false), Loc.T("Căn bảng trái", "Align the table left"), () => SetTableAlign(ParaAlign.Left)));
    }

    private void BuildPictureTab()
    {
        _pagePicture = _ribbon.AddPage("PictureTools", Loc.T("Ảnh", "Picture"), contextual: true);
        var size = _pagePicture.AddGroup("PicSize", Loc.T("Kích thước", "Size"));
        _numPicW = RNum("numPicWidth", 0.2m, 60, 2, 0.1m, v => { var img = _ed.SelectedImage; if (img != null) { float w = PageSetup.Mm((float)v * 10); _ed.ResizeSelectedImage(w, img.Height * w / Math.Max(1, img.Width)); } });
        _numPicH = RNum("numPicHeight", 0.2m, 60, 2, 0.1m, v => { var img = _ed.SelectedImage; if (img != null) { float h = PageSetup.Mm((float)v * 10); _ed.ResizeSelectedImage(img.Width * h / Math.Max(1, img.Height), h); } });
        size.Row(RL(Loc.T("Rộng (cm)", "Width (cm)")), _numPicW);
        size.Row(RL(Loc.T("Cao (cm)  ", "Height (cm)")), _numPicH);
        var adj = _pagePicture.AddGroup("PicAdjust", Loc.T("Điều chỉnh", "Adjust"));
        adj.Large(B("btnPicReset", Loc.T("Kích thước\ngốc", "Reset\nSize"), WI.ResetPicture(true), Loc.T("Đặt lại kích thước gốc của ảnh", "Reset the picture to its original size"), () => { var img = _ed.SelectedImage; if (img != null) { float k = Math.Min(1, _ed.Document.Page.ContentWidth / img.NaturalWidth); _ed.ResizeSelectedImage(img.NaturalWidth * k, img.NaturalHeight * k); } }));
        adj.Large(B("btnPicCenter", Loc.T("Căn giữa", "Center"), WI.G(WI.AlignCenter, true), Loc.T("Căn giữa ảnh trên dòng", "Center the picture"), () => _ed.ApplyPara(p => p with { Align = ParaAlign.Center })));
        adj.Large(B("btnPicDelete", Loc.T("Xóa ảnh", "Delete"), WI.G(WI.Delete, true), Loc.T("Xóa ảnh đã chọn", "Delete the selected picture"), () => _ed.DeleteSelection()));
    }

    // ------------------------------------------------------------------ status bar

    private ToolStrip BuildStatusBar()
    {
        _statusBar = new ToolStrip
        {
            Dock = DockStyle.Bottom, GripStyle = ToolStripGripStyle.Hidden, Renderer = new RibbonRenderer(false), BackColor = Theme.TabStripBack, Name = "statusBar",
            Padding = new Padding(Dpi.Scale(8), Dpi.Scale(1), Dpi.Scale(8), Dpi.Scale(1)), ImageScalingSize = new Size(Dpi.Scale(16), Dpi.Scale(16))
        };
        _lblPage = new ToolStripLabel { Name = "lblPage", ForeColor = Theme.SubtleInk };
        _lblPage.Click += (_, _) => ShowFind(2);
        _lblWords = new ToolStripLabel { Name = "lblWords", ForeColor = Theme.SubtleInk, Margin = new Padding(Dpi.Scale(12), 0, 0, 0) };
        _lblWords.Click += (_, _) => CmdWordCount();
        _lblLang = new ToolStripLabel { Name = "lblLanguage", ForeColor = Theme.SubtleInk, Margin = new Padding(Dpi.Scale(12), 0, 0, 0) };
        _lblLang.Click += (_, _) => CmdSpelling();
        _statusBar.Items.AddRange([_lblPage, _lblWords, _lblLang]);
        _lblZoom = new ToolStripLabel("100%") { Name = "lblZoom", Alignment = ToolStripItemAlignment.Right, AutoSize = false, Width = Dpi.Scale(44) };
        _lblZoom.Click += (_, _) => CmdZoomDialog();
        var zin = new ToolStripButton("+") { Name = "sbZoomIn", Alignment = ToolStripItemAlignment.Right, ToolTipText = Loc.T("Phóng to", "Zoom in") };
        zin.Click += (_, _) => _ed.Zoom = (float)Math.Round(_ed.Zoom + 0.1f, 2);
        _zoomSlider = new ZoomSlider();
        _zoomSlider.ValueChanged += (_, _) => _ed.Zoom = _zoomSlider.Value;
        var host = new ToolStripControlHost(_zoomSlider) { Alignment = ToolStripItemAlignment.Right, Name = "sbZoomSlider", AutoSize = false, Size = _zoomSlider.Size };
        var zout = new ToolStripButton("−") { Name = "sbZoomOut", Alignment = ToolStripItemAlignment.Right, ToolTipText = Loc.T("Thu nhỏ", "Zoom out") };
        zout.Click += (_, _) => _ed.Zoom = (float)Math.Round(_ed.Zoom - 0.1f, 2);
        var pw = new ToolStripButton(WI.PageWidth(false)) { Name = "sbPageWidth", Alignment = ToolStripItemAlignment.Right, ToolTipText = Loc.T("Vừa chiều rộng", "Page width"), DisplayStyle = ToolStripItemDisplayStyle.Image };
        pw.Click += (_, _) => _ed.ZoomMode = ZoomMode.PageWidth;
        _statusBar.Items.AddRange([_lblZoom, zin, host, zout, new ToolStripSeparator { Alignment = ToolStripItemAlignment.Right }, pw]);
        foreach (var it in new ToolStripItem[] { _lblPage, _lblWords, _lblLang, _lblZoom, zin, zout, pw, host }) _items[it.Name!] = it;
        return _statusBar;
    }

    // ------------------------------------------------------------------ state

    private void UpdateUi()
    {
        if (IsDisposed) return;
        _updating = true;
        try
        {
            UpdateTitle();
            var cp = _ed.CurrentCharProps;
            var pp = _ed.CurrentParaProps;
            void Chk(string n, bool v) { if (_cmd.TryGetValue(n, out var c) && c is RibbonButton b) b.Checked = v; }
            Chk("btnBold", cp.Bold); Chk("btnItalic", cp.Italic); Chk("btnUnderline", cp.Underline != UnderlineStyle.None); Chk("btnStrike", cp.Strike);
            Chk("btnSubscript", cp.VertAlign == VertAlign.Subscript); Chk("btnSuperscript", cp.VertAlign == VertAlign.Superscript);
            Chk("btnAlignLeft", pp.Align == ParaAlign.Left); Chk("btnAlignCenter", pp.Align == ParaAlign.Center); Chk("btnAlignRight", pp.Align == ParaAlign.Right); Chk("btnJustify", pp.Align == ParaAlign.Justify);
            Chk("btnBullets", pp.List == ListType.Bullet); Chk("btnNumbering", pp.List is not ListType.None and not ListType.Bullet);
            Chk("btnShowMarks", _ed.ShowMarks); Chk("btnRuler", _ruler.Visible); Chk("btnNavPane", _navPane.Visible); Chk("btnAiAssistant", _aiPane is { Visible: true }); Chk("btnPrintLayout", true);
            Chk("btnGridlines", _ed.ShowTableGridlines); Chk("btnSpellAsYouType", _spell.Enabled); Chk("btnFormatPainter", _painter != null); Chk("btnFullScreen", _fullScreen);
            if (!_cboFont.Focused)
            {
                if (_cboFont.Text != cp.FontName) _cboFont.Text = cp.FontName;
                if (_cboFont.IsHandleCreated) _cboFont.Select(0, 0);
            }
            if (!_cboSize.Focused)
            {
                var sz = cp.Size.ToString("0.#", CultureInfo.InvariantCulture);
                if (_cboSize.Text != sz) _cboSize.Text = sz;
                if (_cboSize.IsHandleCreated) _cboSize.Select(0, 0);
            }
            _gallery.RefreshState(pp.Style);
            if (!_numIndL.Focused) _numIndL.Value = Math.Clamp((decimal)Math.Round(PageSetup.ToMm(pp.IndentLeft) / 10, 2), _numIndL.Minimum, _numIndL.Maximum);
            if (!_numIndR.Focused) _numIndR.Value = Math.Clamp((decimal)Math.Round(PageSetup.ToMm(pp.IndentRight) / 10, 2), _numIndR.Minimum, _numIndR.Maximum);
            if (!_numBefore.Focused) _numBefore.Value = Math.Clamp((decimal)pp.SpaceBefore, _numBefore.Minimum, _numBefore.Maximum);
            if (!_numAfter.Focused) _numAfter.Value = Math.Clamp((decimal)pp.SpaceAfter, _numAfter.Minimum, _numAfter.Maximum);
            _items["qatUndo"].Enabled = _ed.CanUndo;
            _items["qatRedo"].Enabled = _ed.CanRedo;
            bool inTable = _ed.CurrentCell != null;
            _ribbon.SetContextual(_pageTable, inTable);
            var img = _ed.SelectedImage;
            bool picSelected = img != null;
            if (picSelected != _pagePicture.Header.Visible) _ribbon.SetContextual(_pagePicture, picSelected, select: picSelected);
            if (img != null)
            {
                _numPicW.Value = Math.Clamp((decimal)Math.Round(PageSetup.ToMm(img.Width) / 10, 2), _numPicW.Minimum, _numPicW.Maximum);
                _numPicH.Value = Math.Clamp((decimal)Math.Round(PageSetup.ToMm(img.Height) / 10, 2), _numPicH.Minimum, _numPicH.Maximum);
            }
            int pages = _ed.PageCount, cur = _ed.CurrentPage + 1;
            _lblPage.Text = Loc.F("Trang {0}/{1}", "Page {0} of {1}", cur, pages);
            var (words, _, _, _) = _ed.HasSelection ? Editing.Count(_ed.Document, _ed.SelStart, _ed.SelEnd) : Editing.Count(_ed.Document);
            _lblWords.Text = _ed.HasSelection ? Loc.F("{0} từ được chọn", "{0} words selected", words) : Loc.F("{0} từ", "{0} words", words);
            _lblLang.Text = Loc.T("Chính tả: ", "Proofing: ") + (!_spell.Available ? Loc.T("không có", "none") : _spell.Language.StartsWith("vi") ? "Tiếng Việt" : "English")
                + (_spell.Available && !_spell.Enabled ? Loc.T(" (tắt)", " (off)") : _spell.Suppressed ? Loc.T(" (chưa có từ điển tiếng Việt)", " (no Vietnamese dictionary)") : "");
            UpdateZoomUi();
        }
        finally { _updating = false; }
    }

    private void UpdateZoomUi()
    {
        _lblZoom.Text = Math.Round(_ed.Zoom * 100) + "%";
        _zoomSlider.Value = _ed.Zoom;
    }

    // ------------------------------------------------------------------ commands

    private void ApplyFontName(string name)
    {
        name = name.Trim();
        if (name.Length == 0) return;
        _ed.ApplyChar(p => p with { FontName = name });
        _ed.Focus();
    }

    private void ApplyFontSize(string text)
    {
        if (float.TryParse(text.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var s) && s >= 1 && s <= 1638)
            _ed.ApplyChar(p => p with { Size = s });
        else UpdateUi();
        _ed.Focus();
    }

    private void SetList(ListType t)
    {
        if (_ed.CurrentParaProps.List == ListType.None) _ed.ToggleList(t);
        else _ed.ApplyPara(p => p with { List = t });
    }

    private void LineSpacing(float v) => _ed.ApplyPara(p => p with { LineSpacing = v, LineRule = LineRule.Multiple });

    private void ClearFormatting()
    {
        var normal = _ed.Document.Styles[StyleIds.Normal];
        if (_ed.HasSelection) _ed.ApplyChar(_ => normal.Char);
        else _ed.ApplyStyle(StyleIds.Normal);
    }

    private void ApplyShading(Color? c)
    {
        var cell = _ed.CurrentCell;
        if (cell != null)
        {
            var (a, b) = _ed.SelectedCells();
            var cells = CellsBetween(a ?? cell, b ?? cell);
            _ed.Edit("shading", () => { foreach (var x in cells) x.Shading = c; });
        }
        else _ed.ApplyPara(p => p with { Shading = c });
    }

    private static List<TableCell> CellsBetween(TableCell a, TableCell b)
    {
        var t = a.Row?.Table;
        if (t == null || t != b.Row?.Table) return [a];
        var all = t.Rows.SelectMany(r => r.Cells).ToList();
        int i = all.IndexOf(a), j = all.IndexOf(b);
        if (i > j) (i, j) = (j, i);
        int r1 = t.Rows.IndexOf(a.Row!), r2 = t.Rows.IndexOf(b.Row!);
        int g1 = t.GridColumnOf(a.Row!, a), g2 = t.GridColumnOf(b.Row!, b);
        if (r1 > r2) (r1, r2) = (r2, r1);
        if (g1 > g2) (g1, g2) = (g2, g1);
        var list = new List<TableCell>();
        for (int r = r1; r <= r2; r++)
            foreach (var c in t.Rows[r].Cells) { int g = t.GridColumnOf(t.Rows[r], c); if (g >= g1 && g <= g2) list.Add(c); }
        return list;
    }

    private void SetTableBorders(bool on)
    {
        var t = _ed.CurrentCell?.Row?.Table;
        if (t == null) return;
        _ed.Edit("table", () => t.Borders = on);
    }

    private void SetTableAlign(ParaAlign a)
    {
        var t = _ed.CurrentCell?.Row?.Table;
        if (t == null) return;
        _ed.Edit("table", () => t.Align = a);
    }

    private void SetCellVAlign(CellVAlign v)
    {
        var cell = _ed.CurrentCell;
        if (cell == null) return;
        var (a, b) = _ed.SelectedCells();
        var cells = CellsBetween(a ?? cell, b ?? cell);
        _ed.Edit("table", () => { foreach (var c in cells) c.VAlign = v; });
    }

    private void CmdHeaderRow()
    {
        var t = _ed.CurrentCell?.Row?.Table;
        if (t == null) return;
        _ed.Edit("table", () =>
        {
            var row = t.Rows[0];
            row.HeaderRow = true;
            foreach (var c in row.Cells)
            {
                c.Shading ??= Color.FromArgb(221, 235, 247);
                foreach (var p in c.Paragraphs) { foreach (var r in p.Runs) r.Props = r.Props with { Bold = true }; p.MarkProps = p.MarkProps with { Bold = true }; p.Normalize(); }
            }
        });
    }

    private void CmdMergeCells()
    {
        var (a, b) = _ed.SelectedCells();
        if (a == null || b == null || a == b) { Msg.Info(this, Loc.T("Hãy bôi đen từ ô đầu đến ô cuối cần gộp (trong cùng một bảng).", "Select from the first to the last cell to merge (in the same table).")); return; }
        _ed.TableEdit(_ => Editing.MergeCells(_ed.Document, a, b));
    }

    private void CmdInsertTableDialog()
    {
        using var d = new InsertTableDialog();
        if (d.ShowDialog(this) == DialogResult.OK) _ed.InsertTable(d.Rows, d.Cols);
        _ed.Focus();
    }

    private void CmdInsertPicture()
    {
        using var d = new OpenFileDialog { Title = Loc.T("Chèn ảnh", "Insert Picture"), Filter = Loc.T("Ảnh", "Pictures") + " (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp" };
        if (d.ShowDialog(this) != DialogResult.OK) return;
        try { _ed.InsertImageFile(d.FileName); }
        catch (Exception ex) { Msg.Error(this, Loc.T("Không đọc được ảnh: ", "Could not read the picture: ") + ex.Message); }
        _ed.Focus();
    }

    public void CmdHyperlink()
    {
        var sel = _ed.SelectedText;
        var cur = _ed.CurrentCharProps.Link;
        using var d = new HyperlinkDialog(sel.Length > 0 ? sel : cur ?? "", cur);
        if (d.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(d.Url)) { _ed.Focus(); return; }
        var url = d.Url;
        var linkProps = new Func<CharProps, CharProps>(p => p with { Link = url, Color = Color.FromArgb(5, 99, 193), Underline = UnderlineStyle.Single });
        if (_ed.HasSelection && d.LinkText == sel) _ed.ApplyChar(linkProps);
        else
        {
            var text = string.IsNullOrWhiteSpace(d.LinkText) ? url : d.LinkText;
            var props = linkProps(_ed.TypingProps);
            _ed.Edit("link", () =>
            {
                var pos = _ed.HasSelection ? Editing.Delete(_ed.Document, _ed.SelStart, _ed.SelEnd) : _ed.Caret;
                var end = Editing.InsertInline(_ed.Document, pos, text, props, null);
                _ed.SetCaretInternal(end);
            });
        }
        _ed.Focus();
    }

    private void CmdHeaderFooter(bool header)
    {
        var list = header ? _ed.Document.Header : _ed.Document.Footer;
        var text = string.Join("\n", list.Select(p => p.Text.Replace(Run.ObjectChar.ToString(), "").TrimEnd()));
        bool hasPage = list.Any(p => p.Runs.Any(r => r.Object is FieldObject));
        var align = list.FirstOrDefault()?.Props.Align ?? ParaAlign.Center;
        using var d = new HeaderFooterDialog(header, text.Trim(), align, hasPage);
        if (d.ShowDialog(this) != DialogResult.OK) { _ed.Focus(); return; }
        var normal = _ed.Document.Styles[StyleIds.Normal];
        var cp = normal.Char with { Size = Math.Max(8, normal.Char.Size - 2) };
        var pp = new ParaProps { Align = d.Align, SpaceAfter = 0, LineSpacing = 1f };
        var paras = d.Text2.Replace("\r\n", "\n").Split('\n').Where((s, i) => s.Length > 0 || i == 0).Select(s => new Paragraph(s, cp, pp)).ToList();
        if (d.PageNumber)
        {
            var last = paras[^1];
            if (last.Length > 0) last.Runs.Add(new Run(" ", cp));
            last.Runs.Add(new Run(Run.ObjectChar.ToString(), cp, new FieldObject(FieldKind.Page)));
            last.Normalize();
        }
        if (paras.All(p => p.Length == 0)) paras.Clear();
        _ed.DocumentEdit(doc => { if (header) doc.Header = paras; else doc.Footer = paras; });
        _ed.Focus();
    }

    private void SetPageNumber(bool top, ParaAlign align, bool xOfY)
    {
        var normal = _ed.Document.Styles[StyleIds.Normal];
        var cp = normal.Char with { Size = Math.Max(8, normal.Char.Size - 2) };
        var p = new Paragraph("", cp, new ParaProps { Align = align, SpaceAfter = 0, LineSpacing = 1f });
        if (xOfY) p.Runs.Add(new Run(Loc.T("Trang ", "Page "), cp));
        p.Runs.Add(new Run(Run.ObjectChar.ToString(), cp, new FieldObject(FieldKind.Page)));
        if (xOfY) { p.Runs.Add(new Run(" / ", cp)); p.Runs.Add(new Run(Run.ObjectChar.ToString(), cp, new FieldObject(FieldKind.NumPages))); }
        _ed.DocumentEdit(d =>
        {
            RemoveFields(d.Header); RemoveFields(d.Footer);
            (top ? d.Header : d.Footer).Add(p);
        });
    }

    private static void RemoveFields(List<Paragraph> list)
    {
        list.RemoveAll(p => p.Runs.Any(r => r.Object is FieldObject) && p.Runs.All(r => r.Object is FieldObject || r.Text.Trim().Length == 0 || r.Text.Trim() is "/" or "Trang" or "Page"));
        foreach (var p in list) { p.Runs.RemoveAll(r => r.Object is FieldObject); p.Normalize(); }
    }

    private void RemovePageNumbers() => _ed.DocumentEdit(d => { RemoveFields(d.Header); RemoveFields(d.Footer); });

    private void CmdDateTime()
    {
        using var d = new DateTimeDialog();
        if (d.ShowDialog(this) == DialogResult.OK && d.Selected != null) _ed.InsertText(d.Selected, "insert");
        _ed.Focus();
    }

    private void CmdSymbol()
    {
        using var d = new SymbolDialog();
        if (d.ShowDialog(this) == DialogResult.OK && d.Selected != null) _ed.InsertText(d.Selected, "insert");
        _ed.Focus();
    }

    private void CmdHorizontalLine()
    {
        _ed.Edit("hline", () =>
        {
            var doc = _ed.Document;
            var pos = _ed.HasSelection ? Editing.Delete(doc, _ed.SelStart, _ed.SelEnd) : _ed.Caret;
            var para = doc.Paragraphs[pos.Para];
            if (para.Length > 0) pos = Editing.SplitParagraph(doc, new DocPos(pos.Para, para.Length));
            var p = doc.Paragraphs[pos.Para];
            float w = p.Cell != null ? 100 : doc.Page.ContentWidth - p.Props.IndentLeft - p.Props.IndentRight - 1;
            Editing.InsertObject(doc, new DocPos(pos.Para, 0), new ImageObject(HLineImage(), "png", w, 2), p.MarkProps with { Link = null });
            var next = Editing.SplitParagraph(doc, new DocPos(pos.Para, 1));
            _ed.SetCaretInternal(next);
        });
        _ed.Focus();
    }

    private static byte[] HLineImage()
    {
        using var bmp = new Bitmap(400, 3);
        using (var g = Graphics.FromImage(bmp)) { g.Clear(Color.White); using var p = new Pen(Color.FromArgb(120, 120, 120), 1); g.DrawLine(p, 0, 1, 400, 1); }
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        return ms.ToArray();
    }

    private void SetMargins(float t, float b, float l, float r) => _ed.DocumentEdit(d =>
    {
        d.Page.MarginTop = PageSetup.Mm(t); d.Page.MarginBottom = PageSetup.Mm(b); d.Page.MarginLeft = PageSetup.Mm(l); d.Page.MarginRight = PageSetup.Mm(r);
    });

    private void SetOrientation(bool landscape) => _ed.DocumentEdit(d =>
    {
        if (d.Page.Landscape == landscape) return;
        (d.Page.Width, d.Page.Height) = (d.Page.Height, d.Page.Width);
    });

    private void SetPaper(string name, float w, float h) => _ed.DocumentEdit(d =>
    {
        bool land = d.Page.Landscape;
        d.Page.Width = land ? h : w; d.Page.Height = land ? w : h; d.Page.PaperName = name;
    });

    private void SetWatermark(string? text) => _ed.DocumentEdit(d => d.Watermark = text);

    public void CmdPageSetup()
    {
        using var d = new PageSetupDialogEx(_ed.Document.Page);
        if (d.ShowDialog(this) == DialogResult.OK) { var r = d.Result; _ed.DocumentEdit(doc => doc.Page = r); }
        _ed.Focus();
    }

    public void CmdParagraphDialog()
    {
        using var d = new ParagraphDialog(_ed.CurrentParaProps);
        if (d.ShowDialog(this) == DialogResult.OK)
        {
            var r = d.Result;
            _ed.ApplyPara(p => p with
            {
                Align = r.Align, IndentLeft = r.IndentLeft, IndentRight = r.IndentRight, FirstLine = r.FirstLine, SpaceBefore = r.SpaceBefore, SpaceAfter = r.SpaceAfter,
                LineRule = r.LineRule, LineSpacing = r.LineSpacing, PageBreakBefore = r.PageBreakBefore, KeepWithNext = r.KeepWithNext
            });
        }
        _ed.Focus();
    }

    public void CmdFontDialog()
    {
        using var d = new FontDialogEx(_ed.CurrentCharProps);
        if (d.ShowDialog(this) == DialogResult.OK) { var r = d.Result; _ed.ApplyChar(p => r with { Link = p.Link, Highlight = p.Highlight }); }
        _ed.Focus();
    }

    private void CmdSort()
    {
        var a = _ed.SelStart.Para; var b = _ed.SelEnd.Para;
        var ps = _ed.Document.Paragraphs;
        if (a == b || ps[a].Cell != ps[b].Cell) { Msg.Info(this, Loc.T("Hãy chọn nhiều đoạn văn cần sắp xếp.", "Select several paragraphs to sort.")); return; }
        _ed.Edit("sort", () =>
        {
            var list = _ed.Document.Paragraphs.Skip(a).Take(b - a + 1).ToList();
            var sorted = list.OrderBy(p => p.Text, StringComparer.Create(new CultureInfo("vi-VN"), true)).Select(p => p.Clone()).ToList();
            var container = list[0].Cell?.Paragraphs;
            for (int i = 0; i < list.Count; i++)
            {
                list[i].Runs = sorted[i].Runs; list[i].Props = sorted[i].Props; list[i].MarkProps = sorted[i].MarkProps; list[i].Touch();
            }
        });
    }

    private void ToggleMarks() { _ed.ShowMarks = !_ed.ShowMarks; UpdateUi(); }

    // format painter
    private void StartFormatPainter(bool sticky)
    {
        if (_painter != null && !sticky) { _painter = null; _ed.Cursor = Cursors.IBeam; UpdateUi(); return; }
        _painter = (_ed.CurrentCharProps, _ed.CurrentParaProps, sticky);
        UpdateUi();
        _ed.Focus();
    }

    private void ApplyFormatPainter()
    {
        if (_painter is not { } p || !_ed.HasSelection) return;
        var c = p.C; var pp = p.P;
        _ed.ApplyChar(_ => c);
        _ed.ApplyPara(x => pp with { List = x.List, ListId = x.ListId, ListLevel = x.ListLevel });
        if (!p.Sticky) _painter = null;
        UpdateUi();
    }

    private static (CharProps C, ParaProps P)? _formatClipboard;
    private void CopyFormat() => _formatClipboard = (_ed.CurrentCharProps, _ed.CurrentParaProps);
    private void PasteFormat()
    {
        if (_formatClipboard is not { } f) return;
        _ed.ApplyChar(_ => f.C);
    }

    /// <summary>Applies the format painter programmatically (tests).</summary>
    public void PaintFormatOnSelection() => ApplyFormatPainter();

    private void ShowFind(int tab)
    {
        if (_find == null || _find.IsDisposed)
        {
            _find = new FindReplaceDialog(_ed, tab);
            _find.Show(this);
        }
        else
        {
            _find.Activate();
        }
    }

    public FindReplaceDialog OpenFindDialog(int tab) { ShowFind(tab); return _find!; }

    public void CmdSpelling()
    {
        if (!_spell.Available) { Msg.Info(this, Loc.T("Máy tính chưa có bộ kiểm tra chính tả của Windows cho ngôn ngữ này.\nVào Cài đặt Windows > Thời gian và ngôn ngữ > Ngôn ngữ để cài thêm.", "No Windows spell checker is installed for this language.\nInstall one in Windows Settings > Time & language > Language.")); return; }
        _spell.Enabled = true;
        using var d = new SpellingDialog(_ed, _spell);
        d.ShowDialog(this);
        UpdateUi();
    }

    public void CmdWordCount()
    {
        using var d = new WordCountDialog(_ed.Document, _ed.HasSelection ? _ed.SelStart : null, _ed.HasSelection ? _ed.SelEnd : null, _ed.PageCount);
        d.ShowDialog(this);
    }

    private void CmdZoomDialog()
    {
        using var d = new ZoomDialog(_ed.Zoom);
        if (d.ShowDialog(this) == DialogResult.OK) _ed.Zoom = d.Value;
    }

    private void ShowHelp()
    {
        using var d = new WordHelpDialog();
        d.ShowDialog(this);
    }

    // ------------------------------------------------------------------ table of contents

    private bool IsTocPara(Paragraph p) => p.Props.Style is StyleIds.Toc1 or StyleIds.Toc2 or StyleIds.Toc3 or StyleIds.TocHeading;

    public void InsertToc(bool updateOnly = false)
    {
        var doc = _ed.Document;
        var headings = doc.Paragraphs.Where(p => p.Cell == null && p.Props.OutlineLevel is >= 1 and <= 3 && !IsTocPara(p) && p.Text.Trim().Length > 0).ToList();
        if (headings.Count == 0)
        {
            Msg.Info(this, Loc.T("Chưa có tiêu đề nào. Hãy áp dụng kiểu Tiêu đề 1, 2, 3 cho các đề mục (thẻ Trang chủ › Kiểu) rồi chèn lại mục lục.",
                "No headings found. Apply Heading 1–3 styles to your titles (Home › Styles) and insert the table again."));
            return;
        }
        bool exists = doc.Blocks.OfType<Paragraph>().Any(IsTocPara);
        if (updateOnly && !exists) { InsertToc(false); return; }
        _ed.Edit("toc", () =>
        {
            int at;
            var existing = doc.Blocks.OfType<Paragraph>().Where(IsTocPara).ToList();
            if (existing.Count > 0)
            {
                at = doc.Blocks.IndexOf(existing[0]);
                foreach (var p in existing) doc.Blocks.Remove(p);
            }
            else
            {
                var caretPara = doc.Paragraphs[_ed.Caret.Para];
                at = doc.BlockIndexOf(caretPara);
                if (at < 0) at = 0;
            }
            var toc = new List<Paragraph>();
            var hs = doc.Styles[StyleIds.TocHeading];
            toc.Add(new Paragraph(Loc.T("MỤC LỤC", "CONTENTS"), hs.Char, hs.Para));
            foreach (var h in headings)
            {
                var st = doc.Styles[h.Props.OutlineLevel == 1 ? StyleIds.Toc1 : h.Props.OutlineLevel == 2 ? StyleIds.Toc2 : StyleIds.Toc3];
                toc.Add(new Paragraph(h.Text.Replace(Run.ObjectChar.ToString(), "").Replace(Run.LineBreak, ' ').Replace(Run.PageBreak.ToString(), "").Trim() + "\t0", st.Char, st.Para));
            }
            doc.Blocks.InsertRange(at, toc);
            doc.StructureChanged();
            // two passes: the TOC itself can move headings to later pages
            for (int pass = 0; pass < 2; pass++)
            {
                var layout = new DocLayout(doc, _ed.Measurer);
                for (int i = 0; i < headings.Count; i++)
                {
                    int page = layout.ByPara.TryGetValue(headings[i], out var lines) && lines.Count > 0 ? lines[0].Page + 1 : 1;
                    var p = toc[i + 1];
                    var text = p.Text;
                    int tab = text.LastIndexOf('\t');
                    Editing.DeleteInPara(p, tab + 1, p.Length);
                    p.Runs.Add(new Run(page.ToString(), p.Runs.Count > 0 ? p.Runs[^1].Props : p.MarkProps));
                    p.Normalize();
                }
            }
        });
    }

    private void RemoveToc()
    {
        var doc = _ed.Document;
        if (!doc.Blocks.OfType<Paragraph>().Any(IsTocPara)) return;
        _ed.Edit("toc", () =>
        {
            doc.Blocks.RemoveAll(b => b is Paragraph p && IsTocPara(p));
            doc.StructureChanged();
        });
    }
}

/// <summary>Usage guide and keyboard shortcuts.</summary>
public sealed class WordHelpDialog : DialogBase
{
    public WordHelpDialog() : base(Loc.T("Hướng dẫn Microads Office Word", "Microads Office Word guide"), 640, 560)
    {
        Name = "wordHelpDialog";
        var box = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Color.White, Name = "txtHelp", Font = new Font("Segoe UI", 9.5f) };
        box.Text = Loc.T(
            "BẮT ĐẦU\n  • Tệp ▸ Mới: tài liệu trống hoặc mẫu (công văn hành chính NĐ 30/2020, thư, sơ yếu lý lịch, báo cáo).\n  • Mở .docx, .rtf, .txt; Lưu .docx (mặc định), .pdf, .rtf, .html, .txt. Tự lưu bản phục hồi mỗi 2 phút.\n\n" +
            "SOẠN THẢO\n  • Gõ tiếng Việt bằng Unikey/EVKey hoặc bàn phím Telex của Windows.\n  • Bấm đúp chọn từ, bấm 3 lần chọn đoạn; Shift + mũi tên mở rộng vùng chọn; Ctrl + mũi tên nhảy theo từ.\n" +
            "  • Enter: đoạn mới · Shift+Enter: xuống dòng · Ctrl+Enter: ngắt trang · Tab trong danh sách: lùi cấp.\n\n" +
            "ĐỊNH DẠNG\n  • Ctrl+B đậm · Ctrl+I nghiêng · Ctrl+U gạch chân · Ctrl+] / Ctrl+[ tăng/giảm cỡ · Ctrl+= chỉ số dưới · Ctrl+Shift+= chỉ số trên.\n" +
            "  • Ctrl+L/E/R/J: căn trái/giữa/phải/đều · Ctrl+M: thụt lề · Ctrl+D: hộp thoại Phông chữ · Ctrl+Space: xóa định dạng.\n" +
            "  • Sao định dạng: bấm 1 lần dùng 1 lần, bấm đúp để dùng nhiều lần (Esc để tắt). Ctrl+Shift+C / Ctrl+Shift+V.\n  • Kiểu Tiêu đề 1–3 dùng cho Mục lục tự động (thẻ Tham chiếu) và Ngăn điều hướng.\n\n" +
            "BẢNG & ẢNH\n  • Chèn ▸ Bảng: kéo chọn số hàng × cột. Trong bảng: Tab sang ô sau, Tab ở ô cuối thêm hàng. Thẻ Bảng: chèn/xóa hàng cột, gộp ô, tô nền, viền.\n" +
            "  • Chèn ▸ Ảnh hoặc dán ảnh (Ctrl+V), kéo-thả ảnh vào trang. Bấm ảnh để chọn, kéo 8 ô vuông để đổi kích thước.\n\n" +
            "TRANG\n  • Bố cục: lề (có lề văn bản hành chính VN), hướng, khổ giấy, hình mờ, màu trang. Chèn ▸ Đầu/Chân trang, Số trang.\n\n" +
            "TÌM & KHÁC\n  • Ctrl+F tìm · Ctrl+H thay thế · Ctrl+G đi tới trang · Ctrl+K liên kết (Ctrl+bấm để mở) · F7 chính tả · Ctrl+P in · F11 toàn màn hình · Ctrl+F1 thu gọn Ribbon.\n\n" +
            "Microads Office — phát hành bởi Công ty TNHH Công nghệ Microads, luôn luôn miễn phí. microads.vn",
            "GETTING STARTED\n  • File ▸ New: blank document or templates (Vietnamese administrative letter, letter, resume, report).\n  • Open .docx, .rtf, .txt; Save as .docx (default), .pdf, .rtf, .html, .txt. AutoRecover saves every 2 minutes.\n\n" +
            "EDITING\n  • Double-click selects a word, triple-click a paragraph; Shift + arrows extend; Ctrl + arrows move by word.\n" +
            "  • Enter: new paragraph · Shift+Enter: line break · Ctrl+Enter: page break · Tab in a list: demote.\n\n" +
            "FORMATTING\n  • Ctrl+B bold · Ctrl+I italic · Ctrl+U underline · Ctrl+] / Ctrl+[ size · Ctrl+= subscript · Ctrl+Shift+= superscript.\n" +
            "  • Ctrl+L/E/R/J alignment · Ctrl+M indent · Ctrl+D Font dialog · Ctrl+Space clear formatting.\n" +
            "  • Format Painter: click once for one use, double-click to keep it (Esc to stop). Ctrl+Shift+C / Ctrl+Shift+V.\n  • Heading 1–3 styles feed the automatic Table of Contents (References) and the Navigation pane.\n\n" +
            "TABLES & PICTURES\n  • Insert ▸ Table: drag the grid. In a table Tab moves to the next cell; Tab in the last cell adds a row. Table tab: rows, columns, merge, shading, borders.\n" +
            "  • Insert ▸ Picture, paste (Ctrl+V) or drag pictures in. Click a picture and drag its handles to resize.\n\n" +
            "PAGE\n  • Layout: margins, orientation, size, watermark, page colour. Insert ▸ Header, Footer, Page numbers.\n\n" +
            "FIND & MORE\n  • Ctrl+F find · Ctrl+H replace · Ctrl+G go to page · Ctrl+K link (Ctrl+click opens) · F7 spelling · Ctrl+P print · F11 full screen · Ctrl+F1 collapse ribbon.\n\n" +
            "Microads Office — published by Microads Technology Co., Ltd., always free. microads.vn");
        Body.Controls.Add(box);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
    }
}
