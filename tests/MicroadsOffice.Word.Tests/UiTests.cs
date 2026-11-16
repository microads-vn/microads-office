using System.Reflection;
using System.Runtime.InteropServices;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Editor;
using MicroadsOffice.Word.Model;
using MicroadsOffice.Word.UI;

namespace MicroadsOffice.Word.Tests;

/// <summary>
/// In-process UI tests: the real main window is created and every ribbon button / menu item is invoked
/// through its click handler (no mouse is moved). Modal dialogs are closed automatically (or accepted
/// for the scenarios that need it).
/// </summary>
public static class UiTests
{
    // ------------------------------------------------------------------ dialog auto-closer

    private delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] private static extern bool EnumThreadWindows(uint threadId, EnumWindowsProc cb, IntPtr l);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

    /// <summary>Handlers for named dialogs; the default is Cancel.</summary>
    public static readonly Dictionary<string, Action<Form>> DialogHandlers = new();
    public static readonly List<string> SeenDialogs = new();
    public static int NativeDialogsClosed;

    private static System.Windows.Forms.Timer StartCloser()
    {
        var t = new System.Windows.Forms.Timer { Interval = 200 };
        t.Tick += (_, _) =>
        {
            foreach (var f in Application.OpenForms.Cast<Form>().ToList())
            {
                if (!f.Modal || !f.Visible) continue;
                SeenDialogs.Add(f.Name);
                if (DialogHandlers.TryGetValue(f.Name, out var h)) { DialogHandlers.Remove(f.Name); h(f); }
                else { f.DialogResult = DialogResult.Cancel; f.Close(); }
            }
            uint tid = GetCurrentThreadId();
            EnumThreadWindows(tid, (h, _) =>
            {
                var sb = new System.Text.StringBuilder(64);
                GetClassName(h, sb, 64);
                if (sb.ToString() == "#32770" && IsWindowVisible(h))
                {
                    PostMessage(h, 0x0111, (IntPtr)2, IntPtr.Zero); // WM_COMMAND IDCANCEL
                    PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
                    NativeDialogsClosed++;
                }
                return true;
            }, IntPtr.Zero);
        };
        t.Start();
        return t;
    }

    private static readonly MethodInfo OnClick = typeof(Control).GetMethod("OnClick", BindingFlags.NonPublic | BindingFlags.Instance)!;

    private static void Click(Control c)
    {
        if (c is RibbonButton rb)
        {
            ToolStripDropDown? shown = null;
            var factory = rb.DropDownFactory;
            if (factory != null) rb.DropDownFactory = () => shown = factory();
            OnClick.Invoke(rb, [EventArgs.Empty]);
            Pump();
            shown?.Close();
            rb.DropDown?.Close();
            if (factory != null) rb.DropDownFactory = factory;
        }
        else if (c is Button b) OnClick.Invoke(b, [EventArgs.Empty]);
        Pump();
    }

    private static void Pump(int ms = 30)
    {
        var until = DateTime.Now.AddMilliseconds(ms);
        do { Application.DoEvents(); Thread.Sleep(5); } while (DateTime.Now < until);
    }

    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);

    /// <summary>Captures the real window (correct z-order, no mouse) with PrintWindow.</summary>
    private static void Shot(Form f, string name)
    {
        Pump(250);
        using var bmp = new Bitmap(f.Width, f.Height);
        using (var g = Graphics.FromImage(bmp))
        {
            var hdc = g.GetHdc();
            PrintWindow(f.Handle, hdc, 2);
            g.ReleaseHdc(hdc);
        }
        bmp.Save(Path.Combine(T.Out, name + ".png"));
    }

    private static Control? Find(Control root, string name) => root.Controls.Find(name, true).FirstOrDefault();

    public static void Run(Runner r)
    {
        var closer = StartCloser();
        WordMainForm? w = null;
        try
        {
            r.Test("UI01 main window opens on the start screen with templates, recent files and Microads footer", () =>
            {
                w = (WordMainForm)WordApp.CreateMainForm([]);
                w.Show();
                Pump(400);
                T.True(w.BackstageView.Visible && w.BackstageView.CurrentPage == "home", "start screen shown");
                T.True(Find(w, "tpl_blank") != null && Find(w, "tpl_congvan") != null && Find(w, "tpl_cv") != null, "template tiles");
                T.True(Find(w, "promoStrip") != null, "Microads promo strip");
                Shot(w, "word-start");
                ((Button)Find(w.BackstageView, "tpl_blank")!).PerformClick();
                Pump();
                T.False(w.BackstageView.Visible, "blank document opened");
                T.True(w is IOfficeAppWindow, "IOfficeAppWindow");
            });
            if (w == null) return;
            var ed = w.Editor;

            r.Test("UI02 ribbon has all tabs, header items (language flag, update, microads.vn)", () =>
            {
                var names = w.Ribbon.Pages.Select(p => p.Name).ToList();
                foreach (var n in new[] { "pageHome", "pageInsert", "pageLayout", "pageReferences", "pageReview", "pageView", "pageHelp", "pageTableTools", "pagePictureTools" })
                    T.True(names.Contains(n), "tab " + n);
                var header = w.Ribbon.Controls.OfType<ToolStrip>().First();
                T.True(header.Items["btnLanguage"] != null && header.Items["btnWebsite"] != null && header.Items["btnCheckUpdateHeader"] != null, "header items");
                T.Eq("microads.vn", header.Items["btnWebsite"]!.Text);
                foreach (var n in new[] { "btnHelp", "btnCheckUpdate", "btnSettings", "btnAbout", "btnVisitWebsite" }) T.True(w.Commands.ContainsKey(n), "help button " + n);
            });

            r.Test("UI03 typing and character formatting buttons change the document", () =>
            {
                ed.InsertText("Xin chào Microads Office");
                ed.SelectAll();
                Click(w.Commands["btnBold"]);
                T.True(ed.Document.Paragraphs[0].Runs.All(x => x.Props.Bold), "bold");
                T.True(((RibbonButton)w.Commands["btnBold"]).Checked, "bold button checked");
                Click(w.Commands["btnItalic"]); Click(w.Commands["btnUnderline"]); Click(w.Commands["btnStrike"]);
                var cp = ed.Document.Paragraphs[0].Runs[0].Props;
                T.True(cp.Italic && cp.Underline == UnderlineStyle.Single && cp.Strike, "italic/underline/strike");
                Click(w.Commands["btnGrowFont"]);
                T.True(ed.Document.Paragraphs[0].Runs[0].Props.Size > 13, "grow font");
                Click(w.Commands["btnFontColor"]);
                T.True(ed.Document.Paragraphs[0].Runs[0].Props.Color != null, "font colour");
                Click(w.Commands["btnHighlight"]);
                T.True(ed.Document.Paragraphs[0].Runs[0].Props.Highlight != null, "highlight");
                Click(w.Commands["btnClearFormat"]);
                T.False(ed.Document.Paragraphs[0].Runs[0].Props.Bold, "clear formatting");
                var font = (ComboBox)w.Commands["cboFontName"];
                font.Text = "Arial";
                typeof(WordMainForm).GetMethod("ApplyFontName", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(w, ["Arial"]);
                T.Eq("Arial", ed.Document.Paragraphs[0].Runs[0].Props.FontName, "font name combo");
                typeof(WordMainForm).GetMethod("ApplyFontSize", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(w, ["18"]);
                T.Eq(18f, ed.Document.Paragraphs[0].Runs[0].Props.Size, "font size combo");
            });

            r.Test("UI04 paragraph buttons: alignment, lists, indent, spacing, shading, borders, marks", () =>
            {
                ed.SelectAll();
                Click(w.Commands["btnAlignCenter"]); T.Eq(ParaAlign.Center, ed.CurrentParaProps.Align);
                Click(w.Commands["btnJustify"]); T.Eq(ParaAlign.Justify, ed.CurrentParaProps.Align);
                Click(w.Commands["btnBullets"]); T.Eq(ListType.Bullet, ed.CurrentParaProps.List);
                Click(w.Commands["btnNumbering"]); T.Eq(ListType.Decimal, ed.CurrentParaProps.List);
                Click(w.Commands["btnIncreaseIndent"]); T.Eq(1, ed.CurrentParaProps.ListLevel);
                Click(w.Commands["btnDecreaseIndent"]); T.Eq(0, ed.CurrentParaProps.ListLevel);
                Click(w.Commands["btnNumbering"]); T.Eq(ListType.None, ed.CurrentParaProps.List);
                w.MenuItems["miLs15"].PerformClick(); T.Near(1.5f, ed.CurrentParaProps.LineSpacing, 0.01f, "line spacing 1.5");
                Click(w.Commands["btnShading"]); T.True(ed.CurrentParaProps.Shading != null, "shading");
                Click(w.Commands["btnBorders"]); T.True(ed.CurrentParaProps.Border, "border");
                Click(w.Commands["btnShowMarks"]); T.True(ed.ShowMarks, "marks");
                Shot(w, "word-formatting");
                Click(w.Commands["btnShowMarks"]);
                ((NumericUpDown)w.Commands["numSpaceBefore"]).Value = 12;
                T.Near(12, ed.CurrentParaProps.SpaceBefore, 0.01f, "space before numeric");
                ((NumericUpDown)w.Commands["numIndentLeft"]).Value = 1;
                T.Near(PageSetup.Mm(10), ed.CurrentParaProps.IndentLeft, 0.1f, "indent numeric");
            });

            r.Test("UI05 styles gallery, headings and table of contents", () =>
            {
                ed.Document = Document.CreateBlank();
                ed.InsertText("Chương 1\nNội dung một\nChương 2\nNội dung hai");
                ed.SetCaret(new DocPos(0, 0), false);
                Click(w.Commands["style_Heading1"]);
                ed.SetCaret(new DocPos(2, 0), false);
                Click(w.Commands["style_Heading2"]);
                T.Eq(1, ed.Document.Paragraphs[0].Props.OutlineLevel); T.Eq(2, ed.Document.Paragraphs[2].Props.OutlineLevel);
                ed.SetCaret(new DocPos(0, 0), false);
                Click(w.Commands["btnToc"]);
                var toc = ed.Document.Paragraphs.Where(p => p.Props.Style is StyleIds.Toc1 or StyleIds.Toc2).ToList();
                T.Eq(2, toc.Count, "TOC entries");
                T.True(toc[0].Text.StartsWith("Chương 1\t") && char.IsDigit(toc[0].Text[^1]), "entry with page number: " + toc[0].Text);
                Click(w.Commands["btnUpdateToc"]);
                T.Eq(2, ed.Document.Paragraphs.Count(p => p.Props.Style is StyleIds.Toc1 or StyleIds.Toc2), "update keeps one TOC");
                w.MenuItems["miTocRemove"].PerformClick();
                T.Eq(0, ed.Document.Paragraphs.Count(p => p.Props.Style is StyleIds.Toc1 or StyleIds.Toc2), "TOC removed");
                Click(w.Commands["btnNavPane"]);
                Pump(800);
                var nav = (ListBox)Find(w, "lstNavHeadings")!;
                T.Eq(2, nav.Items.Count, "navigation pane headings");
                Click(w.Commands["btnNavPane"]);
            });

            r.Test("UI06 insert: page break, table via dialog, table tools, picture, link, symbol, date, header/footer, page numbers", () =>
            {
                ed.Document = Document.CreateBlank();
                ed.InsertText("Trang một");
                int pages = ed.PageCount;
                Click(w.Commands["btnPageBreak"]);
                T.Eq(pages + 1, ed.PageCount, "page break");
                DialogHandlers["insertTableDialog"] = f => { ((NumericUpDown)Find(f, "numTableRows")!).Value = 2; ((NumericUpDown)Find(f, "numTableCols")!).Value = 3; f.DialogResult = DialogResult.OK; };
                Click(w.Commands["btnInsertTableDialog"]);
                var t = ed.Document.Blocks.OfType<Table>().Single();
                T.Eq(2, t.Rows.Count); T.Eq(3, t.ColumnCount);
                T.True(ed.CurrentCell != null && w.Ribbon.Pages.First(p => p.Name == "pageTableTools").Header.Visible, "table contextual tab");
                Click(w.Commands["btnRowBelow"]); T.Eq(3, t.Rows.Count, "row below");
                Click(w.Commands["btnRowAbove"]); T.Eq(4, t.Rows.Count, "row above");
                Click(w.Commands["btnColRight"]); T.Eq(4, t.ColumnCount, "col right");
                Click(w.Commands["btnColLeft"]); T.Eq(5, t.ColumnCount, "col left");
                Click(w.Commands["btnDeleteCol"]); T.Eq(4, t.ColumnCount, "delete col");
                Click(w.Commands["btnDeleteRow"]); T.Eq(3, t.Rows.Count, "delete row");
                Click(w.Commands["btnHeaderRow"]); T.True(t.Rows[0].HeaderRow && t.Rows[0].Cells[0].Shading != null, "header row");
                Click(w.Commands["btnCellMiddle"]); T.Eq(CellVAlign.Center, ed.CurrentCell!.VAlign);
                Click(w.Commands["btnTableBordersNone"]); T.False(t.Borders, "no borders");
                Click(w.Commands["btnTableBordersAll"]); T.True(t.Borders, "borders");
                var c0 = t.Rows[0].Cells[0]; var c1 = t.Rows[0].Cells[1];
                ed.Select(new DocPos(ed.Document.IndexOf(c0.Paragraphs[0]), 0), new DocPos(ed.Document.IndexOf(c1.Paragraphs[0]), 0));
                Click(w.Commands["btnMergeCells"]); T.Eq(2, t.Rows[0].Cells[0].GridSpan, "merge");
                Click(w.Commands["btnDeleteTable"]); T.Eq(0, ed.Document.Blocks.OfType<Table>().Count(), "delete table");
                ed.SetCaret(Editing.End(ed.Document), false);
                ed.InsertImage(IoTests.Png(200, 100, Color.Teal), "png");
                var img = ed.Document.Paragraphs.SelectMany(p => p.Runs).Select(x => x.Object).OfType<ImageObject>().Single();
                var ip = ed.Document.Paragraphs.First(p => p.Runs.Any(x => x.Object == img));
                ed.Select(new DocPos(ed.Document.IndexOf(ip), 0), new DocPos(ed.Document.IndexOf(ip), 1));
                Pump();
                T.True(ed.SelectedImage == img && w.Ribbon.SelectedPage?.Name == "pagePictureTools", "picture tab selected");
                ((NumericUpDown)w.Commands["numPicWidth"]).Value = 5;
                T.Near(PageSetup.Mm(50), img.Width, 0.5f, "picture width numeric");
                T.Near(img.Width / 2, img.Height, 0.5f, "aspect kept");
                Click(w.Commands["btnPicReset"]); T.Near(img.NaturalWidth, img.Width, 0.5f, "reset size");
                Click(w.Commands["btnPicDelete"]); T.Eq(0, ed.Document.Paragraphs.SelectMany(p => p.Runs).Count(x => x.Object is ImageObject), "picture deleted");
                DialogHandlers["hyperlinkDialog"] = f => { ((TextBox)Find(f, "txtLinkText")!).Text = "Microads"; ((TextBox)Find(f, "txtLinkUrl")!).Text = "microads.vn"; f.DialogResult = DialogResult.OK; };
                Click(w.Commands["btnHyperlink"]);
                T.True(ed.Document.Paragraphs.SelectMany(p => p.Runs).Any(x => x.Props.Link == "https://microads.vn" && x.Text == "Microads"), "hyperlink inserted");
                DialogHandlers["symbolDialog"] = f => ((Button)Find(f, "sym0")!).PerformClick();
                Click(w.Commands["btnSymbol"]);
                T.True(ed.Document.PlainText.Contains('©'), "symbol inserted");
                DialogHandlers["dateTimeDialog"] = f => f.DialogResult = DialogResult.OK;
                Click(w.Commands["btnDateTime"]);
                T.True(ed.Document.PlainText.Contains(DateTime.Now.ToString("dd/MM/yyyy")), "date inserted");
                Click(w.Commands["btnHorizontalLine"]);
                T.True(ed.Document.Paragraphs.SelectMany(p => p.Runs).Any(x => x.Object is ImageObject), "horizontal line");
                w.MenuItems["miPnXofY"].PerformClick();
                T.Eq(2, ed.Document.Footer.SelectMany(p => p.Runs).Count(x => x.Object is FieldObject), "page X of Y");
                w.MenuItems["miPnRemove"].PerformClick();
                T.Eq(0, ed.Document.Footer.SelectMany(p => p.Runs).Count(x => x.Object is FieldObject), "page numbers removed");
                DialogHandlers["headerDialog"] = f => { ((TextBox)Find(f, "txtHeaderFooter")!).Text = "CÔNG TY MICROADS"; ((CheckBox)Find(f, "chkHFPageNumber")!).Checked = true; f.DialogResult = DialogResult.OK; };
                Click(w.Commands["btnHeader"]);
                T.True(ed.Document.Header.Count == 1 && ed.Document.Header[0].Text.StartsWith("CÔNG TY MICROADS") && ed.Document.Header[0].Runs.Any(x => x.Object is FieldObject), "header with page number");
                Shot(w, "word-insert");
            });

            r.Test("UI07 layout: margins, orientation, size, watermark, page colour, page setup dialog", () =>
            {
                w.MenuItems["miMarVN"].PerformClick();
                T.Near(PageSetup.Mm(30), ed.Document.Page.MarginLeft, 0.1f, "VN margins");
                w.MenuItems["miLandscape"].PerformClick(); T.True(ed.Document.Page.Landscape, "landscape");
                w.MenuItems["miPortrait"].PerformClick(); T.False(ed.Document.Page.Landscape, "portrait");
                w.MenuItems["miPaperA5"].PerformClick(); T.Eq("A5", ed.Document.Page.PaperName);
                w.MenuItems["miPaperA4"].PerformClick();
                w.MenuItems["miWmDraft"].PerformClick(); T.True(!string.IsNullOrEmpty(ed.Document.Watermark), "watermark");
                Shot(w, "word-watermark");
                w.MenuItems["miWmRemove"].PerformClick(); T.True(ed.Document.Watermark == null, "watermark removed");
                DialogHandlers["pageSetupDialog"] = f => { ((NumericUpDown)Find(f, "numMarTop")!).Value = 25; ((Button)Find(f, "btnOk")!).PerformClick(); };
                w.CmdPageSetup();
                T.Near(PageSetup.Mm(25), ed.Document.Page.MarginTop, 0.1f, "page setup dialog");
                DialogHandlers["paragraphDialog"] = f => { ((ComboBox)Find(f, "cboParaAlign")!).SelectedIndex = 2; ((Button)Find(f, "btnOk")!).PerformClick(); };
                w.CmdParagraphDialog();
                T.Eq(ParaAlign.Right, ed.CurrentParaProps.Align, "paragraph dialog");
                DialogHandlers["fontDialog"] = f => { ((NumericUpDown)Find(f, "numDlgSize")!).Value = 20; ((Button)Find(f, "btnOk")!).PerformClick(); };
                ed.SelectAll();
                w.CmdFontDialog();
                T.Eq(20f, ed.Document.Paragraphs[0].Runs.First(x => x.Object == null).Props.Size, "font dialog size");
            });

            r.Test("UI08 find / replace / go to", () =>
            {
                ed.Document = Document.CreateBlank();
                ed.InsertText("táo cam táo\nlê táo");
                var f = w.OpenFindDialog(1);
                f.FindText = "táo"; f.ReplaceText = "chuối";
                T.Eq(3, f.Count, "found 3");
                T.True(f.FindNext() && ed.SelectedText == "táo", "find next selects");
                T.Eq(3, f.ReplaceAll(), "replace all");
                T.Eq("chuối cam chuối\nlê chuối", ed.Document.PlainText);
                ed.Undo();
                T.Eq("táo cam táo\nlê táo", ed.Document.PlainText, "replace all is one undo step");
                f.Close();
            });

            r.Test("UI09 view: zoom buttons, slider, ruler, gridlines, full screen, ribbon collapse", () =>
            {
                Click(w.Commands["btnZoom100"]); T.Near(1f, ed.Zoom, 0.001f, "100%");
                Click(w.Commands["btnOnePage"]); T.Eq(ZoomMode.OnePage, ed.ZoomMode);
                Click(w.Commands["btnPageWidth"]); T.Eq(ZoomMode.PageWidth, ed.ZoomMode);
                DialogHandlers["zoomDialog"] = f => { ((NumericUpDown)Find(f, "numZoom")!).Value = 150; f.DialogResult = DialogResult.OK; };
                Click(w.Commands["btnZoom"]); T.Near(1.5f, ed.Zoom, 0.001f, "zoom dialog");
                w.MenuItems["sbZoomOut"].PerformClick(); T.Near(1.4f, ed.Zoom, 0.001f, "status bar zoom out");
                bool ruler = Find(w, "ruler")!.Visible;
                Click(w.Commands["btnRuler"]); T.Eq(!ruler, Find(w, "ruler")!.Visible, "ruler toggle");
                Click(w.Commands["btnRuler"]);
                Click(w.Commands["btnGridlines"]); T.False(ed.ShowTableGridlines, "gridlines off"); Click(w.Commands["btnGridlines"]);
                Click(w.Commands["btnFullScreen"]); T.False(w.Ribbon.Visible, "full screen hides ribbon");
                Click(w.Commands["btnFullScreen"]); T.True(w.Ribbon.Visible, "restored");
                w.Ribbon.SetMinimized(true); T.True(w.Ribbon.Minimized, "collapsed"); w.Ribbon.SetMinimized(false);
                Click(w.Commands["btnZoom100"]);
            });

            r.Test("UI10 ruler drag changes indents; format painter copies formatting", () =>
            {
                ed.Document = Document.CreateBlank();
                ed.InsertText("đoạn một\nđoạn hai");
                ed.SetCaret(new DocPos(0, 0), false);
                ((RulerControl)Find(w, "ruler")!).SimulateDrag(2, 36);
                T.Near(36, ed.CurrentParaProps.IndentLeft, 0.1f, "left indent dragged");
                ed.Select(new DocPos(0, 0), new DocPos(0, 4));
                Click(w.Commands["btnBold"]);
                ed.SetCaret(new DocPos(0, 1), false);
                Click(w.Commands["btnFormatPainter"]);
                ed.Select(new DocPos(1, 0), new DocPos(1, 4));
                w.PaintFormatOnSelection();
                T.True(ed.Document.Paragraphs[1].Runs[0].Props.Bold, "format painted (bold)");
                T.Near(36, ed.Document.Paragraphs[1].Props.IndentLeft, 0.1f, "paragraph format painted");
            });

            r.Test("UI11 every ribbon command and menu item can be invoked without errors", () =>
            {
                ed.Document = IoTests.Rich();
                ed.SetCaret(new DocPos(1, 3), false);
                var skip = new HashSet<string> { "btnVisitWebsite", "btnDoPrint", "btnNewWindow" };
                int n = 0;
                foreach (var (name, c) in w.Commands.ToList())
                {
                    if (skip.Contains(name) || !c.Enabled) continue;
                    if (c is RibbonButton or Button) { Click(c); n++; }
                    if (!w.Visible) throw new Exception("window closed by " + name);
                    if (ed.Document.Paragraphs.Count == 0) throw new Exception("document broken by " + name);
                }
                foreach (var (name, it) in w.MenuItems.ToList())
                {
                    if (name is "qatSave") continue;
                    if (!it.Enabled) continue;
                    it.PerformClick(); Pump(); n++;
                }
                // drop-down factories (colour pickers, table grid) build correctly
                foreach (var c in w.Commands.Values.OfType<RibbonButton>().Where(b => b.DropDownFactory != null))
                {
                    var dd = c.DropDownFactory!();
                    T.True(dd.Items.Count > 0, "drop-down content for " + c.Name);
                    dd.Dispose();
                }
                T.True(n > 120, "commands invoked: " + n);
                T.True(SeenDialogs.Count > 5, "dialogs opened and closed: " + string.Join(",", SeenDialogs.Distinct()));
                T.True(w.Visible, "window still open");
            });

            r.Test("UI12 backstage pages: info, print (range parsing, preview), export, new from template", () =>
            {
                w.ShowBackstage("info"); Pump();
                T.True(((Label)Find(w, "lblInfoStats")!).Text.Length > 10, "info statistics");
                ((TextBox)Find(w, "txtInfoTitle")!).Text = "Tiêu đề thử";
                ((Button)Find(w, "btnInfoApply")!).PerformClick();
                T.Eq("Tiêu đề thử", ed.Document.Title, "properties applied");
                w.ShowBackstage("print"); Pump(300);
                T.True(((ComboBox)Find(w, "cboPrinter")!).Items.Count >= 0, "printer list");
                ((RadioButton)Find(w, "rbPrintSome")!).Checked = true;
                ((TextBox)Find(w, "txtPrintPages")!).Text = "1-2";
                var pages = w.SelectedPrintPages();
                T.True(pages != null && pages.SequenceEqual([0, 1]), "page range 1-2");
                ((TextBox)Find(w, "txtPrintPages")!).Text = "9-1";
                T.True(w.SelectedPrintPages() is { Count: 0 }, "invalid range rejected");
                var pv = (PrintPreviewPanel)Find(w, "printPreview")!;
                pv.Show(0);
                ((Button)Find(w, "btnPreviewNext")!).PerformClick();
                T.Eq(1, pv.PageIndex, "preview next page");
                Shot(w, "word-print");
                w.ShowBackstage("export"); Pump();
                T.True(Find(w, "btnExportPdf") != null, "export pdf button");
                w.ShowBackstage("new"); Pump(200);
                ((Button)Find(Find(w, "startNew")!, "tpl_report")!).PerformClick();
                Pump(300);
                var newest = WordMainForm.OpenWindows.Last();
                T.True(newest.Editor.Document.Paragraphs.Any(p => p.Props.Style == StyleIds.Title), "report template opened");
                if (newest != w) { newest.Editor.MarkSaved(); newest.Close(); Pump(); }
                w.HideBackstage();
            });

            r.Test("UI13 save, reopen (single instance forwarding), recent list, autorecover", () =>
            {
                var dir = Path.Combine(T.Out, "ui");
                Directory.CreateDirectory(dir);
                ed.Document = IoTests.Rich();
                var p1 = Path.Combine(dir, "ui-saved.docx");
                T.True(w.SaveTo(p1) && File.Exists(p1) && !ed.Modified, "saved");
                T.True(w.Text.StartsWith("ui-saved.docx"), "title shows file name: " + w.Text);
                T.True(OfficeConfig.RecentFor(AppKind.Word).Contains(p1), "recent list");
                var p2 = Path.Combine(dir, "ui-second.docx");
                File.Copy(p1, p2, true);
                ((IOfficeAppWindow)w).OpenFiles([p1, p2], true);
                Pump(300);
                T.Eq(2, WordMainForm.OpenWindows.Count, "second document in a new window, first reused");
                T.True(((IOfficeAppWindow)w).OpenDocumentPaths.Count == 2, "open document paths");
                var second = WordMainForm.OpenWindows.First(x => x != w);
                second.Editor.InsertText("thay đổi");
                second.ForceAutoRecoverSave();
                var rec = Path.Combine(MicroadsOffice.Core.AppSettings.DataDirectory, "AutoRecover", "word");
                T.True(Directory.GetFiles(rec, "*.docx").Length >= 1, "autorecover file written");
                second.Editor.MarkSaved();
                second.Close(); Pump(200);
                T.Eq(1, WordMainForm.OpenWindows.Count, "second window closed");
                var pdf = Path.Combine(dir, "ui-saved.pdf");
                DialogHandlers.Clear();
                T.True(WordMainForm.OpenWindows[0].ExportPdfTo(pdf) && File.Exists(pdf), "export pdf from window");
            });

            r.Test("UI14 English UI renders and every command has an English label", () =>
            {
                var old = Loc.Lang;
                Loc.Lang = Loc.English;
                try
                {
                    var en = new WordMainForm([]);
                    en.Show(); Pump(300);
                    en.HideBackstage();
                    en.Editor.Document = IoTests.Rich();
                    Pump(300);
                    T.Eq("Home", en.Ribbon.Pages[0].Title);
                    T.True(en.Commands.Values.OfType<RibbonButton>().All(b => !b.ToolTipText.Any(ch => "ảạấầẩẫậắằẳẵặẻẽẹếềểễệỉịỏọốồổỗộớờởỡợủụứừửữựỳỷỹỵđ".Contains(ch))), "tooltips in English");
                    Shot(en, "word-english");
                    en.Editor.MarkSaved();
                    en.Close();
                }
                finally { Loc.Lang = old; }
            });

            r.Test("UI15 main window screenshots with sample content", () =>
            {
                var tw = new WordMainForm([]);
                tw.Show(); Pump(300);
                tw.Size = new Size(1360, 860);
                tw.HideBackstage();
                tw.Editor.Document = Templates.Create("report");
                tw.Editor.SetCaret(new DocPos(3, 5), false);
                Pump(200);
                Shot(tw, "word-main");
                tw.Editor.Document = Templates.Create("congvan");
                Shot(tw, "word-congvan");
                tw.Ribbon.Select(tw.Ribbon.Pages.First(p => p.Name == "pageInsert")); Shot(tw, "word-tab-insert");
                tw.Ribbon.Select(tw.Ribbon.Pages.First(p => p.Name == "pageLayout")); Shot(tw, "word-tab-layout");
                tw.Ribbon.Select(tw.Ribbon.Pages.First(p => p.Name == "pageView")); Shot(tw, "word-tab-view");
                tw.Ribbon.Select(tw.Ribbon.Pages.First(p => p.Name == "pageHelp")); Shot(tw, "word-tab-help");
                tw.Size = new Size(900, 700); Pump(200);
                tw.Ribbon.Select(tw.Ribbon.Pages[0]); Shot(tw, "word-narrow");
                tw.ShowBackstage("home"); Shot(tw, "word-backstage-home");
                tw.ShowBackstage("print"); Shot(tw, "word-backstage-print");
                tw.Editor.MarkSaved();
                tw.Close();
            });
        }
        finally
        {
            foreach (var x in WordMainForm.OpenWindows.ToList()) { x.Editor.MarkSaved(); x.Close(); }
            Pump(200);
            closer.Dispose();
        }
    }
}
