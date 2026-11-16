using System.Drawing.Imaging;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.IO;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.Excel.Ui;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Tests;

/// <summary>
/// In-process UI tests: the real ExcelForm is created and every ribbon button / menu item is invoked with
/// PerformClick(). Modal dialogs, file pickers and confirmations are answered through <see cref="Ux"/> hooks.
/// No real mouse or keyboard input is used.
/// </summary>
public static class UiTests
{
    public static void Run(TestRunner t, string outDir)
    {
        Exception? fatal = null;
        var th = new Thread(() =>
        {
            try { RunSta(t, outDir); }
            catch (Exception ex) { fatal = ex; }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.Start();
        th.Join();
        if (fatal != null) { t.Failed++; t.Log("FAIL  UI pass crashed: " + fatal); }
    }

    private static readonly List<string> Messages = new();
    private static string _out = "";

    private static void Pump(int n = 3) { for (int i = 0; i < n; i++) Application.DoEvents(); }

    private static void RunSta(TestRunner t, string outDir)
    {
        _out = outDir;
        Environment.SetEnvironmentVariable("MICROADS_NO_UPDATE", "1");
        AppSettings.BaseDirectory = Path.Combine(outDir, "appdata");
        Directory.CreateDirectory(AppSettings.BaseDirectory);
        Application.EnableVisualStyles();
        OfficeShell.Init(AppKind.Excel);
        Loc.Lang = Loc.Vietnamese;
        var dataPng = Path.Combine(outDir, "picture.png");
        using (var bmp = new Bitmap(120, 80)) { using (var g = Graphics.FromImage(bmp)) { g.Clear(Color.SeaGreen); g.FillEllipse(Brushes.Gold, 10, 10, 50, 50); } bmp.Save(dataPng, ImageFormat.Png); }
        var csvPath = Path.Combine(outDir, "import.csv");
        File.WriteAllText(csvPath, "Mã,Tên\n1,Bàn\n2,Ghế\n", new System.Text.UTF8Encoding(true));
        int saveCounter = 0;
        Ux.FileHook = (save, filter) =>
        {
            if (!save) return filter.Contains("*.png") ? dataPng : csvPath;
            var ext = filter.Contains("*.xlsx") ? ".xlsx" : filter.Contains("*.pdf") ? ".pdf" : filter.Contains("*.csv") ? ".csv" : filter.Contains("*.html") ? ".html" : ".txt";
            return Path.Combine(outDir, $"ui-save-{++saveCounter}{ext}");
        };
        Ux.ConfirmHook = _ => true;
        Ux.SaveHook = _ => SaveChoice.DontSave;
        Ux.MessageHook = m => Messages.Add(m);
        Func<Form, DialogResult>? custom = null;
        Ux.DialogHook = d =>
        {
            if (custom != null) return custom(d);
            if (d is FindReplaceDialog fr) { fr.FindText = "a"; fr.FindAll(); return DialogResult.OK; }
            if (d is AskDialog ask && ask.Box.Text.Length == 0) ask.Box.Text = "Ten1";
            return DialogResult.OK;
        };

        var form = (ExcelForm)ExcelApp.CreateMainForm([]);
        form.Show();
        Pump(10);
        form.HideBackstage();
        var grid = form.Grid;

        void Type(int r, int c, string text)
        {
            grid.Select(r, c);
            grid.BeginEdit(text, true);
            T.True(grid.CommitEdit(0, 0), "commit " + text);
        }
        Sheet S() => form.Doc.Sheet;

        t.Test("UI: typing values and formulas, undo/redo", () =>
        {
            Type(0, 0, "Sản phẩm"); Type(0, 1, "Quý 1"); Type(0, 2, "Quý 2");
            string[] names = ["Bàn", "Ghế", "Tủ", "Kệ"];
            for (int i = 0; i < 4; i++) { Type(i + 1, 0, names[i]); Type(i + 1, 1, ((i + 1) * 100).ToString()); Type(i + 1, 2, ((i + 2) * 50).ToString()); }
            Type(5, 1, "=SUM(B2:B5)");
            T.Eq(1000.0, S().GetValue(5, 1).Num, "sum");
            T.Eq("SUM(B2:B5)", S().GetCell(5, 1)!.Formula);
            Type(6, 1, "=sum(b2:b3");
            T.Eq(300.0, S().GetValue(6, 1).Num, "auto-closed paren & lower-case");
            form.CmdUndo();
            T.True(S().GetCell(6, 1) == null, "undo removed");
            form.CmdRedo();
            T.Eq(300.0, S().GetValue(6, 1).Num, "redo");
            Type(6, 1, "");
            Type(7, 0, "15/3/2024");
            T.Eq("dd/mm/yyyy", S().StyleAt(7, 0).NumberFormat, "date auto-format");
            Type(7, 1, "12%");
            T.Eq(0.12, S().GetValue(7, 1).Num, "percent");
            // point mode: "=" then arrow down inserts reference
            grid.Select(8, 1);
            grid.BeginEdit("=", true);
            grid.HandleEditKey(new KeyEventArgs(Keys.Up), form.FormulaBar);
            T.Eq("=B8", grid.EditText, "point mode reference");
            grid.CommitEdit(0, 0);
            T.Eq(0.12, S().GetValue(8, 1).Num);
            // formula bar shows the formula of the active cell
            grid.Select(5, 1);
            form.UpdateCellInfo();
            T.Eq("=SUM(B2:B5)", form.FormulaBar.Text);
            T.Eq("B6", form.NameBox.Text);
            // status bar statistics
            grid.SelectRange(new CellRange(1, 1, 4, 1));
            form.UpdateStats();
            T.True(form.StatusStats.Contains("1.000") || form.StatusStats.Contains("1,000"), "status sum: " + form.StatusStats);
        });

        t.Test("UI: formatting buttons have the expected effect", () =>
        {
            grid.SelectRange(new CellRange(0, 0, 0, 2));
            Click(form, "btnBold");
            T.True(S().StyleAt(0, 1).Bold, "bold");
            Click(form, "btnBold");
            T.True(!S().StyleAt(0, 1).Bold, "bold toggled off");
            Click(form, "btnBold");
            Click(form, "btnAlignCenter");
            T.Eq(HAlign.Center, S().StyleAt(0, 2).HAlign);
            Click(form, "btnFillColor");
            T.True(S().StyleAt(0, 0).Fill != 0, "fill");
            grid.SelectRange(new CellRange(1, 1, 4, 2));
            Click(form, "btnCurrency");
            T.True(S().StyleAt(2, 2).NumberFormat.Contains("₫"), "currency");
            Click(form, "btnDecInc");
            T.True(S().StyleAt(2, 2).NumberFormat.Contains(".0"), "decimals: " + S().StyleAt(2, 2).NumberFormat);
            Click(form, "btnBorders");
            T.True(!S().StyleAt(4, 1).Bottom.IsNone, "bottom border");
            form.ApplyBorders("all");
            T.True(!S().StyleAt(1, 1).Top.IsNone && !S().StyleAt(1, 1).Left.IsNone, "all borders");
            grid.SelectRange(new CellRange(10, 0, 10, 3));
            Click(form, "btnMerge");
            T.True(S().Merges.Contains(new CellRange(10, 0, 10, 3)), "merged");
            Click(form, "btnMerge");
            T.True(!S().Merges.Contains(new CellRange(10, 0, 10, 3)), "unmerged by toggle");
            grid.Select(0, 0);
            Click(form, "btnWrap");
            T.True(S().StyleAt(0, 0).Wrap, "wrap");
            form.SetFontName("Times New Roman");
            T.Eq("Times New Roman", S().StyleAt(0, 0).FontName);
            form.SetFontSize("14");
            T.Eq(14f, S().StyleAt(0, 0).FontSize);
        });

        t.Test("UI: rows, columns, sheets, sort, filter, autosum", () =>
        {
            grid.Select(2, 0);
            form.CmdInsertRows();
            T.Eq("Ghế", S().GetValue(3, 0).Str, "row inserted");
            T.Eq("SUM(B2:B6)", S().GetCell(6, 1)!.Formula, "formula adjusted");
            form.CmdDeleteRows();
            T.Eq("SUM(B2:B5)", S().GetCell(5, 1)!.Formula);
            grid.SelectRange(new CellRange(0, 0, 4, 2), new CellPos(1, 1));
            form.CmdSort(false);
            T.Eq(400.0, S().GetValue(1, 1).Num, "sorted desc");
            T.Eq(1000.0, S().GetValue(5, 1).Num, "total row untouched");
            T.Eq("Sản phẩm", S().GetValue(0, 0).Str, "header kept");
            grid.Select(1, 0);
            form.CmdToggleFilter();
            T.True(S().Filter != null, "filter on");
            form.ShowFilterPopup(0);
            form.LastFilterPopup!.CheckOnly("Bàn");
            form.LastFilterPopup.Apply();
            var visible = Enumerable.Range(1, 4).Where(r => !S().FilteredRows.Contains(r)).ToList();
            T.Eq(1, visible.Count, "one data row visible");
            T.Eq("Bàn", S().GetValue(visible[0], 0).Str);
            form.CmdClearFilter();
            T.Eq(0, S().FilteredRows.Count);
            form.CmdToggleFilter();
            T.True(S().Filter == null, "filter off");
            grid.Select(5, 2);
            form.CmdAutoSum("SUM");
            T.Eq("SUM(C2:C5)", S().GetCell(5, 2)!.Formula);
            int sheets = form.Doc.Book.Sheets.Count;
            form.CmdInsertSheet();
            T.Eq(sheets + 1, form.Doc.Book.Sheets.Count);
            form.CmdDeleteSheet();
            T.Eq(sheets, form.Doc.Book.Sheets.Count);
            form.SwitchSheet(0);
        });

        t.Test("UI: chart, conditional formatting, validation, freeze, names", () =>
        {
            grid.SelectRange(new CellRange(0, 0, 4, 2));
            form.CmdInsertChart(ChartType.Column);
            T.Eq(1, S().Charts.Count);
            T.True(grid.SelectedObject is Chart, "chart selected");
            Click(form, "btnChartLegend");
            T.True(!S().Charts[0].ShowLegend, "legend toggled");
            foreach (var it in form.MenuItems.Where(m => m.Name == "miChartTypeLine").Take(1)) it.PerformClick();
            T.Eq(ChartType.Line, S().Charts[0].Type);
            grid.SelectObject(null);
            grid.SelectRange(new CellRange(1, 1, 4, 1));
            custom = d => { if (d is CfRuleDialog c) c.SetValues("250"); return DialogResult.OK; };
            foreach (var it in form.MenuItems.Where(m => m.Name == "miCfGreater").Take(1)) it.PerformClick();
            custom = null;
            T.True(S().CondFormats.Any(c => c.Type == CfType.GreaterThan && c.Value1 == "250"), "CF added");
            grid.Select(1, 4);
            custom = d => { if (d is ValidationDialog v) v.Set(ValidationType.List, "Có,Không"); return DialogResult.OK; };
            form.CmdDataValidation();
            custom = null;
            T.True(S().Validations.Any(v => v.Formula1 == "Có,Không"), "validation");
            grid.BeginEdit("Sai", true);
            T.True(!grid.CommitEdit(0, 0), "invalid entry rejected");
            grid.CancelEdit();
            grid.Select(1, 1);
            form.CmdFreezePanes();
            T.Eq(1, S().FreezeRows); T.Eq(1, S().FreezeCols);
            form.SetFreeze(0, 0);
            grid.SelectRange(new CellRange(1, 1, 4, 1));
            custom = d => { if (d is AskDialog a) a.Box.Text = "DoanhThu"; return DialogResult.OK; };
            form.CmdDefineName();
            custom = null;
            T.True(form.Doc.Book.Names.ContainsKey("DoanhThu"), "name defined");
            Type(12, 0, "=SUM(DoanhThu)");
            T.Eq(1000.0, S().GetValue(12, 0).Num, "name used");
            Type(13, 0, "=DOCSO(A13)");
            T.Eq("Một nghìn đồng", S().GetValue(13, 0).Str);
        });

        t.Test("UI: save, reopen, export PDF/CSV/HTML, templates, backstage", () =>
        {
            var path = Path.Combine(outDir, "ui-workbook.xlsx");
            T.True(form.SaveTo(form.Doc, path), "save");
            T.True(!form.Doc.Modified, "saved flag");
            var re = XlsxReader.Read(path);
            T.Eq("Bàn", re.Sheets[0].GetValue(4, 0).Str == "Bàn" ? "Bàn" : re.Sheets[0].AllCells.Any(c => c.Value.Value.Str == "Bàn") ? "Bàn" : "?");
            T.Eq(1, re.Sheets[0].Charts.Count);
            var pdf = Path.Combine(outDir, "ui-export.pdf");
            form.ExportPdfTo(pdf, forceImage: true);
            var head = File.ReadAllBytes(pdf).Take(5).ToArray();
            T.Eq("%PDF-", System.Text.Encoding.ASCII.GetString(head), "pdf header");
            T.True(new FileInfo(pdf).Length > 2000, "pdf size");
            foreach (var id in new[] { "budget", "invoice", "timesheet", "list", "blank" })
            {
                form.CmdNewFromTemplate(id);
                T.True(form.Doc.Book.Sheets[0].CellCount > 0 || id == "blank", "template " + id);
            }
            T.True(form.Doc.Book.Sheets[0].CellCount == 0, "blank last");
            int docs = form.Docs.Count;
            form.ShowBackstage("home"); Pump();
            form.ShowBackstage("info"); Pump();
            form.ShowBackstage("print"); Pump(5);
            form.ShowBackstage("export"); Pump();
            form.HideBackstage();
            T.True(form.OpenFile(path), "reopen (already open → activated)");
            T.Eq(path, form.Doc.FilePath);
            T.Eq(docs, form.Docs.Count);
            var copy = Path.Combine(outDir, "ui-workbook-copy.xlsx");
            File.Copy(path, copy, true);
            T.True(form.OpenFile(copy), "open copy");
            T.Eq(docs + 1, form.Docs.Count);
            T.Eq(1000.0, form.Doc.Book.Sheets[0].AllCells.Where(c => c.Value.Formula == "SUM(B2:B5)").Select(c => c.Value.Value.Num).FirstOrDefault(), "formula value after reopen");
            var invoice = form.Docs.First(d => d.Book.Sheets[0].Name is "Hóa đơn" or "Invoice");
            T.True(invoice.Book.Sheets[0].AllCells.Any(c => c.Value.Value.IsText && c.Value.Value.Str!.EndsWith("đồng") && c.Value.Value.Str.Contains("triệu")), "invoice DOCSO computed");
        });

        t.Test("UI: grid keyboard and mouse handlers (simulated events)", () =>
        {
            form.CmdNewFromTemplate("blank");
            var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
            void Key(Keys k) { typeof(GridView).GetMethod("OnKeyDown", flags)!.Invoke(grid, [new KeyEventArgs(k)]); Pump(1); }
            void Char(char c) { typeof(GridView).GetMethod("OnKeyPress", flags)!.Invoke(grid, [new KeyPressEventArgs(c)]); Pump(1); }
            void Mouse(string m, Point p, int clicks = 1) { typeof(GridView).GetMethod(m, flags)!.Invoke(grid, [new MouseEventArgs(MouseButtons.Left, clicks, p.X, p.Y, 0)]); if (m != "OnMouseDown" && m != "OnMouseMove") Pump(1); }
            Point Center(int r, int c) { var rc = grid.CellRect(r, c)!.Value; return new Point(rc.X + rc.Width / 2, rc.Y + rc.Height / 2); }
            grid.Select(0, 0);
            Char('5'); grid.CommitEdit(1, 0);
            T.Eq(5.0, S().GetValue(0, 0).Num, "typed via key press");
            T.Eq(new CellPos(1, 0), grid.Active, "moved down after enter");
            Char('7'); grid.CommitEdit(1, 0);
            Char('9'); grid.CommitEdit(1, 0);
            Key(Keys.Control | Keys.Up);
            T.Eq(0, grid.Active.Row, "ctrl+up to data edge");
            Key(Keys.Control | Keys.Down);
            T.Eq(2, grid.Active.Row, "ctrl+down to data edge");
            Key(Keys.Shift | Keys.Up);
            T.Eq(new CellRange(1, 0, 2, 0), grid.Selection, "shift extends");
            Key(Keys.Right);
            T.Eq(new CellPos(2, 1), grid.Active, "right");
            Key(Keys.Tab);
            T.Eq(new CellPos(2, 2), grid.Active, "tab");
            Key(Keys.Home);
            T.Eq(0, grid.Active.Col, "home");
            Key(Keys.Delete);
            T.True(S().GetValue(2, 0).IsBlank, "delete clears");
            form.CmdUndo();
            T.Eq(9.0, S().GetValue(2, 0).Num, "undo delete");
            // mouse: click, drag-select, fill handle
            grid.ScrollTo(0, 0);
            Mouse("OnMouseDown", Center(0, 0)); Mouse("OnMouseMove", Center(2, 0)); Mouse("OnMouseUp", Center(2, 0));
            T.Eq(new CellRange(0, 0, 2, 0), grid.Selection, "drag selection");
            var fh = grid.FillHandleRect()!.Value;
            Mouse("OnMouseDown", new Point(fh.X + fh.Width / 2, fh.Y + fh.Height / 2));
            Mouse("OnMouseMove", Center(5, 0));
            Mouse("OnMouseUp", Center(5, 0));
            T.Eq(new CellRange(0, 0, 5, 0), grid.Selection, "fill extended selection");
            T.Eq(5.0 + 2 * 3, S().GetValue(3, 0).Num, "series continued (5,7,9 → 11)");
            // copy / paste through the grid
            grid.SelectRange(new CellRange(0, 0, 1, 0));
            grid.CopySelection(false);
            grid.Select(0, 3);
            grid.Paste();
            T.Eq(7.0, S().GetValue(1, 3).Num, "paste");
            // column resize by dragging the header border
            var col = grid.CellRect(0, 1)!.Value;
            var border = new Point(col.Right - 1, grid.ColHeaderH / 2);
            Mouse("OnMouseDown", border); Mouse("OnMouseMove", new Point(border.X + 40, border.Y)); Mouse("OnMouseUp", new Point(border.X + 40, border.Y));
            T.True(S().RawColWidth(1) > Sheet.DefaultColWidth + 20, "column resized: " + S().RawColWidth(1));
            // click column header selects the column
            Mouse("OnMouseDown", new Point(col.X + 5, grid.ColHeaderH / 2)); Mouse("OnMouseUp", new Point(col.X + 5, grid.ColHeaderH / 2));
            T.Eq(new CellRange(0, 1, Addr.MaxRows - 1, 1), grid.Selection, "column selected");
            // double click to edit, F2 toggle and escape
            grid.Select(0, 0);
            Mouse("OnMouseDoubleClick", Center(0, 0), 2);
            T.True(grid.IsEditing, "double click edits");
            grid.CancelEdit();
            Key(Keys.F2);
            T.True(grid.IsEditing, "F2 edits");
            grid.CancelEdit();
            // zoom with the wheel (Ctrl)
            grid.SetZoom(1.5);
            T.Eq(1.5, S().Zoom);
            grid.SetZoom(1);
        });

        t.Test("UI: every ribbon button and menu item can be clicked without errors", () =>
        {
            var skip = new HashSet<string> { "btnCheckUpdate", "btnSettings", "btnAbout", "btnVisitWebsite", "btnHelp", "btnFormulaFx" };
            // fresh workbook with data so that every command has something to act on
            form.CmdNewFromTemplate("budget");
            var errPrefix = Loc.T("Không thực hiện được lệnh", "The command could not be completed");
            int clicked = 0, items = 0;
            foreach (var (name, ctl) in form.CommandControls.OrderBy(k => k.Key))
            {
                if (skip.Contains(name) || ctl is not RibbonButton b) continue;
                if (!b.Enabled) continue;
                PrepareSelection(form, name);
                Messages.Clear();
                if (b.DropDown != null && !b.Split) { clicked++; continue; } // menu-only button: its items are exercised below
                b.PerformClick();
                Pump();
                CloseDropDowns();
                clicked++;
                var err = Messages.FirstOrDefault(m => m.StartsWith(errPrefix));
                if (err != null) throw new AssertException(name + ": " + err);
                if (form.Doc.Sheet.Protected) form.CmdProtectSheet();
            }
            foreach (var it in form.MenuItems.ToList())
            {
                if (it.IsDisposed) continue;
                PrepareSelection(form, it.Name ?? "");
                Messages.Clear();
                it.PerformClick();
                Pump();
                CloseDropDowns();
                items++;
                var err = Messages.FirstOrDefault(m => m.StartsWith(errPrefix));
                if (err != null) throw new AssertException(it.Name + ": " + err);
                if (form.Doc.Sheet.Protected) form.CmdProtectSheet();
                if (form.BackstageVisible) form.HideBackstage();
            }
            // context menus and sheet tab menu
            foreach (var menu in new[] { form.SheetMenu() })
                foreach (ToolStripItem it in menu.Items)
                    if (it is ToolStripMenuItem mi && mi.Name != "smDelete" && mi.Name != "smHide") { mi.PerformClick(); Pump(); CloseDropDowns(); items++; }
            t.Log($"      ribbon buttons: {clicked}, menu items: {items}");
            T.True(clicked > 100, "number of ribbon buttons " + clicked);
            T.True(items > 150, "number of menu items " + items);
            // shell buttons exist (they open shared windows, verified by the suite tests)
            foreach (var n in skip.Where(x => x != "btnFormulaFx")) T.True(form.CommandControls.ContainsKey(n), "has " + n);
        });

        t.Test("UI: screenshot of the main window with data and a chart", () =>
        {
            form.CmdNewFromTemplate("budget");
            form.WindowState = FormWindowState.Normal;
            form.Size = new Size(1400, 860);
            form.HideBackstage();
            form.Ribbon.Select(form.Ribbon.Pages[0]);
            grid.Select(4, 1);
            Pump(10);
            using var bmp = new Bitmap(form.Width, form.Height);
            form.DrawToBitmap(bmp, new Rectangle(0, 0, form.Width, form.Height));
            bmp.Save(Path.Combine(outDir, "excel-main.png"), ImageFormat.Png);
            form.CmdNewFromTemplate("invoice");
            Pump(5);
            using var bmp2 = new Bitmap(form.Width, form.Height);
            form.DrawToBitmap(bmp2, new Rectangle(0, 0, form.Width, form.Height));
            bmp2.Save(Path.Combine(outDir, "excel-invoice.png"), ImageFormat.Png);
            form.ShowBackstage("home");
            Pump(5);
            var bs = form.BackstageView;
            using var bmp3 = new Bitmap(bs.Width, bs.Height);
            bs.DrawToBitmap(bmp3, new Rectangle(0, 0, bs.Width, bs.Height));
            bmp3.Save(Path.Combine(outDir, "excel-start.png"), ImageFormat.Png);
            form.HideBackstage();
            Loc.Lang = Loc.English;
        });

        t.Test("UI: AutoRecover writes a recovery copy and closing cleans it up", () =>
        {
            form.CmdNewFromTemplate("list");
            form.Doc.Modified = true;
            form.AutoRecoverTick(force: true);
            var dir = Path.Combine(AppSettings.DataDirectory, "AutoRecover", "excel");
            T.True(Directory.GetFiles(dir, "*.xlsx").Length >= 1, "recover file written");
            Ux.SaveHook = _ => SaveChoice.DontSave;
            form.Close();
            Pump(5);
            T.True(Directory.GetFiles(dir, "*.xlsx").Length == 0, "recover files removed on normal close");
        });
        Ux.DialogHook = null; Ux.FileHook = null; Ux.ConfirmHook = null; Ux.SaveHook = null; Ux.MessageHook = null;
    }

    private static void Click(ExcelForm f, string name)
    {
        var b = (RibbonButton)f.CommandControls[name];
        if (b.Visible) b.PerformClick();
        // On small screens (e.g. 1024×768 CI runners) the ribbon collapses groups into a drop-down: the button is then
        // hidden until the group is opened. Click it as the user would inside the opened group.
        else typeof(Control).GetMethod("OnClick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(b, [EventArgs.Empty]);
        Pump();
    }

    private static void PrepareSelection(ExcelForm f, string name)
    {
        var g = f.Grid;
        if (g.IsEditing) g.CommitEdit(0, 0);
        f.HideBackstage();
        if (name.StartsWith("btnChart") && name is not ("btnChartColumn" or "btnChartBar" or "btnChartLine" or "btnChartPie" or "btnChartArea" or "btnChartScatter" or "btnChartStacked" or "btnChartDoughnut") || name.StartsWith("miChartType") || name.StartsWith("cmChart"))
        {
            if (f.Doc.Sheet.Charts.Count == 0) { g.SelectRange(new CellRange(9, 0, 16, 1)); f.CmdInsertChart(ChartType.Column); }
            g.SelectObject(f.Doc.Sheet.Charts[0]);
            return;
        }
        if (name.StartsWith("btnPicture") || name.StartsWith("cmPic"))
        {
            if (f.Doc.Sheet.Pictures.Count == 0) { using var img = new Bitmap(40, 30); g.InsertPicture(img); }
            g.SelectObject(f.Doc.Sheet.Pictures[0]);
            return;
        }
        g.SelectObject(null);
        g.SelectRange(new CellRange(10, 0, 16, 1));
    }

    private static void CloseDropDowns()
    {
        foreach (var f in Application.OpenForms.Cast<Form>().ToList())
            if (f is FindReplaceDialog) f.Hide();
        Pump();
    }
}
