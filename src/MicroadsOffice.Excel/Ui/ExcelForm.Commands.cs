using System.Drawing.Printing;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.IO;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

public sealed partial class ExcelForm
{
    private Sheet S => _doc.Sheet;

    private bool Guard()
    {
        if (!S.Protected) return true;
        Ux.Warn(this, Loc.T("Trang tính đang được bảo vệ. Vào Xem lại ▸ Bỏ bảo vệ trang tính để thay đổi.", "The sheet is protected. Go to Review ▸ Unprotect Sheet to make changes."));
        return false;
    }

    private void Changed() { _doc.Recalc(); Grid.RaiseContentChanged(); Grid.RefreshAll(); }

    // ------------------------------------------------------------------ formatting

    /// <summary>Applies a style change to every selected cell (whole rows/columns also update their default style).</summary>
    public void Restyle(Func<CellStyle, CellStyle> f, string label)
    {
        if (!Guard()) return;
        _doc.Checkpoint(label);
        var wb = _doc.Book;
        var map = new Dictionary<int, int>();
        int Map(int id) { if (!map.TryGetValue(id, out var n)) map[id] = n = wb.Styles.Intern(f(wb.Styles[id])); return n; }
        foreach (var r in Grid.AllSelections)
        {
            bool wholeCols = r.R1 == 0 && r.R2 == Addr.MaxRows - 1, wholeRows = r.C1 == 0 && r.C2 == Addr.MaxCols - 1;
            if (wholeCols && !wholeRows) for (int c = r.C1; c <= r.C2; c++) S.ColStyles[c] = Map(S.ColStyles.GetValueOrDefault(c));
            if (wholeRows && !wholeCols) for (int rr = r.R1; rr <= r.R2; rr++) S.RowStyles[rr] = Map(S.RowStyles.GetValueOrDefault(rr));
            if (r.Count <= 250_000 && !wholeCols && !wholeRows)
                foreach (var (row, col) in r.Cells()) S.SetStyle(row, col, Map(S.StyleIdAt(row, col)));
            else
                foreach (var k in S.AllCells.Select(x => x.Key).ToList())
                {
                    var (row, col) = Sheet.FromKey(k);
                    if (r.Contains(row, col)) S.SetStyle(row, col, Map(S.StyleIdAt(row, col)));
                }
        }
        wb.MarkDirty();
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    private void ToggleStyle(Func<CellStyle, bool> get, Func<CellStyle, bool, CellStyle> set)
    {
        bool v = !get(S.StyleAt(Grid.Active.Row, Grid.Active.Col));
        Restyle(s => set(s, v), Loc.T("Định dạng", "Format"));
    }

    private void ToggleAlign(HAlign a)
    {
        var cur = S.StyleAt(Grid.Active.Row, Grid.Active.Col).HAlign;
        Restyle(s => s with { HAlign = cur == a ? HAlign.General : a }, Loc.T("Căn lề", "Align"));
    }

    public void SetFontName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        Restyle(s => s with { FontName = name.Trim() }, Loc.T("Phông chữ", "Font"));
        Grid.Focus();
    }

    public void SetFontSize(string text)
    {
        if (float.TryParse(text.Replace(',', '.'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var sz) && sz >= 1 && sz <= 409)
            Restyle(s => s with { FontSize = sz }, Loc.T("Cỡ chữ", "Font size"));
        Grid.Focus();
    }

    public void SetNumberFormat(string code) => Restyle(s => s with { NumberFormat = code }, Loc.T("Định dạng số", "Number format"));

    private void AdjustDecimals(int delta)
    {
        var v = S.GetValue(Grid.Active.Row, Grid.Active.Col);
        double sample = v.IsNumber ? v.Num : 0;
        Restyle(s => s with { NumberFormat = NumFmt.AdjustDecimals(s.NumberFormat, sample, delta) }, Loc.T("Số thập phân", "Decimals"));
    }

    private void ApplyFill(int argb) => Restyle(s => s with { Fill = argb }, Loc.T("Màu nền", "Fill colour"));

    public void ApplyBorders(string preset, LineStyle style = LineStyle.Thin)
    {
        var line = new BorderLine(preset == "thick" ? LineStyle.Thick : preset == "double" ? LineStyle.Double : style, Color.Black.ToArgb());
        if (!Guard()) return;
        _doc.Checkpoint(Loc.T("Viền", "Borders"));
        var wb = _doc.Book;
        foreach (var r0 in Grid.AllSelections)
        {
            var r = r0.Count > 200_000 ? new CellRange(r0.R1, r0.C1, Math.Min(r0.R2, S.UsedRange().R2), Math.Min(r0.C2, S.UsedRange().C2)) : r0;
            foreach (var (row, col) in r.Cells())
            {
                var st = wb.Styles[S.StyleIdAt(row, col)];
                st = preset switch
                {
                    "none" => st with { Left = BorderLine.None, Right = BorderLine.None, Top = BorderLine.None, Bottom = BorderLine.None },
                    "all" => st with { Left = line, Right = line, Top = line, Bottom = line },
                    "bottom" or "double" => row == r.R2 ? st with { Bottom = line } : st,
                    "top" => row == r.R1 ? st with { Top = line } : st,
                    "left" => col == r.C1 ? st with { Left = line } : st,
                    "right" => col == r.C2 ? st with { Right = line } : st,
                    "outline" or "thick" => st with
                    {
                        Left = col == r.C1 ? line : st.Left, Right = col == r.C2 ? line : st.Right,
                        Top = row == r.R1 ? line : st.Top, Bottom = row == r.R2 ? line : st.Bottom
                    },
                    _ => st
                };
                S.SetStyle(row, col, wb.Styles.Intern(st));
            }
        }
        wb.MarkDirty();
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    public void CmdFormatCells(int tab)
    {
        var a = Grid.Active;
        var v = S.GetValue(a.Row, a.Col);
        bool merged = S.MergeAt(a.Row, a.Col) != null;
        using var d = new FormatCellsDialog(S.StyleAt(a.Row, a.Col), v.IsNumber ? v.Num : 1234.5678, merged, tab);
        if (Ux.Show(d, this) != DialogResult.OK) return;
        Restyle(d.Apply, Loc.T("Định dạng ô", "Format cells"));
        if (d.BorderTouched) ApplyBorders(d.BorderPreset, d.BorderLineStyle);
        if (d.MergeTouched) { if (d.MergeBox.Checked) CmdMerge(false, false); else CmdUnmerge(); }
    }

    public void CmdMerge(bool center, bool across)
    {
        if (!Guard()) return;
        var sel = Grid.Selection;
        if (sel.IsSingle) return;
        if (S.MergeAt(sel.R1, sel.C1) == sel && !across) { CmdUnmerge(); return; }
        var rows = across ? Enumerable.Range(sel.R1, sel.Rows).Select(r => new CellRange(r, sel.C1, r, sel.C2)).ToList() : [sel];
        bool lose = rows.Any(m => m.Cells().Skip(1).Any(p => !S.GetValue(p.R, p.C).IsBlank));
        if (lose && !Ux.Confirm(this, Loc.T("Gộp ô chỉ giữ lại giá trị ở ô trên cùng bên trái, các giá trị khác sẽ bị xóa. Tiếp tục?", "Merging cells only keeps the upper-left value and discards other values. Continue?"))) return;
        _doc.Checkpoint(Loc.T("Gộp ô", "Merge cells"));
        foreach (var m in rows)
        {
            if (m.IsSingle) continue;
            S.Merges.RemoveAll(x => x.Intersects(m));
            foreach (var (r, c) in m.Cells().Skip(1)) S.ClearCell(r, c, true, false);
            S.Merges.Add(m);
            if (center) S.SetStyle(m.R1, m.C1, _doc.Book.Restyle(S.StyleIdAt(m.R1, m.C1), s => s with { HAlign = HAlign.Center, VAlign = VAlign.Center }));
        }
        Changed();
    }

    public void CmdUnmerge()
    {
        var sel = Grid.Selection;
        if (!S.Merges.Any(m => m.Intersects(sel))) return;
        _doc.Checkpoint(Loc.T("Bỏ gộp ô", "Unmerge"));
        S.Merges.RemoveAll(m => m.Intersects(sel));
        Changed();
    }

    private void CmdFormatPainter()
    {
        if (_painter != null) { _painter = null; Grid.Cursor = Cursors.Cross; return; }
        _painter = new FormatPainterState { Clip = SheetOps.Copy(S, Grid.Selection, false) };
        _statusMode.Text = Loc.T("Chọn ô cần áp dụng định dạng", "Select cells to apply the formatting");
    }

    private void ApplyPainterIfActive()
    {
        if (_painter == null) return;
        var p = _painter;
        if (!p.Sticky) _painter = null;
        _doc.Checkpoint(Loc.T("Sao định dạng", "Format painter"));
        SheetOps.Paste(S, p.Clip, Grid.Selection, PasteMode.Formats);
        Changed();
        UpdateCellInfo();
    }

    private void CmdPasteSpecial()
    {
        using var d = new PasteSpecialDialog();
        if (Ux.Show(d, this) != DialogResult.OK) return;
        Grid.Paste(d.Mode, d.Transpose, d.SkipBlanks);
    }

    // ------------------------------------------------------------------ conditional formatting / tables

    private void CmdCfRule(CfType t)
    {
        var a = Grid.Active;
        var nums = Grid.Selection.Cells().Take(5000).Select(p => S.GetValue(p.R, p.C)).Where(v => v.IsNumber).Select(v => v.Num).ToList();
        string suggestion = t == CfType.Expression ? "=" + Addr.ToA1(Grid.Selection.R1, Grid.Selection.C1, false, true) + ">0" : nums.Count > 0 ? NumFmt.General(Math.Round(nums.Average(), 2)) : "";
        using var d = new CfRuleDialog(t, suggestion);
        if (Ux.Show(d, this) != DialogResult.OK) return;
        AddCf(d.Build(Grid.Selection));
    }

    public void AddCf(ConditionalFormat cf)
    {
        if (!Guard()) return;
        var sel = Grid.Selection;
        if (sel.Count > 1_000_000) sel = new CellRange(sel.R1, sel.C1, Math.Min(sel.R2, S.UsedRange().R2), Math.Min(sel.C2, S.UsedRange().C2));
        cf.Range = sel;
        _doc.Checkpoint(Loc.T("Định dạng có điều kiện", "Conditional formatting"));
        S.CondFormats.Add(cf);
        _doc.Book.MarkDirty();
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    private void ClearCf(bool sheet)
    {
        _doc.Checkpoint(Loc.T("Xóa quy tắc", "Clear rules"));
        if (sheet) S.CondFormats.Clear(); else S.CondFormats.RemoveAll(c => c.Range.Intersects(Grid.Selection));
        _doc.Book.MarkDirty();
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    private static readonly (Color Head, Color Band)[] TableColors =
    [
        (Color.FromArgb(33, 115, 70), Color.FromArgb(226, 239, 218)), (Color.FromArgb(46, 117, 182), Color.FromArgb(221, 235, 247)),
        (Color.FromArgb(237, 125, 49), Color.FromArgb(252, 228, 214)), (Color.FromArgb(118, 118, 118), Color.FromArgb(237, 237, 237))
    ];

    public void CmdFormatAsTable(int scheme)
    {
        if (!Guard()) return;
        var r = Grid.Selection.IsSingle ? SheetOps.CurrentRegion(S, Grid.Active.Row, Grid.Active.Col) : Grid.Selection;
        if (r.IsSingle && S.GetValue(r.R1, r.C1).IsBlank) { Ux.Info(this, Loc.T("Hãy chọn vùng dữ liệu (có hàng tiêu đề) trước.", "Select a data range (with a header row) first.")); return; }
        var (head, band) = TableColors[Math.Clamp(scheme, 0, TableColors.Length - 1)];
        _doc.Checkpoint(Loc.T("Định dạng như bảng", "Format as table"));
        var wb = _doc.Book;
        var edge = BorderLine.Thin(Theme.Blend(head, Color.White, 0.5f));
        for (int row = r.R1; row <= r.R2; row++)
            for (int col = r.C1; col <= r.C2; col++)
            {
                var st = wb.Styles[S.StyleIdAt(row, col)];
                st = row == r.R1
                    ? st with { Bold = true, Fill = head.ToArgb(), FontColor = Color.White.ToArgb(), Bottom = edge }
                    : st with { Fill = (row - r.R1) % 2 == 0 ? band.ToArgb() : 0, Bottom = edge };
                S.SetStyle(row, col, wb.Styles.Intern(st));
            }
        S.Filter = new AutoFilter { Range = r };
        SheetOps.ApplyFilter(S);
        Changed();
    }

    // ------------------------------------------------------------------ rows / columns / sheets

    private IEnumerable<int> SelRows() { var s = Grid.Selection; return Enumerable.Range(s.R1, Math.Min(s.Rows, 10000)); }
    private IEnumerable<int> SelCols() { var s = Grid.Selection; return Enumerable.Range(s.C1, Math.Min(s.Cols, 1000)); }

    public void CmdInsertRows()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        _doc.Checkpoint(Loc.T("Chèn hàng", "Insert rows"));
        SheetOps.InsertRows(S, s.R1, Math.Min(s.Rows, 10000));
        Changed();
    }

    public void CmdInsertCols()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        _doc.Checkpoint(Loc.T("Chèn cột", "Insert columns"));
        SheetOps.InsertCols(S, s.C1, Math.Min(s.Cols, 1000));
        Changed();
    }

    public void CmdDeleteRows()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        _doc.Checkpoint(Loc.T("Xóa hàng", "Delete rows"));
        SheetOps.DeleteRows(S, s.R1, s.Rows);
        Grid.Select(s.R1, Grid.Active.Col);
        Changed();
    }

    public void CmdDeleteCols()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        _doc.Checkpoint(Loc.T("Xóa cột", "Delete columns"));
        SheetOps.DeleteCols(S, s.C1, s.Cols);
        Grid.Select(Grid.Active.Row, s.C1);
        Changed();
    }

    public void CmdInsertSheet()
    {
        _doc.Checkpoint(Loc.T("Chèn trang tính", "Insert sheet"));
        var idx = _doc.Book.ActiveSheet + 1;
        _doc.Book.AddSheet(null, idx);
        SwitchSheet(idx);
        OnContentChanged();
    }

    public void CmdDeleteSheet()
    {
        if (_doc.Book.Sheets.Count(s => !s.Hidden) <= 1) { Ux.Warn(this, Loc.T("Sổ làm việc phải có ít nhất một trang tính hiển thị.", "A workbook must contain at least one visible worksheet.")); return; }
        if (S.CellCount > 0 && !Ux.Confirm(this, Loc.F("Xóa vĩnh viễn trang tính \"{0}\"?", "Permanently delete the sheet \"{0}\"?", S.Name), Loc.T("Xóa", "Delete"), Loc.T("Hủy", "Cancel"), true)) return;
        _doc.Checkpoint(Loc.T("Xóa trang tính", "Delete sheet"));
        SheetOps.DeleteSheet(_doc.Book, S);
        Changed();
        SwitchSheet(_doc.Book.ActiveSheet);
        OnContentChanged();
    }

    public void CmdDuplicateSheet()
    {
        _doc.Checkpoint(Loc.T("Nhân bản trang tính", "Duplicate sheet"));
        var copy = SheetOps.DuplicateSheet(_doc.Book, S);
        _doc.Recalc();
        SwitchSheet(_doc.Book.Sheets.IndexOf(copy));
        OnContentChanged();
    }

    public void CmdRenameSheet()
    {
        var name = AskDialog.Ask(this, Loc.T("Đổi tên trang tính", "Rename Sheet"), Loc.T("Tên mới (tối đa 31 ký tự, không chứa : \\ / ? * [ ]):", "New name (max 31 characters, without : \\ / ? * [ ]):"), S.Name);
        if (name == null || name == S.Name) return;
        name = name.Trim();
        if (!Workbook.IsValidSheetName(name)) { Ux.Warn(this, Loc.T("Tên trang tính không hợp lệ.", "That sheet name is not valid.")); return; }
        if (_doc.Book.Sheet(name) != null) { Ux.Warn(this, Loc.T("Đã có trang tính trùng tên.", "That name is already taken.")); return; }
        _doc.Checkpoint(Loc.T("Đổi tên trang tính", "Rename sheet"));
        SheetOps.RenameSheet(_doc.Book, S, name);
        Changed();
        Tabs.Invalidate();
    }

    private void CmdTabColor()
    {
        new ColorPickerDropDown(c => { _doc.Checkpoint(Loc.T("Màu thẻ", "Tab colour")); S.TabColor = c?.ToArgb() ?? 0; OnContentChanged(); }, Loc.T("Không màu", "No Color"))
            .Show(Tabs, Tabs.TabRect(_doc.Book.ActiveSheet)?.Location ?? Point.Empty);
    }

    public void CmdRowHeight()
    {
        var cur = S.RawRowHeight(Grid.Active.Row);
        var t = AskDialog.Ask(this, Loc.T("Chiều cao hàng", "Row Height"), Loc.T("Chiều cao hàng (điểm ảnh):", "Row height (pixels):"), cur.ToString("0"));
        if (t == null || !double.TryParse(t, out var h) || h < 0 || h > 600) return;
        _doc.Checkpoint(Loc.T("Chiều cao hàng", "Row height"));
        foreach (var r in SelRows()) { if (h == 0) S.HiddenRows.Add(r); else { S.HiddenRows.Remove(r); S.RowHeights[r] = h; } }
        Changed();
    }

    public void CmdColWidth()
    {
        var cur = S.RawColWidth(Grid.Active.Col);
        var t = AskDialog.Ask(this, Loc.T("Độ rộng cột", "Column Width"), Loc.T("Độ rộng cột (điểm ảnh):", "Column width (pixels):"), cur.ToString("0"));
        if (t == null || !double.TryParse(t, out var w) || w < 0 || w > 2000) return;
        _doc.Checkpoint(Loc.T("Độ rộng cột", "Column width"));
        foreach (var c in SelCols()) { if (w == 0) S.HiddenCols.Add(c); else { S.HiddenCols.Remove(c); S.ColWidths[c] = w; } }
        Changed();
    }

    public void SetHidden(bool rows, bool hide)
    {
        _doc.Checkpoint(hide ? Loc.T("Ẩn", "Hide") : Loc.T("Hiện", "Unhide"));
        var s = Grid.Selection;
        if (rows)
        {
            if (hide) foreach (var r in SelRows()) S.HiddenRows.Add(r);
            else S.HiddenRows.RemoveWhere(r => r >= Math.Max(0, s.R1 - 1) && r <= s.R2 + 1);
        }
        else
        {
            if (hide) foreach (var c in SelCols()) S.HiddenCols.Add(c);
            else S.HiddenCols.RemoveWhere(c => c >= Math.Max(0, s.C1 - 1) && c <= s.C2 + 1);
        }
        Changed();
    }

    // ------------------------------------------------------------------ editing commands

    public void CmdAutoSum(string fn)
    {
        if (!Guard()) return;
        var sel = Grid.Selection;
        bool IsNum(int r, int c) => S.GetValue(r, c).IsNumber;
        if (sel.IsSingle || S.MergeAt(sel.R1, sel.C1) == sel)
        {
            int r = Grid.Active.Row, c = Grid.Active.Col;
            string target;
            int up = r - 1;
            while (up >= 0 && IsNum(up, c)) up--;
            if (up < r - 1) target = Addr.ToA1(up + 1, c) + ":" + Addr.ToA1(r - 1, c);
            else
            {
                int left = c - 1;
                while (left >= 0 && IsNum(r, left)) left--;
                target = left < c - 1 ? Addr.ToA1(r, left + 1) + ":" + Addr.ToA1(r, c - 1) : "";
            }
            if (target.Length == 0) { Grid.BeginEdit("=" + fn + "(", true); return; }
            _doc.Checkpoint(Loc.T("Tự động tính", "AutoSum"));
            S.SetFormula(r, c, fn + "(" + target + ")");
            Changed();
            return;
        }
        _doc.Checkpoint(Loc.T("Tự động tính", "AutoSum"));
        for (int c = sel.C1; c <= sel.C2; c++)
        {
            int row = sel.R2;
            bool lastEmpty = S.GetValue(sel.R2, c).IsBlank && S.GetCell(sel.R2, c)?.Formula == null;
            int dest = lastEmpty ? sel.R2 : sel.R2 + 1;
            int end = lastEmpty ? sel.R2 - 1 : sel.R2;
            if (end < sel.R1) continue;
            S.SetFormula(dest, c, fn + "(" + Addr.ToA1(sel.R1, c) + ":" + Addr.ToA1(end, c) + ")");
        }
        Changed();
    }

    public void CmdInsertFunction()
    {
        using var d = new InsertFunctionDialog();
        if (Ux.Show(d, this) != DialogResult.OK || d.Selected == null) return;
        InsertFunctionCall(d.Selected);
    }

    private void InsertFunctionCall(string name)
    {
        if (!Grid.CanEditActive()) return;
        if (Grid.IsEditing) { Grid.InsertIntoEditor(name + "("); return; }
        Grid.BeginEdit("=" + name + "(", true);
    }

    private void InsertText(string text)
    {
        if (!Grid.CanEditActive()) return;
        if (Grid.IsEditing) Grid.InsertIntoEditor(text);
        else Grid.BeginEdit(GridView.EditableText(S.GetCell(Grid.Active.Row, Grid.Active.Col), S.StyleAt(Grid.Active.Row, Grid.Active.Col)) + text, false);
    }

    public void CmdFillDown()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        if (s.Rows < 2) { if (s.R1 == 0) return; s = s with { R1 = s.R1 - 1 }; }
        _doc.Checkpoint(Loc.T("Điền xuống", "Fill down"));
        SheetOps.Fill(S, s with { R2 = s.R1 }, s, copyOnly: true);
        Changed();
    }

    public void CmdFillRight()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        if (s.Cols < 2) { if (s.C1 == 0) return; s = s with { C1 = s.C1 - 1 }; }
        _doc.Checkpoint(Loc.T("Điền sang phải", "Fill right"));
        SheetOps.Fill(S, s with { C2 = s.C1 }, s, copyOnly: true);
        Changed();
    }

    public void CmdFillUp()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        if (s.Rows < 2) return;
        _doc.Checkpoint(Loc.T("Điền lên", "Fill up"));
        SheetOps.Fill(S, s with { R1 = s.R2 }, s, copyOnly: true);
        Changed();
    }

    public void CmdFillLeft()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        if (s.Cols < 2) return;
        _doc.Checkpoint(Loc.T("Điền sang trái", "Fill left"));
        SheetOps.Fill(S, s with { C1 = s.C2 }, s, copyOnly: true);
        Changed();
    }

    public void CmdFillSeries()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        if (s.IsSingle) { Ux.Info(this, Loc.T("Hãy chọn vùng cần điền, ô đầu tiên chứa giá trị bắt đầu.", "Select the range to fill; the first cell holds the start value.")); return; }
        var t = AskDialog.Ask(this, Loc.T("Chuỗi", "Series"), Loc.T("Bước nhảy:", "Step value:"), "1");
        if (t == null || !ValueParser.TryCoerceNumber(t, out var step)) return;
        var start = S.GetValue(s.R1, s.C1);
        if (!start.IsNumber) { Ux.Warn(this, Loc.T("Ô đầu tiên phải là số hoặc ngày.", "The first cell must contain a number or date.")); return; }
        _doc.Checkpoint(Loc.T("Điền chuỗi", "Fill series"));
        bool down = s.Rows >= s.Cols;
        int n = down ? s.Rows : s.Cols;
        var style = S.StyleIdAt(s.R1, s.C1);
        for (int i = 1; i < n; i++)
        {
            int r = down ? s.R1 + i : s.R1, c = down ? s.C1 : s.C1 + i;
            S.SetValue(r, c, CellValue.Number(start.Num + step * i));
            S.SetStyle(r, c, style);
        }
        Changed();
    }

    public void CmdNewNote()
    {
        var a = Grid.Active;
        var k = Sheet.Key(a.Row, a.Col);
        S.Notes.TryGetValue(k, out var cur);
        var t = AskDialog.Ask(this, Loc.T("Ghi chú", "Note"), Loc.F("Ghi chú cho ô {0}:", "Note for cell {0}:", Addr.ToA1(a.Row, a.Col)), cur ?? "");
        if (t == null) return;
        _doc.Checkpoint(Loc.T("Ghi chú", "Note"));
        if (t.Trim().Length == 0) S.Notes.Remove(k); else S.Notes[k] = t;
        _doc.Book.MarkDirty();
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    public void CmdDeleteNote()
    {
        var keys = S.Notes.Keys.Where(k => { var (r, c) = Sheet.FromKey(k); return Grid.AllSelections.Any(x => x.Contains(r, c)); }).ToList();
        if (keys.Count == 0) return;
        _doc.Checkpoint(Loc.T("Xóa ghi chú", "Delete note"));
        foreach (var k in keys) S.Notes.Remove(k);
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    private void CmdNextNote()
    {
        if (S.Notes.Count == 0) { Ux.Info(this, Loc.T("Trang tính không có ghi chú.", "This sheet has no notes.")); return; }
        var cur = Sheet.Key(Grid.Active.Row, Grid.Active.Col);
        var keys = S.Notes.Keys.OrderBy(k => k).ToList();
        var next = keys.FirstOrDefault(k => k > cur, keys[0]);
        var (r, c) = Sheet.FromKey(next);
        Grid.Select(r, c);
    }

    public void CmdHyperlink()
    {
        var a = Grid.Active;
        var k = Sheet.Key(a.Row, a.Col);
        S.Links.TryGetValue(k, out var cur);
        var url = AskDialog.Ask(this, Loc.T("Chèn liên kết", "Insert Hyperlink"), Loc.T("Địa chỉ (https://... hoặc #Sheet1!A1):", "Address (https://... or #Sheet1!A1):"), cur ?? "https://");
        if (url == null) return;
        _doc.Checkpoint(Loc.T("Liên kết", "Hyperlink"));
        url = url.Trim();
        if (url.Length == 0 || url == "https://") S.Links.Remove(k);
        else
        {
            S.Links[k] = url;
            if (S.GetValue(a.Row, a.Col).IsBlank) S.SetValue(a.Row, a.Col, CellValue.Text(url.TrimStart('#')));
        }
        Changed();
    }

    private void CmdRemoveHyperlink()
    {
        _doc.Checkpoint(Loc.T("Xóa liên kết", "Remove hyperlink"));
        foreach (var k in S.Links.Keys.ToList()) { var (r, c) = Sheet.FromKey(k); if (Grid.AllSelections.Any(x => x.Contains(r, c))) S.Links.Remove(k); }
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    // ------------------------------------------------------------------ sort / filter / data tools

    private CellRange DataRegion() => Grid.Selection.IsSingle ? SheetOps.CurrentRegion(S, Grid.Active.Row, Grid.Active.Col) : Grid.Selection;

    private bool LooksLikeHeader(CellRange r)
    {
        if (r.Rows < 2) return false;
        int textTop = 0, numBelow = 0;
        for (int c = r.C1; c <= r.C2; c++) { if (S.GetValue(r.R1, c).IsText) textTop++; if (S.GetValue(r.R1 + 1, c).IsNumber) numBelow++; }
        return textTop == r.Cols || (textTop > 0 && numBelow > 0) || S.StyleAt(r.R1, r.C1).Bold;
    }

    public void CmdSort(bool asc)
    {
        if (!Guard()) return;
        var r = S.Filter?.Range is CellRange fr && fr.Contains(Grid.Active.Row, Grid.Active.Col) ? fr : DataRegion();
        bool header = S.Filter != null || LooksLikeHeader(r);
        _doc.Checkpoint(Loc.T("Sắp xếp", "Sort"));
        SheetOps.Sort(S, r, [new SheetOps.SortKey(Grid.Active.Col, asc)], header);
        if (S.Filter != null) SheetOps.ApplyFilter(S);
        Changed();
    }

    public void CmdCustomSort()
    {
        if (!Guard()) return;
        var r = S.Filter?.Range ?? DataRegion();
        using var d = new SortDialog(S, r, S.Filter != null || LooksLikeHeader(r));
        if (Ux.Show(d, this) != DialogResult.OK) return;
        _doc.Checkpoint(Loc.T("Sắp xếp", "Sort"));
        SheetOps.Sort(S, r, d.Keys, d.HasHeader);
        if (S.Filter != null) SheetOps.ApplyFilter(S);
        Changed();
    }

    public void CmdToggleFilter()
    {
        _doc.Checkpoint(Loc.T("Bộ lọc", "Filter"));
        if (S.Filter != null) { S.Filter = null; S.FilteredRows.Clear(); }
        else
        {
            var r = DataRegion();
            S.Filter = new AutoFilter { Range = r };
        }
        Changed();
        UpdateCellInfo();
    }

    public void CmdClearFilter()
    {
        if (S.Filter == null) return;
        _doc.Checkpoint(Loc.T("Xóa lọc", "Clear filter"));
        S.Filter.Allowed.Clear(); S.Filter.Conditions.Clear();
        SheetOps.ApplyFilter(S);
        Changed();
    }

    public void CmdReapplyFilter()
    {
        if (S.Filter == null) return;
        SheetOps.ApplyFilter(S);
        Grid.RaiseContentChanged();
        Grid.RefreshAll();
    }

    /// <summary>Last filter popup shown (tests drive it directly).</summary>
    public FilterPopup? LastFilterPopup { get; private set; }

    public void ShowFilterPopup(int col)
    {
        var f = S.Filter;
        if (f == null) return;
        var used = S.UsedRange();
        var values = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
        for (int r = f.Range.R1 + 1; r <= Math.Max(f.Range.R2, used.R2); r++)
        {
            // values visible considering the other columns' filters
            values.Add(SheetOps.DisplayText(S, r, col));
            if (values.Count > 10000) break;
        }
        f.Allowed.TryGetValue(col, out var allowed);
        f.Conditions.TryGetValue(col, out var cond);
        var popup = new FilterPopup(values.ToList(), allowed, cond);
        popup.SortRequested += asc =>
        {
            _doc.Checkpoint(Loc.T("Sắp xếp", "Sort"));
            SheetOps.Sort(S, f.Range with { R2 = Math.Max(f.Range.R2, used.R2) }, [new SheetOps.SortKey(col, asc)], true);
            SheetOps.ApplyFilter(S);
            Changed();
        };
        popup.Applied += (set, c) => ApplyColumnFilter(col, set, c);
        LastFilterPopup = popup;
        var rect = Grid.FilterButtonRect(f.Range.R1, col) ?? new Rectangle(0, 0, 10, 10);
        if (Ux.DialogHook == null) popup.Show(Grid, new Point(rect.X, rect.Bottom));
    }

    public void ApplyColumnFilter(int col, HashSet<string>? allowed, string? cond)
    {
        var f = S.Filter;
        if (f == null) return;
        _doc.Checkpoint(Loc.T("Lọc", "Filter"));
        if (allowed == null) f.Allowed.Remove(col); else f.Allowed[col] = allowed;
        if (cond == null) f.Conditions.Remove(col); else f.Conditions[col] = cond;
        SheetOps.ApplyFilter(S);
        Changed();
        _statusMode.Text = Loc.F("Đang hiển thị {0} bản ghi khớp bộ lọc", "{0} records shown", Math.Max(0, Math.Max(f.Range.R2, S.UsedRange().R2) - f.Range.R1 - S.FilteredRows.Count));
    }

    public void CmdTextToColumns()
    {
        if (!Guard()) return;
        var s = Grid.Selection;
        if (s.Cols != 1) { Ux.Warn(this, Loc.T("Chỉ chọn một cột để tách.", "Select only one column to split.")); return; }
        using var d = new TextToColumnsDialog();
        if (Ux.Show(d, this) != DialogResult.OK) return;
        _doc.Checkpoint(Loc.T("Văn bản thành cột", "Text to columns"));
        var used = S.UsedRange();
        SheetOps.TextToColumns(S, s.C1, s.R1, Math.Min(s.R2, used.R2), d.Delimiter, d.Consecutive);
        Changed();
    }

    public void CmdRemoveDuplicates()
    {
        if (!Guard()) return;
        var r = DataRegion();
        using var d = new RemoveDuplicatesDialog(S, r, LooksLikeHeader(r));
        if (Ux.Show(d, this) != DialogResult.OK || d.Columns.Count == 0) return;
        _doc.Checkpoint(Loc.T("Loại bỏ trùng lặp", "Remove duplicates"));
        int n = SheetOps.RemoveDuplicates(S, r, d.Columns, d.HasHeader);
        Changed();
        Ux.Info(this, n > 0 ? Loc.F("Đã xóa {0} giá trị trùng lặp.", "{0} duplicate values found and removed.", n) : Loc.T("Không tìm thấy giá trị trùng lặp.", "No duplicate values found."));
    }

    public void CmdDataValidation()
    {
        var existing = Grid.ValidationAt(Grid.Active.Row, Grid.Active.Col);
        using var d = new ValidationDialog(existing);
        if (Ux.Show(d, this) != DialogResult.OK) return;
        _doc.Checkpoint(Loc.T("Xác thực dữ liệu", "Data validation"));
        var sel = Grid.Selection;
        S.Validations.RemoveAll(v => sel.Contains(v.Range));
        if (!d.ClearAll) { var dv = d.Build(sel); if (dv.Type != ValidationType.Any) S.Validations.Add(dv); }
        Changed();
    }

    private void CmdClearValidation()
    {
        _doc.Checkpoint(Loc.T("Xóa xác thực", "Clear validation"));
        S.Validations.RemoveAll(v => v.Range.Intersects(Grid.Selection));
        Changed();
    }

    public void CmdImportCsv()
    {
        var path = Ux.OpenFile(this, "CSV / Text (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt", Loc.T("Nhập dữ liệu", "Import data"));
        if (path == null) return;
        var src = CsvIO.Read(path).Sheets[0];
        _doc.Checkpoint(Loc.T("Nhập CSV", "Import CSV"));
        var name = CsvIO.SafeSheetName(Path.GetFileNameWithoutExtension(path));
        if (_doc.Book.Sheet(name) != null) name = _doc.Book.NextSheetName();
        var s = _doc.Book.AddSheet(name, _doc.Book.ActiveSheet + 1);
        foreach (var (k, cell) in src.AllCells)
        {
            var (r, c) = Sheet.FromKey(k);
            var nc = s.GetOrCreate(r, c);
            nc.Value = cell.Value; nc.Formula = cell.Formula;
            nc.StyleId = _doc.Book.Styles.Intern(src.Book.Styles[cell.StyleId]);
        }
        _doc.Recalc();
        SwitchSheet(_doc.Book.Sheets.IndexOf(s));
        OnContentChanged();
    }

    // ------------------------------------------------------------------ find / go to / select

    public void CmdFind(bool replace)
    {
        if (_find == null || _find.IsDisposed)
        {
            _find = new FindReplaceDialog(() => _doc, () => (_doc.Book, S), (sh, r, c) =>
            {
                int i = _doc.Book.Sheets.IndexOf(sh);
                if (i != _doc.Book.ActiveSheet) SwitchSheet(i);
                Grid.Select(r, c);
            }, () => { Grid.RaiseContentChanged(); Grid.Invalidate(); }, replace);
        }
        if (Ux.DialogHook != null) { Ux.DialogHook(_find); return; }
        if (!_find.Visible) _find.Show(this);
        _find.Activate();
    }

    public void CmdGoTo()
    {
        var t = AskDialog.Ask(this, Loc.T("Đi tới", "Go To"), Loc.T("Tham chiếu (vd. B12, A1:D20, 'Trang 2'!C5 hoặc tên):", "Reference (e.g. B12, A1:D20, Sheet2!C5 or a name):"), "");
        if (t != null) GoToReference(t);
    }

    public bool GoToReference(string text)
    {
        text = text.Trim();
        if (text.Length == 0) return false;
        if (_doc.Book.Names.TryGetValue(text, out var nameRef)) text = nameRef.TrimStart('=');
        string? sheetName = null;
        var refPart = text;
        int bang = text.LastIndexOf('!');
        if (bang > 0) { sheetName = text[..bang].Trim('\''); refPart = text[(bang + 1)..]; }
        if (!Addr.TryParseRange(refPart, out var r)) { Ux.Warn(this, Loc.T("Tham chiếu không hợp lệ.", "Reference isn't valid.")); return false; }
        if (sheetName != null)
        {
            var sh = _doc.Book.Sheet(sheetName);
            if (sh == null) { Ux.Warn(this, Loc.T("Không có trang tính này.", "That sheet doesn't exist.")); return false; }
            SwitchSheet(_doc.Book.Sheets.IndexOf(sh));
        }
        Grid.SelectRange(r);
        return true;
    }

    private void SelectWhere(Func<Cell, bool> pred)
    {
        var hits = S.AllCells.Where(kv => pred(kv.Value)).Select(kv => kv.Key).ToList();
        SelectWhereKeys(hits);
    }

    private void SelectWhereKeys(IEnumerable<long> keys)
    {
        var list = keys.OrderBy(k => k).ToList();
        if (list.Count == 0) { Ux.Info(this, Loc.T("Không tìm thấy ô nào.", "No cells were found.")); return; }
        var (r0, c0) = Sheet.FromKey(list[0]);
        Grid.Select(r0, c0);
        foreach (var k in list.Skip(1).Take(500)) { var (r, c) = Sheet.FromKey(k); Grid.Select(r, c, add: true); }
    }

    private void CmdSelectBlanks()
    {
        var sel = Grid.Selection;
        var used = S.UsedRange();
        var keys = new List<long>();
        for (int r = sel.R1; r <= Math.Min(sel.R2, used.R2); r++)
            for (int c = sel.C1; c <= Math.Min(sel.C2, used.C2); c++)
                if (S.GetValue(r, c).IsBlank && S.GetCell(r, c)?.Formula == null) keys.Add(Sheet.Key(r, c));
        SelectWhereKeys(keys);
    }

    // ------------------------------------------------------------------ insert objects

    public void CmdInsertPicture()
    {
        var path = Ux.OpenFile(this, Loc.T("Ảnh", "Pictures") + " (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif", Loc.T("Chèn ảnh", "Insert Picture"));
        if (path == null) return;
        try
        {
            using var img = Image.FromFile(path);
            Grid.InsertPicture(img);
        }
        catch (Exception ex) { Ux.Error(this, Loc.T("Không đọc được ảnh: ", "Could not read the picture: ") + ex.Message); }
    }

    public void CmdInsertChart(ChartType? type)
    {
        var r = Grid.Selection.IsSingle ? SheetOps.CurrentRegion(S, Grid.Active.Row, Grid.Active.Col) : Grid.Selection;
        if (r.IsSingle && !S.GetValue(r.R1, r.C1).IsNumber)
        {
            Ux.Info(this, Loc.T("Hãy chọn vùng dữ liệu (có tiêu đề cột và giá trị số) trước khi chèn biểu đồ.", "Select a data range (with column headers and numbers) before inserting a chart."));
            return;
        }
        var ch = new Chart
        {
            Type = type ?? ChartType.Column, Data = r, Name = "Chart " + (S.Charts.Count + 1),
            Title = r.Cols > 1 && S.GetValue(r.R1, r.C1 + 1).IsText && r.Cols == 2 ? SheetOps.DisplayText(S, r.R1, r.C1 + 1) : Loc.T("Tiêu đề biểu đồ", "Chart Title")
        };
        if (type == null)
        {
            using var d = new ChartDialog(S, ch);
            if (Ux.Show(d, this) != DialogResult.OK || !d.Apply(ch)) return;
        }
        float x = (float)Grid.SheetX(r.C2 + 1) + 16, y = (float)Grid.SheetY(r.R1);
        ch.Bounds = new RectangleF(x, y, 480, 290);
        _doc.Checkpoint(Loc.T("Chèn biểu đồ", "Insert chart"));
        S.Charts.Add(ch);
        Grid.SelectObject(ch);
        Grid.EnsureVisible(r.R1, r.C2 + 1);
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    public void CmdEditChart()
    {
        if (Grid.SelectedObject is not Chart ch) return;
        var copy = ch.Clone();
        using var d = new ChartDialog(S, copy);
        if (Ux.Show(d, this) != DialogResult.OK || !d.Apply(copy)) return;
        _doc.Checkpoint(Loc.T("Sửa biểu đồ", "Edit chart"));
        d.Apply(ch);
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    private void ChartChange(Action<Chart> a)
    {
        if (Grid.SelectedObject is not Chart ch) return;
        _doc.Checkpoint(Loc.T("Sửa biểu đồ", "Edit chart"));
        a(ch);
        Grid.RaiseContentChanged();
        Grid.Invalidate();
        UpdateContextualTabs();
    }

    private void SetChartType(ChartType t) => ChartChange(c => c.Type = t);

    private void CmdChartTitle()
    {
        if (Grid.SelectedObject is not Chart ch) return;
        var t = AskDialog.Ask(this, Loc.T("Tiêu đề biểu đồ", "Chart Title"), Loc.T("Tiêu đề:", "Title:"), ch.Title);
        if (t != null) ChartChange(c => c.Title = t);
    }

    private void CmdPictureResetSize()
    {
        if (Grid.SelectedObject is not Picture p) return;
        var img = Grid.PictureImage(p);
        if (img == null) return;
        _doc.Checkpoint(Loc.T("Kích thước ảnh", "Picture size"));
        p.Bounds = new RectangleF(p.Bounds.X, p.Bounds.Y, img.Width, img.Height);
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    // ------------------------------------------------------------------ page layout

    private void PageChange(string label, Action<PageSetup> a) { _doc.Checkpoint(label); a(S.Page); OnContentChanged(); _printPage?.RefreshPreview(); }
    private void SetMargins(double t, double b, double l, double r) => PageChange(Loc.T("Lề", "Margins"), p => { p.MarginTop = t; p.MarginBottom = b; p.MarginLeft = l; p.MarginRight = r; });
    private void SetOrientation(bool landscape) => PageChange(Loc.T("Hướng giấy", "Orientation"), p => p.Landscape = landscape);
    private void SetPaper(int id) => PageChange(Loc.T("Khổ giấy", "Paper size"), p => p.PaperSize = id);
    private void CmdSetPrintArea() { var sel = Grid.Selection; PageChange(Loc.T("Vùng in", "Print area"), p => p.PrintArea = sel); }
    private void CmdClearPrintArea() => PageChange(Loc.T("Vùng in", "Print area"), p => p.PrintArea = null);
    private void CmdPrintTitles() { var sel = Grid.Selection; PageChange(Loc.T("In tiêu đề", "Print titles"), p => p.PrintTitleRows = (sel.R1, sel.R2)); }
    private void CmdToggleFitWidth() => PageChange(Loc.T("Co vừa trang", "Fit to page"), p => p.FitToWidth = !p.FitToWidth);

    public void CmdPageSetup()
    {
        using var d = new PageSetupDialog(S.Page);
        if (Ux.Show(d, this) != DialogResult.OK) return;
        PageChange(Loc.T("Thiết lập trang", "Page setup"), d.Apply);
    }

    private void CmdPrintScale()
    {
        var t = AskDialog.Ask(this, Loc.T("Tỉ lệ in", "Scale"), Loc.T("Tỉ lệ in (10–400%):", "Scale (10–400%):"), S.Page.Scale.ToString());
        if (t != null && int.TryParse(t.TrimEnd('%'), out var v) && v >= 10 && v <= 400) PageChange(Loc.T("Tỉ lệ in", "Scale"), p => { p.Scale = v; p.FitToWidth = false; });
    }

    private void CmdToggleGridlines() { S.ShowGridlines = !S.ShowGridlines; _doc.Modified = true; Grid.RefreshAll(); UpdateCellInfo(); }
    private void CmdToggleHeadings() { S.ShowHeadings = !S.ShowHeadings; _doc.Modified = true; Grid.RefreshAll(); UpdateCellInfo(); }

    // ------------------------------------------------------------------ formulas

    public void CmdNameManager()
    {
        using var d = new NameManagerDialog(_doc.Book, () => Addr.QuoteSheet(S.Name) + "!" + AbsRef(Grid.Selection));
        _doc.Checkpoint(Loc.T("Quản lý tên", "Name manager"));
        Ux.Show(d, this);
        _doc.Book.Calc.InvalidateNames();
        Changed();
    }

    private static string AbsRef(CellRange r) => Addr.ToA1(r.R1, r.C1, true, true) + (r.IsSingle ? "" : ":" + Addr.ToA1(r.R2, r.C2, true, true));

    public void CmdDefineName()
    {
        var sel = Grid.Selection;
        var suggestion = S.GetValue(sel.R1, sel.C1).IsText && sel.Rows > 1 ? System.Text.RegularExpressions.Regex.Replace(S.GetValue(sel.R1, sel.C1).Str!, @"[^\p{L}\p{N}_]", "_") : "";
        var n = AskDialog.Ask(this, Loc.T("Định nghĩa tên", "New Name"), Loc.F("Tên cho vùng {0}:", "Name for {0}:", Addr.RangeToString(sel)), suggestion);
        if (string.IsNullOrWhiteSpace(n)) return;
        n = n.Trim();
        if (!System.Text.RegularExpressions.Regex.IsMatch(n, @"^[\p{L}_][\p{L}\p{N}_.]*$") || Addr.TryParseA1(n, out _)) { Ux.Warn(this, Loc.T("Tên không hợp lệ.", "The name is not valid.")); return; }
        _doc.Checkpoint(Loc.T("Định nghĩa tên", "Define name"));
        _doc.Book.Names[n] = Addr.QuoteSheet(S.Name) + "!" + AbsRef(sel);
        _doc.Book.Calc.InvalidateNames();
        Changed();
    }

    private void CmdToggleShowFormulas() { Grid.ShowFormulas = !Grid.ShowFormulas; Grid.Invalidate(); UpdateCellInfo(); }

    public void SetManualCalc(bool manual)
    {
        _doc.Book.ManualCalc = manual;
        _doc.Modified = true;
        if (!manual) CmdCalcNow();
        UpdateMode();
    }

    public void CmdCalcNow()
    {
        _doc.Book.Recalculate(true);
        Grid.RaiseContentChanged();
        Grid.Invalidate();
    }

    // ------------------------------------------------------------------ view

    public void CmdFreezePanes()
    {
        var s = S;
        if (s.FreezeRows > 0 || s.FreezeCols > 0) { SetFreeze(0, 0); return; }
        var a = Grid.Active;
        int fr = Math.Max(0, a.Row - s.ScrollRow) > 0 || a.Row > 0 ? a.Row : 0, fc = a.Col;
        if (fr == 0 && fc == 0) fr = 1;
        SetFreeze(fr, fc);
    }

    public void SetFreeze(int rows, int cols)
    {
        S.FreezeRows = rows; S.FreezeCols = cols;
        S.ScrollRow = rows; S.ScrollCol = cols;
        _doc.Modified = true;
        Grid.RefreshAll();
    }

    private void CmdZoomDialog()
    {
        var t = AskDialog.Ask(this, Loc.T("Thu phóng", "Zoom"), Loc.T("Tỉ lệ (10–400%):", "Magnification (10–400%):"), ((int)Math.Round(S.Zoom * 100)).ToString());
        if (t != null && int.TryParse(t.TrimEnd('%'), out var z)) Grid.SetZoom(z / 100.0);
    }

    private void CmdZoomToSelection()
    {
        var sel = Grid.Selection;
        double w = 0, h = 0;
        for (int c = sel.C1; c <= Math.Min(sel.C2, sel.C1 + 200); c++) w += S.ColWidth(c);
        for (int r = sel.R1; r <= Math.Min(sel.R2, sel.R1 + 500); r++) h += S.RowHeight(r);
        var area = Grid.CellArea;
        double z = Math.Min(area.Width / Math.Max(1, w * Dpi.Factor), area.Height / Math.Max(1, h * Dpi.Factor));
        Grid.SetZoom(Math.Clamp(Math.Floor(z * 20) / 20, 0.1, 4));
        Grid.ScrollTo(sel.R1, sel.C1);
    }

    // ------------------------------------------------------------------ protection

    public void CmdProtectSheet()
    {
        if (S.Protected)
        {
            if (!string.IsNullOrEmpty(S.ProtectionPasswordHash))
            {
                using var d = new ProtectDialog(true);
                if (Ux.Show(d, this) != DialogResult.OK) return;
                if (ProtectDialog.Hash(d.Password.Text) != S.ProtectionPasswordHash) { Ux.Warn(this, Loc.T("Mật khẩu không đúng.", "The password you supplied is not correct.")); return; }
            }
            _doc.Checkpoint(Loc.T("Bỏ bảo vệ", "Unprotect"));
            S.Protected = false; S.ProtectionPasswordHash = null;
        }
        else
        {
            using var d = new ProtectDialog(false);
            if (Ux.Show(d, this) != DialogResult.OK) return;
            _doc.Checkpoint(Loc.T("Bảo vệ trang tính", "Protect sheet"));
            S.Protected = true;
            S.ProtectionPasswordHash = d.Password.Text.Length > 0 ? ProtectDialog.Hash(d.Password.Text) : null;
        }
        OnContentChanged();
        UpdateCellInfo();
        Grid.Invalidate();
    }

    // ------------------------------------------------------------------ undo / redo

    public void CmdUndo()
    {
        if (Grid.IsEditing) { Grid.CancelEdit(); return; }
        if (!_doc.Undo()) return;
        AfterHistory();
    }

    public void CmdRedo()
    {
        if (!_doc.Redo()) return;
        AfterHistory();
    }

    private void AfterHistory()
    {
        Grid.Doc = _doc;
        Grid.RebindSelection();
        Tabs.Book = _doc.Book;
        UpdateAll();
    }

    // ------------------------------------------------------------------ files

    public const string OpenFilter = "Excel (*.xlsx;*.xlsm;*.xltx;*.csv;*.tsv;*.txt)|*.xlsx;*.xlsm;*.xltx;*.csv;*.tsv;*.txt|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|Text (*.txt;*.tsv)|*.txt;*.tsv";

    public void CmdOpen()
    {
        var path = Ux.OpenFile(this, OpenFilter, Loc.T("Mở sổ làm việc", "Open workbook"));
        if (path != null) { HideBackstage(); OpenFile(path); }
    }

    public bool OpenFile(string path)
    {
        try { path = Path.GetFullPath(path); } catch { }
        var existing = Docs.FirstOrDefault(d => string.Equals(d.FilePath, path, StringComparison.OrdinalIgnoreCase));
        if (existing != null) { Activate(existing); HideBackstage(); return true; }
        if (!File.Exists(path))
        {
            Ux.Warn(this, Loc.F("Không tìm thấy tệp:\n{0}", "File not found:\n{0}", path));
            OfficeConfig.RemoveRecent(AppKind.Excel, path);
            return false;
        }
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext == ".xls")
        {
            Ux.Warn(this, Loc.T("Định dạng Excel 97-2003 (.xls) chưa được hỗ trợ. Hãy mở tệp bằng phần mềm khác và lưu lại thành .xlsx.", "The Excel 97-2003 format (.xls) is not supported. Please save the file as .xlsx in another program."));
            return false;
        }
        Workbook wb;
        Cursor = Cursors.WaitCursor;
        try
        {
            wb = ext is ".csv" or ".tsv" or ".txt" ? CsvIO.Read(path) : XlsxReader.Read(path);
        }
        catch (Exception ex)
        {
            Ux.Error(this, Loc.F("Không mở được tệp \"{0}\". Tệp có thể bị hỏng hoặc không đúng định dạng.\n\n{1}", "Could not open \"{0}\". The file may be damaged or in an unsupported format.\n\n{1}", Path.GetFileName(path), ex.Message));
            return false;
        }
        finally { Cursor = Cursors.Default; }
        if (_doc != null && _doc.Pristine && Docs.Count == 1 && _doc.Book.Sheets.All(s => s.CellCount == 0))
        {
            Docs.Remove(_doc);
        }
        var doc = new WorkbookDoc(wb, ext == ".xltx" ? null : path);
        if (ext == ".xltx") doc.UntitledName = Path.GetFileNameWithoutExtension(path);
        Docs.Add(doc);
        Activate(doc);
        HideBackstage();
        if (ext != ".xltx") OfficeConfig.AddRecent(AppKind.Excel, path);
        return true;
    }

    public void CmdNewFromTemplate(string id)
    {
        var wb = ExcelTemplates.Create(id);
        if (_doc != null && _doc.Pristine && Docs.Count == 1 && _doc.Book.Sheets.All(s => s.CellCount == 0)) Docs.Remove(_doc);
        NewDocument(wb, null);
        HideBackstage();
    }

    public bool CmdSave()
    {
        if (Grid.IsEditing && !Grid.CommitEdit(0, 0)) return false;
        if (_doc.FilePath == null) return CmdSaveAs();
        return SaveTo(_doc, _doc.FilePath);
    }

    public bool CmdSaveAs()
    {
        if (Grid.IsEditing && !Grid.CommitEdit(0, 0)) return false;
        var filter = Loc.T("Sổ làm việc Excel", "Excel Workbook") + " (*.xlsx)|*.xlsx|CSV UTF-8 (*.csv)|*.csv|" + Loc.T("Văn bản (Tab)", "Text (Tab delimited)") + " (*.txt)|*.txt|PDF (*.pdf)|*.pdf|" + Loc.T("Trang web", "Web Page") + " (*.html)|*.html";
        var name = Path.GetFileNameWithoutExtension(_doc.FilePath ?? _doc.DisplayName) + ".xlsx";
        var path = Ux.SaveFile(this, filter, name);
        if (path == null) return false;
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext == ".pdf") { ExportPdfTo(path); return true; }
        if (ext is ".html" or ".htm") { CsvIO.WriteHtml(_doc.Book, path); return true; }
        if (ext is ".csv" or ".txt" or ".tsv")
        {
            CsvIO.Write(S, path, ext == ".csv" ? ',' : '\t');
            Ux.Info(this, Loc.T("Đã lưu trang tính hiện hoạt. Định dạng CSV/Text không lưu công thức, định dạng và các trang tính khác.", "The active sheet was saved. CSV/Text files don't keep formulas, formatting or other sheets."));
            return true;
        }
        if (!path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) path += ".xlsx";
        return SaveTo(_doc, path);
    }

    public bool SaveTo(WorkbookDoc doc, string path)
    {
        try
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext is ".csv" or ".tsv" or ".txt") CsvIO.Write(doc.Sheet, path, ext == ".csv" ? ',' : '\t');
            else XlsxWriter.Write(doc.Book, path);
            doc.FilePath = path;
            doc.Modified = false;
            DeleteRecover(doc);
            OfficeConfig.AddRecent(AppKind.Excel, path);
            UpdateTitle();
            if (_docTabs.Visible) RebuildDocTabs();
            _statusMode.Text = Loc.T("Đã lưu", "Saved") + " " + DateTime.Now.ToString("HH:mm");
            return true;
        }
        catch (Exception ex)
        {
            Ux.Error(this, Loc.F("Không lưu được tệp:\n{0}\n\n{1}", "Could not save the file:\n{0}\n\n{1}", path, ex.Message));
            return false;
        }
    }

    public bool CloseDocument(WorkbookDoc doc)
    {
        if (doc.Modified)
        {
            Activate(doc);
            var r = Ux.AskSave(this, doc.DisplayName);
            if (r == SaveChoice.Cancel) return false;
            if (r == SaveChoice.Save && !CmdSave()) return false;
        }
        DeleteRecover(doc);
        Docs.Remove(doc);
        if (Docs.Count == 0) { NewDocument(Workbook.CreateBlank(), null); ShowBackstage("home"); }
        else Activate(Docs[^1]);
        return true;
    }

    public SheetPrintDocument BuildPrintDocument(int what)
    {
        var sheets = what == 1 ? _doc.Book.Sheets.Where(s => !s.Hidden) : new[] { S };
        return new SheetPrintDocument(sheets, what == 2 ? Grid.Selection : null);
    }

    public void DoPrint(string? printer, short copies, int what)
    {
        var doc = BuildPrintDocument(what);
        try
        {
            if (printer != null) doc.PrinterSettings.PrinterName = printer;
            doc.PrinterSettings.Copies = copies;
            doc.Print();
            HideBackstage();
            _statusMode.Text = Loc.T("Đã gửi lệnh in", "Sent to printer");
        }
        catch (Exception ex) { Ux.Error(this, Loc.T("Không in được: ", "Printing failed: ") + ex.Message); }
    }

    public void CmdExportPdf()
    {
        var path = Ux.SaveFile(this, "PDF (*.pdf)|*.pdf", Path.GetFileNameWithoutExtension(_doc.FilePath ?? _doc.DisplayName) + ".pdf");
        if (path != null) ExportPdfTo(path);
    }

    public void ExportPdfTo(string path, bool forceImage = false)
    {
        Cursor = Cursors.WaitCursor;
        try
        {
            PdfExporter.Export(BuildPrintDocument(0), path, forceImage);
            _statusMode.Text = Loc.T("Đã xuất PDF", "PDF exported");
        }
        catch (Exception ex) { Ux.Error(this, Loc.T("Không xuất được PDF: ", "Could not export PDF: ") + ex.Message); }
        finally { Cursor = Cursors.Default; }
    }

    private void CmdExportDelimited(char d)
    {
        var path = Ux.SaveFile(this, d == ',' ? "CSV UTF-8 (*.csv)|*.csv" : "Text (*.txt)|*.txt", Path.GetFileNameWithoutExtension(_doc.FilePath ?? _doc.DisplayName) + (d == ',' ? ".csv" : ".txt"));
        if (path == null) return;
        CsvIO.Write(S, path, d);
        _statusMode.Text = Loc.T("Đã xuất", "Exported");
    }

    private void CmdExportHtml()
    {
        var path = Ux.SaveFile(this, "HTML (*.html)|*.html", Path.GetFileNameWithoutExtension(_doc.FilePath ?? _doc.DisplayName) + ".html");
        if (path == null) return;
        CsvIO.WriteHtml(_doc.Book, path);
        _statusMode.Text = Loc.T("Đã xuất", "Exported");
    }

    private void CmdOptions()
    {
        var style = OfficeConfig.GetAppData("excel.numberStyle") ?? "auto";
        int minutes = int.TryParse(OfficeConfig.GetAppData("excel.autoRecover"), out var m) ? m : 2;
        using var d = new ExcelOptionsDialog(style, _doc.Book.ManualCalc, minutes);
        if (Ux.Show(d, this) != DialogResult.OK) return;
        OfficeConfig.SetAppData("excel.numberStyle", d.NumberStyle);
        OfficeConfig.SetAppData("excel.autoRecover", d.AutoRecoverMinutes.ToString());
        ApplyNumberStyle();
        SetManualCalc(d.ManualCalc);
        Grid.Invalidate();
        UpdateCellInfo();
    }

    private void CmdHelp()
    {
        Ux.Info(this, Loc.T(
            "MICROADS OFFICE EXCEL — PHÍM TẮT\n\n" +
            "Nhập liệu: gõ trực tiếp vào ô, F2 sửa ô, Alt+Enter xuống dòng trong ô, Ctrl+Enter nhập cho cả vùng chọn, Esc hủy.\n" +
            "Công thức: bắt đầu bằng =, ngăn cách đối số bằng dấu phẩy, ví dụ =SUM(A1:A10), =VLOOKUP(A2,Bang!A:C,3,FALSE), =DOCSO(F18).\n" +
            "Khi đang gõ công thức có thể bấm chuột/mũi tên để chọn ô tham chiếu; F4 đổi $A$1 ↔ A1.\n" +
            "Di chuyển: mũi tên, Ctrl+mũi tên tới mép dữ liệu, Ctrl+Home về A1, Ctrl+End ô cuối, Ctrl+PageUp/PageDown đổi trang tính.\n" +
            "Chọn: Shift+mũi tên, Ctrl+bấm chọn nhiều vùng, Ctrl+Space chọn cột, Shift+Space chọn hàng, Ctrl+A chọn vùng/toàn bộ.\n" +
            "Định dạng: Ctrl+B/I/U, Ctrl+1 Định dạng ô, Ctrl+Shift+% phần trăm, Ctrl+Shift+$ tiền tệ.\n" +
            "Lệnh: Ctrl+C/X/V, Ctrl+Alt+V dán đặc biệt, Ctrl+Z/Y, Ctrl+F/H tìm/thay, Ctrl+G đi tới, Ctrl+D/R điền, Alt+= tự tính tổng,\n" +
            "Ctrl+; ngày hôm nay, Ctrl+Shift+: giờ hiện tại, Ctrl+K liên kết, Shift+F2 ghi chú, Ctrl+Shift+L bộ lọc, F9 tính lại, Shift+F11 thêm trang tính.\n" +
            "Kéo ô vuông nhỏ ở góc vùng chọn (tay nắm điền) để điền chuỗi số, ngày, thứ, tháng hoặc sao chép công thức.",
            "MICROADS OFFICE EXCEL — KEYBOARD SHORTCUTS\n\n" +
            "Entering data: type into a cell, F2 edits, Alt+Enter new line in a cell, Ctrl+Enter fills the selection, Esc cancels.\n" +
            "Formulas start with =, arguments are separated by commas, e.g. =SUM(A1:A10), =VLOOKUP(A2,Table!A:C,3,FALSE).\n" +
            "While typing a formula, click cells or use the arrow keys to insert references; F4 toggles $A$1 ↔ A1.\n" +
            "Navigation: arrows, Ctrl+arrows to the data edge, Ctrl+Home to A1, Ctrl+End last cell, Ctrl+PageUp/PageDown change sheet.\n" +
            "Selection: Shift+arrows, Ctrl+click for several ranges, Ctrl+Space column, Shift+Space row, Ctrl+A region/all.\n" +
            "Formatting: Ctrl+B/I/U, Ctrl+1 Format Cells, Ctrl+Shift+% percent, Ctrl+Shift+$ currency.\n" +
            "Commands: Ctrl+C/X/V, Ctrl+Alt+V paste special, Ctrl+Z/Y, Ctrl+F/H, Ctrl+G go to, Ctrl+D/R fill, Alt+= AutoSum,\n" +
            "Ctrl+; today, Ctrl+Shift+: now, Ctrl+K link, Shift+F2 note, Ctrl+Shift+L filter, F9 calculate, Shift+F11 new sheet.\n" +
            "Drag the small square at the corner of the selection (fill handle) to fill series or copy formulas."));
    }

    // ------------------------------------------------------------------ context menus

    private ContextMenuStrip CellMenu()
    {
        if (Grid.SelectedObject is Chart)
            return Menu(Mi("cmChartEdit", "Chọn dữ liệu...", "Select Data...", XlIcons.Small("editData"), CmdEditChart), Mi("cmChartDelete", "Xóa biểu đồ", "Delete Chart", XlIcons.Small("deleteObj"), () => Grid.DeleteSelectedObject()));
        if (Grid.SelectedObject is Picture)
            return Menu(Mi("cmPicReset", "Kích thước gốc", "Reset Size", null, CmdPictureResetSize), Mi("cmPicDelete", "Xóa ảnh", "Delete Picture", XlIcons.Small("deleteObj"), () => Grid.DeleteSelectedObject()));
        return Menu(
            Mi("cmCut", "Cắt", "Cut", Icons.Small(Icons.Cut), () => Grid.CopySelection(true)),
            Mi("cmCopy", "Sao chép", "Copy", Icons.Small(Icons.Copy), () => Grid.CopySelection(false)),
            Mi("cmPaste", "Dán", "Paste", Icons.Small(Icons.Paste), () => Grid.Paste()),
            Mi("cmPasteValues", "Dán giá trị", "Paste Values", null, () => Grid.Paste(PasteMode.Values)),
            Mi("cmPasteSpecial", "Dán đặc biệt...", "Paste Special...", null, CmdPasteSpecial),
            new ToolStripSeparator(),
            Mi("cmInsertRows", "Chèn hàng", "Insert Rows", XlIcons.Small("insertRow"), CmdInsertRows),
            Mi("cmInsertCols", "Chèn cột", "Insert Columns", XlIcons.Small("insertCol"), CmdInsertCols),
            Mi("cmDeleteRows", "Xóa hàng", "Delete Rows", XlIcons.Small("deleteRow"), CmdDeleteRows),
            Mi("cmDeleteCols", "Xóa cột", "Delete Columns", XlIcons.Small("deleteCol"), CmdDeleteCols),
            Mi("cmClear", "Xóa nội dung", "Clear Contents", XlIcons.Small("clear"), () => Grid.ClearSelection(true, false)),
            new ToolStripSeparator(),
            Mi("cmSortAsc", "Sắp xếp A → Z", "Sort A to Z", XlIcons.Small("sortAsc"), () => CmdSort(true)),
            Mi("cmSortDesc", "Sắp xếp Z → A", "Sort Z to A", XlIcons.Small("sortDesc"), () => CmdSort(false)),
            Mi("cmFilter", "Bộ lọc", "Filter", XlIcons.Small("filter"), CmdToggleFilter),
            new ToolStripSeparator(),
            Mi("cmNote", "Chèn/Sửa ghi chú", "Insert/Edit Note", XlIcons.Small("note"), CmdNewNote),
            Mi("cmLink", "Liên kết...", "Link...", XlIcons.Small("link"), CmdHyperlink),
            Mi("cmDefineName", "Định nghĩa tên...", "Define Name...", XlIcons.Small("names"), CmdDefineName),
            Mi("cmFormatCells", "Định dạng ô...", "Format Cells...", XlIcons.Small("format"), () => CmdFormatCells(0)));
    }

    private ContextMenuStrip HeaderMenu(bool cols) => cols
        ? Menu(
            Mi("hmCut", "Cắt", "Cut", Icons.Small(Icons.Cut), () => Grid.CopySelection(true)), Mi("hmCopy", "Sao chép", "Copy", Icons.Small(Icons.Copy), () => Grid.CopySelection(false)),
            Mi("hmPaste", "Dán", "Paste", Icons.Small(Icons.Paste), () => Grid.Paste()), new ToolStripSeparator(),
            Mi("hmInsertCols", "Chèn cột", "Insert", XlIcons.Small("insertCol"), CmdInsertCols), Mi("hmDeleteCols", "Xóa cột", "Delete", XlIcons.Small("deleteCol"), CmdDeleteCols),
            Mi("hmClearCols", "Xóa nội dung", "Clear Contents", null, () => Grid.ClearSelection(true, false)), new ToolStripSeparator(),
            Mi("hmColWidth", "Độ rộng cột...", "Column Width...", XlIcons.Small("colWidth"), CmdColWidth), Mi("hmAutoFitCol", "Tự vừa độ rộng", "AutoFit", null, () => Grid.AutoFitColumns(SelCols())),
            Mi("hmHideCols", "Ẩn", "Hide", null, () => SetHidden(false, true)), Mi("hmUnhideCols", "Hiện", "Unhide", null, () => SetHidden(false, false)))
        : Menu(
            Mi("hmCutR", "Cắt", "Cut", Icons.Small(Icons.Cut), () => Grid.CopySelection(true)), Mi("hmCopyR", "Sao chép", "Copy", Icons.Small(Icons.Copy), () => Grid.CopySelection(false)),
            Mi("hmPasteR", "Dán", "Paste", Icons.Small(Icons.Paste), () => Grid.Paste()), new ToolStripSeparator(),
            Mi("hmInsertRows", "Chèn hàng", "Insert", XlIcons.Small("insertRow"), CmdInsertRows), Mi("hmDeleteRows", "Xóa hàng", "Delete", XlIcons.Small("deleteRow"), CmdDeleteRows),
            Mi("hmClearRows", "Xóa nội dung", "Clear Contents", null, () => Grid.ClearSelection(true, false)), new ToolStripSeparator(),
            Mi("hmRowHeight", "Chiều cao hàng...", "Row Height...", XlIcons.Small("rowHeight"), CmdRowHeight), Mi("hmAutoFitRow", "Tự vừa chiều cao", "AutoFit", null, () => Grid.AutoFitRows(SelRows())),
            Mi("hmHideRows", "Ẩn", "Hide", null, () => SetHidden(true, true)), Mi("hmUnhideRows", "Hiện", "Unhide", null, () => SetHidden(true, false)));

    public ContextMenuStrip SheetMenu() => Menu(
        Mi("smInsert", "Chèn trang tính", "Insert Sheet", XlIcons.Small("newSheetTab"), CmdInsertSheet),
        Mi("smDelete", "Xóa", "Delete", XlIcons.Small("deleteObj"), CmdDeleteSheet),
        Mi("smRename", "Đổi tên", "Rename", null, CmdRenameSheet),
        Mi("smDuplicate", "Nhân bản", "Duplicate", null, CmdDuplicateSheet),
        Mi("smMoveLeft", "Chuyển sang trái", "Move Left", null, () => { int i = _doc.Book.ActiveSheet; if (i > 0) MoveSheet(i, i - 1); }),
        Mi("smMoveRight", "Chuyển sang phải", "Move Right", null, () => { int i = _doc.Book.ActiveSheet; if (i < _doc.Book.Sheets.Count - 1) MoveSheet(i, i + 1); }),
        Mi("smTabColor", "Màu thẻ...", "Tab Color...", null, CmdTabColor),
        Mi("smHide", "Ẩn", "Hide", null, () =>
        {
            if (_doc.Book.Sheets.Count(s => !s.Hidden) <= 1) return;
            _doc.Checkpoint(Loc.T("Ẩn trang tính", "Hide sheet"));
            S.Hidden = true;
            int next = _doc.Book.Sheets.FindIndex(s => !s.Hidden);
            SwitchSheet(next);
            OnContentChanged();
        }),
        Mi("smUnhide", "Hiện trang tính ẩn", "Unhide", null, () =>
        {
            var hidden = _doc.Book.Sheets.Where(s => s.Hidden).ToList();
            if (hidden.Count == 0) return;
            _doc.Checkpoint(Loc.T("Hiện trang tính", "Unhide sheet"));
            foreach (var h in hidden) h.Hidden = false;
            OnContentChanged();
        }),
        Mi("smProtect", "Bảo vệ trang tính...", "Protect Sheet...", XlIcons.Small("protect"), CmdProtectSheet));

    // ------------------------------------------------------------------ keyboard

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (BackstageView.Visible)
        {
            if (keyData == Keys.Escape) { HideBackstage(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        bool textFocus = ActiveControl is TextBox or ComboBox || (ActiveControl is ContainerControl cc && cc.ActiveControl is TextBox) || Grid.IsEditing;
        Action? a = keyData switch
        {
            Keys.Control | Keys.S => () => CmdSave(),
            Keys.F12 => () => CmdSaveAs(),
            Keys.Control | Keys.O => CmdOpen,
            Keys.Control | Keys.N => () => CmdNewFromTemplate("blank"),
            Keys.Control | Keys.P => () => ShowBackstage("print"),
            Keys.Control | Keys.W => () => CloseDocument(_doc),
            Keys.Control | Keys.F1 => () => Ribbon.SetMinimized(!Ribbon.Minimized),
            Keys.F1 => CmdHelp,
            Keys.Shift | Keys.F3 => CmdInsertFunction,
            Keys.F9 => CmdCalcNow,
            _ => null
        };
        if (a == null && !textFocus)
        {
            a = keyData switch
            {
                Keys.Control | Keys.Z => CmdUndo,
                Keys.Control | Keys.Y => CmdRedo,
                Keys.Control | Keys.C => () => Grid.CopySelection(false),
                Keys.Control | Keys.X => () => Grid.CopySelection(true),
                Keys.Control | Keys.V => () => Grid.Paste(),
                Keys.Control | Keys.Alt | Keys.V => CmdPasteSpecial,
                Keys.Control | Keys.B => () => ToggleStyle(s => s.Bold, (s, v) => s with { Bold = v }),
                Keys.Control | Keys.I => () => ToggleStyle(s => s.Italic, (s, v) => s with { Italic = v }),
                Keys.Control | Keys.U => () => ToggleStyle(s => s.Underline, (s, v) => s with { Underline = v }),
                Keys.Control | Keys.D5 => () => ToggleStyle(s => s.Strike, (s, v) => s with { Strike = v }),
                Keys.Control | Keys.D1 => () => CmdFormatCells(0),
                Keys.Control | Keys.F => () => CmdFind(false),
                Keys.Control | Keys.H => () => CmdFind(true),
                Keys.Control | Keys.G or Keys.F5 => CmdGoTo,
                Keys.Control | Keys.D => CmdFillDown,
                Keys.Control | Keys.R => CmdFillRight,
                Keys.Control | Keys.K => CmdHyperlink,
                Keys.Shift | Keys.F2 => CmdNewNote,
                Keys.Shift | Keys.F11 => CmdInsertSheet,
                Keys.Control | Keys.Shift | Keys.L => CmdToggleFilter,
                Keys.Alt | Keys.Oemplus => () => CmdAutoSum("SUM"),
                Keys.Control | Keys.Oemtilde => CmdToggleShowFormulas,
                Keys.Control | Keys.Shift | Keys.D5 => () => SetNumberFormat("0%"),
                Keys.Control | Keys.Shift | Keys.D4 => () => SetNumberFormat("#,##0 \"₫\""),
                Keys.Control | Keys.Shift | Keys.D1 => () => SetNumberFormat("#,##0.00"),
                Keys.Control | Keys.Shift | Keys.D3 => () => SetNumberFormat(NumFmt.ShortDateCode),
                Keys.Control | Keys.OemSemicolon => () => InsertNow(false),
                Keys.Control | Keys.Shift | Keys.OemSemicolon => () => InsertNow(true),
                Keys.Control | Keys.PageDown => () => SwitchSheet(Math.Min(_doc.Book.Sheets.Count - 1, _doc.Book.ActiveSheet + 1)),
                Keys.Control | Keys.PageUp => () => SwitchSheet(Math.Max(0, _doc.Book.ActiveSheet - 1)),
                Keys.Control | Keys.A => SelectAll,
                Keys.Control | Keys.Oemplus or Keys.Control | Keys.Shift | Keys.Oemplus or Keys.Control | Keys.Add => InsertBySelection,
                Keys.Control | Keys.OemMinus or Keys.Control | Keys.Subtract => DeleteBySelection,
                _ => null
            };
        }
        else if (a == null && Grid.IsEditing)
        {
            a = keyData switch
            {
                Keys.Control | Keys.OemSemicolon => () => Grid.InsertIntoEditor(DateTime.Today.ToString(NumFmt.DayFirst ? "dd/MM/yyyy" : "M/d/yyyy")),
                Keys.Control | Keys.Shift | Keys.OemSemicolon => () => Grid.InsertIntoEditor(DateTime.Now.ToString("HH:mm")),
                _ => null
            };
        }
        if (a != null)
        {
            RunCommand("key", a);
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void InsertNow(bool time)
    {
        if (!Grid.CanEditActive()) return;
        var a = Grid.Active;
        _doc.Checkpoint(Loc.T("Nhập ngày giờ", "Insert date/time"));
        var now = DateTime.Now;
        S.SetValue(a.Row, a.Col, CellValue.Number(time ? DateSerial.ToSerial(now) - Math.Floor(DateSerial.ToSerial(now)) : DateSerial.ToSerial(now.Date)));
        if (S.StyleAt(a.Row, a.Col).NumberFormat == "General")
            S.SetStyle(a.Row, a.Col, _doc.Book.Restyle(S.StyleIdAt(a.Row, a.Col), s => s with { NumberFormat = time ? "h:mm" : NumFmt.ShortDateCode }));
        Changed();
    }

    private void SelectAll()
    {
        var region = SheetOps.CurrentRegion(S, Grid.Active.Row, Grid.Active.Col);
        if (!region.IsSingle && Grid.Selection != region) Grid.SelectRange(region, Grid.Active);
        else Grid.SelectRange(new CellRange(0, 0, Addr.MaxRows - 1, Addr.MaxCols - 1), Grid.Active);
    }

    private void InsertBySelection()
    {
        var s = Grid.Selection;
        if (s.C1 == 0 && s.C2 == Addr.MaxCols - 1) CmdInsertRows();
        else if (s.R1 == 0 && s.R2 == Addr.MaxRows - 1) CmdInsertCols();
        else CmdInsertRows();
    }

    private void DeleteBySelection()
    {
        var s = Grid.Selection;
        if (s.R1 == 0 && s.R2 == Addr.MaxRows - 1) CmdDeleteCols();
        else CmdDeleteRows();
    }

    // ------------------------------------------------------------------ closing / autorecover

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (Grid.IsEditing) Grid.CommitEdit(0, 0);
        foreach (var d in Docs.ToList())
        {
            if (!d.Modified) continue;
            Activate(d);
            HideBackstage();
            var r = Ux.AskSave(this, d.DisplayName);
            if (r == SaveChoice.Cancel) { e.Cancel = true; return; }
            if (r == SaveChoice.Save && !CmdSave()) { e.Cancel = true; return; }
        }
        foreach (var d in Docs) DeleteRecover(d);
        _autoRecover.Stop();
        SaveWindow();
        _find?.Dispose();
    }

    private static string RecoverDir => Path.Combine(AppSettings.DataDirectory, "AutoRecover", "excel");

    public void AutoRecoverTick(bool force = false)
    {
        int minutes = int.TryParse(OfficeConfig.GetAppData("excel.autoRecover"), out var m) ? m : 2;
        foreach (var d in Docs)
        {
            if (!d.Modified) continue;
            if (!force && (DateTime.Now - d.LastAutoRecover).TotalMinutes < minutes) continue;
            try
            {
                Directory.CreateDirectory(RecoverDir);
                d.RecoverPath ??= Path.Combine(RecoverDir, Guid.NewGuid().ToString("N") + ".xlsx");
                XlsxWriter.Write(d.Book, d.RecoverPath);
                File.WriteAllText(d.RecoverPath + ".txt", (d.FilePath ?? "") + "\n" + d.DisplayName);
                d.LastAutoRecover = DateTime.Now;
            }
            catch { }
        }
    }

    private static void DeleteRecover(WorkbookDoc d)
    {
        if (d.RecoverPath == null) return;
        try { File.Delete(d.RecoverPath); File.Delete(d.RecoverPath + ".txt"); } catch { }
        d.RecoverPath = null;
    }

    /// <summary>After a crash: offers to reopen the AutoRecover copies.</summary>
    public bool OfferRecovery()
    {
        try
        {
            if (!Directory.Exists(RecoverDir)) return false;
            var files = Directory.GetFiles(RecoverDir, "*.xlsx");
            if (files.Length == 0) return false;
            var names = files.Select(f => File.Exists(f + ".txt") ? File.ReadAllLines(f + ".txt").LastOrDefault() ?? Path.GetFileName(f) : Path.GetFileName(f)).ToList();
            if (!Ux.Confirm(this, Loc.T("Microads Office Excel đã đóng đột ngột. Khôi phục các sổ làm việc chưa lưu?\n\n", "Microads Office Excel closed unexpectedly. Recover the unsaved workbooks?\n\n") + string.Join("\n", names.Select(n => "• " + n)),
                    Loc.T("Khôi phục", "Recover"), Loc.T("Bỏ qua", "Discard")))
            {
                foreach (var f in files) { try { File.Delete(f); File.Delete(f + ".txt"); } catch { } }
                return false;
            }
            bool any = false;
            for (int i = 0; i < files.Length; i++)
            {
                try
                {
                    var wb = XlsxReader.Read(files[i]);
                    var meta = File.Exists(files[i] + ".txt") ? File.ReadAllLines(files[i] + ".txt") : [];
                    if (_doc.Pristine && Docs.Count == 1) Docs.Remove(_doc);
                    var doc = new WorkbookDoc(wb, null) { Modified = true, UntitledName = (meta.LastOrDefault() ?? "Book") + Loc.T(" (đã khôi phục)", " (recovered)") };
                    Docs.Add(doc);
                    Activate(doc);
                    any = true;
                }
                catch { }
                try { File.Delete(files[i]); File.Delete(files[i] + ".txt"); } catch { }
            }
            return any;
        }
        catch { return false; }
    }
}
