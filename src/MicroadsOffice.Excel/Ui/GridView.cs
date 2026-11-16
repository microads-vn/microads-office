using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

public enum HitArea { None, Corner, ColHeader, RowHeader, Cell, FillHandle, ColBorder, RowBorder, Object, ObjectHandle, FilterButton, ValidationArrow }

public readonly record struct Hit(HitArea Area, int Row, int Col, int Handle = -1, object? Obj = null);

/// <summary>
/// Custom painted, virtualised spreadsheet grid: headers, frozen panes, selection (multi-range), fill handle,
/// in-cell editor with point-to-reference, charts/pictures on top, autofilter buttons, validation lists.
/// </summary>
public sealed partial class GridView : Control
{
    private WorkbookDoc _doc = null!;
    private readonly VScrollBar _vbar = new() { Name = "gridVScroll", SmallChange = 1, LargeChange = 10 };
    private readonly HScrollBar _hbar = new() { Name = "gridHScroll", SmallChange = 1, LargeChange = 5 };
    private readonly List<(int Index, int Pos, int Size)> _rows = new();
    private readonly List<(int Index, int Pos, int Size)> _cols = new();
    private int _frozenBottom, _frozenRight;
    private bool _syncingBars;
    internal readonly List<CellRange> ExtraRanges = new();
    /// <summary>Range shown with a dashed "marching ants" border after Copy/Cut.</summary>
    public CellRange? CopyMarquee { get; set; }
    public bool ShowFormulas { get; set; }
    public object? SelectedObject { get; private set; }

    public event EventHandler? SelectionChanged;
    public event EventHandler? ContentChanged;
    public event EventHandler? EditStateChanged;
    public event EventHandler? ObjectSelectionChanged;
    public event EventHandler<int>? ZoomChanged;

    public GridView()
    {
        Name = "grid";
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
        TabStop = true;
        BackColor = Color.White;
        AccessibleName = Loc.T("Lưới bảng tính", "Spreadsheet grid");
        AccessibleRole = AccessibleRole.Table;
        Controls.Add(_vbar);
        Controls.Add(_hbar);
        _vbar.Scroll += (_, e) => { if (!_syncingBars) ScrollTo(e.NewValue, Sheet.ScrollCol); };
        _hbar.Scroll += (_, e) => { if (!_syncingBars) ScrollTo(Sheet.ScrollRow, e.NewValue); };
        InitEditor();
    }

    public WorkbookDoc Doc
    {
        get => _doc;
        set { _doc = value; ExtraRanges.Clear(); CopyMarquee = null; SelectedObject = null; CancelEdit(); RefreshAll(); }
    }

    public Sheet Sheet => _doc.Sheet;
    public Workbook Book => _doc.Book;
    public float Scale => (float)Sheet.Zoom * Dpi.Factor;
    public int ColHeaderH => Sheet.ShowHeadings ? (int)Math.Round(20 * Scale) : 0;
    public int RowHeaderW => Sheet.ShowHeadings ? (int)Math.Round(Math.Max(28, 9 + 8 * (Sheet.ScrollRow + 60).ToString().Length) * Scale) : 0;
    private int BarW => SystemInformation.VerticalScrollBarWidth;
    private int BarH => SystemInformation.HorizontalScrollBarHeight;
    public Rectangle CellArea => new(RowHeaderW, ColHeaderH, Math.Max(0, Width - RowHeaderW - BarW), Math.Max(0, Height - ColHeaderH - BarH));

    public CellPos Active => Sheet.ActiveCell;
    public CellRange Selection => Sheet.Selection.Normalized;
    public IEnumerable<CellRange> AllSelections => new[] { Selection }.Concat(ExtraRanges);

    // ------------------------------------------------------------------ layout

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        _vbar.SetBounds(Width - BarW, 0, BarW, Math.Max(0, Height - BarH));
        _hbar.SetBounds(0, Height - BarH, Math.Max(0, Width - BarW), BarH);
        ComputeLayout();
        UpdateScrollBars();
        PositionEditor();
    }

    internal void ComputeLayout()
    {
        if (_doc == null) return;
        var s = Sheet;
        float sc = Scale;
        _rows.Clear(); _cols.Clear();
        int y = ColHeaderH, bottom = Height - BarH;
        for (int r = 0; r < s.FreezeRows && r < Addr.MaxRows; r++)
        {
            int h = (int)Math.Round(s.RowHeight(r) * sc);
            if (h > 0) { _rows.Add((r, y, h)); y += h; }
        }
        _frozenBottom = y;
        if (s.ScrollRow < s.FreezeRows) s.ScrollRow = s.FreezeRows;
        for (int r = s.ScrollRow; r < Addr.MaxRows && y < bottom; r++)
        {
            int h = (int)Math.Round(s.RowHeight(r) * sc);
            if (h <= 0) continue;
            _rows.Add((r, y, h)); y += h;
        }
        int x = RowHeaderW, right = Width - BarW;
        for (int c = 0; c < s.FreezeCols && c < Addr.MaxCols; c++)
        {
            int w = (int)Math.Round(s.ColWidth(c) * sc);
            if (w > 0) { _cols.Add((c, x, w)); x += w; }
        }
        _frozenRight = x;
        if (s.ScrollCol < s.FreezeCols) s.ScrollCol = s.FreezeCols;
        for (int c = s.ScrollCol; c < Addr.MaxCols && x < right; c++)
        {
            int w = (int)Math.Round(s.ColWidth(c) * sc);
            if (w <= 0) continue;
            _cols.Add((c, x, w)); x += w;
        }
    }

    internal void UpdateScrollBars()
    {
        if (_doc == null) return;
        _syncingBars = true;
        try
        {
            var s = Sheet;
            var used = s.UsedRange();
            int visRows = Math.Max(1, _rows.Count(r => r.Index >= s.FreezeRows) - 1);
            int visCols = Math.Max(1, _cols.Count(c => c.Index >= s.FreezeCols) - 1);
            int maxRow = Math.Min(Addr.MaxRows - 1, Math.Max(used.R2 + visRows, s.ScrollRow + visRows * 2));
            int maxCol = Math.Min(Addr.MaxCols - 1, Math.Max(used.C2 + visCols, s.ScrollCol + visCols * 2));
            _vbar.Minimum = s.FreezeRows; _vbar.Maximum = Math.Max(s.FreezeRows, maxRow); _vbar.LargeChange = visRows;
            _vbar.Value = Math.Clamp(s.ScrollRow, _vbar.Minimum, _vbar.Maximum);
            _hbar.Minimum = s.FreezeCols; _hbar.Maximum = Math.Max(s.FreezeCols, maxCol); _hbar.LargeChange = visCols;
            _hbar.Value = Math.Clamp(s.ScrollCol, _hbar.Minimum, _hbar.Maximum);
        }
        finally { _syncingBars = false; }
    }

    public void ScrollTo(int row, int col)
    {
        var s = Sheet;
        s.ScrollRow = Math.Clamp(row, s.FreezeRows, Addr.MaxRows - 1);
        s.ScrollCol = Math.Clamp(col, s.FreezeCols, Addr.MaxCols - 1);
        ComputeLayout();
        UpdateScrollBars();
        PositionEditor();
        Invalidate();
    }

    public void RefreshAll()
    {
        if (_doc == null) return;
        ComputeLayout();
        UpdateScrollBars();
        PositionEditor();
        Invalidate();
    }

    /// <summary>Screen rectangle of a row/col if visible (merged cells expand to the whole merge).</summary>
    public Rectangle? CellRect(int r, int c, bool expandMerge = true)
    {
        if (expandMerge && Sheet.MergeAt(r, c) is CellRange m) return RangeRect(m);
        int ri = _rows.FindIndex(x => x.Index == r), ci = _cols.FindIndex(x => x.Index == c);
        if (ri < 0 || ci < 0) return null;
        return new Rectangle(_cols[ci].Pos, _rows[ri].Pos, _cols[ci].Size, _rows[ri].Size);
    }

    /// <summary>Visible part of a range (clipped to the grid), or null.</summary>
    public Rectangle? RangeRect(CellRange rg)
    {
        rg = rg.Normalized;
        int x1 = int.MaxValue, y1 = int.MaxValue, x2 = int.MinValue, y2 = int.MinValue;
        foreach (var c in _cols) if (c.Index >= rg.C1 && c.Index <= rg.C2) { x1 = Math.Min(x1, c.Pos); x2 = Math.Max(x2, c.Pos + c.Size); }
        foreach (var r in _rows) if (r.Index >= rg.R1 && r.Index <= rg.R2) { y1 = Math.Min(y1, r.Pos); y2 = Math.Max(y2, r.Pos + r.Size); }
        if (x1 == int.MaxValue || y1 == int.MaxValue) return null;
        // ranges that start above / left of the scrolled area extend beyond the frozen boundary
        return new Rectangle(x1, y1, x2 - x1, y2 - y1);
    }

    /// <summary>Unscaled absolute x of a column's left edge (sheet pixels at 100%).</summary>
    public double SheetX(int col) { double x = 0; for (int c = 0; c < col; c++) x += Sheet.ColWidth(c); return x; }
    public double SheetY(int row) { double y = 0; for (int r = 0; r < row; r++) y += Sheet.RowHeight(r); return y; }

    /// <summary>Converts sheet pixel coordinates (100%) to screen coordinates.</summary>
    public RectangleF SheetToScreen(RectangleF b)
    {
        var s = Sheet;
        double originX = SheetX(s.ScrollCol) - SheetX(s.FreezeCols), originY = SheetY(s.ScrollRow) - SheetY(s.FreezeRows);
        float sc = Scale;
        return new RectangleF(RowHeaderW + (float)((b.X - originX) * sc), ColHeaderH + (float)((b.Y - originY) * sc), b.Width * sc, b.Height * sc);
    }

    public PointF ScreenToSheet(Point p)
    {
        var s = Sheet;
        double originX = SheetX(s.ScrollCol) - SheetX(s.FreezeCols), originY = SheetY(s.ScrollRow) - SheetY(s.FreezeRows);
        float sc = Scale;
        return new PointF((float)((p.X - RowHeaderW) / sc + originX), (float)((p.Y - ColHeaderH) / sc + originY));
    }

    // ------------------------------------------------------------------ hit testing

    public Hit HitTest(Point p)
    {
        if (_doc == null) return new Hit(HitArea.None, -1, -1);
        if (p.X >= Width - BarW || p.Y >= Height - BarH) return new Hit(HitArea.None, -1, -1);
        // objects first (topmost last drawn)
        var objHit = HitObject(p);
        if (objHit != null) return objHit.Value;
        int hb = Dpi.Scale(3);
        if (p.Y < ColHeaderH && p.X < RowHeaderW) return new Hit(HitArea.Corner, -1, -1);
        if (p.Y < ColHeaderH)
        {
            foreach (var c in _cols)
            {
                if (Math.Abs(p.X - (c.Pos + c.Size)) <= hb) return new Hit(HitArea.ColBorder, -1, c.Index);
                if (p.X >= c.Pos && p.X < c.Pos + c.Size) return new Hit(HitArea.ColHeader, -1, c.Index);
            }
            return new Hit(HitArea.None, -1, -1);
        }
        if (p.X < RowHeaderW)
        {
            foreach (var r in _rows)
            {
                if (Math.Abs(p.Y - (r.Pos + r.Size)) <= hb) return new Hit(HitArea.RowBorder, r.Index, -1);
                if (p.Y >= r.Pos && p.Y < r.Pos + r.Size) return new Hit(HitArea.RowHeader, r.Index, -1);
            }
            return new Hit(HitArea.None, -1, -1);
        }
        int row = RowAtY(p.Y), col = ColAtX(p.X);
        if (row < 0 || col < 0) return new Hit(HitArea.None, row, col);
        // fill handle
        if (!IsEditing && FillHandleRect() is Rectangle fh && Rectangle.Inflate(fh, Dpi.Scale(3), Dpi.Scale(3)).Contains(p)) return new Hit(HitArea.FillHandle, row, col);
        // filter buttons
        if (Sheet.Filter is AutoFilter f && row == f.Range.R1 && col >= f.Range.C1 && col <= f.Range.C2 && FilterButtonRect(row, col) is Rectangle fb && fb.Contains(p))
            return new Hit(HitArea.FilterButton, row, col);
        if (ValidationArrowRect() is Rectangle va && va.Contains(p)) return new Hit(HitArea.ValidationArrow, Active.Row, Active.Col);
        return new Hit(HitArea.Cell, row, col);
    }

    public int RowAtY(int y)
    {
        foreach (var r in _rows) if (y >= r.Pos && y < r.Pos + r.Size) return r.Index;
        if (_rows.Count > 0 && y >= _rows[^1].Pos + _rows[^1].Size) return Math.Min(Addr.MaxRows - 1, _rows[^1].Index + 1);
        if (y < ColHeaderH && _rows.Count > 0) return Math.Max(0, _rows.First(r => r.Index >= Sheet.FreezeRows).Index - 1);
        return _rows.Count > 0 ? _rows[0].Index : -1;
    }

    public int ColAtX(int x)
    {
        foreach (var c in _cols) if (x >= c.Pos && x < c.Pos + c.Size) return c.Index;
        if (_cols.Count > 0 && x >= _cols[^1].Pos + _cols[^1].Size) return Math.Min(Addr.MaxCols - 1, _cols[^1].Index + 1);
        if (x < RowHeaderW && _cols.Count > 0) return Math.Max(0, _cols.First(c => c.Index >= Sheet.FreezeCols).Index - 1);
        return _cols.Count > 0 ? _cols[0].Index : -1;
    }

    internal Rectangle? FillHandleRect()
    {
        if (Sheet.Protected) return null;
        var rr = RangeRect(Selection);
        if (rr == null) return null;
        int s = Dpi.Scale(6);
        var r = rr.Value;
        // only if the bottom-right corner is visible
        if (!_rows.Any(x => x.Index == Selection.R2) || !_cols.Any(x => x.Index == Selection.C2)) return null;
        return new Rectangle(r.Right - s / 2 - 1, r.Bottom - s / 2 - 1, s, s);
    }

    internal Rectangle? FilterButtonRect(int r, int c)
    {
        var cr = CellRect(r, c, false);
        if (cr == null) return null;
        int s = Math.Min(cr.Value.Height - 2, (int)Math.Round(16 * Scale));
        return new Rectangle(cr.Value.Right - s - 2, cr.Value.Bottom - s - 1, s, s);
    }

    internal Rectangle? ValidationArrowRect()
    {
        if (IsEditing || ValidationAt(Active.Row, Active.Col) is not { Type: ValidationType.List }) return null;
        var cr = CellRect(Active.Row, Active.Col);
        if (cr == null) return null;
        int s = Math.Max(Dpi.Scale(14), cr.Value.Height);
        return new Rectangle(cr.Value.Right + 1, cr.Value.Top, Dpi.Scale(16), s);
    }

    public DataValidation? ValidationAt(int r, int c) => Sheet.Validations.LastOrDefault(v => v.Range.Contains(r, c));

    // ------------------------------------------------------------------ selection

    public void Select(int row, int col, bool extend = false, bool add = false)
    {
        row = Math.Clamp(row, 0, Addr.MaxRows - 1);
        col = Math.Clamp(col, 0, Addr.MaxCols - 1);
        var s = Sheet;
        if (add)
        {
            ExtraRanges.Add(s.Selection.Normalized);
            s.ActiveCell = new CellPos(row, col);
            s.Selection = ExpandMerges(CellRange.Single(row, col));
        }
        else if (extend)
        {
            var anchor = s.ActiveCell;
            s.Selection = ExpandMerges(new CellRange(anchor.Row, anchor.Col, row, col).Normalized);
        }
        else
        {
            ExtraRanges.Clear();
            s.ActiveCell = new CellPos(row, col);
            var m = s.MergeAt(row, col);
            s.Selection = m ?? CellRange.Single(row, col);
            if (m is CellRange mm) s.ActiveCell = new CellPos(mm.R1, mm.C1);
        }
        SelectObject(null);
        EnsureVisible(row, col);
        Invalidate();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SelectRange(CellRange r, CellPos? active = null)
    {
        ExtraRanges.Clear();
        Sheet.Selection = r.Normalized;
        Sheet.ActiveCell = active ?? new CellPos(r.Normalized.R1, r.Normalized.C1);
        SelectObject(null);
        EnsureVisible(Sheet.ActiveCell.Row, Sheet.ActiveCell.Col);
        Invalidate();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    private CellRange ExpandMerges(CellRange r)
    {
        bool grown = true;
        while (grown)
        {
            grown = false;
            foreach (var m in Sheet.Merges)
                if (m.Intersects(r) && !r.Contains(m)) { r = r.Union(m); grown = true; }
        }
        return r;
    }

    public void EnsureVisible(int row, int col)
    {
        var s = Sheet;
        bool changed = false;
        if (row >= s.FreezeRows)
        {
            if (row < s.ScrollRow) { s.ScrollRow = row; changed = true; }
            else
            {
                var vis = _rows.Where(r => r.Index >= s.FreezeRows && r.Pos + r.Size <= Height - BarH).ToList();
                if (vis.Count > 0 && row > vis[^1].Index)
                {
                    // scroll so that row is the last fully visible one
                    int avail = Height - BarH - _frozenBottom, y = 0, r = row;
                    while (r > s.FreezeRows) { int h = (int)Math.Round(s.RowHeight(r) * Scale); if (y + h > avail) break; y += h; r--; }
                    s.ScrollRow = Math.Min(row, r + 1);
                    changed = true;
                }
            }
        }
        if (col >= s.FreezeCols)
        {
            if (col < s.ScrollCol) { s.ScrollCol = col; changed = true; }
            else
            {
                var vis = _cols.Where(c => c.Index >= s.FreezeCols && c.Pos + c.Size <= Width - BarW).ToList();
                if (vis.Count > 0 && col > vis[^1].Index)
                {
                    int avail = Width - BarW - _frozenRight, x = 0, c = col;
                    while (c > s.FreezeCols) { int w = (int)Math.Round(s.ColWidth(c) * Scale); if (x + w > avail) break; x += w; c--; }
                    s.ScrollCol = Math.Min(col, c + 1);
                    changed = true;
                }
            }
        }
        if (changed) { ComputeLayout(); UpdateScrollBars(); PositionEditor(); }
    }

    public void SelectObject(object? obj)
    {
        if (ReferenceEquals(SelectedObject, obj)) return;
        SelectedObject = obj;
        Invalidate();
        ObjectSelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SetZoom(double z)
    {
        Sheet.Zoom = Math.Clamp(z, 0.1, 4.0);
        RefreshAll();
        ZoomChanged?.Invoke(this, (int)Math.Round(Sheet.Zoom * 100));
    }

    internal void RaiseContentChanged() => ContentChanged?.Invoke(this, EventArgs.Empty);
    internal void RaiseSelectionChanged() => SelectionChanged?.Invoke(this, EventArgs.Empty);
}
