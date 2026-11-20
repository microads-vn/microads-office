using System.Drawing;
using MicroadsOffice.Excel.Formula;

namespace MicroadsOffice.Excel.Model;

public sealed class Cell
{
    /// <summary>Typed constant or the last computed result of the formula.</summary>
    public CellValue Value;
    /// <summary>Formula without the leading '=' (null for constants).</summary>
    public string? Formula;
    public int StyleId;
    internal Expr? Ast;
    internal int CalcStamp = -1;

    // ---- dynamic arrays: a formula whose result is an array "spills" it into the cells to the right and below
    /// <summary>For a spilled value: the formula cell it comes from and the offset from that cell.</summary>
    internal Cell? SpillFrom;
    internal int SpillDr, SpillDc;
    /// <summary>For a formula cell: size of the area it currently spills into (0 when it does not spill).</summary>
    internal int SpillRows, SpillCols;
    /// <summary>The array result could not spill (#SPILL!).</summary>
    internal bool SpillBlocked;
    /// <summary>The last result was an array (saved as a dynamic-array formula).</summary>
    internal bool Dynamic;
    /// <summary>Formula from a file written without dynamic-array support: arrays reduce to one value (implicit intersection) like Excel's "@".</summary>
    internal bool Implicit;

    public bool HasFormula => Formula != null;
    /// <summary>The value was spilled by the formula of another cell.</summary>
    public bool IsSpilled => SpillFrom != null;
    public bool IsEmpty => Formula == null && Value.IsBlank && StyleId == 0;
    public Cell Clone() => new() { Value = Value, Formula = Formula, StyleId = StyleId, Ast = Ast, Implicit = Implicit };
}

public enum CfType { GreaterThan, LessThan, Between, EqualTo, TextContains, Duplicate, Top10, Bottom10, AboveAverage, BelowAverage, DataBar, ColorScale2, ColorScale3, Expression }

public sealed class ConditionalFormat
{
    public CellRange Range;
    public CfType Type;
    public string Value1 = "";
    public string Value2 = "";
    public int Rank = 10;
    /// <summary>Format applied when the rule matches.</summary>
    public int FontColor;
    public int Fill;
    public bool Bold;
    /// <summary>Data bar colour / scale colours (min, mid, max).</summary>
    public int Color1 = Color.FromArgb(99, 142, 198).ToArgb();
    public int Color2 = Color.FromArgb(255, 235, 132).ToArgb();
    public int Color3 = Color.FromArgb(99, 190, 123).ToArgb();
    public ConditionalFormat Clone() => (ConditionalFormat)MemberwiseClone();
}

public enum ValidationType { Any, List, WholeNumber, Decimal, Date, TextLength }
public enum ValidationOp { Between, NotBetween, Equal, NotEqual, Greater, Less, GreaterOrEqual, LessOrEqual }

public sealed class DataValidation
{
    public CellRange Range;
    public ValidationType Type;
    public ValidationOp Op;
    /// <summary>For lists: "a,b,c" or a range reference "=$A$1:$A$9".</summary>
    public string Formula1 = "";
    public string Formula2 = "";
    public bool AllowBlank = true;
    public string ErrorMessage = "";
    public DataValidation Clone() => (DataValidation)MemberwiseClone();
}

public enum ChartType { Column, StackedColumn, Bar, Line, Pie, Doughnut, Area, Scatter }

public sealed class Chart
{
    public ChartType Type = ChartType.Column;
    public string Title = "";
    /// <summary>Data range in the same sheet (first row = series names, first column = categories).</summary>
    public CellRange Data;
    public bool SeriesInRows;
    public bool ShowLegend = true;
    public bool ShowDataLabels;
    /// <summary>Position/size in sheet pixels at 100% zoom.</summary>
    public RectangleF Bounds = new(320, 40, 480, 290);
    public string Name = "Chart 1";
    public Chart Clone() => (Chart)MemberwiseClone();
}

public sealed class Picture
{
    public byte[] Data = [];
    public string Ext = "png";
    public RectangleF Bounds;
    public string Name = "Picture 1";
    public Picture Clone() => (Picture)MemberwiseClone();
}

public sealed class PageSetup
{
    public bool Landscape;
    /// <summary>Paper: 9 = A4, 1 = Letter, 5 = Legal, 8 = A3, 11 = A5.</summary>
    public int PaperSize = 9;
    public double MarginLeft = 0.7, MarginRight = 0.7, MarginTop = 0.75, MarginBottom = 0.75; // inches
    public bool FitToWidth;
    public int Scale = 100;
    public bool PrintGridlines;
    public bool PrintHeadings;
    public CellRange? PrintArea;
    /// <summary>Rows repeated at the top of each page (inclusive).</summary>
    public (int R1, int R2)? PrintTitleRows;
    public bool CenterHorizontally;
    public PageSetup Clone() => (PageSetup)MemberwiseClone();
}

public sealed class AutoFilter
{
    public CellRange Range;
    /// <summary>Column index (absolute) → set of allowed display texts (null = no filter on that column).</summary>
    public Dictionary<int, HashSet<string>> Allowed = new();
    /// <summary>Column → custom condition "&gt;100", "contains:abc".</summary>
    public Dictionary<int, string> Conditions = new();
    public AutoFilter Clone() => new() { Range = Range, Allowed = Allowed.ToDictionary(k => k.Key, v => new HashSet<string>(v.Value)), Conditions = new(Conditions) };
}

public sealed class Sheet
{
    public string Name;
    public readonly Workbook Book;
    internal Dictionary<long, Cell> Cells = new();
    public Dictionary<int, double> ColWidths = new();   // pixels at 100%
    public Dictionary<int, double> RowHeights = new();  // pixels at 100% (set by the user)
    /// <summary>Heights grown automatically because a row contains large fonts (not user-set).</summary>
    public Dictionary<int, double> AutoHeights = new();
    public HashSet<int> HiddenCols = new();
    public HashSet<int> HiddenRows = new();
    /// <summary>Rows hidden by the autofilter (kept separately from manually hidden rows).</summary>
    public HashSet<int> FilteredRows = new();
    public Dictionary<int, int> ColStyles = new();
    public Dictionary<int, int> RowStyles = new();
    public List<CellRange> Merges = new();
    public List<ConditionalFormat> CondFormats = new();
    public List<DataValidation> Validations = new();
    public Dictionary<long, string> Notes = new();
    public Dictionary<long, string> Links = new();
    public List<Chart> Charts = new();
    public List<Picture> Pictures = new();
    public PageSetup Page = new();
    public AutoFilter? Filter;
    public int FreezeRows, FreezeCols;
    public bool ShowGridlines = true;
    public bool ShowHeadings = true;
    public int TabColor;
    public bool Protected;
    public string? ProtectionPasswordHash;
    public double Zoom = 1.0;
    public CellPos ActiveCell;
    public CellRange Selection;
    public int ScrollRow, ScrollCol;
    public bool Hidden;

    public const double DefaultColWidth = 64;
    public const double DefaultRowHeight = 20;

    public Sheet(Workbook book, string name) { Book = book; Name = name; }

    public static long Key(int r, int c) => ((long)r << 14) | (uint)c;
    public static (int R, int C) FromKey(long k) => ((int)(k >> 14), (int)(k & 0x3FFF));

    public Cell? GetCell(int r, int c) => Cells.TryGetValue(Key(r, c), out var cell) ? cell : null;

    public Cell GetOrCreate(int r, int c)
    {
        var k = Key(r, c);
        if (!Cells.TryGetValue(k, out var cell))
        {
            cell = new Cell();
            if (RowStyles.TryGetValue(r, out var rs)) cell.StyleId = rs;
            else if (ColStyles.TryGetValue(c, out var cs)) cell.StyleId = cs;
            Cells[k] = cell;
        }
        return cell;
    }

    public IEnumerable<KeyValuePair<long, Cell>> AllCells => Cells;
    public int CellCount => Cells.Count;

    public CellValue GetValue(int r, int c) => GetCell(r, c)?.Value ?? CellValue.Blank;

    public int StyleIdAt(int r, int c)
    {
        var cell = GetCell(r, c);
        if (cell != null) return cell.StyleId;
        if (RowStyles.TryGetValue(r, out var rs)) return rs;
        if (ColStyles.TryGetValue(c, out var cs)) return cs;
        return 0;
    }

    public CellStyle StyleAt(int r, int c) => Book.Styles[StyleIdAt(r, c)];

    /// <summary>Sets a constant (or clears with Blank). Formula is removed.</summary>
    public void SetValue(int r, int c, CellValue v)
    {
        var cell = GetOrCreate(r, c);
        ForgetSpill(r, c, cell);
        cell.Value = v;
        cell.Formula = null;
        cell.Ast = null;
        CleanupIfEmpty(r, c, cell);
        Book.MarkDirty();
    }

    public void SetFormula(int r, int c, string formula)
    {
        var cell = GetOrCreate(r, c);
        ForgetSpill(r, c, cell);
        cell.Implicit = false;
        cell.Formula = formula.StartsWith('=') ? formula[1..] : formula;
        cell.Ast = null;
        cell.Value = CellValue.Blank;
        Book.MarkDirty();
    }

    public void SetStyle(int r, int c, int styleId)
    {
        var cell = GetOrCreate(r, c);
        cell.StyleId = styleId;
        GrowRowFor(r, Book.Styles[styleId]);
        CleanupIfEmpty(r, c, cell);
    }

    /// <summary>Rows without a user height grow to fit large fonts (like Excel's automatic row height).</summary>
    public void GrowRowFor(int r, CellStyle st)
    {
        if (st.FontSize <= 11.5f || RowHeights.ContainsKey(r)) return;
        double h = Math.Ceiling(st.FontSize * 96 / 72 * 1.3 + 4);
        if (h > AutoHeights.GetValueOrDefault(r, DefaultRowHeight)) AutoHeights[r] = h;
    }

    public void ClearCell(int r, int c, bool contents, bool formats)
    {
        var cell = GetCell(r, c);
        if (cell == null) return;
        if (contents) { ForgetSpill(r, c, cell); cell.Value = CellValue.Blank; cell.Formula = null; cell.Ast = null; Book.MarkDirty(); }
        if (formats) cell.StyleId = 0;
        CleanupIfEmpty(r, c, cell);
    }

    private void CleanupIfEmpty(int r, int c, Cell cell)
    {
        if (cell.IsEmpty && !RowStyles.ContainsKey(r) && !ColStyles.ContainsKey(c)) Cells.Remove(Key(r, c));
    }

    internal void RemoveIfEmpty(long key)
    {
        if (!Cells.TryGetValue(key, out var cell)) return;
        var (r, c) = FromKey(key);
        CleanupIfEmpty(r, c, cell);
    }

    /// <summary>Content typed over a cell: it stops being a spilled value, and a formula cell drops what it spilled.</summary>
    internal void ForgetSpill(int r, int c, Cell cell)
    {
        cell.SpillFrom = null;
        cell.Dynamic = false;
        cell.SpillBlocked = false;
        if (cell.SpillRows > 0) Calc.ClearSpill(this, r, c, cell);
    }

    /// <summary>The formula cell whose array result covers (r, c), with its position (null when the cell is not spilled).</summary>
    public (int Row, int Col, Cell Cell)? SpillSource(int r, int c)
    {
        var cell = GetCell(r, c);
        if (cell?.SpillFrom == null) return null;
        int mr = r - cell.SpillDr, mc = c - cell.SpillDc;
        var m = GetCell(mr, mc);
        return m == cell.SpillFrom && m.Formula != null ? (mr, mc, m) : null;
    }

    /// <summary>Area spilled by the formula at (r, c) (null when it does not spill).</summary>
    public CellRange? SpillArea(int r, int c)
    {
        var cell = GetCell(r, c);
        if (cell?.Formula == null || cell.SpillRows <= 0) return null;
        return new CellRange(r, c, r + cell.SpillRows - 1, c + cell.SpillCols - 1);
    }

    public double ColWidth(int c) => HiddenCols.Contains(c) ? 0 : ColWidths.TryGetValue(c, out var w) ? w : DefaultColWidth;
    public double RowHeight(int r) => HiddenRows.Contains(r) || FilteredRows.Contains(r) ? 0 : RawRowHeight(r);
    public double RawColWidth(int c) => ColWidths.TryGetValue(c, out var w) ? w : DefaultColWidth;
    public double RawRowHeight(int r) => RowHeights.TryGetValue(r, out var h) ? h : AutoHeights.TryGetValue(r, out var a) ? a : DefaultRowHeight;
    public bool IsRowHidden(int r) => HiddenRows.Contains(r) || FilteredRows.Contains(r);

    /// <summary>Smallest rectangle containing every non-empty cell (A1 when empty).</summary>
    public CellRange UsedRange()
    {
        if (Cells.Count == 0) return new CellRange(0, 0, 0, 0);
        int r1 = int.MaxValue, c1 = int.MaxValue, r2 = 0, c2 = 0;
        foreach (var k in Cells.Keys)
        {
            var (r, c) = FromKey(k);
            if (r < r1) r1 = r; if (r > r2) r2 = r; if (c < c1) c1 = c; if (c > c2) c2 = c;
        }
        foreach (var m in Merges) { r2 = Math.Max(r2, m.R2); c2 = Math.Max(c2, m.C2); }
        return new CellRange(r1, c1, r2, c2);
    }

    /// <summary>Used range that only counts cells with content (not formats).</summary>
    public CellRange DataRange()
    {
        int r2 = -1, c2 = -1, r1 = int.MaxValue, c1 = int.MaxValue;
        foreach (var (k, cell) in Cells)
        {
            if (cell.Value.IsBlank && cell.Formula == null) continue;
            var (r, c) = FromKey(k);
            if (r < r1) r1 = r; if (r > r2) r2 = r; if (c < c1) c1 = c; if (c > c2) c2 = c;
        }
        return r2 < 0 ? new CellRange(0, 0, 0, 0) : new CellRange(r1, c1, r2, c2);
    }

    public CellRange? MergeAt(int r, int c)
    {
        foreach (var m in Merges) if (m.Contains(r, c)) return m;
        return null;
    }

    public long PosKey(int r, int c) => Key(r, c);

    public Sheet Clone(Workbook book)
    {
        var s = new Sheet(book, Name)
        {
            ColWidths = new(ColWidths), RowHeights = new(RowHeights), AutoHeights = new(AutoHeights), HiddenCols = new(HiddenCols), HiddenRows = new(HiddenRows), FilteredRows = new(FilteredRows),
            ColStyles = new(ColStyles), RowStyles = new(RowStyles), Merges = new(Merges), CondFormats = CondFormats.Select(x => x.Clone()).ToList(),
            Validations = Validations.Select(x => x.Clone()).ToList(), Notes = new(Notes), Links = new(Links), Charts = Charts.Select(x => x.Clone()).ToList(),
            Pictures = Pictures.Select(x => x.Clone()).ToList(), Page = Page.Clone(), Filter = Filter?.Clone(), FreezeRows = FreezeRows, FreezeCols = FreezeCols,
            ShowGridlines = ShowGridlines, ShowHeadings = ShowHeadings, TabColor = TabColor, Protected = Protected, ProtectionPasswordHash = ProtectionPasswordHash,
            Zoom = Zoom, ActiveCell = ActiveCell, Selection = Selection, ScrollRow = ScrollRow, ScrollCol = ScrollCol, Hidden = Hidden
        };
        s.Cells = new Dictionary<long, Cell>(Cells.Count);
        Dictionary<Cell, Cell>? map = null;
        foreach (var (k, c) in Cells)
        {
            var n = c.Clone();
            n.SpillRows = c.SpillRows; n.SpillCols = c.SpillCols; n.SpillBlocked = c.SpillBlocked; n.Dynamic = c.Dynamic; n.Implicit = c.Implicit;
            n.SpillDr = c.SpillDr; n.SpillDc = c.SpillDc;
            if (c.SpillRows > 0 || c.SpillFrom != null) (map ??= new())[c] = n;
            s.Cells[k] = n;
        }
        if (map != null)
            foreach (var (old, n) in map)
                if (old.SpillFrom != null && map.TryGetValue(old.SpillFrom, out var m)) n.SpillFrom = m;
        return s;
    }
}

public sealed class Workbook
{
    public List<Sheet> Sheets { get; } = new();
    public StyleTable Styles { get; private set; } = new();
    /// <summary>Workbook-level defined names → reference/formula text (without '=').</summary>
    public Dictionary<string, string> Names { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
    public int ActiveSheet;
    public bool ManualCalc;
    public string Title = "", Author = "";
    internal int Version;
    public Calc Calc { get; }

    public Workbook() { Calc = new Calc(this); }

    public static Workbook CreateBlank(int sheets = 1)
    {
        var wb = new Workbook();
        for (int i = 0; i < sheets; i++) wb.AddSheet();
        return wb;
    }

    public Sheet AddSheet(string? name = null, int index = -1)
    {
        name ??= NextSheetName();
        var s = new Sheet(this, name);
        if (index < 0 || index > Sheets.Count) Sheets.Add(s); else Sheets.Insert(index, s);
        MarkDirty();
        return s;
    }

    public string NextSheetName()
    {
        for (int i = Sheets.Count + 1; ; i++)
        {
            var n = "Sheet" + i;
            if (Sheet(n) == null) return n;
        }
    }

    public Sheet? Sheet(string name) => Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
    public Sheet Active => Sheets[Math.Clamp(ActiveSheet, 0, Sheets.Count - 1)];

    public void MarkDirty() => Version++;

    public void Recalculate(bool force = false) => Calc.RecalculateAll(force);

    public Workbook Clone()
    {
        var wb = new Workbook { ActiveSheet = ActiveSheet, ManualCalc = ManualCalc, Title = Title, Author = Author };
        wb.Styles = Styles.Clone();
        wb.Names = new Dictionary<string, string>(Names, StringComparer.OrdinalIgnoreCase);
        foreach (var s in Sheets) wb.Sheets.Add(s.Clone(wb));
        wb.Version = Version;
        return wb;
    }

    /// <summary>Applies a change to the style of a cell and returns the new style id.</summary>
    public int Restyle(int styleId, Func<CellStyle, CellStyle> change) => Styles.Intern(change(Styles[styleId]));

    public static bool IsValidSheetName(string n)
        => !string.IsNullOrWhiteSpace(n) && n.Length <= 31 && n.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) < 0 && !n.StartsWith('\'') && !n.EndsWith('\'');
}
