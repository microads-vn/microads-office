using System.Globalization;
using System.Text;

namespace MicroadsOffice.Excel.Model;

public enum ValueKind : byte { Blank, Number, Text, Bool, Error }

/// <summary>Immutable value of a cell or of a formula result.</summary>
public readonly struct CellValue : IEquatable<CellValue>
{
    public readonly ValueKind Kind;
    public readonly double Num;
    public readonly string? Str;

    private CellValue(ValueKind k, double n, string? s) { Kind = k; Num = n; Str = s; }

    public static readonly CellValue Blank = default;
    public static CellValue Number(double d) => double.IsNaN(d) || double.IsInfinity(d) ? Error(Errors.Num) : new(ValueKind.Number, d, null);
    public static CellValue Text(string s) => new(ValueKind.Text, 0, s ?? "");
    public static CellValue Bool(bool b) => new(ValueKind.Bool, b ? 1 : 0, null);
    public static CellValue Error(string code) => new(ValueKind.Error, 0, code);

    public bool IsBlank => Kind == ValueKind.Blank;
    public bool IsNumber => Kind == ValueKind.Number;
    public bool IsText => Kind == ValueKind.Text;
    public bool IsBool => Kind == ValueKind.Bool;
    public bool IsError => Kind == ValueKind.Error;
    public bool BoolValue => Num != 0;

    public bool Equals(CellValue o) => Kind == o.Kind && (Kind switch
    {
        ValueKind.Number or ValueKind.Bool => Num == o.Num,
        ValueKind.Text or ValueKind.Error => Str == o.Str,
        _ => true
    });
    public override bool Equals(object? obj) => obj is CellValue v && Equals(v);
    public override int GetHashCode() => HashCode.Combine(Kind, Num, Str);
    public static bool operator ==(CellValue a, CellValue b) => a.Equals(b);
    public static bool operator !=(CellValue a, CellValue b) => !a.Equals(b);

    /// <summary>Invariant raw text (used for "show formulas", CSV, tests).</summary>
    public override string ToString() => Kind switch
    {
        ValueKind.Number => FormatGeneralInvariant(Num),
        ValueKind.Text or ValueKind.Error => Str!,
        ValueKind.Bool => Num != 0 ? "TRUE" : "FALSE",
        _ => ""
    };

    public static string FormatGeneralInvariant(double d)
    {
        if (d == Math.Floor(d) && Math.Abs(d) < 1e15) return d.ToString("0", CultureInfo.InvariantCulture);
        var s = d.ToString("G15", CultureInfo.InvariantCulture);
        return s;
    }
}

public static class Errors
{
    public const string Div0 = "#DIV/0!", NA = "#N/A", Name = "#NAME?", Ref = "#REF!", Value = "#VALUE!", Num = "#NUM!", Null = "#NULL!", Calc = "#CALC!", Spill = "#SPILL!";
    public static readonly string[] All = [Div0, NA, Name, Ref, Value, Num, Null, Calc, Spill];
    public static bool IsErrorCode(string s) => All.Contains(s, StringComparer.OrdinalIgnoreCase);
}

/// <summary>Zero-based cell coordinates.</summary>
public readonly record struct CellPos(int Row, int Col)
{
    public override string ToString() => Addr.ColName(Col) + (Row + 1);
}

/// <summary>Inclusive zero-based rectangle of cells.</summary>
public readonly record struct CellRange(int R1, int C1, int R2, int C2)
{
    public static CellRange Single(int r, int c) => new(r, c, r, c);
    public CellRange Normalized => new(Math.Min(R1, R2), Math.Min(C1, C2), Math.Max(R1, R2), Math.Max(C1, C2));
    public int Rows => R2 - R1 + 1;
    public int Cols => C2 - C1 + 1;
    public long Count => (long)Rows * Cols;
    public bool Contains(int r, int c) => r >= R1 && r <= R2 && c >= C1 && c <= C2;
    public bool Contains(CellRange o) => o.R1 >= R1 && o.R2 <= R2 && o.C1 >= C1 && o.C2 <= C2;
    public bool Intersects(CellRange o) => !(o.R2 < R1 || o.R1 > R2 || o.C2 < C1 || o.C1 > C2);
    public bool IsSingle => R1 == R2 && C1 == C2;
    public CellRange Union(CellRange o) => new(Math.Min(R1, o.R1), Math.Min(C1, o.C1), Math.Max(R2, o.R2), Math.Max(C2, o.C2));
    public override string ToString() => IsSingle ? Addr.ToA1(R1, C1) : Addr.ToA1(R1, C1) + ":" + Addr.ToA1(R2, C2);
    public IEnumerable<(int R, int C)> Cells() { for (int r = R1; r <= R2; r++) for (int c = C1; c <= C2; c++) yield return (r, c); }
}

public static class Addr
{
    public const int MaxRows = 1_048_576;
    public const int MaxCols = 16_384;

    public static string ColName(int col)
    {
        var sb = new StringBuilder();
        col++;
        while (col > 0) { int m = (col - 1) % 26; sb.Insert(0, (char)('A' + m)); col = (col - 1) / 26; }
        return sb.ToString();
    }

    public static int ColIndex(string letters)
    {
        int c = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z') return -1;
            c = c * 26 + (ch - 'A' + 1);
            if (c > MaxCols) return -1;
        }
        return c - 1;
    }

    public static string ToA1(int r, int c, bool absR = false, bool absC = false) => (absC ? "$" : "") + ColName(c) + (absR ? "$" : "") + (r + 1);

    /// <summary>Parses "B7", "$B$7" into a position.</summary>
    public static bool TryParseA1(string s, out CellPos p) => TryParseA1(s, out p, out _, out _);

    public static bool TryParseA1(string s, out CellPos p, out bool absR, out bool absC)
    {
        p = default; absR = absC = false;
        s = s.Trim();
        int i = 0;
        if (i < s.Length && s[i] == '$') { absC = true; i++; }
        int start = i;
        while (i < s.Length && char.IsAsciiLetter(s[i])) i++;
        if (i == start || i - start > 3) return false;
        int col = ColIndex(s[start..i]);
        if (i < s.Length && s[i] == '$') { absR = true; i++; }
        int ds = i;
        while (i < s.Length && char.IsAsciiDigit(s[i])) i++;
        if (i == ds || i != s.Length) return false;
        if (!int.TryParse(s[ds..i], out int row) || row < 1 || row > MaxRows || col < 0) return false;
        p = new CellPos(row - 1, col);
        return true;
    }

    /// <summary>Parses "A1", "A1:C9", "A:A", "3:5" (no sheet prefix).</summary>
    public static bool TryParseRange(string s, out CellRange r)
    {
        r = default;
        s = s.Replace("$", "").Trim();
        var parts = s.Split(':');
        if (parts.Length == 1)
        {
            if (!TryParseA1(parts[0], out var p)) return false;
            r = CellRange.Single(p.Row, p.Col);
            return true;
        }
        if (parts.Length != 2) return false;
        if (TryParseA1(parts[0], out var a) && TryParseA1(parts[1], out var b)) { r = new CellRange(a.Row, a.Col, b.Row, b.Col).Normalized; return true; }
        int c1 = ColIndex(parts[0]), c2 = ColIndex(parts[1]);
        if (parts[0].All(char.IsAsciiLetter) && parts[1].All(char.IsAsciiLetter) && c1 >= 0 && c2 >= 0) { r = new CellRange(0, Math.Min(c1, c2), MaxRows - 1, Math.Max(c1, c2)); return true; }
        if (int.TryParse(parts[0], out int r1) && int.TryParse(parts[1], out int r2) && r1 >= 1 && r2 >= 1 && r1 <= MaxRows && r2 <= MaxRows)
        { r = new CellRange(Math.Min(r1, r2) - 1, 0, Math.Max(r1, r2) - 1, MaxCols - 1); return true; }
        return false;
    }

    public static string RangeToString(CellRange r)
    {
        if (r.R1 == 0 && r.R2 == MaxRows - 1) return ColName(r.C1) + ":" + ColName(r.C2);
        if (r.C1 == 0 && r.C2 == MaxCols - 1) return (r.R1 + 1) + ":" + (r.R2 + 1);
        return r.ToString();
    }

    public static string QuoteSheet(string name)
    {
        bool plain = name.Length > 0 && (char.IsLetter(name[0]) || name[0] == '_') && name.All(ch => char.IsLetterOrDigit(ch) || ch == '_' || ch == '.')
                     && !TryParseA1(name, out _);
        return plain ? name : "'" + name.Replace("'", "''") + "'";
    }
}
