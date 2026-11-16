using System.Globalization;
using System.Text.RegularExpressions;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Format;

/// <summary>Interprets what the user typed into a cell: numbers (display locale), %, currency, dates, times, booleans.</summary>
public static partial class ValueParser
{
    /// <summary>Returns the typed value and, for dates/percent/currency, a format to apply when the cell is still "General".</summary>
    public static CellValue Parse(string text, out string? suggestedFormat)
    {
        suggestedFormat = null;
        if (text.Length == 0) return CellValue.Blank;
        if (text.StartsWith('\'')) return CellValue.Text(text[1..]);
        var t = text.Trim();
        if (t.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return CellValue.Bool(true);
        if (t.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return CellValue.Bool(false);
        if (Errors.IsErrorCode(t)) return CellValue.Error(Errors.All.First(e => e.Equals(t, StringComparison.OrdinalIgnoreCase)));
        if (TryNumber(t, out double d, out var fmt)) { suggestedFormat = fmt; return CellValue.Number(d); }
        if (TryDateTime(t, out d, out fmt)) { suggestedFormat = fmt; return CellValue.Number(d); }
        return CellValue.Text(text);
    }

    /// <summary>Number in the display locale ("1.234,5" in vi, "1,234.5" in en), with optional %, ₫/$/đ, parentheses for negatives.</summary>
    public static bool TryNumber(string t, out double d, out string? fmt)
    {
        d = 0; fmt = null;
        t = t.Trim();
        if (t.Length == 0) return false;
        bool neg = false, pct = false;
        string? currency = null;
        if (t.StartsWith('(') && t.EndsWith(')')) { neg = true; t = t[1..^1].Trim(); }
        if (t.StartsWith('-')) { neg = !neg; t = t[1..].Trim(); }
        else if (t.StartsWith('+')) t = t[1..].Trim();
        if (t.EndsWith('%')) { pct = true; t = t[..^1].Trim(); }
        foreach (var sym in new[] { "₫", "VND", "VNĐ", "đ", "$" })
        {
            if (t.EndsWith(sym, StringComparison.OrdinalIgnoreCase)) { currency = sym; t = t[..^sym.Length].Trim(); break; }
            if (t.StartsWith(sym, StringComparison.OrdinalIgnoreCase)) { currency = sym; t = t[sym.Length..].Trim(); break; }
        }
        if (t.StartsWith('-')) { neg = !neg; t = t[1..].Trim(); }
        if (t.Length == 0) return false;
        char dec = NumFmt.DecimalSep, grp = NumFmt.GroupSep;
        bool hasGroup = false;
        // In Vietnamese display mode "1.5" is still understood as one and a half (only "1.500" is a thousands group).
        if (dec == ',' && !t.Contains(',') && t.Count(c => c == '.') == 1 && t.Length - t.IndexOf('.') - 1 != 3 && t.All(c => char.IsAsciiDigit(c) || c == '.'))
        { dec = '.'; grp = ','; }
        // validate characters
        foreach (var ch in t) if (!(char.IsAsciiDigit(ch) || ch == dec || ch == grp || ch == 'e' || ch == 'E' || ch == '+' || ch == '-')) return false;
        string norm;
        if (t.Contains('e') || t.Contains('E'))
        {
            norm = t.Replace(dec, '.');
            if (dec != '.' && t.Contains('.') && grp == '.') return false;
        }
        else
        {
            int decCount = t.Count(c => c == dec);
            if (decCount > 1) return false;
            var parts = t.Split(dec);
            var intPart = parts[0];
            if (intPart.Contains(grp))
            {
                var groups = intPart.Split(grp);
                if (groups[0].Length == 0 || groups[0].Length > 3 || groups.Skip(1).Any(g => g.Length != 3)) return false;
                hasGroup = true;
                intPart = intPart.Replace(grp.ToString(), "");
            }
            if (parts.Length > 1 && parts[1].Contains(grp)) return false;
            norm = intPart + (parts.Length > 1 ? "." + parts[1] : "");
            if (norm.StartsWith('.')) norm = "0" + norm;
        }
        if (!double.TryParse(norm, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return false;
        if (neg) d = -d;
        if (pct) { d /= 100; fmt = norm.Contains('.') ? "0.00%" : "0%"; }
        else if (currency != null)
        {
            bool vnd = currency != "$";
            fmt = vnd ? "#,##0 \"₫\"" : "\"$\"#,##0.00";
        }
        else if (hasGroup) fmt = norm.Contains('.') ? "#,##0.00" : "#,##0";
        return true;
    }

    [GeneratedRegex(@"^(\d{1,4})([/\-\.])(\d{1,2})(?:\2(\d{1,4}))?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|SA|CH)?)?$", RegexOptions.IgnoreCase)]
    private static partial Regex DateRx();

    [GeneratedRegex(@"^(\d{1,2}):(\d{2})(?::(\d{2})(?:[\.,](\d+))?)?\s*(AM|PM|SA|CH)?$", RegexOptions.IgnoreCase)]
    private static partial Regex TimeRx();

    public static bool TryDateTime(string t, out double serial, out string? fmt)
    {
        serial = 0; fmt = null;
        var tm = TimeRx().Match(t);
        if (tm.Success)
        {
            int h = int.Parse(tm.Groups[1].Value), m = int.Parse(tm.Groups[2].Value), s = tm.Groups[3].Success ? int.Parse(tm.Groups[3].Value) : 0;
            var ap = tm.Groups[5].Value.ToUpperInvariant();
            if (ap is "PM" or "CH" && h < 12) h += 12;
            if (ap is "AM" or "SA" && h == 12) h = 0;
            if (m > 59 || s > 59) return false;
            serial = (h * 3600 + m * 60 + s) / 86400.0;
            fmt = tm.Groups[3].Success ? "h:mm:ss" : "h:mm";
            if (ap.Length > 0) fmt += " AM/PM";
            return true;
        }
        var dm = DateRx().Match(t);
        if (!dm.Success) return false;
        int a = int.Parse(dm.Groups[1].Value), b = int.Parse(dm.Groups[3].Value);
        int y, mo, d;
        if (dm.Groups[1].Value.Length == 4) { y = a; mo = b; d = dm.Groups[4].Success ? int.Parse(dm.Groups[4].Value) : 1; }
        else
        {
            if (!dm.Groups[4].Success)
            {
                // "15/3" → current year
                y = DateTime.Today.Year;
                if (NumFmt.DayFirst) { d = a; mo = b; } else { mo = a; d = b; }
            }
            else
            {
                y = int.Parse(dm.Groups[4].Value);
                if (dm.Groups[4].Value.Length <= 2) y += y < 30 ? 2000 : 1900;
                if (NumFmt.DayFirst) { d = a; mo = b; } else { mo = a; d = b; }
            }
        }
        if (mo < 1 || mo > 12 || d < 1 || d > DateTime.DaysInMonth(Math.Clamp(y, 1, 9999), mo) || y < 1900 || y > 9999) return false;
        serial = DateSerial.ToSerial(new DateTime(y, mo, d));
        fmt = NumFmt.ShortDateCode;
        if (dm.Groups[5].Success)
        {
            int h = int.Parse(dm.Groups[5].Value), m = int.Parse(dm.Groups[6].Value), s = dm.Groups[7].Success ? int.Parse(dm.Groups[7].Value) : 0;
            var ap = dm.Groups[8].Value.ToUpperInvariant();
            if (ap is "PM" or "CH" && h < 12) h += 12;
            if (ap is "AM" or "SA" && h == 12) h = 0;
            serial += (h * 3600 + m * 60 + s) / 86400.0;
            fmt += " h:mm";
        }
        return true;
    }

    /// <summary>Text → number coercion used inside formulas (VALUE, arithmetic on text): accepts invariant and display formats.</summary>
    public static bool TryCoerceNumber(string s, out double d)
    {
        s = s.Trim();
        d = 0;
        if (s.Length == 0) return false;
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return true;
        if (TryNumber(s, out d, out _)) return true;
        if (TryDateTime(s, out d, out _)) return true;
        if (s.EndsWith('%') && double.TryParse(s[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out d)) { d /= 100; return true; }
        if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out d)) return true;
        return false;
    }
}
