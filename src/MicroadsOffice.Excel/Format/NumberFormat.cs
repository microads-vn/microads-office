using System.Drawing;
using System.Globalization;
using System.Text;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Format;

/// <summary>Excel date serials (1900 date system incl. the 1900-02-29 quirk).</summary>
public static class DateSerial
{
    private static readonly DateTime Epoch = new(1899, 12, 30);

    public static double ToSerial(DateTime d)
    {
        var s = (d - Epoch).TotalDays;
        if (s < 61) s -= 1; // dates before 1900-03-01
        return s;
    }

    public static DateTime FromSerial(double s)
    {
        if (s < 61) s += 1;
        if (s < 0) s = 0;
        if (s > 2958465) s = 2958465;
        return Epoch.AddDays(Math.Floor(s)).AddTicks((long)Math.Round((s - Math.Floor(s)) * TimeSpan.TicksPerDay / 10000.0) * 10000);
    }

    public static double FromParts(int y, int m, int d)
    {
        if (y < 1900) y += 1900;
        var dt = new DateTime(Math.Clamp(y, 1900, 9999), 1, 1).AddMonths(m - 1).AddDays(d - 1);
        if (y == 1900 && m == 2 && d == 29) return 60;
        return ToSerial(dt);
    }
}

/// <summary>Formats values with Excel number format codes; display separators follow the chosen locale.</summary>
public static class NumFmt
{
    public static char DecimalSep { get; private set; } = ',';
    public static char GroupSep { get; private set; } = '.';
    public static bool DayFirst { get; private set; } = true;
    public static bool Vietnamese { get; private set; } = true;

    public static void UseLocale(bool vietnameseStyle)
    {
        Vietnamese = vietnameseStyle;
        DecimalSep = vietnameseStyle ? ',' : '.';
        GroupSep = vietnameseStyle ? '.' : ',';
        DayFirst = vietnameseStyle;
    }

    public static string ShortDateCode => DayFirst ? "dd/mm/yyyy" : "m/d/yyyy";

    /// <summary>Built-in number format ids used by .xlsx files.</summary>
    public static string? BuiltIn(int id) => id switch
    {
        0 => "General", 1 => "0", 2 => "0.00", 3 => "#,##0", 4 => "#,##0.00", 9 => "0%", 10 => "0.00%", 11 => "0.00E+00", 12 => "# ?/?", 13 => "# ??/??",
        14 => "m/d/yyyy", 15 => "d-mmm-yy", 16 => "d-mmm", 17 => "mmm-yy", 18 => "h:mm AM/PM", 19 => "h:mm:ss AM/PM", 20 => "h:mm", 21 => "h:mm:ss", 22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)", 38 => "#,##0 ;[Red](#,##0)", 39 => "#,##0.00;(#,##0.00)", 40 => "#,##0.00;[Red](#,##0.00)", 45 => "mm:ss", 46 => "[h]:mm:ss", 47 => "mmss.0", 48 => "##0.0E+0", 49 => "@",
        _ => null
    };

    public static int? BuiltInId(string code)
    {
        for (int i = 0; i <= 49; i++) if (BuiltIn(i) == code) return i;
        return null;
    }

    // ------------------------------------------------------------------ parsed format cache

    private enum TK { Lit, Digit, Dec, Group, Pct, Exp, Slash, Date, AmPm, Text, Elapsed }
    private readonly record struct Tok(TK Kind, string S);

    private sealed class Section
    {
        public readonly List<Tok> Toks = new();
        public Color? Color;
        public string? CondOp;
        public double CondVal;
        public bool IsDate, HasDigits, IsText, Thousands, IsFraction, IsExp;
        public int Percent, Scale;
    }

    private static readonly Dictionary<string, Section[]> Cache = new();

    public static bool IsDateFormat(string code)
    {
        if (string.IsNullOrEmpty(code) || code == "General") return false;
        return Parse(code)[0].IsDate;
    }

    public static bool IsTextFormat(string code) => code == "@";

    private static Section[] Parse(string code)
    {
        lock (Cache)
        {
            if (Cache.TryGetValue(code, out var cached)) return cached;
        }
        var sections = new List<Section>();
        var cur = new Section();
        int i = 0;
        string lower = code.ToLowerInvariant();
        while (i < code.Length)
        {
            char ch = code[i];
            switch (ch)
            {
                case ';': sections.Add(cur); cur = new Section(); i++; continue;
                case '"':
                {
                    int j = code.IndexOf('"', i + 1);
                    if (j < 0) j = code.Length;
                    cur.Toks.Add(new Tok(TK.Lit, code[(i + 1)..j]));
                    i = j + 1; continue;
                }
                case '\\': if (i + 1 < code.Length) cur.Toks.Add(new Tok(TK.Lit, code[i + 1].ToString())); i += 2; continue;
                case '_': cur.Toks.Add(new Tok(TK.Lit, " ")); i += 2; continue;
                case '*': i += 2; continue;
                case '[':
                {
                    int j = code.IndexOf(']', i);
                    if (j < 0) { i++; continue; }
                    var inner = code[(i + 1)..j];
                    var il = inner.ToLowerInvariant();
                    if (il is "h" or "hh" or "m" or "mm" or "s" or "ss") { cur.Toks.Add(new Tok(TK.Elapsed, il)); cur.IsDate = true; }
                    else if (inner.StartsWith('$'))
                    {
                        var sym = inner[1..];
                        int dash = sym.IndexOf('-');
                        if (dash >= 0) sym = sym[..dash];
                        if (sym.Length > 0) cur.Toks.Add(new Tok(TK.Lit, sym));
                    }
                    else if (inner.Length > 0 && (inner[0] is '<' or '>' or '='))
                    {
                        int k = 1;
                        if (k < inner.Length && inner[k] is '=' or '>') k++;
                        cur.CondOp = inner[..k];
                        double.TryParse(inner[k..], NumberStyles.Float, CultureInfo.InvariantCulture, out cur.CondVal);
                    }
                    else
                    {
                        var c = il switch
                        {
                            "red" => Color.Red, "blue" => Color.Blue, "green" => Color.FromArgb(0, 128, 0), "black" => Color.Black, "white" => Color.White,
                            "yellow" => Color.FromArgb(255, 192, 0), "magenta" => Color.Magenta, "cyan" => Color.Cyan, _ => (Color?)null
                        };
                        if (c != null) cur.Color = c;
                    }
                    i = j + 1; continue;
                }
                case '0': case '#': case '?': cur.Toks.Add(new Tok(TK.Digit, ch.ToString())); cur.HasDigits = true; i++; continue;
                case '.':
                    if (cur.IsDate && i + 1 < code.Length && code[i + 1] == '0') { int j = i + 1; while (j < code.Length && code[j] == '0') j++; cur.Toks.Add(new Tok(TK.Date, code[i..j])); i = j; continue; }
                    cur.Toks.Add(new Tok(TK.Dec, ".")); i++; continue;
                case ',': cur.Toks.Add(new Tok(TK.Group, ",")); i++; continue;
                case '%': cur.Toks.Add(new Tok(TK.Pct, "%")); cur.Percent++; i++; continue;
                case '/':
                    if (cur.HasDigits && !cur.IsDate) { cur.Toks.Add(new Tok(TK.Slash, "/")); cur.IsFraction = true; }
                    else cur.Toks.Add(new Tok(TK.Lit, "/"));
                    i++; continue;
                case '@': cur.Toks.Add(new Tok(TK.Text, "@")); cur.IsText = true; i++; continue;
                case 'E': case 'e':
                    if (i + 1 < code.Length && (code[i + 1] == '+' || code[i + 1] == '-') && cur.HasDigits) { cur.Toks.Add(new Tok(TK.Exp, "E" + code[i + 1])); cur.IsExp = true; i += 2; continue; }
                    break;
            }
            if (lower.Length >= i + 5 && lower.Substring(i, 5) == "am/pm") { cur.Toks.Add(new Tok(TK.AmPm, "AM/PM")); cur.IsDate = true; i += 5; continue; }
            if (lower.Length >= i + 3 && lower.Substring(i, 3) == "a/p") { cur.Toks.Add(new Tok(TK.AmPm, "A/P")); cur.IsDate = true; i += 3; continue; }
            char lc = char.ToLowerInvariant(ch);
            if (lc is 'y' or 'm' or 'd' or 'h' or 's')
            {
                int j = i;
                while (j < code.Length && char.ToLowerInvariant(code[j]) == lc) j++;
                cur.Toks.Add(new Tok(TK.Date, code[i..j].ToLowerInvariant()));
                cur.IsDate = true;
                i = j; continue;
            }
            if (code.Length >= i + 7 && code.Substring(i, 7).Equals("General", StringComparison.OrdinalIgnoreCase)) { cur.Toks.Add(new Tok(TK.Lit, "\u0001GENERAL")); i += 7; continue; }
            cur.Toks.Add(new Tok(TK.Lit, ch.ToString()));
            i++;
        }
        sections.Add(cur);
        foreach (var s in sections)
        {
            if (s.IsDate) { ResolveMinutes(s); continue; }
            // thousands separator: a ',' between digit placeholders; trailing ',' after the last digit = scale by 1000
            int lastDigit = s.Toks.FindLastIndex(t => t.Kind == TK.Digit);
            int firstDigit = s.Toks.FindIndex(t => t.Kind == TK.Digit);
            int decIdx = s.Toks.FindIndex(t => t.Kind == TK.Dec);
            for (int k = 0; k < s.Toks.Count; k++)
            {
                if (s.Toks[k].Kind != TK.Group) continue;
                bool betweenInt = k > firstDigit && k < lastDigit && (decIdx < 0 || k < decIdx);
                if (betweenInt) s.Thousands = true;
                else if (k > lastDigit || (decIdx >= 0 && k > decIdx && !s.Toks.Skip(k + 1).Any(t => t.Kind == TK.Digit))) s.Scale++;
                else if (!s.HasDigits) s.Toks[k] = new Tok(TK.Lit, ",");
            }
        }
        var arr = sections.ToArray();
        lock (Cache) { if (Cache.Count > 2000) Cache.Clear(); Cache[code] = arr; }
        return arr;
    }

    private static void ResolveMinutes(Section s)
    {
        for (int k = 0; k < s.Toks.Count; k++)
        {
            var t = s.Toks[k];
            if (t.Kind != TK.Date || t.S[0] != 'm' || t.S.Length > 2) continue;
            bool afterHour = false, beforeSecond = false;
            for (int p = k - 1; p >= 0; p--) { if (s.Toks[p].Kind is TK.Date or TK.Elapsed) { afterHour = s.Toks[p].S[0] == 'h'; break; } }
            for (int p = k + 1; p < s.Toks.Count; p++) { if (s.Toks[p].Kind is TK.Date or TK.Elapsed) { beforeSecond = s.Toks[p].S[0] == 's'; break; } }
            if (afterHour || beforeSecond) s.Toks[k] = new Tok(TK.Date, t.S.Length == 1 ? "n" : "nn");
        }
    }

    // ------------------------------------------------------------------ formatting

    /// <summary>Display text for a cell value.</summary>
    public static string Format(CellValue v, string? code, out Color? color)
    {
        color = null;
        code = string.IsNullOrEmpty(code) ? "General" : code;
        switch (v.Kind)
        {
            case ValueKind.Blank: return "";
            case ValueKind.Error: return v.Str!;
            case ValueKind.Bool: return v.Num != 0 ? "TRUE" : "FALSE";
            case ValueKind.Text:
            {
                if (code == "General" || code == "@") return v.Str!;
                var secs = Parse(code);
                var ts = secs.Length >= 4 ? secs[3] : secs.FirstOrDefault(s => s.IsText);
                if (ts == null) return v.Str!;
                color = ts.Color;
                var sb = new StringBuilder();
                foreach (var t in ts.Toks) sb.Append(t.Kind == TK.Text ? v.Str : t.Kind == TK.Lit ? t.S : "");
                return sb.ToString();
            }
        }
        return FormatNumber(v.Num, code, out color);
    }

    public static string FormatNumber(double d, string code, out Color? color)
    {
        color = null;
        if (code == "General" || string.IsNullOrEmpty(code)) return General(d);
        if (code == "@") return General(d);
        var secs = Parse(code);
        Section sec;
        bool useAbs = false;
        if (secs.Any(s => s.CondOp != null))
        {
            sec = secs.FirstOrDefault(s => s.CondOp != null && Cond(s, d)) ?? secs.FirstOrDefault(s => s.CondOp == null && !s.IsText) ?? secs[^1];
            useAbs = d < 0 && sec != secs[0];
        }
        else if (secs.Length == 1 || (secs.Length == 2 && secs[1].IsText && !secs[1].HasDigits)) sec = secs[0];
        else if (d > 0 || (d == 0 && secs.Length < 3)) sec = secs[0];
        else if (d < 0) { sec = secs[1]; useAbs = true; }
        else sec = secs[2];
        color = sec.Color;
        if (sec.IsText && !sec.HasDigits && !sec.IsDate) return General(d);
        if (sec.IsDate) return FormatDate(d, sec);
        double x = useAbs ? Math.Abs(d) : d;
        for (int i = 0; i < sec.Percent; i++) x *= 100;
        for (int i = 0; i < sec.Scale; i++) x /= 1000;
        if (!sec.HasDigits)
        {
            var sb = new StringBuilder();
            foreach (var t in sec.Toks) sb.Append(t.Kind == TK.Lit ? (t.S == "\u0001GENERAL" ? General(x) : t.S) : t.Kind == TK.Pct ? "%" : "");
            if (x < 0 && !useAbs && !sb.ToString().Contains(General(x))) sb.Insert(0, '-');
            return sb.ToString();
        }
        if (sec.IsExp) return FormatExp(x, sec);
        if (sec.IsFraction) return FormatFraction(x, sec);
        return FormatFixed(x, sec);
    }

    private static bool Cond(Section s, double d) => s.CondOp switch
    {
        ">" => d > s.CondVal, ">=" => d >= s.CondVal, "<" => d < s.CondVal, "<=" => d <= s.CondVal, "=" => d == s.CondVal, "<>" => d != s.CondVal, _ => false
    };

    /// <summary>Excel "General": up to 11 characters, scientific for very large/small numbers.</summary>
    public static string General(double d)
    {
        if (d == 0) return "0";
        double a = Math.Abs(d);
        string s;
        if (a >= 1e11 || a < 1e-9)
        {
            s = d.ToString("0.#####E+00", CultureInfo.InvariantCulture);
        }
        else
        {
            int intDigits = a >= 1 ? (int)Math.Floor(Math.Log10(a)) + 1 : 1;
            int dec = Math.Max(0, 10 - intDigits);
            s = Math.Round(d, dec, MidpointRounding.AwayFromZero).ToString("0." + new string('#', Math.Max(dec, 0)), CultureInfo.InvariantCulture);
            if (s.EndsWith('.')) s = s[..^1];
        }
        return DecimalSep == '.' ? s : s.Replace('.', DecimalSep);
    }

    private static string FormatFixed(double x, Section sec)
    {
        int decIdx = sec.Toks.FindIndex(t => t.Kind == TK.Dec);
        var intToks = decIdx >= 0 ? sec.Toks.Take(decIdx).ToList() : sec.Toks.ToList();
        var decToks = decIdx >= 0 ? sec.Toks.Skip(decIdx + 1).ToList() : new List<Tok>();
        int decPlaces = decToks.Count(t => t.Kind == TK.Digit);
        bool neg = x < 0;
        double r = Math.Round(Math.Abs(x), Math.Min(decPlaces, 15), MidpointRounding.AwayFromZero);
        string digits = r.ToString("F" + decPlaces, CultureInfo.InvariantCulture);
        string intDigits = decPlaces > 0 ? digits[..digits.IndexOf('.')] : digits;
        string fracDigits = decPlaces > 0 ? digits[(digits.IndexOf('.') + 1)..] : "";
        if (intDigits == "0") intDigits = "";
        if (neg && r == 0) neg = false;

        // integer part: map digits right-to-left into placeholders
        var intPh = intToks.Where(t => t.Kind == TK.Digit).ToList();
        var outInt = new StringBuilder();
        int di = intDigits.Length - 1;
        int phCount = intPh.Count;
        var resultTokens = new List<string>();
        int phSeen = 0;
        for (int k = intToks.Count - 1; k >= 0; k--)
        {
            var t = intToks[k];
            if (t.Kind == TK.Digit)
            {
                phSeen++;
                bool leftmost = phSeen == phCount;
                string piece;
                if (leftmost)
                {
                    var rest = di >= 0 ? intDigits[..(di + 1)] : "";
                    di = -1;
                    piece = rest.Length > 0 ? rest : t.S == "0" ? "0" : t.S == "?" ? " " : "";
                }
                else if (di >= 0) piece = intDigits[di--].ToString();
                else piece = t.S == "0" ? "0" : t.S == "?" ? " " : "";
                resultTokens.Add("\u0002" + piece);
            }
            else if (t.Kind == TK.Lit) resultTokens.Add(t.S == "\u0001GENERAL" ? "" : t.S);
            else if (t.Kind == TK.Pct) resultTokens.Add("%");
        }
        resultTokens.Reverse();
        // join digit pieces into one run so grouping can be applied
        var sb = new StringBuilder();
        var run = new StringBuilder();
        void FlushRun()
        {
            if (run.Length == 0) return;
            var s = run.ToString();
            if (sec.Thousands)
            {
                var trimmed = s.TrimStart(' ');
                int pad = s.Length - trimmed.Length;
                sb.Append(' ', pad);
                sb.Append(Group(trimmed));
            }
            else sb.Append(s);
            run.Clear();
        }
        foreach (var piece in resultTokens)
        {
            if (piece.StartsWith('\u0002')) run.Append(piece[1..]);
            else { FlushRun(); sb.Append(piece); }
        }
        FlushRun();
        if (decIdx >= 0)
        {
            bool anyDec = false;
            var decSb = new StringBuilder();
            int fi = 0;
            foreach (var t in decToks)
            {
                if (t.Kind == TK.Digit)
                {
                    char dch = fi < fracDigits.Length ? fracDigits[fi] : '0';
                    fi++;
                    decSb.Append(dch);
                    anyDec = true;
                }
                else if (t.Kind == TK.Lit) decSb.Append('\u0003').Append(t.S).Append('\u0003');
                else if (t.Kind == TK.Pct) decSb.Append('\u0003').Append('%').Append('\u0003');
            }
            // trailing '#' digits are dropped when zero, '?' become spaces
            var decPh = decToks.Where(t => t.Kind == TK.Digit).Select(t => t.S[0]).ToList();
            var plain = decSb.ToString();
            var digitsOnly = new StringBuilder(new string(plain.Where(char.IsAsciiDigit).ToArray()));
            for (int k = digitsOnly.Length - 1; k >= 0; k--)
            {
                if (digitsOnly[k] != '0') break;
                if (decPh[k] == '#') digitsOnly.Length = k;
                else if (decPh[k] == '?') digitsOnly[k] = ' ';
                else break;
            }
            // rebuild with literals
            var final = new StringBuilder();
            int dd = 0;
            bool inLit = false;
            foreach (var ch in plain)
            {
                if (ch == '\u0003') { inLit = !inLit; continue; }
                if (inLit) { final.Append(ch); continue; }
                if (dd < digitsOnly.Length) final.Append(digitsOnly[dd]);
                dd++;
            }
            if (anyDec && digitsOnly.Length > 0) sb.Append(DecimalSep);
            else if (anyDec && decPh.All(c => c == '#')) { /* "0.##" with integer value: Excel keeps the point */ sb.Append(DecimalSep); }
            sb.Append(final);
        }
        var res = sb.ToString();
        if (neg) res = "-" + res;
        return res;
    }

    private static string Group(string digits)
    {
        if (digits.Length <= 3) return digits;
        var sb = new StringBuilder();
        int first = digits.Length % 3;
        if (first > 0) sb.Append(digits, 0, first);
        for (int i = first; i < digits.Length; i += 3)
        {
            if (sb.Length > 0) sb.Append(GroupSep);
            sb.Append(digits, i, 3);
        }
        return sb.ToString();
    }

    private static string FormatExp(double x, Section sec)
    {
        int expIdx = sec.Toks.FindIndex(t => t.Kind == TK.Exp);
        var mant = sec.Toks.Take(expIdx).ToList();
        var expToks = sec.Toks.Skip(expIdx + 1).Where(t => t.Kind == TK.Digit).ToList();
        int decIdx = mant.FindIndex(t => t.Kind == TK.Dec);
        int intPh = (decIdx >= 0 ? mant.Take(decIdx) : mant).Count(t => t.Kind == TK.Digit);
        int decPh = decIdx >= 0 ? mant.Skip(decIdx + 1).Count(t => t.Kind == TK.Digit) : 0;
        if (x == 0) return (0.0).ToString("F" + decPh, CultureInfo.InvariantCulture).Replace('.', DecimalSep) + "E+" + new string('0', Math.Max(2, expToks.Count));
        int step = intPh > 1 && mant.Take(decIdx < 0 ? mant.Count : decIdx).First(t => t.Kind == TK.Digit).S == "#" ? intPh : 1;
        int e = (int)Math.Floor(Math.Log10(Math.Abs(x)));
        e = (int)Math.Floor(e / (double)step) * step;
        double m = x / Math.Pow(10, e);
        m = Math.Round(m, decPh);
        if (Math.Abs(m) >= Math.Pow(10, step)) { m /= Math.Pow(10, step); e += step; }
        var ms = m.ToString("F" + decPh, CultureInfo.InvariantCulture).Replace('.', DecimalSep);
        var sign = e < 0 ? "-" : sec.Toks[expIdx].S == "E+" ? "+" : "";
        return ms + "E" + sign + Math.Abs(e).ToString(new string('0', Math.Max(1, expToks.Count)));
    }

    private static string FormatFraction(double x, Section sec)
    {
        int slash = sec.Toks.FindIndex(t => t.Kind == TK.Slash);
        var denToks = sec.Toks.Skip(slash + 1).Where(t => t.Kind == TK.Digit).ToList();
        var before = sec.Toks.Take(slash).ToList();
        // whole part exists if there is a literal (space) separating integer and numerator placeholders
        int numStart = before.FindLastIndex(t => t.Kind != TK.Digit);
        bool hasWhole = numStart >= 0 && before.Take(numStart).Any(t => t.Kind == TK.Digit);
        int maxDen = (int)Math.Pow(10, Math.Max(1, denToks.Count)) - 1;
        // fixed denominator like "# ?/8"
        var denLit = sec.Toks.Skip(slash + 1).TakeWhile(t => t.Kind is TK.Lit or TK.Digit).Select(t => t.S).Aggregate("", (a, b) => a + b);
        bool fixedDen = int.TryParse(denLit.Trim(), out int fd) && fd > 0 && denToks.Count == 0;
        bool neg = x < 0;
        double a = Math.Abs(x);
        long whole = hasWhole ? (long)Math.Floor(a) : 0;
        double frac = a - whole;
        long bestN = 0, bestD = 1;
        if (fixedDen) { bestD = fd; bestN = (long)Math.Round(frac * fd); }
        else
        {
            double bestErr = double.MaxValue;
            for (int d = 1; d <= maxDen; d++)
            {
                long n = (long)Math.Round(frac * d);
                double err = Math.Abs(frac - (double)n / d);
                if (err < bestErr - 1e-12) { bestErr = err; bestN = n; bestD = d; }
                if (err < 1e-12) break;
            }
        }
        if (bestN == bestD && hasWhole) { whole++; bestN = 0; }
        string res;
        if (bestN == 0) res = hasWhole ? whole.ToString(CultureInfo.InvariantCulture) : "0";
        else res = (hasWhole && whole > 0 ? whole + " " : "") + bestN + "/" + bestD;
        return (neg ? "-" : "") + res;
    }

    private static string FormatDate(double d, Section sec)
    {
        if (d < 0) return new string('#', 8);
        var dt = DateSerial.FromSerial(d);
        bool ampm = sec.Toks.Any(t => t.Kind == TK.AmPm);
        var culture = Vietnamese ? VnCulture : CultureInfo.InvariantCulture;
        var sb = new StringBuilder();
        foreach (var t in sec.Toks)
        {
            switch (t.Kind)
            {
                case TK.Lit: sb.Append(t.S == "\u0001GENERAL" ? General(d) : t.S); break;
                case TK.Digit: sb.Append(t.S == "0" ? "0" : ""); break;
                case TK.Dec: sb.Append(DecimalSep); break;
                case TK.Group: sb.Append(','); break;
                case TK.AmPm:
                    sb.Append(t.S == "A/P" ? (dt.Hour < 12 ? "A" : "P") : (dt.Hour < 12 ? "AM" : "PM"));
                    break;
                case TK.Elapsed:
                {
                    double totalH = d * 24;
                    sb.Append(t.S[0] switch
                    {
                        'h' => Math.Floor(totalH).ToString(t.S.Length > 1 ? "00" : "0", CultureInfo.InvariantCulture),
                        'm' => Math.Floor(d * 1440 + 1e-9).ToString(t.S.Length > 1 ? "00" : "0", CultureInfo.InvariantCulture),
                        _ => Math.Floor(d * 86400 + 1e-6).ToString(t.S.Length > 1 ? "00" : "0", CultureInfo.InvariantCulture)
                    });
                    break;
                }
                case TK.Date:
                {
                    var s = t.S;
                    if (s.StartsWith('.'))
                    {
                        double frac = dt.Millisecond / 1000.0;
                        sb.Append(DecimalSep).Append(Math.Round(frac * Math.Pow(10, s.Length - 1)).ToString(new string('0', s.Length - 1), CultureInfo.InvariantCulture));
                        break;
                    }
                    int hour12 = dt.Hour % 12 == 0 ? 12 : dt.Hour % 12;
                    sb.Append(s switch
                    {
                        "yy" or "y" => (dt.Year % 100).ToString("00"),
                        _ when s.StartsWith("yyy") => dt.Year.ToString("0000"),
                        "m" => dt.Month.ToString(),
                        "mm" => dt.Month.ToString("00"),
                        "mmm" => culture.DateTimeFormat.GetAbbreviatedMonthName(dt.Month),
                        "mmmmm" => culture.DateTimeFormat.GetMonthName(dt.Month)[..1],
                        _ when s.StartsWith("mmmm") => culture.DateTimeFormat.GetMonthName(dt.Month),
                        "d" => dt.Day.ToString(),
                        "dd" => dt.Day.ToString("00"),
                        "ddd" => culture.DateTimeFormat.GetAbbreviatedDayName(dt.DayOfWeek),
                        _ when s.StartsWith("dddd") => culture.DateTimeFormat.GetDayName(dt.DayOfWeek),
                        "h" => (ampm ? hour12 : dt.Hour).ToString(),
                        "hh" => (ampm ? hour12 : dt.Hour).ToString("00"),
                        "n" => dt.Minute.ToString(),
                        "nn" => dt.Minute.ToString("00"),
                        "s" => dt.Second.ToString(),
                        "ss" => dt.Second.ToString("00"),
                        _ when s.StartsWith('h') => (ampm ? hour12 : dt.Hour).ToString("00"),
                        _ when s.StartsWith('s') => dt.Second.ToString("00"),
                        _ when s.StartsWith('m') => dt.Month.ToString("00"),
                        _ => s
                    });
                    break;
                }
                case TK.Pct: sb.Append('%'); break;
            }
        }
        return sb.ToString();
    }

    private static readonly CultureInfo VnCulture = CreateVn();
    private static CultureInfo CreateVn()
    {
        try { return CultureInfo.GetCultureInfo("vi-VN"); } catch { return CultureInfo.InvariantCulture; }
    }

    /// <summary>Changes the number of decimals of a format code (Increase / Decrease Decimal buttons).</summary>
    public static string AdjustDecimals(string code, double sample, int delta)
    {
        if (code == "General" || string.IsNullOrEmpty(code))
        {
            var g = General(sample);
            int dec = g.Contains(DecimalSep) ? g.Length - g.IndexOf(DecimalSep) - 1 : 0;
            dec = Math.Max(0, dec + delta);
            return dec == 0 ? "0" : "0." + new string('0', dec);
        }
        var sections = SplitSections(code).Select(sec =>
        {
            int dot = IndexOutsideQuotes(sec, '.');
            if (delta > 0)
            {
                if (dot >= 0) return sec.Insert(dot + 1, "0");
                int last = LastIndexOutsideQuotes(sec, c => c is '0' or '#');
                return last < 0 ? sec : sec.Insert(last + 1, ".0");
            }
            if (dot < 0) return sec;
            int count = 0;
            int k = dot + 1;
            while (k < sec.Length && sec[k] is '0' or '#') { k++; count++; }
            if (count <= 1) return sec.Remove(dot, 1 + count);
            return sec.Remove(dot + 1, 1);
        });
        return string.Join(";", sections);
    }

    private static IEnumerable<string> SplitSections(string code)
    {
        var sb = new StringBuilder();
        bool q = false;
        foreach (var ch in code)
        {
            if (ch == '"') q = !q;
            if (ch == ';' && !q) { yield return sb.ToString(); sb.Clear(); continue; }
            sb.Append(ch);
        }
        yield return sb.ToString();
    }

    private static int IndexOutsideQuotes(string s, char target)
    {
        bool q = false, br = false;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '"') q = !q;
            else if (!q && s[i] == '[') br = true;
            else if (!q && s[i] == ']') br = false;
            else if (!q && !br && s[i] == target) return i;
        }
        return -1;
    }

    private static int LastIndexOutsideQuotes(string s, Func<char, bool> pred)
    {
        bool q = false, br = false;
        int last = -1;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '"') q = !q;
            else if (!q && s[i] == '[') br = true;
            else if (!q && s[i] == ']') br = false;
            else if (!q && !br && pred(s[i])) last = i;
        }
        return last;
    }
}
