using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Formula;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Result of conditional formatting for one cell.</summary>
public struct CfResult
{
    public int Fill, FontColor;
    public bool Bold;
    public double? Bar;
    public int BarColor;
}

/// <summary>Evaluates conditional formatting rules (statistics are computed once per rule and version).</summary>
public sealed class CfEvaluator
{
    private readonly Sheet _s;
    private readonly Dictionary<ConditionalFormat, object> _stats = new();
    public CfEvaluator(Sheet s) { _s = s; }

    private List<double> Numbers(CellRange r)
    {
        var used = _s.UsedRange();
        var res = new List<double>();
        for (int row = r.R1; row <= Math.Min(r.R2, used.R2); row++)
            for (int c = r.C1; c <= Math.Min(r.C2, used.C2); c++) { var v = _s.GetValue(row, c); if (v.IsNumber) res.Add(v.Num); }
        return res;
    }

    public CfResult? Evaluate(int r, int c)
    {
        CfResult? res = null;
        foreach (var cf in _s.CondFormats)
        {
            if (!cf.Range.Contains(r, c)) continue;
            var v = _s.GetValue(r, c);
            bool hit = false;
            switch (cf.Type)
            {
                case CfType.GreaterThan: hit = v.IsNumber && TryNum(cf.Value1, r, c, cf, out var a) && v.Num > a; break;
                case CfType.LessThan: hit = v.IsNumber && TryNum(cf.Value1, r, c, cf, out var b) && v.Num < b; break;
                case CfType.Between: hit = v.IsNumber && TryNum(cf.Value1, r, c, cf, out var lo) && TryNum(cf.Value2, r, c, cf, out var hi) && v.Num >= Math.Min(lo, hi) && v.Num <= Math.Max(lo, hi); break;
                case CfType.EqualTo:
                    hit = TryNum(cf.Value1, r, c, cf, out var eq) ? v.IsNumber && Math.Abs(v.Num - eq) < 1e-9 : !v.IsBlank && string.Equals(Calc.ToTextSafe(v), cf.Value1.Trim('"'), StringComparison.CurrentCultureIgnoreCase);
                    break;
                case CfType.TextContains: hit = !v.IsBlank && Calc.ToTextSafe(v).Contains(cf.Value1, StringComparison.CurrentCultureIgnoreCase); break;
                case CfType.Duplicate:
                {
                    if (!_stats.TryGetValue(cf, out var o))
                    {
                        var counts = new Dictionary<string, int>();
                        var used = _s.UsedRange();
                        for (int rr = cf.Range.R1; rr <= Math.Min(cf.Range.R2, used.R2); rr++)
                            for (int cc = cf.Range.C1; cc <= Math.Min(cf.Range.C2, used.C2); cc++)
                            {
                                var x = _s.GetValue(rr, cc);
                                if (x.IsBlank) continue;
                                var k = x.Kind + ":" + x.ToString().ToUpperInvariant();
                                counts[k] = counts.GetValueOrDefault(k) + 1;
                            }
                        _stats[cf] = o = counts;
                    }
                    hit = !v.IsBlank && ((Dictionary<string, int>)o).GetValueOrDefault(v.Kind + ":" + v.ToString().ToUpperInvariant()) > 1;
                    break;
                }
                case CfType.Top10: case CfType.Bottom10:
                {
                    if (!_stats.TryGetValue(cf, out var o))
                    {
                        var nums = Numbers(cf.Range);
                        nums.Sort();
                        double th = nums.Count == 0 ? double.NaN : cf.Type == CfType.Top10 ? nums[Math.Max(0, nums.Count - cf.Rank)] : nums[Math.Min(nums.Count - 1, cf.Rank - 1)];
                        _stats[cf] = o = th;
                    }
                    var t = (double)o;
                    hit = v.IsNumber && !double.IsNaN(t) && (cf.Type == CfType.Top10 ? v.Num >= t : v.Num <= t);
                    break;
                }
                case CfType.AboveAverage: case CfType.BelowAverage:
                {
                    if (!_stats.TryGetValue(cf, out var o)) { var nums = Numbers(cf.Range); _stats[cf] = o = nums.Count == 0 ? double.NaN : nums.Average(); }
                    var avg = (double)o;
                    hit = v.IsNumber && (cf.Type == CfType.AboveAverage ? v.Num > avg : v.Num < avg);
                    break;
                }
                case CfType.Expression:
                {
                    var f = RefAdjust.Shift(cf.Value1.TrimStart('='), r - cf.Range.R1, c - cf.Range.C1);
                    var res2 = _s.Book.Calc.EvaluateFormula(_s, r, c, f);
                    hit = res2.IsBool ? res2.BoolValue : res2.IsNumber && res2.Num != 0;
                    break;
                }
                case CfType.DataBar:
                {
                    if (!v.IsNumber) continue;
                    var (mn, mx) = MinMax(cf);
                    var cur = res ?? new CfResult();
                    cur.Bar = mx > mn ? Math.Clamp((v.Num - Math.Min(0, mn)) / (mx - Math.Min(0, mn)), 0, 1) : 1;
                    cur.BarColor = cf.Color1;
                    res = cur;
                    continue;
                }
                case CfType.ColorScale2: case CfType.ColorScale3:
                {
                    if (!v.IsNumber) continue;
                    var (mn, mx) = MinMax(cf);
                    double t = mx > mn ? (v.Num - mn) / (mx - mn) : 0.5;
                    Color col;
                    if (cf.Type == CfType.ColorScale2) col = Lerp(Color.FromArgb(cf.Color1), Color.FromArgb(cf.Color3), t);
                    else col = t < 0.5 ? Lerp(Color.FromArgb(cf.Color1), Color.FromArgb(cf.Color2), t * 2) : Lerp(Color.FromArgb(cf.Color2), Color.FromArgb(cf.Color3), (t - 0.5) * 2);
                    var cur = res ?? new CfResult();
                    if (cur.Fill == 0) cur.Fill = col.ToArgb();
                    res = cur;
                    continue;
                }
            }
            if (!hit) continue;
            var curr = res ?? new CfResult();
            if (curr.Fill == 0) curr.Fill = cf.Fill;
            if (curr.FontColor == 0) curr.FontColor = cf.FontColor;
            curr.Bold |= cf.Bold;
            res = curr;
        }
        return res;
    }

    private (double, double) MinMax(ConditionalFormat cf)
    {
        if (!_stats.TryGetValue(cf, out var o))
        {
            var nums = Numbers(cf.Range);
            _stats[cf] = o = nums.Count == 0 ? (0.0, 0.0) : (nums.Min(), nums.Max());
        }
        return ((double, double))o;
    }

    private bool TryNum(string text, int r, int c, ConditionalFormat cf, out double d)
    {
        text = text.Trim();
        if (text.StartsWith('='))
        {
            var res = _s.Book.Calc.EvaluateFormula(_s, r, c, RefAdjust.Shift(text[1..], r - cf.Range.R1, c - cf.Range.C1));
            d = res.IsNumber ? res.Num : 0;
            return res.IsNumber;
        }
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return true;
        return ValueParser.TryCoerceNumber(text, out d);
    }

    private static Color Lerp(Color a, Color b, double t)
    {
        t = Math.Clamp(t, 0, 1);
        return Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));
    }
}

/// <summary>Draws cells (fills, gridlines, CF, text with overflow/wrap, borders) for screen and print.</summary>
public static class CellPainter
{
    public readonly record struct Line(int Index, float Pos, float Size);

    private static readonly Dictionary<(string, float, FontStyle), Font> Fonts = new();

    public static Font GetFont(string name, float size, FontStyle style)
    {
        size = (float)Math.Round(Math.Max(1, size), 2);
        var key = (name, size, style);
        if (Fonts.TryGetValue(key, out var f)) return f;
        if (Fonts.Count > 400) { foreach (var x in Fonts.Values) x.Dispose(); Fonts.Clear(); }
        try { f = new Font(name, size, style, GraphicsUnit.Point); }
        catch { f = new Font("Segoe UI", size, style, GraphicsUnit.Point); }
        Fonts[key] = f;
        return f;
    }

    public static FontStyle StyleOf(CellStyle st, bool forceBold = false)
    {
        var fs = FontStyle.Regular;
        if (st.Bold || forceBold) fs |= FontStyle.Bold;
        if (st.Italic) fs |= FontStyle.Italic;
        if (st.Underline) fs |= FontStyle.Underline;
        if (st.Strike) fs |= FontStyle.Strikeout;
        return fs;
    }

    public static string DisplayText(Sheet s, Cell? cell, CellStyle st, bool showFormulas)
    {
        if (cell == null) return "";
        if (showFormulas && cell.Formula != null) return "=" + cell.Formula;
        if (showFormulas && cell.Value.IsNumber) return CellValue.FormatGeneralInvariant(cell.Value.Num);
        return NumFmt.Format(cell.Value, st.NumberFormat, out _);
    }

    /// <param name="fontScale">Multiplier for font point sizes (zoom).</param>
    /// <param name="gdi">True on screen (ClearType text), false for printing (GDI+ text, device independent).</param>
    public static void Paint(Graphics g, Sheet s, IReadOnlyList<Line> rows, IReadOnlyList<Line> cols, float scale, float fontScale, bool gridlines, bool gdi, bool showFormulas, CfEvaluator? cf, RectangleF clip)
    {
        if (rows.Count == 0 || cols.Count == 0) return;
        var styles = s.Book.Styles;
        var rowIdx = new Dictionary<int, int>(rows.Count);
        for (int i = 0; i < rows.Count; i++) rowIdx[rows[i].Index] = i;
        var colIdx = new Dictionary<int, int>(cols.Count);
        for (int i = 0; i < cols.Count; i++) colIdx[cols[i].Index] = i;
        float left = cols[0].Pos, top = rows[0].Pos, right = cols[^1].Pos + cols[^1].Size, bottom = rows[^1].Pos + rows[^1].Size;
        // visible merges
        var merges = s.Merges.Where(m => m.R2 >= rows[0].Index && m.R1 <= rows[^1].Index && m.C2 >= cols[0].Index && m.C1 <= cols[^1].Index).ToList();
        RectangleF? MergeRect(CellRange m)
        {
            float x1 = float.MaxValue, x2 = float.MinValue, y1 = float.MaxValue, y2 = float.MinValue;
            foreach (var c in cols) if (c.Index >= m.C1 && c.Index <= m.C2) { x1 = Math.Min(x1, c.Pos); x2 = Math.Max(x2, c.Pos + c.Size); }
            foreach (var r in rows) if (r.Index >= m.R1 && r.Index <= m.R2) { y1 = Math.Min(y1, r.Pos); y2 = Math.Max(y2, r.Pos + r.Size); }
            if (x1 == float.MaxValue || y1 == float.MaxValue) return null;
            return new RectangleF(x1, y1, x2 - x1, y2 - y1);
        }
        bool InMergeNotAnchor(int r, int c) { foreach (var m in merges) if (m.Contains(r, c) && !(m.R1 == r && m.C1 == c)) return true; return false; }

        // 1. gridlines
        if (gridlines)
        {
            using var gp = new Pen(Color.FromArgb(218, 220, 224));
            foreach (var r in rows) g.DrawLine(gp, left, r.Pos + r.Size - 1, right - 1, r.Pos + r.Size - 1);
            foreach (var c in cols) g.DrawLine(gp, c.Pos + c.Size - 1, top, c.Pos + c.Size - 1, bottom - 1);
        }
        // collect content cells in view
        var cache = new Dictionary<(int, int), (Cell? Cell, CellStyle St, CfResult? Cf)>();
        (Cell? Cell, CellStyle St, CfResult? Cf) Info(int r, int c)
        {
            if (cache.TryGetValue((r, c), out var x)) return x;
            var cell = s.GetCell(r, c);
            var st = styles[cell?.StyleId ?? s.StyleIdAt(r, c)];
            var res = cf != null && s.CondFormats.Count > 0 ? cf.Evaluate(r, c) : null;
            x = (cell, st, res);
            cache[(r, c)] = x;
            return x;
        }
        // 2. fills (+ merged areas cover inner gridlines)
        foreach (var r in rows)
            foreach (var c in cols)
            {
                if (InMergeNotAnchor(r.Index, c.Index)) continue;
                var (cell, st, cfr) = Info(r.Index, c.Index);
                int fill = cfr?.Fill is int cfFill && cfFill != 0 ? cfFill : st.Fill;
                RectangleF rect = new(c.Pos, r.Pos, c.Size - (gridlines ? 1 : 0), r.Size - (gridlines ? 1 : 0));
                var merge = merges.FirstOrDefault(m => m.R1 == r.Index && m.C1 == c.Index && !m.IsSingle);
                if (merge != default && MergeRect(merge) is RectangleF mr) rect = new RectangleF(mr.X, mr.Y, mr.Width - (gridlines ? 1 : 0), mr.Height - (gridlines ? 1 : 0));
                if (fill != 0)
                {
                    using var b = new SolidBrush(Color.FromArgb(255, Color.FromArgb(fill)));
                    g.FillRectangle(b, rect.X, rect.Y, rect.Width + (gridlines ? 1 : 0), rect.Height + (gridlines ? 1 : 0));
                }
                else if (merge != default) g.FillRectangle(Brushes.White, rect);
                if (cfr?.Bar is double bar && bar > 0)
                {
                    var bc = Color.FromArgb(cfr.Value.BarColor);
                    var br = new RectangleF(rect.X + 1, rect.Y + 2, (rect.Width - 3) * (float)bar, rect.Height - 4);
                    if (br.Width > 0)
                    {
                        using var lb = new LinearGradientBrush(new RectangleF(br.X - 1, br.Y, br.Width + 2, br.Height), bc, Color.FromArgb(40, bc), LinearGradientMode.Horizontal);
                        g.FillRectangle(lb, br);
                        using var bp = new Pen(bc);
                        g.DrawRectangle(bp, br.X, br.Y, br.Width, br.Height);
                    }
                }
            }
        // 3. text
        var oldHint = g.TextRenderingHint;
        if (!gdi) g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
        float pad = 3 * scale;
        for (int ri = 0; ri < rows.Count; ri++)
        {
            var r = rows[ri];
            for (int ci = 0; ci < cols.Count; ci++)
            {
                var c = cols[ci];
                if (InMergeNotAnchor(r.Index, c.Index)) continue;
                var (cell, st, cfr) = Info(r.Index, c.Index);
                if (cell == null || (cell.Value.IsBlank && !(showFormulas && cell.Formula != null))) continue;
                string text = DisplayText(s, cell, st, showFormulas);
                if (text.Length == 0) continue;
                Color color = st.FontColorValue;
                if (!showFormulas) { NumFmt.Format(cell.Value, st.NumberFormat, out var fc); if (fc is Color fcc) color = fcc; }
                if (cfr?.FontColor is int cfc && cfc != 0) color = Color.FromArgb(cfc);
                if (s.Links.ContainsKey(Sheet.Key(r.Index, c.Index)) && st.FontColor == 0) color = Color.FromArgb(5, 99, 193);
                var font = GetFont(st.FontName, st.FontSize * fontScale, StyleOf(st, cfr?.Bold == true) | (s.Links.ContainsKey(Sheet.Key(r.Index, c.Index)) ? FontStyle.Underline : 0));
                var v = cell.Value;
                var align = st.HAlign;
                if (align == HAlign.General) align = showFormulas ? HAlign.Left : v.IsNumber ? HAlign.Right : v.IsBool || v.IsError ? HAlign.Center : HAlign.Left;
                RectangleF rect = new(c.Pos, r.Pos, c.Size, r.Size);
                var merge = merges.FirstOrDefault(m => m.R1 == r.Index && m.C1 == c.Index && !m.IsSingle);
                if (merge != default && MergeRect(merge) is RectangleF mr) rect = mr;
                float indent = st.Indent * 9 * scale;
                var textRect = new RectangleF(rect.X + pad + (align == HAlign.Left ? indent : 0), rect.Y, rect.Width - 2 * pad - indent, rect.Height - 1);
                var size = Measure(g, text, font, gdi, st.Wrap ? textRect.Width : 0);
                bool numeric = v.IsNumber && !showFormulas;
                if (numeric && !st.Wrap && size.Width > textRect.Width + 1 && merge == default)
                {
                    text = new string('#', Math.Max(1, (int)(textRect.Width / Math.Max(1, Measure(g, "#", font, gdi, 0).Width))));
                    size = Measure(g, text, font, gdi, 0);
                }
                // overflow into empty neighbours for unwrapped text
                RectangleF clipRect = rect;
                if (!st.Wrap && !numeric && merge == default && size.Width > textRect.Width && align is HAlign.Left or HAlign.Center or HAlign.Right)
                {
                    float extraR = 0, extraL = 0;
                    if (align is HAlign.Left or HAlign.Center)
                        for (int k = ci + 1; k < cols.Count && size.Width > textRect.Width + extraR; k++)
                        {
                            if (!IsEmpty(s, r.Index, cols[k].Index) || s.MergeAt(r.Index, cols[k].Index) != null) break;
                            extraR += cols[k].Size;
                        }
                    if (align is HAlign.Right or HAlign.Center)
                        for (int k = ci - 1; k >= 0 && size.Width > textRect.Width + extraL; k--)
                        {
                            if (!IsEmpty(s, r.Index, cols[k].Index) || s.MergeAt(r.Index, cols[k].Index) != null) break;
                            extraL += cols[k].Size;
                        }
                    if (align == HAlign.Center) { float e = Math.Min(extraL, extraR); extraL = extraR = e; }
                    clipRect = new RectangleF(rect.X - extraL, rect.Y, rect.Width + extraL + extraR, rect.Height);
                    if (align == HAlign.Left) textRect.Width += extraR;
                    else if (align == HAlign.Right) { textRect.X -= extraL; textRect.Width += extraL; }
                    else { textRect.X -= extraL; textRect.Width += extraL + extraR; }
                }
                float ty = st.VAlign switch
                {
                    VAlign.Top => textRect.Y + 1,
                    VAlign.Center => textRect.Y + (textRect.Height - Math.Min(size.Height, textRect.Height)) / 2,
                    _ => textRect.Bottom - Math.Min(size.Height, textRect.Height) - 1 * scale
                };
                if (st.Wrap && size.Height > textRect.Height) ty = textRect.Y;
                float tx = align switch
                {
                    HAlign.Right => textRect.Right - size.Width,
                    HAlign.Center or HAlign.CenterAcross => textRect.X + (textRect.Width - size.Width) / 2,
                    _ => textRect.X
                };
                var state = g.Save();
                var ic = RectangleF.Intersect(clipRect, clip);
                g.SetClip(ic);
                DrawText(g, text, font, color, st.Wrap ? new RectangleF(textRect.X, ty, textRect.Width, textRect.Bottom - ty) : new RectangleF(tx, ty, size.Width + 2, size.Height),
                    gdi, st.Wrap, align);
                g.Restore(state);
            }
        }
        g.TextRenderingHint = oldHint;
        // 4. borders
        var pens = new Dictionary<(LineStyle, int), Pen>();
        Pen PenOf(BorderLine b)
        {
            if (!pens.TryGetValue((b.Style, b.Argb), out var p))
            {
                float w = b.Style switch { LineStyle.Medium => 2f, LineStyle.Thick => 3f, LineStyle.Double => 3f, _ => 1f } * Math.Max(1, scale * 0.85f);
                p = new Pen(b.Argb == 0 ? Color.Black : Color.FromArgb(255, Color.FromArgb(b.Argb)), w);
                if (b.Style == LineStyle.Dashed) p.DashStyle = DashStyle.Dash;
                if (b.Style == LineStyle.Dotted || b.Style == LineStyle.Hair) p.DashStyle = DashStyle.Dot;
                if (b.Style == LineStyle.Double) p.CompoundArray = [0f, 0.33f, 0.67f, 1f];
                pens[(b.Style, b.Argb)] = p;
            }
            return p;
        }
        foreach (var r in rows)
            foreach (var c in cols)
            {
                var st = Info(r.Index, c.Index).St;
                if (st.Left.IsNone && st.Right.IsNone && st.Top.IsNone && st.Bottom.IsNone) continue;
                float x1 = c.Pos - 1, y1 = r.Pos - 1, x2 = c.Pos + c.Size - 1, y2 = r.Pos + r.Size - 1;
                var m = s.MergeAt(r.Index, c.Index);
                if (!st.Top.IsNone && (m == null || m.Value.R1 == r.Index)) g.DrawLine(PenOf(st.Top), x1, y1, x2, y1);
                if (!st.Bottom.IsNone && (m == null || m.Value.R2 == r.Index)) g.DrawLine(PenOf(st.Bottom), x1, y2, x2, y2);
                if (!st.Left.IsNone && (m == null || m.Value.C1 == c.Index)) g.DrawLine(PenOf(st.Left), x1, y1, x1, y2);
                if (!st.Right.IsNone && (m == null || m.Value.C2 == c.Index)) g.DrawLine(PenOf(st.Right), x2, y1, x2, y2);
            }
        foreach (var p in pens.Values) p.Dispose();
    }

    private static bool IsEmpty(Sheet s, int r, int c)
    {
        var cell = s.GetCell(r, c);
        return cell == null || (cell.Value.IsBlank && cell.Formula == null) || (cell.Formula != null && cell.Value.IsText && cell.Value.Str!.Length == 0);
    }

    public static SizeF Measure(Graphics g, string text, Font f, bool gdi, float wrapWidth)
    {
        if (gdi)
        {
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix;
            if (wrapWidth > 0) return TextRenderer.MeasureText(g, text, f, new Size((int)Math.Max(1, wrapWidth), int.MaxValue), flags | TextFormatFlags.WordBreak);
            return TextRenderer.MeasureText(g, text, f, new Size(int.MaxValue, int.MaxValue), flags);
        }
        using var sf = (StringFormat)StringFormat.GenericTypographic.Clone();
        sf.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
        return wrapWidth > 0 ? g.MeasureString(text, f, (int)Math.Max(1, wrapWidth), sf) : g.MeasureString(text, f, int.MaxValue, sf);
    }

    private static void DrawText(Graphics g, string text, Font f, Color color, RectangleF r, bool gdi, bool wrap, HAlign align)
    {
        if (gdi)
        {
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping;
            if (wrap)
            {
                flags |= TextFormatFlags.WordBreak;
                flags |= align switch { HAlign.Right => TextFormatFlags.Right, HAlign.Center or HAlign.CenterAcross => TextFormatFlags.HorizontalCenter, _ => TextFormatFlags.Left };
            }
            TextRenderer.DrawText(g, text, f, Rectangle.Round(r), color, flags);
            return;
        }
        using var b = new SolidBrush(color);
        using var sf = (StringFormat)StringFormat.GenericTypographic.Clone();
        sf.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
        if (!wrap) sf.FormatFlags |= StringFormatFlags.NoWrap;
        sf.Alignment = wrap ? align switch { HAlign.Right => StringAlignment.Far, HAlign.Center or HAlign.CenterAcross => StringAlignment.Center, _ => StringAlignment.Near } : StringAlignment.Near;
        g.DrawString(text, f, b, r, sf);
    }
}
