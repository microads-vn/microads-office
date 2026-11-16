using System.Drawing.Drawing2D;
using System.Drawing.Text;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.IO;
using MicroadsOffice.Excel.Model;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Draws charts (column, stacked, bar, line, area, pie, doughnut, scatter) with title, axes and legend.</summary>
public static class ChartRenderer
{
    public sealed record SeriesData(string Name, double[] Values, Color Color);

    public static (string[] Categories, List<SeriesData> Series, double[]? XValues) Data(Sheet s, Chart ch)
    {
        var (cats, series) = ChartData.Layout(ch);
        string Txt(int r, int c) => SheetOps.DisplayText(s, r, c);
        string[] categories;
        int n;
        if (cats is CellRange cr)
        {
            categories = cr.Cells().Select(p => Txt(p.R, p.C)).ToArray();
            n = categories.Length;
        }
        else
        {
            n = series.Count > 0 ? (int)series[0].Values.Count : 0;
            categories = Enumerable.Range(1, n).Select(i => i.ToString()).ToArray();
        }
        double[]? xs = null;
        if (ch.Type == ChartType.Scatter && cats is CellRange xr)
            xs = xr.Cells().Select(p => { var v = s.GetValue(p.R, p.C); return v.IsNumber ? v.Num : double.NaN; }).ToArray();
        var list = new List<SeriesData>();
        int i = 0;
        foreach (var ser in series)
        {
            var name = ser.NameCell is CellPos np ? Txt(np.Row, np.Col) : Model.Addr.ColName(ser.Values.C1);
            if (string.IsNullOrWhiteSpace(name)) name = "Series " + (i + 1);
            var vals = ser.Values.Cells().Select(p => { var v = s.GetValue(p.R, p.C); return v.IsNumber ? v.Num : 0; }).ToArray();
            list.Add(new SeriesData(name, vals, ChartData.Palette[i % ChartData.Palette.Length]));
            i++;
        }
        return (categories, list, xs);
    }

    public static void Draw(Graphics g, RectangleF r, Sheet s, Chart ch, float scale = 1f)
    {
        var state = g.Save();
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
        g.SetClip(r);
        g.FillRectangle(Brushes.White, r);
        using var border = new Pen(Color.FromArgb(200, 202, 206));
        g.DrawRectangle(border, r.X, r.Y, r.Width - 1, r.Height - 1);
        var (cats, series, xs) = Data(s, ch);
        using var titleFont = new Font("Calibri", 13f * scale, FontStyle.Regular, GraphicsUnit.Pixel);
        using var tf = new Font("Calibri", 14f * scale, FontStyle.Regular, GraphicsUnit.Point);
        using var lblFont = new Font("Calibri", 8.5f * scale, FontStyle.Regular, GraphicsUnit.Point);
        using var ink = new SolidBrush(Color.FromArgb(89, 89, 89));
        float pad = 10 * scale;
        var area = RectangleF.Inflate(r, -pad, -pad);
        if (ch.Title.Length > 0)
        {
            var sz = g.MeasureString(ch.Title, tf);
            g.DrawString(ch.Title, tf, ink, r.X + (r.Width - sz.Width) / 2, area.Y);
            area.Y += sz.Height + 4 * scale; area.Height -= sz.Height + 4 * scale;
        }
        bool pie = ch.Type is ChartType.Pie or ChartType.Doughnut;
        if (ch.ShowLegend && series.Count > 0)
        {
            var items = pie ? cats.Select((c, i) => (c, ChartData.Palette[i % ChartData.Palette.Length])).ToList() : series.Select(x => (x.Name, x.Color)).ToList();
            if (pie)
            {
                float lw = Math.Min(area.Width * 0.35f, items.Max(it => g.MeasureString(it.Item1, lblFont).Width) + 22 * scale);
                float ly = area.Y + Math.Max(0, (area.Height - items.Count * 16 * scale) / 2);
                foreach (var (name, col) in items)
                {
                    using var b = new SolidBrush(col);
                    g.FillRectangle(b, area.Right - lw, ly + 3 * scale, 9 * scale, 9 * scale);
                    g.DrawString(name, lblFont, ink, area.Right - lw + 13 * scale, ly);
                    ly += 16 * scale;
                }
                area.Width -= lw + 6 * scale;
            }
            else
            {
                float totalW = items.Sum(it => g.MeasureString(it.Item1, lblFont).Width + 24 * scale);
                float lx = r.X + Math.Max(pad, (r.Width - totalW) / 2), ly = area.Bottom - 15 * scale;
                foreach (var (name, col) in items)
                {
                    using var b = new SolidBrush(col);
                    g.FillRectangle(b, lx, ly + 3 * scale, 9 * scale, 9 * scale);
                    g.DrawString(name, lblFont, ink, lx + 12 * scale, ly);
                    lx += g.MeasureString(name, lblFont).Width + 24 * scale;
                }
                area.Height -= 20 * scale;
            }
        }
        if (series.Count == 0 || cats.Length == 0) { g.Restore(state); return; }
        if (pie) DrawPie(g, area, ch, series[0], cats, lblFont, ink, scale);
        else DrawAxes(g, area, ch, series, cats, xs, lblFont, ink, scale);
        g.Restore(state);
    }

    private static void DrawPie(Graphics g, RectangleF area, Chart ch, SeriesData ser, string[] cats, Font f, Brush ink, float scale)
    {
        double total = ser.Values.Where(v => v > 0).Sum();
        if (total <= 0) return;
        float d = Math.Min(area.Width, area.Height);
        var pr = new RectangleF(area.X + (area.Width - d) / 2, area.Y + (area.Height - d) / 2, d, d);
        float start = -90;
        using var white = new Pen(Color.White, 1.5f * scale);
        for (int i = 0; i < ser.Values.Length; i++)
        {
            var v = ser.Values[i];
            if (v <= 0) continue;
            float sweep = (float)(v / total * 360);
            using var b = new SolidBrush(ChartData.Palette[i % ChartData.Palette.Length]);
            g.FillPie(b, pr.X, pr.Y, pr.Width, pr.Height, start, sweep);
            g.DrawPie(white, pr.X, pr.Y, pr.Width, pr.Height, start, sweep);
            if (ch.ShowDataLabels)
            {
                var mid = (start + sweep / 2) * Math.PI / 180;
                float rad = d * (ch.Type == ChartType.Doughnut ? 0.39f : 0.32f);
                var txt = (v / total).ToString("0%");
                var sz = g.MeasureString(txt, f);
                g.DrawString(txt, f, Brushes.White, pr.X + d / 2 + rad * (float)Math.Cos(mid) - sz.Width / 2, pr.Y + d / 2 + rad * (float)Math.Sin(mid) - sz.Height / 2);
            }
            start += sweep;
        }
        if (ch.Type == ChartType.Doughnut)
        {
            float h = d * 0.55f;
            g.FillEllipse(Brushes.White, pr.X + (d - h) / 2, pr.Y + (d - h) / 2, h, h);
        }
    }

    private static void DrawAxes(Graphics g, RectangleF area, Chart ch, List<SeriesData> series, string[] cats, double[]? xs, Font f, Brush ink, float scale)
    {
        bool horizontal = ch.Type == ChartType.Bar;
        bool stacked = ch.Type == ChartType.StackedColumn;
        int n = Math.Min(cats.Length, series.Max(x => x.Values.Length));
        double min = 0, max = 0;
        if (stacked)
            for (int i = 0; i < n; i++)
            {
                double pos = series.Sum(x => i < x.Values.Length && x.Values[i] > 0 ? x.Values[i] : 0), neg = series.Sum(x => i < x.Values.Length && x.Values[i] < 0 ? x.Values[i] : 0);
                max = Math.Max(max, pos); min = Math.Min(min, neg);
            }
        else foreach (var x in series) foreach (var v in x.Values.Take(n)) { max = Math.Max(max, v); min = Math.Min(min, v); }
        if (max == min) max = min + 1;
        var (lo, hi, step) = NiceScale(min, max);
        // value axis labels
        var labels = new List<double>();
        for (double v = lo; v <= hi + step / 2; v += step) labels.Add(Math.Round(v, 10));
        float labelW = labels.Max(v => g.MeasureString(NumFmt.General(v), f).Width) + 4 * scale;
        float catH = g.MeasureString("Ag", f).Height + 4 * scale;
        RectangleF plot;
        if (horizontal)
        {
            float catW = Math.Min(area.Width * 0.3f, cats.Take(n).Select(c => g.MeasureString(c, f).Width).DefaultIfEmpty(0).Max() + 6 * scale);
            plot = new RectangleF(area.X + catW, area.Y + 4 * scale, area.Width - catW - 8 * scale, area.Height - catH - 4 * scale);
        }
        else plot = new RectangleF(area.X + labelW, area.Y + 6 * scale, area.Width - labelW - 6 * scale, area.Height - catH - 6 * scale);
        if (plot.Width < 10 || plot.Height < 10) return;
        using var gridPen = new Pen(Color.FromArgb(217, 217, 217), 1);
        using var axisPen = new Pen(Color.FromArgb(191, 191, 191), 1);
        float ValPos(double v) => horizontal ? plot.X + (float)((v - lo) / (hi - lo)) * plot.Width : plot.Bottom - (float)((v - lo) / (hi - lo)) * plot.Height;
        foreach (var v in labels)
        {
            var p = ValPos(v);
            var t = NumFmt.General(v);
            var sz = g.MeasureString(t, f);
            if (horizontal) { g.DrawLine(gridPen, p, plot.Y, p, plot.Bottom); g.DrawString(t, f, ink, p - sz.Width / 2, plot.Bottom + 2 * scale); }
            else { g.DrawLine(gridPen, plot.X, p, plot.Right, p); g.DrawString(t, f, ink, plot.X - sz.Width - 2 * scale, p - sz.Height / 2); }
        }
        if (ch.Type == ChartType.Scatter && xs != null)
        {
            double xmin = xs.Where(x => !double.IsNaN(x)).DefaultIfEmpty(0).Min(), xmax = xs.Where(x => !double.IsNaN(x)).DefaultIfEmpty(1).Max();
            if (xmax == xmin) xmax = xmin + 1;
            var (xlo, xhi, xstep) = NiceScale(xmin, xmax);
            for (double v = xlo; v <= xhi + xstep / 2; v += xstep)
            {
                float px = plot.X + (float)((v - xlo) / (xhi - xlo)) * plot.Width;
                var t = NumFmt.General(Math.Round(v, 10));
                var sz = g.MeasureString(t, f);
                g.DrawString(t, f, ink, px - sz.Width / 2, plot.Bottom + 2 * scale);
            }
            foreach (var ser in series)
            {
                using var b = new SolidBrush(ser.Color);
                var pts = new List<PointF>();
                for (int i = 0; i < Math.Min(xs.Length, ser.Values.Length); i++)
                {
                    if (double.IsNaN(xs[i])) continue;
                    var p = new PointF(plot.X + (float)((xs[i] - xlo) / (xhi - xlo)) * plot.Width, ValPos(ser.Values[i]));
                    pts.Add(p);
                    g.FillEllipse(b, p.X - 3.5f * scale, p.Y - 3.5f * scale, 7 * scale, 7 * scale);
                }
            }
            g.DrawLine(axisPen, plot.X, plot.Bottom, plot.Right, plot.Bottom);
            return;
        }
        // category axis
        float slot = (horizontal ? plot.Height : plot.Width) / Math.Max(1, n);
        for (int i = 0; i < n; i++)
        {
            var t = cats[i];
            var sz = g.MeasureString(t, f);
            if (horizontal) g.DrawString(t, f, ink, plot.X - sz.Width - 3 * scale, plot.Y + slot * i + (slot - sz.Height) / 2);
            else
            {
                if (sz.Width > slot && slot < 30 * scale && i % 2 == 1) continue;
                g.DrawString(t, f, ink, new RectangleF(plot.X + slot * i, plot.Bottom + 2 * scale, slot, catH), new StringFormat { Alignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap });
            }
        }
        float zero = ValPos(Math.Clamp(0, lo, hi));
        if (horizontal) g.DrawLine(axisPen, zero, plot.Y, zero, plot.Bottom); else g.DrawLine(axisPen, plot.X, zero, plot.Right, zero);
        switch (ch.Type)
        {
            case ChartType.Column: case ChartType.Bar:
            {
                int k = series.Count;
                float group = slot * 0.7f, bw = group / k;
                for (int si = 0; si < k; si++)
                {
                    using var b = new SolidBrush(series[si].Color);
                    for (int i = 0; i < n && i < series[si].Values.Length; i++)
                    {
                        var v = series[si].Values[i];
                        float p = ValPos(v);
                        float start = slot * i + (slot - group) / 2 + bw * si;
                        RectangleF bar = horizontal
                            ? new RectangleF(Math.Min(zero, p), plot.Y + start, Math.Abs(p - zero), bw * 0.92f)
                            : new RectangleF(plot.X + start, Math.Min(zero, p), bw * 0.92f, Math.Abs(p - zero));
                        g.FillRectangle(b, bar);
                        if (ch.ShowDataLabels)
                        {
                            var t = NumFmt.General(v);
                            var sz = g.MeasureString(t, f);
                            if (horizontal) g.DrawString(t, f, ink, bar.Right + 2 * scale, bar.Y + (bar.Height - sz.Height) / 2);
                            else g.DrawString(t, f, ink, bar.X + (bar.Width - sz.Width) / 2, bar.Y - sz.Height);
                        }
                    }
                }
                break;
            }
            case ChartType.StackedColumn:
            {
                float bw = slot * 0.6f;
                for (int i = 0; i < n; i++)
                {
                    double pos = 0, neg = 0;
                    foreach (var ser in series)
                    {
                        if (i >= ser.Values.Length) continue;
                        var v = ser.Values[i];
                        double from = v >= 0 ? pos : neg, to = from + v;
                        if (v >= 0) pos = to; else neg = to;
                        float y1 = ValPos(from), y2 = ValPos(to);
                        using var b = new SolidBrush(ser.Color);
                        g.FillRectangle(b, plot.X + slot * i + (slot - bw) / 2, Math.Min(y1, y2), bw, Math.Abs(y2 - y1));
                    }
                }
                break;
            }
            case ChartType.Line: case ChartType.Area:
                foreach (var ser in series)
                {
                    var pts = Enumerable.Range(0, Math.Min(n, ser.Values.Length)).Select(i => new PointF(plot.X + slot * i + slot / 2, ValPos(ser.Values[i]))).ToArray();
                    if (pts.Length == 0) continue;
                    if (ch.Type == ChartType.Area)
                    {
                        var poly = new List<PointF> { new(pts[0].X, zero) };
                        poly.AddRange(pts);
                        poly.Add(new PointF(pts[^1].X, zero));
                        using var ab = new SolidBrush(Color.FromArgb(200, ser.Color));
                        g.FillPolygon(ab, poly.ToArray());
                    }
                    else
                    {
                        using var lp = new Pen(ser.Color, 2.25f * scale) { LineJoin = LineJoin.Round };
                        if (pts.Length > 1) g.DrawLines(lp, pts);
                        using var mb = new SolidBrush(ser.Color);
                        foreach (var p in pts) g.FillEllipse(mb, p.X - 3 * scale, p.Y - 3 * scale, 6 * scale, 6 * scale);
                    }
                    if (ch.ShowDataLabels)
                        for (int i = 0; i < pts.Length; i++)
                        {
                            var t = NumFmt.General(ser.Values[i]);
                            var sz = g.MeasureString(t, f);
                            g.DrawString(t, f, ink, pts[i].X - sz.Width / 2, pts[i].Y - sz.Height - 2 * scale);
                        }
                }
                break;
        }
    }

    public static (double Lo, double Hi, double Step) NiceScale(double min, double max)
    {
        double range = max - min;
        double rough = range / 5;
        double mag = Math.Pow(10, Math.Floor(Math.Log10(rough)));
        double norm = rough / mag;
        double step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
        double lo = Math.Floor(min / step) * step, hi = Math.Ceiling(max / step) * step;
        if (hi == lo) hi = lo + step;
        return (lo, hi, step);
    }
}
