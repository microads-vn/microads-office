using System.Drawing.Drawing2D;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Vector icons for spreadsheet commands (drawn on a 16×16 design grid, scaled to any size).</summary>
public static class XlIcons
{
    public static readonly Color Green = Color.FromArgb(33, 115, 70);
    public static readonly Color Blue = Color.FromArgb(46, 117, 182);
    public static readonly Color Red = Color.FromArgb(200, 50, 50);
    public static readonly Color Orange = Color.FromArgb(237, 125, 49);
    public static readonly Color Gold = Color.FromArgb(255, 192, 0);
    public static readonly Color Gray = Color.FromArgb(150, 150, 155);

    public static Bitmap Small(string key) => Get(key, Dpi.Scale(16));
    public static Bitmap Large(string key) => Get(key, Dpi.Scale(32));

    public static Bitmap Get(string key, int size) => Icons.Painted("xl." + key, size, Theme.Ink, (g, s, ink) => Paint(g, key, s / 16f, ink));

    private static void Paint(Graphics g, string key, float k, Color ink)
    {
        g.ScaleTransform(k, k);
        using var pen = new Pen(ink, 1.1f);
        using var thin = new Pen(Color.FromArgb(140, ink), 0.8f);
        using var inkB = new SolidBrush(ink);
        using var green = new SolidBrush(Green);
        using var blue = new SolidBrush(Blue);
        using var greenPen = new Pen(Green, 1.4f);
        void Grid(float x, float y, float w, float h, int rows, int cols, Pen? p = null)
        {
            p ??= thin;
            g.DrawRectangle(p, x, y, w, h);
            for (int i = 1; i < rows; i++) g.DrawLine(p, x, y + h * i / rows, x + w, y + h * i / rows);
            for (int i = 1; i < cols; i++) g.DrawLine(p, x + w * i / cols, y, x + w * i / cols, y + h);
        }
        void Text(string t, float x, float y, float sz, Brush b, FontStyle st = FontStyle.Bold)
        {
            using var f = new Font("Segoe UI", sz, st, GraphicsUnit.Pixel);
            g.DrawString(t, f, b, x, y);
        }
        void Arrow(float x1, float y1, float x2, float y2, Pen p)
        {
            g.DrawLine(p, x1, y1, x2, y2);
            var ang = Math.Atan2(y2 - y1, x2 - x1);
            float a = 3f;
            g.DrawLine(p, x2, y2, x2 - a * (float)Math.Cos(ang - 0.6), y2 - a * (float)Math.Sin(ang - 0.6));
            g.DrawLine(p, x2, y2, x2 - a * (float)Math.Cos(ang + 0.6), y2 - a * (float)Math.Sin(ang + 0.6));
        }
        switch (key)
        {
            case "merge":
                Grid(1.5f, 3.5f, 13, 9, 1, 1, pen);
                Arrow(4, 8, 7, 8, pen); Arrow(12, 8, 9, 8, pen);
                g.DrawLine(thin, 8, 4, 8, 6); g.DrawLine(thin, 8, 10, 8, 12);
                break;
            case "wrap":
                g.DrawLine(pen, 2, 4, 14, 4);
                g.DrawLine(pen, 2, 8, 12, 8); g.DrawArc(pen, 9, 8, 4, 4, -90, 180); Arrow(11, 12, 8, 12, pen);
                g.DrawLine(pen, 2, 12, 6, 12);
                break;
            case "borders":
                Grid(2.5f, 2.5f, 11, 11, 2, 2, thin);
                g.DrawLine(new Pen(ink, 2f), 2, 13.5f, 14, 13.5f);
                break;
            case "borderAll": Grid(2.5f, 2.5f, 11, 11, 2, 2, pen); break;
            case "borderOutside": Grid(2.5f, 2.5f, 11, 11, 2, 2, thin); g.DrawRectangle(new Pen(ink, 1.5f), 2.5f, 2.5f, 11, 11); break;
            case "borderThick": Grid(2.5f, 2.5f, 11, 11, 2, 2, thin); g.DrawRectangle(new Pen(ink, 2.4f), 3, 3, 10, 10); break;
            case "borderNone": { using var dp = new Pen(Color.FromArgb(150, ink), 0.8f) { DashStyle = DashStyle.Dot }; Grid(2.5f, 2.5f, 11, 11, 2, 2, dp); break; }
            case "borderTop": Grid(2.5f, 2.5f, 11, 11, 2, 2, thin); g.DrawLine(new Pen(ink, 2f), 2, 2.5f, 14, 2.5f); break;
            case "borderLeft": Grid(2.5f, 2.5f, 11, 11, 2, 2, thin); g.DrawLine(new Pen(ink, 2f), 2.5f, 2, 2.5f, 14); break;
            case "borderRight": Grid(2.5f, 2.5f, 11, 11, 2, 2, thin); g.DrawLine(new Pen(ink, 2f), 13.5f, 2, 13.5f, 14); break;
            case "borderBottom": Grid(2.5f, 2.5f, 11, 11, 2, 2, thin); g.DrawLine(new Pen(ink, 2f), 2, 13.5f, 14, 13.5f); break;
            case "fill":
            {
                var pts = new[] { new PointF(3, 7), new PointF(8, 2), new PointF(13, 7), new PointF(8, 12) };
                g.FillPolygon(Brushes.White, pts); g.DrawPolygon(pen, pts);
                g.FillEllipse(inkB, 12.5f, 8.5f, 2.5f, 3.5f);
                break;
            }
            case "fontColor": Text("A", 2.5f, -1.5f, 13, inkB); break;
            case "growFont": Text("A", 0, 0, 12, inkB); g.FillPolygon(inkB, new[] { new PointF(10.5f, 6), new PointF(15.5f, 6), new PointF(13, 2.5f) }); break;
            case "shrinkFont": Text("A", 1, 3, 9, inkB); g.FillPolygon(inkB, new[] { new PointF(10.5f, 3), new PointF(15.5f, 3), new PointF(13, 6.5f) }); break;
            case "strike": Text("S", 3, -0.5f, 12, inkB, FontStyle.Regular); g.DrawLine(new Pen(ink, 1.3f), 1.5f, 8.5f, 14.5f, 8.5f); break;
            case "decInc": Text(".0", 0, 1, 7, inkB); Text(".00", 5, 7, 7, inkB); Arrow(3, 13, 3, 9, pen); break;
            case "decDec": Text(".00", 0, 1, 7, inkB); Text(".0", 7, 7, 7, inkB); Arrow(3, 9, 3, 13, pen); break;
            case "percent": Text("%", 1, -1, 14, inkB); break;
            case "comma": Text("000", -0.5f, 2, 8.5f, inkB); break;
            case "currency": Text("₫", 2, -1, 14, green); break;
            case "dollar": Text("$", 3, -1, 14, green); break;
            case "alignTop": g.DrawLine(pen, 2, 2.5f, 14, 2.5f); g.FillRectangle(inkB, 5, 4, 2, 7); g.FillRectangle(inkB, 9, 4, 2, 4); break;
            case "alignMiddle": g.DrawLine(pen, 2, 8, 14, 8); g.FillRectangle(inkB, 5, 4, 2, 8); g.FillRectangle(inkB, 9, 5.5f, 2, 5); break;
            case "alignBottom": g.DrawLine(pen, 2, 13.5f, 14, 13.5f); g.FillRectangle(inkB, 5, 5, 2, 7); g.FillRectangle(inkB, 9, 8, 2, 4); break;
            case "alignLeft": for (int i = 0; i < 4; i++) g.DrawLine(pen, 2, 3 + i * 3.3f, i % 2 == 0 ? 14 : 10, 3 + i * 3.3f); break;
            case "alignCenter": for (int i = 0; i < 4; i++) g.DrawLine(pen, i % 2 == 0 ? 2 : 4, 3 + i * 3.3f, i % 2 == 0 ? 14 : 12, 3 + i * 3.3f); break;
            case "alignRight": for (int i = 0; i < 4; i++) g.DrawLine(pen, i % 2 == 0 ? 2 : 6, 3 + i * 3.3f, 14, 3 + i * 3.3f); break;
            case "indentInc": for (int i = 0; i < 4; i++) g.DrawLine(pen, i is 1 or 2 ? 7 : 2, 3 + i * 3.3f, 14, 3 + i * 3.3f); Arrow(2, 8, 5, 8, pen); break;
            case "indentDec": for (int i = 0; i < 4; i++) g.DrawLine(pen, i is 1 or 2 ? 7 : 2, 3 + i * 3.3f, 14, 3 + i * 3.3f); Arrow(5, 8, 2, 8, pen); break;
            case "condFormat":
                Grid(1.5f, 1.5f, 13, 13, 3, 1, thin);
                g.FillRectangle(new SolidBrush(Red), 2.5f, 2.5f, 9, 3); g.FillRectangle(new SolidBrush(Gold), 2.5f, 6.8f, 6, 3); g.FillRectangle(green, 2.5f, 11, 11, 3);
                break;
            case "table":
                g.FillRectangle(blue, 1.5f, 2, 13, 3);
                Grid(1.5f, 2, 13, 12, 4, 3, thin);
                g.FillRectangle(new SolidBrush(Color.FromArgb(70, Blue)), 2, 8, 12, 3);
                break;
            case "cellStyles":
                g.FillRectangle(new SolidBrush(Color.FromArgb(198, 239, 206)), 1.5f, 2, 6, 5); g.FillRectangle(new SolidBrush(Color.FromArgb(255, 199, 206)), 8.5f, 2, 6, 5);
                g.FillRectangle(new SolidBrush(Color.FromArgb(255, 235, 156)), 1.5f, 9, 6, 5); g.DrawRectangle(thin, 8.5f, 9, 6, 5);
                break;
            case "insertCells": Grid(1.5f, 4.5f, 9, 9, 2, 2); g.DrawLine(greenPen, 12.5f, 1, 12.5f, 7); g.DrawLine(greenPen, 9.5f, 4, 15.5f, 4); break;
            case "deleteCells": Grid(1.5f, 4.5f, 9, 9, 2, 2); g.DrawLine(new Pen(Red, 1.4f), 10.5f, 1.5f, 15, 6); g.DrawLine(new Pen(Red, 1.4f), 15, 1.5f, 10.5f, 6); break;
            case "insertRow": Grid(1.5f, 1.5f, 13, 13, 3, 1); g.FillRectangle(new SolidBrush(Color.FromArgb(120, Green)), 2, 6, 12, 4); break;
            case "insertCol": Grid(1.5f, 1.5f, 13, 13, 1, 3); g.FillRectangle(new SolidBrush(Color.FromArgb(120, Green)), 6, 2, 4, 12); break;
            case "deleteRow": Grid(1.5f, 1.5f, 13, 13, 3, 1); g.FillRectangle(new SolidBrush(Color.FromArgb(120, Red)), 2, 6, 12, 4); break;
            case "deleteCol": Grid(1.5f, 1.5f, 13, 13, 1, 3); g.FillRectangle(new SolidBrush(Color.FromArgb(120, Red)), 6, 2, 4, 12); break;
            case "format": Grid(1.5f, 2.5f, 13, 11, 2, 2); g.FillRectangle(new SolidBrush(Gold), 9, 8.5f, 5, 4.5f); break;
            case "rowHeight": Grid(4.5f, 1.5f, 10, 13, 3, 1); Arrow(2, 8, 2, 2.5f, thin); Arrow(2, 8, 2, 13.5f, thin); break;
            case "colWidth": Grid(1.5f, 4.5f, 13, 10, 1, 3); Arrow(8, 2, 2.5f, 2, thin); Arrow(8, 2, 13.5f, 2, thin); break;
            case "sum": Text("Σ", 2, -2, 15, inkB, FontStyle.Regular); break;
            case "fillDown": Grid(3.5f, 1.5f, 9, 13, 3, 1); Arrow(8, 3, 8, 13, new Pen(Blue, 1.4f)); break;
            case "clear": g.FillPolygon(new SolidBrush(Color.FromArgb(240, 150, 160)), new[] { new PointF(2, 10), new PointF(8, 4), new PointF(13, 9), new PointF(9, 13), new PointF(5, 13) }); g.DrawPolygon(pen, new[] { new PointF(2, 10), new PointF(8, 4), new PointF(13, 9), new PointF(9, 13), new PointF(5, 13) }); g.DrawLine(pen, 1, 14.5f, 15, 14.5f); break;
            case "sortFilter": case "filter":
                g.FillPolygon(inkB, new[] { new PointF(2, 3), new PointF(14, 3), new PointF(9.5f, 8.5f), new PointF(9.5f, 13.5f), new PointF(6.5f, 12), new PointF(6.5f, 8.5f) });
                break;
            case "sortAsc": Text("A", 1, -1, 7.5f, blue); Text("Z", 1, 7, 7.5f, inkB); Arrow(12, 2, 12, 14, pen); break;
            case "sortDesc": Text("Z", 1, -1, 7.5f, inkB); Text("A", 1, 7, 7.5f, blue); Arrow(12, 2, 12, 14, pen); break;
            case "sortCustom": Text("A", 0.5f, -1, 6.5f, blue); Text("Z", 0.5f, 7, 6.5f, inkB); Arrow(8, 2, 8, 14, pen); g.DrawLine(thin, 10, 4, 15, 4); g.DrawLine(thin, 10, 8, 15, 8); g.DrawLine(thin, 10, 12, 15, 12); break;
            case "clearFilter": g.FillPolygon(inkB, new[] { new PointF(1, 2), new PointF(11, 2), new PointF(7.5f, 7), new PointF(7.5f, 12), new PointF(4.5f, 11), new PointF(4.5f, 7) }); g.DrawLine(new Pen(Red, 1.5f), 10, 9, 15, 14); g.DrawLine(new Pen(Red, 1.5f), 15, 9, 10, 14); break;
            case "find": g.DrawEllipse(pen, 2, 2, 8, 8); g.DrawLine(new Pen(ink, 2f), 9, 9, 14, 14); break;
            case "colChart": g.FillRectangle(blue, 2, 8, 3, 6); g.FillRectangle(new SolidBrush(Orange), 6.5f, 4, 3, 10); g.FillRectangle(green, 11, 2, 3, 12); g.DrawLine(thin, 1, 14.5f, 15, 14.5f); break;
            case "barChart": g.FillRectangle(blue, 2, 2, 6, 3); g.FillRectangle(new SolidBrush(Orange), 2, 6.5f, 10, 3); g.FillRectangle(green, 2, 11, 12, 3); g.DrawLine(thin, 1.5f, 1, 1.5f, 15); break;
            case "stackChart": g.FillRectangle(blue, 2, 9, 3, 5); g.FillRectangle(new SolidBrush(Orange), 2, 5, 3, 4); g.FillRectangle(blue, 6.5f, 7, 3, 7); g.FillRectangle(new SolidBrush(Orange), 6.5f, 3, 3, 4); g.FillRectangle(blue, 11, 5, 3, 9); g.FillRectangle(new SolidBrush(Orange), 11, 2, 3, 3); break;
            case "lineChart": g.DrawLines(new Pen(Blue, 1.6f), new[] { new PointF(1, 12), new PointF(5, 7), new PointF(9, 9), new PointF(15, 3) }); g.DrawLines(new Pen(Orange, 1.6f), new[] { new PointF(1, 14), new PointF(5, 11), new PointF(9, 12), new PointF(15, 8) }); break;
            case "pieChart": g.FillPie(blue, 1.5f, 1.5f, 13, 13, -90, 220); g.FillPie(new SolidBrush(Orange), 1.5f, 1.5f, 13, 13, 130, 90); g.FillPie(green, 1.5f, 1.5f, 13, 13, 220, 50); break;
            case "doughnutChart": g.FillPie(blue, 1.5f, 1.5f, 13, 13, -90, 220); g.FillPie(new SolidBrush(Orange), 1.5f, 1.5f, 13, 13, 130, 140); g.FillEllipse(Brushes.White, 5, 5, 6, 6); break;
            case "areaChart": g.FillPolygon(new SolidBrush(Color.FromArgb(180, Blue)), new[] { new PointF(1, 15), new PointF(1, 9), new PointF(6, 5), new PointF(10, 8), new PointF(15, 3), new PointF(15, 15) }); break;
            case "scatterChart": foreach (var (x, y) in new[] { (3f, 11f), (5f, 8f), (8f, 9f), (10f, 5f), (13f, 3f), (12f, 7f) }) g.FillEllipse(blue, x - 1.3f, y - 1.3f, 2.6f, 2.6f); g.DrawLine(thin, 1.5f, 1, 1.5f, 14.5f); g.DrawLine(thin, 1.5f, 14.5f, 15, 14.5f); break;
            case "function": Text("fx", 0.5f, 0, 12, inkB, FontStyle.Italic | FontStyle.Bold); break;
            case "names": g.DrawRectangle(pen, 1.5f, 4.5f, 13, 7); Text("ab", 3, 3.5f, 8, inkB); break;
            case "showFormulas": Text("fx", 0, 0, 9, inkB, FontStyle.Italic); Grid(7.5f, 7.5f, 7, 7, 2, 2); break;
            case "calc": Grid(2.5f, 1.5f, 11, 13, 1, 1, pen); g.FillRectangle(new SolidBrush(Color.FromArgb(200, 230, 210)), 4, 3, 8, 3); for (int i = 0; i < 3; i++) for (int j = 0; j < 3; j++) g.FillRectangle(inkB, 4.2f + j * 2.8f, 7.5f + i * 2.2f, 1.6f, 1.4f); break;
            case "freeze": Grid(1.5f, 1.5f, 13, 13, 3, 3, thin); g.DrawLine(new Pen(Blue, 1.8f), 1.5f, 5.8f, 14.5f, 5.8f); g.DrawLine(new Pen(Blue, 1.8f), 5.8f, 1.5f, 5.8f, 14.5f); break;
            case "gridlines": Grid(1.5f, 1.5f, 13, 13, 3, 3, pen); break;
            case "headings": Grid(1.5f, 1.5f, 13, 13, 3, 3, thin); g.FillRectangle(new SolidBrush(Color.FromArgb(160, Gray)), 2, 2, 12, 3.3f); g.FillRectangle(new SolidBrush(Color.FromArgb(160, Gray)), 2, 2, 3.3f, 12); break;
            case "zoomSel": g.DrawEllipse(pen, 2, 2, 9, 9); g.DrawLine(new Pen(ink, 2f), 10, 10, 14.5f, 14.5f); g.DrawRectangle(thin, 4.5f, 4.5f, 4, 4); break;
            case "removeDup": Grid(1.5f, 1.5f, 9, 13, 4, 1); g.DrawLine(new Pen(Red, 1.5f), 10, 9, 15, 14); g.DrawLine(new Pen(Red, 1.5f), 15, 9, 10, 14); break;
            case "textToCols": g.DrawRectangle(thin, 1.5f, 3.5f, 5, 9); g.DrawRectangle(thin, 9.5f, 3.5f, 5, 9); Arrow(5, 8, 11, 8, pen); break;
            case "validation": g.DrawRectangle(thin, 1.5f, 2.5f, 13, 11); g.DrawLines(new Pen(Green, 1.6f), new[] { new PointF(4, 8), new PointF(6.5f, 10.5f), new PointF(12, 5) }); break;
            case "protect": g.DrawArc(pen, 4.5f, 1.5f, 7, 8, 180, 180); g.FillRectangle(new SolidBrush(Gold), 3, 6, 10, 8.5f); g.DrawRectangle(pen, 3, 6, 10, 8.5f); break;
            case "note": g.FillRectangle(new SolidBrush(Color.FromArgb(255, 245, 180)), 2, 2, 12, 11); g.DrawRectangle(thin, 2, 2, 12, 11); g.FillPolygon(new SolidBrush(Red), new[] { new PointF(10, 2), new PointF(14, 2), new PointF(14, 6) }); g.DrawLine(thin, 4, 6, 11, 6); g.DrawLine(thin, 4, 9, 11, 9); break;
            case "symbol": Text("Ω", 1.5f, -1, 13, inkB, FontStyle.Regular); break;
            case "picture": g.DrawRectangle(pen, 1.5f, 2.5f, 13, 11); g.FillPolygon(green, new[] { new PointF(2, 13), new PointF(6, 8), new PointF(9, 11), new PointF(11, 9), new PointF(14, 13) }); g.FillEllipse(new SolidBrush(Gold), 10, 4, 3, 3); break;
            case "link": g.DrawArc(new Pen(ink, 1.5f), 1.5f, 6, 7, 5, 90, 180); g.DrawArc(new Pen(ink, 1.5f), 7.5f, 6, 7, 5, -90, 180); g.DrawLine(new Pen(ink, 1.5f), 5, 6, 11, 6); g.DrawLine(new Pen(ink, 1.5f), 5, 11, 11, 11); break;
            case "margins": g.DrawRectangle(pen, 2.5f, 1.5f, 11, 13); using (var dp = new Pen(Blue, 0.8f) { DashStyle = DashStyle.Dash }) g.DrawRectangle(dp, 4.5f, 3.5f, 7, 9); break;
            case "orientation": g.DrawRectangle(thin, 1.5f, 1.5f, 7, 10); g.DrawRectangle(pen, 4.5f, 7.5f, 10, 7); break;
            case "paperSize": g.DrawRectangle(pen, 3.5f, 1.5f, 9, 13); Text("A4", 3.6f, 4, 6.5f, inkB); break;
            case "printArea": g.DrawRectangle(thin, 1.5f, 1.5f, 13, 13); using (var dp = new Pen(Blue, 1.2f) { DashStyle = DashStyle.Dash }) g.DrawRectangle(dp, 3.5f, 3.5f, 8, 7); break;
            case "printTitles": Grid(1.5f, 1.5f, 13, 13, 4, 1, thin); g.FillRectangle(new SolidBrush(Color.FromArgb(160, Blue)), 2, 2, 12, 3); break;
            case "fitWidth": g.DrawRectangle(pen, 2.5f, 1.5f, 11, 13); Arrow(8, 8, 3.5f, 8, thin); Arrow(8, 8, 12.5f, 8, thin); break;
            case "scale": Text("%", 3, 0, 12, inkB, FontStyle.Regular); break;
            case "addSheet": g.DrawLine(new Pen(ink, 1.4f), 8, 3, 8, 13); g.DrawLine(new Pen(ink, 1.4f), 3, 8, 13, 8); break;
            case "chartTitle": g.DrawRectangle(thin, 1.5f, 1.5f, 13, 13); g.FillRectangle(inkB, 4, 3.5f, 8, 2); g.FillRectangle(blue, 4, 9, 2, 4); g.FillRectangle(blue, 8, 7, 2, 6); break;
            case "legend": g.DrawRectangle(thin, 1.5f, 1.5f, 13, 13); g.FillRectangle(blue, 3, 4, 3, 2); g.FillRectangle(new SolidBrush(Orange), 3, 9, 3, 2); g.DrawLine(thin, 7.5f, 5, 12.5f, 5); g.DrawLine(thin, 7.5f, 10, 12.5f, 10); break;
            case "dataLabels": g.FillRectangle(blue, 3, 9, 3, 5); g.FillRectangle(blue, 10, 6, 3, 8); Text("5", 2.5f, 2, 6, inkB); Text("8", 9.5f, -0.5f, 6, inkB); break;
            case "switchRowCol": Grid(1.5f, 1.5f, 13, 13, 2, 2, thin); Arrow(4, 11, 11, 4, pen); break;
            case "editData": Grid(1.5f, 1.5f, 10, 10, 2, 2, thin); g.DrawLine(new Pen(Orange, 2f), 7, 14, 14, 7); break;
            case "deleteObj": g.DrawLine(new Pen(Red, 1.8f), 3, 3, 13, 13); g.DrawLine(new Pen(Red, 1.8f), 13, 3, 3, 13); break;
            case "options": g.DrawEllipse(pen, 4, 4, 8, 8); g.FillEllipse(inkB, 6.5f, 6.5f, 3, 3); for (int i = 0; i < 8; i++) { var a = i * Math.PI / 4; g.DrawLine(new Pen(ink, 1.8f), 8 + 4 * (float)Math.Cos(a), 8 + 4 * (float)Math.Sin(a), 8 + 6.5f * (float)Math.Cos(a), 8 + 6.5f * (float)Math.Sin(a)); } break;
            case "newSheetTab": g.DrawRectangle(pen, 2.5f, 3.5f, 11, 9); g.DrawLine(greenPen, 8, 5, 8, 11); g.DrawLine(greenPen, 5, 8, 11, 8); break;
            case "formatPainter": g.FillRectangle(new SolidBrush(Gold), 2, 2, 10, 4); g.DrawRectangle(pen, 2, 2, 10, 4); g.DrawLine(pen, 12, 4, 14, 4); g.DrawLine(pen, 14, 4, 14, 8); g.DrawLine(pen, 14, 8, 8, 8); g.FillRectangle(inkB, 7, 8, 2, 7); break;
            default:
                g.DrawRectangle(pen, 2.5f, 2.5f, 11, 11);
                break;
        }
    }
}
