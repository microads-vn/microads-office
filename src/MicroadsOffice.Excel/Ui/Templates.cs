using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Format;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.UI;

namespace MicroadsOffice.Excel.Ui;

/// <summary>Built-in workbook templates.</summary>
public static class ExcelTemplates
{
    public static IReadOnlyList<DocTemplate> All => new List<DocTemplate>
    {
        new("blank", Loc.T("Sổ làm việc trống", "Blank workbook"), (w, h) => Thumb(Create("blank"), w, h, blank: true)),
        new("budget", Loc.T("Ngân sách cá nhân", "Personal budget"), (w, h) => Thumb(Create("budget"), w, h)),
        new("invoice", Loc.T("Hóa đơn / Báo giá", "Invoice / Quote"), (w, h) => Thumb(Create("invoice"), w, h)),
        new("timesheet", Loc.T("Bảng chấm công", "Timesheet"), (w, h) => Thumb(Create("timesheet"), w, h)),
        new("list", Loc.T("Danh sách liên hệ", "Contact list"), (w, h) => Thumb(Create("list"), w, h)),
    };

    private static int C(int r, int g, int b) => Color.FromArgb(r, g, b).ToArgb();

    public static Workbook Create(string id)
    {
        var wb = Workbook.CreateBlank();
        var s = wb.Sheets[0];
        int St(CellStyle st) => wb.Styles.Intern(st);
        void Put(int r, int c, string text, int style = 0) { SheetOps.SetInput(s, r, c, text); if (style != 0) s.SetStyle(r, c, style); }
        void StyleRange(CellRange rg, int style) { foreach (var (r, c) in rg.Cells()) s.SetStyle(r, c, style); }
        var green = C(33, 115, 70);
        var title = St(CellStyle.Default with { FontSize = 18, Bold = true, FontColor = green });
        var head = St(CellStyle.Default with { Bold = true, Fill = green, FontColor = C(255, 255, 255), HAlign = HAlign.Center, VAlign = VAlign.Center, Wrap = true, Left = BorderLine.Thin(Color.FromArgb(200, 200, 200)), Right = BorderLine.Thin(Color.FromArgb(200, 200, 200)) });
        var cellB = St(CellStyle.Default with { Left = BorderLine.Thin(Color.FromArgb(200, 200, 200)), Right = BorderLine.Thin(Color.FromArgb(200, 200, 200)), Top = BorderLine.Thin(Color.FromArgb(200, 200, 200)), Bottom = BorderLine.Thin(Color.FromArgb(200, 200, 200)) });
        var money = St(wb.Styles[cellB] with { NumberFormat = "#,##0" });
        var totalMoney = St(wb.Styles[money] with { Bold = true, Fill = C(226, 239, 218) });
        var totalLbl = St(wb.Styles[cellB] with { Bold = true, Fill = C(226, 239, 218) });
        switch (id)
        {
            case "budget":
            {
                s.Name = Loc.T("Ngân sách", "Budget");
                Put(0, 0, Loc.T("NGÂN SÁCH CÁ NHÂN THÁNG", "MONTHLY PERSONAL BUDGET"), title);
                Put(1, 0, Loc.T("Tháng:", "Month:")); Put(1, 1, DateTime.Today.ToString("MM/yyyy"));
                Put(3, 0, Loc.T("Thu nhập", "Income"), head); Put(3, 1, Loc.T("Số tiền (₫)", "Amount"), head);
                string[] inc = Loc.IsVi ? ["Lương", "Thưởng", "Thu nhập khác"] : ["Salary", "Bonus", "Other income"];
                double[] incV = [15000000, 2000000, 1000000];
                for (int i = 0; i < 3; i++) { Put(4 + i, 0, inc[i], cellB); s.SetValue(4 + i, 1, CellValue.Number(incV[i])); s.SetStyle(4 + i, 1, money); }
                Put(7, 0, Loc.T("Tổng thu", "Total income"), totalLbl); Put(7, 1, "=SUM(B5:B7)", totalMoney);
                Put(9, 0, Loc.T("Chi tiêu", "Expenses"), head); Put(9, 1, Loc.T("Số tiền (₫)", "Amount"), head);
                string[] exp = Loc.IsVi ? ["Nhà ở", "Ăn uống", "Đi lại", "Điện, nước, Internet", "Giáo dục", "Giải trí", "Tiết kiệm"] : ["Housing", "Food", "Transport", "Utilities", "Education", "Entertainment", "Savings"];
                double[] expV = [5000000, 4000000, 1200000, 900000, 1500000, 800000, 3000000];
                for (int i = 0; i < exp.Length; i++) { Put(10 + i, 0, exp[i], cellB); s.SetValue(10 + i, 1, CellValue.Number(expV[i])); s.SetStyle(10 + i, 1, money); }
                Put(17, 0, Loc.T("Tổng chi", "Total expenses"), totalLbl); Put(17, 1, "=SUM(B11:B17)", totalMoney);
                Put(19, 0, Loc.T("Còn lại", "Balance"), totalLbl); Put(19, 1, "=B8-B18", totalMoney);
                s.ColWidths[0] = 190; s.ColWidths[1] = 130;
                s.CondFormats.Add(new ConditionalFormat { Range = new CellRange(10, 1, 16, 1), Type = CfType.DataBar, Color1 = C(99, 190, 123) });
                s.Charts.Add(new Chart { Type = ChartType.Doughnut, Title = Loc.T("Cơ cấu chi tiêu", "Spending breakdown"), Data = new CellRange(9, 0, 16, 1), Bounds = new RectangleF(360, 60, 400, 300), ShowDataLabels = true, Name = "Chart 1" });
                break;
            }
            case "invoice":
            {
                s.Name = Loc.T("Hóa đơn", "Invoice");
                s.ShowGridlines = false;
                Put(0, 0, Loc.T("CÔNG TY CỦA BẠN", "YOUR COMPANY"), St(CellStyle.Default with { Bold = true, FontSize = 13, FontColor = green }));
                Put(1, 0, Loc.T("Địa chỉ: ...   Điện thoại: ...   Email: ...", "Address: ...   Phone: ...   Email: ..."));
                Put(3, 0, Loc.T("HÓA ĐƠN / BÁO GIÁ", "INVOICE / QUOTATION"), St(CellStyle.Default with { Bold = true, FontSize = 18, HAlign = HAlign.Center }));
                s.Merges.Add(new CellRange(3, 0, 3, 5));
                Put(5, 0, Loc.T("Khách hàng:", "Customer:"), St(CellStyle.Default with { Bold = true })); Put(5, 2, Loc.T("Công ty ABC", "ABC Company"));
                Put(6, 0, Loc.T("Ngày:", "Date:"), St(CellStyle.Default with { Bold = true })); Put(6, 2, "=TODAY()", St(CellStyle.Default with { NumberFormat = "dd/mm/yyyy", HAlign = HAlign.Left }));
                s.Merges.Add(new CellRange(5, 0, 5, 1)); s.Merges.Add(new CellRange(6, 0, 6, 1)); s.Merges.Add(new CellRange(5, 2, 5, 5)); s.Merges.Add(new CellRange(6, 2, 6, 5));
                string[] h = Loc.IsVi ? ["STT", "Tên hàng hóa, dịch vụ", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"] : ["No.", "Description", "Unit", "Qty", "Unit price", "Amount"];
                for (int c = 0; c < 6; c++) Put(8, c, h[c], head);
                s.RowHeights[8] = 30;
                string[] items = Loc.IsVi ? ["Thiết kế website", "Phần mềm quản lý (bản quyền 1 năm)", "Bảo trì hệ thống"] : ["Website design", "Management software (1-year licence)", "System maintenance"];
                string[] units = Loc.IsVi ? ["Gói", "Bộ", "Tháng"] : ["Package", "Set", "Month"];
                double[] qty = [1, 2, 6], price = [15000000, 8500000, 1200000];
                var center = St(wb.Styles[cellB] with { HAlign = HAlign.Center });
                for (int i = 0; i < 6; i++)
                {
                    int r = 9 + i;
                    if (i < 3)
                    {
                        Put(r, 0, (i + 1).ToString(), center); Put(r, 1, items[i], cellB); Put(r, 2, units[i], center);
                        s.SetValue(r, 3, CellValue.Number(qty[i])); s.SetStyle(r, 3, center);
                        s.SetValue(r, 4, CellValue.Number(price[i])); s.SetStyle(r, 4, money);
                    }
                    else { foreach (var c in new[] { 0, 2, 3 }) s.SetStyle(r, c, center); s.SetStyle(r, 1, cellB); s.SetStyle(r, 4, money); }
                    Put(r, 5, $"=IF(D{r + 1}=\"\",\"\",D{r + 1}*E{r + 1})", money);
                }
                Put(15, 4, Loc.T("Cộng tiền hàng", "Subtotal"), totalLbl); Put(15, 5, "=SUM(F10:F15)", totalMoney);
                Put(16, 4, Loc.T("Thuế GTGT 10%", "VAT 10%"), totalLbl); Put(16, 5, "=ROUND(F16*10%,0)", totalMoney);
                Put(17, 4, Loc.T("TỔNG THANH TOÁN", "TOTAL"), totalLbl); Put(17, 5, "=F16+F17", St(wb.Styles[totalMoney] with { FontColor = C(192, 0, 0) }));
                Put(19, 0, "=\"" + Loc.T("Số tiền bằng chữ: ", "Amount in words: ") + "\"&DOCSO(F18)", St(CellStyle.Default with { Italic = true, Bold = true }));
                s.Merges.Add(new CellRange(19, 0, 19, 5));
                Put(21, 1, Loc.T("Người mua hàng", "Buyer"), St(CellStyle.Default with { Bold = true, HAlign = HAlign.Center }));
                Put(21, 4, Loc.T("Người bán hàng", "Seller"), St(CellStyle.Default with { Bold = true, HAlign = HAlign.Center }));
                s.Merges.Add(new CellRange(21, 4, 21, 5));
                s.ColWidths[0] = 50; s.ColWidths[1] = 260; s.ColWidths[2] = 70; s.ColWidths[3] = 75; s.ColWidths[4] = 140; s.ColWidths[5] = 130;
                s.Page.FitToWidth = true;
                break;
            }
            case "timesheet":
            {
                s.Name = Loc.T("Chấm công", "Timesheet");
                Put(0, 0, Loc.T("BẢNG CHẤM CÔNG THÁNG ", "TIMESHEET FOR ") + DateTime.Today.ToString("MM/yyyy"), title);
                Put(1, 0, Loc.T("Ký hiệu: x = đi làm, p = nghỉ phép, o = nghỉ ốm", "Legend: x = worked, p = paid leave, o = sick leave"), St(CellStyle.Default with { Italic = true, FontColor = C(90, 90, 90) }));
                Put(3, 0, Loc.T("STT", "No."), head); Put(3, 1, Loc.T("Họ và tên", "Name"), head);
                int days = DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month);
                var dayHead = St(wb.Styles[head] with { FontSize = 9 });
                var weekend = St(wb.Styles[cellB] with { Fill = C(242, 242, 242), HAlign = HAlign.Center });
                var mark = St(wb.Styles[cellB] with { HAlign = HAlign.Center });
                for (int d = 1; d <= days; d++)
                {
                    var date = new DateTime(DateTime.Today.Year, DateTime.Today.Month, d);
                    Put(3, 1 + d, d.ToString(), dayHead);
                    s.ColWidths[1 + d] = 26;
                }
                Put(3, days + 2, Loc.T("Công", "Days"), head); Put(3, days + 3, Loc.T("Phép", "Leave"), head);
                string[] names = Loc.IsVi ? ["Nguyễn Văn An", "Trần Thị Bình", "Lê Hoàng Cường", "Phạm Thu Dung", "Hoàng Minh Đức"] : ["Alex Nguyen", "Bella Tran", "Chris Le", "Diana Pham", "Ethan Hoang"];
                for (int i = 0; i < names.Length; i++)
                {
                    int r = 4 + i;
                    Put(r, 0, (i + 1).ToString(), mark); Put(r, 1, names[i], cellB);
                    for (int d = 1; d <= days; d++)
                    {
                        var date = new DateTime(DateTime.Today.Year, DateTime.Today.Month, d);
                        bool we = date.DayOfWeek is DayOfWeek.Sunday;
                        if (!we && d <= DateTime.Today.Day) Put(r, 1 + d, (i + d) % 13 == 0 ? "p" : "x", mark);
                        s.SetStyle(r, 1 + d, we ? weekend : mark);
                    }
                    string last = Addr.ColName(1 + days);
                    Put(r, days + 2, $"=COUNTIF(C{r + 1}:{last}{r + 1},\"x\")", St(wb.Styles[mark] with { Bold = true }));
                    Put(r, days + 3, $"=COUNTIF(C{r + 1}:{last}{r + 1},\"p\")", mark);
                }
                s.ColWidths[0] = 40; s.ColWidths[1] = 170;
                s.FreezeRows = 4; s.FreezeCols = 2;
                s.Page.Landscape = true; s.Page.FitToWidth = true;
                break;
            }
            case "list":
            {
                s.Name = Loc.T("Danh bạ", "Contacts");
                string[] h = Loc.IsVi ? ["Họ và tên", "Công ty", "Chức vụ", "Điện thoại", "Email", "Ghi chú"] : ["Full name", "Company", "Title", "Phone", "Email", "Notes"];
                for (int c = 0; c < h.Length; c++) Put(0, c, h[c], head);
                string[,] data = Loc.IsVi
                    ? new[,] { { "Nguyễn Văn An", "Công ty ABC", "Giám đốc", "'0901234567", "an@abc.vn", "" }, { "Trần Thị Bình", "Công ty XYZ", "Kế toán", "'0912345678", "binh@xyz.vn", "Khách hàng mới" }, { "Lê Hoàng Cường", "Microads", "Kỹ sư", "'0987654321", "cuong@microads.vn", "" } }
                    : new[,] { { "Alex Nguyen", "ABC Ltd", "Director", "'0901234567", "alex@abc.com", "" }, { "Bella Tran", "XYZ Corp", "Accountant", "'0912345678", "bella@xyz.com", "New customer" }, { "Chris Le", "Microads", "Engineer", "'0987654321", "chris@microads.vn", "" } };
                for (int r = 0; r < data.GetLength(0); r++) for (int c = 0; c < h.Length; c++) Put(r + 1, c, data[r, c], cellB);
                s.ColWidths[0] = 160; s.ColWidths[1] = 140; s.ColWidths[2] = 110; s.ColWidths[3] = 110; s.ColWidths[4] = 180; s.ColWidths[5] = 160;
                s.FreezeRows = 1;
                s.Filter = new AutoFilter { Range = new CellRange(0, 0, 3, 5) };
                break;
            }
        }
        wb.Recalculate();
        s.ActiveCell = new CellPos(0, 0);
        s.Selection = CellRange.Single(0, 0);
        return wb;
    }

    private static Bitmap Thumb(Workbook wb, int w, int h, bool blank = false)
    {
        var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.White);
        var s = wb.Sheets[0];
        float scale = 0.42f;
        var rows = new List<CellPainter.Line>();
        float y = 0;
        for (int r = 0; y < h; r++) { float rh = (float)s.RowHeight(r) * scale; rows.Add(new CellPainter.Line(r, y, rh)); y += rh; }
        var cols = new List<CellPainter.Line>();
        float x = 0;
        for (int c = 0; x < w; c++) { float cw = (float)s.ColWidth(c) * scale; cols.Add(new CellPainter.Line(c, x, cw)); x += cw; }
        if (blank)
        {
            using var p = new Pen(Color.FromArgb(225, 227, 230));
            foreach (var r in rows) g.DrawLine(p, 0, r.Pos + r.Size, w, r.Pos + r.Size);
            foreach (var c in cols) g.DrawLine(p, c.Pos + c.Size, 0, c.Pos + c.Size, h);
            using var acc = new Pen(Theme.Accent, 2);
            g.DrawRectangle(acc, cols[1].Pos, rows[1].Pos, cols[1].Size, rows[1].Size);
            return bmp;
        }
        CellPainter.Paint(g, s, rows, cols, scale, scale, s.ShowGridlines, false, false, new CfEvaluator(s), new RectangleF(0, 0, w, h));
        foreach (var ch in s.Charts)
            ChartRenderer.Draw(g, new RectangleF(ch.Bounds.X * scale, ch.Bounds.Y * scale, ch.Bounds.Width * scale, ch.Bounds.Height * scale), s, ch, scale);
        return bmp;
    }
}
