using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using MicroadsOffice.AI.UI;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Excel.Ui;
using MicroadsOffice.Word.Model;
using MicroadsOffice.Word.UI;

namespace MicroadsOffice.AI.Tests;

/// <summary>
/// Microsoft Store listing screenshots (1920×1080) taken from the real app windows with a scripted AI answer, plus the
/// square Store logos. usage: MicroadsOffice.AI.Tests shots-word|shots-excel outDir
/// </summary>
public static class StoreShots
{
    const int W = 1920, H = 1080;

    const string Contract = """
        Tôi đã soạn hợp đồng lao động mẫu theo Bộ luật Lao động 2019, các thông tin chưa biết để trống (……).
        ```microads-document mode=replace_document
        :::center
        **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM**
        **Độc lập - Tự do - Hạnh phúc**
        :::

        # HỢP ĐỒNG LAO ĐỘNG
        :::center
        *Số: ……/……/HĐLĐ*
        :::

        *Căn cứ Bộ luật Lao động số 45/2019/QH14 ngày 20 tháng 11 năm 2019;*
        *Căn cứ nhu cầu sử dụng lao động và năng lực của người lao động;*

        Hôm nay, ngày …… tháng …… năm 20……, tại Thành phố Hồ Chí Minh, chúng tôi gồm:

        ## BÊN A: NGƯỜI SỬ DỤNG LAO ĐỘNG
        - Tên doanh nghiệp: ……………………………………
        - Địa chỉ trụ sở: ……………………………………
        - Mã số thuế: …………………  Điện thoại: …………………
        - Đại diện bởi: Ông/Bà ………………… Chức vụ: …………………

        ## BÊN B: NGƯỜI LAO ĐỘNG
        - Họ và tên: …………………  Ngày sinh: …………………
        - Số CCCD: …………………  Nơi cư trú: …………………

        ## Điều 1. Công việc, địa điểm và thời hạn
        | Nội dung | Thỏa thuận |
        |---|---|
        | Chức danh / công việc | …………………………… |
        | Địa điểm làm việc | …………………………… |
        | Loại hợp đồng | Xác định thời hạn …… tháng |

        ## Điều 2. Chế độ làm việc
        Thời giờ làm việc 8 giờ/ngày, 40 giờ/tuần; nghỉ hằng năm theo quy định của pháp luật.
        ```
        """;

    public static void RunWord(Runner r, string outDir)
    {
        var closer = Ui.StartCloser();
        var ai = new ScriptedProvider { DelayMs = 1 };
        AiService.Override = _ => ai;
        AiSettings.Update(x => { x.Provider = AiProviderKind.ChatGpt; x.ChatGptModel = ""; /* no invented model name in a Store screenshot */ });
        WordMainForm? w = null;
        try
        {
            r.Test("S01 Word + AI assistant drafting a contract (1920×1080)", () =>
            {
                w = (WordMainForm)Word.WordApp.CreateMainForm([]);
                w.StartPosition = FormStartPosition.Manual; w.Location = new Point(0, 0); w.Size = new Size(W, H);
                w.Show(); Ui.Pump(400); w.HideBackstage(); w.Size = new Size(W, H); Ui.Pump(200);
                w.ShowAi(true); Ui.Pump(200);
                ai.Replies.Enqueue(Contract);
                w.AiPane!.Ask("Soạn hợp đồng lao động mẫu cho công ty");
                Ui.WaitRequest(w.AiPane);
                w.Editor.SetCaret(new DocPos(0, 0), false);
                Ui.Shot(w, "01-word-ai");
                Move(outDir, "01-word-ai");
            });

            r.Test("S03 AI settings: Continue with ChatGPT, Claude, API keys, own server", () =>
            {
                using var d = new AiSettingsDialog(AiSettingsDialog.Section.ChatGpt);
                d.StartPosition = FormStartPosition.Manual;
                d.Show(w); Ui.Pump(200);
                d.Location = new Point(w!.Left + (W - d.Width) / 2, w.Top + (H - d.Height) / 2);
                Ui.Pump(300);
                using var back = Capture(w);
                using var front = Capture(d);
                using var g = Graphics.FromImage(back);
                using (var dim = new SolidBrush(Color.FromArgb(70, 0, 0, 0))) g.FillRectangle(dim, 0, 0, back.Width, back.Height);
                g.DrawImage(front, (back.Width - front.Width) / 2, (back.Height - front.Height) / 2);
                back.Save(Path.Combine(outDir, "03-ai-settings.png"), ImageFormat.Png);
                d.Close();
            });

            r.Test("S05 square Store logos (300, 1080) and poster (720×1080)", () =>
            {
                var logos = Path.Combine(outDir, "..", "logos");
                Directory.CreateDirectory(logos);
                foreach (var s in new[] { 300, 1080 }) Logo(s, s, Path.Combine(logos, $"StoreLogo-{s}.png"));
                Logo(720, 1080, Path.Combine(logos, "Poster-720x1080.png"));
            });
        }
        finally
        {
            AiService.Override = null;
            try { foreach (var x in Application.OpenForms.OfType<WordMainForm>().ToList()) { x.Editor.MarkSaved(); x.Close(); } Ui.Pump(300); } catch { }
            closer.Stop();
        }
    }

    public static void RunExcel(Runner r, string outDir)
    {
        var closer = Ui.StartCloser();
        var ai = new ScriptedProvider { DelayMs = 1 };
        AiService.Override = _ => ai;
        AiSettings.Update(x => { x.Provider = AiProviderKind.ClaudeKey; x.ClaudeModel = "claude-opus-5-5"; });
        ExcelForm? f = null;
        try
        {
            r.Test("S02 Excel + AI assistant building a revenue table and chart (1920×1080)", () =>
            {
                f = (ExcelForm)Excel.ExcelApp.CreateMainForm([]);
                f.StartPosition = FormStartPosition.Manual; f.Location = new Point(0, 0); f.Size = new Size(W, H);
                f.Show(); Ui.Pump(400); f.HideBackstage(); f.Size = new Size(W, H); Ui.Pump(200);
                f.ShowAi(true); Ui.Pump(200);
                ai.Replies.Enqueue("""
                    Đã tạo bảng doanh thu theo chi nhánh, cột Tổng và dòng Cộng dùng công thức, kèm biểu đồ cột.
                    ```microads-sheet
                    {"actions":[
                      {"range":"A1","values":[["Chi nhánh","Quý 1","Quý 2","Quý 3","Quý 4","Tổng"],
                        ["Hà Nội",1250,1380,1420,1610,"=SUM(B2:E2)"],["TP. Hồ Chí Minh",2310,2450,2600,2890,"=SUM(B3:E3)"],
                        ["Đà Nẵng",680,720,790,860,"=SUM(B4:E4)"],["Cần Thơ",410,455,470,530,"=SUM(B5:E5)"],
                        ["Cộng","=SUM(B2:B5)","=SUM(C2:C5)","=SUM(D2:D5)","=SUM(E2:E5)","=SUM(F2:F5)"]]},
                      {"range":"A1:F1","bold":true,"fill":"#1F6F43","fontColor":"#FFFFFF","align":"center","border":true},
                      {"range":"A6:F6","bold":true,"fill":"#E2EFDA","border":true},
                      {"range":"A2:F5","border":true},
                      {"range":"B2:F6","numberFormat":"#,##0"},
                      {"autofit":"A:F"},
                      {"chart":{"type":"column","range":"A1:E5","title":"Doanh thu theo quý (triệu đồng)","at":"H2"}}
                    ]}
                    ```
                    """);
                f.AiPane!.Ask("Tạo bảng doanh thu 4 chi nhánh theo quý, có tổng và biểu đồ");
                Ui.WaitRequest(f.AiPane);
                f.Grid.Select(0, 0);
                Ui.Shot(f, "02-excel-ai");
                Move(outDir, "02-excel-ai");
            });

            r.Test("S04 Excel: explain a formula", () =>
            {
                f!.Grid.Select(1, 5);
                ai.Replies.Enqueue("**=SUM(B2:E2)** cộng doanh thu 4 quý của chi nhánh Hà Nội (ô B2 đến E2): 1.250 + 1.380 + 1.420 + 1.610 = **5.660** triệu đồng.\n\nKhi bạn sửa số của một quý, ô Tổng tự tính lại.");
                Ui.Click(Ui.Find<Button>(f.AiPane!, "btnAiQ_Explain"));
                Ui.WaitRequest(f.AiPane!);
                Ui.Shot(f, "04-excel-explain");
                Move(outDir, "04-excel-explain");
            });
        }
        finally
        {
            AiService.Override = null;
            try { if (f != null) { foreach (var d in f.Docs) d.Modified = false; f.Close(); } Ui.Pump(300); } catch { }
            closer.Stop();
        }
    }

    static void Move(string outDir, string name)
    {
        var src = Path.Combine(T.Out, name + ".png");
        var dst = Path.Combine(outDir, name + ".png");
        if (!string.Equals(Path.GetFullPath(src), Path.GetFullPath(dst), StringComparison.OrdinalIgnoreCase)) File.Copy(src, dst, true);
        using var img = Image.FromFile(dst);
        T.True(img.Width >= 1366 && img.Height >= 768, $"{name}: {img.Width}×{img.Height} is at least 1366×768");
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);

    static Bitmap Capture(Form f)
    {
        var bmp = new Bitmap(f.Width, f.Height);
        using var g = Graphics.FromImage(bmp);
        var hdc = g.GetHdc();
        PrintWindow(f.Handle, hdc, 2);
        g.ReleaseHdc(hdc);
        return bmp;
    }

    static void Logo(int w, int h, string path)
    {
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias; g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.Clear(Color.White);
            int icon = (int)(Math.Min(w, h) * (h > w ? 0.5 : 0.62));
            using var art = BrandArt.RenderSuiteIcon(icon);
            int top = h > w ? (int)(h * 0.22) : (h - icon) / 2;
            g.DrawImage(art, (w - icon) / 2, top, icon, icon);
            if (h > w)
            {
                // pixel sizes, shrunk until the name fits 90% of the width (a point size overflowed: "Microads Off")
                float size = w / 9f;
                Font font = new("Segoe UI Semibold", size, GraphicsUnit.Pixel);
                while (g.MeasureString("Microads Office", font).Width > w * 0.9f && size > 8) { font.Dispose(); size -= 2; font = new Font("Segoe UI Semibold", size, GraphicsUnit.Pixel); }
                using var _font = font;
                using var sub = new Font("Segoe UI", size * 0.45f, GraphicsUnit.Pixel);
                var fmt = new StringFormat { Alignment = StringAlignment.Center };
                g.DrawString("Microads Office", font, Brushes.Black, new RectangleF(0, top + icon + h * 0.05f, w, h * 0.12f), fmt);
                g.DrawString("Word · Excel · PDF · AI", sub, Brushes.DimGray, new RectangleF(0, top + icon + h * 0.16f, w, h * 0.08f), fmt);
            }
        }
        bmp.Save(path, ImageFormat.Png);
    }
}
