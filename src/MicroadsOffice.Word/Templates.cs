using System.Drawing;
using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word;

/// <summary>Built-in document templates (blank, Vietnamese administrative letter, letter, CV, report).</summary>
public static class Templates
{
    public static IReadOnlyList<DocTemplate> All => new[]
    {
        T("blank", Loc.T("Tài liệu trống", "Blank document")),
        T("congvan", Loc.T("Công văn hành chính (NĐ 30/2020)", "VN administrative letter")),
        T("letter", Loc.T("Thư", "Letter")),
        T("cv", Loc.T("Sơ yếu lý lịch", "Resume / CV")),
        T("report", Loc.T("Báo cáo", "Report")),
    };

    private static DocTemplate T(string id, string title) => new(id, title, (w, h) => Thumbnail(Create(id), w, h));

    public static Document Create(string id) => id switch
    {
        "congvan" => CongVan(),
        "letter" => Letter(),
        "cv" => Cv(),
        "report" => Report(),
        _ => Document.CreateBlank()
    };

    /// <summary>Renders the first page of a document into a thumbnail.</summary>
    public static Bitmap Thumbnail(Document d, int w, int h)
    {
        using var m = new TextMeasurer();
        var layout = new DocLayout(d, m);
        var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.White);
        float s = Math.Min(w / d.Page.Width, h / d.Page.Height);
        PageRenderer.Setup(g, 0, 0, s);
        PageRenderer.DrawPage(g, layout, 0, new RenderOptions { Print = true });
        return bmp;
    }

    private static (CharProps C, ParaProps P) N(Document d) => (d.Styles[StyleIds.Normal].Char, d.Styles[StyleIds.Normal].Para);

    private static Paragraph P(Document d, string text, Func<CharProps, CharProps>? c = null, Func<ParaProps, ParaProps>? p = null)
    {
        var (nc, np) = N(d);
        return new Paragraph(text, c != null ? c(nc) : nc, p != null ? p(np) : np);
    }

    private static Paragraph Styled(Document d, string style, string text)
    {
        var s = d.Styles[style];
        return new Paragraph(text, s.Char, s.Para);
    }

    private static TableCell Cell(params Paragraph[] ps) => new() { Paragraphs = ps.ToList() };

    private static Document CongVan()
    {
        var d = Document.CreateBlank("Times New Roman", 14);
        d.Blocks.Clear();
        d.Page.MarginTop = PageSetup.Mm(20); d.Page.MarginBottom = PageSetup.Mm(20);
        d.Page.MarginLeft = PageSetup.Mm(30); d.Page.MarginRight = PageSetup.Mm(15);
        ParaProps C(ParaProps p) => p with { Align = ParaAlign.Center, SpaceAfter = 0, LineSpacing = 1f };
        var header = new Table { Borders = false, ColumnWidths = [PageSetup.Mm(65), d.Page.ContentWidth - PageSetup.Mm(65)] };
        header.Rows.Add(new TableRow
        {
            Cells =
            [
                Cell(P(d, "TÊN CQ, TC CHỦ QUẢN", c => c with { Size = 13 }, C), P(d, "TÊN CƠ QUAN, TỔ CHỨC", c => c with { Size = 13, Bold = true }, C),
                     P(d, "_______", c => c with { Size = 13 }, C), P(d, "Số:     /…-…", c => c with { Size = 13 }, C), P(d, "V/v ……………………", c => c with { Size = 12 }, C)),
                Cell(P(d, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", c => c with { Size = 13, Bold = true }, C), P(d, "Độc lập - Tự do - Hạnh phúc", c => c with { Size = 14, Bold = true }, C),
                     P(d, "_______________________", c => c with { Size = 13 }, C),
                     P(d, "Hà Nội, ngày      tháng      năm 20…", c => c with { Size = 14, Italic = true }, p => C(p) with { SpaceBefore = 6 }))
            ]
        });
        d.Blocks.Add(header);
        d.Blocks.Add(P(d, ""));
        d.Blocks.Add(P(d, "Kính gửi: ……………………………………", null, p => p with { Align = ParaAlign.Center, SpaceAfter = 12 }));
        ParaProps Body(ParaProps p) => p with { Align = ParaAlign.Justify, FirstLine = PageSetup.Mm(10), SpaceAfter = 6, LineSpacing = 1.15f };
        d.Blocks.Add(P(d, "Nội dung công văn trình bày tại đây: phông chữ Times New Roman cỡ 13-14, căn đều hai bên, thụt đầu dòng 1 cm, khoảng cách giữa các đoạn tối thiểu 6pt.", null, Body));
        d.Blocks.Add(P(d, "……………………………………………………………………………………………………………………", null, Body));
        d.Blocks.Add(P(d, "Trân trọng./.", null, Body));
        d.Blocks.Add(P(d, ""));
        var sign = new Table { Borders = false, ColumnWidths = [d.Page.ContentWidth / 2, d.Page.ContentWidth / 2] };
        sign.Rows.Add(new TableRow
        {
            Cells =
            [
                Cell(P(d, "Nơi nhận:", c => c with { Size = 12, Bold = true, Italic = true }, p => p with { SpaceAfter = 0, LineSpacing = 1f }),
                     P(d, "- Như trên;", c => c with { Size = 11 }, p => p with { SpaceAfter = 0, LineSpacing = 1f }),
                     P(d, "- ……;", c => c with { Size = 11 }, p => p with { SpaceAfter = 0, LineSpacing = 1f }),
                     P(d, "- Lưu: VT, … .", c => c with { Size = 11 }, p => p with { SpaceAfter = 0, LineSpacing = 1f })),
                Cell(P(d, "QUYỀN HẠN, CHỨC VỤ CỦA NGƯỜI KÝ", c => c with { Bold = true }, C), P(d, "(Chữ ký của người có thẩm quyền,", c => c with { Italic = true }, C),
                     P(d, "dấu/chữ ký số của cơ quan, tổ chức)", c => c with { Italic = true }, C), P(d, "", null, C), P(d, "", null, C),
                     P(d, "Họ và tên", c => c with { Bold = true }, C))
            ]
        });
        d.Blocks.Add(sign);
        d.Title = "Công văn";
        d.StructureChanged();
        return d;
    }

    private static Document Letter()
    {
        var d = Document.CreateBlank("Times New Roman", 13);
        d.Blocks.Clear();
        d.Blocks.Add(P(d, Loc.T("[Tên của bạn]", "[Your name]"), c => c with { Bold = true, Size = 16 }, p => p with { SpaceAfter = 0 }));
        d.Blocks.Add(P(d, Loc.T("[Địa chỉ] · [Điện thoại] · [Email]", "[Address] · [Phone] · [Email]"), c => c with { Color = Color.FromArgb(89, 89, 89) }, p => p with { SpaceAfter = 18 }));
        d.Blocks.Add(P(d, DateTime.Now.ToString(Loc.IsVi ? "'Ngày' dd 'tháng' MM 'năm' yyyy" : "MMMM d, yyyy", Loc.IsVi ? new System.Globalization.CultureInfo("vi-VN") : System.Globalization.CultureInfo.InvariantCulture), null, p => p with { SpaceAfter = 18 }));
        d.Blocks.Add(P(d, Loc.T("[Tên người nhận]", "[Recipient name]"), c => c with { Bold = true }, p => p with { SpaceAfter = 0 }));
        d.Blocks.Add(P(d, Loc.T("[Chức danh] · [Công ty] · [Địa chỉ]", "[Title] · [Company] · [Address]"), null, p => p with { SpaceAfter = 18 }));
        d.Blocks.Add(P(d, Loc.T("Kính gửi [Tên người nhận],", "Dear [Recipient name],")));
        d.Blocks.Add(P(d, Loc.T("Nhập nội dung thư tại đây. Bấm vào đoạn văn mẫu bất kỳ và gõ để thay thế. Bạn có thể đổi phông chữ, màu sắc và bố cục trên thẻ Trang chủ và Bố cục.",
            "Type your letter here. Click any placeholder text and start typing to replace it. Change fonts, colours and layout on the Home and Layout tabs."), null, p => p with { Align = ParaAlign.Justify }));
        d.Blocks.Add(P(d, Loc.T("Trân trọng,", "Sincerely,"), null, p => p with { SpaceBefore = 12, SpaceAfter = 36 }));
        d.Blocks.Add(P(d, Loc.T("[Tên của bạn]", "[Your name]"), c => c with { Bold = true }));
        d.StructureChanged();
        return d;
    }

    private static Document Cv()
    {
        var d = Document.CreateBlank("Calibri", 11);
        d.Blocks.Clear();
        d.Page.MarginLeft = d.Page.MarginRight = PageSetup.Mm(20);
        var accent = Color.FromArgb(31, 78, 121);
        d.Blocks.Add(P(d, Loc.T("NGUYỄN VĂN A", "JOHN SMITH"), c => c with { Size = 26, Bold = true, Color = accent }, p => p with { SpaceAfter = 0 }));
        d.Blocks.Add(P(d, Loc.T("Chuyên viên kinh doanh", "Sales specialist"), c => c with { Size = 13, Color = Color.FromArgb(89, 89, 89) }, p => p with { SpaceAfter = 4 }));
        d.Blocks.Add(P(d, Loc.T("Hà Nội · 09xx xxx xxx · email@vidu.vn", "City · +1 555 0100 · email@example.com"), c => c with { Size = 10 }, p => p with { SpaceAfter = 14, Border = false }));
        void Section(string vi, string en)
        {
            d.Blocks.Add(P(d, Loc.T(vi, en).ToUpperInvariant(), c => c with { Bold = true, Size = 12, Color = accent }, p => p with { SpaceBefore = 10, SpaceAfter = 4, Style = StyleIds.Heading2, OutlineLevel = 2, Shading = Color.FromArgb(231, 238, 246) }));
        }
        Section("Mục tiêu nghề nghiệp", "Objective");
        d.Blocks.Add(P(d, Loc.T("Mô tả ngắn gọn mục tiêu và thế mạnh của bạn trong 2-3 câu.", "Describe your goals and strengths in two or three sentences."), null, p => p with { Align = ParaAlign.Justify }));
        Section("Kinh nghiệm làm việc", "Experience");
        d.Blocks.Add(P(d, Loc.T("2022 – nay · Công ty ABC", "2022 – present · ABC Company"), c => c with { Bold = true }, p => p with { SpaceAfter = 2 }));
        int id = 7;
        foreach (var s in new[] { Loc.T("Thành tích hoặc nhiệm vụ chính", "Key achievement or responsibility"), Loc.T("Kết quả đo lường được (doanh số, dự án...)", "Measurable result (sales, projects...)") })
            d.Blocks.Add(P(d, s, null, p => p with { List = ListType.Bullet, ListId = id, IndentLeft = 36, FirstLine = -18, SpaceAfter = 2 }));
        Section("Học vấn", "Education");
        d.Blocks.Add(P(d, Loc.T("Cử nhân Quản trị kinh doanh · Đại học XYZ · 2018 – 2022", "BA in Business Administration · XYZ University · 2018 – 2022")));
        Section("Kỹ năng", "Skills");
        foreach (var s in new[] { Loc.T("Tin học văn phòng (Word, Excel)", "Office software (Word, Excel)"), Loc.T("Giao tiếp, đàm phán", "Communication, negotiation"), Loc.T("Tiếng Anh giao tiếp", "Languages") })
            d.Blocks.Add(P(d, s, null, p => p with { List = ListType.Bullet, ListId = 8, IndentLeft = 36, FirstLine = -18, SpaceAfter = 2 }));
        d.StructureChanged();
        return d;
    }

    private static Document Report()
    {
        var d = Document.CreateBlank("Times New Roman", 13);
        d.Blocks.Clear();
        d.Blocks.Add(Styled(d, StyleIds.Title, Loc.T("BÁO CÁO", "REPORT")));
        d.Blocks.Add(Styled(d, StyleIds.Subtitle, Loc.T("Tiêu đề phụ của báo cáo", "Report subtitle")));
        d.Blocks.Add(Styled(d, StyleIds.Heading1, Loc.T("1. Giới thiệu", "1. Introduction")));
        d.Blocks.Add(P(d, Loc.T("Trình bày bối cảnh, mục đích và phạm vi của báo cáo. Dùng các kiểu Tiêu đề để có thể chèn Mục lục tự động (thẻ Tham chiếu).",
            "Describe the background, purpose and scope. Use Heading styles so you can insert an automatic Table of Contents (References tab)."), null, p => p with { Align = ParaAlign.Justify, FirstLine = 28 }));
        d.Blocks.Add(Styled(d, StyleIds.Heading1, Loc.T("2. Kết quả", "2. Results")));
        var t = Table.Create(4, 3, d.Page.ContentWidth, d.Styles[StyleIds.Normal].Char);
        string[,] cells = { { Loc.T("Chỉ tiêu", "Item"), Loc.T("Kế hoạch", "Plan"), Loc.T("Thực hiện", "Actual") }, { "A", "100", "112" }, { "B", "80", "76" }, { Loc.T("Tổng", "Total"), "180", "188" } };
        for (int r = 0; r < 4; r++)
            for (int c = 0; c < 3; c++)
            {
                var p = t.Rows[r].Cells[c].Paragraphs[0];
                var cp = p.MarkProps with { Bold = r == 0 || r == 3 };
                p.Runs.Add(new Run(cells[r, c], cp));
                p.Props = p.Props with { Align = c == 0 ? ParaAlign.Left : ParaAlign.Right };
                if (r == 0) t.Rows[r].Cells[c].Shading = Color.FromArgb(221, 235, 247);
            }
        d.Blocks.Add(t);
        d.Blocks.Add(Styled(d, StyleIds.Heading1, Loc.T("3. Kết luận và kiến nghị", "3. Conclusions")));
        d.Blocks.Add(P(d, Loc.T("Tóm tắt kết luận chính và các kiến nghị.", "Summarise the main conclusions and recommendations."), null, p => p with { Align = ParaAlign.Justify, FirstLine = 28 }));
        d.Footer.Add(new Paragraph("", d.Styles[StyleIds.Normal].Char with { Size = 10 }, new ParaProps { Align = ParaAlign.Center, SpaceAfter = 0, LineSpacing = 1 }));
        d.Footer[0].Runs.Add(new Run(Run.ObjectChar.ToString(), d.Styles[StyleIds.Normal].Char with { Size = 10 }, new FieldObject(FieldKind.Page)));
        d.StructureChanged();
        return d;
    }
}
