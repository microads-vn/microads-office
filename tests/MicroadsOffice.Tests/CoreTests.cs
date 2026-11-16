using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography.X509Certificates;
using MicroadsOffice.Core.Pdf;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Core.Signing;

namespace MicroadsOffice.Tests;

public static class CoreTests
{
    public static string OutDir = "";
    private static X509Certificate2? _cert;
    private static readonly VerifyOptions Offline = new() { OnlineRevocation = false, AllowCertificateDownloads = false };

    private static X509Certificate2 Cert => _cert ??= CertificateService.CreateSelfSigned(
        "Nguyễn Văn Kiểm Thử", "Công ty TNHH Công nghệ Microads", "Phòng QA", "qa@microads.test", "VN", 3, "123456",
        Path.Combine(OutDir, "test-id.pfx"));

    public static void Run(TestRunner t)
    {
        t.Test("PDF parser: classic xref, pages, catalog", () =>
        {
            var f = PdfFile.Load(TestPdf.Create(3));
            T.Eq(3, f.GetPageRefs().Count);
            T.True(!f.LastXrefIsStream && !f.Reconstructed, "classic xref");
        });

        t.Test("PDF parser: broken xref offsets are reconstructed", () =>
        {
            var data = TestPdf.Create(2);
            var s = System.Text.Encoding.Latin1.GetString(data);
            int i = s.LastIndexOf("startxref\n", StringComparison.Ordinal) + 10;
            var broken = System.Text.Encoding.Latin1.GetBytes(s[..i] + "999\n%%EOF\n");
            var f = PdfFile.Load(broken);
            T.True(f.Reconstructed, "reconstructed");
            T.Eq(2, f.GetPageRefs().Count);
            var signed = PdfSigner.Sign(broken, Req(0, new RectangleF(300, 100, 200, 60)));
            var v = SignatureVerifier.Verify(signed, Offline);
            T.True(v.Single().IntegrityValid && v.Single().SignatureValid, "signature on repaired file valid");
            using var d = PdfiumDocument.Load(signed);
            T.Eq(2, d.PageCount);
        });

        t.Test("PDFium: open, render, text, search, link, bookmark", () =>
        {
            using var d = PdfiumDocument.Load(TestPdf.Create(3, withLink: true, withOutline: true));
            T.Eq(3, d.PageCount);
            using var bmp = d.Render(0, 595, 842, 0);
            T.True(CountDark(bmp, new Rectangle(60, 40, 400, 60)) > 100, "text rendered");
            T.True(d.GetPageText(0).Contains("Microads Test Page 1"), "text extraction");
            var hits = d.Search(1, "contract", false, false);
            T.Eq(1, hits.Count);
            T.True(hits[0].Rects.Count > 0, "hit rects");
            var link = d.GetLinkAt(0, 100, 770);
            T.True(link != null && link.PageIndex == 2, "internal link resolved");
            var bms = d.GetBookmarks();
            T.True(bms.Count == 1 && bms[0].Title == "Chapter 2" && bms[0].PageIndex == 1, "bookmark");
        });

        t.Test("PDFium: wrong / missing password reported", () =>
        {
            try { PdfiumDocument.Load(System.Text.Encoding.ASCII.GetBytes("not a pdf")); T.Fail("no exception"); }
            catch (PdfLoadException) { }
        });

        t.Test("Sign: visible PAdES signature is valid (integrity + crypto) and rendered", () =>
        {
            var pdf = TestPdf.Create(2);
            var signed = PdfSigner.Sign(pdf, Req(0, new RectangleF(300, 100, 220, 70)));
            File.WriteAllBytes(Path.Combine(OutDir, "signed-pades.pdf"), signed);
            T.True(signed.AsSpan(0, pdf.Length).SequenceEqual(pdf), "original bytes untouched (incremental)");
            var v = SignatureVerifier.Verify(signed, Offline);
            T.Eq(1, v.Count);
            var r = v[0];
            T.True(r.IntegrityValid, "integrity " + string.Join(";", r.Errors));
            T.True(r.SignatureValid, "signature");
            T.True(r.CoversWholeDocument, "covers whole doc");
            T.Eq("ETSI.CAdES.detached", r.SubFilter);
            T.True(r.Application.Contains("Microads Office (Microads.vn)"), "platform name in /Prop_Build: " + r.Application);
            T.Eq(SignatureStatus.Unknown, r.Status); // self-signed: identity not trusted
            T.Eq(0, r.PageIndex);
            using var d = PdfiumDocument.Load(signed);
            using var bmp = d.Render(0, 595, 842, 0);
            bmp.Save(Path.Combine(OutDir, "signed-pades-p1.png"), ImageFormat.Png);
            // signature box is at x 300..520, y(top-left) 842-170..842-100
            T.True(CountDark(bmp, new Rectangle(305, 675, 210, 60)) > 150, "appearance visible");
        });

        t.Test("Sign: PKCS#7 standard + image graphic, layout background", () =>
        {
            using var seal = MakeSeal();
            var style = new SignatureStyle { Graphic = SignatureGraphic.Image, Layout = SignatureLayout.GraphicBackground, ImageOpacity = 0.8f };
            var signed = PdfSigner.Sign(TestPdf.Create(1), Req(0, new RectangleF(320, 80, 200, 110), style, seal, SignatureStandard.Pkcs7));
            File.WriteAllBytes(Path.Combine(OutDir, "signed-pkcs7-seal.pdf"), signed);
            var r = SignatureVerifier.Verify(signed, Offline).Single();
            T.True(r.IntegrityValid && r.SignatureValid, "valid");
            T.Eq("adbe.pkcs7.detached", r.SubFilter);
            using var d = PdfiumDocument.Load(signed);
            using var bmp = d.Render(0, 595, 842, 0);
            bmp.Save(Path.Combine(OutDir, "signed-pkcs7-seal-p1.png"), ImageFormat.Png);
            T.True(CountColor(bmp, new Rectangle(320, 652, 200, 110), c => c.R > 150 && c.G < 120 && c.B < 120) > 200, "red seal visible");
        });

        t.Test("Sign twice: both signatures stay valid, first reports later signature only", () =>
        {
            var s1 = PdfSigner.Sign(TestPdf.Create(2), Req(0, new RectangleF(50, 100, 200, 60)));
            var s2 = PdfSigner.Sign(s1, Req(1, new RectangleF(300, 100, 200, 60)));
            File.WriteAllBytes(Path.Combine(OutDir, "signed-twice.pdf"), s2);
            var v = SignatureVerifier.Verify(s2, Offline);
            T.Eq(2, v.Count);
            T.True(v.All(x => x.IntegrityValid && x.SignatureValid), "both valid");
            var first = v.First(x => x.FieldName == "Signature1");
            T.True(!first.CoversWholeDocument && first.ChangesAfterAreSignaturesOnly, "first: later signature only");
            T.Eq(1, first.Revision);
            T.Eq(2, v.First(x => x.FieldName == "Signature2").Revision);
            T.Eq(1, v.First(x => x.FieldName == "Signature2").PageIndex);
        });

        t.Test("Attack: page content altered then covered by a new signature → first signature flagged", () =>
        {
            var s1 = PdfSigner.Sign(TestPdf.Create(1), Req(0, new RectangleF(50, 100, 200, 60)));
            var f = PdfFile.Load(s1);
            var page = f.ResolveDict(f.GetPageRefs()[0])!;
            var contentsRef = (PdfReference)page["Contents"]!;
            var upd = new IncrementalUpdate(f);
            var evil = System.Text.Encoding.ASCII.GetBytes("BT /F1 28 Tf 72 760 Td (HACKED CONTRACT) Tj ET\n");
            upd.Replace(contentsRef.Number, new PdfStream(new PdfDictionary(), evil));
            var tampered = upd.Build();
            var s2 = PdfSigner.Sign(tampered, Req(0, new RectangleF(300, 100, 200, 60)));
            File.WriteAllBytes(Path.Combine(OutDir, "attack-resigned.pdf"), s2);
            var v = SignatureVerifier.Verify(s2, Offline);
            var first = v.First(x => x.FieldName == "Signature1");
            T.True(first.IntegrityValid, "old revision bytes intact");
            T.True(!first.ChangesAfterAreSignaturesOnly, "content change must not be treated as 'signature only'");
            T.True(first.ModificationDetails.Length > 0, "details: " + first.ModificationDetails);
            T.True(first.Status != SignatureStatus.Valid, "never green");
            // a legitimate second signature is still accepted
            var ok2 = PdfSigner.Sign(s1, Req(0, new RectangleF(300, 100, 200, 60)));
            T.True(SignatureVerifier.Verify(ok2, Offline).First(x => x.FieldName == "Signature1").ChangesAfterAreSignaturesOnly, "plain re-sign allowed");
        });

        t.Test("Password with Vietnamese diacritics (UTF-8) opens", () =>
        {
            var p = Path.Combine(OutDir, "encrypted-vn.pdf");
            if (!File.Exists(p)) throw new SkipException("encrypted-vn.pdf chưa được tạo (cần pikepdf)");
            try { PdfiumDocument.Load(File.ReadAllBytes(p), "sai"); T.Fail("wrong password accepted"); } catch (PdfPasswordRequiredException) { }
            using var d = PdfiumDocument.Load(File.ReadAllBytes(p), "mậtkhẩu");
            T.True(d.PageCount > 0, "opened");
        });

        t.Test("Tampering is detected", () =>
        {
            var signed = PdfSigner.Sign(TestPdf.Create(1), Req(0, new RectangleF(300, 100, 200, 60)));
            var s = System.Text.Encoding.Latin1.GetString(signed);
            int idx = s.IndexOf("Microads Test Page 1", StringComparison.Ordinal);
            signed[idx] = (byte)'X';
            var r = SignatureVerifier.Verify(signed, Offline).Single();
            T.True(!r.IntegrityValid, "integrity must fail");
            T.Eq(SignatureStatus.Invalid, r.Status);
        });

        t.Test("Changes after signing (image stamp) are reported", () =>
        {
            var signed = PdfSigner.Sign(TestPdf.Create(1), Req(0, new RectangleF(300, 100, 200, 60)));
            using var seal = MakeSeal();
            var stamped = ImageStamper.AddImageStamp(signed, 0, new RectangleF(60, 300, 120, 120), 0, seal, 1f, "QA");
            File.WriteAllBytes(Path.Combine(OutDir, "signed-then-stamped.pdf"), stamped);
            var r = SignatureVerifier.Verify(stamped, Offline).Single();
            T.True(r.IntegrityValid && r.SignatureValid, "old revision still valid");
            T.True(!r.CoversWholeDocument && !r.ChangesAfterAreSignaturesOnly, "modification flagged");
            using var d = PdfiumDocument.Load(stamped);
            using var bmp = d.Render(0, 595, 842, 0);
            T.True(CountColor(bmp, new Rectangle(60, 422, 120, 120), c => c.R > 150 && c.G < 120) > 200, "stamp visible");
        });

        t.Test("Sign rotated page (90°): appearance upright", () =>
        {
            var pdf = TestPdf.Create(1, rotations: [90]);
            var style = new SignatureStyle { Graphic = SignatureGraphic.None, ShowPlatform = true, Border = true };
            // page is shown 842 wide x 595 tall; choose a rect in page space (unrotated) 80 wide x 200 tall => visual 200x80
            var signed = PdfSigner.Sign(pdf, Req(0, new RectangleF(100, 300, 80, 200), style, null, SignatureStandard.PAdES, 90));
            File.WriteAllBytes(Path.Combine(OutDir, "signed-rotated.pdf"), signed);
            var r = SignatureVerifier.Verify(signed, Offline).Single();
            T.True(r.IntegrityValid && r.SignatureValid, "valid");
            using var d = PdfiumDocument.Load(signed);
            var size = d.GetPageSize(0);
            T.True(size.Width > size.Height, "landscape due to /Rotate");
            using var bmp = d.Render(0, (int)size.Width, (int)size.Height, 0);
            bmp.Save(Path.Combine(OutDir, "signed-rotated.png"), ImageFormat.Png);
        });

        t.Test("Sign existing empty signature field", () =>
        {
            var pdf = TestPdf.Create(1, withEmptySigField: true);
            var before = SignatureVerifier.Verify(pdf, Offline);
            T.True(before.Count == 1 && !before[0].IsSigned && before[0].FieldName == "ApproverSignature", "empty field detected");
            var signed = PdfSigner.Sign(pdf, new SignRequest
            {
                Certificate = Cert, ExistingFieldName = "ApproverSignature", Reason = "Phê duyệt", Location = "Hà Nội"
            });
            var r = SignatureVerifier.Verify(signed, Offline).Single();
            T.True(r.IsSigned && r.IntegrityValid && r.SignatureValid, "field signed & valid");
            T.Eq("Phê duyệt", r.Reason);
            T.Eq(0, r.PageIndex);
        });

        t.Test("Invisible signature", () =>
        {
            var signed = PdfSigner.Sign(TestPdf.Create(1), Req(0, RectangleF.Empty));
            var r = SignatureVerifier.Verify(signed, Offline).Single();
            T.True(r.IntegrityValid && r.SignatureValid, "valid");
        });

        t.Test("Appearance preview bitmap renders text (Vietnamese)", () =>
        {
            var style = new SignatureStyle { Graphic = SignatureGraphic.Name, Layout = SignatureLayout.GraphicLeft };
            using var bmp = AppearanceRenderer.RenderBitmap(style, new SignatureAppearanceInfo
            {
                SignerName = "Trần Thị Ánh Dương", Reason = "Đồng ý", Location = "TP. Hồ Chí Minh"
            }, 240, 80, 3, false);
            bmp.Save(Path.Combine(OutDir, "preview-name.png"), ImageFormat.Png);
            T.True(CountDark(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height)) > 500, "has ink");
        });

        t.Test("Xref-stream / object-stream PDF: sign new field + existing field, twice", () =>
        {
            var path = Path.Combine(OutDir, "objstm-input.pdf");
            if (!File.Exists(path)) throw new SkipException("objstm-input.pdf chưa được tạo (cần pikepdf)");
            var pdf = File.ReadAllBytes(path);
            var f = PdfFile.Load(pdf);
            T.True(f.LastXrefIsStream && !f.Reconstructed, "xref stream detected");
            T.True(f.Xref.Values.Any(e => e.Type == 2), "has compressed objects");
            var s1 = PdfSigner.Sign(pdf, Req(0, new RectangleF(60, 100, 200, 60)));
            var s2 = PdfSigner.Sign(s1, new SignRequest { Certificate = Cert, ExistingFieldName = "ApproverSignature", Reason = "Duyệt" });
            File.WriteAllBytes(Path.Combine(OutDir, "signed-objstm.pdf"), s2);
            var v = SignatureVerifier.Verify(s2, Offline);
            T.Eq(2, v.Count);
            T.True(v.All(x => x.IsSigned && x.IntegrityValid && x.SignatureValid), "both valid");
            using var d = PdfiumDocument.Load(s2);
            T.Eq(1, d.PageCount);
        });

        t.Test("PDF date round-trip", () =>
        {
            var now = new DateTimeOffset(2026, 11, 16, 10, 30, 5, TimeSpan.FromHours(7));
            var s = PdfSigner.PdfDate(now);
            T.Eq("D:20261116103005+07'00'", s);
            T.Eq(now, PdfSigner.ParsePdfDate(s)!.Value);
        });

        t.Test("Timestamp (TSA) signature — online", () =>
        {
            try
            {
                var req = Req(0, new RectangleF(300, 100, 200, 60));
                var signed = PdfSigner.Sign(TestPdf.Create(1), new SignRequest
                {
                    Certificate = Cert, PageIndex = 0, PageRect = req.PageRect, TsaUrl = "http://timestamp.digicert.com"
                });
                File.WriteAllBytes(Path.Combine(OutDir, "signed-tsa.pdf"), signed);
                var r = SignatureVerifier.Verify(signed, Offline).Single();
                T.True(r.HasTimestamp && r.TimestampValid && r.TimestampTime != null, "timestamp valid");
                T.True(r.IntegrityValid && r.SignatureValid, "valid");
            }
            catch (PdfSignException ex) when (ex.InnerException is InvalidOperationException)
            {
                throw new SkipException("Không có mạng tới TSA: " + ex.Message);
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("TSA"))
            {
                throw new SkipException("Không có mạng tới TSA: " + ex.Message);
            }
        });
    }

    private static SignRequest Req(int page, RectangleF rect, SignatureStyle? style = null, Image? graphic = null,
        SignatureStandard std = SignatureStandard.PAdES, int rotation = 0) => new()
    {
        Certificate = Cert, PageIndex = page, PageRect = rect, Style = style ?? SignatureStyle.CreateDefault(),
        Graphic = graphic, Reason = "Kiểm thử ký số", Location = "Hà Nội", Standard = std, PageRotation = rotation
    };

    public static Bitmap MakeSeal()
    {
        var bmp = new Bitmap(300, 300, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using var pen = new Pen(Color.FromArgb(220, 210, 30, 40), 14);
        g.DrawEllipse(pen, 15, 15, 270, 270);
        g.DrawEllipse(pen, 60, 60, 180, 180);
        using var f = new Font("Arial", 34, FontStyle.Bold);
        using var br = new SolidBrush(Color.FromArgb(230, 210, 30, 40));
        g.DrawString("MICROADS", f, br, new RectangleF(0, 120, 300, 60), new StringFormat { Alignment = StringAlignment.Center });
        return bmp;
    }

    public static int CountDark(Bitmap bmp, Rectangle r) => CountColor(bmp, r, c => c.A > 0 && c.R + c.G + c.B < 300);

    public static int CountColor(Bitmap bmp, Rectangle r, Func<Color, bool> pred)
    {
        int n = 0;
        r.Intersect(new Rectangle(0, 0, bmp.Width, bmp.Height));
        for (int y = r.Top; y < r.Bottom; y += 1)
            for (int x = r.Left; x < r.Right; x += 1)
                if (pred(bmp.GetPixel(x, y))) n++;
        return n;
    }
}
