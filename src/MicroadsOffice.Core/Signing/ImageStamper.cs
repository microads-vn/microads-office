using MicroadsOffice.Core.Office;
using System.Drawing;
using MicroadsOffice.Core.Pdf;

namespace MicroadsOffice.Core.Signing;

/// <summary>Adds an image stamp annotation (e.g. company seal, scanned signature) via incremental update.</summary>
public static class ImageStamper
{
    public static byte[] AddImageStamp(byte[] pdf, int pageIndex, RectangleF pageRect, int pageRotation, Image image,
        float opacity, string author)
    {
        var file = PdfFile.Load(pdf);
        if (file.IsEncrypted) throw new PdfSignException(Loc.T("Tài liệu được mã hóa — không thể chèn dấu.", "The document is encrypted — a stamp cannot be inserted."));
        if (PdfSigner.DocMdpLevel(file) is int lvl && lvl < 3)
            throw new PdfSignException(Loc.T("Tài liệu đã được chứng nhận (certified) không cho phép thêm chú thích/con dấu — chèn dấu sẽ làm mất hiệu lực chứng nhận.", "The document is certified and does not allow annotations/stamps — inserting a stamp would invalidate the certification."));
        var pages = file.GetPageRefs();
        if (pageIndex < 0 || pageIndex >= pages.Count) throw new PdfSignException(Loc.T("Trang không hợp lệ.", "Invalid page."));
        var upd = new IncrementalUpdate(file);
        int rot = ((pageRotation % 360) + 360) % 360;
        float vw = rot % 180 == 0 ? pageRect.Width : pageRect.Height;
        float vh = rot % 180 == 0 ? pageRect.Height : pageRect.Width;

        var style = new SignatureStyle
        {
            Graphic = SignatureGraphic.Image, ImageOpacity = opacity,
            ShowName = false, ShowDate = false, ShowReason = false, ShowLocation = false, ShowPlatform = false,
            ShowContact = false, ShowDistinguishedName = false
        };
        var ap = AppearanceRenderer.BuildFormXObject(upd, style, new SignatureAppearanceInfo { Graphic = image }, vw, vh, rot);
        var now = PdfSigner.PdfDate(DateTimeOffset.Now);
        var annot = new PdfDictionary()
            .Set("Type", "Annot").Set("Subtype", "Stamp")
            .Set("Rect", PdfArray.Of(pageRect.Left, pageRect.Top, pageRect.Right, pageRect.Bottom))
            .Set("F", 4)
            .Set("Name", "MicroadsImage")
            .Set("T", PdfString.FromText(author))
            .Set("Contents", PdfString.FromText("Dấu hình ảnh"))
            .Set("M", PdfString.FromText(now))
            .Set("CreationDate", PdfString.FromText(now))
            .Set("P", pages[pageIndex])
            .Set("AP", new PdfDictionary().Set("N", ap));
        var annotRef = upd.Add(annot);
        AnnotationHelper.AppendAnnotation(upd, file, pages[pageIndex], annotRef);
        return upd.Build();
    }
}
