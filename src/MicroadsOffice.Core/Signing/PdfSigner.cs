using MicroadsOffice.Core.Office;
using System.Drawing;
using System.Globalization;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using MicroadsOffice.Core.Pdf;

namespace MicroadsOffice.Core.Signing;

public sealed class SignRequest
{
    public required X509Certificate2 Certificate { get; init; }
    public int PageIndex { get; init; }
    /// <summary>Rectangle in page user space (unrotated, points). Empty = invisible signature.</summary>
    public RectangleF PageRect { get; init; }
    /// <summary>Page /Rotate in degrees (0, 90, 180, 270) so the appearance is upright.</summary>
    public int PageRotation { get; init; }
    /// <summary>When set, the existing empty signature field with this full name is signed.</summary>
    public string? ExistingFieldName { get; init; }
    public SignatureStyle Style { get; init; } = SignatureStyle.CreateDefault();
    public System.Drawing.Image? Graphic { get; init; }
    public string Reason { get; init; } = "";
    public string Location { get; init; } = "";
    public string ContactInfo { get; init; } = "";
    public SignatureStandard Standard { get; init; } = SignatureStandard.PAdES;
    public string? TsaUrl { get; init; }
    public HashAlgorithmName Hash { get; init; } = HashAlgorithmName.SHA256;
    public DateTimeOffset SigningTime { get; init; } = DateTimeOffset.Now;
}

public sealed class PdfSignException(string message, Exception? inner = null) : Exception(message, inner);

/// <summary>Applies a PAdES / PKCS#7 digital signature to a PDF using an incremental update.</summary>
public static class PdfSigner
{
    public static string SignerDisplayName(X509Certificate2 cert)
    {
        var cn = cert.GetNameInfo(X509NameType.SimpleName, false);
        return string.IsNullOrWhiteSpace(cn) ? cert.Subject : cn;
    }

    public static byte[] Sign(byte[] pdf, SignRequest req)
    {
        PdfFile file;
        try { file = PdfFile.Load(pdf); }
        catch (Exception ex) { throw new PdfSignException(Loc.T("Không đọc được cấu trúc PDF để ký: ", "Could not read the PDF structure for signing: ") + ex.Message, ex); }
        if (file.IsEncrypted)
            throw new PdfSignException(Loc.T("Tài liệu đang được mã hóa/bảo vệ bằng mật khẩu. Hãy gỡ bảo vệ trước khi ký số.", "The document is encrypted/password-protected. Remove the protection before signing."));
        CheckDocMdp(file);

        var pages = file.GetPageRefs();
        var upd = new IncrementalUpdate(file);
        var info = new SignatureAppearanceInfo
        {
            SignerName = SignerDisplayName(req.Certificate),
            DistinguishedName = req.Certificate.Subject,
            SigningTime = req.SigningTime,
            Reason = req.Reason,
            Location = req.Location,
            Contact = req.ContactInfo,
            Graphic = req.Graphic
        };

        var sigDict = new SignatureDictionary(req, info.SignerName);
        var sigRef = upd.Add(sigDict);

        if (!string.IsNullOrEmpty(req.ExistingFieldName))
        {
            var field = SignatureFields.Find(file).FirstOrDefault(x => x.FullName == req.ExistingFieldName)
                        ?? throw new PdfSignException(Loc.T($"Không tìm thấy trường chữ ký '{req.ExistingFieldName}'.", $"Signature field '{req.ExistingFieldName}' was not found."));
            if (field.IsSigned) throw new PdfSignException(Loc.T("Trường chữ ký này đã được ký.", "This signature field is already signed."));
            if (field.FieldRef == null) throw new PdfSignException(Loc.T("Trường chữ ký không hợp lệ.", "Invalid signature field."));
            var fieldDict = (file.ResolveDict(field.FieldRef) ?? throw new PdfSignException(Loc.T("Trường chữ ký hỏng.", "Corrupted signature field."))).CloneShallow();
            fieldDict["V"] = sigRef;
            var widgetDict = field.WidgetRef != null && field.WidgetRef.Number != field.FieldRef.Number
                ? file.ResolveDict(field.WidgetRef)?.CloneShallow() : fieldDict;
            if (widgetDict != null && field.Rect.Width > 0 && field.Rect.Height > 0)
            {
                int rot = req.PageRotation;
                float vw = rot % 180 == 0 ? field.Rect.Width : field.Rect.Height;
                float vh = rot % 180 == 0 ? field.Rect.Height : field.Rect.Width;
                var ap = AppearanceRenderer.BuildFormXObject(upd, req.Style, info, vw, vh, rot);
                widgetDict["AP"] = new PdfDictionary().Set("N", ap);
            }
            upd.Replace(field.FieldRef.Number, fieldDict);
            if (widgetDict != null && !ReferenceEquals(widgetDict, fieldDict) && field.WidgetRef != null)
                upd.Replace(field.WidgetRef.Number, widgetDict);
            EnsureSigFlags(upd, file, null);
        }
        else
        {
            if (req.PageIndex < 0 || req.PageIndex >= pages.Count) throw new PdfSignException(Loc.T("Trang không hợp lệ.", "Invalid page."));
            var pageRef = pages[req.PageIndex];
            var names = SignatureFields.AllFieldNames(file);
            int n = 1;
            while (names.Contains("Signature" + n)) n++;
            var widget = new PdfDictionary()
                .Set("Type", "Annot").Set("Subtype", "Widget").Set("FT", "Sig")
                .Set("T", PdfString.FromText("Signature" + n))
                .Set("V", sigRef)
                .Set("F", 132) // Print + Locked
                .Set("P", pageRef);
            bool visible = req.PageRect.Width > 1 && req.PageRect.Height > 1;
            if (visible)
            {
                var r = req.PageRect;
                widget["Rect"] = PdfArray.Of(r.Left, r.Top, r.Right, r.Bottom);
                int rot = ((req.PageRotation % 360) + 360) % 360;
                float vw = rot % 180 == 0 ? r.Width : r.Height;
                float vh = rot % 180 == 0 ? r.Height : r.Width;
                var ap = AppearanceRenderer.BuildFormXObject(upd, req.Style, info, vw, vh, rot);
                widget["AP"] = new PdfDictionary().Set("N", ap);
            }
            else
            {
                widget["Rect"] = PdfArray.Of(0, 0, 0, 0);
                widget["F"] = new PdfNumber(132 | 2); // hidden as well
            }
            var widgetRef = upd.Add(widget);
            AnnotationHelper.AppendAnnotation(upd, file, pageRef, widgetRef);
            EnsureSigFlags(upd, file, widgetRef);
        }

        // Build, then fill ByteRange and Contents. Retry with a bigger reservation if the CMS does not fit.
        int reserve = string.IsNullOrWhiteSpace(req.TsaUrl) ? 12000 : 24000;
        reserve += req.Certificate.RawData.Length * 4;
        for (int attempt = 0; attempt < 3; attempt++)
        {
            sigDict.ContentsReserve = reserve;
            var output = upd.Build();
            long a = sigDict.ContentsStart, b = sigDict.ContentsEnd;
            var byteRange = $"[0 {a} {b} {output.Length - b}]";
            if (byteRange.Length > SignatureDictionary.ByteRangeReserve) throw new PdfSignException(Loc.T("Tệp quá lớn.", "The file is too large."));
            var brBytes = Encoding.ASCII.GetBytes(byteRange.PadRight(SignatureDictionary.ByteRangeReserve));
            Buffer.BlockCopy(brBytes, 0, output, (int)sigDict.ByteRangeStart, brBytes.Length);

            var toSign = new byte[a + (output.Length - b)];
            Buffer.BlockCopy(output, 0, toSign, 0, (int)a);
            Buffer.BlockCopy(output, (int)b, toSign, (int)a, (int)(output.Length - b));
            byte[] cms;
            try
            {
                cms = CmsSignatureBuilder.Sign(toSign, req.Certificate, req.Standard, req.SigningTime, req.TsaUrl, req.Hash);
            }
            catch (CryptographicException ex)
            {
                throw new PdfSignException(Loc.T("Không thể ký bằng chứng thư số này: ", "Could not sign with this certificate: ") + ex.Message +
                                           Loc.T("\n(Kiểm tra USB Token đã cắm, mã PIN, hoặc khóa riêng của chứng thư.)", "\n(Check that the USB token is plugged in, the PIN, or the certificate's private key.)"), ex);
            }
            if (cms.Length > reserve) { reserve = cms.Length + 4096; continue; }
            var hex = Encoding.ASCII.GetBytes(Convert.ToHexString(cms));
            Buffer.BlockCopy(hex, 0, output, (int)a + 1, hex.Length);
            return output;
        }
        throw new PdfSignException(Loc.T("Không đủ chỗ cho dữ liệu chữ ký.", "Not enough space for the signature data."));
    }

    /// <summary>DocMDP permission level of a certified document (1 = no changes, 2 = form fill + sign, 3 = + annotations), or null.</summary>
    public static int? DocMdpLevel(PdfFile file)
    {
        var perms = file.ResolveDict(file.Catalog?["Perms"]);
        var docMdp = file.ResolveDict(perms?["DocMDP"]);
        if (docMdp == null) return null;
        if (file.Resolve(docMdp["Reference"]) is PdfArray refs)
            foreach (var r in refs.Items)
            {
                var tp = file.ResolveDict(file.ResolveDict(r)?["TransformParams"]);
                if (tp != null) return tp["P"] is PdfNumber p ? p.IntValue : 2;
            }
        return 2;
    }

    private static void CheckDocMdp(PdfFile file)
    {
        var perms = file.ResolveDict(file.Catalog?["Perms"]);
        var docMdp = file.ResolveDict(perms?["DocMDP"]);
        if (docMdp == null) return;
        if (file.Resolve(docMdp["Reference"]) is PdfArray refs)
        {
            foreach (var r in refs.Items)
            {
                var rd = file.ResolveDict(r);
                var tp = file.ResolveDict(rd?["TransformParams"]);
                if (tp?["P"] is PdfNumber p && p.IntValue == 1)
                    throw new PdfSignException(Loc.T("Tài liệu đã được chứng nhận (certified) và không cho phép thay đổi, kể cả thêm chữ ký.", "The document is certified and does not allow any changes, including new signatures."));
            }
        }
    }

    private static void EnsureSigFlags(IncrementalUpdate upd, PdfFile file, PdfReference? newField)
    {
        var rootRef = file.Trailer["Root"] as PdfReference ?? throw new PdfSignException(Loc.T("Thiếu Catalog.", "Missing Catalog."));
        var catalog = upd.Get(rootRef.Number) as PdfDictionary ?? throw new PdfSignException(Loc.T("Catalog hỏng.", "Corrupted Catalog."));
        var acroObj = catalog["AcroForm"];
        if (acroObj is PdfReference acroRef)
        {
            var acro = (upd.Get(acroRef.Number) as PdfDictionary ?? new PdfDictionary()).CloneShallow();
            AddField(upd, file, acro, newField);
            acro["SigFlags"] = new PdfNumber(3);
            upd.Replace(acroRef.Number, acro);
        }
        else
        {
            var cat = catalog.CloneShallow();
            var acro = acroObj is PdfDictionary inline ? inline.CloneShallow() : new PdfDictionary();
            AddField(upd, file, acro, newField);
            acro["SigFlags"] = new PdfNumber(3);
            cat["AcroForm"] = acroObj is PdfDictionary ? acro : upd.Add(acro);
            upd.Replace(rootRef.Number, cat);
        }
    }

    private static void AddField(IncrementalUpdate upd, PdfFile file, PdfDictionary acro, PdfReference? field)
    {
        if (field == null) return;
        if (acro["Fields"] is PdfReference fr)
        {
            var arr = new PdfArray((upd.Get(fr.Number) as PdfArray)?.Items ?? []);
            arr.Add(field);
            upd.Replace(fr.Number, arr);
        }
        else
        {
            var arr = new PdfArray((acro["Fields"] as PdfArray)?.Items ?? []);
            arr.Add(field);
            acro["Fields"] = arr;
        }
    }

    /// <summary>Formats a PDF date string, e.g. D:20261116103000+07'00'.</summary>
    public static string PdfDate(DateTimeOffset t)
    {
        var off = t.Offset;
        string sign = off < TimeSpan.Zero ? "-" : "+";
        off = off.Duration();
        return "D:" + t.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) + $"{sign}{off.Hours:D2}'{off.Minutes:D2}'";
    }

    public static DateTimeOffset? ParsePdfDate(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Trim();
        if (s.StartsWith("D:")) s = s[2..];
        try
        {
            int Y = int.Parse(s[..4]);
            int Get(int pos, int def) => s.Length >= pos + 2 && char.IsDigit(s[pos]) ? int.Parse(s.Substring(pos, 2)) : def;
            int M = Get(4, 1), D = Get(6, 1), h = Get(8, 0), m = Get(10, 0), sec = Get(12, 0);
            var off = TimeSpan.Zero;
            if (s.Length > 14 && (s[14] == '+' || s[14] == '-'))
            {
                int oh = Get(15, 0);
                int om = s.Length >= 20 ? Get(18, 0) : 0;
                off = new TimeSpan(oh, om, 0);
                if (s[14] == '-') off = -off;
            }
            return new DateTimeOffset(Y, M, D, h, m, sec, off);
        }
        catch { return null; }
    }

    /// <summary>The /Sig dictionary with fixed-size placeholders for /ByteRange and /Contents.</summary>
    private sealed class SignatureDictionary : PdfObject
    {
        public const int ByteRangeReserve = 64;
        private readonly SignRequest _req;
        private readonly string _name;
        public int ContentsReserve { get; set; } = 12000;
        public long ByteRangeStart { get; private set; }
        public long ContentsStart { get; private set; }
        public long ContentsEnd { get; private set; }

        public SignatureDictionary(SignRequest req, string name) { _req = req; _name = name; }

        public override void WriteTo(PdfByteWriter w)
        {
            var d = new PdfDictionary()
                .Set("Type", "Sig")
                .Set("Filter", "Adobe.PPKLite")
                .Set("SubFilter", _req.Standard == SignatureStandard.PAdES ? "ETSI.CAdES.detached" : "adbe.pkcs7.detached")
                .Set("M", PdfString.FromText(PdfDate(_req.SigningTime)))
                .Set("Name", PdfString.FromText(_name));
            if (!string.IsNullOrWhiteSpace(_req.Reason)) d["Reason"] = PdfString.FromText(_req.Reason);
            if (!string.IsNullOrWhiteSpace(_req.Location)) d["Location"] = PdfString.FromText(_req.Location);
            if (!string.IsNullOrWhiteSpace(_req.ContactInfo)) d["ContactInfo"] = PdfString.FromText(_req.ContactInfo);
            var date = PdfString.FromText(DateTime.Now.ToString("MMM dd yyyy", CultureInfo.InvariantCulture));
            d["Prop_Build"] = new PdfDictionary()
                .Set("Filter", new PdfDictionary().Set("Name", "Adobe.PPKLite").Set("R", 131104).Set("Date", date))
                .Set("PubSec", new PdfDictionary().Set("R", 131105).Set("Date", date).Set("NonEFontNoWarn", new PdfBool(true)))
                .Set("App", new PdfDictionary()
                    .Set("Name", AppearanceRenderer.PlatformFullName)
                    .Set("REx", PdfString.FromText(AppearanceRenderer.PlatformVersion))
                    .Set("R", 65536)
                    .Set("OS", new PdfArray([new PdfName("Win")])));
            d["Prop_AuthType"] = new PdfName("PIN");

            // Serialise manually so that ByteRange/Contents are last and their offsets are known.
            var inner = d.ToBytes();
            w.Write(inner.AsSpan(0, inner.Length - 2)); // without closing ">>"
            w.Write("/ByteRange ");
            ByteRangeStart = w.Position;
            w.Write(new string(' ', ByteRangeReserve));
            w.Write("\n/Contents ");
            ContentsStart = w.Position;
            w.Write("<");
            w.Write(new string('0', ContentsReserve * 2));
            w.Write(">");
            ContentsEnd = w.Position;
            w.Write("\n>>");
        }
    }
}

public static class AnnotationHelper
{
    /// <summary>Appends an annotation reference to the page's /Annots array (handling direct and indirect arrays).</summary>
    public static void AppendAnnotation(IncrementalUpdate upd, PdfFile file, PdfReference pageRef, PdfReference annotRef)
    {
        var page = upd.Get(pageRef.Number) as PdfDictionary ?? throw new PdfSignException(Loc.T("Trang PDF hỏng.", "Corrupted PDF page."));
        if (page["Annots"] is PdfReference ar && upd.Get(ar.Number) is PdfArray existing)
        {
            var arr = new PdfArray(existing.Items);
            arr.Add(annotRef);
            upd.Replace(ar.Number, arr);
        }
        else
        {
            var p2 = page.CloneShallow();
            var arr = new PdfArray((page["Annots"] as PdfArray)?.Items ?? []);
            arr.Add(annotRef);
            p2["Annots"] = arr;
            upd.Replace(pageRef.Number, p2);
        }
    }
}
