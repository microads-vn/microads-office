using MicroadsOffice.Core.Office;
using System.Drawing;
using System.Formats.Asn1;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using MicroadsOffice.Core.Pdf;

namespace MicroadsOffice.Core.Signing;

public enum SignatureStatus
{
    Valid,          // integrity OK, signature OK, signer trusted
    Unknown,        // integrity & signature OK but signer identity not trusted / could not be checked
    Invalid,        // modified / broken signature
    Unsigned        // empty signature field
}

public sealed class SignatureValidationResult
{
    public string FieldName { get; init; } = "";
    public int PageIndex { get; init; } = -1;
    public RectangleF Rect { get; init; }
    public bool IsSigned { get; init; }
    public string SignerName { get; set; } = "";
    public string Reason { get; set; } = "";
    public string Location { get; set; } = "";
    public string ContactInfo { get; set; } = "";
    public string SubFilter { get; set; } = "";
    public string Application { get; set; } = "";
    public DateTimeOffset? ClaimedSigningTime { get; set; }
    public DateTimeOffset? TimestampTime { get; set; }
    public bool HasTimestamp { get; set; }
    public bool TimestampValid { get; set; }
    public bool IntegrityValid { get; set; }
    public bool SignatureValid { get; set; }
    public bool CertificateTrusted { get; set; }
    public string TrustDetails { get; set; } = "";
    public string ModificationDetails { get; set; } = "";
    public bool CoversWholeDocument { get; set; }
    public bool ChangesAfterAreSignaturesOnly { get; set; }
    public long RevisionLength { get; set; }
    public int Revision { get; set; }
    public string DigestAlgorithm { get; set; } = "";
    public X509Certificate2? Certificate { get; set; }
    public List<X509Certificate2> Chain { get; } = new();
    public List<string> Errors { get; } = new();

    public SignatureStatus Status
    {
        get
        {
            if (!IsSigned) return SignatureStatus.Unsigned;
            if (!IntegrityValid || !SignatureValid) return SignatureStatus.Invalid;
            if (!CoversWholeDocument && !ChangesAfterAreSignaturesOnly) return SignatureStatus.Unknown;
            return CertificateTrusted ? SignatureStatus.Valid : SignatureStatus.Unknown;
        }
    }

    public string StatusText => Status switch
    {
        SignatureStatus.Valid => Loc.T("Chữ ký hợp lệ", "Signature is valid"),
        SignatureStatus.Unknown => !CoversWholeDocument && !ChangesAfterAreSignaturesOnly
            ? Loc.T("Chữ ký hợp lệ nhưng tài liệu đã được thay đổi sau khi ký", "Signature is valid but the document was changed after signing")
            : Loc.T("Chữ ký toàn vẹn — danh tính người ký chưa được xác thực", "Signature intact — the signer's identity is not verified"),
        SignatureStatus.Invalid => Loc.T("Chữ ký KHÔNG hợp lệ", "Signature is NOT valid"),
        _ => Loc.T("Trường chữ ký chưa ký", "Unsigned signature field")
    };
}

public sealed class VerifyOptions
{
    public bool OnlineRevocation { get; set; } = true;
    public bool AllowCertificateDownloads { get; set; } = true;
    public TimeSpan UrlTimeout { get; set; } = TimeSpan.FromSeconds(8);
}

/// <summary>Validates every signature in a PDF: byte-range integrity, CMS signature, certificate trust, timestamps.</summary>
public static class SignatureVerifier
{
    public static List<SignatureValidationResult> Verify(byte[] pdf, VerifyOptions? options = null)
    {
        options ??= new VerifyOptions();
        var results = new List<SignatureValidationResult>();
        PdfFile file;
        try { file = PdfFile.Load(pdf); }
        catch { return results; }
        List<SignatureFieldInfo> fields;
        try { fields = SignatureFields.Find(file); }
        catch { return results; }

        var signedEnds = new List<long>();
        foreach (var f in fields)
        {
            var r = new SignatureValidationResult
            {
                FieldName = f.FullName,
                PageIndex = f.PageIndex,
                Rect = f.Rect,
                IsSigned = f.IsSigned
            };
            results.Add(r);
            if (!f.IsSigned) continue;
            try { VerifyOne(file, f.SignatureDict!, r, options); }
            catch (Exception ex) { r.Errors.Add(Loc.T("Lỗi khi xác thực: ", "Validation error: ") + ex.Message); }
            if (r.RevisionLength > 0) signedEnds.Add(r.RevisionLength);
        }

        // Revision numbering and "changes after signing" analysis.
        var ordered = signedEnds.Distinct().OrderBy(x => x).ToList();
        long maxEnd = ordered.Count > 0 ? ordered[^1] : 0;
        foreach (var r in results.Where(x => x.IsSigned))
        {
            r.Revision = ordered.IndexOf(r.RevisionLength) + 1;
            r.CoversWholeDocument = r.RevisionLength == pdf.Length || IsOnlyTrailingWhitespace(pdf, r.RevisionLength);
            if (!r.CoversWholeDocument && r.RevisionLength > 0 && r.RevisionLength < pdf.Length)
            {
                // Object-level diff between the signed revision and the final file.
                r.ChangesAfterAreSignaturesOnly = RevisionDiff.OnlySignatureChanges(pdf, r.RevisionLength, out var why);
                r.ModificationDetails = r.ChangesAfterAreSignaturesOnly ? "" : why;
            }
        }
        return results;
    }

    private static bool IsOnlyTrailingWhitespace(byte[] pdf, long from)
    {
        if (from <= 0 || from > pdf.Length) return false;
        for (long i = from; i < pdf.Length; i++)
            if (!PdfParser.IsWhite(pdf[i])) return false;
        return true;
    }

    private static void VerifyOne(PdfFile file, PdfDictionary sig, SignatureValidationResult r, VerifyOptions opt)
    {
        var pdf = file.Data;
        r.SubFilter = sig.GetName("SubFilter") ?? "";
        r.Reason = (file.Resolve(sig["Reason"]) as PdfString)?.ToText() ?? "";
        r.Location = (file.Resolve(sig["Location"]) as PdfString)?.ToText() ?? "";
        r.ContactInfo = (file.Resolve(sig["ContactInfo"]) as PdfString)?.ToText() ?? "";
        r.SignerName = (file.Resolve(sig["Name"]) as PdfString)?.ToText() ?? "";
        r.ClaimedSigningTime = PdfSigner.ParsePdfDate((file.Resolve(sig["M"]) as PdfString)?.ToText());
        var app = file.ResolveDict(file.ResolveDict(sig["Prop_Build"])?["App"]);
        if (app != null)
        {
            r.Application = app.GetName("Name") ?? (app["Name"] as PdfString)?.ToText() ?? "";
            if (app["REx"] is PdfString rex) r.Application += " " + rex.ToText();
        }

        if (file.Resolve(sig["ByteRange"]) is not PdfArray br || br.Count < 4 || br.Count % 2 != 0)
        { r.Errors.Add(Loc.T("Thiếu hoặc sai /ByteRange.", "Missing or wrong /ByteRange.")); return; }
        var ranges = br.Items.Select(x => (file.Resolve(x) as PdfNumber)?.LongValue ?? -1).ToArray();
        if (file.Resolve(sig["Contents"]) is not PdfString contents) { r.Errors.Add(Loc.T("Thiếu /Contents.", "Missing /Contents.")); return; }
        long total = 0;
        for (int i = 0; i < ranges.Length; i += 2)
        {
            if (ranges[i] < 0 || ranges[i + 1] < 0 || ranges[i] + ranges[i + 1] > pdf.Length)
            { r.Errors.Add(Loc.T("/ByteRange vượt ngoài tệp — tài liệu đã bị cắt hoặc thay đổi.", "/ByteRange exceeds the file — the document was truncated or changed.")); return; }
            total += ranges[i + 1];
        }
        r.RevisionLength = ranges[^2] + ranges[^1];
        // The gap must be exactly the hex string of /Contents.
        bool gapOk = ranges.Length == 4 && ranges[0] == 0 && ranges[2] > ranges[1] &&
                     pdf[ranges[1]] == '<' && pdf[ranges[2] - 1] == '>' &&
                     ranges[2] - ranges[1] == contents.Bytes.Length * 2L + 2; // the gap is exactly this /Contents hex string
        if (!gapOk) r.Errors.Add(Loc.T("Vùng /ByteRange không bao trùm đúng tài liệu.", "The /ByteRange does not cover the document correctly."));

        var signed = new byte[total];
        long p = 0;
        for (int i = 0; i < ranges.Length; i += 2)
        {
            Buffer.BlockCopy(pdf, (int)ranges[i], signed, (int)p, (int)ranges[i + 1]);
            p += ranges[i + 1];
        }

        var der = TrimDer(contents.Bytes);
        if (r.SubFilter == "ETSI.RFC3161")
        {
            VerifyDocTimestamp(der, signed, r, gapOk, opt);
            return;
        }
        if (r.SubFilter is not ("adbe.pkcs7.detached" or "ETSI.CAdES.detached" or "adbe.pkcs7.sha1"))
        {
            r.Errors.Add(Loc.T("Định dạng chữ ký không được hỗ trợ: ", "Unsupported signature format: ") + r.SubFilter);
            return;
        }

        var cms = new SignedCms();
        if (r.SubFilter == "adbe.pkcs7.sha1")
        {
            cms.Decode(der);
        }
        else
        {
            cms = new SignedCms(new ContentInfo(signed), detached: true);
            cms.Decode(der);
        }
        var si = cms.SignerInfos[0];
        r.Certificate = si.Certificate;
        r.DigestAlgorithm = si.DigestAlgorithm.FriendlyName ?? si.DigestAlgorithm.Value ?? "";
        if (r.Certificate != null) r.SignerName = PdfSigner.SignerDisplayName(r.Certificate);

        // Integrity: compare the messageDigest attribute with our own hash of the byte ranges.
        byte[]? md = null;
        DateTimeOffset? cmsTime = null;
        foreach (var attr in si.SignedAttributes)
        {
            if (attr.Oid.Value == "1.2.840.113549.1.9.4")
                md = new Pkcs9MessageDigest { RawData = attr.Values[0].RawData }.MessageDigest;
            if (attr.Oid.Value == "1.2.840.113549.1.9.5")
            {
                var st = new Pkcs9SigningTime(attr.Values[0].RawData);
                cmsTime = new DateTimeOffset(st.SigningTime.ToUniversalTime(), TimeSpan.Zero);
            }
        }
        r.ClaimedSigningTime ??= cmsTime?.ToLocalTime();
        byte[] dataForHash = signed;
        if (r.SubFilter == "adbe.pkcs7.sha1")
        {
            // encapsulated content is SHA1(byte ranges)
            r.IntegrityValid = cms.ContentInfo.Content.AsSpan().SequenceEqual(SHA1.HashData(signed));
        }
        else
        {
            var ourHash = HashFor(si.DigestAlgorithm.Value, dataForHash);
            r.IntegrityValid = gapOk && md != null && ourHash != null && md.AsSpan().SequenceEqual(ourHash);
            if (md == null) r.Errors.Add(Loc.T("Chữ ký thiếu thuộc tính messageDigest.", "The signature is missing the messageDigest attribute."));
        }
        if (!r.IntegrityValid) r.Errors.Add(Loc.T("Nội dung tài liệu đã bị thay đổi kể từ khi ký (sai giá trị băm).", "The document content has changed since signing (hash mismatch)."));

        try
        {
            si.CheckSignature(verifySignatureOnly: true);
            r.SignatureValid = true;
        }
        catch (CryptographicException ex)
        {
            r.SignatureValid = false;
            r.Errors.Add(Loc.T("Giá trị chữ ký không hợp lệ: ", "Invalid signature value: ") + ex.Message);
        }

        // Timestamp (unsigned attribute)
        foreach (var attr in si.UnsignedAttributes)
        {
            if (attr.Oid.Value != CmsSignatureBuilder.OidTimeStampToken) continue;
            r.HasTimestamp = true;
            try
            {
                if (Rfc3161TimestampToken.TryDecode(attr.Values[0].RawData, out var tst, out _))
                {
                    r.TimestampTime = tst.TokenInfo.Timestamp.ToLocalTime();
                    r.TimestampValid = tst.VerifySignatureForSignerInfo(si, out _);
                }
            }
            catch { r.TimestampValid = false; }
            if (!r.TimestampValid) r.Errors.Add(Loc.T("Dấu thời gian không hợp lệ.", "Invalid timestamp."));
        }

        if (r.Certificate != null)
            CheckTrust(r, cms.Certificates, (r.TimestampValid ? r.TimestampTime : r.ClaimedSigningTime) ?? DateTimeOffset.Now, opt);
    }

    private static void VerifyDocTimestamp(byte[] der, byte[] signed, SignatureValidationResult r, bool gapOk, VerifyOptions opt)
    {
        if (!Rfc3161TimestampToken.TryDecode(der, out var tst, out _)) { r.Errors.Add(Loc.T("Dấu thời gian tài liệu hỏng.", "Corrupted document timestamp.")); return; }
        r.HasTimestamp = true;
        r.TimestampTime = tst.TokenInfo.Timestamp.ToLocalTime();
        r.ClaimedSigningTime = r.TimestampTime;
        X509Certificate2? signerCert = null;
        r.IntegrityValid = gapOk && tst.VerifySignatureForData(signed, out signerCert);
        r.SignatureValid = r.IntegrityValid;
        r.TimestampValid = r.IntegrityValid;
        r.Certificate = signerCert;
        if (signerCert != null) r.SignerName = PdfSigner.SignerDisplayName(signerCert);
        if (!r.IntegrityValid) r.Errors.Add(Loc.T("Dấu thời gian tài liệu không khớp nội dung.", "The document timestamp does not match the content."));
        if (signerCert != null) CheckTrust(r, tst.AsSignedCms().Certificates, r.TimestampTime ?? DateTimeOffset.Now, opt);
    }

    private static void CheckTrust(SignatureValidationResult r, X509Certificate2Collection extra, DateTimeOffset at, VerifyOptions opt)
    {
        using var chain = new X509Chain();
        chain.ChainPolicy.ExtraStore.AddRange(extra);
        chain.ChainPolicy.RevocationMode = opt.OnlineRevocation ? X509RevocationMode.Online : X509RevocationMode.NoCheck;
        chain.ChainPolicy.RevocationFlag = X509RevocationFlag.EntireChain;
        chain.ChainPolicy.DisableCertificateDownloads = !opt.AllowCertificateDownloads;
        chain.ChainPolicy.UrlRetrievalTimeout = opt.UrlTimeout;
        chain.ChainPolicy.VerificationTime = at.LocalDateTime;
        chain.ChainPolicy.VerificationFlags = X509VerificationFlags.IgnoreNotTimeNested;
        bool ok;
        try { ok = chain.Build(r.Certificate!); }
        catch (CryptographicException ex) { r.TrustDetails = ex.Message; return; }
        foreach (var el in chain.ChainElements) r.Chain.Add(el.Certificate);
        var problems = chain.ChainStatus
            .Where(s => s.Status != X509ChainStatusFlags.NoError)
            .Where(s => !(s.Status is X509ChainStatusFlags.RevocationStatusUnknown or X509ChainStatusFlags.OfflineRevocation) || opt.OnlineRevocation)
            .ToList();
        var revokedOrBroken = problems.Where(s => s.Status is X509ChainStatusFlags.Revoked or X509ChainStatusFlags.NotSignatureValid).ToList();
        if (revokedOrBroken.Count > 0)
        {
            r.CertificateTrusted = false;
            r.SignatureValid = false;
            r.Errors.Add(Loc.T("Chứng thư số đã bị thu hồi hoặc chữ ký chứng thư sai: ", "The certificate is revoked or its signature is wrong: ") + string.Join("; ", revokedOrBroken.Select(TranslateStatus)));
        }
        // Unknown revocation status (offline) is reported but does not make the identity untrusted.
        var blocking = problems.Where(s => s.Status is not (X509ChainStatusFlags.RevocationStatusUnknown or X509ChainStatusFlags.OfflineRevocation)).ToList();
        r.CertificateTrusted = blocking.Count == 0 && (ok || problems.Count > 0);
        r.TrustDetails = problems.Count == 0
            ? Loc.T("Chứng thư hợp lệ và được tin cậy bởi kho chứng thư gốc của Windows.", "The certificate is valid and trusted by the Windows root store.")
            : string.Join("\n", problems.Select(TranslateStatus).Distinct());
    }

    private static string TranslateStatus(X509ChainStatus s) => s.Status switch
    {
        X509ChainStatusFlags.UntrustedRoot => Loc.T("Chứng thư gốc (Root CA) không nằm trong danh sách tin cậy.", "The root CA is not in the trusted list."),
        X509ChainStatusFlags.PartialChain => Loc.T("Không xây dựng được chuỗi chứng thư tới CA gốc.", "Could not build the certificate chain to a root CA."),
        X509ChainStatusFlags.NotTimeValid => Loc.T("Chứng thư hết hạn hoặc chưa có hiệu lực tại thời điểm ký.", "The certificate was expired or not yet valid at signing time."),
        X509ChainStatusFlags.Revoked => Loc.T("Chứng thư đã bị thu hồi.", "The certificate has been revoked."),
        X509ChainStatusFlags.RevocationStatusUnknown => Loc.T("Không kiểm tra được tình trạng thu hồi (CRL/OCSP).", "Revocation status (CRL/OCSP) could not be checked."),
        X509ChainStatusFlags.OfflineRevocation => Loc.T("Máy chủ kiểm tra thu hồi không phản hồi.", "The revocation server did not respond."),
        X509ChainStatusFlags.NotValidForUsage => Loc.T("Chứng thư không dùng cho mục đích ký.", "The certificate is not intended for signing."),
        X509ChainStatusFlags.NotSignatureValid => Loc.T("Chữ ký trên chứng thư không hợp lệ.", "The signature on the certificate is invalid."),
        _ => s.StatusInformation.Trim()
    };

    private static byte[]? HashFor(string? oid, byte[] data) => oid switch
    {
        "2.16.840.1.101.3.4.2.1" => SHA256.HashData(data),
        "2.16.840.1.101.3.4.2.2" => SHA384.HashData(data),
        "2.16.840.1.101.3.4.2.3" => SHA512.HashData(data),
        "1.3.14.3.2.26" => SHA1.HashData(data),
        _ => null
    };

    /// <summary>/Contents is zero-padded: cut the DER blob to its declared length.</summary>
    public static byte[] TrimDer(byte[] data)
    {
        try
        {
            var reader = new AsnReader(data, AsnEncodingRules.BER);
            var enc = reader.ReadEncodedValue();
            return enc.ToArray();
        }
        catch
        {
            int end = data.Length;
            while (end > 0 && data[end - 1] == 0) end--;
            return data[..end];
        }
    }
}
