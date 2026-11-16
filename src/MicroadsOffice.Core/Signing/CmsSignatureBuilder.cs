using MicroadsOffice.Core.Office;
using System.Formats.Asn1;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace MicroadsOffice.Core.Signing;

/// <summary>Creates detached CMS (PKCS#7 / CAdES) signatures with optional RFC 3161 timestamps.</summary>
public static class CmsSignatureBuilder
{
    public const string OidSigningCertificateV2 = "1.2.840.113549.1.9.16.2.47";
    public const string OidTimeStampToken = "1.2.840.113549.1.9.16.2.14";

    public static HttpMessageHandler? HttpHandlerOverride { get; set; }

    public static byte[] Sign(byte[] content, X509Certificate2 cert, SignatureStandard standard, DateTimeOffset signingTime,
        string? tsaUrl, HashAlgorithmName hash, bool silent = false)
    {
        var digestOid = hash.Name switch
        {
            "SHA384" => new Oid("2.16.840.1.101.3.4.2.2"),
            "SHA512" => new Oid("2.16.840.1.101.3.4.2.3"),
            _ => new Oid("2.16.840.1.101.3.4.2.1"),
        };
        var cms = new SignedCms(new ContentInfo(content), detached: true);

        AsymmetricAlgorithm? key = null;
        try { key = (AsymmetricAlgorithm?)cert.GetRSAPrivateKey() ?? cert.GetECDsaPrivateKey(); }
        catch (CryptographicException) { key = null; }

        var signer = key != null
            ? new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, cert, key)
            : new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, cert);
        signer.DigestAlgorithm = digestOid;
        signer.IncludeOption = X509IncludeOption.WholeChain;
        // PAdES baseline forbids the CMS signing-time attribute (time goes into the signature dictionary /M).
        if (standard == SignatureStandard.Pkcs7)
            signer.SignedAttributes.Add(new Pkcs9SigningTime(signingTime.UtcDateTime));
        signer.SignedAttributes.Add(new AsnEncodedData(OidSigningCertificateV2, BuildSigningCertificateV2(cert, hash)));

        try
        {
            cms.ComputeSignature(signer, silent);
        }
        catch (CryptographicException) when (key != null)
        {
            // Some token/CSP keys refuse the explicit-key path: retry letting CMS acquire the key itself.
            var s2 = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, cert)
            {
                DigestAlgorithm = digestOid,
                IncludeOption = X509IncludeOption.WholeChain
            };
            foreach (var a in signer.SignedAttributes) s2.SignedAttributes.Add(a);
            cms = new SignedCms(new ContentInfo(content), detached: true);
            cms.ComputeSignature(s2, silent);
        }
        finally { key?.Dispose(); }

        if (!string.IsNullOrWhiteSpace(tsaUrl))
        {
            var si = cms.SignerInfos[0];
            var token = RequestTimestamp(si.GetSignature(), tsaUrl!, hash);
            si.AddUnsignedAttribute(new AsnEncodedData(OidTimeStampToken, token));
        }
        return cms.Encode();
    }

    /// <summary>ESS signing-certificate-v2 (RFC 5035) binding the signer certificate to the signature.</summary>
    private static byte[] BuildSigningCertificateV2(X509Certificate2 cert, HashAlgorithmName hash)
    {
        byte[] certHash = hash.Name switch
        {
            "SHA384" => SHA384.HashData(cert.RawData),
            "SHA512" => SHA512.HashData(cert.RawData),
            _ => SHA256.HashData(cert.RawData)
        };
        var w = new AsnWriter(AsnEncodingRules.DER);
        using (w.PushSequence())              // SigningCertificateV2
        {
            using (w.PushSequence())          // certs SEQUENCE OF ESSCertIDv2
            {
                using (w.PushSequence())      // ESSCertIDv2
                {
                    if (hash.Name is "SHA384" or "SHA512")
                    {
                        using (w.PushSequence())
                        {
                            w.WriteObjectIdentifier(hash.Name == "SHA384" ? "2.16.840.1.101.3.4.2.2" : "2.16.840.1.101.3.4.2.3");
                        }
                    }
                    w.WriteOctetString(certHash);
                    using (w.PushSequence())  // IssuerSerial
                    {
                        using (w.PushSequence()) // GeneralNames
                        {
                            using (w.PushSequence(new Asn1Tag(TagClass.ContextSpecific, 4, true))) // directoryName
                            {
                                w.WriteEncodedValue(cert.IssuerName.RawData);
                            }
                        }
                        w.WriteInteger(cert.SerialNumberBytes.Span);
                    }
                }
            }
        }
        return w.Encode();
    }

    public static byte[] RequestTimestamp(byte[] signatureValue, string tsaUrl, HashAlgorithmName hash)
    {
        var nonce = RandomNumberGenerator.GetBytes(8);
        nonce[0] &= 0x7F;
        var req = Rfc3161TimestampRequest.CreateFromData(signatureValue, hash, requestSignerCertificates: true, nonce: nonce);
        using var http = HttpHandlerOverride != null ? new HttpClient(HttpHandlerOverride, false) : new HttpClient();
        http.Timeout = TimeSpan.FromSeconds(30);
        using var body = new ByteArrayContent(req.Encode());
        body.Headers.ContentType = new MediaTypeHeaderValue("application/timestamp-query");
        HttpResponseMessage resp;
        try
        {
            resp = http.PostAsync(tsaUrl, body).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(Loc.T($"Không kết nối được máy chủ dấu thời gian (TSA) {tsaUrl}: {ex.Message}", $"Could not reach the timestamp server (TSA) {tsaUrl}: {ex.Message}"), ex);
        }
        using (resp)
        {
            if (!resp.IsSuccessStatusCode)
                throw new InvalidOperationException(Loc.T($"Máy chủ dấu thời gian trả về lỗi HTTP {(int)resp.StatusCode}.", $"The timestamp server returned HTTP error {(int)resp.StatusCode}."));
            var bytes = resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            var token = req.ProcessResponse(bytes, out _);
            return token.AsSignedCms().Encode();
        }
    }
}
