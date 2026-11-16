using MicroadsOffice.Core.Office;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace MicroadsOffice.Core.Signing;

public sealed record CertificateEntry(X509Certificate2 Certificate, string Source)
{
    public string DisplayName => PdfSigner.SignerDisplayName(Certificate);
    public string Issuer => Certificate.GetNameInfo(X509NameType.SimpleName, true);
    public bool IsExpired => DateTime.Now > Certificate.NotAfter || DateTime.Now < Certificate.NotBefore;
    public override string ToString() =>
        $"{DisplayName}  —  {Loc.T("Cấp bởi", "Issued by")}: {Issuer}  —  {Loc.T("HSD", "Expires")}: {Certificate.NotAfter:dd/MM/yyyy}{(IsExpired ? "  " + Loc.T("(HẾT HẠN)", "(EXPIRED)") : "")}  [{Source}]";
}

/// <summary>Finds signing certificates (Windows store / USB token, PFX files) and creates self-signed digital IDs.</summary>
public static class CertificateService
{
    /// <summary>Certificates with a private key in CurrentUser\My (includes USB tokens registered by CA drivers).</summary>
    public static List<CertificateEntry> GetStoreCertificates()
    {
        var list = new List<CertificateEntry>();
        try
        {
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
            foreach (var c in store.Certificates)
            {
                if (!c.HasPrivateKey) continue;
                if (!IsUsableForSigning(c)) continue;
                list.Add(new CertificateEntry(c, "Windows / USB Token"));
            }
        }
        catch (CryptographicException) { }
        return list.OrderBy(x => x.IsExpired).ThenBy(x => x.DisplayName).ToList();
    }

    public static bool IsUsableForSigning(X509Certificate2 c)
    {
        foreach (var ext in c.Extensions)
        {
            if (ext is X509KeyUsageExtension ku)
            {
                if ((ku.KeyUsages & (X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation)) == 0) return false;
            }
        }
        return true;
    }

    public static X509Certificate2 LoadPfx(string path, string? password)
    {
        try
        {
            return X509CertificateLoader.LoadPkcs12FromFile(path, password,
                X509KeyStorageFlags.EphemeralKeySet | X509KeyStorageFlags.Exportable);
        }
        catch (CryptographicException ex)
        {
            throw new CryptographicException(Loc.T("Không mở được tệp chứng thư (.pfx/.p12). Kiểm tra lại mật khẩu. ", "Could not open the certificate file (.pfx/.p12). Please check the password. ") + ex.Message, ex);
        }
    }

    /// <summary>Creates a self-signed digital ID (RSA 2048, SHA-256) and saves it as a password-protected PFX.</summary>
    public static X509Certificate2 CreateSelfSigned(string name, string organization, string orgUnit, string email,
        string country, int years, string password, string pfxPath)
    {
        var parts = new List<string> { "CN=" + Escape(name) };
        if (!string.IsNullOrWhiteSpace(orgUnit)) parts.Add("OU=" + Escape(orgUnit));
        if (!string.IsNullOrWhiteSpace(organization)) parts.Add("O=" + Escape(organization));
        if (!string.IsNullOrWhiteSpace(email)) parts.Add("E=" + Escape(email));
        if (!string.IsNullOrWhiteSpace(country)) parts.Add("C=" + Escape(country.Trim().ToUpperInvariant()[..Math.Min(2, country.Trim().Length)]));
        var dn = new X500DistinguishedName(string.Join(", ", parts), X500DistinguishedNameFlags.UseUTF8Encoding);
        using var rsa = RSA.Create(2048);
        var req = new CertificateRequest(dn, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        req.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation, true));
        req.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.4"), new Oid("1.3.6.1.4.1.311.10.3.12") }, false));
        req.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(req.PublicKey, false));
        req.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        var now = DateTimeOffset.Now.AddMinutes(-5);
        using var cert = req.CreateSelfSigned(now, now.AddYears(Math.Clamp(years, 1, 30)));
        var pfx = cert.Export(X509ContentType.Pfx, password);
        Directory.CreateDirectory(Path.GetDirectoryName(pfxPath)!);
        File.WriteAllBytes(pfxPath, pfx);
        return X509CertificateLoader.LoadPkcs12(pfx, password, X509KeyStorageFlags.EphemeralKeySet | X509KeyStorageFlags.Exportable);
    }

    private static string Escape(string v)
    {
        v = v.Trim();
        if (v.IndexOfAny([',', '+', '"', '\\', '<', '>', ';', '=']) >= 0) return "\"" + v.Replace("\"", "\\\"") + "\"";
        return v;
    }
}
