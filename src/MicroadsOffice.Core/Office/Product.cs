using System.Reflection;

namespace MicroadsOffice.Core.Office;

public enum AppKind { Hub, Word, Excel, Pdf }

/// <summary>Publisher / product identity shared by every Microads Office application.</summary>
public static class Product
{
    public const string SuiteName = "Microads Office";
    public const string Website = "https://microads.vn";
    public const string WebsiteLabel = "microads.vn";
    public const string UpdatePage = "https://microads.vn/office";
    public const string ManifestUrl = "https://microads.vn/office/api/manifest.json";
    public const string PrivacyUrl = "https://microads.vn/office/privacy";
    public const string ExeName = "MicroadsOffice.exe";

    /// <summary>
    /// Public half of the release-signing key (ECDSA P-256, SubjectPublicKeyInfo, base64).
    /// Every update is verified against it before it is installed, so a tampered download
    /// or a compromised web server cannot push a malicious program to users.
    /// </summary>
    public const string UpdatePublicKey = "MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEXSjzXUOQ/LgAupMzJCQNytd+BZaRwafJixpx9FX1LW4jyHov1tOAxAi86oetJjZFXeVbutcMCa52lXxcX42CXw==";

    public static string CompanyName => Loc.T("Công ty TNHH Công nghệ Microads", "Microads Technology Co., Ltd.");
    public static string PublishedBy => Loc.T("Phát hành bởi Công ty TNHH Công nghệ Microads", "Published by Microads Technology Co., Ltd.");
    public static string AlwaysFree => Loc.T("Luôn luôn miễn phí", "Always free");

    public static Version Version
    {
        get
        {
            var v = Assembly.GetEntryAssembly()?.GetName().Version ?? typeof(Product).Assembly.GetName().Version ?? new Version(1, 0, 0);
            return new Version(v.Major, v.Minor, Math.Max(0, v.Build));
        }
    }

    public static string VersionText => $"{Version.Major}.{Version.Minor}.{Version.Build}";

    public static string AppName(AppKind k) => k switch
    {
        AppKind.Word => "Microads Office Word",
        AppKind.Excel => "Microads Office Excel",
        AppKind.Pdf => "Microads Office PDF",
        _ => SuiteName
    };

    public static string AppShortName(AppKind k) => k switch
    {
        AppKind.Word => "Word",
        AppKind.Excel => "Excel",
        AppKind.Pdf => "PDF",
        _ => "Office"
    };

    public static string AppDescription(AppKind k) => k switch
    {
        AppKind.Word => Loc.T("Soạn thảo văn bản", "Word processing"),
        AppKind.Excel => Loc.T("Bảng tính", "Spreadsheets"),
        AppKind.Pdf => Loc.T("Đọc, ký số và in PDF", "Read, sign and print PDF"),
        _ => Loc.T("Bộ ứng dụng văn phòng miễn phí", "Free office suite")
    };

    public static readonly string[] WordExtensions = [".docx", ".docm", ".dotx", ".rtf", ".txt"];
    public static readonly string[] ExcelExtensions = [".xlsx", ".xlsm", ".xltx", ".csv", ".tsv"];
    public static readonly string[] PdfExtensions = [".pdf"];

    public static AppKind? KindForFile(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (WordExtensions.Contains(ext) || ext == ".doc") return AppKind.Word;
        if (ExcelExtensions.Contains(ext) || ext == ".xls") return AppKind.Excel;
        if (PdfExtensions.Contains(ext)) return AppKind.Pdf;
        return null;
    }

    public static AppKind? ParseKind(string? s) => s?.ToLowerInvariant() switch
    {
        "word" or "doc" or "docx" => AppKind.Word,
        "excel" or "xls" or "xlsx" => AppKind.Excel,
        "pdf" => AppKind.Pdf,
        "hub" or "office" => AppKind.Hub,
        _ => null
    };
}
