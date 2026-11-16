namespace MicroadsOffice.Core.Office;

/// <summary>
/// Bilingual strings. Every user-visible text is written inline as a Vietnamese / English pair:
/// <c>Loc.T("Mở tệp", "Open file")</c>. The language is chosen right after installation and can be
/// changed at any time with the flag button (all Microads Office apps share the choice).
/// </summary>
public static class Loc
{
    public const string Vietnamese = "vi";
    public const string English = "en";

    public static string Lang { get; set; } = Vietnamese;
    public static bool IsVi => Lang != English;

    public static string T(string vi, string en) => IsVi ? vi : en;

    /// <summary>Formats with the current language: <c>Loc.F("Trang {0}/{1}", "Page {0} of {1}", a, b)</c>.</summary>
    public static string F(string vi, string en, params object?[] args) => string.Format(T(vi, en), args);

    public static string Normalize(string? lang) => lang == English ? English : Vietnamese;
}
