using MicroadsOffice.Core.Office;
using System.Text.Json;
using System.Text.Json.Serialization;
using MicroadsOffice.Core.Signing;

namespace MicroadsOffice.Core;

/// <summary>
/// Portable settings: everything is stored in the "Data" folder next to the executable
/// (nothing is written to the user profile / system drive).
/// </summary>
public sealed class AppSettings
{
    public List<string> RecentFiles { get; set; } = new();
    public List<SignatureStyle> Styles { get; set; } = new();
    public string? LastStyleId { get; set; }
    /// <summary>
    /// Last signing reason. Until the user signs once nothing is stored, so the suggestion follows the
    /// current language; a value already present in settings.json is kept unchanged.
    /// </summary>
    [JsonIgnore]
    public string LastReason
    {
        get => string.IsNullOrWhiteSpace(StoredLastReason) ? DefaultReason : StoredLastReason!;
        set => StoredLastReason = value;
    }

    [JsonPropertyName("LastReason")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StoredLastReason { get; set; }

    [JsonIgnore] public static string DefaultReason => Loc.T("Tôi đồng ý với nội dung văn bản", "I agree with the content of this document");
    public string LastLocation { get; set; } = "";
    public string LastContact { get; set; } = "";
    public string? LastCertThumbprint { get; set; }
    public SignatureStandard Standard { get; set; } = SignatureStandard.PAdES;
    public bool UseTimestamp { get; set; }
    public string TsaUrl { get; set; } = "http://timestamp.digicert.com";
    public bool OnlineRevocationCheck { get; set; } = true;
    public bool ShowNavigationPane { get; set; } = true;
    public string SignedFileSuffix { get; set; } = "_signed";
    public List<string> PfxFiles { get; set; } = new();
    public int WindowState { get; set; }
    public int[]? WindowBounds { get; set; }

    [JsonIgnore] public static string BaseDirectory { get; set; } = AppContext.BaseDirectory;
    [JsonIgnore] public static string DataDirectory => Path.Combine(BaseDirectory, "Data");
    [JsonIgnore] public static string StylesDirectory => Path.Combine(DataDirectory, "Styles");
    [JsonIgnore] public static string IdsDirectory => Path.Combine(DataDirectory, "DigitalIDs");
    [JsonIgnore] public static string SettingsPath => Path.Combine(DataDirectory, "settings.json");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters = { new JsonStringEnumConverter() }
    };

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var s = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsPath), JsonOpts);
                if (s != null) { s.Normalize(); return s; }
            }
        }
        catch { /* corrupted settings: start fresh */ }
        var d = new AppSettings();
        d.Normalize();
        return d;
    }

    private void Normalize()
    {
        if (Styles.Count == 0)
        {
            Styles.Add(SignatureStyle.CreateDefault());
            Styles.Add(new SignatureStyle
            {
                Id = "seal-overlay", Title = Loc.T("Con dấu nền + thông tin (kiểu Việt Nam)", "Seal background + details (Vietnamese style)"),
                Graphic = SignatureGraphic.Image, Layout = SignatureLayout.GraphicBackground,
                ImageOpacity = 0.85f, ShowLocation = false
            });
        }
        RecentFiles = RecentFiles.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(15).ToList();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(DataDirectory);
            var tmp = SettingsPath + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(this, JsonOpts));
            File.Move(tmp, SettingsPath, true);
        }
        catch { /* read-only location: ignore */ }
    }

    public void AddRecent(string path)
    {
        RecentFiles.RemoveAll(x => string.Equals(x, path, StringComparison.OrdinalIgnoreCase));
        RecentFiles.Insert(0, path);
        if (RecentFiles.Count > 15) RecentFiles.RemoveRange(15, RecentFiles.Count - 15);
    }
}
