using System.Text.Json;
using System.Text.Json.Serialization;

namespace MicroadsOffice.Core.Office;

/// <summary>
/// Preferences shared by all Microads Office applications (language, update state, cached promo,
/// recent files). Stored in Data/office.json. Several apps may run at the same time, so every
/// change is a read-modify-write under a machine-wide mutex (<see cref="Update"/>).
/// </summary>
public sealed class OfficeConfig
{
    /// <summary>"vi" / "en"; null until the user picked a language (first run).</summary>
    public string? Language { get; set; }
    public bool AutoCheckUpdates { get; set; } = true;
    public DateTime? LastUpdateCheckUtc { get; set; }
    public string? SkippedVersion { get; set; }
    /// <summary>Id of the last "update everyone" command already applied, so it runs once.</summary>
    public string? AppliedForceCommand { get; set; }
    /// <summary>Version installed by the last automatic update (shown once as "updated to ...").</summary>
    public string? JustUpdatedTo { get; set; }
    public PromoInfo? Promo { get; set; }
    public Dictionary<string, List<string>> Recent { get; set; } = new();
    public bool ShowPromoOnFileOpen { get; set; } = true;
    /// <summary>Per-app free-form settings (e.g. Word/Excel options) as JSON strings.</summary>
    public Dictionary<string, string> AppData { get; set; } = new();

    [JsonIgnore] public static string FilePath => Path.Combine(AppSettings.DataDirectory, "office.json");

    private static readonly JsonSerializerOptions Opts = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNameCaseInsensitive = true
    };

    private static OfficeConfig? _current;

    /// <summary>Last loaded snapshot (cheap). Call <see cref="Reload"/> to pick up changes from other apps.</summary>
    public static OfficeConfig Current => _current ??= Reload();

    public static OfficeConfig Reload()
    {
        _current = WithLock(ReadUnlocked);
        return _current;
    }

    private static OfficeConfig ReadUnlocked()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<OfficeConfig>(File.ReadAllText(FilePath), Opts) ?? new OfficeConfig();
        }
        catch { /* corrupted: start fresh */ }
        return new OfficeConfig();
    }

    /// <summary>Atomically applies a change to the shared configuration and saves it.</summary>
    public static OfficeConfig Update(Action<OfficeConfig> change)
    {
        _current = WithLock(() =>
        {
            var c = ReadUnlocked();
            change(c);
            try
            {
                Directory.CreateDirectory(AppSettings.DataDirectory);
                var tmp = FilePath + "." + Environment.ProcessId + ".tmp";
                File.WriteAllText(tmp, JsonSerializer.Serialize(c, Opts));
                File.Move(tmp, FilePath, true);
            }
            catch { /* read-only location */ }
            return c;
        });
        return _current;
    }

    private static T WithLock<T>(Func<T> body)
    {
        Mutex? m = null;
        bool owned = false;
        try
        {
            var name = "Local\\MicroadsOffice.Config." + Math.Abs(StringComparer.OrdinalIgnoreCase.GetHashCode(AppSettings.DataDirectory)).ToString("x");
            m = new Mutex(false, name);
            try { owned = m.WaitOne(3000); } catch (AbandonedMutexException) { owned = true; }
            return body();
        }
        finally
        {
            if (owned) try { m!.ReleaseMutex(); } catch { }
            m?.Dispose();
        }
    }

    // ------------------------------------------------------------------ recent files

    public static IReadOnlyList<string> RecentFor(AppKind kind)
        => Current.Recent.TryGetValue(kind.ToString(), out var l) ? l : [];

    public static void AddRecent(AppKind kind, string path)
    {
        Update(c =>
        {
            if (!c.Recent.TryGetValue(kind.ToString(), out var l)) c.Recent[kind.ToString()] = l = new();
            l.RemoveAll(x => string.Equals(x, path, StringComparison.OrdinalIgnoreCase));
            l.Insert(0, path);
            if (l.Count > 25) l.RemoveRange(25, l.Count - 25);
        });
    }

    public static void RemoveRecent(AppKind kind, string? path)
    {
        Update(c =>
        {
            if (!c.Recent.TryGetValue(kind.ToString(), out var l)) return;
            if (path == null) l.Clear();
            else l.RemoveAll(x => string.Equals(x, path, StringComparison.OrdinalIgnoreCase));
        });
    }

    public static string? GetAppData(string key) => Current.AppData.TryGetValue(key, out var v) ? v : null;
    public static void SetAppData(string key, string value) => Update(c => c.AppData[key] = value);
}

/// <summary>
/// Promotional information shown in About, start screens and the "file opened" notice. Defaults are built in;
/// the server (microads.vn/office) can change any field without a new release.
/// </summary>
public sealed class PromoInfo
{
    public bool Enabled { get; set; } = true;
    public bool ShowOnFileOpen { get; set; } = true;
    public string Website { get; set; } = Product.Website;
    public string WebsiteLabel { get; set; } = Product.WebsiteLabel;
    public LocText Company { get; set; } = new("Công ty TNHH Công nghệ Microads", "Microads Technology Co., Ltd.");
    public LocText Headline { get; set; } = new(
        "Microads Office — phát hành bởi Công ty TNHH Công nghệ Microads, luôn luôn miễn phí",
        "Microads Office — published by Microads Technology Co., Ltd., always free");
    public LocText Tagline { get; set; } = new(
        "Chuyên các giải pháp công nghệ cho doanh nghiệp",
        "Technology solutions for businesses");
    public LocList Services { get; set; } = new(
        ["Lập trình phần mềm doanh nghiệp theo yêu cầu", "Bảo mật hệ thống, an toàn thông tin", "Giải pháp CRM quản lý khách hàng", "Hệ thống ERP quản trị doanh nghiệp", "Thiết kế website chuyên nghiệp, chuẩn SEO"],
        ["Custom enterprise software development", "System security & information safety", "CRM customer management solutions", "ERP business management systems", "Professional, SEO-ready website design"]);
    public LocText CallToAction { get; set; } = new("Truy cập microads.vn", "Visit microads.vn");
    public string Phone { get; set; } = "";
    public string Email { get; set; } = "";

    public static PromoInfo Default => new();

    public static PromoInfo Effective => OfficeConfig.Current.Promo ?? Default;

    /// <summary>Only http(s) links may be opened from promo content.</summary>
    public static bool IsSafeLink(string? url)
        => Uri.TryCreate(url, UriKind.Absolute, out var u) && (u.Scheme == Uri.UriSchemeHttps || u.Scheme == Uri.UriSchemeHttp);
}

public sealed class LocText
{
    public string Vi { get; set; } = "";
    public string En { get; set; } = "";
    public LocText() { }
    public LocText(string vi, string en) { Vi = vi; En = en; }
    public override string ToString() => Loc.IsVi || string.IsNullOrWhiteSpace(En) ? Vi : En;
}

public sealed class LocList
{
    public List<string> Vi { get; set; } = new();
    public List<string> En { get; set; } = new();
    public LocList() { }
    public LocList(List<string> vi, List<string> en) { Vi = vi; En = en; }
    public IReadOnlyList<string> Items => Loc.IsVi || En.Count == 0 ? Vi : En;
}
