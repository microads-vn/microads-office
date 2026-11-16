using System.Diagnostics;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MicroadsOffice.Core.Office;

public sealed class UpdateManifest
{
    public int Schema { get; set; } = 1;
    public ReleaseInfo? Latest { get; set; }
    public ForceUpdateCommand? ForceUpdate { get; set; }
    public PromoInfo? Promo { get; set; }
}

public sealed class ReleaseInfo
{
    public string Version { get; set; } = "";
    public string Date { get; set; } = "";
    public string Url { get; set; } = "";
    public long Size { get; set; }
    public string Sha256 { get; set; } = "";
    public string Signature { get; set; } = "";
    public LocText? Notes { get; set; }

    public Version? ParsedVersion => System.Version.TryParse(Version, out var v) ? Norm(v) : null;
    internal static Version Norm(Version v) => new(v.Major, v.Minor, Math.Max(0, v.Build));
}

/// <summary>"Update everyone" order issued from the admin page.</summary>
public sealed class ForceUpdateCommand
{
    public bool Enabled { get; set; }
    public string CommandId { get; set; } = "";
}

public enum ApplyResult { Applied, NeedsElevation, Failed }

/// <summary>
/// Checks microads.vn/office for new versions, downloads them, verifies the ECDSA signature
/// and swaps the executable in place (a running .exe can be renamed on Windows, so the
/// new version is used from the next start without any installer).
/// </summary>
public static class UpdateService
{
    private static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true };

    /// <summary>Manifest URL; tests may point it to a local server with MICROADS_UPDATE_URL (http only on localhost).</summary>
    public static string ManifestUrl
    {
        get
        {
            var o = Environment.GetEnvironmentVariable("MICROADS_UPDATE_URL");
            if (!string.IsNullOrWhiteSpace(o) && Uri.TryCreate(o, UriKind.Absolute, out var u) && (u.Scheme == "https" || u.IsLoopback)) return o;
            return Product.ManifestUrl;
        }
    }

    public static Version CurrentVersion { get; set; } = Product.Version;

    private static HttpClient NewClient(TimeSpan timeout)
    {
        var h = new HttpClient(new HttpClientHandler { AllowAutoRedirect = true }) { Timeout = timeout };
        h.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("MicroadsOffice", CurrentVersion.ToString(3)));
        h.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue { NoCache = true };
        return h;
    }

    /// <summary>Downloads the manifest; caches the promo block for offline use. Returns null when offline.</summary>
    public static async Task<UpdateManifest?> FetchManifestAsync(CancellationToken ct = default, TimeSpan? timeout = null)
    {
        try
        {
            using var http = NewClient(timeout ?? TimeSpan.FromSeconds(15));
            var url = ManifestUrl + (ManifestUrl.Contains('?') ? "&" : "?") + "v=" + CurrentVersion.ToString(3);
            var json = await http.GetStringAsync(url, ct).ConfigureAwait(false);
            var m = JsonSerializer.Deserialize<UpdateManifest>(json, Json);
            if (m == null) return null;
            OfficeConfig.Update(c =>
            {
                c.LastUpdateCheckUtc = DateTime.UtcNow;
                if (m.Promo != null) c.Promo = Sanitize(m.Promo);
            });
            return m;
        }
        catch { return null; }
    }

    private static PromoInfo Sanitize(PromoInfo p)
    {
        var d = PromoInfo.Default;
        if (!PromoInfo.IsSafeLink(p.Website)) p.Website = d.Website;
        if (string.IsNullOrWhiteSpace(p.WebsiteLabel)) p.WebsiteLabel = d.WebsiteLabel;
        p.Company ??= d.Company; p.Headline ??= d.Headline; p.Tagline ??= d.Tagline; p.Services ??= d.Services; p.CallToAction ??= d.CallToAction;
        p.Phone ??= ""; p.Email ??= "";
        return p;
    }

    public static bool IsNewer(ReleaseInfo? r) => r?.ParsedVersion is { } v && v > ReleaseInfo.Norm(CurrentVersion);

    /// <summary>True when the server ordered everyone to update and this order was not applied yet.</summary>
    public static bool IsForced(UpdateManifest m)
        => m.ForceUpdate is { Enabled: true } f && IsNewer(m.Latest) && f.CommandId != OfficeConfig.Current.AppliedForceCommand;

    public static bool VerifySignature(ReleaseInfo r, byte[] file)
    {
        try
        {
            var hash = Convert.ToHexString(SHA256.HashData(file)).ToLowerInvariant();
            if (!string.Equals(hash, r.Sha256, StringComparison.OrdinalIgnoreCase)) return false;
            using var ec = ECDsa.Create();
            ec.ImportSubjectPublicKeyInfo(Convert.FromBase64String(Product.UpdatePublicKey), out _);
            var msg = Encoding.UTF8.GetBytes($"MicroadsOffice|{r.Version}|{hash}");
            return ec.VerifyData(msg, Convert.FromBase64String(r.Signature), HashAlgorithmName.SHA256, DSASignatureFormat.Rfc3279DerSequence);
        }
        catch { return false; }
    }

    /// <summary>Downloads the release and returns the bytes only when the signature is valid.</summary>
    public static async Task<byte[]> DownloadAsync(ReleaseInfo r, IProgress<double>? progress, CancellationToken ct)
    {
        if (!Uri.TryCreate(r.Url, UriKind.Absolute, out var u) || !(u.Scheme == "https" || u.IsLoopback))
            throw new InvalidOperationException(Loc.T("Đường dẫn tải bản cập nhật không hợp lệ.", "Invalid update download link."));
        using var http = NewClient(TimeSpan.FromMinutes(30));
        using var resp = await http.GetAsync(u, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();
        long total = resp.Content.Headers.ContentLength ?? r.Size;
        await using var s = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var ms = new MemoryStream(total > 0 ? (int)Math.Min(total, int.MaxValue) : 1 << 20);
        var buf = new byte[1 << 16];
        int n;
        while ((n = await s.ReadAsync(buf, ct).ConfigureAwait(false)) > 0)
        {
            ms.Write(buf, 0, n);
            if (total > 0) progress?.Report(Math.Min(1.0, ms.Length / (double)total));
        }
        var bytes = ms.ToArray();
        if (!VerifySignature(r, bytes))
            throw new CryptographicException(Loc.T("Bản cập nhật không có chữ ký hợp lệ của Microads — đã hủy cài đặt để bảo vệ máy tính.",
                "The update is not signed by Microads — installation cancelled to protect your computer."));
        return bytes;
    }

    /// <summary>Replaces the running executable with the verified new version.</summary>
    public static ApplyResult Apply(byte[] newExe, string version, out string? error)
    {
        error = null;
        if (StorePackage.IsPackaged) { error = "Microsoft Store edition"; return ApplyResult.Failed; } // updated by the Store
        var exe = Environment.ProcessPath!;
        var dir = Path.GetDirectoryName(exe)!;
        var staged = Path.Combine(dir, Product.ExeName + ".new");
        var old = Path.Combine(dir, $"{Product.ExeName}.old-{Environment.ProcessId}");
        try
        {
            File.WriteAllBytes(staged, newExe);
        }
        catch (UnauthorizedAccessException)
        {
            // Installed in a protected folder (Program Files): hand over to the elevated installer.
            try
            {
                var tmp = Path.Combine(AppSettings.DataDirectory, "Updates");
                Directory.CreateDirectory(tmp);
                var setup = Path.Combine(tmp, $"MicroadsOffice_Setup_{version}.exe");
                File.WriteAllBytes(setup, newExe);
                Process.Start(new ProcessStartInfo(setup, $"--install --silent --dir \"{dir}\"") { UseShellExecute = true, Verb = "runas" });
                return ApplyResult.NeedsElevation;
            }
            catch (Exception ex) { error = ex.Message; return ApplyResult.Failed; }
        }
        catch (Exception ex) { error = ex.Message; return ApplyResult.Failed; }

        try
        {
            var target = Path.Combine(dir, Path.GetFileName(exe));
            File.Move(target, old, true);         // allowed while running
            try { File.Move(staged, target, true); }
            catch { File.Move(old, target, true); throw; }
            OfficeConfig.Update(c => c.JustUpdatedTo = version);
            return ApplyResult.Applied;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            try { File.Delete(staged); } catch { }
            return ApplyResult.Failed;
        }
    }

    /// <summary>Deletes leftovers of previous updates (old executables still locked are skipped).</summary>
    public static void CleanupOldFiles()
    {
        try
        {
            var dir = Path.GetDirectoryName(Environment.ProcessPath!)!;
            foreach (var f in Directory.EnumerateFiles(dir, Product.ExeName + ".old-*")) try { File.Delete(f); } catch { }
            foreach (var f in Directory.EnumerateFiles(dir, Product.ExeName + ".new")) try { File.Delete(f); } catch { }
            var up = Path.Combine(AppSettings.DataDirectory, "Updates");
            if (Directory.Exists(up)) foreach (var f in Directory.EnumerateFiles(up)) try { File.Delete(f); } catch { }
        }
        catch { }
    }

    /// <summary>Restarts the given app with the (new) executable.</summary>
    public static void Restart(string arguments)
    {
        try { Process.Start(new ProcessStartInfo(Environment.ProcessPath!, arguments) { UseShellExecute = false }); } catch { }
    }
}
