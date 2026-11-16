using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace MicroadsOffice.WebSetup
{
    /// <summary>
    /// Latest release as published in https://microads.vn/office/api/manifest.json (same manifest as the in-app updater).
    /// MANIFEST CONTRACT (this program never changes): latest.version = "a.b.c" (digits only), latest.url = https,
    /// latest.size = integer bytes (≤ 1 GB), latest.sha256 = hex, latest.signature = base64 DER ECDSA P-256 over
    /// "MicroadsOffice|version|sha256" made with the release key or the backup key below. Requests carry ?v=web-setup,
    /// so the server can give this program its own "latest" if ever needed.
    /// </summary>
    [DataContract]
    public sealed class ReleaseInfo
    {
        [DataMember(Name = "version")] public string Version { get; set; }
        [DataMember(Name = "url")] public string Url { get; set; }
        [DataMember(Name = "size")] public long Size { get; set; }
        [DataMember(Name = "sha256")] public string Sha256 { get; set; }
        [DataMember(Name = "signature")] public string Signature { get; set; }
    }

    [DataContract]
    internal sealed class Manifest
    {
        [DataMember(Name = "latest")] public ReleaseInfo Latest { get; set; }
    }

    public enum FailureKind { Network, NoRelease, Signature, Disk, Launch, Other }

    public sealed class SetupFailure : Exception
    {
        public FailureKind Kind { get; }
        public SetupFailure(FailureKind kind, string detail, Exception inner = null) : base(detail, inner) { Kind = kind; }
    }

    /// <summary>
    /// Download → verify → start. Everything that can be checked is checked before the downloaded program runs:
    /// HTTPS only, size limit, minimum version (no downgrade to old builds), SHA-256 and the ECDSA P-256 release signature
    /// of Microads; the verified file stays locked against writes until Windows has started it.
    /// </summary>
    public static class Engine
    {
        public const string ManifestUrl = "https://microads.vn/office/api/manifest.json";
        public const string PageUrl = "https://microads.vn/office/";
        public const string Website = "https://microads.vn";

        /// <summary>Public half of the Microads release-signing key (SubjectPublicKeyInfo, P-256) — Product.UpdatePublicKey in the app.</summary>
        public const string ReleasePublicKey = "MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEXSjzXUOQ/LgAupMzJCQNytd+BZaRwafJixpx9FX1LW4jyHov1tOAxAi86oetJjZFXeVbutcMCa52lXxcX42CXw==";
        /// <summary>Backup release key (private half kept offline): lets Microads replace the main key without a new web installer.</summary>
        public const string BackupPublicKey = "MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE8SeJ06fZgIVTGrCRRKL2z2z059+TTq+9/JeBIBsdtvtVU06sJ4oHaYL3dehfpGNrbBEujY3UI0Jqzi1eYAw7Sg==";
        /// <summary>Oldest release this program installs (current release when it was frozen): a compromised server cannot push old signed builds.</summary>
        public static readonly Version MinimumVersion = new Version(2, 0, 1);

        private const long MaxSize = 1L << 30; // 1 GB
        private static bool _tlsFallbackUsed;

        static Engine()
        {
            // .NET Framework 4.6.2 defaults to TLS 1.0 unless the OS policy says otherwise: allow TLS 1.2 (and 1.3 where Windows has it).
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; } catch { }
            try { if (Environment.OSVersion.Version.Build >= 22000) ServicePointManager.SecurityProtocol |= (SecurityProtocolType)12288; } catch { }
        }

        /// <summary>Manifest address; automated tests may use MICROADS_UPDATE_URL (https, or http on this computer only).</summary>
        public static string EffectiveManifestUrl
        {
            get
            {
                var o = Environment.GetEnvironmentVariable("MICROADS_UPDATE_URL");
                if (!string.IsNullOrWhiteSpace(o) && Uri.TryCreate(o, UriKind.Absolute, out var u) && (u.Scheme == Uri.UriSchemeHttps || u.IsLoopback)) return o;
                return ManifestUrl;
            }
        }

        private static HttpWebRequest NewRequest(Uri u, int timeoutMs)
        {
            var r = (HttpWebRequest)WebRequest.Create(u);
            r.UserAgent = "MicroadsOfficeWebSetup/1.0 (Windows NT " + Environment.OSVersion.Version.Major + "." + Environment.OSVersion.Version.Minor + ")";
            r.Timeout = timeoutMs;
            r.ReadWriteTimeout = 60000;
            r.AllowAutoRedirect = true;
            r.MaximumAutomaticRedirections = 5;
            r.AutomaticDecompression = DecompressionMethods.None;
            r.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);
            // Company networks: use the system proxy with the user's Windows credentials.
            try { var p = WebRequest.GetSystemWebProxy(); p.Credentials = CredentialCache.DefaultCredentials; r.Proxy = p; } catch { }
            return r;
        }

        /// <summary>
        /// Opens a response; if the TLS handshake fails (e.g. the server one day drops the protocols chosen above), retries
        /// once with the operating system's own TLS defaults (SecurityProtocolType.SystemDefault on .NET Framework 4.7+).
        /// </summary>
        private static HttpWebResponse Open(Uri u, int timeoutMs, CancellationToken ct)
        {
            for (int attempt = 0; ; attempt++)
            {
                var req = NewRequest(u, timeoutMs);
                using (ct.Register(() => { try { req.Abort(); } catch { } }))
                {
                    try { return (HttpWebResponse)req.GetResponse(); }
                    catch (WebException ex) when (attempt == 0 && !_tlsFallbackUsed && !ct.IsCancellationRequested && ex.Status == WebExceptionStatus.SecureChannelFailure)
                    {
                        _tlsFallbackUsed = true;
                        bool switched = false;
                        try { ServicePointManager.SecurityProtocol = 0; switched = true; } catch { }
                        if (!switched) throw; // runtime without SystemDefault: report the original failure
                    }
                }
            }
        }

        private static readonly Regex VersionPattern = new Regex(@"^[0-9]{1,4}\.[0-9]{1,4}\.[0-9]{1,6}\z", RegexOptions.CultureInvariant);
        private static readonly Regex Sha256Pattern = new Regex(@"^[0-9a-fA-F]{64}\z", RegexOptions.CultureInvariant);

        public static bool IsAllowedUrl(string url)
            => Uri.TryCreate(url, UriKind.Absolute, out var u) && (u.Scheme == Uri.UriSchemeHttps || (u.Scheme == Uri.UriSchemeHttp && u.IsLoopback));

        /// <summary>Asks microads.vn for the latest release.</summary>
        public static ReleaseInfo FetchLatest(CancellationToken ct)
        {
            var url = EffectiveManifestUrl;
            url += (url.Contains("?") ? "&" : "?") + "v=web-setup";
            try
            {
                using (var resp = Open(new Uri(url), 20000, ct))
                using (var s = resp.GetResponseStream())
                {
                    var ms = new MemoryStream();
                    CopyLimited(s, ms, 4 << 20, ct);
                    ms.Position = 0;
                    Manifest m;
                    try { m = (Manifest)new DataContractJsonSerializer(typeof(Manifest)).ReadObject(ms); }
                    catch (Exception ex) { throw new SetupFailure(FailureKind.Network, "unreadable manifest", ex); }
                    var r = m?.Latest;
                    if (r == null || string.IsNullOrEmpty(r.Version)) throw new SetupFailure(FailureKind.NoRelease, "no release");
                    if (!VersionPattern.IsMatch(r.Version)) throw new SetupFailure(FailureKind.Signature, "invalid version number in the manifest");
                    if (new Version(r.Version) < MinimumVersion) throw new SetupFailure(FailureKind.Signature, "version older than " + MinimumVersion.ToString(3) + " refused");
                    if (!IsAllowedUrl(r.Url)) throw new SetupFailure(FailureKind.Signature, "download address is not HTTPS");
                    if (string.IsNullOrEmpty(r.Sha256) || !Sha256Pattern.IsMatch(r.Sha256) || string.IsNullOrEmpty(r.Signature)) throw new SetupFailure(FailureKind.Signature, "release is not signed");
                    if (r.Size <= 0 || r.Size > MaxSize) throw new SetupFailure(FailureKind.Signature, "invalid size");
                    return r;
                }
            }
            catch (SetupFailure) { throw; }
            catch (Exception ex) when (ct.IsCancellationRequested) { throw new OperationCanceledException("cancelled", ex, ct); }
            catch (WebException ex) { throw new SetupFailure(FailureKind.Network, ex.Message, ex); }
            catch (IOException ex) { throw new SetupFailure(FailureKind.Network, ex.Message, ex); }
        }

        private static void CopyLimited(Stream s, Stream dst, long limit, CancellationToken ct)
        {
            var buf = new byte[81920];
            int n;
            long total = 0;
            while ((n = s.Read(buf, 0, buf.Length)) > 0)
            {
                ct.ThrowIfCancellationRequested();
                total += n;
                if (total > limit) throw new SetupFailure(FailureKind.Network, "response too large");
                dst.Write(buf, 0, n);
            }
        }

        /// <summary>Creates a new private temporary folder for the download (and removes leftovers older than a day).</summary>
        public static string CreateWorkFolder()
        {
            var temp = Path.GetTempPath();
            try
            {
                foreach (var old in Directory.GetDirectories(temp, "MicroadsOffice-Setup-*"))
                    try { if (Directory.GetCreationTimeUtc(old) < DateTime.UtcNow.AddDays(-1)) Directory.Delete(old, true); } catch { }
            }
            catch { }
            for (int i = 0; i < 5; i++)
            {
                var dir = Path.Combine(temp, "MicroadsOffice-Setup-" + Guid.NewGuid().ToString("N"));
                if (Directory.Exists(dir)) continue;
                try { Directory.CreateDirectory(dir); return dir; }
                catch (Exception ex) { throw new SetupFailure(FailureKind.Disk, ex.Message, ex); }
            }
            throw new SetupFailure(FailureKind.Disk, "cannot create a temporary folder");
        }

        private static bool IsDiskFull(Exception ex)
        {
            int h = ex.HResult;
            return h == unchecked((int)0x80070070) || h == unchecked((int)0x80070027); // ERROR_DISK_FULL, ERROR_HANDLE_DISK_FULL
        }

        private static bool IsSharingViolation(Exception ex)
        {
            int h = ex.HResult;
            return h == unchecked((int)0x80070020) || h == unchecked((int)0x80070021); // ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION
        }

        /// <summary>Downloads the release into <paramref name="dir"/>. <paramref name="progress"/> gets (bytes done, bytes total).</summary>
        public static string Download(ReleaseInfo r, string dir, Action<long, long> progress, CancellationToken ct)
        {
            var path = Path.Combine(dir, "MicroadsOffice_Setup_" + r.Version + ".exe");
            try
            {
                using (var resp = Open(new Uri(r.Url), 30000, ct))
                {
                    if (!IsAllowedUrl(resp.ResponseUri.AbsoluteUri)) throw new SetupFailure(FailureKind.Signature, "redirected to a non-HTTPS address");
                    long total = resp.ContentLength > 0 ? resp.ContentLength : r.Size;
                    if (total > MaxSize) throw new SetupFailure(FailureKind.Signature, "file too large");
                    using (var s = resp.GetResponseStream())
                    using (var f = OpenOutput(path))
                    {
                        var buf = new byte[1 << 16];
                        long done = 0;
                        int n;
                        progress?.Invoke(0, total);
                        while ((n = s.Read(buf, 0, buf.Length)) > 0)
                        {
                            ct.ThrowIfCancellationRequested();
                            done += n;
                            if (done > MaxSize || done > r.Size + (1 << 20)) throw new SetupFailure(FailureKind.Signature, "file larger than announced");
                            f.Write(buf, 0, n);
                            progress?.Invoke(done, total);
                        }
                        if (done != r.Size) throw new SetupFailure(FailureKind.Network, $"incomplete download ({done} of {r.Size} bytes)");
                    }
                }
                return path;
            }
            catch (Exception ex)
            {
                TryDelete(path);
                if (ex is SetupFailure) throw;
                if (ct.IsCancellationRequested) throw new OperationCanceledException("cancelled", ex, ct);
                if (IsDiskFull(ex)) throw new SetupFailure(FailureKind.Disk, ex.Message, ex);
                if (ex is WebException || ex is IOException) throw new SetupFailure(FailureKind.Network, ex.Message, ex);
                throw;
            }
        }

        private static FileStream OpenOutput(string path)
        {
            try { return new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None); }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { throw new SetupFailure(FailureKind.Disk, ex.Message, ex); }
        }

        /// <summary>
        /// Opens the downloaded file read-only (nobody may change it while it is open), checks its SHA-256 and the
        /// Microads signature. Returns the open handle: keep it until the setup process has been created.
        /// </summary>
        public static FileStream VerifyAndLock(string path, ReleaseInfo r, CancellationToken ct)
        {
            FileStream fs = null;
            for (int i = 0; fs == null; i++)
            {
                // an antivirus or backup tool may still be reading/writing the new file for a moment
                try { fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read); }
                catch (IOException ex) when (IsSharingViolation(ex) && i < 40) { ct.ThrowIfCancellationRequested(); Thread.Sleep(250); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { throw new SetupFailure(FailureKind.Disk, ex.Message, ex); }
            }
            try
            {
                string hash;
                using (var sha = new SHA256Cng())
                {
                    var buf = new byte[1 << 20];
                    int n;
                    while ((n = fs.Read(buf, 0, buf.Length)) > 0)
                    {
                        ct.ThrowIfCancellationRequested();
                        sha.TransformBlock(buf, 0, n, null, 0);
                    }
                    sha.TransformFinalBlock(buf, 0, 0);
                    hash = Hex(sha.Hash);
                }
                if (!string.Equals(hash, r.Sha256, StringComparison.OrdinalIgnoreCase))
                    throw new SetupFailure(FailureKind.Signature, "SHA-256 mismatch");
                if (!VerifySignature(r.Version, hash, r.Signature))
                    throw new SetupFailure(FailureKind.Signature, "invalid Microads signature");
                return fs;
            }
            catch { fs.Dispose(); throw; }
        }

        /// <summary>ECDSA P-256 / SHA-256 over "MicroadsOffice|version|sha256" (DER signature) with the release key or the backup key.</summary>
        public static bool VerifySignature(string version, string sha256Hex, string signatureB64)
            => VerifySignature(ReleasePublicKey, version, sha256Hex, signatureB64) || VerifySignature(BackupPublicKey, version, sha256Hex, signatureB64);

        private static bool VerifySignature(string publicKey, string version, string sha256Hex, string signatureB64)
        {
            try
            {
                var spki = Convert.FromBase64String(publicKey);
                byte[] prefix = { 0x30, 0x59, 0x30, 0x13, 0x06, 0x07, 0x2A, 0x86, 0x48, 0xCE, 0x3D, 0x02, 0x01, 0x06, 0x08, 0x2A, 0x86, 0x48, 0xCE, 0x3D, 0x03, 0x01, 0x07, 0x03, 0x42, 0x00, 0x04 };
                if (spki.Length != prefix.Length + 64) return false;
                for (int i = 0; i < prefix.Length; i++) if (spki[i] != prefix[i]) return false;
                var blob = new byte[8 + 64];
                BitConverter.GetBytes(0x31534345).CopyTo(blob, 0); // BCRYPT_ECDSA_PUBLIC_P256_MAGIC "ECS1"
                BitConverter.GetBytes(32).CopyTo(blob, 4);
                Buffer.BlockCopy(spki, prefix.Length, blob, 8, 64);
                var sig = DerToP1363(Convert.FromBase64String(signatureB64));
                if (sig == null) return false;
                byte[] digest;
                using (var sha = new SHA256Cng()) digest = sha.ComputeHash(Encoding.UTF8.GetBytes("MicroadsOffice|" + version + "|" + sha256Hex.ToLowerInvariant()));
                using (var key = CngKey.Import(blob, CngKeyBlobFormat.EccPublicBlob))
                using (var ec = new ECDsaCng(key))
                    return ec.VerifyHash(digest, sig);
            }
            catch { return false; }
        }

        /// <summary>DER SEQUENCE { INTEGER r, INTEGER s } → r‖s (32 + 32 bytes).</summary>
        internal static byte[] DerToP1363(byte[] der)
        {
            int p = 0;
            if (der.Length < 8 || der[p++] != 0x30) return null;
            int len = der[p++];
            if (len == 0x81) len = der[p++];
            else if (len > 0x80) return null;
            if (p + len != der.Length) return null;
            var outp = new byte[64];
            for (int k = 0; k < 2; k++)
            {
                if (p + 2 > der.Length || der[p++] != 0x02) return null;
                int l = der[p++];
                if (l < 1 || l > 33 || p + l > der.Length) return null;
                int start = p, cnt = l;
                while (cnt > 32 && der[start] == 0) { start++; cnt--; }
                if (cnt > 32) return null;
                Buffer.BlockCopy(der, start, outp, k * 32 + (32 - cnt), cnt);
                p += l;
            }
            return p == der.Length ? outp : null;
        }

        private static string Hex(byte[] b)
        {
            var sb = new StringBuilder(b.Length * 2);
            foreach (var x in b) sb.Append(x.ToString("x2"));
            return sb.ToString();
        }

        /// <summary>
        /// Starts the verified setup. Downloaded by this program, it carries no "Mark of the Web". Should a future setup
        /// require administrator rights (Windows error 740), it is started again through the shell, which shows the UAC prompt.
        /// </summary>
        public static Process StartSetup(string path, string arguments)
        {
            try
            {
                var psi = new ProcessStartInfo(path, arguments) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(path) };
                try { return Process.Start(psi) ?? throw new SetupFailure(FailureKind.Launch, "process not started"); }
                catch (Win32Exception ex) when (ex.NativeErrorCode == 740)
                {
                    psi.UseShellExecute = true;
                    return Process.Start(psi) ?? throw new SetupFailure(FailureKind.Launch, "process not started");
                }
            }
            catch (SetupFailure) { throw; }
            catch (Exception ex) { throw new SetupFailure(FailureKind.Launch, ex.Message, ex); }
        }

        /// <summary>Removes the temporary download folder (retries while Windows still holds the file).</summary>
        public static void Cleanup(string dir)
        {
            if (string.IsNullOrEmpty(dir)) return;
            for (int i = 0; i < 20 && Directory.Exists(dir); i++)
            {
                try { Directory.Delete(dir, true); }
                catch { Thread.Sleep(500); }
            }
        }

        private static void TryDelete(string f) { try { if (File.Exists(f)) File.Delete(f); } catch { } }

        /// <summary>Opens a web page in the default browser (tests: MICROADS_WEBSETUP_OPENURL_LOG records it instead).</summary>
        public static void OpenUrl(string url)
        {
            var log = Environment.GetEnvironmentVariable("MICROADS_WEBSETUP_OPENURL_LOG");
            if (!string.IsNullOrEmpty(log)) { try { File.AppendAllText(log, url + Environment.NewLine); } catch { } return; }
            try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { }
        }

        public static void Log(string text)
        {
            try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "MicroadsOffice_Setup.log"), "[" + DateTime.Now.ToString("s") + "] web-setup: " + text + Environment.NewLine); } catch { }
        }
    }
}
