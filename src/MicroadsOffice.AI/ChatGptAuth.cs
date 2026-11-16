using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI;

/// <summary>Saved ChatGPT sign-in (encrypted with DPAPI). Field names follow OpenAI's credential record example.</summary>
public sealed class ChatGptCredential
{
    public string Email { get; set; } = "";
    public string Name { get; set; } = "";
    public string Issuer { get; set; } = "";
    public string Subject { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ExtAgentHostId { get; set; } = "";
    public string IdToken { get; set; } = "";
    public string AccessToken { get; set; } = "";
    public string RefreshToken { get; set; } = "";
    public string TokenType { get; set; } = "Bearer";
    public DateTimeOffset ExpiresAt { get; set; }
    public List<string> Scopes { get; set; } = [];
    public DateTimeOffset SavedAt { get; set; }

    public bool SignedIn => !string.IsNullOrEmpty(AccessToken) && !string.IsNullOrEmpty(RefreshToken);
    public bool PlanUsage => Scopes.Contains(ChatGptAuth.PlanScope);
}

/// <summary>
/// "Continue with ChatGPT" for open-source apps (OpenAI Sign in with ChatGPT — ChatGPT plan usage).
/// https://developers.openai.com/siwc/token-sharing-open-source/sign-in
/// </summary>
public static class ChatGptAuth
{
    public const string PlanScope = "chatgpt.tokens.use.direct";
    public const string Scopes = "openid profile email offline_access resource.invoke chatgpt.tokens.use.direct";
    public const string Resource = "https://api.openai.com/v1";
    public const string DynamicClient = "dynamic_agent_client";
    public const string AgentName = "Microads Office";
    public const int PreferredPort = 1455;

    /// <summary>Issuer, overridable for automated tests with MICROADS_AI_AUTH_BASE (https, or http on loopback only).</summary>
    public static string Issuer => SafeOverride("MICROADS_AI_AUTH_BASE") ?? "https://auth.openai.com";
    /// <summary>Responses API base, overridable for tests with MICROADS_AI_OPENAI_BASE.</summary>
    public static string ApiBase => SafeOverride("MICROADS_AI_OPENAI_BASE") ?? AiDefaults.OpenAiApiBase;
    static string AuthorizeUrl => Issuer + "/api/accounts/authorize";
    static string TokenUrl => Issuer + "/api/accounts/oauth/token";

    /// <summary>Opens the system browser. Tests replace it to play the browser's part.</summary>
    public static Func<string, bool> OpenBrowser = url =>
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); return true; } catch { return false; }
    };

    /// <summary>Test override: only a server on this computer (loopback) is accepted, so tokens can never be redirected elsewhere.</summary>
    public static string? SafeOverride(string env)
    {
        var v = Environment.GetEnvironmentVariable(env);
        if (string.IsNullOrWhiteSpace(v) || !Uri.TryCreate(v, UriKind.Absolute, out var u)) return null;
        return (u.Scheme == "http" || u.Scheme == "https") && u.IsLoopback ? v.TrimEnd('/') : null;
    }

    // ---------- storage ----------

    static readonly JsonSerializerOptions JsonOpt = new() { WriteIndented = false };
    static readonly object FileLock = new();

    public static ChatGptCredential? Load()
    {
        var raw = SecretStore.Read(SecretNames.ChatGpt);
        if (string.IsNullOrEmpty(raw)) return null;
        try { return JsonSerializer.Deserialize<ChatGptCredential>(raw, JsonOpt); } catch { return null; }
    }

    static void Save(ChatGptCredential c) { lock (FileLock) SecretStore.Write(SecretNames.ChatGpt, JsonSerializer.Serialize(c, JsonOpt)); }

    /// <summary>Stable opaque host id for this installation (urn:uuid), created once before the first sign-in.</summary>
    public static string HostId
    {
        get
        {
            var p = AiPaths.File("host-id.txt");
            try { var s = File.Exists(p) ? File.ReadAllText(p).Trim() : ""; if (s.StartsWith("urn:uuid:")) return s; } catch { }
            var id = "urn:uuid:" + Guid.NewGuid().ToString("D");
            File.WriteAllText(p, id);
            return id;
        }
    }

    public static ChatGptCredential? Current => Load();
    public static bool IsSignedIn => Load() is { SignedIn: true };

    // ---------- PKCE helpers (public for tests) ----------

    public static string RandomUrlSafe(int bytes = 32) => B64Url(RandomNumberGenerator.GetBytes(bytes));
    public static string Challenge(string verifier) => B64Url(SHA256.HashData(Encoding.ASCII.GetBytes(verifier)));
    public static string B64Url(byte[] b) => Convert.ToBase64String(b).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    public static byte[] FromB64Url(string s)
    {
        s = s.Replace('-', '+').Replace('_', '/');
        switch (s.Length % 4) { case 2: s += "=="; break; case 3: s += "="; break; }
        return Convert.FromBase64String(s);
    }

    public static string BuildAuthorizeUrl(string clientId, string redirectUri, string state, string nonce, string challenge, string hostId, string? idTokenHint, string? loginHint)
    {
        var q = new List<(string, string)>
        {
            ("client_id", clientId),
            ("response_type", "code"),
            ("redirect_uri", redirectUri),
            ("scope", Scopes),
            ("resource", Resource),
            ("state", state),
            ("nonce", nonce),
            ("code_challenge_method", "S256"),
            ("code_challenge", challenge),
            ("ext_agent_host_id", hostId),
        };
        if (clientId == DynamicClient) q.Add(("agent_name_hint", AgentName));
        if (!string.IsNullOrEmpty(idTokenHint)) q.Add(("id_token_hint", idTokenHint));
        if (!string.IsNullOrEmpty(loginHint)) q.Add(("login_hint", loginHint));
        return AuthorizeUrl + "?" + string.Join("&", q.Select(p => Uri.EscapeDataString(p.Item1) + "=" + Uri.EscapeDataString(p.Item2)));
    }

    // ---------- sign-in ----------

    /// <summary>
    /// Runs the browser sign-in. <paramref name="newAccount"/> forces a fresh registration (another ChatGPT account);
    /// otherwise the saved registration (client id) is reused. Returns the validated credential.
    /// </summary>
    public static async Task<ChatGptCredential> SignInAsync(bool newAccount, Action<string>? status, CancellationToken ct)
    {
        var previous = Load();
        bool reuse = !newAccount && previous != null && !string.IsNullOrEmpty(previous.ClientId);
        string clientId = reuse ? previous!.ClientId : DynamicClient;
        string hostId = HostId;
        string state = RandomUrlSafe(), nonce = RandomUrlSafe(), verifier = RandomUrlSafe(48);

        using var listener = StartListener(out int port);
        string redirect = $"http://127.0.0.1:{port}/auth/callback";
        var url = BuildAuthorizeUrl(clientId, redirect, state, nonce, Challenge(verifier), hostId,
            reuse && !string.IsNullOrEmpty(previous!.IdToken) ? previous.IdToken : null,
            reuse ? previous!.Email : null);

        status?.Invoke(Loc.T("Đang mở trình duyệt để đăng nhập ChatGPT…", "Opening the browser to sign in to ChatGPT…"));
        if (!OpenBrowser(url))
            throw new AiException(Loc.T("Không mở được trình duyệt. Hãy đặt một trình duyệt mặc định trong Windows rồi thử lại.", "Could not open a web browser. Set a default browser in Windows and try again."));
        status?.Invoke(Loc.T("Hoàn tất đăng nhập trong trình duyệt, rồi quay lại Microads Office.", "Finish signing in in the browser, then come back to Microads Office."));

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromMinutes(10));
        Dictionary<string, string> cb;
        try { cb = await WaitForCallbackAsync(listener, state, timeout.Token); }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested) { throw new AiException(Loc.T("Hết thời gian chờ đăng nhập (10 phút).", "Sign-in timed out (10 minutes).")); }

        if (!cb.TryGetValue("state", out var gotState) || !CryptographicOperations.FixedTimeEquals(Encoding.ASCII.GetBytes(gotState), Encoding.ASCII.GetBytes(state)))
            throw new AiException(Loc.T("Phản hồi đăng nhập không hợp lệ (sai state). Hãy thử lại.", "Invalid sign-in response (state mismatch). Please try again."), "invalid_state");
        if (cb.TryGetValue("error", out var error))
            throw error == "access_denied"
                ? new AiException(Loc.T("Bạn đã hủy đăng nhập ChatGPT.", "You cancelled the ChatGPT sign-in."), error)
                : new AiException(Loc.T("ChatGPT báo lỗi đăng nhập: ", "ChatGPT reported a sign-in error: ") + error + (cb.TryGetValue("error_description", out var ed) ? " — " + ed : ""), error);
        if (!cb.TryGetValue("code", out var code) || string.IsNullOrEmpty(code))
            throw new AiException(Loc.T("Không nhận được mã đăng nhập.", "No authorization code was received."), "no_code");

        cb.TryGetValue("client_id", out var issued);
        if (!reuse)
        {
            if (string.IsNullOrEmpty(issued) || issued == DynamicClient)
                throw new AiException(Loc.T("Đăng ký ứng dụng với ChatGPT chưa hoàn tất. Hãy thử lại.", "Registration with ChatGPT did not complete. Please try again."), "registration_incomplete");
            clientId = issued;
        }
        else if (!string.IsNullOrEmpty(issued) && issued != clientId)
            throw new AiException(Loc.T("ChatGPT trả về ứng dụng khác với tài khoản đã chọn. Hãy chọn \"Dùng tài khoản ChatGPT khác\".", "ChatGPT returned a different app registration. Choose \"Use another ChatGPT account\"."), "client_mismatch");

        status?.Invoke(Loc.T("Đang xác minh tài khoản…", "Verifying the account…"));
        var tokens = await PostFormAsync(TokenUrl, new()
        {
            ["grant_type"] = "authorization_code",
            ["client_id"] = clientId,
            ["code"] = code,
            ["code_verifier"] = verifier,
            ["redirect_uri"] = redirect,
            ["resource"] = Resource,
        }, ct);

        var idToken = tokens["id_token"]?.GetValue<string>() ?? throw new AiException(Loc.T("ChatGPT không trả về ID token.", "ChatGPT returned no ID token."), "no_id_token");
        var claims = await ValidateIdTokenAsync(idToken, clientId, nonce, ct);
        var sub = claims["sub"]?.GetValue<string>() ?? "";
        if (reuse && !string.IsNullOrEmpty(previous!.Subject) && previous.Subject != sub)
            throw new AiException(Loc.T("Bạn đã đăng nhập một tài khoản ChatGPT khác với tài khoản đã lưu. Hãy chọn \"Dùng tài khoản ChatGPT khác\".", "You signed in with a different ChatGPT account than the saved one. Choose \"Use another ChatGPT account\"."), "account_mismatch");

        var cred = FromTokenResponse(tokens, new ChatGptCredential
        {
            Issuer = Issuer, Subject = sub, ClientId = clientId, ExtAgentHostId = hostId,
            Email = claims["email"]?.GetValue<string>() ?? "", Name = claims["name"]?.GetValue<string>() ?? "",
        });
        cred.IdToken = idToken;
        Save(cred);
        return cred;
    }

    static ChatGptCredential FromTokenResponse(JsonNode t, ChatGptCredential c)
    {
        c.AccessToken = t["access_token"]?.GetValue<string>() ?? throw new AiException(Loc.T("ChatGPT không trả về access token.", "ChatGPT returned no access token."), "no_access_token");
        var rt = t["refresh_token"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(rt)) c.RefreshToken = rt;
        if (string.IsNullOrEmpty(c.RefreshToken)) throw new AiException(Loc.T("ChatGPT không trả về refresh token.", "ChatGPT returned no refresh token."), "no_refresh_token");
        c.TokenType = t["token_type"]?.GetValue<string>() ?? "Bearer";
        long exp = 3600; try { exp = t["expires_in"]?.GetValue<long>() ?? 3600; } catch { }
        c.ExpiresAt = DateTimeOffset.UtcNow.AddSeconds(exp);
        var scope = t["scope"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(scope)) c.Scopes = scope.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
        c.SavedAt = DateTimeOffset.UtcNow;
        return c;
    }

    static TcpListener StartListener(out int port)
    {
        foreach (var p in new[] { PreferredPort, 0 })
        {
            try
            {
                var l = new TcpListener(IPAddress.Loopback, p);
                l.Start();
                port = ((IPEndPoint)l.LocalEndpoint).Port;
                return l;
            }
            catch (SocketException) { }
        }
        throw new AiException(Loc.T("Không mở được cổng chờ đăng nhập trên máy.", "Could not open the local sign-in port."));
    }

    static async Task<Dictionary<string, string>> WaitForCallbackAsync(TcpListener listener, string expectedState, CancellationToken ct)
    {
        while (true)
        {
            using var client = await listener.AcceptTcpClientAsync(ct);
            client.ReceiveTimeout = 5000;
            await using var stream = client.GetStream();
            var buf = new byte[16384]; int n = 0;
            using (var rc = CancellationTokenSource.CreateLinkedTokenSource(ct))
            {
                rc.CancelAfter(5000);
                try
                {
                    while (n < buf.Length)
                    {
                        int r = await stream.ReadAsync(buf.AsMemory(n), rc.Token);
                        if (r == 0) break;
                        n += r;
                        if (Encoding.ASCII.GetString(buf, 0, n).Contains("\r\n\r\n")) break;
                    }
                }
                catch (OperationCanceledException) when (!ct.IsCancellationRequested) { continue; }
            }
            var head = Encoding.ASCII.GetString(buf, 0, n);
            var first = head.Split("\r\n")[0].Split(' ');
            if (first.Length < 2 || first[0] != "GET" || !first[1].StartsWith("/auth/callback"))
            {
                await Respond(stream, 404, "Not found");
                continue;
            }
            var query = first[1].Contains('?') ? first[1][(first[1].IndexOf('?') + 1)..] : "";
            var dict = ParseQuery(query);
            // a request without our state (another web page, a stale tab) must not end the sign-in: refuse it and keep waiting
            if (!dict.TryGetValue("state", out var st) || !CryptographicOperations.FixedTimeEquals(Encoding.ASCII.GetBytes(st), Encoding.ASCII.GetBytes(expectedState)))
            {
                await Respond(stream, 400, CallbackPage(false));
                continue;
            }
            bool ok = dict.ContainsKey("code") && !dict.ContainsKey("error");
            await Respond(stream, 200, CallbackPage(ok));
            return dict;
        }
    }

    public static Dictionary<string, string> ParseQuery(string q)
    {
        var d = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var part in q.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            int i = part.IndexOf('=');
            var k = Uri.UnescapeDataString((i < 0 ? part : part[..i]).Replace('+', ' '));
            var v = i < 0 ? "" : Uri.UnescapeDataString(part[(i + 1)..].Replace('+', ' '));
            d[k] = v;
        }
        return d;
    }

    static async Task Respond(NetworkStream s, int status, string body)
    {
        var bytes = Encoding.UTF8.GetBytes(body);
        var head = $"HTTP/1.1 {status} {(status == 200 ? "OK" : status == 400 ? "Bad Request" : "Not Found")}\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {bytes.Length}\r\nCache-Control: no-store\r\nConnection: close\r\n\r\n";
        await s.WriteAsync(Encoding.ASCII.GetBytes(head));
        await s.WriteAsync(bytes);
        await s.FlushAsync();
    }

    static string CallbackPage(bool ok)
    {
        string title = ok ? Loc.T("Đã đăng nhập ChatGPT", "Signed in to ChatGPT") : Loc.T("Chưa đăng nhập được", "Sign-in not completed");
        string text = ok ? Loc.T("Bạn có thể đóng thẻ này và quay lại Microads Office.", "You can close this tab and return to Microads Office.")
                         : Loc.T("Hãy quay lại Microads Office để xem chi tiết.", "Return to Microads Office for details.");
        return $"<!doctype html><html lang=\"{Loc.Lang}\"><meta charset=\"utf-8\"><title>Microads Office</title><body style=\"font-family:Segoe UI,Arial,sans-serif;text-align:center;padding:60px;color:#222\"><h2>{WebUtility.HtmlEncode(title)}</h2><p>{WebUtility.HtmlEncode(text)}</p></body></html>";
    }

    // ---------- token endpoint / refresh / revoke ----------

    static async Task<JsonNode> PostFormAsync(string url, Dictionary<string, string> form, CancellationToken ct)
    {
        using var http = AiProvider.NewHttp(TimeSpan.FromSeconds(60));
        HttpResponseMessage resp;
        try { resp = await http.PostAsync(url, new FormUrlEncodedContent(form), ct); }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested) { throw new AiException(AiProvider.Network(e), inner: e); }
        using (resp)
        {
            var text = await resp.Content.ReadAsStringAsync(ct);
            JsonNode? j = null; try { j = JsonNode.Parse(text); } catch { }
            if (!resp.IsSuccessStatusCode || j == null)
            {
                var code = j?["error"] is JsonValue ev ? ev.GetValue<string>() : j?["error"]?["code"]?.GetValue<string>();
                var desc = j?["error_description"]?.GetValue<string>() ?? j?["error"]?["message"]?.GetValue<string>() ?? "";
                throw new AiException(Loc.T("Máy chủ đăng nhập ChatGPT báo lỗi: ", "The ChatGPT sign-in server reported an error: ") + (code ?? ((int)resp.StatusCode).ToString()) + (desc.Length > 0 ? " — " + desc : ""), code, (int)resp.StatusCode);
            }
            return j;
        }
    }

    /// <summary>
    /// Cross-process lock for refreshing (Word, Excel and PDF run as separate processes and share one rotating refresh token):
    /// an exclusively opened lock file. Not thread-affine (the refresh awaits I/O) and released by Windows if the process dies.
    /// </summary>
    static async Task<FileStream> AcquireRefreshLockAsync(CancellationToken ct)
    {
        var path = AiPaths.File("refresh.lock");
        var until = DateTime.UtcNow.AddSeconds(60);
        while (true)
        {
            try { return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) when (DateTime.UtcNow < until) { await Task.Delay(150, ct); }
            catch (IOException)
            {
                throw new AiException(Loc.T("Đang chờ ứng dụng khác gia hạn phiên ChatGPT quá lâu. Thử lại.", "Timed out waiting for another app to renew the ChatGPT session. Try again."), "refresh_busy");
            }
        }
    }

    /// <summary>Returns a valid access token, refreshing it near expiry. Refreshes are serialised across Word/Excel/PDF processes.</summary>
    public static async Task<string> GetAccessTokenAsync(CancellationToken ct)
    {
        var c = Load();
        if (c == null || !c.SignedIn)
            throw new AiException(Loc.T("Chưa đăng nhập ChatGPT. Mở Cài đặt › Trợ lý AI › Continue with ChatGPT.", "Not signed in to ChatGPT. Open Settings › AI assistant › Continue with ChatGPT."), "not_signed_in") { NeedsSignIn = true };
        if (!c.PlanUsage)
            throw new AiException(Loc.T("Bạn chưa cho phép Microads Office dùng gói ChatGPT. Đăng nhập lại và đồng ý quyền \"dùng gói ChatGPT\", hoặc dùng khóa API.", "You have not allowed Microads Office to use your ChatGPT plan. Sign in again and grant \"use your ChatGPT plan\", or use an API key."), "plan_not_enabled") { NeedsSignIn = true };
        if (c.ExpiresAt > DateTimeOffset.UtcNow.AddMinutes(2)) return c.AccessToken;
        return (await RefreshAsync(force: false, ct)).AccessToken;
    }

    public static async Task<ChatGptCredential> RefreshAsync(bool force, CancellationToken ct)
    {
        using var held = await AcquireRefreshLockAsync(ct);
        {
            var c = Load() ?? throw new AiException(Loc.T("Chưa đăng nhập ChatGPT.", "Not signed in to ChatGPT."), "not_signed_in") { NeedsSignIn = true };
            if (!force && c.ExpiresAt > DateTimeOffset.UtcNow.AddMinutes(2)) return c; // another app refreshed meanwhile
            JsonNode t;
            try
            {
                t = await PostFormAsync(TokenUrl, new()
                {
                    ["grant_type"] = "refresh_token",
                    ["client_id"] = c.ClientId,
                    ["refresh_token"] = c.RefreshToken,
                    ["resource"] = Resource,
                }, CancellationToken.None); // never abandon a rotation half-way: the new refresh token must be saved
            }
            catch (AiException e) when (e.Code is "invalid_grant" or "invalid_refresh_token" or "token_expired" or "refresh_token_expired" or "refresh_token_invalidated" or "refresh_token_reused")
            {
                ClearTokens(c);
                throw new AiException(Loc.T("Phiên đăng nhập ChatGPT đã hết hạn hoặc đã bị ngắt kết nối. Hãy bấm Continue with ChatGPT để đăng nhập lại.", "The ChatGPT session expired or was disconnected. Click Continue with ChatGPT to sign in again."), e.Code, e.Status, e) { NeedsSignIn = true };
            }
            var updated = FromTokenResponse(t, c);
            var newId = t["id_token"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(newId)) updated.IdToken = newId;
            Save(updated);
            return updated;
        }
    }

    static void ClearTokens(ChatGptCredential c)
    {
        c.AccessToken = ""; c.RefreshToken = ""; c.IdToken = ""; c.Scopes = [];
        Save(c); // keep email / subject / client id for the next sign-in to the same account
    }

    /// <summary>Signs out: revokes the refresh token, then clears tokens (the registration is kept). Returns false if revocation was not confirmed.</summary>
    public static async Task<bool> SignOutAsync(CancellationToken ct)
    {
        var c = Load();
        if (c == null) return true;
        bool confirmed = false;
        if (!string.IsNullOrEmpty(c.RefreshToken))
        {
            for (int attempt = 0; attempt < 3 && !confirmed; attempt++)
            {
                try
                {
                    var endpoint = (await DiscoveryAsync(ct))["revocation_endpoint"]?.GetValue<string>() ?? Issuer + "/api/accounts/oauth/revoke";
                    using var http = AiProvider.NewHttp(TimeSpan.FromSeconds(20));
                    using var resp = await http.PostAsync(endpoint, new FormUrlEncodedContent(new Dictionary<string, string>
                    {
                        ["token"] = c.RefreshToken, ["token_type_hint"] = "refresh_token", ["client_id"] = c.ClientId,
                    }), ct);
                    if (resp.IsSuccessStatusCode) confirmed = true;
                    else if ((int)resp.StatusCode < 500) break;
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch { }
                if (!confirmed) await Task.Delay(400 * (attempt + 1), ct);
            }
        }
        else confirmed = true;
        ClearTokens(c);
        return confirmed;
    }

    /// <summary>Forgets the account completely (used by "Use another ChatGPT account" after signing out).</summary>
    public static void Forget() => SecretStore.Write(SecretNames.ChatGpt, null);

    // ---------- ID token validation ----------

    static JsonNode? _discovery; static DateTime _discoveryAt;
    static async Task<JsonNode> DiscoveryAsync(CancellationToken ct)
    {
        if (_discovery != null && DateTime.UtcNow - _discoveryAt < TimeSpan.FromHours(1)) return _discovery;
        using var http = AiProvider.NewHttp(TimeSpan.FromSeconds(30));
        var text = await http.GetStringAsync(Issuer + "/.well-known/openid-configuration", ct);
        _discovery = JsonNode.Parse(text) ?? new JsonObject(); _discoveryAt = DateTime.UtcNow;
        return _discovery;
    }

    /// <summary>Verifies an RS256 ID token against the issuer's JWKS, plus issuer, audience, expiry and nonce. Returns its claims.</summary>
    public static async Task<JsonNode> ValidateIdTokenAsync(string idToken, string clientId, string nonce, CancellationToken ct)
    {
        var parts = idToken.Split('.');
        if (parts.Length != 3) throw Invalid("format");
        var header = JsonNode.Parse(FromB64Url(parts[0])) ?? throw Invalid("header");
        var claims = JsonNode.Parse(FromB64Url(parts[1])) ?? throw Invalid("payload");
        if (header["alg"]?.GetValue<string>() != "RS256") throw Invalid("alg");
        var kid = header["kid"]?.GetValue<string>();

        string jwksUri;
        try { jwksUri = (await DiscoveryAsync(ct))["jwks_uri"]?.GetValue<string>() ?? Issuer + "/.well-known/jwks.json"; }
        catch (Exception e) when (e is not OperationCanceledException) { jwksUri = Issuer + "/.well-known/jwks.json"; }
        using var http = AiProvider.NewHttp(TimeSpan.FromSeconds(30));
        var jwks = JsonNode.Parse(await http.GetStringAsync(jwksUri, ct))?["keys"] as JsonArray ?? throw Invalid("jwks");
        bool verified = false;
        foreach (var k in jwks)
        {
            if (k?["kty"]?.GetValue<string>() != "RSA") continue;
            if (kid != null && k["kid"]?.GetValue<string>() is { } kk && kk != kid) continue;
            using var rsa = RSA.Create();
            rsa.ImportParameters(new RSAParameters { Modulus = FromB64Url(k["n"]!.GetValue<string>()), Exponent = FromB64Url(k["e"]!.GetValue<string>()) });
            if (rsa.VerifyData(Encoding.ASCII.GetBytes(parts[0] + "." + parts[1]), FromB64Url(parts[2]), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
            { verified = true; break; }
        }
        if (!verified) throw Invalid("signature");
        if (claims["iss"]?.GetValue<string>()?.TrimEnd('/') != Issuer.TrimEnd('/')) throw Invalid("issuer");
        var aud = claims["aud"];
        bool audOk = aud is JsonArray arr ? arr.Any(a => a?.GetValue<string>() == clientId) : aud?.GetValue<string>() == clientId;
        if (!audOk) throw Invalid("audience");
        long exp = claims["exp"]?.GetValue<long>() ?? 0;
        if (DateTimeOffset.FromUnixTimeSeconds(exp) < DateTimeOffset.UtcNow.AddMinutes(-5)) throw Invalid("expired");
        if (claims["nonce"]?.GetValue<string>() != nonce) throw Invalid("nonce");
        if (string.IsNullOrEmpty(claims["sub"]?.GetValue<string>())) throw Invalid("sub");
        return claims;
    }

    static AiException Invalid(string what) =>
        new(Loc.T("Không xác minh được tài khoản ChatGPT (", "Could not verify the ChatGPT account (") + what + Loc.T("). Hãy thử lại.", "). Please try again."), "invalid_id_token_" + what);
}
