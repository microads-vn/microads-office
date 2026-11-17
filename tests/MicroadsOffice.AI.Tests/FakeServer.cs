using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using MicroadsOffice.AI;

namespace MicroadsOffice.AI.Tests;

public sealed record FakeRequest(string Method, string Path, string Query, Dictionary<string, string> Headers, string Body)
{
    public Dictionary<string, string> Form => ChatGptAuth.ParseQuery(Body);
    public JsonNode? Json => JsonNode.Parse(Body);
}

public sealed record FakeResponse(int Status, string ContentType, string Body)
{
    public static FakeResponse Json(object o, int status = 200) => new(status, "application/json", o is string s ? s : System.Text.Json.JsonSerializer.Serialize(o));
    public static FakeResponse Sse(string body) => new(200, "text/event-stream", body);
}

/// <summary>
/// Minimal HTTP/1.1 server on 127.0.0.1 playing OpenAI (sign-in + Responses API), Anthropic (Messages API) and an
/// OpenAI-compatible server (Ollama) for the tests. One request per connection.
/// </summary>
public sealed class FakeServer : IDisposable
{
    readonly TcpListener _l;
    readonly CancellationTokenSource _cts = new();
    public int Port { get; }
    public string Base => $"http://127.0.0.1:{Port}";
    public readonly List<FakeRequest> Requests = [];
    public Func<FakeRequest, FakeResponse?>? Override;

    // ---- OpenAI sign-in state ----
    public readonly RSA Key = RSA.Create(2048);
    public RSA SigningKey; // swapped to make signatures invalid
    public string IssuedClientId = "oaiapp_test123";
    public string? Challenge, Nonce, LastRedirect;
    public string AccessToken = "AT-1", RefreshToken = "RT-1";
    public int TokenSerial = 1;
    public string Scope = "chatgpt.tokens.use.direct email offline_access openid profile resource.invoke";
    public string Subject = "user-1", Email = "nhanvien@example.com";
    public readonly List<string> Revoked = [];
    public int ExpiresIn = 3600;

    public FakeServer()
    {
        SigningKey = Key;
        _l = new TcpListener(IPAddress.Loopback, 0);
        _l.Start();
        Port = ((IPEndPoint)_l.LocalEndpoint).Port;
        _ = Task.Run(Loop);
    }

    async Task Loop()
    {
        while (!_cts.IsCancellationRequested)
        {
            TcpClient c;
            try { c = await _l.AcceptTcpClientAsync(_cts.Token); } catch { break; }
            _ = Task.Run(() => Handle(c));
        }
    }

    async Task Handle(TcpClient c)
    {
        using (c)
        {
            try
            {
                var s = c.GetStream();
                var buf = new List<byte>();
                var one = new byte[1];
                while (true)
                {
                    int n = await s.ReadAsync(one);
                    if (n == 0) return;
                    buf.Add(one[0]);
                    if (buf.Count >= 4 && buf[^1] == '\n' && buf[^2] == '\r' && buf[^3] == '\n' && buf[^4] == '\r') break;
                }
                var head = Encoding.ASCII.GetString(buf.ToArray()).Split("\r\n");
                var first = head[0].Split(' ');
                var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var h in head.Skip(1)) { int i = h.IndexOf(':'); if (i > 0) headers[h[..i].Trim()] = h[(i + 1)..].Trim(); }
                string body = "";
                if (headers.TryGetValue("Content-Length", out var cl) && int.Parse(cl) > 0)
                {
                    var b = new byte[int.Parse(cl)]; int read = 0;
                    while (read < b.Length) { int n = await s.ReadAsync(b.AsMemory(read)); if (n == 0) break; read += n; }
                    body = Encoding.UTF8.GetString(b, 0, read);
                }
                else if (headers.TryGetValue("Transfer-Encoding", out var te) && te.Contains("chunked"))
                {
                    var ms = new MemoryStream();
                    while (true)
                    {
                        var line = await ReadLine(s);
                        int size = Convert.ToInt32(line.Split(';')[0].Trim(), 16);
                        if (size == 0) { await ReadLine(s); break; }
                        var b = new byte[size]; int read = 0;
                        while (read < size) { int n = await s.ReadAsync(b.AsMemory(read)); if (n == 0) break; read += n; }
                        ms.Write(b, 0, read);
                        await ReadLine(s);
                    }
                    body = Encoding.UTF8.GetString(ms.ToArray());
                }
                var target = first[1];
                var path = target.Split('?')[0];
                var query = target.Contains('?') ? target[(target.IndexOf('?') + 1)..] : "";
                var req = new FakeRequest(first[0], path, query, headers, body);
                lock (Requests) Requests.Add(req);
                var resp = Override?.Invoke(req) ?? Route(req);
                var bytes = Encoding.UTF8.GetBytes(resp.Body);
                int n2; lock (Requests) n2 = Requests.Count;
                var hdr = $"HTTP/1.1 {resp.Status} X\r\nContent-Type: {resp.ContentType}; charset=utf-8\r\nContent-Length: {bytes.Length}\r\nx-request-id: req_test_{n2}\r\nConnection: close\r\n\r\n";
                await s.WriteAsync(Encoding.ASCII.GetBytes(hdr));
                await s.WriteAsync(bytes);
                await s.FlushAsync();
            }
            catch { }
        }
    }

    static async Task<string> ReadLine(NetworkStream s)
    {
        var sb = new StringBuilder(); var one = new byte[1];
        while (true) { int n = await s.ReadAsync(one); if (n == 0) break; if (one[0] == '\n') break; if (one[0] != '\r') sb.Append((char)one[0]); }
        return sb.ToString();
    }

    public FakeRequest? Last(string path) { lock (Requests) return Requests.LastOrDefault(r => r.Path == path); }
    public int Count(string path) { lock (Requests) return Requests.Count(r => r.Path == path); }

    // ------------------------------------------------------------------ default routes

    FakeResponse Route(FakeRequest r)
    {
        switch (r.Path)
        {
            case "/.well-known/openid-configuration":
                return FakeResponse.Json(new JsonObject { ["issuer"] = Base, ["jwks_uri"] = Base + "/jwks", ["revocation_endpoint"] = Base + "/oauth/revoke" }.ToJsonString());
            case "/jwks":
                var p = Key.ExportParameters(false);
                return FakeResponse.Json(new JsonObject
                {
                    ["keys"] = new JsonArray(new JsonObject { ["kty"] = "RSA", ["kid"] = "k1", ["use"] = "sig", ["alg"] = "RS256", ["n"] = ChatGptAuth.B64Url(p.Modulus!), ["e"] = ChatGptAuth.B64Url(p.Exponent!) }),
                }.ToJsonString());
            case "/api/accounts/oauth/token":
                return Token(r);
            case "/oauth/revoke":
                Revoked.Add(r.Form.GetValueOrDefault("token") ?? "");
                return new FakeResponse(200, "text/plain", "");
            case "/v1/models":
                if (r.Headers.GetValueOrDefault("Authorization", "").StartsWith("Bearer sk-ant-"))
                    return FakeResponse.Json("""{"data":[{"type":"model","id":"claude-opus-5-5","display_name":"Claude Opus 5.5","created_at":"2026-01-01T00:00:00Z"},{"type":"model","id":"claude-sonnet-5-5","display_name":"Claude Sonnet 5.5","created_at":"2026-01-01T00:00:00Z"}],"has_more":false,"first_id":"claude-opus-5-5","last_id":"claude-sonnet-5-5"}""");
                if (r.Headers.GetValueOrDefault("Authorization") == "Bearer " + AccessToken)
                    return FakeResponse.Json("""{"models":[{"slug":"gpt-test","display_name":"GPT Test","visibility":"list"},{"slug":"gpt-hidden","display_name":"Hidden","visibility":"hide"}]}""");
                if (r.Headers.GetValueOrDefault("Authorization") == "Bearer sk-test-openai")
                    return FakeResponse.Json("""{"data":[{"id":"text-embedding-3","created":1},{"id":"gpt-old","created":10},{"id":"gpt-new","created":20},{"id":"gpt-new-mini","created":30}]}""");
                return FakeResponse.Json("""{"error":{"code":"invalid_api_key","message":"bad key"}}""", 401);
            case "/v1/responses":
                if (r.Headers.GetValueOrDefault("Authorization") is not ("Bearer sk-test-openai") && r.Headers.GetValueOrDefault("Authorization") != "Bearer " + AccessToken)
                    return FakeResponse.Json("""{"error":{"code":"subscription_sharing_invalid_user","message":"bad token"}}""", 401);
                return FakeResponse.Sse(ResponsesSse(Reply ?? "Xin chào từ ChatGPT."));
            case "/v1/messages":
                if (Environment.GetEnvironmentVariable("FAKE_DEBUG") == "1") Console.WriteLine("DEBUG /v1/messages headers: " + string.Join(" | ", r.Headers.Select(h => h.Key + "=" + h.Value)) + " BODY: " + r.Body);
                if (r.Headers.GetValueOrDefault("Authorization") != "Bearer sk-ant-test")
                    return FakeResponse.Json("""{"type":"error","error":{"type":"authentication_error","message":"invalid x-api-key"}}""", 401);
                return FakeResponse.Sse(ClaudeSse(Reply ?? "Chào bạn từ Claude."));
            case "/report.php":
                return FakeResponse.Json("{\"ok\":true,\"id\":\"abc123\"}");
            case "/compat/v1/models":
                return FakeResponse.Json("""{"object":"list","data":[{"id":"qwen-local"},{"id":"llama-local"}]}""");
            case "/compat/v1/chat/completions":
                return FakeResponse.Sse(CompatSse(Reply ?? "Trả lời từ máy chủ nội bộ."));
        }
        return new FakeResponse(404, "text/plain", "not found");
    }

    /// <summary>The text the fake AI answers (null = a default greeting).</summary>
    public string? Reply;

    FakeResponse Token(FakeRequest r)
    {
        var f = r.Form;
        if (f.GetValueOrDefault("grant_type") == "authorization_code")
        {
            if (f.GetValueOrDefault("client_id") != IssuedClientId) return FakeResponse.Json("""{"error":"invalid_client"}""", 401);
            if (f.GetValueOrDefault("code") != "CODE-1") return FakeResponse.Json("""{"error":"invalid_grant"}""", 400);
            if (ChatGptAuth.Challenge(f.GetValueOrDefault("code_verifier") ?? "") != Challenge) return FakeResponse.Json("""{"error":"invalid_grant","error_description":"PKCE"}""", 400);
            if (f.GetValueOrDefault("redirect_uri") != LastRedirect) return FakeResponse.Json("""{"error":"invalid_grant","error_description":"redirect"}""", 400);
            if (f.GetValueOrDefault("resource") != ChatGptAuth.Resource) return FakeResponse.Json("""{"error":"invalid_target"}""", 400);
            return FakeResponse.Json(TokenJson(true));
        }
        if (f.GetValueOrDefault("grant_type") == "refresh_token")
        {
            if (f.GetValueOrDefault("refresh_token") != RefreshToken || f.GetValueOrDefault("client_id") != IssuedClientId)
                return FakeResponse.Json("""{"error":"invalid_grant","error_description":"refresh token invalid"}""", 400);
            TokenSerial++;
            AccessToken = "AT-" + TokenSerial; RefreshToken = "RT-" + TokenSerial;
            return FakeResponse.Json(TokenJson(false));
        }
        return FakeResponse.Json("""{"error":"unsupported_grant_type"}""", 400);
    }

    string TokenJson(bool withId)
    {
        var o = new JsonObject
        {
            ["access_token"] = AccessToken, ["refresh_token"] = RefreshToken, ["token_type"] = "Bearer", ["expires_in"] = ExpiresIn, ["scope"] = Scope,
        };
        if (withId) o["id_token"] = IdToken(Nonce ?? "");
        return o.ToJsonString();
    }

    public string IdToken(string nonce, string? aud = null, string? iss = null)
    {
        var header = ChatGptAuth.B64Url(Encoding.UTF8.GetBytes("""{"alg":"RS256","kid":"k1","typ":"JWT"}"""));
        var payload = ChatGptAuth.B64Url(Encoding.UTF8.GetBytes(new JsonObject
        {
            ["iss"] = iss ?? Base, ["aud"] = aud ?? IssuedClientId, ["sub"] = Subject, ["email"] = Email, ["name"] = "Nhân Viên",
            ["nonce"] = nonce, ["iat"] = DateTimeOffset.UtcNow.ToUnixTimeSeconds(), ["exp"] = DateTimeOffset.UtcNow.AddHours(1).ToUnixTimeSeconds(),
        }.ToJsonString()));
        var sig = SigningKey.SignData(Encoding.ASCII.GetBytes(header + "." + payload), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return header + "." + payload + "." + ChatGptAuth.B64Url(sig);
    }

    // ------------------------------------------------------------------ streams

    public static string ResponsesSse(string text)
    {
        var sb = new StringBuilder();
        foreach (var piece in Chunks(text))
            sb.Append("event: response.output_text.delta\ndata: ").Append(new JsonObject { ["type"] = "response.output_text.delta", ["delta"] = piece }.ToJsonString()).Append("\n\n");
        sb.Append("event: response.completed\ndata: {\"type\":\"response.completed\",\"response\":{\"status\":\"completed\"}}\n\n");
        return sb.ToString();
    }

    public static string ClaudeSse(string text)
    {
        var sb = new StringBuilder();
        sb.Append("event: message_start\ndata: {\"type\":\"message_start\",\"message\":{\"id\":\"msg_1\",\"type\":\"message\",\"role\":\"assistant\",\"model\":\"claude-opus-5-5\",\"content\":[],\"stop_reason\":null,\"stop_sequence\":null,\"usage\":{\"input_tokens\":10,\"output_tokens\":1}}}\n\n");
        sb.Append("event: content_block_start\ndata: {\"type\":\"content_block_start\",\"index\":0,\"content_block\":{\"type\":\"text\",\"text\":\"\"}}\n\n");
        foreach (var piece in Chunks(text))
            sb.Append("event: content_block_delta\ndata: ").Append(new JsonObject { ["type"] = "content_block_delta", ["index"] = 0, ["delta"] = new JsonObject { ["type"] = "text_delta", ["text"] = piece } }.ToJsonString()).Append("\n\n");
        sb.Append("event: content_block_stop\ndata: {\"type\":\"content_block_stop\",\"index\":0}\n\n");
        sb.Append("event: message_delta\ndata: {\"type\":\"message_delta\",\"delta\":{\"stop_reason\":\"end_turn\",\"stop_sequence\":null},\"usage\":{\"output_tokens\":5}}\n\n");
        sb.Append("event: message_stop\ndata: {\"type\":\"message_stop\"}\n\n");
        return sb.ToString();
    }

    public static string CompatSse(string text)
    {
        var sb = new StringBuilder();
        foreach (var piece in Chunks(text))
            sb.Append("data: ").Append(new JsonObject { ["choices"] = new JsonArray(new JsonObject { ["index"] = 0, ["delta"] = new JsonObject { ["content"] = piece } }) }.ToJsonString()).Append("\n\n");
        sb.Append("data: {\"choices\":[{\"index\":0,\"delta\":{},\"finish_reason\":\"stop\"}]}\n\n");
        sb.Append("data: [DONE]\n\n");
        return sb.ToString();
    }

    static IEnumerable<string> Chunks(string s)
    {
        for (int i = 0; i < s.Length; i += 7) yield return s.Substring(i, Math.Min(7, s.Length - i));
    }

    public void Dispose() { _cts.Cancel(); _l.Stop(); }
}
