using System.Text;
using System.Text.Json.Nodes;
using MicroadsOffice.AI.Mcp;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Tests;

public static class CoreTests
{
    static Dictionary<string, string> _lastAuthorize = [];

    /// <summary>Plays the browser: records the authorize request, then calls the app's loopback callback.</summary>
    static void Browser(FakeServer s, Func<Dictionary<string, string>, string>? callback = null)
    {
        ChatGptAuth.OpenBrowser = url =>
        {
            var u = new Uri(url);
            var q = ChatGptAuth.ParseQuery(u.Query.TrimStart('?'));
            _lastAuthorize = q;
            s.Challenge = q["code_challenge"]; s.Nonce = q["nonce"]; s.LastRedirect = q["redirect_uri"];
            var cb = callback?.Invoke(q) ?? DefaultCallback(s, q);
            _ = Task.Run(async () =>
            {
                await Task.Delay(80);
                using var h = new HttpClient();
                try { await h.GetAsync(q["redirect_uri"] + "?" + cb); } catch { }
            });
            return true;
        };
    }

    static string DefaultCallback(FakeServer s, Dictionary<string, string> q) =>
        "code=CODE-1&state=" + Uri.EscapeDataString(q["state"]) + "&scope=" + Uri.EscapeDataString(s.Scope) +
        (q["client_id"] == ChatGptAuth.DynamicClient ? "&client_id=" + s.IssuedClientId : "");

    static T2 Run<T2>(Func<Task<T2>> f) => Task.Run(f).GetAwaiter().GetResult();

    public static void Run(Runner r)
    {
        using var s = new FakeServer();
        Environment.SetEnvironmentVariable("MICROADS_AI_AUTH_BASE", s.Base);
        Environment.SetEnvironmentVariable("MICROADS_AI_OPENAI_BASE", s.Base + "/v1");
        Environment.SetEnvironmentVariable("MICROADS_AI_ANTHROPIC_BASE", s.Base);

        r.Test("C01 secrets are encrypted with DPAPI on disk and read back", () =>
        {
            SecretStore.Write("unit", "sk-bí-mật-123");
            var raw = File.ReadAllBytes(AiPaths.File("unit.bin"));
            T.False(Encoding.UTF8.GetString(raw).Contains("sk-bí-mật-123"), "plain text on disk");
            T.Eq("sk-bí-mật-123", SecretStore.Read("unit"));
            T.True(SecretStore.Has("unit"), "Has");
            SecretStore.Write("unit", null);
            T.False(File.Exists(AiPaths.File("unit.bin")), "removed");
            T.Eq(null, SecretStore.Read("unit"));
        });

        r.Test("C02 PKCE S256 matches the RFC 7636 test vector", () =>
            T.Eq("E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM", ChatGptAuth.Challenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk")));

        r.Test("C03 first sign-in request follows OpenAI's open-source contract", () =>
        {
            var url = ChatGptAuth.BuildAuthorizeUrl(ChatGptAuth.DynamicClient, "http://127.0.0.1:1455/auth/callback", "st", "no", "ch", "urn:uuid:x", null, null);
            var q = ChatGptAuth.ParseQuery(new Uri(url).Query.TrimStart('?'));
            T.True(url.StartsWith(s.Base + "/api/accounts/authorize?"), url);
            T.Eq("dynamic_agent_client", q["client_id"]);
            T.Eq("Microads Office", q["agent_name_hint"]);
            T.Eq("code", q["response_type"]);
            T.Eq("S256", q["code_challenge_method"]);
            T.Eq("https://api.openai.com/v1", q["resource"]);
            T.Eq("urn:uuid:x", q["ext_agent_host_id"]);
            foreach (var sc in new[] { "openid", "profile", "email", "offline_access", "resource.invoke", "chatgpt.tokens.use.direct" })
                T.True(q["scope"].Split(' ').Contains(sc), "scope " + sc);
            T.False(q.ContainsKey("id_token_hint"), "no hint on first sign-in");
        });

        r.Test("C04 Continue with ChatGPT: browser sign-in, PKCE, ID token verified, credential saved encrypted", () =>
        {
            ChatGptAuth.Forget();
            Browser(s);
            var statuses = new List<string>();
            var c = Run(() => ChatGptAuth.SignInAsync(false, statuses.Add, CancellationToken.None));
            T.Eq("oaiapp_test123", c.ClientId);
            T.Eq("user-1", c.Subject);
            T.Eq("nhanvien@example.com", c.Email);
            T.True(c.PlanUsage, "plan scope granted");
            T.True(ChatGptAuth.IsSignedIn, "signed in");
            T.True(_lastAuthorize["redirect_uri"].StartsWith("http://127.0.0.1:") && _lastAuthorize["redirect_uri"].EndsWith("/auth/callback"), "loopback callback " + _lastAuthorize["redirect_uri"]);
            T.True(c.ExtAgentHostId.StartsWith("urn:uuid:"), "host id");
            T.Eq(c.ExtAgentHostId, _lastAuthorize["ext_agent_host_id"]);
            T.True(statuses.Count >= 2, "status messages shown");
            var disk = Encoding.UTF8.GetString(File.ReadAllBytes(AiPaths.File(SecretNames.ChatGpt + ".bin")));
            T.False(disk.Contains("AT-1") || disk.Contains("RT-1") || disk.Contains("nhanvien"), "tokens not readable on disk");
            T.Eq("AT-1", Run(() => ChatGptAuth.GetAccessTokenAsync(CancellationToken.None)));
        });

        r.Test("C05 access token is refreshed near expiry and the rotating refresh token is stored", () =>
        {
            var c = ChatGptAuth.Load()!;
            c.ExpiresAt = DateTimeOffset.UtcNow.AddSeconds(30);
            SecretStore.Write(SecretNames.ChatGpt, System.Text.Json.JsonSerializer.Serialize(c));
            var t = Run(() => ChatGptAuth.GetAccessTokenAsync(CancellationToken.None));
            T.Eq("AT-2", t);
            var last = s.Last("/api/accounts/oauth/token")!;
            T.Eq("refresh_token", last.Form["grant_type"]);
            T.Eq("oaiapp_test123", last.Form["client_id"]);
            T.Eq("https://api.openai.com/v1", last.Form["resource"]);
            T.False(last.Form.ContainsKey("scope"), "scope omitted on refresh");
            T.Eq("RT-2", ChatGptAuth.Load()!.RefreshToken);
            T.True(ChatGptAuth.Load()!.ExpiresAt > DateTimeOffset.UtcNow.AddMinutes(50), "new expiry");
        });

        r.Test("C06 signing in again reuses the issued client id with id_token_hint and login_hint", () =>
        {
            Browser(s);
            var c = Run(() => ChatGptAuth.SignInAsync(false, null, CancellationToken.None));
            T.Eq("oaiapp_test123", _lastAuthorize["client_id"]);
            T.False(_lastAuthorize.ContainsKey("agent_name_hint"), "no agent_name_hint on reauthorization");
            T.True(_lastAuthorize.ContainsKey("id_token_hint"), "id_token_hint");
            T.Eq("nhanvien@example.com", _lastAuthorize["login_hint"]);
            T.Eq("oaiapp_test123", c.ClientId);
        });

        r.Test("C07 cancelling in the browser (access_denied) keeps the previous sign-in", () =>
        {
            int tokenCalls = s.Count("/api/accounts/oauth/token");
            Browser(s, q => "error=access_denied&state=" + Uri.EscapeDataString(q["state"]));
            var e = T.Throws<AiException>(() => Run(() => ChatGptAuth.SignInAsync(false, null, CancellationToken.None)), "cancel");
            T.Eq("access_denied", e.Code);
            T.Eq(tokenCalls, s.Count("/api/accounts/oauth/token"));
            T.True(ChatGptAuth.IsSignedIn, "still signed in");
        });

        r.Test("C08 a forged callback (wrong state, e.g. from another web page) is refused with 400 and ignored; the real one still signs in", () =>
        {
            int tokenCalls = s.Count("/api/accounts/oauth/token");
            int forgedStatus = 0;
            ChatGptAuth.OpenBrowser = url =>
            {
                var q = ChatGptAuth.ParseQuery(new Uri(url).Query.TrimStart('?'));
                s.Challenge = q["code_challenge"]; s.Nonce = q["nonce"]; s.LastRedirect = q["redirect_uri"];
                _ = Task.Run(async () =>
                {
                    await Task.Delay(80);
                    using var h = new HttpClient();
                    forgedStatus = (int)(await h.GetAsync(q["redirect_uri"] + "?code=EVIL&state=forged")).StatusCode;
                    await h.GetAsync(q["redirect_uri"] + "?" + DefaultCallback(s, q));
                });
                return true;
            };
            var c = Run(() => ChatGptAuth.SignInAsync(false, null, CancellationToken.None));
            T.Eq(400, forgedStatus);
            T.True(c.SignedIn, "real callback completed the sign-in");
            T.Eq(tokenCalls + 1, s.Count("/api/accounts/oauth/token"));
            T.False(s.Requests.Any(x => x.Path == "/api/accounts/oauth/token" && x.Form.GetValueOrDefault("code") == "EVIL"), "forged code never exchanged");
        });

        r.Test("C09 ID tokens with a bad signature, nonce or audience are rejected", () =>
        {
            var nonce = "n-123";
            s.Nonce = nonce;
            T.Eq("user-1", Run(() => ChatGptAuth.ValidateIdTokenAsync(s.IdToken(nonce), s.IssuedClientId, nonce, CancellationToken.None))["sub"]!.GetValue<string>());
            T.Eq("invalid_id_token_nonce", T.Throws<AiException>(() => Run(() => ChatGptAuth.ValidateIdTokenAsync(s.IdToken("other"), s.IssuedClientId, nonce, CancellationToken.None)), "nonce").Code);
            T.Eq("invalid_id_token_audience", T.Throws<AiException>(() => Run(() => ChatGptAuth.ValidateIdTokenAsync(s.IdToken(nonce, aud: "someone-else"), s.IssuedClientId, nonce, CancellationToken.None)), "aud").Code);
            T.Eq("invalid_id_token_issuer", T.Throws<AiException>(() => Run(() => ChatGptAuth.ValidateIdTokenAsync(s.IdToken(nonce, iss: "https://evil.example"), s.IssuedClientId, nonce, CancellationToken.None)), "iss").Code);
            s.SigningKey = System.Security.Cryptography.RSA.Create(2048);
            try { T.Eq("invalid_id_token_signature", T.Throws<AiException>(() => Run(() => ChatGptAuth.ValidateIdTokenAsync(s.IdToken(nonce), s.IssuedClientId, nonce, CancellationToken.None)), "sig").Code); }
            finally { s.SigningKey = s.Key; }
        });

        r.Test("C10 sign-in without the plan permission keeps the account but blocks requests with a clear message", () =>
        {
            s.Scope = "email offline_access openid profile";
            try
            {
                Browser(s);
                var c = Run(() => ChatGptAuth.SignInAsync(false, null, CancellationToken.None));
                T.False(c.PlanUsage, "plan not granted");
                var e = T.Throws<AiException>(() => Run(() => ChatGptAuth.GetAccessTokenAsync(CancellationToken.None)), "blocked");
                T.Eq("plan_not_enabled", e.Code);
                T.True(e.NeedsSignIn, "offers sign-in");
                T.False(AiService.IsReady(new AiSettings { Provider = AiProviderKind.ChatGpt }), "not ready");
            }
            finally { s.Scope = "chatgpt.tokens.use.direct email offline_access openid profile resource.invoke"; }
            Browser(s);
            T.True(Run(() => ChatGptAuth.SignInAsync(false, null, CancellationToken.None)).PlanUsage, "granted again");
        });

        r.Test("C11 a revoked refresh token signs the user out (account and client id are kept)", () =>
        {
            var c = ChatGptAuth.Load()!;
            c.ExpiresAt = DateTimeOffset.UtcNow; c.RefreshToken = "RT-revoked";
            SecretStore.Write(SecretNames.ChatGpt, System.Text.Json.JsonSerializer.Serialize(c));
            var e = T.Throws<AiException>(() => Run(() => ChatGptAuth.GetAccessTokenAsync(CancellationToken.None)), "refresh fails");
            T.True(e.NeedsSignIn, "needs sign-in");
            var after = ChatGptAuth.Load()!;
            T.False(after.SignedIn, "tokens cleared");
            T.Eq("oaiapp_test123", after.ClientId);
            T.Eq("nhanvien@example.com", after.Email);
            Browser(s);
            T.True(Run(() => ChatGptAuth.SignInAsync(false, null, CancellationToken.None)).SignedIn, "sign in again works");
        });

        r.Test("C12 Responses API (ChatGPT plan): streamed answer, store=false, stream=true, no unsupported fields", () =>
        {
            s.Reply = "Xin chào, tôi là trợ lý.";
            var p = AiService.Create(new AiSettings { Provider = AiProviderKind.ChatGpt })!;
            var deltas = new List<string>();
            var text = Run(() => p.CompleteAsync(new AiRequest("Be brief.", [new AiMessage(AiRole.User, "Chào")], "gpt-test"), deltas.Add, CancellationToken.None));
            T.Eq("Xin chào, tôi là trợ lý.", text);
            T.True(deltas.Count > 1, "streamed in pieces");
            var req = s.Last("/v1/responses")!;
            var j = req.Json!;
            T.Eq(false, j["store"]!.GetValue<bool>());
            T.Eq(true, j["stream"]!.GetValue<bool>());
            T.Eq("Be brief.", j["instructions"]!.GetValue<string>());
            T.Eq("gpt-test", j["model"]!.GetValue<string>());
            foreach (var bad in new[] { "temperature", "max_output_tokens", "metadata", "user", "previous_response_id", "top_p" }) T.False(j.AsObject().ContainsKey(bad), bad);
            T.False(j["input"]!.AsArray().Any(m => m!["role"]!.GetValue<string>() == "system"), "no system role items");
            T.True(req.Headers["Authorization"].StartsWith("Bearer AT-"), "bearer access token");
            s.Reply = null;
        });

        r.Test("C13 ChatGPT plan errors become clear messages (not eligible, usage limit, interrupted stream)", () =>
        {
            var p = AiService.Create(new AiSettings { Provider = AiProviderKind.ChatGpt })!;
            var req = new AiRequest("", [new AiMessage(AiRole.User, "x")], "gpt-test");
            s.Override = q => q.Path == "/v1/responses" ? FakeResponse.Json("""{"error":{"code":"subscription_sharing_user_not_eligible","message":"no"}}""", 403) : null;
            var e1 = T.Throws<AiException>(() => Run(() => p.CompleteAsync(req, _ => { }, CancellationToken.None)), "403");
            T.Eq("subscription_sharing_user_not_eligible", e1.Code);
            T.True(e1.Message.Contains("Plus") || e1.Message.Contains("Pro"), e1.Message);
            s.Override = q => q.Path == "/v1/responses" ? FakeResponse.Sse("data: {\"type\":\"response.output_text.delta\",\"delta\":\"A\"}\n\ndata: {\"type\":\"response.failed\",\"response\":{\"error\":{\"code\":\"subscription_sharing_usage_limit_exceeded\",\"message\":\"limit\"}}}\n\n") : null;
            var e2 = T.Throws<AiException>(() => Run(() => p.CompleteAsync(req, _ => { }, CancellationToken.None)), "limit");
            T.Eq(AiDefaults.ChatGptUsage, e2.HelpUrl);
            s.Override = q => q.Path == "/v1/responses" ? FakeResponse.Sse("data: {\"type\":\"response.output_text.delta\",\"delta\":\"A\"}\n\n") : null;
            T.Eq("stream_interrupted", T.Throws<AiException>(() => Run(() => p.CompleteAsync(req, _ => { }, CancellationToken.None)), "cut").Code);
            s.Override = q => q.Path == "/v1/responses" ? FakeResponse.Json("""{"detail":"Direct routing is unavailable"}""", 503) : null;
            T.True(T.Throws<AiException>(() => Run(() => p.CompleteAsync(req, _ => { }, CancellationToken.None)), "503").Message.Contains("Direct routing"), "detail kept");
            s.Override = null;
        });

        r.Test("C14 model lists: ChatGPT catalogue hides non-listed models; API list prefers full GPT models", () =>
        {
            var plan = AiService.Create(new AiSettings { Provider = AiProviderKind.ChatGpt })!;
            var m = Run(() => plan.ListModelsAsync(CancellationToken.None));
            T.Eq(1, m.Count); T.Eq("gpt-test", m[0].Id); T.Eq("GPT Test", m[0].DisplayName);
            SecretStore.Write(SecretNames.OpenAiKey, "sk-test-openai");
            var key = AiService.Create(new AiSettings { Provider = AiProviderKind.OpenAiKey })!;
            var km = Run(() => key.ListModelsAsync(CancellationToken.None));
            T.False(km.Any(x => x.Id.Contains("embedding")), "no embedding model");
            T.Eq("gpt-new", key.DefaultModel(km));
            T.Eq("Xin chào từ ChatGPT.", Run(() => key.CompleteAsync(new AiRequest("", [new AiMessage(AiRole.User, "x")], "gpt-new"), _ => { }, CancellationToken.None)));
            SecretStore.Write(SecretNames.OpenAiKey, "sk-wrong");
            T.True(T.Throws<AiException>(() => Run(() => AiService.Create(new AiSettings { Provider = AiProviderKind.OpenAiKey })!.ListModelsAsync(CancellationToken.None)), "bad key").NeedsSignIn, "bad key → re-enter");
            SecretStore.Write(SecretNames.OpenAiKey, null);
        });

        r.Test("C15 Claude (official SDK): streamed answer, API key header, fallbacks + effort sent, model list, bad key", () =>
        {
            SecretStore.Write(SecretNames.ClaudeKey, "sk-ant-test");
            s.Reply = "Chào bạn, tôi là Claude.";
            var p = AiService.Create(new AiSettings { Provider = AiProviderKind.ClaudeKey })!;
            var deltas = new List<string>();
            var text = Run(() => p.CompleteAsync(new AiRequest("Sys", [new AiMessage(AiRole.User, "Chào"), new AiMessage(AiRole.Assistant, "Chào!"), new AiMessage(AiRole.User, "Tiếp")], "claude-opus-5-5"), deltas.Add, CancellationToken.None));
            T.Eq("Chào bạn, tôi là Claude.", text);
            T.True(deltas.Count > 1, "streamed");
            var req = s.Last("/v1/messages")!;
            T.Eq("Bearer sk-ant-test", req.Headers["Authorization"]); // API version 2026-10-21: API keys travel as Authorization: Bearer
            T.Eq("2026-10-21", req.Headers["anthropic-version"]);
            T.True(req.Headers.GetValueOrDefault("anthropic-beta", "").Contains("server-side-fallback-2026-07-01"), "beta header: " + req.Headers.GetValueOrDefault("anthropic-beta"));
            var j = req.Json!;
            T.Eq("claude-opus-5-5", j["model"]!.GetValue<string>());
            T.Eq("default", j["fallbacks"]?.ToString());
            T.Eq("medium", j["output_config"]?["effort"]?.ToString());
            T.Eq("Sys", j["system"]?.ToString());
            T.Eq(3, j["messages"]!.AsArray().Count);
            T.True(j["stream"]?.GetValue<bool>() == true, "stream");
            T.False(j.AsObject().ContainsKey("temperature"), "no temperature");
            var models = Run(() => p.ListModelsAsync(CancellationToken.None));
            T.True(models.Any(m => m.Id == "claude-opus-5-5"), "models");
            T.Eq("claude-opus-5-5", p.DefaultModel(models));
            SecretStore.Write(SecretNames.ClaudeKey, "sk-ant-wrong");
            T.True(T.Throws<AiException>(() => Run(() => AiService.Create(new AiSettings { Provider = AiProviderKind.ClaudeKey })!.CompleteAsync(new AiRequest("", [new AiMessage(AiRole.User, "x")], "claude-opus-5-5"), _ => { }, CancellationToken.None)), "bad key").NeedsSignIn, "re-enter key");
            SecretStore.Write(SecretNames.ClaudeKey, null);
            s.Reply = null;
        });

        r.Test("C16 OpenAI-compatible server (Ollama/internal): streaming chat completions and model list", () =>
        {
            var st = new AiSettings { Provider = AiProviderKind.Compatible, CompatibleBaseUrl = s.Base + "/compat/v1", CompatibleModel = "qwen-local" };
            T.True(AiService.IsReady(st), "ready with url + model");
            var p = AiService.Create(st)!;
            T.Eq("Trả lời từ máy chủ nội bộ.", Run(() => p.CompleteAsync(new AiRequest("Sys", [new AiMessage(AiRole.User, "x")], "qwen-local"), _ => { }, CancellationToken.None)));
            var j = s.Last("/compat/v1/chat/completions")!.Json!;
            T.Eq("system", j["messages"]![0]!["role"]!.GetValue<string>());
            T.Eq(2, Run(() => p.ListModelsAsync(CancellationToken.None)).Count);
            T.False(AiService.IsReady(new AiSettings { Provider = AiProviderKind.Compatible, CompatibleBaseUrl = s.Base, CompatibleModel = "" }), "needs a model");
        });

        r.Test("C17 sign out revokes the refresh token and clears tokens", () =>
        {
            var rt = ChatGptAuth.Load()!.RefreshToken;
            T.True(Run(() => ChatGptAuth.SignOutAsync(CancellationToken.None)), "revocation confirmed");
            T.True(s.Revoked.Contains(rt), "server received the refresh token");
            T.False(ChatGptAuth.IsSignedIn, "signed out");
            T.Eq("nhanvien@example.com", ChatGptAuth.Load()!.Email);
            T.Eq("not_signed_in", T.Throws<AiException>(() => Run(() => ChatGptAuth.GetAccessTokenAsync(CancellationToken.None)), "no token").Code);
        });

        r.Test("C18 Markdown: headings, alignment blocks, lists, tables, emphasis, code, page break", () =>
        {
            var b = Markdown.Parse(":::center\nCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n**Độc lập - Tự do - Hạnh phúc**\n:::\n\n# HỢP ĐỒNG\n\nĐiều *một* có **chữ đậm** và `mã`.\n\n- Ý 1\n  - Ý con\n1. Bước một\n\n| STT | Tên |\n|---|---|\n| 1 | A \\| B |\n\n```\ncode\n```\n\\pagebreak\n> trích dẫn");
            var p0 = (MdParagraph)b[0];
            T.Eq(MdAlign.Center, p0.Align);
            T.True(Markdown.Plain(p0.Spans).Contains('\v'), "line break inside centred block");
            T.True(p0.Spans.Any(s => s.Bold && s.Text.StartsWith("Độc lập")), "bold motto");
            T.Eq(1, ((MdHeading)b[1]).Level);
            var para = (MdParagraph)b[2];
            T.True(para.Spans.Any(x => x.Italic && x.Text == "một") && para.Spans.Any(x => x.Bold && x.Text == "chữ đậm") && para.Spans.Any(x => x.Code && x.Text == "mã"), "inline styles");
            var lists = b.OfType<MdListItem>().ToList();
            T.Eq(3, lists.Count); T.Eq(1, lists[1].Level); T.True(lists[2].Ordered, "ordered");
            var t = b.OfType<MdTable>().Single();
            T.Eq(2, t.Rows.Count); T.Eq("A | B", Markdown.Plain(t.Rows[1][1]));
            T.Eq("code", b.OfType<MdCode>().Single().Text);
            T.Eq(1, b.OfType<MdPageBreak>().Count());
            T.Eq(1, b.OfType<MdQuote>().Count());
            T.True(Markdown.ToPlainText("# T\n\n**a** *b*").Contains("a b"), "plain text");
            T.Eq("2 * 3 = 6", Markdown.Plain(Markdown.Inline("2 * 3 = 6")));
        });

        r.Test("C19 action blocks are split from the chat text; live steps follow the headings", () =>
        {
            var ans = "Tôi đã soạn hợp đồng.\n```microads-document mode=replace_selection\n# Điều 1\nNội dung\n```python\nx=1\n```\n## Điều 2\n```\nXong.";
            var (blocks, chat) = ActionBlocks.Extract(ans);
            T.Eq(1, blocks.Count);
            T.Eq("replace_selection", blocks[0].Mode);
            T.Eq("document", blocks[0].Kind);
            T.True(blocks[0].Body.Contains("x=1") && blocks[0].Body.Contains("## Điều 2"), "nested fence kept");
            T.Eq("Tôi đã soạn hợp đồng.\nXong.", chat);
            var steps = ActionBlocks.Steps("Ok\n```microads-document\n# Điều 1\nabc\n## Điều 2\n");
            T.Eq(2, steps.Count); T.Eq("Điều 2", steps[1]);
            var sheet = ActionBlocks.Extract("```microads-sheet\n{\"actions\":[{\"cell\":\"A1\",\"value\":1}]}\n```");
            T.Eq("sheet", sheet.Blocks[0].Kind);
            T.True(ActionBlocks.Steps("```microads-sheet\n{\"actions\":[{\"cell\":\"A1\",\"value\":1},{\"range\":\"A2\"").Last().Contains('2'), "sheet steps count");
        });

        r.Test("C20 MCP server: initialize, tool list, unknown method, notifications get no reply", () =>
        {
            var m = new McpServer(TextReader.Null, TextWriter.Null, _ => false);
            var init = Run(() => m.HandleAsync(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "initialize", ["params"] = new JsonObject { ["protocolVersion"] = "2025-06-18", ["capabilities"] = new JsonObject(), ["clientInfo"] = new JsonObject { ["name"] = "claude-ai", ["version"] = "1" } } }, CancellationToken.None))!;
            T.Eq("2025-06-18", init["result"]!["protocolVersion"]!.GetValue<string>());
            T.Eq("microads-office", init["result"]!["serverInfo"]!["name"]!.GetValue<string>());
            T.Eq(null, Run(() => m.HandleAsync(new JsonObject { ["jsonrpc"] = "2.0", ["method"] = "notifications/initialized" }, CancellationToken.None)));
            var tools = Run(() => m.HandleAsync(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = 2, ["method"] = "tools/list" }, CancellationToken.None))!["result"]!["tools"]!.AsArray();
            var names = tools.Select(t => t!["name"]!.GetValue<string>()).ToList();
            foreach (var n in new[] { "list_open_apps", "open_app", "read_document", "get_selection", "get_editing_guide", "write_document", "edit_sheet", "undo" }) T.True(names.Contains(n), n);
            T.Eq(-32601, Run(() => m.HandleAsync(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = 3, ["method"] = "nope" }, CancellationToken.None))!["error"]!["code"]!.GetValue<int>());
            // a line-based round trip through RunAsync
            var input = new StringReader("{\"jsonrpc\":\"2.0\",\"id\":7,\"method\":\"ping\"}\n");
            var output = new StringWriter();
            Run(async () => { await new McpServer(input, output, _ => false).RunAsync(CancellationToken.None); return 0; });
            T.True(output.ToString().Contains("\"id\":7") && output.ToString().Contains("\"result\""), output.ToString());
        });

        r.Test("C21 Claude Desktop ↔ open document: refused until allowed, then read / write / undo through the bridge", () =>
        {
            var host = new FakeHost();
            using var ui = new UiThread();
            ui.Run(() => AiBridge.Start(AppKind.Word, () => host, ui.Control));
            Thread.Sleep(200);
            var m = new McpServer(TextReader.Null, TextWriter.Null, _ => false);
            JsonNode Call(string tool, JsonObject args) => Run(() => m.HandleAsync(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = 9, ["method"] = "tools/call", ["params"] = new JsonObject { ["name"] = tool, ["arguments"] = args } }, CancellationToken.None))!["result"]!;
            string Text(JsonNode res) => res["content"]![0]!["text"]!.GetValue<string>();
            AiSettings.Update(x => x.AllowClaudeDesktop = false);
            var refused = Call("read_document", new JsonObject { ["app"] = "word" });
            T.True(refused["isError"]!.GetValue<bool>(), "refused");
            T.True(Text(refused).Contains("Claude Desktop"), Text(refused));
            AiSettings.Update(x => x.AllowClaudeDesktop = true);
            var list = Text(Call("list_open_apps", []));
            T.True(list.Contains("word: open") && list.Contains("Hợp đồng.docx") && list.Contains("excel: not open"), list);
            T.True(Text(Call("read_document", new JsonObject { ["app"] = "word" })).Contains("# Nội dung giả"), "read");
            T.Eq("Đoạn chọn", Text(Call("get_selection", new JsonObject { ["app"] = "word" })));
            T.True(Text(Call("get_editing_guide", new JsonObject { ["app"] = "word" })).Contains("GUIDE"), "guide");
            var w = Call("write_document", new JsonObject { ["content"] = "## Mới", ["mode"] = "append" });
            T.False(w["isError"]!.GetValue<bool>(), Text(w));
            T.Eq("append|## Mới", host.Applied.Single());
            Call("undo", new JsonObject { ["app"] = "word" });
            T.Eq(1, host.Undos);
            T.True(Call("read_document", new JsonObject { ["app"] = "excel" })["isError"]!.GetValue<bool>(), "excel not open → error");
            T.True(ui.Thread != Thread.CurrentThread && host.AppliedOnThread == ui.Thread, "edits run on the app's UI thread");
            AiBridge.Stop(AppKind.Word);
            AiSettings.Update(x => x.AllowClaudeDesktop = false);
        });

        r.Test("C22 connecting Claude Desktop keeps the user's other MCP servers and makes a backup", () =>
        {
            var root = Environment.GetEnvironmentVariable("MICROADS_CLAUDE_CONFIG_ROOT")!;
            var dir = Path.Combine(root, "Roaming", "Claude");
            Directory.CreateDirectory(dir);
            var f = Path.Combine(dir, "claude_desktop_config.json");
            File.WriteAllText(f, "{\n  // user comment\n  \"mcpServers\": { \"other\": { \"command\": \"x.exe\" } },\n  \"theme\": \"dark\",\n}");
            ClaudeDesktop.ExePath = @"C:\Program Files\Microads Office\MicroadsOffice.exe";
            T.False(ClaudeDesktop.IsConnected(), "not yet");
            ClaudeDesktop.Connect();
            var j = JsonNode.Parse(File.ReadAllText(f))!;
            T.Eq("x.exe", j["mcpServers"]!["other"]!["command"]!.GetValue<string>());
            T.Eq("dark", j["theme"]!.GetValue<string>());
            T.Eq(ClaudeDesktop.ExePath, j["mcpServers"]!["microads-office"]!["command"]!.GetValue<string>());
            T.Eq("--mcp", j["mcpServers"]!["microads-office"]!["args"]![0]!.GetValue<string>());
            T.True(File.Exists(f + ".microads-backup"), "backup");
            T.True(ClaudeDesktop.IsConnected(), "connected");
            ClaudeDesktop.Disconnect();
            T.False(ClaudeDesktop.IsConnected(), "disconnected");
            T.Eq("x.exe", JsonNode.Parse(File.ReadAllText(f))!["mcpServers"]!["other"]!["command"]!.GetValue<string>());
            T.False(Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Claude", "claude_desktop_config.json.microads-backup")), "real config untouched");
        });

        r.Test("C23 settings are saved in office.json; secrets never are", () =>
        {
            AiSettings.Update(x => { x.Provider = AiProviderKind.ClaudeKey; x.ClaudeModel = "claude-sonnet-5-5"; x.Instructions = "Văn phong trang trọng"; });
            SecretStore.Write(SecretNames.ClaudeKey, "sk-ant-secret-xyz");
            var s2 = AiSettings.Load();
            T.Eq(AiProviderKind.ClaudeKey, s2.Provider);
            T.Eq("claude-sonnet-5-5", s2.ModelFor(AiProviderKind.ClaudeKey));
            T.Eq("Văn phong trang trọng", s2.Instructions);
            var json = File.ReadAllText(Path.Combine(AppSettings.DataDirectory, "office.json"));
            T.False(json.Contains("sk-ant-secret-xyz"), "key not in office.json");
            T.True(AiService.IsReady(s2), "ready with key");
            SecretStore.Write(SecretNames.ClaudeKey, null);
            T.False(AiService.IsReady(AiSettings.Load()), "not ready without key");
            T.False(AiService.IsReady(new AiSettings()), "None is never ready");
            AiSettings.Update(x => x.Provider = AiProviderKind.None);
        });

        r.Test("C24 prompts: document is marked untrusted, action guide only for editable hosts, context carries selection", () =>
        {
            var h = new FakeHost();
            var ins = Prompts.Instructions(h, new AiSettings { Instructions = "Luôn xưng hô anh/chị" });
            T.True(ins.Contains("untrusted"), "injection warning");
            T.True(ins.Contains("GUIDE"), "action guide");
            T.True(ins.Contains("Luôn xưng hô anh/chị"), "standing instructions");
            h.Editable = false;
            T.False(Prompts.Instructions(h, new AiSettings()).Contains("GUIDE"), "no guide for read-only");
            var ctx = Prompts.WithContext(h, "Câu hỏi?", true).Replace("\r\n", "\n");
            T.True(ctx.Contains("<document name=\"Hợp đồng.docx\">") && ctx.Contains("<selection>\nĐoạn chọn") && ctx.EndsWith("Câu hỏi?"), ctx);
            T.Eq("Câu hỏi?", Prompts.WithContext(h, "Câu hỏi?", false));
        });

        r.Test("C25 Store edition: Claude Desktop gets a paste-in configuration that starts the execution alias with --mcp", () =>
        {
            Environment.SetEnvironmentVariable("MICROADS_CLAUDE_MANUAL", "1");
            try
            {
                T.False(ClaudeDesktop.CanEditConfig, "manual mode");
                ClaudeDesktop.ExePath = @"C:\Users\u\AppData\Local\Microsoft\WindowsApps\MicroadsOffice.exe";
                var j = JsonNode.Parse(ClaudeDesktop.ConfigSnippet())!;
                T.Eq(ClaudeDesktop.ExePath, j["mcpServers"]!["microads-office"]!["command"]!.GetValue<string>());
                T.Eq("--mcp", j["mcpServers"]!["microads-office"]!["args"]![0]!.GetValue<string>());
            }
            finally { Environment.SetEnvironmentVariable("MICROADS_CLAUDE_MANUAL", null); }
            T.True(ClaudeDesktop.CanEditConfig, "classic edition edits the file itself");
            var manifest = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tools", "store", "AppxManifest.template.xml"));
            T.True(manifest.Contains("<uap5:ExecutionAlias Alias=\"MicroadsOffice.exe\" />"), "MSIX declares the execution alias");
        });

        r.Test("C26 Claude Desktop 4.x (Store): config goes to %LOCALAPPDATA%\\Claude-Data\\claude_<pkg>, not the folder it left", () =>
        {
            var root = Environment.GetEnvironmentVariable("MICROADS_CLAUDE_CONFIG_ROOT")!;
            foreach (var sub in new[] { "Roaming", "Local" }) { var p = Path.Combine(root, sub); if (Directory.Exists(p)) Directory.Delete(p, true); }
            var left = Path.Combine(root, "Local", "Packages", "Claude_pzs8sxrjxfjjc", "LocalCache", "Roaming", "Claude");
            Directory.CreateDirectory(left);
            File.WriteAllText(Path.Combine(left, ".user-data-left.claude_pzs8sxrjxfjjc"), "");
            var data = Path.Combine(root, "Local", "Claude-Data", "claude_pzs8sxrjxfjjc");
            Directory.CreateDirectory(data);
            File.WriteAllText(Path.Combine(data, "claude_desktop_config.json"), "{\"coworkUserFilesPath\":\"C:\\\\Users\\\\u\\\\Claude\",\"preferences\":{\"sidebarMode\":\"chat\"}}");
            T.Eq(1, ClaudeDesktop.ConfigDirs().Count);
            T.Eq(data, ClaudeDesktop.ConfigDirs()[0]);
            T.True(ClaudeDesktop.IsInstalled(), "detected");
            ClaudeDesktop.Connect();
            var j = JsonNode.Parse(File.ReadAllText(Path.Combine(data, "claude_desktop_config.json")))!;
            T.Eq("chat", j["preferences"]!["sidebarMode"]!.GetValue<string>());
            T.True(j["mcpServers"]!["microads-office"] != null, "server added where Claude reads it");
            T.False(File.Exists(Path.Combine(left, "claude_desktop_config.json")), "nothing written to the abandoned folder");
            T.True(ClaudeDesktop.IsConnected(), "connected");
            ClaudeDesktop.Disconnect();
            T.False(ClaudeDesktop.IsConnected(), "disconnected");
        });

        r.Test("C27 cut-off answers are never usable: unclosed blocks, streams without an end, Claude stops other than end_turn", () =>
        {
            var (blocks, _) = ActionBlocks.Extract("Đây\n```microads-document mode=replace_document\n# Tiêu đề\nĐoạn bị cắt");
            T.False(blocks.Single().Closed, "unclosed block flagged");
            T.True(ActionBlocks.Extract("```microads-document\nx\n```").Blocks.Single().Closed, "closed block ok");
            var compat = new CompatibleProvider(s.Base + "/compat/v1", null);
            s.Override = q => q.Path == "/compat/v1/chat/completions" ? FakeResponse.Sse("data: {\"choices\":[{\"index\":0,\"delta\":{\"content\":\"nửa câu\"}}]}\n\n") : null;
            T.Eq("stream_interrupted", T.Throws<AiException>(() => Run(() => compat.CompleteAsync(new AiRequest("", [new AiMessage(AiRole.User, "x")], "m"), _ => { }, CancellationToken.None)), "no [DONE]").Code);
            SecretStore.Write(SecretNames.ClaudeKey, "sk-ant-test");
            var claude = AiService.Create(new AiSettings { Provider = AiProviderKind.ClaudeKey })!;
            var req = new AiRequest("", [new AiMessage(AiRole.User, "x")], "claude-opus-5-5");
            s.Override = q => q.Path == "/v1/messages" ? FakeResponse.Sse(FakeServer.ClaudeSse("Một phần").Replace("\"end_turn\"", "\"max_tokens\"")) : null;
            T.Eq("max_tokens", T.Throws<AiException>(() => Run(() => claude.CompleteAsync(req, _ => { }, CancellationToken.None)), "max_tokens").Code);
            s.Override = q => q.Path == "/v1/messages" ? FakeResponse.Sse(FakeServer.ClaudeSse("Một phần").Replace("\"end_turn\"", "\"refusal\"")) : null;
            T.Eq("refusal", T.Throws<AiException>(() => Run(() => claude.CompleteAsync(req, _ => { }, CancellationToken.None)), "refusal with text").Code);
            s.Override = q => q.Path == "/v1/messages" ? FakeResponse.Sse(FakeServer.ClaudeSse("Một phần").Split("event: message_delta")[0]) : null;
            T.Eq("stream_interrupted", T.Throws<AiException>(() => Run(() => claude.CompleteAsync(req, _ => { }, CancellationToken.None)), "no stop").Code);
            s.Override = null;
            SecretStore.Write(SecretNames.ClaudeKey, null);
        });

        r.Test("C28 malformed MCP messages get JSON-RPC errors and the server keeps running", () =>
        {
            var input = new StringReader(
                "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":42}\n" +
                "{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/call\",\"params\":{\"name\":[1,2]}}\n" +
                "{\"jsonrpc\":\"2.0\",\"id\":3,\"method\":\"initialize\",\"params\":{\"protocolVersion\":7}}\n" +
                "not json\n" +
                "{\"jsonrpc\":\"2.0\",\"id\":4,\"method\":\"ping\"}\n");
            var output = new StringWriter();
            Run(async () => { await new McpServer(input, output, _ => false).RunAsync(CancellationToken.None); return 0; });
            var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(l => JsonNode.Parse(l)!).ToList();
            T.Eq(-32600, lines.Single(l => l["id"]?.GetValue<int>() == 1)["error"]!["code"]!.GetValue<int>());
            T.Eq(-32602, lines.Single(l => l["id"]?.GetValue<int>() == 2)["error"]!["code"]!.GetValue<int>());
            T.Eq(-32602, lines.Single(l => l["id"]?.GetValue<int>() == 3)["error"]!["code"]!.GetValue<int>());
            T.True(lines.Any(l => l["error"]?["code"]?.GetValue<int>() == -32700), "parse error reported");
            T.True(lines.Single(l => l["id"]?.GetValue<int>() == 4)["result"] != null, "still answering after bad input");
            var undo = McpServer.Tools().Single(t => t!["name"]!.GetValue<string>() == "undo")!;
            T.True(undo["annotations"]!["destructiveHint"]!.GetValue<bool>(), "undo is marked as changing the document");
        });

        r.Test("C29 Claude Desktop config: a file we do not understand is never replaced; the original is backed up once", () =>
        {
            var root = Environment.GetEnvironmentVariable("MICROADS_CLAUDE_CONFIG_ROOT")!;
            foreach (var sub in new[] { "Roaming", "Local" }) { var p = Path.Combine(root, sub); if (Directory.Exists(p)) Directory.Delete(p, true); }
            var dir = Path.Combine(root, "Roaming", "Claude");
            Directory.CreateDirectory(dir);
            var f = Path.Combine(dir, "claude_desktop_config.json");
            File.WriteAllText(f, "[1,2,3]");
            T.Throws<InvalidOperationException>(() => ClaudeDesktop.Connect(), "array root refused");
            T.Eq("[1,2,3]", File.ReadAllText(f));
            File.WriteAllText(f, "{\"mcpServers\":\"oops\"}");
            T.Throws<InvalidOperationException>(() => ClaudeDesktop.Connect(), "non-object mcpServers refused");
            File.Delete(f + ".microads-backup");
            File.WriteAllText(f, "{\"theme\":\"original\"}");
            ClaudeDesktop.Connect();
            ClaudeDesktop.Disconnect();
            ClaudeDesktop.Connect();
            T.Eq("{\"theme\":\"original\"}", File.ReadAllText(f + ".microads-backup"));
            ClaudeDesktop.Disconnect();
        });

        r.Test("C30 self-repair: an allowed connection that Claude lost (update / Cowork rewrite / exe moved) is written again", () =>
        {
            var root = Environment.GetEnvironmentVariable("MICROADS_CLAUDE_CONFIG_ROOT")!;
            foreach (var sub in new[] { "Roaming", "Local" }) { var p = Path.Combine(root, sub); if (Directory.Exists(p)) Directory.Delete(p, true); }
            var data = Path.Combine(root, "Local", "Claude-Data", "claude_x");
            Directory.CreateDirectory(data);
            File.WriteAllText(Path.Combine(data, "claude_desktop_config.json"), "{\"preferences\":{}}");
            AiSettings.Update(x => x.AllowClaudeDesktop = false);
            T.False(ClaudeDesktop.EnsureConnected(), "nothing written while not allowed");
            AiSettings.Update(x => x.AllowClaudeDesktop = true);
            ClaudeDesktop.ExePath = @"E:\Old\MicroadsOffice.exe";
            T.True(ClaudeDesktop.EnsureConnected(), "repaired");
            T.False(ClaudeDesktop.EnsureConnected(), "no rewrite when already right");
            ClaudeDesktop.ExePath = @"E:\New\MicroadsOffice.exe";
            T.True(ClaudeDesktop.EnsureConnected(), "follows a moved executable");
            T.Eq(@"E:\New\MicroadsOffice.exe", JsonNode.Parse(File.ReadAllText(Path.Combine(data, "claude_desktop_config.json")))!["mcpServers"]!["microads-office"]!["command"]!.GetValue<string>());
            ClaudeDesktop.ClearHandshake();
            T.Eq(null, ClaudeDesktop.LastHandshake());
            var m = new McpServer(TextReader.Null, TextWriter.Null, _ => false);
            Run(() => m.HandleAsync(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "initialize", ["params"] = new JsonObject { ["protocolVersion"] = "2025-06-18", ["clientInfo"] = new JsonObject { ["name"] = "claude-ai", ["version"] = "4.23" } } }, CancellationToken.None));
            var hs = ClaudeDesktop.LastHandshake()!;
            T.Eq("claude-ai", hs.Client);
            T.True((DateTimeOffset.Now - hs.Time).TotalMinutes < 1, "handshake recorded now");
            AiSettings.Update(x => x.AllowClaudeDesktop = false);
            ClaudeDesktop.Disconnect();
        });

        r.Test("C31 test overrides cannot redirect tokens or keys to another host", () =>
        {
            Environment.SetEnvironmentVariable("MICROADS_AI_X_TEST", "https://evil.example.com/v1");
            T.Eq(null, ChatGptAuth.SafeOverride("MICROADS_AI_X_TEST"));
            Environment.SetEnvironmentVariable("MICROADS_AI_X_TEST", "http://127.0.0.1:5000/v1/");
            T.Eq("http://127.0.0.1:5000/v1", ChatGptAuth.SafeOverride("MICROADS_AI_X_TEST"));
            Environment.SetEnvironmentVariable("MICROADS_AI_X_TEST", null);
        });

        ChatGptAuth.OpenBrowser = _ => false;
    }
}

/// <summary>A document host that records what the assistant does.</summary>
public sealed class FakeHost : IAiHost
{
    public readonly List<string> Applied = [];
    public int Undos;
    public bool Editable = true;
    public Thread? AppliedOnThread;
    public AppKind Kind => AppKind.Word;
    public Control Owner { get; } = new();
    public string DocumentName => "Hợp đồng.docx";
    public string? SelectionText => "Đoạn chọn";
    public string GetDocumentContext(int maxChars) => "# Nội dung giả\nĐoạn văn.";
    public string ActionInstructions => "GUIDE";
    public IReadOnlyList<AiQuickAction> QuickActions => [];
    public bool CanEdit => Editable;
    public AiApplyResult Apply(AiActionBlock block) { AppliedOnThread = Thread.CurrentThread; Applied.Add(block.Mode + "|" + block.Body); EditSeq++; return new AiApplyResult("OK"); }
    public AiApplyResult InsertAnswer(string markdown, bool replaceSelection) { Applied.Add("answer|" + markdown); return new AiApplyResult("OK"); }
    public void Undo() { Undos++; EditSeq++; }
    public long EditSeq { get; set; }
}

/// <summary>A separate STA thread with a message loop (stands in for an app's UI thread).</summary>
public sealed class UiThread : IDisposable
{
    public Control Control { get; private set; } = null!;
    public Thread Thread { get; }
    readonly ManualResetEventSlim _ready = new();
    Form? _form;
    public UiThread()
    {
        Thread = new Thread(() =>
        {
            _form = new Form { ShowInTaskbar = false, WindowState = FormWindowState.Minimized, Opacity = 0 };
            Control = _form;
            _form.Load += (_, _) => _ready.Set();
            Application.Run(_form);
        }) { IsBackground = true };
        Thread.SetApartmentState(ApartmentState.STA);
        Thread.Start();
        _ready.Wait(5000);
    }
    public void Run(Action a) => Control.Invoke(a);
    public void Dispose() { try { _form?.Invoke(() => _form.Close()); } catch { } }
}
