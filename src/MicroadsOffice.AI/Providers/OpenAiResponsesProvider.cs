using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Providers;

/// <summary>
/// OpenAI Responses API (POST /v1/responses, streamed). Used with an OpenAI API key, and with a ChatGPT plan access token from
/// "Continue with ChatGPT" (then the request follows OpenAI's plan-usage contract: store=false, stream=true, instructions instead of
/// system messages, none of the unsupported fields such as temperature/max_output_tokens).
/// </summary>
public class OpenAiResponsesProvider : AiProvider
{
    readonly Func<CancellationToken, Task<string>> _token;
    readonly bool _plan;
    public string BaseUrl { get; }

    public OpenAiResponsesProvider(string baseUrl, Func<CancellationToken, Task<string>> token, bool chatGptPlan)
    { BaseUrl = baseUrl.TrimEnd('/'); _token = token; _plan = chatGptPlan; }

    public override AiProviderKind Kind => _plan ? AiProviderKind.ChatGpt : AiProviderKind.OpenAiKey;

    /// <summary>Renews the credential after a 401 (ChatGPT plan); the request is then retried once.</summary>
    public Func<CancellationToken, Task<string>>? ForceRefresh { get; init; }

    /// <summary>Waits before retrying a 503 (bounded backoff). Tests shorten it.</summary>
    public static int[] RetryDelaysMs = [1000, 3000];
    static DateTime PlanPausedUntil; static AiException? PlanPause;
    /// <summary>Forgets a usage-limit pause (after the user changed the limit / for tests).</summary>
    public static void ResetPlanPause() { PlanPausedUntil = default; PlanPause = null; }

    /// <summary>OpenAI's request id, kept in error messages so a failure can be diagnosed.</summary>
    static string? RequestId(HttpResponseMessage r) =>
        r.Headers.TryGetValues("x-request-id", out var v) ? v.FirstOrDefault() : r.Headers.TryGetValues("openai-request-id", out var v2) ? v2.FirstOrDefault() : null;

    public static string BuildBody(AiRequest req)
    {
        var input = new JsonArray();
        foreach (var m in req.Messages)
            input.Add(new JsonObject { ["role"] = m.Role == AiRole.User ? "user" : "assistant", ["content"] = m.Text });
        var body = new JsonObject
        {
            ["model"] = req.Model,
            ["instructions"] = req.Instructions,
            ["input"] = input,
            ["store"] = false,
            ["stream"] = true,
        };
        return body.ToJsonString();
    }

    public override async Task<string> CompleteAsync(AiRequest req, Action<string> onDelta, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Model)) throw new AiException(Loc.T("Chưa chọn mô hình AI. Mở Cài đặt › Trợ lý AI để chọn.", "No AI model selected. Open Settings › AI assistant to choose one."));
        var token = await _token(ct);
        using var http = NewHttp(TimeSpan.FromMinutes(10));
        var body = BuildBody(req);
        async Task<HttpResponseMessage> Send(string bearer)
        {
            using var msg = new HttpRequestMessage(HttpMethod.Post, BaseUrl + "/responses") { Content = new StringContent(body, Encoding.UTF8, "application/json") };
            msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
            msg.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
            try { return await http.SendAsync(msg, HttpCompletionOption.ResponseHeadersRead, ct); }
            catch (Exception e) when (e is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested) { throw new AiException(Network(e), inner: e); }
        }
        // OpenAI: after usage_limit_exceeded, pause new plan-billed requests instead of hammering the limit
        if (_plan && PlanPausedUntil > DateTime.UtcNow && PlanPause != null) throw PlanPause;
        var resp = await Send(token);
        if ((int)resp.StatusCode == 401 && ForceRefresh != null)
        {
            resp.Dispose();
            resp = await Send(await ForceRefresh(ct));
        }
        // 503 = temporarily unavailable (routing / usage check / workspace info): keep the credentials, bounded backoff
        for (int attempt = 0; (int)resp.StatusCode == 503 && attempt < RetryDelaysMs.Length; attempt++)
        {
            resp.Dispose();
            await Task.Delay(RetryDelaysMs[attempt], ct);
            resp = await Send(_plan ? await _token(ct) : token);
        }
        using (resp)
        {
            if (!resp.IsSuccessStatusCode)
            {
                var e = MapError((int)resp.StatusCode, await resp.Content.ReadAsStringAsync(ct), _plan, RequestId(resp));
                if (_plan && e.Code == "subscription_sharing_usage_limit_exceeded") { PlanPause = e; PlanPausedUntil = DateTime.UtcNow.AddMinutes(1); }
                throw e;
            }
            var sb = new StringBuilder();
            bool completed = false;
            await using var stream = await resp.Content.ReadAsStreamAsync(ct);
            await foreach (var (ev, data) in Sse.ReadAsync(stream, ct))
            {
                if (data == "[DONE]") break;
                JsonNode? j; try { j = JsonNode.Parse(data); } catch { continue; }
                var type = j?["type"]?.GetValue<string>() ?? ev;
                switch (type)
                {
                    case "response.output_text.delta":
                        var d = j?["delta"]?.GetValue<string>() ?? "";
                        if (d.Length > 0) { sb.Append(d); onDelta(d); }
                        break;
                    case "response.completed":
                        completed = true;
                        break;
                    case "response.failed":
                        {
                            var err = j?["response"]?["error"];
                            throw MapCode(err?["code"]?.GetValue<string>(), err?["message"]?.GetValue<string>(), 0, _plan);
                        }
                    case "response.incomplete":
                        {
                            var reason = j?["response"]?["incomplete_details"]?["reason"]?.GetValue<string>() ?? "";
                            throw new AiException(Loc.T("Câu trả lời bị cắt ngang (", "The answer was cut off (") + reason + ").", reason);
                        }
                    case "error":
                        throw MapCode(j?["code"]?.GetValue<string>() ?? j?["error"]?["code"]?.GetValue<string>(),
                                      j?["message"]?.GetValue<string>() ?? j?["error"]?["message"]?.GetValue<string>(), 0, _plan);
                }
                if (completed) break;
            }
            if (!completed) throw new AiException(Loc.T("Kết nối bị ngắt trước khi AI trả lời xong.", "The connection closed before the answer was complete."), "stream_interrupted");
            return sb.ToString();
        }
    }

    public override async Task<IReadOnlyList<AiModel>> ListModelsAsync(CancellationToken ct)
    {
        var token = await _token(ct);
        using var http = NewHttp(TimeSpan.FromSeconds(30));
        using var msg = new HttpRequestMessage(HttpMethod.Get, BaseUrl + "/models");
        msg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        HttpResponseMessage resp;
        try { resp = await http.SendAsync(msg, ct); }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested) { throw new AiException(Network(e), inner: e); }
        using (resp)
        {
            var text = await resp.Content.ReadAsStringAsync(ct);
            if (!resp.IsSuccessStatusCode) throw MapError((int)resp.StatusCode, text, _plan, RequestId(resp));
            return ParseModels(text, _plan);
        }
    }

    /// <summary>ChatGPT plan catalogue: {"models":[{slug,display_name,visibility}]}; API: {"data":[{id,created}]}.</summary>
    public static IReadOnlyList<AiModel> ParseModels(string json, bool plan)
    {
        var list = new List<AiModel>();
        var root = JsonNode.Parse(json);
        if (root?["models"] is JsonArray models)
        {
            foreach (var m in models)
            {
                if (m?["visibility"]?.GetValue<string>() != "list") continue; // OpenAI: keep only entries whose visibility is "list"
                var slug = m?["slug"]?.GetValue<string>(); if (string.IsNullOrEmpty(slug)) continue;
                list.Add(new AiModel(slug, m?["display_name"]?.GetValue<string>() ?? slug));
            }
            return list;
        }
        if (root?["data"] is JsonArray data)
        {
            var items = new List<(string Id, long Created)>();
            foreach (var m in data)
            {
                var id = m?["id"]?.GetValue<string>(); if (string.IsNullOrEmpty(id)) continue;
                long created = 0; try { created = m?["created"]?.GetValue<long>() ?? 0; } catch { }
                items.Add((id, created));
            }
            // Text models first (newest first); audio/image/embedding models are useless for the assistant.
            static bool Chat(string id) => !(id.Contains("embedding") || id.Contains("tts") || id.Contains("whisper") || id.Contains("audio")
                || id.Contains("realtime") || id.Contains("image") || id.Contains("dall-e") || id.Contains("transcribe") || id.Contains("moderation")
                || id.Contains("search") || id.StartsWith("babbage") || id.StartsWith("davinci") || id.Contains("sora"));
            foreach (var it in items.Where(i => Chat(i.Id)).OrderByDescending(i => i.Id.StartsWith("gpt") ? 1 : 0).ThenByDescending(i => i.Created).ThenBy(i => i.Id))
                list.Add(new AiModel(it.Id, it.Id));
        }
        return list;
    }

    public override string DefaultModel(IReadOnlyList<AiModel> models)
    {
        if (models.Count == 0) return "";
        if (_plan) return models[0].Id; // server ordering is meaningful for the ChatGPT catalogue
        return (models.FirstOrDefault(m => m.Id.StartsWith("gpt") && !m.Id.Contains("mini") && !m.Id.Contains("nano")) ?? models[0]).Id;
    }

    /// <summary>Maps an HTTP error body ({"error":{code,message}} or {"detail":"..."}) to a user-facing exception.</summary>
    public static AiException MapError(int status, string body, bool plan, string? requestId = null)
    {
        string? code = null, message = null, param = null;
        try
        {
            var j = JsonNode.Parse(body);
            code = j?["error"]?["code"]?.GetValue<string>() ?? j?["error"]?["type"]?.GetValue<string>();
            message = j?["error"]?["message"]?.GetValue<string>() ?? j?["detail"]?.ToString();
            param = j?["error"]?["param"]?.ToString();
        }
        catch { message = body.Length > 300 ? body[..300] : body; }
        return MapCode(code, message, status, plan, requestId, param);
    }

    public static AiException MapCode(string? code, string? message, int status, bool plan, string? requestId = null, string? param = null)
    {
        string detail = (string.IsNullOrWhiteSpace(message) ? "" : " (" + message + ")") + (string.IsNullOrEmpty(requestId) ? "" : Loc.T(" Mã yêu cầu: ", " Request ID: ") + requestId);
        switch (code)
        {
            case "subscription_sharing_user_not_eligible":
                return new AiException(Loc.T("Tài khoản/không gian làm việc ChatGPT này chưa được dùng gói ChatGPT trong ứng dụng khác (cần ChatGPT Plus hoặc Pro, hoặc quản trị viên đã tắt). Có thể dùng khóa API thay thế.",
                    "This ChatGPT account or workspace cannot use its ChatGPT plan in other apps (ChatGPT Plus or Pro is required, or an admin disabled it). You can use an API key instead.") + detail, code, status);
            case "subscription_sharing_usage_limit_exceeded":
                return new AiException(Loc.T("Đã hết hạn mức dùng gói ChatGPT cho Microads Office. Xem và chỉnh hạn mức trong ChatGPT › Settings › Usage.",
                    "The ChatGPT plan usage limit for Microads Office has been reached. Review and adjust it in ChatGPT › Settings › Usage.") + detail, code, status) { HelpUrl = AiDefaults.ChatGptUsage };
            case "subscription_sharing_usage_unavailable":
            case "subscription_sharing_user_unavailable":
                return new AiException(Loc.T("ChatGPT tạm thời không kiểm tra được hạn mức. Thử lại sau ít phút.", "ChatGPT could not check usage right now. Try again in a few minutes.") + detail, code, status);
            case "subscription_sharing_unsupported_capability":
                return new AiException(Loc.T("Mô hình hoặc tính năng này không dùng được với gói ChatGPT", "This model or feature is not available with the ChatGPT plan")
                    + (string.IsNullOrEmpty(param) ? "" : " [" + param + "]") + Loc.T(". Chọn mô hình khác trong Cài đặt › Trợ lý AI.", ". Choose another model in Settings › AI assistant.") + detail, code, status);
            case "subscription_sharing_route_not_supported":
                return new AiException(Loc.T("ChatGPT không cho phép thao tác này với gói ChatGPT. Hãy báo lỗi cho Microads (nút Báo cáo) kèm mã yêu cầu.", "ChatGPT does not allow this operation with the ChatGPT plan. Please report it to Microads (Report button) with the request ID.") + detail, code, status);
            // OpenAI: the subscriber context could not be validated — a forced refresh already succeeded, so the sign-in itself is
            // fine; ask to sign in again only after a confirmed revocation / terminal refresh error (those come from ChatGptAuth)
            case "subscription_sharing_invalid_user":
                return new AiException(Loc.T("ChatGPT chưa xác thực được phiên dùng gói lúc này. Thử lại sau ít phút; nếu vẫn lỗi, mở Cài đặt › Trợ lý AI, Đăng xuất rồi Continue with ChatGPT.", "ChatGPT could not validate the plan session right now. Try again in a few minutes; if it persists, open Settings › AI assistant, sign out and Continue with ChatGPT.") + detail, code, status);
            // configuration of the grant, not the user's session: signing in again or retrying does not help
            case "chatpass_v2_scope_not_authorized":
            case "chatpass_v2_invalid_authorization_context":
                return new AiException(Loc.T("Quyền đã cấp cho Microads Office không cho phép thao tác này (lỗi cấu hình phía ứng dụng). Hãy báo lỗi cho Microads (nút Báo cáo) kèm mã yêu cầu; trong lúc chờ có thể dùng khóa API.", "The permission granted to Microads Office does not allow this operation (an app configuration problem). Please report it to Microads (Report button) with the request ID; meanwhile you can use an API key.") + detail, code, status);
            case "invalid_api_key":
                return new AiException(Loc.T("Khóa API không đúng hoặc đã bị thu hồi.", "The API key is wrong or was revoked."), code, status) { NeedsSignIn = true };
            case "insufficient_quota":
                return new AiException(Loc.T("Tài khoản OpenAI đã hết hạn mức/tín dụng. Nạp thêm tại platform.openai.com.", "The OpenAI account has no quota/credit left. Add credit at platform.openai.com."), code, status);
            case "model_not_found":
                return new AiException(Loc.T("Không có mô hình này. Chọn mô hình khác trong Cài đặt › Trợ lý AI.", "This model is not available. Choose another one in Settings › AI assistant.") + detail, code, status);
            case "rate_limit_exceeded":
                return new AiException(Loc.T("Gửi quá nhiều yêu cầu. Chờ một lát rồi thử lại.", "Too many requests. Wait a moment and try again."), code, status);
        }
        return status switch
        {
            401 => new AiException((plan ? Loc.T("Phiên đăng nhập ChatGPT đã hết hạn hoặc bị ngắt kết nối. Hãy đăng nhập lại.", "The ChatGPT session expired or was disconnected. Please sign in again.")
                                         : Loc.T("Khóa API không được chấp nhận.", "The API key was not accepted.")) + detail, code, status) { NeedsSignIn = true },
            403 => new AiException(Loc.T("Máy chủ AI từ chối yêu cầu (quyền hoặc khu vực).", "The AI server refused the request (permission or region).") + detail, code, status),
            429 => new AiException(Loc.T("Đã chạm giới hạn sử dụng. Chờ một lát rồi thử lại.", "Usage limit reached. Wait a moment and try again.") + detail, code, status) { HelpUrl = plan ? AiDefaults.ChatGptUsage : null },
            >= 500 => new AiException(Loc.T("Máy chủ AI đang gặp sự cố. Thử lại sau.", "The AI server has a problem. Try again later.") + detail, code, status),
            _ => new AiException(Loc.T("Máy chủ AI báo lỗi", "The AI server reported an error") + (status > 0 ? $" {status}" : "") + detail, code, status),
        };
    }
}
