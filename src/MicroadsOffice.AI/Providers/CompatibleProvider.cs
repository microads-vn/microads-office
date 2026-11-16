using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Providers;

/// <summary>
/// Any server that speaks the OpenAI Chat Completions protocol: Ollama / LM Studio on this computer, an internal company
/// server, or a hosted service. POST {base}/chat/completions with stream=true; GET {base}/models.
/// </summary>
public sealed class CompatibleProvider : AiProvider
{
    readonly string _base; readonly string? _key;
    public CompatibleProvider(string baseUrl, string? apiKey) { _base = baseUrl.TrimEnd('/'); _key = string.IsNullOrWhiteSpace(apiKey) ? null : apiKey; }
    public override AiProviderKind Kind => AiProviderKind.Compatible;

    public static string BuildBody(AiRequest req)
    {
        var msgs = new JsonArray();
        if (!string.IsNullOrEmpty(req.Instructions)) msgs.Add(new JsonObject { ["role"] = "system", ["content"] = req.Instructions });
        foreach (var m in req.Messages) msgs.Add(new JsonObject { ["role"] = m.Role == AiRole.User ? "user" : "assistant", ["content"] = m.Text });
        return new JsonObject { ["model"] = req.Model, ["messages"] = msgs, ["stream"] = true }.ToJsonString();
    }

    HttpRequestMessage Req(HttpMethod m, string path, string? body = null)
    {
        var r = new HttpRequestMessage(m, _base + path);
        if (body != null) r.Content = new StringContent(body, Encoding.UTF8, "application/json");
        if (_key != null) r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _key);
        return r;
    }

    public override async Task<string> CompleteAsync(AiRequest req, Action<string> onDelta, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Model)) throw new AiException(Loc.T("Chưa chọn mô hình AI. Mở Cài đặt › Trợ lý AI để chọn.", "No AI model selected. Open Settings › AI assistant to choose one."));
        using var http = NewHttp(TimeSpan.FromMinutes(15));
        using var msg = Req(HttpMethod.Post, "/chat/completions", BuildBody(req));
        HttpResponseMessage resp;
        try { resp = await http.SendAsync(msg, HttpCompletionOption.ResponseHeadersRead, ct); }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested) { throw new AiException(Network(e), inner: e); }
        using (resp)
        {
            if (!resp.IsSuccessStatusCode) throw OpenAiResponsesProvider.MapError((int)resp.StatusCode, await resp.Content.ReadAsStringAsync(ct), false);
            var sb = new StringBuilder(); bool done = false;
            await using var stream = await resp.Content.ReadAsStreamAsync(ct);
            await foreach (var (_, data) in Sse.ReadAsync(stream, ct))
            {
                if (data.Trim() == "[DONE]") { done = true; break; }
                JsonNode? j; try { j = JsonNode.Parse(data); } catch { continue; }
                if (j?["error"] is JsonNode err) throw OpenAiResponsesProvider.MapCode(err["code"]?.ToString(), err["message"]?.ToString() ?? err.ToString(), 0, false);
                var choice = j?["choices"]?[0];
                var d = choice?["delta"]?["content"]?.GetValue<string>();
                if (!string.IsNullOrEmpty(d)) { sb.Append(d); onDelta(d); }
                if (choice?["finish_reason"] is JsonNode fr && fr.GetValueKind() == System.Text.Json.JsonValueKind.String)
                {
                    var reason = fr.GetValue<string>();
                    done = true;
                    if (reason == "length") throw new AiException(Loc.T("Câu trả lời bị cắt ngang (quá dài).", "The answer was cut off (too long)."), "length");
                }
            }
            if (!done)
                throw sb.Length == 0
                    ? new AiException(Loc.T("Máy chủ AI không trả lời.", "The AI server returned nothing."), "empty")
                    : new AiException(Loc.T("Kết nối bị ngắt trước khi AI trả lời xong.", "The connection closed before the answer was complete."), "stream_interrupted");
            return sb.ToString();
        }
    }

    public override async Task<IReadOnlyList<AiModel>> ListModelsAsync(CancellationToken ct)
    {
        using var http = NewHttp(TimeSpan.FromSeconds(20));
        using var msg = Req(HttpMethod.Get, "/models");
        HttpResponseMessage resp;
        try { resp = await http.SendAsync(msg, ct); }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested) { throw new AiException(Network(e), inner: e); }
        using (resp)
        {
            var text = await resp.Content.ReadAsStringAsync(ct);
            if (!resp.IsSuccessStatusCode) throw OpenAiResponsesProvider.MapError((int)resp.StatusCode, text, false);
            var list = new List<AiModel>();
            if (JsonNode.Parse(text)?["data"] is JsonArray data)
                foreach (var m in data) { var id = m?["id"]?.GetValue<string>(); if (!string.IsNullOrEmpty(id)) list.Add(new AiModel(id, id)); }
            return list;
        }
    }
}
