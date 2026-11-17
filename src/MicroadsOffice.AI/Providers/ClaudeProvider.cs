using System.Text;
using Anthropic;
using Anthropic.Exceptions;
using Anthropic.Models.Beta;
using Anthropic.Models.Beta.Messages;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Providers;

/// <summary>
/// Claude through the official Anthropic C# SDK with the user's own Claude Console API key (Anthropic does not allow third-party
/// apps to sign in with Claude.ai subscriptions — subscribers use the Claude Desktop connection instead, see Mcp/).
/// Streams with adaptive thinking (always on for Opus 5.5) and server-side refusal fallbacks enabled.
/// </summary>
public sealed class ClaudeProvider : AiProvider
{
    readonly string _key; readonly string? _baseUrl;
    public ClaudeProvider(string apiKey, string? baseUrl = null) { _key = apiKey; _baseUrl = baseUrl ?? ChatGptAuth.SafeOverride("MICROADS_AI_ANTHROPIC_BASE"); }
    public override AiProviderKind Kind => AiProviderKind.ClaudeKey;

    AnthropicClient Client()
    {
        var c = new AnthropicClient { ApiKey = _key, MaxRetries = 2, Timeout = TimeSpan.FromMinutes(10) };
        if (!string.IsNullOrEmpty(_baseUrl)) c = new AnthropicClient { ApiKey = _key, MaxRetries = 0, Timeout = TimeSpan.FromMinutes(10), BaseUrl = _baseUrl };
        return c;
    }

    /// <summary>
    /// output_config.effort returns a 400 on Sonnet 4.5 / Haiku 4.5 and older models; every newer model accepts "medium".
    /// The user can pick any model the key lists, so older ones must still work.
    /// </summary>
    public static bool SupportsEffort(string model)
    {
        var m = model.ToLowerInvariant();
        string[] old = ["claude-3", "claude-sonnet-4-5", "claude-haiku-4-5", "claude-opus-4-1", "claude-opus-4-0", "claude-opus-4-2025", "claude-sonnet-4-0", "claude-sonnet-4-2025"];
        return !old.Any(m.StartsWith);
    }

    /// <summary>Server-side refusal fallbacks ("default" mode) are documented for Fable / Mythos 5.x, Opus 5 / 5.5 and Sonnet 5.5.</summary>
    public static bool SupportsFallbacks(string model)
    {
        var m = model.ToLowerInvariant();
        return m.StartsWith("claude-fable-5") || m.StartsWith("claude-mythos-5") || m.StartsWith("claude-opus-5") || m.StartsWith("claude-sonnet-5-5");
    }

    public static MessageCreateParams BuildParams(AiRequest req)
    {
        var messages = new List<BetaMessageParam>();
        foreach (var m in req.Messages)
            messages.Add(m.Role == AiRole.User ? new BetaUserMessageParam { Content = m.Text } : new BetaAssistantMessageParam { Content = m.Text });
        var p = new MessageCreateParams
        {
            Model = req.Model,
            MaxTokens = 64000,
            System = req.Instructions,
            Messages = messages,
        };
        if (SupportsEffort(req.Model)) p = p with { OutputConfig = new BetaOutputConfig { Effort = Effort.Medium } };
        if (SupportsFallbacks(req.Model)) p = p with { Betas = [AnthropicBeta.ServerSideFallback2026_07_01], Fallbacks = new Default() };
        return p;
    }

    public override async Task<string> CompleteAsync(AiRequest req, Action<string> onDelta, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Model)) throw new AiException(Loc.T("Chưa chọn mô hình Claude.", "No Claude model selected."));
        var sb = new StringBuilder();
        string? stop = null;
        try
        {
            await foreach (var ev in Client().Beta.Messages.CreateStreaming(BuildParams(req), ct))
            {
                if (ev.TryPickContentBlockDelta(out var delta) && delta.Delta.TryPickText(out var text))
                {
                    sb.Append(text.Text); onDelta(text.Text);
                }
                else if (ev.TryPickDelta(out var md) && md.Delta.StopReason is { } sr)
                {
                    stop = sr.ToString();
                }
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (AnthropicUnauthorizedException e) { throw new AiException(Loc.T("Khóa API Claude không đúng hoặc đã bị thu hồi.", "The Claude API key is wrong or was revoked."), "authentication_error", 401, e) { NeedsSignIn = true }; }
        catch (AnthropicRateLimitException e) { throw new AiException(Loc.T("Đã chạm giới hạn sử dụng Claude. Chờ một lát rồi thử lại.", "Claude rate limit reached. Wait a moment and try again."), "rate_limit_error", 429, e); }
        catch (AnthropicNotFoundException e) { throw new AiException(Loc.T("Không có mô hình Claude này. Chọn mô hình khác trong Cài đặt › Trợ lý AI.", "This Claude model is not available. Choose another one in Settings › AI assistant."), "not_found_error", 404, e); }
        catch (AnthropicBadRequestException e) { throw new AiException(Loc.T("Claude từ chối yêu cầu: ", "Claude rejected the request: ") + e.Message, "invalid_request_error", 400, e); }
        catch (Anthropic5xxException e) { throw new AiException(Loc.T("Máy chủ Claude đang quá tải hoặc gặp sự cố. Thử lại sau.", "Claude is overloaded or has a problem. Try again later."), "api_error", 500, e); }
        catch (AnthropicApiException e) { throw new AiException(Loc.T("Claude báo lỗi: ", "Claude reported an error: ") + e.Message, "api_error", 0, e); }
        catch (AnthropicIOException e) { throw new AiException(Loc.T("Không kết nối được Claude: ", "Could not reach Claude: ") + e.Message, "io", 0, e); }
        catch (HttpRequestException e) { throw new AiException(Network(e), inner: e); }
        // only a normal end is a usable answer: anything else must never be written into a document
        var norm = (stop ?? "").Replace("_", "").ToLowerInvariant();
        if (norm.Contains("endturn") || norm.Contains("stopsequence")) return sb.ToString();
        if (norm.Contains("refusal")) throw new AiException(Loc.T("Claude từ chối trả lời yêu cầu này.", "Claude declined this request."), "refusal");
        if (norm.Contains("maxtokens")) throw new AiException(Loc.T("Câu trả lời bị cắt ngang (quá dài).", "The answer was cut off (too long)."), "max_tokens");
        if (stop == null) throw new AiException(Loc.T("Kết nối bị ngắt trước khi Claude trả lời xong.", "The connection closed before Claude finished."), "stream_interrupted");
        throw new AiException(Loc.T("Claude dừng trả lời bất thường (", "Claude stopped unexpectedly (") + stop + ").", "stop_" + norm);
    }

    public override async Task<IReadOnlyList<AiModel>> ListModelsAsync(CancellationToken ct)
    {
        try
        {
            // every page, not only the first (older models would otherwise be missing from the list)
            var page = await Client().Models.List(cancellationToken: ct);
            var list = new List<AiModel>();
            await foreach (var m in page.Paginate(ct))
                list.Add(new AiModel(m.ID, string.IsNullOrEmpty(m.DisplayName) ? m.ID : m.DisplayName));
            return list;
        }
        catch (AnthropicUnauthorizedException e) { throw new AiException(Loc.T("Khóa API Claude không đúng hoặc đã bị thu hồi.", "The Claude API key is wrong or was revoked."), "authentication_error", 401, e) { NeedsSignIn = true }; }
        catch (AnthropicApiException e) { throw new AiException(Loc.T("Claude báo lỗi: ", "Claude reported an error: ") + e.Message, "api_error", 0, e); }
        catch (AnthropicIOException e) { throw new AiException(Loc.T("Không kết nối được Claude: ", "Could not reach Claude: ") + e.Message, "io", 0, e); }
        catch (HttpRequestException e) { throw new AiException(Network(e), inner: e); }
    }

    public override string DefaultModel(IReadOnlyList<AiModel> models) =>
        models.Any(m => m.Id == AiDefaults.ClaudeModel) ? AiDefaults.ClaudeModel : models.Count > 0 ? models[0].Id : AiDefaults.ClaudeModel;
}
