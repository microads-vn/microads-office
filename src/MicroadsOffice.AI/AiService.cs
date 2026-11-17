using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI;

/// <summary>Builds the provider the user chose and reports whether it is ready (credentials present).</summary>
public static class AiService
{
    /// <summary>Tests can inject a provider (e.g. pointing at a local fake server).</summary>
    public static Func<AiSettings, AiProvider?>? Override;

    public static AiProvider? Create(AiSettings s) => Override?.Invoke(s) ?? Create(s, s.Provider);

    public static AiProvider? Create(AiSettings s, AiProviderKind kind) => kind switch
    {
        AiProviderKind.ChatGpt => new OpenAiResponsesProvider(ChatGptAuth.ApiBase, ChatGptAuth.GetAccessTokenAsync, chatGptPlan: true)
        {
            // a 401 can mean the access token was revoked early: renew once before asking the user to sign in again
            ForceRefresh = async ct => (await ChatGptAuth.RefreshAsync(force: true, ct)).AccessToken,
        },
        AiProviderKind.OpenAiKey => SecretStore.Read(SecretNames.OpenAiKey) is { Length: > 0 } k
            ? new OpenAiResponsesProvider(ChatGptAuth.SafeOverride("MICROADS_AI_OPENAI_BASE") ?? AiDefaults.OpenAiApiBase, _ => Task.FromResult(k), chatGptPlan: false)
            : null,
        AiProviderKind.ClaudeKey => SecretStore.Read(SecretNames.ClaudeKey) is { Length: > 0 } ck ? new ClaudeProvider(ck) : null,
        AiProviderKind.Compatible => string.IsNullOrWhiteSpace(s.CompatibleBaseUrl) ? null : new CompatibleProvider(s.CompatibleBaseUrl, SecretStore.Read(SecretNames.CompatibleKey)),
        _ => null,
    };

    public static bool IsReady(AiSettings s)
    {
        if (Override != null) return true;
        return s.Provider switch
        {
            AiProviderKind.ChatGpt => ChatGptAuth.Current is { SignedIn: true, PlanUsage: true },
            AiProviderKind.OpenAiKey => SecretStore.Has(SecretNames.OpenAiKey),
            AiProviderKind.ClaudeKey => SecretStore.Has(SecretNames.ClaudeKey),
            AiProviderKind.Compatible => !string.IsNullOrWhiteSpace(s.CompatibleBaseUrl) && !string.IsNullOrWhiteSpace(s.CompatibleModel),
            _ => false,
        };
    }

    /// <summary>One line for the pane header, e.g. "ChatGPT · a@b.com · GPT-x".</summary>
    public static string StatusText(AiSettings s)
    {
        if (!IsReady(s))
            return s.AllowClaudeDesktop ? Loc.T("Chưa kết nối AI trong khung này (Claude Desktop: đã kết nối)", "AI not connected in this pane (Claude Desktop: connected)")
                                        : Loc.T("Chưa kết nối AI", "AI not connected");
        var model = s.ModelFor(s.Provider);
        string who = s.Provider == AiProviderKind.ChatGpt && ChatGptAuth.Current is { } c && c.Email.Length > 0 ? " · " + c.Email : "";
        return AiSettings.ProviderName(s.Provider) + who + (model.Length > 0 ? " · " + model : "");
    }
}
