using System.Net;
using System.Net.Http.Headers;
using System.Text;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Providers;

public enum AiRole { User, Assistant }

public sealed record AiMessage(AiRole Role, string Text);

/// <summary>One assistant request: standing instructions + the conversation (the last message is the new question).</summary>
public sealed record AiRequest(string Instructions, IReadOnlyList<AiMessage> Messages, string Model);

public sealed record AiModel(string Id, string DisplayName);

/// <summary>A failure the user can understand. <see cref="Code"/> is the provider's machine-readable code when present.</summary>
public sealed class AiException : Exception
{
    public string? Code { get; }
    public int Status { get; }
    /// <summary>The user must sign in / enter a key again (do not retry).</summary>
    public bool NeedsSignIn { get; init; }
    /// <summary>Link the error message should offer (e.g. ChatGPT usage settings).</summary>
    public string? HelpUrl { get; init; }
    public AiException(string message, string? code = null, int status = 0, Exception? inner = null) : base(message, inner) { Code = code; Status = status; }
}

/// <summary>A service that can answer. Implementations stream text deltas through <paramref name="onDelta"/>.</summary>
public abstract class AiProvider
{
    public abstract AiProviderKind Kind { get; }
    public abstract Task<string> CompleteAsync(AiRequest request, Action<string> onDelta, CancellationToken ct);
    public abstract Task<IReadOnlyList<AiModel>> ListModelsAsync(CancellationToken ct);

    /// <summary>Picks a sensible default when the user has not chosen a model yet.</summary>
    public virtual string DefaultModel(IReadOnlyList<AiModel> models) => models.Count > 0 ? models[0].Id : "";

    internal static HttpClient NewHttp(TimeSpan timeout)
    {
        var h = new HttpClient(new SocketsHttpHandler { AutomaticDecompression = DecompressionMethods.All, PooledConnectionLifetime = TimeSpan.FromMinutes(5) })
        { Timeout = timeout };
        h.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("MicroadsOffice", Product.VersionText.Split(' ')[0]));
        return h;
    }

    internal static string Network(Exception e) => e switch
    {
        TaskCanceledException => Loc.T("Hết thời gian chờ phản hồi từ máy chủ AI.", "The AI server did not respond in time."),
        HttpRequestException h when h.HttpRequestError == HttpRequestError.NameResolutionError =>
            Loc.T("Không có kết nối Internet (không tìm thấy máy chủ).", "No Internet connection (server not found)."),
        HttpRequestException => Loc.T("Không kết nối được máy chủ AI: ", "Could not reach the AI server: ") + e.Message,
        _ => e.Message,
    };
}

/// <summary>Minimal Server-Sent Events reader (event:/data: lines, blank line ends an event).</summary>
public static class Sse
{
    public static async IAsyncEnumerable<(string Event, string Data)> ReadAsync(Stream s, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        using var r = new StreamReader(s, Encoding.UTF8);
        string ev = ""; var data = new StringBuilder();
        while (true)
        {
            var line = await r.ReadLineAsync(ct);
            if (line == null) break;
            if (line.Length == 0)
            {
                if (data.Length > 0) yield return (ev, data.ToString());
                ev = ""; data.Clear();
                continue;
            }
            if (line.StartsWith(':')) continue;
            int c = line.IndexOf(':');
            string field = c < 0 ? line : line[..c];
            string value = c < 0 ? "" : line[(c + 1)..];
            if (value.StartsWith(' ')) value = value[1..];
            if (field == "event") ev = value;
            else if (field == "data") { if (data.Length > 0) data.Append('\n'); data.Append(value); }
        }
        if (data.Length > 0) yield return (ev, data.ToString());
    }
}
