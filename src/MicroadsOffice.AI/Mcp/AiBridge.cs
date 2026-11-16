using System.IO.Pipes;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Mcp;

/// <summary>
/// Local bridge between a running Word/Excel/PDF window and the MCP server process that Claude Desktop starts.
/// Named pipe restricted to the current Windows user; one JSON object per line. Every request is refused unless the user
/// enabled "Connect to Claude Desktop" (AiSettings.AllowClaudeDesktop). Edits run on the UI thread as one undo step.
/// </summary>
public static class AiBridge
{
    public static string PipeName(AppKind kind)
    {
        var who = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Environment.UserDomainName + "\\" + Environment.UserName + "|" + MicroadsOffice.Core.AppSettings.DataDirectory.ToLowerInvariant())))[..16];
        return $"MicroadsOffice.AI.{who}.{kind.ToString().ToLowerInvariant()}";
    }

    // ---------------- server (inside the app) ----------------

    static readonly Dictionary<AppKind, CancellationTokenSource> Servers = [];

    /// <summary>Starts serving <paramref name="hostProvider"/> (returns the active document's host, or null).</summary>
    public static void Start(AppKind kind, Func<IAiHost?> hostProvider, Control uiThread)
    {
        lock (Servers)
        {
            if (Servers.ContainsKey(kind)) return;
            var cts = new CancellationTokenSource();
            Servers[kind] = cts;
            _ = Task.Run(() => ServeLoop(kind, hostProvider, uiThread, cts.Token));
            // keep Claude Desktop pointing at this installation (its updates move the config; Cowork can rewrite it)
            if (Servers.Count == 1) _ = Task.Run(() => { try { ClaudeDesktop.EnsureConnected(); } catch { } });
        }
    }

    public static void Stop(AppKind kind)
    {
        lock (Servers) { if (Servers.Remove(kind, out var c)) c.Cancel(); }
    }

    static async Task ServeLoop(AppKind kind, Func<IAiHost?> hostProvider, Control ui, CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            NamedPipeServerStream? pipe = null;
            try
            {
                pipe = new NamedPipeServerStream(PipeName(kind), PipeDirection.InOut, 4, PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
                await pipe.WaitForConnectionAsync(ct);
                var p = pipe; pipe = null;
                _ = Task.Run(() => Serve(p, hostProvider, ui, ct));
            }
            catch (OperationCanceledException) { pipe?.Dispose(); break; }
            catch (IOException) { pipe?.Dispose(); await Task.Delay(500, ct).ContinueWith(_ => { }); }
            catch { pipe?.Dispose(); await Task.Delay(1000, ct).ContinueWith(_ => { }); }
        }
    }

    static async Task Serve(NamedPipeServerStream pipe, Func<IAiHost?> hostProvider, Control ui, CancellationToken ct)
    {
        using (pipe)
        {
            var reader = new StreamReader(pipe, new UTF8Encoding(false));
            var writer = new StreamWriter(pipe, new UTF8Encoding(false)) { AutoFlush = true, NewLine = "\n" };
            while (!ct.IsCancellationRequested && pipe.IsConnected)
            {
                var line = await reader.ReadLineAsync(ct);
                if (line == null) break;
                JsonObject resp;
                try
                {
                    var req = JsonNode.Parse(line) as JsonObject ?? throw new InvalidOperationException("bad request");
                    resp = Handle(req, hostProvider, ui);
                }
                catch (Exception e) { resp = new JsonObject { ["ok"] = false, ["error"] = e.Message }; }
                await writer.WriteLineAsync(resp.ToJsonString());
            }
        }
    }

    static JsonObject Handle(JsonObject req, Func<IAiHost?> hostProvider, Control ui)
    {
        var op = req["op"]?.GetValue<string>() ?? "";
        if (op == "ping") return Ok(new JsonObject { ["app"] = OfficeKind(hostProvider) });
        if (!AiSettings.Load().AllowClaudeDesktop)
            return Err(Loc.T("Microads Office chưa cho phép Claude Desktop truy cập. Trong Microads Office, mở Trợ lý AI › Cài đặt › Kết nối với Claude Desktop.",
                             "Microads Office has not allowed Claude Desktop access. In Microads Office open AI assistant › Settings › Connect to Claude Desktop."));
        var args = req["args"] as JsonObject ?? [];
        JsonObject result = [];
        string? error = null;
        void Work()
        {
            try
            {
                var host = hostProvider() ?? throw new InvalidOperationException(Loc.T("Chưa có văn bản nào đang mở.", "No document is open."));
                switch (op)
                {
                    case "info":
                        result = new JsonObject
                        {
                            ["app"] = host.Kind.ToString(), ["document"] = host.DocumentName, ["canEdit"] = host.CanEdit,
                            ["hasSelection"] = !string.IsNullOrEmpty(host.SelectionText),
                        };
                        break;
                    case "read":
                        int max = args["maxChars"]?.GetValue<int>() ?? AiDefaults.MaxContextChars;
                        result = new JsonObject { ["document"] = host.DocumentName, ["content"] = host.GetDocumentContext(Math.Clamp(max, 1000, 400_000)) };
                        break;
                    case "selection":
                        result = new JsonObject { ["selection"] = host.SelectionText ?? "" };
                        break;
                    case "guide":
                        result = new JsonObject { ["guide"] = host.ActionInstructions };
                        break;
                    case "apply":
                        if (!host.CanEdit) throw new InvalidOperationException(Loc.T("Ứng dụng này chỉ đọc.", "This app is read-only."));
                        var r = host.Apply(new AiActionBlock(args["kind"]?.GetValue<string>() ?? (host.Kind == AppKind.Excel ? "sheet" : "document"),
                                                             args["mode"]?.GetValue<string>() ?? "insert", args["body"]?.GetValue<string>() ?? ""));
                        result = new JsonObject { ["summary"] = r.Summary };
                        break;
                    case "undo":
                        host.Undo();
                        result = new JsonObject { ["summary"] = Loc.T("Đã hoàn tác thay đổi gần nhất.", "Undid the last change.") };
                        break;
                    default: throw new InvalidOperationException("unknown op " + op);
                }
            }
            catch (Exception e) { error = e.Message; }
        }
        if (ui.InvokeRequired) ui.Invoke(Work); else Work();
        return error != null ? Err(error) : Ok(result);
    }

    static string OfficeKind(Func<IAiHost?> hp) { try { return hp()?.Kind.ToString() ?? ""; } catch { return ""; } }
    static JsonObject Ok(JsonNode r) => new() { ["ok"] = true, ["result"] = r };
    static JsonObject Err(string m) => new() { ["ok"] = false, ["error"] = m };

    // ---------------- client (inside the MCP server process) ----------------

    public static async Task<JsonNode> CallAsync(AppKind kind, string op, JsonObject? args, int connectTimeoutMs, CancellationToken ct)
    {
        using var pipe = new NamedPipeClientStream(".", PipeName(kind), PipeDirection.InOut, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        try { await pipe.ConnectAsync(connectTimeoutMs, ct); }
        catch (TimeoutException) { throw new BridgeUnavailableException(kind); }
        var writer = new StreamWriter(pipe, new UTF8Encoding(false)) { AutoFlush = true, NewLine = "\n" };
        var reader = new StreamReader(pipe, new UTF8Encoding(false));
        await writer.WriteLineAsync(new JsonObject { ["op"] = op, ["args"] = args ?? [] }.ToJsonString());
        var line = await reader.ReadLineAsync(ct) ?? throw new IOException("no response");
        var resp = JsonNode.Parse(line)!;
        if (resp["ok"]?.GetValue<bool>() != true) throw new InvalidOperationException(resp["error"]?.GetValue<string>() ?? "error");
        return resp["result"] ?? new JsonObject();
    }
}

public sealed class BridgeUnavailableException(AppKind kind) : Exception(Loc.F("Microads {0} chưa mở.", "Microads {0} is not open.", kind))
{
    public AppKind Kind { get; } = kind;
}
