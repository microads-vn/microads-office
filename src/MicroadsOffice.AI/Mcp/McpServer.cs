using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.AI.Mcp;

/// <summary>
/// Model Context Protocol server over stdio ("MicroadsOffice.exe --mcp"), started by Claude Desktop. It exposes tools that read
/// and edit the documents open in Microads Word / Excel / PDF through <see cref="AiBridge"/>. JSON-RPC 2.0, one message per line.
/// </summary>
public sealed class McpServer
{
    static readonly string[] Versions = ["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05"];
    readonly TextReader _in; readonly TextWriter _out;
    readonly Func<AppKind, bool> _launch;

    public McpServer(TextReader input, TextWriter output, Func<AppKind, bool>? launch = null)
    { _in = input; _out = output; _launch = launch ?? LaunchApp; }

    /// <summary>Entry point for "--mcp": reads stdin until it closes.</summary>
    public static int RunStdio()
    {
        var utf8 = new UTF8Encoding(false);
        using var stdin = new StreamReader(Console.OpenStandardInput(), utf8);
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), utf8) { AutoFlush = true, NewLine = "\n" };
        new McpServer(stdin, stdout).RunAsync(CancellationToken.None).GetAwaiter().GetResult();
        return 0;
    }

    public async Task RunAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            var line = await _in.ReadLineAsync(ct);
            if (line == null) break;
            if (string.IsNullOrWhiteSpace(line)) continue;
            JsonNode? resp;
            try { resp = await HandleAsync(JsonNode.Parse(line) as JsonObject, ct); }
            catch (JsonException) { resp = Error(null, -32700, "Parse error"); }
            if (resp != null) { await _out.WriteLineAsync(resp.ToJsonString()); await _out.FlushAsync(ct); }
        }
    }

    public async Task<JsonNode?> HandleAsync(JsonObject? msg, CancellationToken ct)
    {
        if (msg == null) return Error(null, -32600, "Invalid Request");
        var id = msg["id"]?.DeepClone();
        string? method;
        try { method = msg["method"]?.GetValue<string>(); }
        catch (Exception) { return id == null ? null : Error(id, -32600, "Invalid Request: method must be a string"); }
        if (method == null) return null; // a response or garbage: ignore
        bool notification = id == null;
        try
        {
            JsonNode result = method switch
            {
                "initialize" => Initialize(msg["params"] as JsonObject),
                "ping" => new JsonObject(),
                "tools/list" => new JsonObject { ["tools"] = Tools() },
                "tools/call" => await CallToolAsync(msg["params"] as JsonObject, ct),
                "resources/list" => new JsonObject { ["resources"] = new JsonArray() },
                "prompts/list" => new JsonObject { ["prompts"] = new JsonArray() },
                _ when method.StartsWith("notifications/") => new JsonObject(),
                _ => throw new McpError(-32601, "Method not found: " + method),
            };
            return notification ? null : new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };
        }
        catch (McpError e) { return notification ? null : Error(id, e.Code, e.Message); }
        // a malformed message must never stop the server (Claude Desktop would show it as disconnected)
        catch (Exception e) when (e is InvalidOperationException or FormatException or JsonException or KeyNotFoundException)
        { return notification ? null : Error(id, -32602, "Invalid params: " + e.Message); }
        catch (OperationCanceledException) { throw; }
        catch (Exception e) { return notification ? null : Error(id, -32603, "Internal error: " + e.Message); }
    }

    static JsonObject Error(JsonNode? id, int code, string message) =>
        new() { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JsonObject { ["code"] = code, ["message"] = message } };

    sealed class McpError(int code, string message) : Exception(message) { public int Code { get; } = code; }

    static JsonObject Initialize(JsonObject? p)
    {
        ClaudeDesktop.RecordHandshake(p?["clientInfo"]?["name"]?.GetValue<string>() ?? "", p?["clientInfo"]?["version"]?.GetValue<string>() ?? "");
        var requested = p?["protocolVersion"]?.GetValue<string>();
        var version = requested != null && Versions.Contains(requested) ? requested : Versions[0];
        return new JsonObject
        {
            ["protocolVersion"] = version,
            ["capabilities"] = new JsonObject { ["tools"] = new JsonObject { ["listChanged"] = false } },
            ["serverInfo"] = new JsonObject { ["name"] = ClaudeDesktop.ServerName, ["title"] = "Microads Office", ["version"] = Product.VersionText },
            ["instructions"] = "Microads Office is a free office suite (Word, Excel, PDF). Use these tools to read and edit the document the user has open in Microads Word / Excel / PDF. " +
                               "Call list_open_apps first. Before editing call get_editing_guide for the app to learn the content format. Edits are applied as one undo step each; tell the user they can press Ctrl+Z in Microads Office.",
        };
    }

    static JsonObject AppParam(string desc, params string[] allowed) => new()
    {
        ["type"] = "string", ["enum"] = new JsonArray(allowed.Select(a => (JsonNode)a).ToArray()), ["description"] = desc,
    };

    static JsonObject Tool(string name, string title, string description, JsonObject properties, string[] required, bool readOnly, bool destructive = false) => new()
    {
        ["name"] = name,
        ["title"] = title,
        ["description"] = description,
        ["inputSchema"] = new JsonObject
        {
            ["type"] = "object", ["properties"] = properties,
            ["required"] = new JsonArray(required.Select(r => (JsonNode)r).ToArray()), ["additionalProperties"] = false,
        },
        ["annotations"] = new JsonObject { ["readOnlyHint"] = readOnly, ["destructiveHint"] = destructive, ["openWorldHint"] = false },
    };

    public static JsonArray Tools() =>
    [
        Tool("list_open_apps", "List open Microads Office apps", "Lists which Microads Office apps (word, excel, pdf) are open and the active document in each.", [], [], true),
        Tool("open_app", "Open a Microads Office app", "Starts Microads Word, Excel or PDF if it is not open yet (the user then opens or creates a document).",
            new JsonObject { ["app"] = AppParam("Which app", "word", "excel", "pdf") }, ["app"], false),
        Tool("read_document", "Read the open document", "Returns the active document's content: Word as Markdown, Excel as the active sheet's cells (address, value, formula), PDF as page text.",
            new JsonObject { ["app"] = AppParam("Which app", "word", "excel", "pdf"), ["max_chars"] = new JsonObject { ["type"] = "integer", ["description"] = "Maximum characters (default 120000)" } }, ["app"], true),
        Tool("get_selection", "Read the selection", "Returns the text (Word, PDF) or cells (Excel) the user has selected.",
            new JsonObject { ["app"] = AppParam("Which app", "word", "excel", "pdf") }, ["app"], true),
        Tool("get_editing_guide", "How to edit", "Returns the exact format that write_document / edit_sheet accept for the app. Call it before the first edit.",
            new JsonObject { ["app"] = AppParam("Which app", "word", "excel") }, ["app"], true),
        Tool("write_document", "Write into the Word document", "Writes Markdown content into the open Microads Word document (headings, bold, lists, tables, ':::center' blocks). mode: insert (at the cursor), replace_selection, or replace_document. One undo step.",
            new JsonObject
            {
                ["content"] = new JsonObject { ["type"] = "string", ["description"] = "Markdown content" },
                ["mode"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("insert", "replace_selection", "replace_document", "append") },
            }, ["content", "mode"], false, destructive: true),
        Tool("edit_sheet", "Edit the Excel workbook", "Applies a JSON list of actions to the open Microads Excel workbook (values, formulas, formats, sheets, charts). See get_editing_guide for the schema. One undo step.",
            new JsonObject { ["actions_json"] = new JsonObject { ["type"] = "string", ["description"] = "JSON text: {\"actions\":[...]}" } }, ["actions_json"], false, destructive: true),
        Tool("undo", "Undo", "Undoes the last change in the app (same as Ctrl+Z).",
            new JsonObject { ["app"] = AppParam("Which app", "word", "excel") }, ["app"], false, destructive: true),
    ];

    static AppKind ParseApp(JsonObject? a)
    {
        return (a?["app"]?.GetValue<string>() ?? "").ToLowerInvariant() switch
        {
            "word" => AppKind.Word, "excel" => AppKind.Excel, "pdf" => AppKind.Pdf,
            _ => throw new McpError(-32602, "app must be word, excel or pdf"),
        };
    }

    async Task<JsonNode> CallToolAsync(JsonObject? p, CancellationToken ct)
    {
        var name = p?["name"]?.GetValue<string>() ?? throw new McpError(-32602, "missing tool name");
        var a = p?["arguments"] as JsonObject ?? [];
        try
        {
            switch (name)
            {
                case "list_open_apps":
                    {
                        var sb = new StringBuilder();
                        foreach (var k in new[] { AppKind.Word, AppKind.Excel, AppKind.Pdf })
                        {
                            try
                            {
                                var info = await AiBridge.CallAsync(k, "info", null, 400, ct);
                                sb.AppendLine($"{k.ToString().ToLowerInvariant()}: open — document \"{info["document"]}\"{(info["hasSelection"]?.GetValue<bool>() == true ? " (has a selection)" : "")}");
                            }
                            catch (BridgeUnavailableException) { sb.AppendLine($"{k.ToString().ToLowerInvariant()}: not open"); }
                            catch (Exception e) { sb.AppendLine($"{k.ToString().ToLowerInvariant()}: {e.Message}"); }
                        }
                        return Text(sb.ToString().TrimEnd());
                    }
                case "open_app":
                    {
                        var k = ParseApp(a);
                        try { await AiBridge.CallAsync(k, "ping", null, 400, ct); return Text($"Microads {k} is already open."); } catch (BridgeUnavailableException) { }
                        if (!_launch(k)) return Text($"Could not start Microads {k}.", true);
                        for (int i = 0; i < 30; i++)
                        {
                            await Task.Delay(500, ct);
                            try { await AiBridge.CallAsync(k, "ping", null, 300, ct); return Text($"Microads {k} started. Ask the user to open or create a document if needed."); }
                            catch (BridgeUnavailableException) { }
                        }
                        return Text($"Microads {k} was started but is not ready yet; try again in a few seconds.", true);
                    }
                case "read_document":
                    {
                        var args = new JsonObject();
                        if (a["max_chars"] is JsonNode mc) args["maxChars"] = mc.GetValue<int>();
                        var r = await AiBridge.CallAsync(ParseApp(a), "read", args, 1500, ct);
                        return Text($"Document: {r["document"]}\n\n{r["content"]}");
                    }
                case "get_selection":
                    {
                        var r = await AiBridge.CallAsync(ParseApp(a), "selection", null, 1500, ct);
                        var s = r["selection"]?.GetValue<string>() ?? "";
                        return Text(s.Length == 0 ? "(nothing selected)" : s);
                    }
                case "get_editing_guide":
                    {
                        var r = await AiBridge.CallAsync(ParseApp(a), "guide", null, 1500, ct);
                        return Text(r["guide"]?.GetValue<string>() ?? "");
                    }
                case "write_document":
                    {
                        var mode = a["mode"]?.GetValue<string>() ?? "insert";
                        var r = await AiBridge.CallAsync(AppKind.Word, "apply", new JsonObject { ["kind"] = "document", ["mode"] = mode, ["body"] = a["content"]?.GetValue<string>() ?? "" }, 1500, ct);
                        return Text(r["summary"]?.GetValue<string>() ?? "OK");
                    }
                case "edit_sheet":
                    {
                        var r = await AiBridge.CallAsync(AppKind.Excel, "apply", new JsonObject { ["kind"] = "sheet", ["mode"] = "apply", ["body"] = a["actions_json"]?.GetValue<string>() ?? "" }, 1500, ct);
                        return Text(r["summary"]?.GetValue<string>() ?? "OK");
                    }
                case "undo":
                    {
                        var r = await AiBridge.CallAsync(ParseApp(a), "undo", null, 1500, ct);
                        return Text(r["summary"]?.GetValue<string>() ?? "OK");
                    }
                default: throw new McpError(-32602, "Unknown tool: " + name);
            }
        }
        catch (McpError) { throw; }
        catch (BridgeUnavailableException e) { return Text(e.Message + " Use open_app to start it, or ask the user to open it.", true); }
        catch (OperationCanceledException) { throw; }
        catch (Exception e) { return Text(e.Message, true); }
    }

    static JsonObject Text(string text, bool isError = false) => new()
    {
        ["content"] = new JsonArray(new JsonObject { ["type"] = "text", ["text"] = text }),
        ["isError"] = isError,
    };

    static bool LaunchApp(AppKind k)
    {
        try
        {
            var exe = Environment.ProcessPath;
            if (exe == null) return false;
            Process.Start(new ProcessStartInfo(exe, "--app=" + k.ToString().ToLowerInvariant()) { UseShellExecute = false });
            return true;
        }
        catch { return false; }
    }
}
