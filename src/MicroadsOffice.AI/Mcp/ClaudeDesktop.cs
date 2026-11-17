using System.Text.Json;
using System.Text.Json.Nodes;

namespace MicroadsOffice.AI.Mcp;

/// <summary>
/// Registers Microads Office as a local MCP server in Claude Desktop's configuration (claude_desktop_config.json → mcpServers).
/// Other entries in the file are preserved; a backup is written before each change.
/// </summary>
public static class ClaudeDesktop
{
    public const string ServerName = "microads-office";

    /// <summary>Executable Claude Desktop should start (the installed MicroadsOffice.exe; for the Store edition its execution alias). Tests override it.</summary>
    public static string ExePath { get; set; } = MicroadsOffice.Core.Office.StorePackage.IsPackaged
        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WindowsApps", "MicroadsOffice.exe")
        : Environment.ProcessPath ?? "MicroadsOffice.exe";

    /// <summary>
    /// The Store (MSIX) edition does not change Claude Desktop's file: Windows redirects new files a packaged app creates under
    /// AppData (our temp file + backup) into the app's private copy, so an atomic rewrite would never reach Claude. There the
    /// user pastes <see cref="ConfigSnippet"/> in Claude Desktop › Settings › Developer › Edit Config (which opens the file Claude reads).
    /// MICROADS_CLAUDE_MANUAL=1 forces this path (tests).
    /// </summary>
    public static bool CanEditConfig => !MicroadsOffice.Core.Office.StorePackage.IsPackaged && Environment.GetEnvironmentVariable("MICROADS_CLAUDE_MANUAL") != "1";

    /// <summary>The configuration to paste into claude_desktop_config.json.</summary>
    public static string ConfigSnippet() => new JsonObject
    {
        ["mcpServers"] = new JsonObject
        {
            [ServerName] = new JsonObject { ["command"] = ExePath, ["args"] = new JsonArray("--mcp") },
        },
    }.ToJsonString(new JsonSerializerOptions { WriteIndented = true });

    /// <summary>Tests point this at a scratch folder (MICROADS_CLAUDE_CONFIG_ROOT) so the real Claude Desktop configuration is never touched.</summary>
    static string? Root => Environment.GetEnvironmentVariable("MICROADS_CLAUDE_CONFIG_ROOT") is { Length: > 0 } r ? r : null;
    static string AppData => Root is { } r ? Path.Combine(r, "Roaming") : Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
    static string LocalAppData => Root is { } r ? Path.Combine(r, "Local") : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

    /// <summary>Folders where Claude Desktop reads its configuration (classic installer and the Microsoft Store/MSIX package).</summary>
    /// <remarks>
    /// Claude Desktop 4.x keeps its user data in %LOCALAPPDATA%\Claude-Data\claude_&lt;package&gt;\ and leaves a
    /// ".user-data-left.*" marker in the folder it moved away from; older versions use %APPDATA%\Claude (classic installer)
    /// or the MSIX package's LocalCache\Roaming\Claude.
    /// </remarks>
    public static List<string> ConfigDirs()
    {
        var dirs = new List<string>();
        static bool Left(string d) { try { return Directory.GetFiles(d, ".user-data-left*").Length > 0; } catch { return false; } }
        try
        {
            var data = Path.Combine(LocalAppData, "Claude-Data");
            if (Directory.Exists(data))
                foreach (var d in Directory.GetDirectories(data, "claude*"))
                    if (!Left(d)) dirs.Add(d);
        }
        catch { }
        var classic = Path.Combine(AppData, "Claude");
        if (Directory.Exists(classic) && !Left(classic)) dirs.Add(classic);
        try
        {
            var packages = Path.Combine(LocalAppData, "Packages");
            if (Directory.Exists(packages))
                foreach (var p in Directory.GetDirectories(packages, "Claude_*"))
                {
                    var d = Path.Combine(p, "LocalCache", "Roaming", "Claude");
                    if (Directory.Exists(d) && !Left(d)) dirs.Add(d);
                }
        }
        catch { }
        if (dirs.Count == 0) dirs.Add(classic);
        return dirs;
    }

    public static bool IsInstalled()
    {
        try
        {
            return Directory.Exists(Path.Combine(LocalAppData, "AnthropicClaude")) || Directory.Exists(Path.Combine(AppData, "Claude"))
                || Directory.Exists(Path.Combine(LocalAppData, "Claude-Data"))
                || (Directory.Exists(Path.Combine(LocalAppData, "Packages")) && Directory.GetDirectories(Path.Combine(LocalAppData, "Packages"), "Claude_*").Length > 0);
        }
        catch { return false; }
    }

    public static bool IsConnected()
    {
        foreach (var d in ConfigDirs())
        {
            try
            {
                var f = Path.Combine(d, "claude_desktop_config.json");
                if (File.Exists(f) && JsonNode.Parse(File.ReadAllText(f))?["mcpServers"]?[ServerName] != null) return true;
            }
            catch { }
        }
        return false;
    }

    public static void Connect() => Edit(servers => servers[ServerName] = new JsonObject
    {
        ["command"] = ExePath,
        ["args"] = new JsonArray("--mcp"),
    });

    /// <summary>
    /// Self-repair (classic edition, user allowed Claude Desktop): re-adds the entry where it is missing or points to another
    /// executable — Claude Desktop updates move its data folder and its Cowork tab has been seen rewriting the file.
    /// Only existing Claude folders are touched. Returns true when something was repaired.
    /// </summary>
    public static bool EnsureConnected()
    {
        if (!CanEditConfig || !AiSettings.Load().AllowClaudeDesktop || !IsInstalled()) return false;
        bool repaired = false;
        foreach (var d in ConfigDirs())
        {
            if (!Directory.Exists(d)) continue;
            var f = Path.Combine(d, "claude_desktop_config.json");
            try
            {
                var entry = File.Exists(f) ? JsonNode.Parse(File.ReadAllText(f), documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true })?["mcpServers"]?[ServerName] : null;
                if (entry?["command"]?.GetValue<string>() is { } cmd && string.Equals(cmd, ExePath, StringComparison.OrdinalIgnoreCase)) continue;
                EditOne(d, servers => servers[ServerName] = new JsonObject { ["command"] = ExePath, ["args"] = new JsonArray("--mcp") });
                repaired = true;
            }
            catch { }
        }
        return repaired;
    }

    /// <summary>Folder to show the user ("Mở thư mục cấu hình Claude"): the first place Claude Desktop reads.</summary>
    public static string PrimaryConfigDir => ConfigDirs()[0];

    // ---------------- handshake: proves Claude Desktop really started the server ----------------

    public sealed record Handshake(DateTimeOffset Time, string Client, string Version);

    /// <summary>Written by the MCP server when a client initializes (Data\ai\mcp-last.json).</summary>
    public static void RecordHandshake(string client, string version)
    {
        try
        {
            File.WriteAllText(AiPaths.File("mcp-last.json"), new JsonObject
            {
                ["time"] = DateTimeOffset.Now.ToString("o"), ["client"] = client, ["version"] = version,
            }.ToJsonString());
        }
        catch { }
    }

    /// <summary>Forgets the last handshake (on connect / disconnect the status must reflect the new state).</summary>
    public static void ClearHandshake() { try { File.Delete(AiPaths.File("mcp-last.json")); } catch { } }

    public static Handshake? LastHandshake()
    {
        try
        {
            var p = AiPaths.File("mcp-last.json");
            if (!File.Exists(p)) return null;
            var j = JsonNode.Parse(File.ReadAllText(p))!;
            return new Handshake(DateTimeOffset.Parse(j["time"]!.GetValue<string>()), j["client"]?.GetValue<string>() ?? "", j["version"]?.GetValue<string>() ?? "");
        }
        catch { return null; }
    }

    public static void Disconnect() => Edit(servers => servers.Remove(ServerName));

    static void Edit(Action<JsonObject> change)
    {
        foreach (var d in ConfigDirs()) EditOne(d, change);
    }

    static void EditOne(string d, Action<JsonObject> change)
    {
        {
            Directory.CreateDirectory(d);
            var f = Path.Combine(d, "claude_desktop_config.json");
            JsonObject root;
            if (File.Exists(f))
            {
                var text = File.ReadAllText(f);
                JsonNode? parsed = null;
                var malformed = new InvalidOperationException("claude_desktop_config.json " + MicroadsOffice.Core.Office.Loc.T("bị lỗi định dạng; hãy sửa tệp trong Claude Desktop › Settings › Developer.", "is malformed; fix it in Claude Desktop › Settings › Developer."));
                try { parsed = text.Trim().Length == 0 ? new JsonObject() : JsonNode.Parse(text, documentOptions: new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true }); }
                catch (JsonException) { throw malformed; }
                // never replace a file we do not understand (top level not an object, or "mcpServers" not an object)
                root = parsed as JsonObject ?? throw malformed;
                if (root["mcpServers"] is { } ms && ms is not JsonObject) throw malformed;
                // keep the user's ORIGINAL file once (later edits must not overwrite that backup)
                if (!File.Exists(f + ".microads-backup")) File.Copy(f, f + ".microads-backup");
            }
            else root = new JsonObject();
            if (root["mcpServers"] is not JsonObject servers) { servers = new JsonObject(); root["mcpServers"] = servers; }
            change(servers);
            var tmp = f + ".tmp";
            File.WriteAllText(tmp, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
            File.Move(tmp, f, true);
        }
    }
}
