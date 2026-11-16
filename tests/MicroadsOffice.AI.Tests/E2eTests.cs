using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;

namespace MicroadsOffice.AI.Tests;

/// <summary>
/// Cross-process end-to-end test with the real executable: Microads Word / Excel run as separate processes, and a second
/// "MicroadsOffice.exe --mcp" process (exactly what Claude Desktop starts) reads and edits their documents over stdio.
/// usage: MicroadsOffice.AI.Tests e2e outDir path\to\MicroadsOffice.exe
/// </summary>
public static class E2eTests
{
    public static void Run(Runner r, string exe, string outDir)
    {
        var data = Path.Combine(outDir, "e2e-data");
        if (Directory.Exists(data)) Directory.Delete(data, true);
        Directory.CreateDirectory(Path.Combine(data, "Data"));
        // the user has connected Claude Desktop (AllowClaudeDesktop) and switched update checks off
        File.WriteAllText(Path.Combine(data, "Data", "office.json"),
            new JsonObject
            {
                ["Language"] = "vi", ["AutoCheckUpdates"] = false,
                ["AppData"] = new JsonObject { ["ai"] = "{\"AllowClaudeDesktop\":true}" },
            }.ToJsonString());
        var apps = new List<Process>();
        Process Start(string args)
        {
            var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false };
            psi.Environment["MICROADS_DATA_DIR"] = data;
            psi.Environment["MICROADS_NO_PROMO"] = "1";
            psi.Environment["MICROADS_NO_UPDATE"] = "1";
            return Process.Start(psi)!;
        }
        Mcp? mcp = null;
        try
        {
            mcp = new Mcp(exe, data);

            r.Test("E01 '--mcp' answers initialize and lists the tools (real exe, stdio)", () =>
            {
                var init = mcp.Call("initialize", new JsonObject { ["protocolVersion"] = "2025-06-18", ["capabilities"] = new JsonObject(), ["clientInfo"] = new JsonObject { ["name"] = "e2e", ["version"] = "1" } });
                T.Eq("microads-office", init["serverInfo"]!["name"]!.GetValue<string>());
                mcp.Notify("notifications/initialized");
                T.True(mcp.Call("tools/list", null)["tools"]!.AsArray().Count >= 8, "tools");
                T.True(mcp.Tool("list_open_apps", []).Contains("word: not open"), "nothing open yet");
            });

            r.Test("E02 open_app starts Microads Word and Claude writes into its document (one undo step)", () =>
            {
                var started = mcp.Tool("open_app", new JsonObject { ["app"] = "word" });
                T.True(started.Contains("started") || started.Contains("already open"), started);
                foreach (var p in OurProcesses(exe)) if (!apps.Any(a => a.Id == p.Id) && p.Id != mcp.Pid) apps.Add(p);
                T.True(mcp.Tool("list_open_apps", []).Contains("word: open"), "word open");
                var w = mcp.Tool("write_document", new JsonObject { ["content"] = "# Tiêu đề E2E\n\nĐoạn có chữ **đậm** và *nghiêng*.\n\n| A | B |\n|---|---|\n| 1 | 2 |", ["mode"] = "insert" });
                T.True(w.Contains("Đã chèn"), w);
                var doc = mcp.Tool("read_document", new JsonObject { ["app"] = "word" });
                T.True(doc.Contains("# Tiêu đề E2E") && doc.Contains("**đậm**") && doc.Contains("*nghiêng*") && doc.Contains("| 1 | 2 |"), doc);
                mcp.Tool("undo", new JsonObject { ["app"] = "word" });
                T.False(mcp.Tool("read_document", new JsonObject { ["app"] = "word" }).Contains("Tiêu đề E2E"), "undone");
            });

            r.Test("E03 Claude fills an Excel sheet with values and formulas; the formula is calculated", () =>
            {
                var ex = Start("--app=excel");
                apps.Add(ex);
                var until = DateTime.Now.AddSeconds(40);
                while (!mcp.Tool("list_open_apps", []).Contains("excel: open")) { if (DateTime.Now > until) throw new AssertException("excel did not start"); Thread.Sleep(500); }
                var guide = mcp.Tool("get_editing_guide", new JsonObject { ["app"] = "excel" });
                T.True(guide.Contains("microads-sheet") && guide.Contains("VLOOKUP"), "guide");
                var res = mcp.Tool("edit_sheet", new JsonObject { ["actions_json"] = "{\"actions\":[{\"range\":\"A1\",\"values\":[[1,2,\"=A1+B1\"]]},{\"range\":\"A1:C1\",\"bold\":true}]}" });
                T.True(res.Contains("Đã ghi"), res);
                var sheet = mcp.Tool("read_document", new JsonObject { ["app"] = "excel" });
                T.True(sheet.Contains("C1: 3  [=A1+B1]"), sheet);
                var bad = mcp.ToolRaw("edit_sheet", new JsonObject { ["actions_json"] = "{\"actions\":[{\"cell\":\"D1\",\"formula\":\"=NOPE(1)\"}]}" });
                T.True(bad["isError"]!.GetValue<bool>() && bad["content"]![0]!["text"]!.GetValue<string>().Contains("NOPE"), "invalid formula refused");
            });

            r.Test("E04 access is refused after the user disconnects Claude Desktop", () =>
            {
                var cfg = JsonNode.Parse(File.ReadAllText(Path.Combine(data, "Data", "office.json")))!;
                cfg["AppData"]!["ai"] = "{\"AllowClaudeDesktop\":false}";
                File.WriteAllText(Path.Combine(data, "Data", "office.json"), cfg.ToJsonString());
                var res = mcp.ToolRaw("read_document", new JsonObject { ["app"] = "word" });
                T.True(res["isError"]!.GetValue<bool>(), "refused");
            });
        }
        finally
        {
            mcp?.Dispose();
            // close only the processes started from THIS test executable (never the user's installed Microads Office)
            foreach (var p in apps.Concat(OurProcesses(exe))) { try { if (!p.HasExited) p.Kill(true); } catch { } }
        }
    }

    static IEnumerable<Process> OurProcesses(string exe)
    {
        var full = Path.GetFullPath(exe);
        foreach (var p in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(exe)))
        {
            string? path = null;
            try { path = p.MainModule?.FileName; } catch { }
            if (path != null && string.Equals(Path.GetFullPath(path), full, StringComparison.OrdinalIgnoreCase)) yield return p;
        }
    }

    sealed class Mcp : IDisposable
    {
        readonly Process _p;
        int _id;
        public int Pid => _p.Id;
        public Mcp(string exe, string data)
        {
            var psi = new ProcessStartInfo(exe, "--mcp")
            {
                UseShellExecute = false, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = new UTF8Encoding(false), StandardInputEncoding = new UTF8Encoding(false),
            };
            psi.Environment["MICROADS_DATA_DIR"] = data;
            psi.Environment["MICROADS_NO_PROMO"] = "1";
            psi.Environment["MICROADS_NO_UPDATE"] = "1";
            _p = Process.Start(psi)!;
            _p.StandardInput.AutoFlush = true;
        }

        public JsonNode Call(string method, JsonObject? p)
        {
            int id = ++_id;
            _p.StandardInput.Write(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["method"] = method, ["params"] = p ?? new JsonObject() }.ToJsonString() + "\n");
            var t = _p.StandardOutput.ReadLineAsync();
            if (!t.Wait(60000)) throw new AssertException("no MCP response to " + method);
            var j = JsonNode.Parse(t.Result ?? throw new AssertException("MCP closed"))!;
            if (j["error"] != null) throw new AssertException("MCP error: " + j["error"]!.ToJsonString());
            return j["result"]!;
        }

        public void Notify(string method) => _p.StandardInput.Write(new JsonObject { ["jsonrpc"] = "2.0", ["method"] = method }.ToJsonString() + "\n");

        public JsonNode ToolRaw(string name, JsonObject args) => Call("tools/call", new JsonObject { ["name"] = name, ["arguments"] = args });

        public string Tool(string name, JsonObject args)
        {
            var r = ToolRaw(name, args);
            var text = r["content"]![0]!["text"]!.GetValue<string>();
            if (r["isError"]?.GetValue<bool>() == true && name != "list_open_apps") throw new AssertException(name + " failed: " + text);
            return text;
        }

        public void Dispose() { try { _p.StandardInput.Close(); if (!_p.WaitForExit(5000)) _p.Kill(); } catch { } }
    }
}
