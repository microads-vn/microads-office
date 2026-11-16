// End-to-end test of the Microads Office web installer (MicroadsOffice_Setup.exe), driven through UI Automation patterns only
// (Invoke / Value / Toggle — the real mouse and keyboard are never used).
//   args: <web installer exe> <full setup exe> <full setup .sig.json> <work folder> [<python.exe> <signkey.py> <backup key .pem>]
//   (the optional three enable the minimum-version and backup-key tests, which sign the release again)
// A local HTTP server plays microads.vn (manifest + download). The Microads Office setup that the web installer starts runs in
// sandbox mode (shortcuts + registry redirected into the work folder / HKCU\Software\MicroadsOfficeSetupSandbox, removed at the end).
using System.Diagnostics;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using Microsoft.Win32;

var stubSrc = Path.GetFullPath(args[0]);
var fullSetup = Path.GetFullPath(args[1]);
var sig = JsonDocument.Parse(File.ReadAllText(args[2])).RootElement;
var work = Path.GetFullPath(args[3]);
string? pyExe = args.Length > 6 ? args[4] : null, signKey = args.Length > 6 ? args[5] : null, backupKey = args.Length > 6 ? args[6] : null;
if (Directory.Exists(work)) Directory.Delete(work, true);
Directory.CreateDirectory(work);
var tmp = Path.Combine(work, "tmp"); Directory.CreateDirectory(tmp);
var sandbox = Path.Combine(work, "sandbox");
var urlLog = Path.Combine(work, "opened-urls.txt");
var inst = Path.Combine(work, "inst");
var inst2 = Path.Combine(work, "inst2");
var inst3 = Path.Combine(work, "inst3");
var VER = sig.GetProperty("version").GetString()!;
var SHA = sig.GetProperty("sha256").GetString()!;
var SIG = sig.GetProperty("signature").GetString()!;
var fullBytes = File.ReadAllBytes(fullSetup);
int pass = 0, fail = 0;
var report = new List<string>();
void Log(string s) { Console.WriteLine(s); report.Add(s); }
void Check(bool ok, string what) { if (ok) pass++; else fail++; Log((ok ? "PASS  " : "FAIL  ") + what); }

// ---------------------------------------------------------------- local "microads.vn"
int port = 18766;
string manifest = "";
int manifestStatus = 200;
byte[] payload = fullBytes;
int throttle = 0; // bytes per second, 0 = full speed
var http = new HttpListener();
http.Prefixes.Add($"http://localhost:{port}/");
http.Start();
_ = Task.Run(async () =>
{
    while (http.IsListening)
    {
        HttpListenerContext ctx;
        try { ctx = await http.GetContextAsync(); } catch { break; }
        _ = Task.Run(async () =>
        {
            try
            {
                var isExe = ctx.Request.Url!.AbsolutePath.EndsWith(".exe");
                if (!isExe && manifestStatus != 200) { ctx.Response.StatusCode = manifestStatus; ctx.Response.Close(); return; }
                var body = isExe ? payload : Encoding.UTF8.GetBytes(manifest);
                ctx.Response.ContentType = isExe ? "application/vnd.microsoft.portable-executable" : "application/json";
                ctx.Response.ContentLength64 = body.Length;
                if (isExe && throttle > 0)
                {
                    int chunk = Math.Max(4096, throttle / 10);
                    for (int o = 0; o < body.Length; o += chunk) { await ctx.Response.OutputStream.WriteAsync(body.AsMemory(o, Math.Min(chunk, body.Length - o))); await Task.Delay(100); }
                }
                else await ctx.Response.OutputStream.WriteAsync(body);
                ctx.Response.Close();
            }
            catch { try { ctx.Response.Abort(); } catch { } }
        });
    }
});
string Manifest(string? version = null, string? url = null, string? signature = null, bool noRelease = false) => JsonSerializer.Serialize(new
{
    schema = 1,
    latest = noRelease ? null : new
    {
        version = version ?? VER, date = "2026-11-16", url = url ?? $"http://localhost:{port}/download/MicroadsOffice_Setup_{VER}.exe",
        size = fullBytes.Length, sha256 = SHA, signature = signature ?? SIG, notes = new { vi = "", en = "" }
    },
    forceUpdate = new { enabled = false, commandId = "" },
    promo = (object?)null
});

using var A = new UIA3Automation();
var desktop = A.GetDesktop();

Process Start(string exe, string arguments, Dictionary<string, string>? extra = null)
{
    var psi = new ProcessStartInfo(exe, arguments) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(exe)! };
    psi.Environment["MICROADS_UPDATE_URL"] = $"http://localhost:{port}/api/manifest.json";
    psi.Environment["MICROADS_SETUP_SANDBOX"] = sandbox;
    psi.Environment["MICROADS_DATA_DIR"] = Path.Combine(work, "data");
    psi.Environment["MICROADS_WEBSETUP_OPENURL_LOG"] = urlLog;
    psi.Environment["TEMP"] = tmp; psi.Environment["TMP"] = tmp;
    if (extra != null) foreach (var kv in extra) psi.Environment[kv.Key] = kv.Value;
    return Process.Start(psi)!;
}
T? WaitFor<T>(Func<T?> f, int ms = 20000) where T : class
{
    var sw = Stopwatch.StartNew();
    while (sw.ElapsedMilliseconds < ms)
    {
        try { var v = f(); if (v != null) return v; } catch { }
        Thread.Sleep(200);
    }
    return null;
}
bool WaitTrue(Func<bool> f, int ms = 20000) => WaitFor(() => f() ? "ok" : null, ms) != null;
string[] ProcNames() => ["MicroadsOffice_Setup", "MicroadsOffice_Setup_" + VER, "MicroadsOffice_Setup_9.9.9", "MicroadsOffice"];
HashSet<int> OurPids()
{
    var set = new HashSet<int>();
    foreach (var n in ProcNames())
        foreach (var p in Process.GetProcessesByName(n))
            try { if (p.MainModule?.FileName?.StartsWith(work, StringComparison.OrdinalIgnoreCase) == true) set.Add(p.Id); } catch { }
    return set;
}
bool FullSetupRunning() => Process.GetProcessesByName("MicroadsOffice_Setup_" + VER).Any(p => { try { return p.MainModule?.FileName?.StartsWith(work, StringComparison.OrdinalIgnoreCase) == true; } catch { return false; } });
AutomationElement? TopWindow(string automationId)
{
    var pids = OurPids();
    return desktop.FindAllChildren().FirstOrDefault(w => pids.Contains(w.Properties.ProcessId.ValueOrDefault) && w.Properties.AutomationId.ValueOrDefault == automationId);
}
AutomationElement? Find(AutomationElement root, string id) => root.FindFirstDescendant(cf => cf.ByAutomationId(id));
AutomationElement Need(AutomationElement root, string id, int ms = 10000) => WaitFor(() => Find(root, id), ms) ?? throw new Exception("element not found: " + id);
string Text(AutomationElement e) => e.Properties.Name.ValueOrDefault ?? "";
void Invoke(AutomationElement e)
{
    if (e.Patterns.Invoke.IsSupported) e.Patterns.Invoke.Pattern.Invoke();
    else if (e.Patterns.LegacyIAccessible.IsSupported) e.Patterns.LegacyIAccessible.Pattern.DoDefaultAction();
    else throw new Exception("not invokable: " + e.Properties.AutomationId.ValueOrDefault);
}
void KillAll() { foreach (var id in OurPids()) try { Process.GetProcessById(id).Kill(true); } catch { } }
string[] TempWork() => Directory.GetDirectories(tmp, "MicroadsOffice-Setup-*");
string[] OpenedUrls() => File.Exists(urlLog) ? File.ReadAllLines(urlLog) : [];
void Step(string name, Action a)
{
    Log($"\n== {name}");
    try { a(); }
    catch (Exception ex) { fail++; Log($"FAIL  {name}: {ex.GetType().Name}: {ex.Message}"); }
    finally { KillAll(); Thread.Sleep(800); throttle = 0; payload = fullBytes; manifestStatus = 200; }
}
JsonElement Sign(string version, string? keyFile)
{
    var psi = new ProcessStartInfo(pyExe!, $"-I \"{signKey}\" sign \"{fullSetup}\" {version}") { RedirectStandardOutput = true, UseShellExecute = false };
    psi.Environment["PYTHONDONTWRITEBYTECODE"] = "1";
    if (keyFile != null) psi.Environment["MICROADS_SIGNING_KEY"] = keyFile;
    using var p = Process.Start(psi)!;
    var o = p.StandardOutput.ReadToEnd();
    p.WaitForExit();
    return JsonDocument.Parse(o).RootElement;
}
static bool HasMotw(string f)
{
    try { using var _ = new FileStream(f + ":Zone.Identifier", FileMode.Open, FileAccess.Read); return true; }
    catch { return false; }
}
bool IsVi(AutomationElement w) => Text(Need(w, "btnLang")) == "English";

/// Drives the Microads Office setup started by the web installer up to the finish page.
void CompleteSetup(AutomationElement s, string dir, bool vi)
{
    Invoke(Need(s, "btnSetupNext"));                                   // language page → options
    var box = Need(s, "txtInstallDir");
    box.Patterns.Value.Pattern.SetValue(dir);
    var launch = Need(s, "chkLaunch");
    if (launch.Patterns.Toggle.Pattern.ToggleState.Value == ToggleState.On) launch.Patterns.Toggle.Pattern.Toggle();
    Invoke(Need(s, "btnSetupNext"));                                   // Install
    Check(WaitTrue(() => Text(Need(s, "lblSetupTitle", 60000)).Contains(vi ? "hoàn tất" : "complete"), 90000), "setup installs Microads Office " + VER);
    Invoke(Need(s, "btnSetupNext"));                                   // Finish
}

try
{
    var dlDir = Path.Combine(work, "Downloads"); Directory.CreateDirectory(dlDir);
    var stub = Path.Combine(dlDir, "MicroadsOffice_Setup.exe");
    File.Copy(stubSrc, stub);
    // what Chrome / Edge do with a downloaded file: attach the "Mark of the Web"
    File.WriteAllText(stub + ":Zone.Identifier", "[ZoneTransfer]\r\nZoneId=3\r\nReferrerUrl=https://microads.vn/office/\r\nHostUrl=https://microads.vn/office/download/MicroadsOffice_Setup.exe\r\n");
    Check(HasMotw(stub), "test precondition: the downloaded web installer carries the Mark of the Web");
    var fv = FileVersionInfo.GetVersionInfo(stub);
    Check(fv.CompanyName?.Contains("Microads") == true && fv.ProductName == "Microads Office" && fv.FileDescription == "Microads Office Setup" && fv.OriginalFilename == "MicroadsOffice_Setup.exe",
        $"publisher information embedded in the file ({fv.CompanyName} / {fv.ProductName} / {fv.FileDescription})");
    Check(new FileInfo(stub).Length < 300_000, $"web installer is small ({new FileInfo(stub).Length / 1024} KB)");

    // ============================================================ 1. normal path
    Step("download → verify → setup → installed", () =>
    {
        manifest = Manifest();
        var p = Start(stub, "");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("web installer window missing");
        bool vi = IsVi(w);
        Log($"      language detected: {(vi ? "vi" : "en")}");
        Check(Text(Need(w, "lblTitle")) == "Microads Office", "window shows Microads Office");
        Check(Text(Need(w, "lblNote")).Contains("microads.vn") && Text(Need(w, "btnWebsite")).Contains("microads.vn"), "window explains the download source and shows the publisher (microads.vn)");
        var s = WaitFor(() => TopWindow("setupForm"), 60000) ?? throw new Exception("setup window did not appear");
        Check(true, "latest version downloaded, signature verified, Microads Office setup started");
        Check(WaitTrue(() => TopWindow("webSetupForm") == null, 10000), "web installer window steps aside while the setup is open");
        var dl = TempWork().SelectMany(d => Directory.GetFiles(d, "*.exe")).ToArray();
        Check(dl.Length == 1 && Path.GetFileName(dl[0]) == $"MicroadsOffice_Setup_{VER}.exe", "setup downloaded into a private temporary folder");
        Check(dl.Length == 1 && Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(dl[0]))).Equals(SHA, StringComparison.OrdinalIgnoreCase), "downloaded file is byte-identical to the signed release");
        Check(dl.Length == 1 && !HasMotw(dl[0]), "downloaded setup has NO Mark of the Web → Windows does not show a second warning");
        Check(Text(Need(s, "btnSetupNext")) == (vi ? "Tiếp tục" : "Next"), "setup opens in the web installer's language");
        CompleteSetup(s, inst, vi);
        Check(WaitTrue(() => p.HasExited, 30000) && p.ExitCode == 0, $"web installer ends with exit code 0 after the setup ({(p.HasExited ? p.ExitCode : -1)})");
        Check(WaitTrue(() => TempWork().Length == 0, 15000), "temporary download folder removed");
        var exe = Path.Combine(inst, "MicroadsOffice.exe");
        Check(File.Exists(exe) && Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(exe))).Equals(SHA, StringComparison.OrdinalIgnoreCase), "Microads Office installed (verified program)");
        Check(File.Exists(exe) && !HasMotw(exe), "installed program has no Mark of the Web (Word/Excel/PDF open without warnings)");
        using var un = Registry.CurrentUser.OpenSubKey(@"Software\MicroadsOfficeSetupSandbox\Software\Microsoft\Windows\CurrentVersion\Uninstall\MicroadsOffice");
        Check(un?.GetValue("DisplayVersion") as string == VER, "Settings › Apps entry registered with version " + VER);
    });

    // ============================================================ 2. server error, language switch, full installer, retry
    Step("server error → English → full installer link → retry", () =>
    {
        manifest = Manifest();
        manifestStatus = 500;
        var p = Start(stub, "--lang=vi");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("Không tải được"), 30000), "connection problem reported in Vietnamese");
        Check(Find(w, "btnRetry")?.IsOffscreen == false && Find(w, "btnFullInstaller")?.IsOffscreen == false, "Thử lại and Tải bộ cài đầy đủ buttons shown");
        Check(Text(Need(w, "btnCancel")) == "Đóng", "Đóng (close) button shown");
        Invoke(Need(w, "btnLang"));
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("Could not download")), "English button switches the window to English");
        Check(Text(Need(w, "btnLang")) == "Tiếng Việt" && Text(Need(w, "btnCancel")) == "Close" && Text(Need(w, "btnRetry")) == "Try again", "all buttons translated");
        Invoke(Need(w, "btnFullInstaller"));
        Check(WaitTrue(() => OpenedUrls().Contains("https://microads.vn/office/"), 5000), "Full installer button opens microads.vn/office");
        manifestStatus = 200;
        Invoke(Need(w, "btnRetry"));
        var s = WaitFor(() => TopWindow("setupForm"), 60000);
        Check(s != null, "Try again downloads and starts the setup once the server answers");
        Check(s != null && Text(Need(s, "btnSetupNext")) == "Next", "setup follows the language switched in the web installer (English)");
        if (s != null) Invoke(Need(s, "btnSetupCancel"));
        Check(WaitTrue(() => p.HasExited, 20000) && p.ExitCode == 1, "closing the setup without installing ends the web installer (code 1)");
        Check(WaitTrue(() => TempWork().Length == 0, 15000), "temporary download removed");
    });

    // ============================================================ 3. cancel while downloading, website link
    Step("cancel while downloading + microads.vn link", () =>
    {
        manifest = Manifest();
        throttle = 300_000;
        var p = Start(stub, "--lang=vi");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("Đang tải Microads Office " + VER), 20000), "status: downloading version " + VER);
        Check(WaitTrue(() => Text(Need(w, "lblDetail")).Contains("MB") && Need(w, "progressDownload").Patterns.RangeValue.Pattern.Value.Value > 0, 15000),
            $"progress bar and MB counter move ({Text(Need(w, "lblDetail"))})");
        Invoke(Need(w, "btnWebsite"));
        Check(WaitTrue(() => OpenedUrls().Contains("https://microads.vn"), 5000), "publisher link opens microads.vn");
        Invoke(Need(w, "btnCancel"));
        Check(WaitTrue(() => p.HasExited, 15000) && p.ExitCode == 2, "Hủy stops the download and closes (code 2)");
        Check(!FullSetupRunning(), "nothing was started");
        Check(WaitTrue(() => TempWork().Length == 0, 15000), "partial download removed");
    });

    // ============================================================ 4. tampered file
    Step("tampered download is refused", () =>
    {
        manifest = Manifest();
        var bad = (byte[])fullBytes.Clone();
        bad[bad.Length / 2] ^= 0xFF; // same size, one byte changed
        payload = bad;
        var p = Start(stub, "--lang=vi");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("không hợp lệ"), 60000), "modified file rejected (SHA-256 / signature)");
        Check(!FullSetupRunning(), "modified file was never run");
        Check(TempWork().SelectMany(d => Directory.GetFiles(d)).Count() == 0, "modified file deleted");
        Invoke(Need(w, "btnCancel"));
        Check(WaitTrue(() => p.HasExited, 10000) && p.ExitCode == 5, "Đóng closes (code 5)");
    });

    // ============================================================ 5. signature of another version
    Step("forged release information is refused", () =>
    {
        manifest = Manifest(version: "9.9.9"); // real file, but the signature covers 2.0.x, not 9.9.9
        var p = Start(stub, "--lang=en");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("Invalid download"), 60000), "signature not matching the announced version → refused");
        Check(Text(Need(w, "lblDetail")).Contains("signature"), "explains: no valid Microads signature");
        Check(!Process.GetProcessesByName("MicroadsOffice_Setup_9.9.9").Any(), "nothing started");
        Invoke(Need(w, "btnCancel"));
        Check(WaitTrue(() => p.HasExited, 10000) && p.ExitCode == 5, "Close closes (code 5)");
    });

    // ============================================================ 6. insecure address
    Step("non-HTTPS download address is refused", () =>
    {
        manifest = Manifest(url: $"http://example.com/MicroadsOffice_Setup_{VER}.exe");
        var p = Start(stub, "--lang=vi");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("không hợp lệ") && Text(Need(w, "lblDetail")).Contains("HTTPS"), 30000), "plain-HTTP link in the manifest refused before downloading");
        Invoke(Need(w, "btnCancel"));
        Check(WaitTrue(() => p.HasExited, 10000), "closes");
    });

    // ============================================================ 7. nothing published yet
    Step("no release published", () =>
    {
        manifest = Manifest(noRelease: true);
        var p = Start(stub, "--lang=vi");
        var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
        Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("Chưa có phiên bản"), 30000), "explains that no version is available yet");
        Invoke(Need(w, "btnCancel"));
        Check(WaitTrue(() => p.HasExited, 10000) && p.ExitCode == 4, "closes (code 4)");
    });

    // ============================================================ 8. silent (IT / many computers)
    Step("silent install (--silent --dir)", () =>
    {
        manifest = Manifest();
        bool sawWindow = false;
        var p = Start(stub, $"--silent --dir \"{inst2}\"");
        var sw = Stopwatch.StartNew();
        while (!p.HasExited && sw.ElapsedMilliseconds < 120000)
        {
            if (TopWindow("webSetupForm") != null || TopWindow("setupForm") != null) sawWindow = true;
            Thread.Sleep(300);
        }
        Check(p.HasExited && p.ExitCode == 0, $"silent install succeeds (code {(p.HasExited ? p.ExitCode : -1)})");
        Check(!sawWindow, "no window shown");
        Check(File.Exists(Path.Combine(inst2, "MicroadsOffice.exe")), "installed into the --dir folder");
        Check(TempWork().Length == 0, "temporary download removed");
        var q = Start(stub, "--silent", new() { ["MICROADS_UPDATE_URL"] = "http://localhost:1/api/manifest.json" });
        Check(WaitTrue(() => q.HasExited, 60000) && q.ExitCode == 4, "silent install without network reports exit code 4");
        Check(File.Exists(Path.Combine(tmp, "MicroadsOffice_Setup.log")) && File.ReadAllText(Path.Combine(tmp, "MicroadsOffice_Setup.log")).Contains("failed (Network)"), "reason written to %TEMP%\\MicroadsOffice_Setup.log");
    });

    // ============================================================ 8b. --dir must be a full path
    Step("relative --dir is refused", () =>
    {
        manifest = Manifest();
        var log = Path.Combine(tmp, "MicroadsOffice_Setup.log");
        var p = Start(stub, "--silent --dir \"Microads Office\"");
        Check(WaitTrue(() => p.HasExited, 30000) && p.ExitCode == 6, "relative install folder refused before downloading (code 6) — it would land in the temporary folder");
        Check(TempWork().Length == 0 && !FullSetupRunning(), "nothing downloaded or started");
        Check(File.Exists(log) && File.ReadAllText(log).Contains("invalid --dir"), "reason written to the log");
    });

    if (pyExe != null)
    {
        // ============================================================ 8c. no downgrade to an older signed build
        Step("older signed release is refused", () =>
        {
            var old = Sign("2.0.0", null);
            Check(old.GetProperty("signature").GetString()!.Length > 0, "test precondition: release signed as 2.0.0 with the real key");
            manifest = Manifest(version: "2.0.0", signature: old.GetProperty("signature").GetString());
            var p = Start(stub, "--lang=vi");
            var w = WaitFor(() => TopWindow("webSetupForm"), 20000) ?? throw new Exception("window missing");
            Check(WaitTrue(() => Text(Need(w, "lblStatus")).Contains("không hợp lệ") && Text(Need(w, "lblDetail")).Contains("2.0.1"), 30000),
                "validly signed but older than 2.0.1 → refused (a hacked server cannot push old builds)");
            Check(!FullSetupRunning() && TempWork().Length == 0, "nothing downloaded or started");
            Invoke(Need(w, "btnCancel"));
            Check(WaitTrue(() => p.HasExited, 10000) && p.ExitCode == 5, "closes (code 5)");
        });

        // ============================================================ 8d. backup signing key + %VARIABLE% in --dir
        Step("backup signing key accepted, --dir with %VARIABLE%", () =>
        {
            var b = Sign(VER, backupKey);
            manifest = Manifest(signature: b.GetProperty("signature").GetString());
            Check(b.GetProperty("signature").GetString() != SIG, "test precondition: release signed with the backup key");
            var p = Start(stub, "--silent --dir \"%MO_E2E_DIR%\"", new() { ["MO_E2E_DIR"] = inst3 });
            Check(WaitTrue(() => p.HasExited, 120000) && p.ExitCode == 0, $"release signed with the backup key installs (code {(p.HasExited ? p.ExitCode : -1)})");
            Check(File.Exists(Path.Combine(inst3, "MicroadsOffice.exe")), "%MO_E2E_DIR% expanded to the install folder");
        });
    }
    else Log("NOTE  minimum-version and backup-key tests skipped (no python/signkey arguments)");

    // ============================================================ 9. clean up the test installations
    Step("uninstall test installations", () =>
    {
        foreach (var d in new[] { inst, inst2, inst3 })
        {
            var exe = Path.Combine(d, "MicroadsOffice.exe");
            if (!File.Exists(exe)) continue;
            var p = Start(exe, "--uninstall --silent");
            WaitTrue(() => p.HasExited, 20000);
            Check(WaitTrue(() => !File.Exists(exe), 20000), "removed " + d);
        }
    });
}
finally
{
    KillAll();
    http.Stop();
    try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\MicroadsOfficeSetupSandbox", false); } catch { }
    Log($"\nTOTAL: {pass} passed, {fail} failed");
    File.WriteAllLines(Path.Combine(work, "..", "report-websetup.txt"), report);
}
return fail == 0 ? 0 : 1;
