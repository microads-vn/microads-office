using System.Diagnostics;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice.AI.Tests;

// Microads Office — AI assistant tests (in-process, against a local fake server; no real account, no network).
// usage: MicroadsOffice.AI.Tests [core|word|excel] [outDir] [vi|en]
internal static class Program
{
    /// <summary>Sign-in pages the app tried to open in a browser (none is really opened during tests).</summary>
    public static readonly List<string> BrowserUrls = [];

    [STAThread]
    static int Main(string[] args)
    {
        var mode = args.Length > 0 ? args[0] : "core";
        var outDir = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "ai-test");
        if (Directory.Exists(Path.Combine(outDir, "appdata-" + mode))) Directory.Delete(Path.Combine(outDir, "appdata-" + mode), true);
        Directory.CreateDirectory(outDir);
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        T.Out = outDir;
        AppSettings.BaseDirectory = Path.Combine(outDir, "appdata-" + mode);
        Directory.CreateDirectory(AppSettings.DataDirectory);
        // Everything stays inside outDir: Claude Desktop config, update checks off, no promo, no modal welcome.
        Environment.SetEnvironmentVariable("MICROADS_CLAUDE_CONFIG_ROOT", Path.Combine(outDir, "claude-" + mode));
        Environment.SetEnvironmentVariable("MICROADS_UPDATE_URL", "http://127.0.0.1:9/manifest.json");
        Environment.SetEnvironmentVariable("MICROADS_NO_PROMO", "1");
        Environment.SetEnvironmentVariable("MICROADS_NO_UPDATE", "1");
        Environment.SetEnvironmentVariable("MICROADS_AI_NO_MODALS", "1");
        Loc.Lang = args.Length > 2 && args[2].Length == 2 ? Loc.Normalize(args[2]) : Loc.Vietnamese;
        // never open the user's real browser (e.g. the setup card's "Continue with ChatGPT" starts a sign-in); tests that
        // need the browser's part replace this with a fake that calls the loopback callback
        ChatGptAuth.OpenBrowser = url => { BrowserUrls.Add(url); return false; };
        var r = new Runner();
        // watchdog: a test mode never runs longer than 15 minutes (it takes about one); report and exit instead of hanging
        using var watchdog = new System.Threading.Timer(_ =>
        {
            try { File.WriteAllLines(Path.Combine(outDir, $"report-ai-{mode}.txt"), [.. r.Report, "", "TIMEOUT after 15 minutes — a window or dialog was left waiting"]); } catch { }
            Environment.Exit(4);
        }, null, TimeSpan.FromMinutes(15), Timeout.InfiniteTimeSpan);
        switch (mode)
        {
            case "core": CoreTests.Run(r); break;
            case "word": OfficeShell.Init(AppKind.Word); Loc.Lang = args.Length > 2 ? Loc.Normalize(args[2]) : Loc.Vietnamese; WordUiTests.Run(r); break;
            case "excel": OfficeShell.Init(AppKind.Excel); Loc.Lang = args.Length > 2 ? Loc.Normalize(args[2]) : Loc.Vietnamese; ExcelUiTests.Run(r); break;
            case "pdf": OfficeShell.Init(AppKind.Pdf); Loc.Lang = Loc.Vietnamese; MicroadsOffice.Core.Pdfium.PdfiumLoader.Register(typeof(Program).Assembly); PdfUiTests.Run(r); break;
            case "e2e": E2eTests.Run(r, Path.GetFullPath(args[2]), outDir); break;
            case "shots-word": OfficeShell.Init(AppKind.Word); Loc.Lang = Loc.Vietnamese; StoreShots.RunWord(r, outDir); break;
            case "shots-excel": OfficeShell.Init(AppKind.Excel); Loc.Lang = Loc.Vietnamese; StoreShots.RunExcel(r, outDir); break;
            default: Console.WriteLine("mode: core|word|excel|pdf|e2e"); return 2;
        }
        r.Log($"\nTOTAL: {r.Passed} passed, {r.Failed} failed");
        File.WriteAllLines(Path.Combine(outDir, $"report-ai-{mode}.txt"), r.Report);
        // never linger after the report (a stray window or foreground thread must not keep a build waiting)
        Environment.Exit(r.Failed == 0 ? 0 : 1);
        return 0;
    }
}

public sealed class AssertException(string m) : Exception(m);

public static class T
{
    public static string Out = "";
    public static void True(bool c, string what) { if (!c) throw new AssertException("Expected true: " + what); }
    public static void False(bool c, string what) { if (c) throw new AssertException("Expected false: " + what); }
    public static void Eq<TV>(TV expected, TV actual, string what = "")
    {
        if (!EqualityComparer<TV>.Default.Equals(expected, actual)) throw new AssertException($"Expected <{expected}> but was <{actual}> {what}");
    }
    public static TEx Throws<TEx>(Action a, string what) where TEx : Exception
    {
        try { a(); } catch (TEx e) { return e; } catch (AggregateException ae) when (ae.InnerException is TEx inner) { return inner; }
        throw new AssertException("Expected " + typeof(TEx).Name + ": " + what);
    }
}

public sealed class Runner
{
    public int Passed, Failed;
    public readonly List<string> Report = new();

    public void Test(string name, Action body)
    {
        var sw = Stopwatch.StartNew();
        try { body(); Passed++; Log($"PASS  {name} ({sw.ElapsedMilliseconds} ms)"); }
        catch (Exception e) { Failed++; Log($"FAIL  {name}: {e.GetType().Name}: {e.Message}\n{e.StackTrace?.Split('\n').FirstOrDefault(l => l.Contains("Tests"))?.Trim()}"); }
    }

    public void Log(string s) { Console.WriteLine(s); Report.Add(s); }
}
