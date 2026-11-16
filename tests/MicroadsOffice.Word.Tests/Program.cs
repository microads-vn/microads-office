using System.Diagnostics;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Tests;

namespace MicroadsOffice.Word.Tests
{
    internal static class Program
    {
    // Microads Office Word — automated tests (in-process; no real mouse/keyboard is used).
    // usage: MicroadsOffice.Word.Tests [all|model|io|ui] [outDir] [vi|en]
        [STAThread]
        static int Main(string[] args)
    {
        var mode = args.Length > 0 ? args[0] : "all";
        var outDir = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "word-test");
        Directory.CreateDirectory(outDir);
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        T.Out = outDir;
        AppSettings.BaseDirectory = Path.Combine(outDir, "appdata");
        Environment.SetEnvironmentVariable("MICROADS_UPDATE_URL", "http://127.0.0.1:9/manifest.json");
        Environment.SetEnvironmentVariable("MICROADS_NO_PROMO", "1");
        OfficeShell.Init(AppKind.Word);
        Loc.Lang = args.Length > 2 ? Loc.Normalize(args[2]) : Loc.Vietnamese;
        var r = new Runner();
        if (mode is "all" or "model") ModelTests.Run(r);
        if (mode is "all" or "io") IoTests.Run(r);
        if (mode is "all" or "ui") UiTests.Run(r);
        r.Log($"\nTOTAL: {r.Passed} passed, {r.Failed} failed");
        File.WriteAllLines(Path.Combine(outDir, $"report-word-{mode}.txt"), r.Report);
        return r.Failed == 0 ? 0 : 1;
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
        public static void Near(float expected, float actual, float tol, string what) { if (Math.Abs(expected - actual) > tol) throw new AssertException($"Expected ~{expected} but was {actual} {what}"); }
    }

    public sealed class Runner
    {
        public int Passed, Failed;
        public readonly List<string> Report = new();

        public void Test(string name, Action body)
        {
            var sw = Stopwatch.StartNew();
            try { body(); Passed++; Log($"PASS  {name} ({sw.ElapsedMilliseconds} ms)"); }
            catch (Exception ex)
            {
                Failed++;
                var inner = ex is System.Reflection.TargetInvocationException tie && tie.InnerException != null ? tie.InnerException : ex;
                Log($"FAIL  {name}: {inner.GetType().Name}: {inner.Message}\n      {inner.StackTrace?.Split('\n').FirstOrDefault()?.Trim()}");
            }
        }

        public void Log(string s) { Report.Add(s); Console.WriteLine(s); }
    }
}
