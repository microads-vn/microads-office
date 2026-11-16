using System.Diagnostics;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using Application = FlaUI.Core.Application;

namespace MicroadsOffice.Tests;

public static partial class UiTests
{
    public static string Shots = "";

    public static Application Launch(string exe, string dataDir, params string[] files)
    {
        // The suite exe dispatches by --app; the PDF app is tested in Vietnamese (expected texts below are Vietnamese).
        var psi = new ProcessStartInfo(exe, "--app=pdf " + string.Join(" ", files.Select(f => "\"" + f + "\""))) { UseShellExecute = false };
        psi.Environment["MICROADS_DATA_DIR"] = dataDir;
        psi.Environment["MICROADS_NO_PROMO"] = "1";
        psi.Environment["MICROADS_NO_UPDATE"] = "1";
        var dataFolder = Path.Combine(dataDir, "Data");
        Directory.CreateDirectory(dataFolder);
        var officeJson = Path.Combine(dataFolder, "office.json");
        if (!File.Exists(officeJson)) File.WriteAllText(officeJson, "{\"Language\":\"vi\",\"AutoCheckUpdates\":false}");
        psi.Environment["MICROADS_TEST_PRINT_TO_FILE"] = Path.Combine(dataDir, "guard-print.pdf");
        // Run the app with the real user profile (the build shell redirects profile variables to E:).
        var profile = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory)!, "Users", Environment.UserName);
        if ((psi.Environment.TryGetValue("USERPROFILE", out var up) ? up : "")?.Contains(".build") == true && Directory.Exists(profile))
        {
            psi.Environment["USERPROFILE"] = profile;
            psi.Environment["APPDATA"] = Path.Combine(profile, "AppData", "Roaming");
            psi.Environment["LOCALAPPDATA"] = Path.Combine(profile, "AppData", "Local");
            psi.Environment.Remove("HOME");
        }
        return Application.Launch(psi);
    }

    public static Window MainWindow(Application app, UIA3Automation a)
    {
        var w = Retry.WhileNull(() => app.GetMainWindow(a, TimeSpan.FromSeconds(2)), TimeSpan.FromSeconds(30)).Result;
        if (w == null) throw new AssertException("Main window not found");
        return w;
    }

    public static Func<bool>? ShotAllowed;

    public static void Shot(AutomationElement el, string name)
    {
        if (ShotAllowed != null && !ShotAllowed()) return; // another app is on top: never capture it (privacy)
        try { Capture.Element(el).ToFile(Path.Combine(Shots, name + ".png")); } catch { }
    }

    public static void Dump(AutomationElement el, TextWriter w, int depth = 0, int maxDepth = 12)
    {
        if (depth > maxDepth) return;
        string id = "", name = "", type = "";
        try { id = el.Properties.AutomationId.ValueOrDefault ?? ""; } catch { }
        try { name = el.Properties.Name.ValueOrDefault ?? ""; } catch { }
        try { type = el.Properties.ControlType.ValueOrDefault.ToString(); } catch { }
        w.WriteLine($"{new string(' ', depth * 2)}{type} id='{id}' name='{name.Replace("\n", "\\n")}'");
        foreach (var c in el.FindAllChildren()) Dump(c, w, depth + 1, maxDepth);
    }

    public static void Run(TestRunner t, string exe, string outDir)
    {
        Shots = Path.Combine(outDir, "shots");
        Directory.CreateDirectory(Shots);
        var mode = Environment.GetEnvironmentVariable("UI_MODE") ?? "full";
        if (mode == "explore") Explore(t, exe, outDir);
        else FullUi.Run(t, exe, outDir);
    }

    private static void Explore(TestRunner t, string exe, string outDir)
    {
        var data = Path.Combine(outDir, "ui-data-explore");
        Directory.CreateDirectory(data);
        var sample = Path.Combine(outDir, "sample-3pages.pdf");
        File.WriteAllBytes(sample, TestPdf.Create(3, withLink: true, withOutline: true));
        using var app = Launch(exe, data, sample);
        using var a = new UIA3Automation();
        try
        {
            var w = MainWindow(app, a);
            Thread.Sleep(2500);
            Shot(w, "explore-main");
            using var sw = new StreamWriter(Path.Combine(outDir, "uia-tree.txt"));
            Dump(w, sw);
            t.Log("explored");
        }
        finally { app.Close(); Thread.Sleep(500); if (!app.HasExited) app.Kill(); }
    }
}
