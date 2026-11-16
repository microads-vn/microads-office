using MicroadsOffice.Tests;

var mode = args.Length > 0 ? args[0] : "core";
var outDir = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "test-output");
Directory.CreateDirectory(outDir);
CoreTests.OutDir = outDir;
var runner = new TestRunner();

switch (mode)
{
    case "core":
        CoreTests.Run(runner);
        PrintTests.Run(runner);
        break;
    case "ui":
        UiTests.Run(runner, args.Length > 2 ? args[2] : throw new ArgumentException("exe path required"), outDir);
        break;
    case "brand":
        MicroadsOffice.Core.Branding.BrandArt.WriteIcon(args.Length > 2 ? args[2] : Path.Combine(outDir, "app.ico"));
        foreach (var sz in new[] { 16, 32, 48, 256 }) using (var b = MicroadsOffice.Core.Branding.BrandArt.RenderAppIcon(sz)) b.Save(Path.Combine(outDir, $"icon{sz}.png"));
        using (var wm = MicroadsOffice.Core.Branding.BrandArt.RenderWordmark(100)) wm.Save(Path.Combine(outDir, "wordmark.png"));
        return 0;
    case "preview":
        PreviewProbe.Run(args[2], Path.Combine(outDir, "probe.png"), args[3]);
        return 0;
    case "render":
        RenderProbe.Run(outDir);
        return 0;
    case "samples":
        SampleFiles.Write(outDir);
        return 0;
    default:
        Console.WriteLine("usage: tests [core|ui|samples] outDir [exePath]");
        return 2;
}

runner.Log($"\nTOTAL: {runner.Passed} passed, {runner.Failed} failed, {runner.Skipped} skipped");
File.WriteAllLines(Path.Combine(outDir, $"report-{mode}.txt"), runner.Report);
return runner.Failed == 0 ? 0 : 1;
