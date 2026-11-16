using MicroadsOffice.Excel.Tests;

// Microads Office Excel tests: engine, file formats and an in-process UI pass (no real mouse/keyboard).
var outDir = args.Length > 1 ? args[1] : @"E:\Office\.build\excel-test";
Directory.CreateDirectory(outDir);
var mode = args.Length > 0 ? args[0] : "all";
var t = new TestRunner();
if (mode is "all" or "engine") EngineTests.Run(t);
if (mode is "all" or "io") IoTests.Run(t, outDir);
if (mode is "all" or "ui") UiTests.Run(t, outDir);
t.Log($"\nTOTAL: {t.Passed} passed, {t.Failed} failed");
File.WriteAllLines(Path.Combine(outDir, $"report-{mode}.txt"), t.Report);
return t.Failed == 0 ? 0 : 1;
