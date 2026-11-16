using System.Diagnostics;

namespace MicroadsOffice.Excel.Tests;

public sealed class AssertException(string message) : Exception(message);

public static class T
{
    public static void True(bool cond, string what) { if (!cond) throw new AssertException("Expected true: " + what); }
    public static void Eq<TV>(TV expected, TV actual, string what = "")
    {
        if (!EqualityComparer<TV>.Default.Equals(expected, actual))
            throw new AssertException($"Expected <{expected}> but was <{actual}> {what}");
    }
}

public sealed class TestRunner
{
    public int Passed, Failed;
    public readonly List<string> Report = new();

    public void Test(string name, Action body)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            body();
            Passed++;
            Log($"PASS  {name} ({sw.ElapsedMilliseconds} ms)");
        }
        catch (Exception ex)
        {
            Failed++;
            Log($"FAIL  {name}: {ex.GetType().Name}: {ex.Message}\n      {ex.StackTrace?.Split('\n').FirstOrDefault()?.Trim()}");
        }
    }

    public void Log(string s)
    {
        Report.Add(s);
        Console.WriteLine(s);
    }
}
