using System.Diagnostics;

namespace MicroadsOffice.Tests;

public sealed class SkipException(string message) : Exception(message);
public sealed class AssertException(string message) : Exception(message);

public static class T
{
    public static void True(bool cond, string what) { if (!cond) throw new AssertException("Expected true: " + what); }
    public static void Eq<TV>(TV expected, TV actual, string what = "")
    {
        if (!EqualityComparer<TV>.Default.Equals(expected, actual))
            throw new AssertException($"Expected <{expected}> but was <{actual}> {what}");
    }
    public static void Fail(string msg) => throw new AssertException(msg);
}

public sealed class TestRunner
{
    public int Passed, Failed, Skipped;
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
        catch (SkipException ex)
        {
            Skipped++;
            Log($"SKIP  {name}: {ex.Message}");
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
