using System.IO.Pipes;
using System.Security.Cryptography;
using System.Text;

namespace MicroadsOffice;

/// <summary>
/// One running window per data folder: a second launch forwards its file arguments to the
/// first instance through a named pipe (in memory, nothing is written to disk) and exits.
/// </summary>
internal static class SingleInstance
{
    public static string KeyFor(string dataDirectory)
    {
        var h = SHA256.HashData(Encoding.UTF8.GetBytes(dataDirectory.ToUpperInvariant()));
        return "MicroadsOffice_" + Convert.ToHexString(h, 0, 8);
    }

    public static bool TrySend(string key, IEnumerable<string> files)
    {
        try
        {
            using var client = new NamedPipeClientStream(".", key, PipeDirection.Out);
            client.Connect(2000);
            using var w = new StreamWriter(client, Encoding.UTF8);
            w.WriteLine("ACTIVATE");
            foreach (var f in files) w.WriteLine(f);
            w.Flush();
            return true;
        }
        catch { return false; }
    }

    public static void Listen(string key, Action<List<string>> onMessage)
    {
        var t = new Thread(() =>
        {
            while (true)
            {
                try
                {
                    using var server = new NamedPipeServerStream(key, PipeDirection.In, 1, PipeTransmissionMode.Byte, PipeOptions.None);
                    server.WaitForConnection();
                    using var r = new StreamReader(server, Encoding.UTF8);
                    var lines = new List<string>();
                    string? line;
                    while ((line = r.ReadLine()) != null) if (line != "ACTIVATE" && line.Length > 0) lines.Add(line);
                    onMessage(lines);
                }
                catch (IOException) { Thread.Sleep(200); }
                catch (ObjectDisposedException) { return; }
            }
        }) { IsBackground = true, Name = "SingleInstancePipe" };
        t.Start();
    }
}
