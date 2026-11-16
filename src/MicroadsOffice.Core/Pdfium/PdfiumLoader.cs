using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace MicroadsOffice.Core.Pdfium;

/// <summary>
/// The PDFium engine ships inside the single Microads Office executable as an embedded resource.
/// On first use it is written next to the executable (or to the Data folder when that is
/// read-only) and P/Invoke calls are redirected to that copy. No separate DLL has to be shipped.
/// </summary>
public static class PdfiumLoader
{
    private static string? _path;
    private static bool _registered;

    /// <param name="resourceAssembly">Assembly that embeds "pdfium.dll" (the launcher executable).</param>
    public static void Register(Assembly resourceAssembly)
    {
        if (_registered) return;
        _registered = true;
        NativeLibrary.SetDllImportResolver(typeof(PdfiumLoader).Assembly, (name, asm, search) =>
        {
            if (!name.StartsWith("pdfium", StringComparison.OrdinalIgnoreCase)) return IntPtr.Zero;
            var path = _path ??= Extract(resourceAssembly);
            if (path != null && NativeLibrary.TryLoad(path, out var h)) return h;
            return IntPtr.Zero; // default probing (pdfium.dll next to the exe in development builds)
        });
    }

    private static string? Extract(Assembly asm)
    {
        using var res = asm.GetManifestResourceStream("pdfium.dll");
        if (res == null) return null;
        byte[] data;
        using (var ms = new MemoryStream()) { res.CopyTo(ms); data = ms.ToArray(); }
        var hash = Convert.ToHexString(SHA256.HashData(data))[..16];
        var candidates = new[]
        {
            // the Store edition never writes into its (read-only) install folder
            MicroadsOffice.Core.Office.StorePackage.IsPackaged ? Path.Combine(AppSettings.DataDirectory, "bin", hash, "pdfium.dll") : Path.Combine(AppContext.BaseDirectory, "pdfium.dll"),
            Path.Combine(AppSettings.DataDirectory, "bin", hash, "pdfium.dll"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microads Office", "bin", hash, "pdfium.dll")
        };
        foreach (var c in candidates)
        {
            try
            {
                if (File.Exists(c) && SameContent(c, data)) return c;
            }
            catch { }
        }
        foreach (var c in candidates)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(c)!);
                var tmp = c + "." + Environment.ProcessId + ".tmp";
                File.WriteAllBytes(tmp, data);
                try { File.Move(tmp, c, true); }
                catch
                {
                    // the old dll is loaded by another running app: keep it if identical, otherwise use the next location
                    File.Delete(tmp);
                    if (File.Exists(c) && SameContent(c, data)) return c;
                    continue;
                }
                return c;
            }
            catch { }
        }
        return null;
    }

    private static bool SameContent(string path, byte[] data)
    {
        var fi = new FileInfo(path);
        if (fi.Length != data.Length) return false;
        return SHA256.HashData(File.ReadAllBytes(path)).AsSpan().SequenceEqual(SHA256.HashData(data));
    }
}
