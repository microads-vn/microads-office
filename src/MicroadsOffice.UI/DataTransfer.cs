using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace MicroadsOffice.UI;

/// <summary>
/// Reads what the user copied or dragged from any Windows program, in the forms Windows programs really use:
/// a screenshot (PrtSc / Snipping Tool: PNG, DIB, bitmap), files and folders copied in Explorer (file list),
/// an image dragged out of a browser / chat app / mail (virtual file: FileGroupDescriptorW + FileContents),
/// an &lt;img&gt; in copied HTML (data: or file: address — never downloaded from the Internet), and text.
/// Shared by Word, Excel and PDF so Ctrl+V and drag &amp; drop behave the same everywhere.
/// </summary>
public static class DataTransfer
{
    /// <summary>An image ready to insert: PNG, JPEG or GIF bytes (anything else is converted to PNG).</summary>
    public sealed record ImageData(byte[] Bytes, string Ext, string? Name = null)
    {
        public Image Load() => System.Drawing.Image.FromStream(new MemoryStream(Bytes));
    }

    static readonly string[] ImageExts = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".ico", ".jfif"];

    public static bool IsImageFile(string path) => ImageExts.Contains(Path.GetExtension(path).ToLowerInvariant());

    /// <summary>The clipboard, retried while another program holds it open (chat apps and clipboard managers do).</summary>
    public static IDataObject? ReadClipboard()
    {
        for (int i = 0; i < 10; i++)
        {
            try { return Clipboard.GetDataObject(); }
            catch (ExternalException) { Thread.Sleep(80); }
            catch { return null; }
        }
        return null;
    }

    /// <summary>Writes data to the clipboard, retried while another program holds it.</summary>
    public static bool WriteClipboard(DataObject data)
    {
        try { Clipboard.SetDataObject(data, true, 10, 100); return true; } catch { return false; }
    }

    static bool Has(IDataObject d, string format)
    {
        try { return d.GetDataPresent(format); } catch { return false; }
    }

    static object? Get(IDataObject d, string format)
    {
        try { return d.GetData(format); } catch { return null; }
    }

    /// <summary>Files and folders copied or dragged from Explorer (only paths that exist).</summary>
    public static string[] Files(IDataObject? d)
    {
        if (d == null || !Has(d, DataFormats.FileDrop)) return [];
        return Get(d, DataFormats.FileDrop) is string[] f ? f.Where(p => File.Exists(p) || Directory.Exists(p)).ToArray() : [];
    }

    /// <summary>Plain text (Unicode first).</summary>
    public static string? Text(IDataObject? d)
    {
        if (d == null) return null;
        if (Has(d, DataFormats.UnicodeText) && Get(d, DataFormats.UnicodeText) is string u) return u;
        if (Has(d, DataFormats.Text) && Get(d, DataFormats.Text) is string t) return t;
        return null;
    }

    /// <summary>Anything this class can turn into document content (used to accept a drag).</summary>
    public static bool HasContent(IDataObject? d) =>
        d != null && (Has(d, DataFormats.FileDrop) || Has(d, "PNG") || Has(d, DataFormats.Bitmap) || Has(d, DataFormats.Dib)
                      || Has(d, "FileGroupDescriptorW") || Has(d, DataFormats.Html) || Has(d, DataFormats.Rtf)
                      || Has(d, DataFormats.UnicodeText) || Has(d, DataFormats.Text));

    /// <summary>Readable images among <paramref name="paths"/> (in order).</summary>
    public static List<ImageData> ImageFiles(IEnumerable<string> paths)
    {
        var list = new List<ImageData>();
        foreach (var p in paths)
        {
            if (!IsImageFile(p) || !File.Exists(p)) continue;
            try { if (Normalize(File.ReadAllBytes(p), Path.GetExtension(p), Path.GetFileName(p)) is { } img) list.Add(img); } catch { }
        }
        return list;
    }

    /// <summary>
    /// The picture in <paramref name="d"/> that is not a file list: screenshot / copied image (PNG, bitmap, DIB),
    /// an image dragged from a browser or chat app (virtual file), or the first data:/file: &lt;img&gt; of copied HTML.
    /// </summary>
    public static ImageData? Picture(IDataObject? d)
    {
        if (d == null) return null;
        // PNG keeps transparency (Snipping Tool, browsers, Office)
        foreach (var f in new[] { "PNG", "image/png" })
            if (Has(d, f) && Bytes(Get(d, f)) is { Length: > 0 } png && Normalize(png, ".png") is { } a) return a;
        // virtual file dragged from a browser / mail / chat app
        if (VirtualFile(d) is { } vf && IsImageFile(vf.Name) && Normalize(vf.Bytes, Path.GetExtension(vf.Name), vf.Name) is { } v) return v;
        // bitmap (PrtSc puts CF_BITMAP / CF_DIB; Windows converts between them)
        if (Has(d, DataFormats.Bitmap) && Get(d, DataFormats.Bitmap) is Image bmp)
        {
            try { return FromImage(bmp); } catch { }
        }
        if (Has(d, DataFormats.Dib) && Bytes(Get(d, DataFormats.Dib)) is { Length: > 40 } dib && FromDib(dib) is { } b) return b;
        // <img src="data:..."> or a local file in copied HTML
        if (Has(d, DataFormats.Html) && Get(d, DataFormats.Html) is string html && FromHtml(html) is { } h) return h;
        return null;
    }

    /// <summary>Puts a picture on the clipboard data the way other programs read it: bitmap and PNG (keeps transparency).</summary>
    public static void PutImage(DataObject data, Image img)
    {
        try { data.SetImage(new Bitmap(img)); } catch { }
        try
        {
            var ms = new MemoryStream();
            img.Save(ms, ImageFormat.Png);
            ms.Position = 0;
            data.SetData("PNG", false, ms);
        }
        catch { }
    }

    // ------------------------------------------------------------------ helpers

    static byte[]? Bytes(object? o) => o switch
    {
        MemoryStream ms => ms.ToArray(),
        Stream s => ReadAll(s),
        byte[] b => b,
        _ => null,
    };

    static byte[] ReadAll(Stream s) { using var ms = new MemoryStream(); s.CopyTo(ms); return ms.ToArray(); }

    /// <summary>Checks the bytes are an image and keeps PNG / JPEG / GIF as they are; other formats become PNG.</summary>
    public static ImageData? Normalize(byte[] bytes, string ext, string? name = null)
    {
        try
        {
            using var img = System.Drawing.Image.FromStream(new MemoryStream(bytes));
            if (img.Width < 1 || img.Height < 1) return null;
            var e = ext.TrimStart('.').ToLowerInvariant();
            if (img.RawFormat.Guid == ImageFormat.Png.Guid) return new ImageData(bytes, "png", name);
            if (img.RawFormat.Guid == ImageFormat.Jpeg.Guid) return new ImageData(bytes, "jpg", name);
            if (img.RawFormat.Guid == ImageFormat.Gif.Guid) return new ImageData(bytes, "gif", name);
            _ = e;
            return FromImage(img, name);
        }
        catch { return null; }
    }

    static ImageData FromImage(Image img, string? name = null)
    {
        using var ms = new MemoryStream();
        // a copy drops the source's palette quirks / locked stream
        using (var copy = new Bitmap(img)) copy.Save(ms, ImageFormat.Png);
        return new ImageData(ms.ToArray(), "png", name);
    }

    /// <summary>CF_DIB / CF_DIBV5 bytes → a BMP file GDI+ can read.</summary>
    static ImageData? FromDib(byte[] dib)
    {
        try
        {
            int headerSize = BitConverter.ToInt32(dib, 0);
            int bitCount = BitConverter.ToInt16(dib, 14);
            int compression = BitConverter.ToInt32(dib, 16);
            int clrUsed = headerSize >= 36 ? BitConverter.ToInt32(dib, 32) : 0;
            int palette = clrUsed > 0 ? clrUsed * 4 : bitCount <= 8 ? (1 << bitCount) * 4 : 0;
            int masks = headerSize == 40 && compression == 3 ? 12 : headerSize == 40 && compression == 6 ? 16 : 0;
            int offset = 14 + headerSize + masks + palette;
            var file = new byte[14 + dib.Length];
            file[0] = (byte)'B'; file[1] = (byte)'M';
            BitConverter.GetBytes(file.Length).CopyTo(file, 2);
            BitConverter.GetBytes(offset).CopyTo(file, 10);
            dib.CopyTo(file, 14);
            return Normalize(file, ".bmp");
        }
        catch { return null; }
    }

    static readonly Regex ImgSrc = new(@"<img\b[^>]*?\bsrc\s*=\s*[""']([^""']+)[""']", RegexOptions.IgnoreCase);

    static ImageData? FromHtml(string html)
    {
        var m = ImgSrc.Match(html);
        if (!m.Success) return null;
        var src = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value.Trim());
        try
        {
            if (src.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase))
            {
                int comma = src.IndexOf(','), semi = src.IndexOf(';');
                if (comma < 0 || semi < 0 || !src[semi..comma].Contains("base64", StringComparison.OrdinalIgnoreCase)) return null;
                var ext = "." + src[11..semi];
                return Normalize(Convert.FromBase64String(src[(comma + 1)..]), ext);
            }
            if (Uri.TryCreate(src, UriKind.Absolute, out var u) && u.IsFile && File.Exists(u.LocalPath))
                return Normalize(File.ReadAllBytes(u.LocalPath), Path.GetExtension(u.LocalPath), Path.GetFileName(u.LocalPath));
        }
        catch { }
        return null; // web addresses are never downloaded
    }

    // ------------------------------------------------------------------ virtual files (FileGroupDescriptorW + FileContents)

    /// <summary>The first file of a virtual file list: what browsers, mail and chat apps offer when an image is dragged out.</summary>
    public static (string Name, byte[] Bytes)? VirtualFile(IDataObject d)
    {
        if (!Has(d, "FileGroupDescriptorW")) return null;
        try
        {
            var desc = Bytes(Get(d, "FileGroupDescriptorW"));
            if (desc == null || desc.Length < 4 + 592 || BitConverter.ToInt32(desc, 0) < 1) return null;
            // FILEDESCRIPTORW: 72 bytes of fields, then cFileName[260] (UTF-16)
            var name = Encoding.Unicode.GetString(desc, 4 + 72, 520);
            int z = name.IndexOf('\0'); if (z >= 0) name = name[..z];
            name = Path.GetFileName(name);
            if (name.Length == 0) return null;
            byte[]? bytes = null;
            try { bytes = FileContents(d, 0); } catch { }
            // some sources only answer the plain request (no item index)
            if (bytes is not { Length: > 0 }) bytes = Bytes(Get(d, "FileContents"));
            return bytes is { Length: > 0 } ? (name, bytes) : null;
        }
        catch { return null; }
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern ushort RegisterClipboardFormat(string format);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr h);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr h);
    [DllImport("kernel32.dll")] static extern UIntPtr GlobalSize(IntPtr h);
    [DllImport("ole32.dll")] static extern void ReleaseStgMedium(ref ComTypes.STGMEDIUM medium);

    /// <summary>FileContents needs an item index (lindex), which the managed GetData cannot pass: ask the COM data object directly.</summary>
    static byte[]? FileContents(IDataObject d, int index)
    {
        if (d is not ComTypes.IDataObject com) return Bytes(Get(d, "FileContents"));
        var fmt = new ComTypes.FORMATETC
        {
            cfFormat = unchecked((short)RegisterClipboardFormat("FileContents")),
            dwAspect = ComTypes.DVASPECT.DVASPECT_CONTENT, lindex = index, ptd = IntPtr.Zero,
            tymed = ComTypes.TYMED.TYMED_HGLOBAL | ComTypes.TYMED.TYMED_ISTREAM,
        };
        com.GetData(ref fmt, out var medium);
        try
        {
            if (medium.tymed == ComTypes.TYMED.TYMED_HGLOBAL)
            {
                var p = GlobalLock(medium.unionmember);
                try
                {
                    var size = (int)(ulong)GlobalSize(medium.unionmember);
                    var buf = new byte[size];
                    Marshal.Copy(p, buf, 0, size);
                    return buf;
                }
                finally { GlobalUnlock(medium.unionmember); }
            }
            if (medium.tymed == ComTypes.TYMED.TYMED_ISTREAM && Marshal.GetObjectForIUnknown(medium.unionmember) is ComTypes.IStream stream)
            {
                stream.Stat(out var stat, 1);
                var buf = new byte[stat.cbSize];
                var read = Marshal.AllocHGlobal(sizeof(int));
                try
                {
                    int total = 0;
                    while (total < buf.Length)
                    {
                        var chunk = new byte[Math.Min(1 << 20, buf.Length - total)];
                        stream.Read(chunk, chunk.Length, read);
                        int n = Marshal.ReadInt32(read);
                        if (n <= 0) break;
                        Array.Copy(chunk, 0, buf, total, n);
                        total += n;
                    }
                    return total == buf.Length ? buf : buf[..total];
                }
                finally { Marshal.FreeHGlobal(read); }
            }
            return null;
        }
        finally { ReleaseStgMedium(ref medium); }
    }
}
