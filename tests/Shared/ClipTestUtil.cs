using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

namespace MicroadsOffice.TestShared;

/// <summary>
/// Puts data on the Windows clipboard / into drag data exactly the way other programs do (for paste and drop tests),
/// and saves / restores the user's clipboard around the tests.
/// </summary>
public static class ClipTestUtil
{
    [DllImport("user32.dll", SetLastError = true)] static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")] static extern bool CloseClipboard();
    [DllImport("user32.dll")] static extern bool EmptyClipboard();
    [DllImport("user32.dll")] static extern IntPtr SetClipboardData(uint format, IntPtr mem);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr h);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr h);
    const uint CF_BITMAP = 2, CF_DIB = 8, GMEM_MOVEABLE = 2;

    public static Bitmap TestImage(int w = 160, int h = 90, Color? c = null)
    {
        var b = new Bitmap(w, h);
        using var g = Graphics.FromImage(b);
        g.Clear(c ?? Color.SteelBlue);
        g.FillEllipse(Brushes.Gold, 10, 10, h - 20, h - 20);
        return b;
    }

    public static byte[] Png(Image img) { using var ms = new MemoryStream(); img.Save(ms, ImageFormat.Png); return ms.ToArray(); }

    static void Native(Action put)
    {
        for (int i = 0; i < 20 && !OpenClipboard(IntPtr.Zero); i++) Thread.Sleep(50);
        try { EmptyClipboard(); put(); }
        finally { CloseClipboard(); }
    }

    /// <summary>Like PrtSc / most screenshot tools: only a device-independent bitmap (Windows adds BITMAP / DIBV5 itself).</summary>
    public static void SetNativeDib(Bitmap bmp)
    {
        // BITMAPINFOHEADER + 32-bit bottom-up pixels
        int w = bmp.Width, h = bmp.Height, stride = w * 4;
        var dib = new byte[40 + stride * h];
        BitConverter.GetBytes(40).CopyTo(dib, 0);
        BitConverter.GetBytes(w).CopyTo(dib, 4);
        BitConverter.GetBytes(h).CopyTo(dib, 8);
        BitConverter.GetBytes((short)1).CopyTo(dib, 12);
        BitConverter.GetBytes((short)32).CopyTo(dib, 14);
        var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            for (int y = 0; y < h; y++) Marshal.Copy(data.Scan0 + (h - 1 - y) * data.Stride, dib, 40 + y * stride, stride);
        }
        finally { bmp.UnlockBits(data); }
        Native(() =>
        {
            var mem = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)dib.Length);
            var p = GlobalLock(mem); Marshal.Copy(dib, 0, p, dib.Length); GlobalUnlock(mem);
            SetClipboardData(CF_DIB, mem);
        });
    }

    /// <summary>Like old programs: only a GDI bitmap handle.</summary>
    public static void SetNativeBitmap(Bitmap bmp) => Native(() => SetClipboardData(CF_BITMAP, bmp.GetHbitmap()));

    /// <summary>Like the Snipping Tool / browsers: only the registered "PNG" format.</summary>
    public static void SetPngOnly(Image img)
    {
        var d = new DataObject();
        d.SetData("PNG", false, new MemoryStream(Png(img)));
        Clipboard.SetDataObject(d, true, 10, 100);
    }

    public static void SetFiles(params string[] paths)
    {
        var d = new DataObject();
        d.SetData(DataFormats.FileDrop, paths);
        Clipboard.SetDataObject(d, true, 10, 100);
    }

    /// <summary>What a browser / mail / chat app offers when an image is dragged out of it: a virtual file.</summary>
    public static DataObject VirtualFile(string name, byte[] bytes)
    {
        var desc = new byte[4 + 592];
        BitConverter.GetBytes(1).CopyTo(desc, 0);
        BitConverter.GetBytes(0x4000 | 0x40).CopyTo(desc, 4); // FD_UNICODE | FD_FILESIZE
        BitConverter.GetBytes(bytes.Length).CopyTo(desc, 4 + 68);
        Encoding.Unicode.GetBytes(name).CopyTo(desc, 4 + 72);
        var d = new DataObject();
        d.SetData("FileGroupDescriptorW", new MemoryStream(desc));
        d.SetData("FileContents", new MemoryStream(bytes));
        return d;
    }

    /// <summary>Copied web content whose picture is a data: address.</summary>
    public static DataObject HtmlWithImage(byte[] png)
    {
        var html = "<html><body><!--StartFragment--><img src=\"data:image/png;base64," + Convert.ToBase64String(png) + "\"><!--EndFragment--></body></html>";
        var d = new DataObject();
        d.SetData(DataFormats.Html, html);
        return d;
    }

    // ---------------------------------------------------------------- keep the user's clipboard

    public sealed class Snapshot
    {
        internal string? Text; internal Image? Image; internal System.Collections.Specialized.StringCollection? Files;
    }

    public static Snapshot Save()
    {
        var s = new Snapshot();
        try
        {
            if (Clipboard.ContainsFileDropList()) s.Files = Clipboard.GetFileDropList();
            else if (Clipboard.ContainsImage()) s.Image = Clipboard.GetImage();
            if (Clipboard.ContainsText()) s.Text = Clipboard.GetText();
        }
        catch { }
        return s;
    }

    public static void Restore(Snapshot s)
    {
        try
        {
            var d = new DataObject();
            bool any = false;
            if (s.Files is { Count: > 0 }) { d.SetFileDropList(s.Files); any = true; }
            if (s.Image != null) { d.SetImage(s.Image); any = true; }
            if (!string.IsNullOrEmpty(s.Text)) { d.SetText(s.Text); any = true; }
            if (any) Clipboard.SetDataObject(d, true, 10, 100); else Clipboard.Clear();
        }
        catch { }
    }
}
