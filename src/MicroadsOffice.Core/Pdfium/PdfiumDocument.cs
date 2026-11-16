using MicroadsOffice.Core.Office;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using static MicroadsOffice.Core.Pdfium.PdfiumNative;

namespace MicroadsOffice.Core.Pdfium;

public sealed class PdfPasswordRequiredException : Exception
{
    public bool WrongPassword { get; }
    public PdfPasswordRequiredException(bool wrong) : base(wrong ? Loc.T("Mật khẩu không đúng.", "Incorrect password.") : Loc.T("Tài liệu được bảo vệ bằng mật khẩu.", "The document is password-protected.")) => WrongPassword = wrong;
}

public sealed class PdfLoadException(string message) : Exception(message);

public sealed record PdfBookmark(string Title, int PageIndex, float? Y, IReadOnlyList<PdfBookmark> Children);

public sealed record PdfLinkTarget(int PageIndex, float? Y, string? Uri);

public sealed record PdfSearchHit(int PageIndex, int CharStart, int CharCount, IReadOnlyList<RectangleF> Rects, string Context);

/// <summary>
/// Affine mapping from PDF page space (points, origin bottom-left) to normalised device space
/// ([0..1] x [0..1] of the rendered page box, origin top-left) for a given view rotation.
/// </summary>
public readonly record struct PageTransform(double A, double B, double C, double D, double E, double F)
{
    public PointF ToDevice(double px, double py, float w, float h)
        => new((float)((A * px + C * py + E) * w), (float)((B * px + D * py + F) * h));

    public (double X, double Y) ToPage(double dx, double dy, float w, float h)
    {
        double nx = dx / w, ny = dy / h;
        double det = A * D - B * C;
        double x = (D * (nx - E) - C * (ny - F)) / det;
        double y = (-B * (nx - E) + A * (ny - F)) / det;
        return (x, y);
    }

    public RectangleF ToDevice(RectangleF pageRect, float w, float h)
    {
        var p1 = ToDevice(pageRect.Left, pageRect.Top, w, h);
        var p2 = ToDevice(pageRect.Right, pageRect.Bottom, w, h);
        return RectangleF.FromLTRB(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), Math.Max(p1.X, p2.X), Math.Max(p1.Y, p2.Y));
    }
}

/// <summary>Thread-safe (globally serialised) wrapper around a PDFium document.</summary>
public sealed class PdfiumDocument : IDisposable
{
    /// <summary>PDFium is not thread-safe: every native call is serialised through this lock.</summary>
    public static readonly object Sync = new();

    private static bool _initialised;
    private readonly byte[] _data;
    private GCHandle _pin;
    private IntPtr _doc;
    private IntPtr _form;
    private IntPtr _formInfo;
    private readonly Dictionary<int, IntPtr> _pages = new();
    private readonly Dictionary<int, IntPtr> _textPages = new();
    private readonly LinkedList<int> _pageLru = new();
    private const int MaxCachedPages = 24;
    private readonly SizeF[] _sizes;
    private readonly Dictionary<(int, int), PageTransform> _transforms = new();

    public int PageCount { get; }
    public byte[] Data => _data;
    public bool IsEncrypted { get; }

    private static void EnsureInit()
    {
        if (_initialised) return;
        var cfg = new FPDF_LIBRARY_CONFIG { version = 2 };
        FPDF_InitLibraryWithConfig(ref cfg);
        _initialised = true;
    }

    private PdfiumDocument(byte[] data, string? password)
    {
        _data = data;
        lock (Sync)
        {
            EnsureInit();
            _pin = GCHandle.Alloc(_data, GCHandleType.Pinned);
            _doc = FPDF_LoadMemDocument64(_pin.AddrOfPinnedObject(), (nuint)_data.Length, password);
            if (_doc == IntPtr.Zero)
            {
                var err = FPDF_GetLastError();
                _pin.Free();
                if (err == FPDF_ERR_PASSWORD) throw new PdfPasswordRequiredException(password != null);
                throw new PdfLoadException(err switch
                {
                    FPDF_ERR_FORMAT => Loc.T("Tệp không phải PDF hợp lệ hoặc đã bị hỏng.", "The file is not a valid PDF or is corrupted."),
                    FPDF_ERR_SECURITY => Loc.T("Không hỗ trợ cơ chế bảo mật của tài liệu này.", "This document's security handler is not supported."),
                    FPDF_ERR_FILE => Loc.T("Không thể đọc tệp.", "Could not read the file."),
                    _ => Loc.T("Không thể mở tài liệu (mã lỗi ", "Could not open the document (error code ") + err + ")."
                });
            }
            PageCount = FPDF_GetPageCount(_doc);
            IsEncrypted = FPDF_GetSecurityHandlerRevision(_doc) != -1;
            _sizes = new SizeF[PageCount];
            for (int i = 0; i < PageCount; i++)
            {
                _sizes[i] = FPDF_GetPageSizeByIndexF(_doc, i, out var s) != 0 && s.Width > 0 && s.Height > 0
                    ? new SizeF(s.Width, s.Height) : new SizeF(612, 792);
            }
            // Form-fill environment (all callbacks null) so that AcroForm fields render like in desktop readers.
            _formInfo = Marshal.AllocHGlobal(1024);
            unsafe { new Span<byte>((void*)_formInfo, 1024).Clear(); }
            Marshal.WriteInt32(_formInfo, 1);
            _form = FPDFDOC_InitFormFillEnvironment(_doc, _formInfo);
        }
    }

    public static PdfiumDocument Load(byte[] data, string? password = null) => new(data, password);

    /// <summary>Page size in points with the page's own /Rotate applied.</summary>
    public SizeF GetPageSize(int index) => _sizes[index];

    private IntPtr GetPage(int index)
    {
        if (_doc == IntPtr.Zero) throw new ObjectDisposedException(nameof(PdfiumDocument));
        if (_pages.TryGetValue(index, out var p))
        {
            _pageLru.Remove(index); _pageLru.AddFirst(index);
            return p;
        }
        p = FPDF_LoadPage(_doc, index);
        if (p == IntPtr.Zero) throw new PdfLoadException(Loc.T($"Không tải được trang {index + 1}.", $"Could not load page {index + 1}."));
        if (_form != IntPtr.Zero) FORM_OnAfterLoadPage(p, _form);
        _pages[index] = p;
        _pageLru.AddFirst(index);
        while (_pageLru.Count > MaxCachedPages)
        {
            var victim = _pageLru.Last!.Value;
            _pageLru.RemoveLast();
            ClosePage(victim);
        }
        return p;
    }

    private void ClosePage(int index)
    {
        if (_textPages.Remove(index, out var tp)) FPDFText_ClosePage(tp);
        if (_pages.Remove(index, out var p))
        {
            if (_form != IntPtr.Zero) FORM_OnBeforeClosePage(p, _form);
            FPDF_ClosePage(p);
        }
    }

    private IntPtr GetTextPage(int index)
    {
        var page = GetPage(index);
        if (_textPages.TryGetValue(index, out var tp)) return tp;
        tp = FPDFText_LoadPage(page);
        _textPages[index] = tp;
        return tp;
    }

    /// <summary>
    /// Renders a page (or a clipped region of it) into a 32bpp bitmap.
    /// fullWidth/fullHeight define the size of the whole rendered page; clip selects the region returned.
    /// </summary>
    public Bitmap Render(int index, int fullWidth, int fullHeight, int rotation, Rectangle? clip = null,
        bool annotations = true, bool printing = false, bool grayscale = false)
    {
        var region = clip ?? new Rectangle(0, 0, fullWidth, fullHeight);
        if (region.Width <= 0 || region.Height <= 0) region = new Rectangle(0, 0, 1, 1);
        var bmp = new Bitmap(region.Width, region.Height, PixelFormat.Format32bppArgb);
        var bd = bmp.LockBits(new Rectangle(0, 0, region.Width, region.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        try
        {
            lock (Sync)
            {
                var page = GetPage(index);
                var fb = FPDFBitmap_CreateEx(region.Width, region.Height, FPDFBitmap_BGRA, bd.Scan0, bd.Stride);
                try
                {
                    FPDFBitmap_FillRect(fb, 0, 0, region.Width, region.Height, 0xFFFFFFFF);
                    int flags = (annotations ? FPDF_ANNOT : 0) | (printing ? FPDF_PRINTING : 0) | (grayscale ? FPDF_GRAYSCALE : 0);
                    FPDF_RenderPageBitmap(fb, page, -region.X, -region.Y, fullWidth, fullHeight, rotation & 3, flags);
                    if (annotations && _form != IntPtr.Zero)
                        FPDF_FFLDraw(_form, fb, page, -region.X, -region.Y, fullWidth, fullHeight, rotation & 3, flags);
                }
                finally { FPDFBitmap_Destroy(fb); }
            }
        }
        catch
        {
            bmp.UnlockBits(bd);
            bmp.Dispose();
            throw;
        }
        bmp.UnlockBits(bd);
        return bmp;
    }

    public PageTransform GetTransform(int index, int rotation)
    {
        rotation &= 3;
        lock (_transforms) { if (_transforms.TryGetValue((index, rotation), out var cached)) return cached; }
        lock (Sync)
        {
            PageTransform t;
            var page = GetPage(index);
            const int S = 1_000_000;
            FPDF_PageToDevice(page, 0, 0, S, S, rotation, 0, 0, out int x0, out int y0);
            FPDF_PageToDevice(page, 0, 0, S, S, rotation, 1000, 0, out int x1, out int y1);
            FPDF_PageToDevice(page, 0, 0, S, S, rotation, 0, 1000, out int x2, out int y2);
            t = new PageTransform(
                (x1 - x0) / 1000.0 / S, (y1 - y0) / 1000.0 / S,
                (x2 - x0) / 1000.0 / S, (y2 - y0) / 1000.0 / S,
                (double)x0 / S, (double)y0 / S);
            lock (_transforms) _transforms[(index, rotation)] = t;
            return t;
        }
    }

    public int GetPageRotation(int index)
    {
        lock (Sync) return FPDFPage_GetRotation(GetPage(index));
    }

    // ---------------- Text ----------------

    public int GetCharCount(int index)
    {
        lock (Sync) return FPDFText_CountChars(GetTextPage(index));
    }

    public int GetCharIndexAt(int index, double x, double y, double tolerance)
    {
        lock (Sync) return FPDFText_GetCharIndexAtPos(GetTextPage(index), x, y, tolerance, tolerance);
    }

    public string GetText(int index, int start, int count)
    {
        if (count <= 0) return string.Empty;
        lock (Sync)
        {
            var tp = GetTextPage(index);
            var buf = new ushort[count + 1];
            int n = FPDFText_GetText(tp, start, count, buf);
            if (n <= 1) return string.Empty;
            var chars = new char[n - 1];
            for (int i = 0; i < n - 1; i++) chars[i] = (char)buf[i];
            return new string(chars);
        }
    }

    public string GetPageText(int index)
    {
        int n = GetCharCount(index);
        return GetText(index, 0, n);
    }

    public List<RectangleF> GetTextRects(int index, int start, int count)
    {
        var list = new List<RectangleF>();
        if (count <= 0) return list;
        lock (Sync)
        {
            var tp = GetTextPage(index);
            int n = FPDFText_CountRects(tp, start, count);
            for (int i = 0; i < n; i++)
            {
                if (FPDFText_GetRect(tp, i, out double l, out double t, out double r, out double b) != 0)
                    list.Add(RectangleF.FromLTRB((float)l, (float)b, (float)r, (float)t)); // page coords: Top=bottom y
            }
        }
        return list;
    }

    public List<PdfSearchHit> Search(int index, string query, bool matchCase, bool wholeWord)
    {
        var hits = new List<PdfSearchHit>();
        if (string.IsNullOrEmpty(query)) return hits;
        lock (Sync)
        {
            var tp = GetTextPage(index);
            uint flags = (matchCase ? (uint)FPDF_MATCHCASE : 0) | (wholeWord ? (uint)FPDF_MATCHWHOLEWORD : 0);
            var h = FPDFText_FindStart(tp, query, flags, 0);
            if (h == IntPtr.Zero) return hits;
            try
            {
                int total = FPDFText_CountChars(tp);
                while (FPDFText_FindNext(h) != 0 && hits.Count < 10000)
                {
                    int s = FPDFText_GetSchResultIndex(h), c = FPDFText_GetSchCount(h);
                    var rects = new List<RectangleF>();
                    int n = FPDFText_CountRects(tp, s, c);
                    for (int i = 0; i < n; i++)
                        if (FPDFText_GetRect(tp, i, out double l, out double t, out double r, out double b) != 0)
                            rects.Add(RectangleF.FromLTRB((float)l, (float)b, (float)r, (float)t));
                    int cs = Math.Max(0, s - 30), ce = Math.Min(total, s + c + 30);
                    var buf = new ushort[ce - cs + 1];
                    int got = FPDFText_GetText(tp, cs, ce - cs, buf);
                    var sb = new StringBuilder();
                    for (int i = 0; i < Math.Max(0, got - 1); i++) sb.Append(buf[i] is '\r' or '\n' ? ' ' : (char)buf[i]);
                    hits.Add(new PdfSearchHit(index, s, c, rects, sb.ToString().Trim()));
                }
            }
            finally { FPDFText_FindClose(h); }
        }
        return hits;
    }

    // ---------------- Links & bookmarks ----------------

    public PdfLinkTarget? GetLinkAt(int index, double x, double y)
    {
        lock (Sync)
        {
            var page = GetPage(index);
            var link = FPDFLink_GetLinkAtPoint(page, x, y);
            if (link == IntPtr.Zero) return null;
            var dest = FPDFLink_GetDest(_doc, link);
            if (dest != IntPtr.Zero) return DestToTarget(dest);
            var action = FPDFLink_GetAction(link);
            return action == IntPtr.Zero ? null : ActionToTarget(action);
        }
    }

    private PdfLinkTarget? DestToTarget(IntPtr dest)
    {
        int pi = FPDFDest_GetDestPageIndex(_doc, dest);
        if (pi < 0) return null;
        float? y = null;
        if (FPDFDest_GetLocationInPage(dest, out _, out int hasY, out _, out _, out float fy, out _) != 0 && hasY != 0) y = fy;
        return new PdfLinkTarget(pi, y, null);
    }

    private PdfLinkTarget? ActionToTarget(IntPtr action)
    {
        var type = FPDFAction_GetType(action);
        if (type == PDFACTION_GOTO)
        {
            var d = FPDFAction_GetDest(_doc, action);
            return d == IntPtr.Zero ? null : DestToTarget(d);
        }
        if (type == PDFACTION_URI)
        {
            uint len = FPDFAction_GetURIPath(_doc, action, null, 0);
            if (len == 0) return null;
            var buf = new byte[len];
            FPDFAction_GetURIPath(_doc, action, buf, len);
            return new PdfLinkTarget(-1, null, Encoding.UTF8.GetString(buf, 0, (int)len - 1));
        }
        return null;
    }

    public IReadOnlyList<PdfBookmark> GetBookmarks()
    {
        lock (Sync)
        {
            int budget = 20000;
            return ReadBookmarks(IntPtr.Zero, 0, ref budget);
        }
    }

    private List<PdfBookmark> ReadBookmarks(IntPtr parent, int depth, ref int budget)
    {
        var list = new List<PdfBookmark>();
        if (depth > 32) return list;
        var seen = new HashSet<IntPtr>();
        var bm = FPDFBookmark_GetFirstChild(_doc, parent);
        while (bm != IntPtr.Zero && budget-- > 0 && seen.Add(bm))
        {
            uint len = FPDFBookmark_GetTitle(bm, null, 0);
            string title = "";
            if (len > 2)
            {
                var buf = new byte[len];
                FPDFBookmark_GetTitle(bm, buf, len);
                title = Encoding.Unicode.GetString(buf, 0, (int)len - 2);
            }
            int page = -1; float? y = null;
            var dest = FPDFBookmark_GetDest(_doc, bm);
            PdfLinkTarget? t = null;
            if (dest != IntPtr.Zero) t = DestToTarget(dest);
            else
            {
                var act = FPDFBookmark_GetAction(bm);
                if (act != IntPtr.Zero) t = ActionToTarget(act);
            }
            if (t != null) { page = t.PageIndex; y = t.Y; }
            var children = ReadBookmarks(bm, depth + 1, ref budget);
            list.Add(new PdfBookmark(title, page, y, children));
            bm = FPDFBookmark_GetNextSibling(_doc, bm);
        }
        return list;
    }

    // ---------------- Metadata ----------------

    public string GetMeta(string tag)
    {
        lock (Sync)
        {
            uint len = FPDF_GetMetaText(_doc, tag, null, 0);
            if (len <= 2) return "";
            var buf = new byte[len];
            FPDF_GetMetaText(_doc, tag, buf, len);
            return Encoding.Unicode.GetString(buf, 0, (int)len - 2);
        }
    }

    public string GetPageLabel(int index)
    {
        lock (Sync)
        {
            uint len = FPDF_GetPageLabel(_doc, index, null, 0);
            if (len <= 2) return "";
            var buf = new byte[len];
            FPDF_GetPageLabel(_doc, index, buf, len);
            return Encoding.Unicode.GetString(buf, 0, (int)len - 2);
        }
    }

    public string GetFileVersion()
    {
        lock (Sync) return FPDF_GetFileVersion(_doc, out int v) != 0 ? $"{v / 10}.{v % 10}" : "?";
    }

    public uint GetPermissions()
    {
        lock (Sync) return FPDF_GetDocPermissions(_doc);
    }

    public void Dispose()
    {
        lock (Sync)
        {
            if (_doc == IntPtr.Zero) return;
            foreach (var i in _pages.Keys.ToList()) ClosePage(i);
            if (_form != IntPtr.Zero) { FPDFDOC_ExitFormFillEnvironment(_form); _form = IntPtr.Zero; }
            FPDF_CloseDocument(_doc);
            _doc = IntPtr.Zero;
            if (_formInfo != IntPtr.Zero) { Marshal.FreeHGlobal(_formInfo); _formInfo = IntPtr.Zero; }
            if (_pin.IsAllocated) _pin.Free();
        }
    }
}
