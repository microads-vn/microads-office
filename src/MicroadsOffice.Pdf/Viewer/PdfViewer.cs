using MicroadsOffice.Core.Office;
using System.Drawing.Drawing2D;
using System.Text;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.UI;

namespace MicroadsOffice.Viewer;

public enum ViewerTool { Hand, SelectText, PlaceSignature, PlaceStamp }
public enum PageLayoutMode { Continuous, SinglePage }
public enum ZoomMode { Custom, FitWidth, FitPage }

public sealed record SignatureMark(int Index, int PageIndex, RectangleF PageRect, SignatureStatus Status, string FieldName);

/// <summary>Continuous-scroll PDF page viewer with zoom, rotation, selection, search highlights and placement overlay.</summary>
public sealed class PdfViewer : ScrollableControl
{
    private const int PageGap = 14;
    private const int PageMargin = 18;
    private const long FullRenderMaxPixels = 14_000_000;
    private const int TileSize = 1024;

    private PdfiumDocument? _doc;
    private readonly RenderWorker _worker = new();
    private readonly Dictionary<TileKey, Bitmap> _cache = new();
    private readonly LinkedList<TileKey> _lru = new();
    private long _cacheBytes;
    private const long CacheLimit = 450L * 1024 * 1024;

    private Rectangle[] _pageRects = [];
    private float _zoom = 1f;
    private int _rotation;
    private int _currentPage;
    private ZoomMode _zoomMode = ZoomMode.FitWidth;
    private PageLayoutMode _layoutMode = PageLayoutMode.Continuous;
    private ViewerTool _tool = ViewerTool.Hand;

    // interaction state
    private Point _dragStart;
    private Point _scrollStart;
    private bool _dragging;
    private bool _moved;
    private int _selStartPage = -1, _selStartChar = -1, _selEndPage = -1, _selEndChar = -1;
    private readonly Dictionary<int, List<RectangleF>> _selRectsCache = new();
    private List<PdfSearchHit> _hits = new();
    private int _currentHit = -1;
    private List<SignatureMark> _sigMarks = new();
    private List<SignatureFieldInfo> _emptyFields = new();

    // placement
    private int _placePage = -1;
    private Rectangle _placeRectDoc;       // while drawing, doc coords
    private bool _placing;
    private PendingPlacement? _pending;

    public PdfViewer()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint |
                 ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
        BackColor = Theme.ViewerBack;
        AutoScroll = true;
        TabStop = true;
        AccessibleName = Loc.T("Vùng hiển thị tài liệu", "Document view");
        AccessibleRole = AccessibleRole.Document;
        _worker.Completed += OnRendered;
    }

    // ------------------------------------------------------------------ public API

    public event EventHandler? PageChanged;
    public event EventHandler? ZoomChanged;
    public event EventHandler? ToolChanged;
    public event EventHandler? SelectionChanged;
    public event Action<string>? UriClicked;
    public event Action<int>? SignatureClicked;
    public event Action<SignatureFieldInfo>? EmptyFieldClicked;
    public event Action<int, Rectangle>? PlacementDrawn;   // page, doc-rect (pixels)
    public event Action<int, RectangleF>? PlacementConfirmed; // page, page-space rect (points)
    public event EventHandler? PlacementCancelled;

    public PdfiumDocument? Document => _doc;
    public int PageCount => _doc?.PageCount ?? 0;
    public float Zoom => _zoom;
    public ZoomMode ZoomMode => _zoomMode;
    public int ViewRotation => _rotation;
    public PageLayoutMode LayoutMode => _layoutMode;
    public ViewerTool Tool => _tool;
    public int CurrentPage => _currentPage;
    public bool HasSelection => _selStartPage >= 0 && _selEndPage >= 0 && (_selStartPage != _selEndPage || _selStartChar != _selEndChar || _selectAll);
    public bool HasPending => _pending != null;
    private bool _selectAll;

    private float RenderScale => _zoom * DeviceDpi / 72f;

    public void SetDocument(PdfiumDocument? doc, bool keepView = false)
    {
        var oldPage = _currentPage;
        var oldScroll = AutoScrollPosition;
        CancelPending();
        _doc = doc;
        _worker.SetDocument(doc);
        ClearCache();
        _hits.Clear(); _currentHit = -1;
        ClearSelection();
        _sigMarks.Clear(); _emptyFields.Clear();
        if (!keepView) { _currentPage = 0; }
        else _currentPage = Math.Min(oldPage, Math.Max(0, PageCount - 1));
        RecalcLayout();
        if (keepView) AutoScrollPosition = new Point(-oldScroll.X, -oldScroll.Y);
        else AutoScrollPosition = Point.Empty;
        Invalidate();
        PageChanged?.Invoke(this, EventArgs.Empty);
        ZoomChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SetTool(ViewerTool tool)
    {
        if (_tool == tool) return;
        _tool = tool;
        _placing = false;
        Cursor = tool switch
        {
            ViewerTool.Hand => Cursors.Hand,
            ViewerTool.SelectText => Cursors.IBeam,
            _ => Cursors.Cross
        };
        Invalidate();
        ToolChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SetZoom(float zoom, Point? anchorClient = null)
    {
        zoom = Math.Clamp(zoom, 0.1f, 16f);
        _zoomMode = ZoomMode.Custom;
        ApplyZoom(zoom, anchorClient);
    }

    public void SetZoomMode(ZoomMode mode)
    {
        _zoomMode = mode;
        if (mode != ZoomMode.Custom) ApplyZoom(ComputeFitZoom(), null);
    }

    public void ZoomStep(bool zoomIn, Point? anchor = null)
    {
        float[] steps = [0.1f, 0.25f, 0.33f, 0.5f, 0.66f, 0.75f, 0.9f, 1f, 1.1f, 1.25f, 1.5f, 1.75f, 2f, 2.5f, 3f, 4f, 6f, 8f, 12f, 16f];
        float z = _zoom;
        float next = zoomIn ? steps.FirstOrDefault(s => s > z + 0.001f, 16f) : steps.LastOrDefault(s => s < z - 0.001f, 0.1f);
        SetZoom(next, anchor);
    }

    public void SetLayoutMode(PageLayoutMode mode)
    {
        _layoutMode = mode;
        int page = _currentPage;
        RecalcLayout();
        GoToPage(page);
    }

    public void Rotate(bool clockwise)
    {
        int page = _currentPage;
        _rotation = (_rotation + (clockwise ? 1 : 3)) % 4;
        _worker.Reset();
        ClearCache();
        _selRectsCache.Clear();
        if (_zoomMode != ZoomMode.Custom) _zoom = ComputeFitZoom();
        RecalcLayout();
        GoToPage(page);
        ZoomChanged?.Invoke(this, EventArgs.Empty);
    }

    public void GoToPage(int index, float? pageY = null)
    {
        if (_doc == null || PageCount == 0) return;
        index = Math.Clamp(index, 0, PageCount - 1);
        if (_layoutMode == PageLayoutMode.SinglePage && index != _currentPage)
        {
            _currentPage = index;
            RecalcLayout();
        }
        var r = _pageRects[index];
        int y = r.Top - PageGap / 2;
        if (pageY.HasValue && _rotation % 2 == 0)
        {
            var t = _doc.GetTransform(index, _rotation);
            var p = t.ToDevice(0, pageY.Value, r.Width, r.Height);
            y = r.Top + (int)p.Y - 10;
        }
        AutoScrollPosition = new Point(-AutoScrollPosition.X, Math.Max(0, y));
        SetCurrentPage(index);
        Invalidate();
    }

    public void NextPage() => GoToPage(_currentPage + 1);
    public void PrevPage() => GoToPage(_currentPage - 1);

    public void SetSearchHits(List<PdfSearchHit> hits, int current)
    {
        _hits = hits;
        _currentHit = current;
        if (current >= 0 && current < hits.Count)
        {
            var h = hits[current];
            if (h.Rects.Count > 0) EnsureVisible(h.PageIndex, h.Rects[0]);
            else GoToPage(h.PageIndex);
        }
        Invalidate();
    }

    public void SetSignatureMarks(List<SignatureMark> marks, List<SignatureFieldInfo> emptyFields)
    {
        _sigMarks = marks;
        _emptyFields = emptyFields;
        Invalidate();
    }

    public void EnsureVisible(int page, RectangleF pageRect)
    {
        if (_doc == null) return;
        if (_layoutMode == PageLayoutMode.SinglePage && page != _currentPage) { _currentPage = page; RecalcLayout(); }
        var dr = PageToDoc(page, pageRect);
        var view = new Rectangle(-AutoScrollPosition.X, -AutoScrollPosition.Y, ClientSize.Width, ClientSize.Height);
        if (!view.Contains(Rectangle.Round(dr)))
        {
            int x = dr.Left < view.Left || dr.Right > view.Right ? (int)dr.Left - 60 : view.Left;
            int y = (int)dr.Top - ClientSize.Height / 3;
            AutoScrollPosition = new Point(Math.Max(0, x), Math.Max(0, y));
        }
        UpdateCurrentPageFromScroll();
        Invalidate();
    }

    // ------------------------------------------------------------------ layout

    private SizeF PagePoints(int i)
    {
        var s = _doc!.GetPageSize(i);
        return _rotation % 2 == 1 ? new SizeF(s.Height, s.Width) : s;
    }

    private float ComputeFitZoom()
    {
        if (_doc == null || PageCount == 0) return 1f;
        int availW = Math.Max(50, ClientSize.Width - 2 * PageMargin - SystemInformation.VerticalScrollBarWidth);
        int availH = Math.Max(50, ClientSize.Height - 2 * PageMargin);
        float dpiK = DeviceDpi / 72f;
        if (_zoomMode == ZoomMode.FitWidth)
        {
            float maxW = 0;
            int n = Math.Min(PageCount, 50);
            for (int i = 0; i < n; i++) maxW = Math.Max(maxW, PagePoints(i).Width);
            return Math.Clamp(availW / (maxW * dpiK), 0.1f, 16f);
        }
        var p = PagePoints(Math.Clamp(_currentPage, 0, PageCount - 1));
        return Math.Clamp(Math.Min(availW / (p.Width * dpiK), availH / (p.Height * dpiK)), 0.1f, 16f);
    }

    private void ApplyZoom(float zoom, Point? anchorClient)
    {
        if (_doc == null) { _zoom = zoom; ZoomChanged?.Invoke(this, EventArgs.Empty); return; }
        var anchor = anchorClient ?? new Point(ClientSize.Width / 2, ClientSize.Height / 2);
        // remember the page point under the anchor
        var docPt = ClientToDoc(anchor);
        int page = PageAtDoc(docPt, nearest: true);
        PointF rel = PointF.Empty;
        if (page >= 0)
        {
            var r = _pageRects[page];
            rel = new PointF((docPt.X - r.Left) / (float)Math.Max(1, r.Width), (docPt.Y - r.Top) / (float)Math.Max(1, r.Height));
        }
        bool changed = Math.Abs(zoom - _zoom) > 0.0001f;
        _zoom = zoom;
        if (changed) { _worker.Reset(); _selRectsCache.Clear(); }
        RecalcLayout();
        if (page >= 0)
        {
            var r = _pageRects[page];
            int nx = (int)(r.Left + rel.X * r.Width) - anchor.X;
            int ny = (int)(r.Top + rel.Y * r.Height) - anchor.Y;
            AutoScrollPosition = new Point(Math.Max(0, nx), Math.Max(0, ny));
        }
        UpdateCurrentPageFromScroll();
        Invalidate();
        ZoomChanged?.Invoke(this, EventArgs.Empty);
    }

    private void RecalcLayout()
    {
        if (_doc == null || PageCount == 0)
        {
            _pageRects = [];
            AutoScrollMinSize = Size.Empty;
            return;
        }
        int n = PageCount;
        var rects = new Rectangle[n];
        float k = RenderScale;
        int maxW = 0;
        var sizes = new Size[n];
        for (int i = 0; i < n; i++)
        {
            var p = PagePoints(i);
            sizes[i] = new Size(Math.Max(1, (int)Math.Round(p.Width * k)), Math.Max(1, (int)Math.Round(p.Height * k)));
            maxW = Math.Max(maxW, sizes[i].Width);
        }
        int totalW = maxW + 2 * PageMargin;
        int areaW = Math.Max(totalW, ClientSize.Width);
        int y = PageMargin;
        for (int i = 0; i < n; i++)
        {
            if (_layoutMode == PageLayoutMode.SinglePage && i != _currentPage)
            {
                rects[i] = new Rectangle(-100000, -100000, sizes[i].Width, sizes[i].Height);
                continue;
            }
            int x = (areaW - sizes[i].Width) / 2;
            rects[i] = new Rectangle(x, y, sizes[i].Width, sizes[i].Height);
            y += sizes[i].Height + PageGap;
        }
        _pageRects = rects;
        AutoScrollMinSize = new Size(totalW, y + PageMargin - PageGap);
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        if (_doc == null) return;
        if (_zoomMode != ZoomMode.Custom)
        {
            float z = ComputeFitZoom();
            if (Math.Abs(z - _zoom) > 0.001f) { ApplyZoom(z, null); return; }
        }
        RecalcLayout();
        Invalidate();
    }

    protected override void OnScroll(ScrollEventArgs se)
    {
        base.OnScroll(se);
        UpdateCurrentPageFromScroll();
        PositionPendingButtons();
        Invalidate();
    }

    private void UpdateCurrentPageFromScroll()
    {
        if (_doc == null || PageCount == 0 || _layoutMode == PageLayoutMode.SinglePage) return;
        int viewTop = -AutoScrollPosition.Y, viewMid = viewTop + ClientSize.Height / 3;
        int best = _currentPage;
        for (int i = 0; i < _pageRects.Length; i++)
        {
            var r = _pageRects[i];
            if (r.Top - PageGap <= viewMid && r.Bottom + PageGap >= viewMid) { best = i; break; }
            if (r.Top > viewMid) { best = Math.Max(0, i - 1); break; }
            best = i;
        }
        SetCurrentPage(best);
    }

    private void SetCurrentPage(int p)
    {
        if (p == _currentPage) return;
        _currentPage = p;
        PageChanged?.Invoke(this, EventArgs.Empty);
    }

    // ------------------------------------------------------------------ coordinates

    private Point ClientToDoc(Point c) => new(c.X - AutoScrollPosition.X, c.Y - AutoScrollPosition.Y);
    private Rectangle DocToClient(Rectangle r) => new(r.X + AutoScrollPosition.X, r.Y + AutoScrollPosition.Y, r.Width, r.Height);
    private RectangleF DocToClient(RectangleF r) => new(r.X + AutoScrollPosition.X, r.Y + AutoScrollPosition.Y, r.Width, r.Height);

    private int PageAtDoc(Point p, bool nearest = false)
    {
        for (int i = 0; i < _pageRects.Length; i++)
            if (_pageRects[i].Contains(p)) return i;
        if (!nearest || _pageRects.Length == 0) return -1;
        int best = -1; long bestD = long.MaxValue;
        for (int i = 0; i < _pageRects.Length; i++)
        {
            var r = _pageRects[i];
            if (r.X < -50000) continue;
            long dy = p.Y < r.Top ? r.Top - p.Y : p.Y > r.Bottom ? p.Y - r.Bottom : 0;
            if (dy < bestD) { bestD = dy; best = i; }
        }
        return best;
    }

    /// <summary>Page-space rect (points) → doc pixel rect.</summary>
    private RectangleF PageToDoc(int page, RectangleF pageRect)
    {
        var r = _pageRects[page];
        var t = _doc!.GetTransform(page, _rotation);
        var d = t.ToDevice(pageRect, r.Width, r.Height);
        d.Offset(r.Left, r.Top);
        return d;
    }

    private (double X, double Y) DocToPage(int page, Point doc)
    {
        var r = _pageRects[page];
        var t = _doc!.GetTransform(page, _rotation);
        return t.ToPage(doc.X - r.Left, doc.Y - r.Top, r.Width, r.Height);
    }

    /// <summary>Converts a doc-pixel rectangle on a page to a page-space rectangle in points.</summary>
    public RectangleF DocRectToPageRect(int page, Rectangle docRect)
    {
        var a = DocToPage(page, new Point(docRect.Left, docRect.Top));
        var b = DocToPage(page, new Point(docRect.Right, docRect.Bottom));
        return RectangleF.FromLTRB((float)Math.Min(a.X, b.X), (float)Math.Min(a.Y, b.Y), (float)Math.Max(a.X, b.X), (float)Math.Max(a.Y, b.Y));
    }

    // ------------------------------------------------------------------ rendering

    private void OnRendered(TileKey key, Bitmap bmp, int gen)
    {
        if (IsDisposed || !IsHandleCreated) { bmp.Dispose(); return; }
        try
        {
            BeginInvoke(() =>
            {
                if (gen != _worker.Generation) { bmp.Dispose(); return; }
                if (_cache.TryGetValue(key, out var old)) { _cacheBytes -= (long)old.Width * old.Height * 4; old.Dispose(); _lru.Remove(key); }
                _cache[key] = bmp;
                _lru.AddFirst(key);
                _cacheBytes += (long)bmp.Width * bmp.Height * 4;
                TrimCache();
                Invalidate();
            });
        }
        catch (InvalidOperationException) { bmp.Dispose(); }
    }

    private int CurrentW(int page) => page < _pageRects.Length ? _pageRects[page].Width : -1;

    private void TrimCache()
    {
        while (_cacheBytes > CacheLimit && _lru.Count > 0)
        {
            var k = _lru.Last!.Value;
            _lru.RemoveLast();
            if (_cache.Remove(k, out var b)) { _cacheBytes -= (long)b.Width * b.Height * 4; b.Dispose(); }
        }
    }

    private void ClearCache()
    {
        foreach (var b in _cache.Values) b.Dispose();
        _cache.Clear();
        _lru.Clear();
        _cacheBytes = 0;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(BackColor);
        if (_doc == null || PageCount == 0) return;
        g.InterpolationMode = InterpolationMode.HighQualityBilinear;
        var view = new Rectangle(-AutoScrollPosition.X, -AutoScrollPosition.Y, ClientSize.Width, ClientSize.Height);
        using var shadow = new SolidBrush(Color.FromArgb(40, 0, 0, 0));
        for (int i = 0; i < _pageRects.Length; i++)
        {
            var r = _pageRects[i];
            if (!r.IntersectsWith(view)) continue;
            var cr = DocToClient(r);
            g.FillRectangle(shadow, cr.X + 2, cr.Y + 3, cr.Width, cr.Height);
            g.FillRectangle(Brushes.White, cr);
            DrawPageContent(g, i, r, view);
            DrawOverlays(g, i, r);
            g.DrawRectangle(Pens.Silver, cr.X - 1, cr.Y - 1, cr.Width + 1, cr.Height + 1);
        }
        if (_placing && _placePage >= 0)
        {
            var cr = DocToClient(_placeRectDoc);
            using var pen = new Pen(Theme.Accent, 2) { DashStyle = DashStyle.Dash };
            using var fill = new SolidBrush(Color.FromArgb(40, Theme.Accent));
            g.FillRectangle(fill, cr);
            g.DrawRectangle(pen, cr);
        }
        if (_pending != null) DrawPending(g);
    }

    private void DrawPageContent(Graphics g, int page, Rectangle r, Rectangle view)
    {
        long pixels = (long)r.Width * r.Height;
        if (pixels <= FullRenderMaxPixels)
        {
            var key = new TileKey(page, r.Width, r.Height, _rotation, -1, -1);
            if (_cache.TryGetValue(key, out var bmp))
            {
                _lru.Remove(key); _lru.AddFirst(key);
                g.DrawImageUnscaled(bmp, DocToClient(r).Location);
                return;
            }
            DrawFallback(g, page, r);
            _worker.Request(key, null);
            return;
        }
        // Tiled rendering for very large zoom levels.
        DrawFallback(g, page, r);
        var vis = Rectangle.Intersect(r, view);
        int tx0 = (vis.Left - r.Left) / TileSize, tx1 = (vis.Right - r.Left - 1) / TileSize;
        int ty0 = (vis.Top - r.Top) / TileSize, ty1 = (vis.Bottom - r.Top - 1) / TileSize;
        for (int ty = ty0; ty <= ty1; ty++)
        for (int tx = tx0; tx <= tx1; tx++)
        {
            var key = new TileKey(page, r.Width, r.Height, _rotation, tx, ty);
            var clip = new Rectangle(tx * TileSize, ty * TileSize, Math.Min(TileSize, r.Width - tx * TileSize), Math.Min(TileSize, r.Height - ty * TileSize));
            if (_cache.TryGetValue(key, out var tile))
            {
                _lru.Remove(key); _lru.AddFirst(key);
                g.DrawImageUnscaled(tile, DocToClient(r).X + clip.X, DocToClient(r).Y + clip.Y);
            }
            else _worker.Request(key, clip);
        }
    }

    private void DrawFallback(Graphics g, int page, Rectangle r)
    {
        // Use any cached full-page bitmap of this page (other zoom) as a temporary preview.
        foreach (var (k, b) in _cache)
        {
            if (k.Page == page && k.TX < 0 && k.Rot == _rotation)
            {
                g.DrawImage(b, DocToClient(r));
                return;
            }
        }
    }

    private void DrawOverlays(Graphics g, int page, Rectangle r)
    {
        // search hits
        if (_hits.Count > 0)
        {
            using var hitBrush = new SolidBrush(Color.FromArgb(90, 255, 214, 0));
            using var curBrush = new SolidBrush(Color.FromArgb(120, 255, 120, 0));
            for (int i = 0; i < _hits.Count; i++)
            {
                var h = _hits[i];
                if (h.PageIndex != page) continue;
                foreach (var rect in h.Rects)
                    g.FillRectangle(i == _currentHit ? curBrush : hitBrush, DocToClient(PageToDoc(page, rect)));
            }
        }
        // text selection
        var sel = GetSelectionRects(page);
        if (sel.Count > 0)
        {
            using var selBrush = new SolidBrush(Theme.Selection);
            foreach (var rect in sel) g.FillRectangle(selBrush, DocToClient(PageToDoc(page, rect)));
        }
        // signature validity badges (like Foxit/Adobe)
        foreach (var m in _sigMarks)
        {
            if (m.PageIndex != page || m.PageRect.Width <= 0) continue;
            var dr = DocToClient(PageToDoc(page, m.PageRect));
            int s = Dpi.Scale(18);
            var color = m.Status switch
            {
                SignatureStatus.Valid => Theme.OkGreen,
                SignatureStatus.Unknown => Theme.WarnAmber,
                _ => Theme.ErrRed
            };
            var glyph = m.Status switch
            {
                SignatureStatus.Valid => Icons.Check,
                SignatureStatus.Unknown => Icons.Unknown,
                _ => Icons.Error
            };
            using var pen = new Pen(Color.FromArgb(140, color), 1) { DashStyle = DashStyle.Dot };
            g.DrawRectangle(pen, dr.X, dr.Y, dr.Width, dr.Height);
            float bx = dr.X - s - 4 >= 0 ? dr.X - s - 4 : dr.Right + 4, by = dr.Y;
            g.FillEllipse(Brushes.White, bx - 2, by - 2, s + 4, s + 4);
            g.DrawImage(Icons.Get(glyph, s, color), bx, by, s, s);
        }
        // empty signature fields: highlight when signing tool is active
        if (_tool == ViewerTool.PlaceSignature)
        {
            using var b = new SolidBrush(Color.FromArgb(50, Theme.AccentGreen));
            using var pen = new Pen(Theme.AccentGreen, 1.5f) { DashStyle = DashStyle.Dash };
            foreach (var f in _emptyFields.Where(x => x.PageIndex == page && x.Rect.Width > 0))
            {
                var dr = DocToClient(PageToDoc(page, f.Rect));
                g.FillRectangle(b, dr);
                g.DrawRectangle(pen, dr.X, dr.Y, dr.Width, dr.Height);
                using var font = new Font("Segoe UI", 8.5f);
                g.DrawString(Loc.T("Bấm để ký vào: ", "Click to sign: ") + f.FullName, font, Brushes.DarkGreen, dr.X + 2, dr.Y + 2);
            }
        }
    }

    // ------------------------------------------------------------------ text selection

    public void ClearSelection()
    {
        _selStartPage = _selStartChar = _selEndPage = _selEndChar = -1;
        _selectAll = false;
        _selRectsCache.Clear();
        Invalidate();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SelectAll()
    {
        if (_doc == null || PageCount == 0) return;
        _selectAll = true;
        _selStartPage = 0; _selStartChar = 0;
        _selEndPage = PageCount - 1; _selEndChar = int.MaxValue;
        _selRectsCache.Clear();
        Invalidate();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    private (int From, int To)? SelectionRangeOnPage(int page)
    {
        if (_selStartPage < 0 || _selEndPage < 0 || _doc == null) return null;
        (int p1, int c1, int p2, int c2) = (_selStartPage, _selStartChar, _selEndPage, _selEndChar);
        if (p1 > p2 || (p1 == p2 && c1 > c2)) (p1, c1, p2, c2) = (p2, c2, p1, c1);
        if (page < p1 || page > p2) return null;
        int count = _doc.GetCharCount(page);
        if (count == 0) return null;
        int from = page == p1 ? Math.Max(0, c1) : 0;
        int to = page == p2 ? Math.Min(count - 1, c2) : count - 1;
        return to >= from ? (from, to) : null;
    }

    private List<RectangleF> GetSelectionRects(int page)
    {
        if (_selRectsCache.TryGetValue(page, out var list)) return list;
        list = new List<RectangleF>();
        var range = SelectionRangeOnPage(page);
        if (range != null && (HasSelection))
            list = _doc!.GetTextRects(page, range.Value.From, range.Value.To - range.Value.From + 1);
        _selRectsCache[page] = list;
        return list;
    }

    public string GetSelectedText()
    {
        if (!HasSelection || _doc == null) return "";
        var sb = new StringBuilder();
        int p1 = Math.Min(_selStartPage, _selEndPage), p2 = Math.Max(_selStartPage, _selEndPage);
        for (int p = p1; p <= p2; p++)
        {
            var range = SelectionRangeOnPage(p);
            if (range == null) continue;
            if (sb.Length > 0) sb.AppendLine();
            sb.Append(_doc.GetText(p, range.Value.From, range.Value.To - range.Value.From + 1));
        }
        return sb.ToString();
    }

    public bool CopySelection()
    {
        var text = GetSelectedText();
        if (string.IsNullOrEmpty(text)) return false;
        try { Clipboard.SetText(text); return true; }
        catch { return false; }
    }

    // ------------------------------------------------------------------ mouse

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        Focus();
        if (_doc == null) return;
        _dragStart = e.Location;
        _scrollStart = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y);
        _moved = false;
        if (e.Button != MouseButtons.Left) return;

        if (_pending != null && _pending.HitTest(this, e.Location) is { } hit)
        {
            _pending.BeginDrag(hit, e.Location);
            _dragging = true;
            return;
        }
        var doc = ClientToDoc(e.Location);
        int page = PageAtDoc(doc);
        switch (_tool)
        {
            case ViewerTool.Hand:
                _dragging = true;
                Cursor = Cursors.SizeAll;
                break;
            case ViewerTool.SelectText:
                ClearSelection();
                if (page >= 0)
                {
                    var (px, py) = DocToPage(page, doc);
                    int idx = _doc.GetCharIndexAt(page, px, py, 6 / _zoom + 2);
                    if (idx >= 0) { _selStartPage = page; _selStartChar = idx; }
                }
                _dragging = true;
                break;
            case ViewerTool.PlaceSignature:
            case ViewerTool.PlaceStamp:
                if (_pending != null) return;
                if (page >= 0)
                {
                    _placing = true;
                    _placePage = page;
                    _placeRectDoc = new Rectangle(doc, Size.Empty);
                    _placeAnchorPoint = doc;
                }
                break;
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (_doc == null) return;
        if (Math.Abs(e.X - _dragStart.X) + Math.Abs(e.Y - _dragStart.Y) > 3) _moved = true;
        var doc = ClientToDoc(e.Location);

        if (_pending != null && _pending.IsDragging)
        {
            _pending.DragTo(this, e.Location);
            Invalidate();
            return;
        }
        if (_placing && e.Button == MouseButtons.Left)
        {
            var pr = _pageRects[_placePage];
            var p = new Point(Math.Clamp(doc.X, pr.Left, pr.Right), Math.Clamp(doc.Y, pr.Top, pr.Bottom));
            var s = _placeAnchorPoint;
            _placeRectDoc = Rectangle.FromLTRB(Math.Min(s.X, p.X), Math.Min(s.Y, p.Y), Math.Max(s.X, p.X), Math.Max(s.Y, p.Y));
            Invalidate();
            return;
        }
        if (_dragging && e.Button == MouseButtons.Left)
        {
            if (_tool == ViewerTool.Hand)
            {
                AutoScrollPosition = new Point(_scrollStart.X - (e.X - _dragStart.X), _scrollStart.Y - (e.Y - _dragStart.Y));
                UpdateCurrentPageFromScroll();
                PositionPendingButtons();
                Invalidate();
            }
            else if (_tool == ViewerTool.SelectText && _selStartPage >= 0)
            {
                int page = PageAtDoc(doc, nearest: true);
                if (page >= 0)
                {
                    var (px, py) = DocToPage(page, doc);
                    int idx = _doc.GetCharIndexAt(page, px, py, 30);
                    if (idx >= 0)
                    {
                        _selEndPage = page; _selEndChar = idx;
                        _selRectsCache.Clear();
                        Invalidate();
                    }
                }
                AutoScrollWhileDragging(e.Location);
            }
            return;
        }
        // hover feedback
        if (_pending != null)
        {
            var hit = _pending.HitTest(this, e.Location);
            Cursor = hit switch
            {
                PendingPlacement.Hit.Move => Cursors.SizeAll,
                PendingPlacement.Hit.N or PendingPlacement.Hit.S => Cursors.SizeNS,
                PendingPlacement.Hit.E or PendingPlacement.Hit.W => Cursors.SizeWE,
                PendingPlacement.Hit.NW or PendingPlacement.Hit.SE => Cursors.SizeNWSE,
                PendingPlacement.Hit.NE or PendingPlacement.Hit.SW => Cursors.SizeNESW,
                _ => ToolCursor()
            };
            return;
        }
        int hp = PageAtDoc(doc);
        bool overLink = false, overSig = false;
        if (hp >= 0 && _tool is ViewerTool.Hand or ViewerTool.SelectText)
        {
            var (px, py) = DocToPage(hp, doc);
            overLink = _doc.GetLinkAt(hp, px, py) != null;
            overSig = _sigMarks.Any(m => m.PageIndex == hp && m.PageRect.Contains((float)px, (float)py));
        }
        if (hp >= 0 && _tool == ViewerTool.PlaceSignature)
        {
            var (px, py) = DocToPage(hp, doc);
            overSig = _emptyFields.Any(f => f.PageIndex == hp && f.Rect.Contains((float)px, (float)py));
        }
        Cursor = overLink || overSig ? Cursors.Hand : ToolCursor();
    }

    private Point _placeAnchorPoint;

    private Cursor ToolCursor() => _tool switch
    {
        ViewerTool.Hand => Cursors.Hand,
        ViewerTool.SelectText => Cursors.IBeam,
        _ => Cursors.Cross
    };

    private void AutoScrollWhileDragging(Point client)
    {
        int dy = client.Y < 0 ? client.Y : client.Y > ClientSize.Height ? client.Y - ClientSize.Height : 0;
        if (dy != 0)
        {
            AutoScrollPosition = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y + dy);
            UpdateCurrentPageFromScroll();
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (_doc == null) return;
        var doc = ClientToDoc(e.Location);
        if (_pending != null && _pending.IsDragging)
        {
            _pending.EndDrag();
            _dragging = false;
            PositionPendingButtons();
            Invalidate();
            return;
        }
        if (_placing)
        {
            _placing = false;
            var r = _placeRectDoc;
            var pr = _pageRects[_placePage];
            if (r.Width < 12 || r.Height < 12)
            {
                // simple click: default size centred on the click point
                float k = RenderScale;
                int w = (int)((_tool == ViewerTool.PlaceStamp ? 110 : 210) * k), h = (int)((_tool == ViewerTool.PlaceStamp ? 110 : 70) * k);
                var c = new Point(Math.Clamp(doc.X, pr.Left, pr.Right), Math.Clamp(doc.Y, pr.Top, pr.Bottom));
                r = new Rectangle(c.X - w / 2, c.Y - h / 2, w, h);
                if (r.Right > pr.Right) r.X = pr.Right - r.Width;
                if (r.Bottom > pr.Bottom) r.Y = pr.Bottom - r.Height;
                if (r.Left < pr.Left) r.X = pr.Left;
                if (r.Top < pr.Top) r.Y = pr.Top;
                r.Intersect(pr);
            }
            // Clicking an existing empty signature field signs that field instead.
            if (_tool == ViewerTool.PlaceSignature && !_moved)
            {
                int page = PageAtDoc(doc);
                if (page >= 0)
                {
                    var (px, py) = DocToPage(page, doc);
                    var field = _emptyFields.FirstOrDefault(f => f.PageIndex == page && f.Rect.Contains((float)px, (float)py));
                    if (field != null) { Invalidate(); EmptyFieldClicked?.Invoke(field); return; }
                }
            }
            Invalidate();
            PlacementDrawn?.Invoke(_placePage, r);
            return;
        }
        if (_dragging)
        {
            _dragging = false;
            Cursor = ToolCursor();
            if (_tool == ViewerTool.SelectText) SelectionChanged?.Invoke(this, EventArgs.Empty);
        }
        if (!_moved && e.Button == MouseButtons.Left && _tool is ViewerTool.Hand or ViewerTool.SelectText)
        {
            int page = PageAtDoc(doc);
            if (page < 0) return;
            var (px, py) = DocToPage(page, doc);
            var sig = _sigMarks.FirstOrDefault(m => m.PageIndex == page && m.PageRect.Contains((float)px, (float)py));
            if (sig != null) { SignatureClicked?.Invoke(sig.Index); return; }
            var link = _doc.GetLinkAt(page, px, py);
            if (link != null)
            {
                if (link.Uri != null) UriClicked?.Invoke(link.Uri);
                else if (link.PageIndex >= 0) GoToPage(link.PageIndex, link.Y);
            }
        }
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if (_doc == null) { base.OnMouseWheel(e); return; }
        if ((ModifierKeys & Keys.Control) != 0)
        {
            ZoomStep(e.Delta > 0, e.Location);
            return;
        }
        if (_layoutMode == PageLayoutMode.SinglePage)
        {
            int before = -AutoScrollPosition.Y;
            base.OnMouseWheel(e);
            int after = -AutoScrollPosition.Y;
            if (before == after)
            {
                if (e.Delta < 0 && _currentPage < PageCount - 1) GoToPage(_currentPage + 1);
                else if (e.Delta > 0 && _currentPage > 0)
                {
                    GoToPage(_currentPage - 1);
                    AutoScrollPosition = new Point(-AutoScrollPosition.X, AutoScrollMinSize.Height);
                }
            }
        }
        else base.OnMouseWheel(e);
        UpdateCurrentPageFromScroll();
        PositionPendingButtons();
        Invalidate();
    }

    protected override bool IsInputKey(Keys keyData) =>
        keyData is Keys.Up or Keys.Down or Keys.Left or Keys.Right or Keys.PageUp or Keys.PageDown or Keys.Home or Keys.End || base.IsInputKey(keyData);

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (_doc == null) return;
        int line = Dpi.Scale(48);
        var pos = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y);
        switch (e.KeyCode)
        {
            case Keys.Down: pos.Y += line; break;
            case Keys.Up: pos.Y -= line; break;
            case Keys.Left: pos.X -= line; break;
            case Keys.Right: pos.X += line; break;
            case Keys.PageDown:
                if (_layoutMode == PageLayoutMode.SinglePage) { NextPage(); return; }
                pos.Y += ClientSize.Height - line; break;
            case Keys.PageUp:
                if (_layoutMode == PageLayoutMode.SinglePage) { PrevPage(); return; }
                pos.Y -= ClientSize.Height - line; break;
            case Keys.Home when e.Control || _layoutMode == PageLayoutMode.Continuous: GoToPage(0); return;
            case Keys.End when e.Control || _layoutMode == PageLayoutMode.Continuous: GoToPage(PageCount - 1); return;
            case Keys.Escape:
                if (_pending != null) { CancelPending(); return; }
                if (_tool is ViewerTool.PlaceSignature or ViewerTool.PlaceStamp) SetTool(ViewerTool.Hand);
                ClearSelection();
                return;
            case Keys.Enter when _pending != null:
                ConfirmPending();
                return;
            default: return;
        }
        AutoScrollPosition = pos;
        UpdateCurrentPageFromScroll();
        PositionPendingButtons();
        Invalidate();
        e.Handled = true;
    }

    // ------------------------------------------------------------------ pending placement (move / resize before applying)

    /// <summary>Shows a movable / resizable preview of the signature or stamp before it is applied.</summary>
    public void BeginPending(int page, Rectangle docRect, Func<SizeF, Bitmap?> preview, string confirmText, string? cancelText = null)
    {
        CancelPending();
        var pr = _pageRects[page];
        var norm = new RectangleF((docRect.X - pr.X) / (float)pr.Width, (docRect.Y - pr.Y) / (float)pr.Height,
            docRect.Width / (float)pr.Width, docRect.Height / (float)pr.Height);
        _pending = new PendingPlacement(page, norm, preview);
        var ok = new Button
        {
            Text = confirmText, Name = "btnPendingConfirm", AutoSize = true, BackColor = Theme.Accent, ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat, Font = Theme.UiFontBold, Cursor = Cursors.Hand, Padding = new Padding(6, 2, 6, 2)
        };
        ok.FlatAppearance.BorderSize = 0;
        ok.Click += (_, _) => ConfirmPending();
        var cancel = new Button
        {
            Text = cancelText ?? Loc.T("Hủy", "Cancel"), Name = "btnPendingCancel", AutoSize = true, BackColor = Color.White, FlatStyle = FlatStyle.Flat,
            Font = Theme.UiFont, Cursor = Cursors.Hand, Padding = new Padding(6, 2, 6, 2)
        };
        cancel.FlatAppearance.BorderColor = Theme.Border;
        cancel.Click += (_, _) => CancelPending();
        _pending.Buttons.AddRange([ok, cancel]);
        Controls.Add(ok);
        Controls.Add(cancel);
        PositionPendingButtons();
        Invalidate();
    }

    public void RefreshPendingPreview()
    {
        _pending?.InvalidatePreview();
        Invalidate();
    }

    public void ConfirmPending()
    {
        if (_pending == null || _doc == null) return;
        var p = _pending;
        var pr = _pageRects[p.Page];
        var dr = Rectangle.Round(new RectangleF(pr.X + p.Norm.X * pr.Width, pr.Y + p.Norm.Y * pr.Height, p.Norm.Width * pr.Width, p.Norm.Height * pr.Height));
        var pageRect = DocRectToPageRect(p.Page, dr);
        RemovePending();
        PlacementConfirmed?.Invoke(p.Page, pageRect);
    }

    public void CancelPending()
    {
        if (_pending == null) return;
        RemovePending();
        PlacementCancelled?.Invoke(this, EventArgs.Empty);
    }

    private void RemovePending()
    {
        if (_pending == null) return;
        foreach (var b in _pending.Buttons) { Controls.Remove(b); b.Dispose(); }
        _pending.Dispose();
        _pending = null;
        Invalidate();
    }

    internal Rectangle PendingClientRect()
    {
        var p = _pending!;
        var pr = _pageRects[p.Page];
        return DocToClient(Rectangle.Round(new RectangleF(pr.X + p.Norm.X * pr.Width, pr.Y + p.Norm.Y * pr.Height, p.Norm.Width * pr.Width, p.Norm.Height * pr.Height)));
    }

    internal Rectangle PendingPageClientRect() => DocToClient(_pageRects[_pending!.Page]);

    private void PositionPendingButtons()
    {
        if (_pending == null || _pending.Buttons.Count < 2) return;
        var r = PendingClientRect();
        var ok = _pending.Buttons[0]; var cancel = _pending.Buttons[1];
        int y = r.Bottom + 6;
        if (y + ok.Height > ClientSize.Height) y = r.Top - ok.Height - 6;
        int x = Math.Max(4, Math.Min(r.Left, ClientSize.Width - ok.Width - cancel.Width - 12));
        ok.Location = new Point(x, y);
        cancel.Location = new Point(x + ok.Width + 6, y);
    }

    private void DrawPending(Graphics g)
    {
        var p = _pending!;
        var r = PendingClientRect();
        float k = RenderScale;
        var upright = _rotation % 2 == 1 ? new SizeF(r.Height / k, r.Width / k) : new SizeF(r.Width / k, r.Height / k);
        var bmp = p.GetPreview(upright, _rotation);
        if (bmp != null) g.DrawImage(bmp, r);
        using var pen = new Pen(Theme.Accent, 1.5f) { DashStyle = DashStyle.Dash };
        g.DrawRectangle(pen, r);
        foreach (var h in PendingPlacement.Handles(r))
        {
            g.FillRectangle(Brushes.White, h);
            g.DrawRectangle(Pens.Firebrick, h);
        }
        PositionPendingButtons();
    }

    internal void MovePendingNorm(RectangleF norm)
    {
        if (_pending == null) return;
        _pending.Norm = norm;
        _pending.InvalidatePreview();
    }

    internal Rectangle PageClientRectOf(int page) => DocToClient(_pageRects[page]);

    // ------------------------------------------------------------------ cleanup

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _worker.Dispose();
            ClearCache();
            _pending?.Dispose();
        }
        base.Dispose(disposing);
    }


}

/// <summary>State of a signature / stamp preview that the user can move and resize.</summary>
internal sealed class PendingPlacement : IDisposable
{
    public enum Hit { None, Move, N, S, E, W, NE, NW, SE, SW }

    public int Page { get; }
    public RectangleF Norm { get; set; }
    public List<Control> Buttons { get; } = new();
    private readonly Func<SizeF, Bitmap?> _preview;
    private Bitmap? _bmp;
    private (SizeF Size, int Rot) _bmpKey;
    private Hit _drag = Hit.None;
    private Point _dragStart;
    private RectangleF _normStart;

    public PendingPlacement(int page, RectangleF norm, Func<SizeF, Bitmap?> preview)
    {
        Page = page; Norm = norm; _preview = preview;
    }

    public bool IsDragging => _drag != Hit.None;

    public Bitmap? GetPreview(SizeF uprightPts, int viewRotation)
    {
        var key = (new SizeF(MathF.Round(uprightPts.Width), MathF.Round(uprightPts.Height)), viewRotation);
        if (_bmp != null && _bmpKey == key) return _bmp;
        _bmp?.Dispose();
        _bmp = _preview(key.Item1);
        if (_bmp != null && viewRotation != 0)
            _bmp.RotateFlip(viewRotation switch { 1 => RotateFlipType.Rotate90FlipNone, 2 => RotateFlipType.Rotate180FlipNone, _ => RotateFlipType.Rotate270FlipNone });
        _bmpKey = key;
        return _bmp;
    }

    public void InvalidatePreview()
    {
        _bmp?.Dispose();
        _bmp = null;
    }

    public static IEnumerable<Rectangle> Handles(Rectangle r)
    {
        int s = Dpi.Scale(8), h = s / 2;
        int cx = r.Left + r.Width / 2, cy = r.Top + r.Height / 2;
        foreach (var p in new[] { new Point(r.Left, r.Top), new Point(cx, r.Top), new Point(r.Right, r.Top), new Point(r.Right, cy),
                     new Point(r.Right, r.Bottom), new Point(cx, r.Bottom), new Point(r.Left, r.Bottom), new Point(r.Left, cy) })
            yield return new Rectangle(p.X - h, p.Y - h, s, s);
    }

    public Hit? HitTest(PdfViewer v, Point client)
    {
        var r = v.PendingClientRect();
        int tol = Dpi.Scale(6);
        bool nearL = Math.Abs(client.X - r.Left) <= tol, nearR = Math.Abs(client.X - r.Right) <= tol;
        bool nearT = Math.Abs(client.Y - r.Top) <= tol, nearB = Math.Abs(client.Y - r.Bottom) <= tol;
        bool inX = client.X >= r.Left - tol && client.X <= r.Right + tol, inY = client.Y >= r.Top - tol && client.Y <= r.Bottom + tol;
        if (!inX || !inY) return null;
        if (nearT && nearL) return Hit.NW;
        if (nearT && nearR) return Hit.NE;
        if (nearB && nearL) return Hit.SW;
        if (nearB && nearR) return Hit.SE;
        if (nearT) return Hit.N;
        if (nearB) return Hit.S;
        if (nearL) return Hit.W;
        if (nearR) return Hit.E;
        return r.Contains(client) ? Hit.Move : null;
    }

    public void BeginDrag(Hit hit, Point client)
    {
        _drag = hit;
        _dragStart = client;
        _normStart = Norm;
    }

    public void DragTo(PdfViewer v, Point client)
    {
        var pr = v.PageClientRectOf(Page);
        float dx = (client.X - _dragStart.X) / (float)pr.Width, dy = (client.Y - _dragStart.Y) / (float)pr.Height;
        float l = _normStart.Left, t = _normStart.Top, r = _normStart.Right, b = _normStart.Bottom;
        float minW = 20f / pr.Width, minH = 14f / pr.Height;
        switch (_drag)
        {
            case Hit.Move:
                float w = r - l, h = b - t;
                l = Math.Clamp(l + dx, 0, 1 - w); t = Math.Clamp(t + dy, 0, 1 - h);
                r = l + w; b = t + h;
                break;
            default:
                if (_drag is Hit.W or Hit.NW or Hit.SW) l = Math.Clamp(l + dx, 0, r - minW);
                if (_drag is Hit.E or Hit.NE or Hit.SE) r = Math.Clamp(r + dx, l + minW, 1);
                if (_drag is Hit.N or Hit.NW or Hit.NE) t = Math.Clamp(t + dy, 0, b - minH);
                if (_drag is Hit.S or Hit.SW or Hit.SE) b = Math.Clamp(b + dy, t + minH, 1);
                break;
        }
        Norm = RectangleF.FromLTRB(l, t, r, b);
        if (_drag != Hit.Move) InvalidatePreview();
    }

    public void EndDrag() => _drag = Hit.None;

    public void Dispose()
    {
        _bmp?.Dispose();
        _bmp = null;
    }
}
