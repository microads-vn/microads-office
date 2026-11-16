using MicroadsOffice.Core.Office;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.UI;

namespace MicroadsOffice.Views;

/// <summary>Scrollable page thumbnails rendered lazily in the background.</summary>
public sealed class ThumbnailList : ScrollableControl
{
    private PdfiumDocument? _doc;
    private readonly Dictionary<int, Bitmap> _thumbs = new();
    private readonly HashSet<int> _pending = new();
    private int _current;
    private int _gen;
    private int ThumbW => Dpi.Scale(118);
    private int ItemH => Dpi.Scale(150) + Dpi.Scale(26);

    public event Action<int>? PageClicked;

    public ThumbnailList()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        AutoScroll = true;
        BackColor = Theme.PanelBack;
        AccessibleName = Loc.T("Hình thu nhỏ trang", "Page thumbnails");
        AccessibleRole = AccessibleRole.List;
    }

    public void SetDocument(PdfiumDocument? doc)
    {
        _gen++;
        foreach (var b in _thumbs.Values) b.Dispose();
        _thumbs.Clear();
        _pending.Clear();
        _doc = doc;
        AutoScrollMinSize = new Size(0, (doc?.PageCount ?? 0) * ItemH + Dpi.Scale(8));
        Invalidate();
    }

    public void SetCurrent(int page)
    {
        if (page == _current) return;
        _current = page;
        int top = page * ItemH, scroll = -AutoScrollPosition.Y;
        if (top < scroll || top + ItemH > scroll + ClientSize.Height)
            AutoScrollPosition = new Point(0, Math.Max(0, top - ClientSize.Height / 3));
        Invalidate();
    }

    protected override void OnScroll(ScrollEventArgs se) { base.OnScroll(se); Invalidate(); }
    protected override void OnMouseWheel(MouseEventArgs e) { base.OnMouseWheel(e); Invalidate(); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(BackColor);
        if (_doc == null) return;
        int scroll = -AutoScrollPosition.Y;
        int first = Math.Max(0, scroll / ItemH), last = Math.Min(_doc.PageCount - 1, (scroll + ClientSize.Height) / ItemH);
        using var sf = new StringFormat { Alignment = StringAlignment.Center };
        for (int i = first; i <= last; i++)
        {
            int y = i * ItemH - scroll + Dpi.Scale(8);
            var size = _doc.GetPageSize(i);
            int maxH = Dpi.Scale(150);
            float k = Math.Min((float)ThumbW / size.Width, (float)maxH / size.Height);
            int w = Math.Max(1, (int)(size.Width * k)), h = Math.Max(1, (int)(size.Height * k));
            var r = new Rectangle((ClientSize.Width - w) / 2, y + (maxH - h) / 2, w, h);
            if (i == _current)
            {
                using var sel = new SolidBrush(Theme.Hover);
                g.FillRectangle(sel, new Rectangle(Dpi.Scale(4), y - Dpi.Scale(4), ClientSize.Width - Dpi.Scale(8), ItemH - Dpi.Scale(4)));
            }
            g.FillRectangle(Brushes.White, r);
            if (_thumbs.TryGetValue(i, out var bmp)) g.DrawImage(bmp, r);
            else Enqueue(i, w, h);
            using var pen = new Pen(i == _current ? Theme.Accent : Theme.Border, i == _current ? 2 : 1);
            g.DrawRectangle(pen, r);
            g.DrawString((i + 1).ToString(), Theme.UiFont, i == _current ? Brushes.Firebrick : Brushes.DimGray,
                new RectangleF(0, y + maxH + Dpi.Scale(3), ClientSize.Width, Dpi.Scale(18)), sf);
        }
    }

    private void Enqueue(int page, int w, int h)
    {
        if (!_pending.Add(page) || _doc == null) return;
        var doc = _doc;
        int gen = _gen;
        Task.Run(() =>
        {
            Bitmap? bmp = null;
            try { bmp = doc.Render(page, w, h, 0); } catch { }
            if (bmp == null || IsDisposed) { bmp?.Dispose(); return; }
            try
            {
                BeginInvoke(() =>
                {
                    if (gen != _gen) { bmp.Dispose(); return; }
                    _thumbs[page] = bmp;
                    _pending.Remove(page);
                    Invalidate();
                });
            }
            catch { bmp.Dispose(); }
        });
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        base.OnMouseClick(e);
        if (_doc == null) return;
        int idx = (e.Y - AutoScrollPosition.Y) / ItemH;
        if (idx >= 0 && idx < _doc.PageCount) PageClicked?.Invoke(idx);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { _gen++; foreach (var b in _thumbs.Values) b.Dispose(); _thumbs.Clear(); }
        base.Dispose(disposing);
    }
}
