using MicroadsOffice.Core.Pdfium;

namespace MicroadsOffice.Viewer;

public readonly record struct TileKey(int Page, int W, int H, int Rot, int TX, int TY);

/// <summary>Background renderer: one worker thread, newest requests first, stale requests dropped.</summary>
internal sealed class RenderWorker : IDisposable
{
    private readonly object _lock = new();
    private readonly List<(TileKey Key, Rectangle? Clip, int Gen)> _queue = new();
    private readonly HashSet<TileKey> _queued = new();
    private readonly AutoResetEvent _signal = new(false);
    private readonly Thread _thread;
    private volatile bool _stop;
    private PdfiumDocument? _doc;
    private int _gen;

    public event Action<TileKey, Bitmap, int>? Completed;

    public RenderWorker()
    {
        _thread = new Thread(Loop) { IsBackground = true, Name = "PdfRender", Priority = ThreadPriority.BelowNormal };
        _thread.Start();
    }

    public int Generation => _gen;

    public void SetDocument(PdfiumDocument? doc)
    {
        lock (_lock)
        {
            _doc = doc;
            _gen++;
            _queue.Clear();
            _queued.Clear();
        }
    }

    /// <summary>Drops pending requests (e.g. after zoom change) but keeps the document.</summary>
    public void Reset()
    {
        lock (_lock)
        {
            _gen++;
            _queue.Clear();
            _queued.Clear();
        }
    }

    public void Request(TileKey key, Rectangle? clip)
    {
        lock (_lock)
        {
            if (!_queued.Add(key)) return;
            _queue.Add((key, clip, _gen));
        }
        _signal.Set();
    }

    private void Loop()
    {
        while (!_stop)
        {
            _signal.WaitOne(200);
            while (!_stop)
            {
                (TileKey Key, Rectangle? Clip, int Gen) item;
                PdfiumDocument? doc;
                lock (_lock)
                {
                    if (_queue.Count == 0) break;
                    item = _queue[^1]; // newest first (what the user is looking at now)
                    _queue.RemoveAt(_queue.Count - 1);
                    _queued.Remove(item.Key);
                    doc = _doc;
                    if (item.Gen != _gen || doc == null) continue;
                }
                try
                {
                    var bmp = doc.Render(item.Key.Page, item.Key.W, item.Key.H, item.Key.Rot, item.Clip);
                    Completed?.Invoke(item.Key, bmp, item.Gen);
                }
                catch
                {
                    // document closed / page broken: ignore
                }
            }
        }
    }

    public void Dispose()
    {
        _stop = true;
        _signal.Set();
        _thread.Join(1000);
        _signal.Dispose();
    }
}
