using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.IO;
using MicroadsOffice.Word.Layout;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.Editor;

public enum ZoomMode { Custom, PageWidth, OnePage }

/// <summary>
/// Print-layout editor: pages stacked vertically, custom caret and selection, keyboard and IME input
/// (Unikey / Windows Telex), mouse selection, image resizing, multi-level undo and rich clipboard.
/// </summary>
public sealed class WordEditor : ScrollableControl
{
    public const float PageGap = 18f;

    private Document _doc;
    private DocLayout _layout;
    private readonly TextMeasurer _m = new();
    private DocPos _anchor, _caret;
    private float _desiredX = -1;
    private CharProps? _pending;
    private float _zoom = 1f;
    private ZoomMode _zoomMode = ZoomMode.Custom;
    private bool _showMarks;
    private bool _caretOn = true;
    private readonly System.Windows.Forms.Timer _blink = new() { Interval = 530 };
    private readonly List<(Document Doc, DocPos A, DocPos C)> _undo = new(), _redo = new();
    private string? _lastEditKind;
    private DateTime _lastEditTime;
    private bool _mouseSelecting;
    private int _clickCount;
    private DateTime _lastClick;
    private Point _lastClickPt;
    private ImageObject? _selImage;
    private (ImageObject Img, int Handle, RectangleF Start, PointF Mouse)? _resizing;
    private int _savedVersion;
    private int _version;

    public event EventHandler? SelectionChanged;
    public event EventHandler? DocumentChanged;
    public event EventHandler? ZoomChanged;
    /// <summary>Raised after each mouse selection ends (format painter).</summary>
    public event EventHandler? MouseSelectionDone;

    public Func<Paragraph, IReadOnlyList<(int, int)>>? Misspelled { get; set; }

    public WordEditor()
    {
        Name = "wordEditor";
        _doc = Document.CreateBlank();
        _layout = new DocLayout(_doc, _m);
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
        AutoScroll = true;
        BackColor = Color.FromArgb(232, 234, 237);
        Cursor = Cursors.IBeam;
        TabStop = true;
        ImeMode = ImeMode.On;
        AccessibleName = Loc.T("Vùng soạn thảo", "Document editing area");
        AccessibleRole = AccessibleRole.Text;
        _blink.Tick += (_, _) => { _caretOn = !_caretOn; InvalidateCaret(); };
        UpdateScrollSize();
    }

    protected override bool CanEnableIme => true;

    // ------------------------------------------------------------------ document / state

    public Document Document
    {
        get => _doc;
        set
        {
            _doc = value;
            EditSeq++;
            _undo.Clear(); _redo.Clear();
            _anchor = _caret = new DocPos(0, 0);
            _pending = null;
            _selImage = null;
            _version++;
            _savedVersion = _version;
            Relayout();
            AutoScrollPosition = Point.Empty;
            DocumentChanged?.Invoke(this, EventArgs.Empty);
            OnSelectionChanged();
        }
    }

    public DocLayout LayoutResult => _layout;
    public TextMeasurer Measurer => _m;
    public DocPos Caret => _caret;
    public DocPos SelAnchor => _anchor;
    public DocPos SelStart => _anchor < _caret ? _anchor : _caret;
    public DocPos SelEnd => _anchor < _caret ? _caret : _anchor;
    public bool HasSelection => _anchor != _caret;
    public bool Modified => _version != _savedVersion;
    public void MarkSaved() { _savedVersion = _version; DocumentChanged?.Invoke(this, EventArgs.Empty); }
    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;
    public int PageCount => _layout.Pages.Count;
    public int CurrentPage => _layout.PageOf(_caret);
    public ImageObject? SelectedImage => _selImage;
    public Paragraph CurrentParagraph => _doc.Paragraphs[Editing.Clamp(_doc, _caret).Para];
    public TableCell? CurrentCell => CurrentParagraph.Cell;

    public bool ShowMarks { get => _showMarks; set { _showMarks = value; Invalidate(); } }
    private bool _gridlines = true;
    public bool ShowTableGridlines { get => _gridlines; set { _gridlines = value; Invalidate(); } }

    /// <summary>Moves the caret inside an <see cref="Edit"/> callback (no events, no scrolling).</summary>
    public void SetCaretInternal(DocPos p) { _anchor = _caret = p; }

    public CharProps CurrentCharProps => _pending ?? (HasSelection ? Editing.CharAt(_doc, SelStart.Offset < _doc.Paragraphs[SelStart.Para].Length ? new DocPos(SelStart.Para, SelStart.Offset + 1) : SelStart) : Editing.CharAt(_doc, _caret));
    public ParaProps CurrentParaProps => CurrentParagraph.Props;

    public void Relayout()
    {
        _doc.StructureChanged();
        _layout = new DocLayout(_doc, _m);
        UpdateScrollSize();
        Invalidate();
    }

    private void Changed()
    {
        _version++;
        if (_resizing == null) RefreshSelectedImage();
        Relayout();
        DocumentChanged?.Invoke(this, EventArgs.Empty);
        OnSelectionChanged();
        ScrollToCaret();
    }

    // ------------------------------------------------------------------ undo

    /// <summary>Runs a document change as one undoable step.</summary>
    /// <summary>Increases on every change, undo, redo and new document (the AI pane uses it to undo only its own change).</summary>
    public long EditSeq { get; private set; }

    public void Edit(string kind, Action change)
    {
        EditSeq++;
        bool coalesce = kind is "type" or "back" && _lastEditKind == kind && (DateTime.Now - _lastEditTime).TotalSeconds < 3 && _undo.Count > 0;
        if (!coalesce) PushUndo();
        _redo.Clear();
        _lastEditKind = kind;
        _lastEditTime = DateTime.Now;
        change();
        _caret = Editing.Clamp(_doc, _caret);
        _anchor = Editing.Clamp(_doc, _anchor);
        Changed();
    }

    private void PushUndo()
    {
        _undo.Add((_doc.Clone(), _anchor, _caret));
        if (_undo.Count > 200) _undo.RemoveAt(0);
    }

    public void Undo()
    {
        if (_undo.Count == 0) return;
        EditSeq++;
        _redo.Add((_doc, _anchor, _caret));
        var (d, a, c) = _undo[^1];
        _undo.RemoveAt(_undo.Count - 1);
        _doc = d; _anchor = Editing.Clamp(d, a); _caret = Editing.Clamp(d, c);
        _lastEditKind = null;
        _selImage = null;
        Changed();
    }

    public void Redo()
    {
        if (_redo.Count == 0) return;
        EditSeq++;
        _undo.Add((_doc, _anchor, _caret));
        var (d, a, c) = _redo[^1];
        _redo.RemoveAt(_redo.Count - 1);
        _doc = d; _anchor = Editing.Clamp(d, a); _caret = Editing.Clamp(d, c);
        _lastEditKind = null;
        _selImage = null;
        Changed();
    }

    // ------------------------------------------------------------------ geometry

    private float PxScale => _zoom * DeviceDpi / 72f;
    private float PageWpx => _doc.Page.Width * PxScale;
    private float PageHpx => _doc.Page.Height * PxScale;
    private float GapPx => PageGap * DeviceDpi / 96f;

    private float PageLeft => Math.Max(GapPx, (ClientSize.Width - PageWpx) / 2) + AutoScrollPosition.X;
    /// <summary>Client x of the page's left edge (for the ruler).</summary>
    public float PageLeftPx => PageLeft;
    /// <summary>Screen pixels per point at the current zoom.</summary>
    public float PxPerPoint => PxScale;
    private float PageTop(int i) => GapPx + i * (PageHpx + GapPx) + AutoScrollPosition.Y;

    private void UpdateScrollSize()
    {
        int n = Math.Max(1, _layout.Pages.Count);
        AutoScrollMinSize = new Size((int)(PageWpx + 2 * GapPx), (int)(GapPx + n * (PageHpx + GapPx)));
    }

    public float Zoom
    {
        get => _zoom;
        set
        {
            var z = Math.Clamp(value, 0.1f, 5f);
            _zoomMode = ZoomMode.Custom;
            SetZoomInternal(z);
        }
    }

    public ZoomMode ZoomMode
    {
        get => _zoomMode;
        set { _zoomMode = value; ApplyZoomMode(); }
    }

    private void SetZoomInternal(float z)
    {
        if (Math.Abs(z - _zoom) < 0.0001f) return;
        var caretPage = CurrentPage;
        _zoom = z;
        UpdateScrollSize();
        Invalidate();
        ZoomChanged?.Invoke(this, EventArgs.Empty);
        ScrollToCaret();
    }

    private void ApplyZoomMode()
    {
        if (ClientSize.Width < 50) return;
        float dpi = DeviceDpi / 72f;
        if (_zoomMode == ZoomMode.PageWidth) SetZoomInternal(Math.Clamp((ClientSize.Width - 2 * GapPx - SystemInformation.VerticalScrollBarWidth) / (_doc.Page.Width * dpi), 0.1f, 5f));
        else if (_zoomMode == ZoomMode.OnePage)
            SetZoomInternal(Math.Clamp(Math.Min((ClientSize.Width - 2 * GapPx) / (_doc.Page.Width * dpi), (ClientSize.Height - 2 * GapPx) / (_doc.Page.Height * dpi)), 0.1f, 5f));
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        if (_zoomMode != ZoomMode.Custom) ApplyZoomMode();
        Invalidate();
    }

    /// <summary>Converts a client point to (page, point in page units).</summary>
    private (int Page, PointF Pt) ClientToPage(Point p)
    {
        float y = p.Y - AutoScrollPosition.Y - GapPx;
        int page = (int)Math.Floor(y / (PageHpx + GapPx));
        page = Math.Clamp(page, 0, Math.Max(0, _layout.Pages.Count - 1));
        float px = (p.X - PageLeft) / PxScale, py = (p.Y - PageTop(page)) / PxScale;
        return (page, new PointF(px, py));
    }

    private RectangleF CaretClientRect()
    {
        var (page, x, y, h) = _layout.CaretAt(_caret);
        return new RectangleF(PageLeft + x * PxScale, PageTop(page) + y * PxScale, Math.Max(1, DeviceDpi / 96f), h * PxScale);
    }

    public void ScrollToCaret()
    {
        if (!IsHandleCreated) return;
        var r = CaretClientRect();
        int margin = (int)(30 * DeviceDpi / 96f);
        var pos = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y);
        if (r.Top < margin) pos.Y += (int)(r.Top - margin);
        else if (r.Bottom > ClientSize.Height - margin) pos.Y += (int)(r.Bottom - ClientSize.Height + margin);
        if (r.Left < 0) pos.X += (int)r.Left - margin;
        else if (r.Right > ClientSize.Width) pos.X += (int)(r.Right - ClientSize.Width) + margin;
        AutoScrollPosition = pos;
        UpdateImeWindow();
        Invalidate();
    }

    public void GoToPage(int index)
    {
        index = Math.Clamp(index, 0, _layout.Pages.Count - 1);
        var line = _layout.Pages[index].Lines.FirstOrDefault();
        if (line != null) SetCaret(_layout.PosOf(line, line.Line.Start), false);
        AutoScrollPosition = new Point(-AutoScrollPosition.X, (int)(GapPx + index * (PageHpx + GapPx) - GapPx / 2));
        Invalidate();
    }

    // ------------------------------------------------------------------ painting

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(BackColor);
        var opts = new RenderOptions
        {
            SelStart = _anchor, SelEnd = _caret, ShowMarks = _showMarks, SelectedImage = _selImage, ShowGridlines = _gridlines,
            Misspelled = Misspelled == null ? null : p => Misspelled(p).Select(t => (t.Item1, t.Item2)).ToList()
        };
        float s = PxScale;
        for (int i = 0; i < _layout.Pages.Count; i++)
        {
            float top = PageTop(i);
            if (top > e.ClipRectangle.Bottom || top + PageHpx < e.ClipRectangle.Top) continue;
            float left = PageLeft;
            using (var sh = new SolidBrush(Color.FromArgb(40, 0, 0, 0))) g.FillRectangle(sh, left + 2, top + 3, PageWpx, PageHpx);
            var st = g.Save();
            PageRenderer.Setup(g, left, top, s);
            g.SetClip(new RectangleF(0, 0, _doc.Page.Width, _doc.Page.Height));
            PageRenderer.DrawPage(g, _layout, i, opts);
            if (_showMarks || true) DrawMarginCorners(g);
            g.Restore(st);
        }
        if (Focused && !HasSelection && _caretOn && _selImage == null)
        {
            var r = CaretClientRect();
            using var b = new SolidBrush(Color.Black);
            g.FillRectangle(b, r.X, r.Y, Math.Max(1.5f, DeviceDpi / 96f * 1.2f), r.Height);
        }
    }

    private void DrawMarginCorners(Graphics g)
    {
        var ps = _doc.Page;
        using var pen = new Pen(Color.FromArgb(150, 160, 160, 165), 0.5f);
        float l = ps.MarginLeft, t = ps.MarginTop, r = ps.Width - ps.MarginRight, b = ps.Height - ps.MarginBottom, k = 10;
        g.DrawLine(pen, l - k, t, l, t); g.DrawLine(pen, l, t - k, l, t);
        g.DrawLine(pen, r, t, r + k, t); g.DrawLine(pen, r, t - k, r, t);
        g.DrawLine(pen, l - k, b, l, b); g.DrawLine(pen, l, b, l, b + k);
        g.DrawLine(pen, r, b, r + k, b); g.DrawLine(pen, r, b, r, b + k);
    }

    private void InvalidateCaret()
    {
        if (!IsHandleCreated) return;
        var r = CaretClientRect();
        Invalidate(Rectangle.Inflate(Rectangle.Round(r), 3, 3));
    }

    protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); _caretOn = true; _blink.Start(); UpdateImeWindow(); Invalidate(); }
    protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); _blink.Stop(); Invalidate(); }

    // ------------------------------------------------------------------ selection

    public void SetCaret(DocPos p, bool extend, bool keepDesiredX = false)
    {
        p = Editing.Clamp(_doc, p);
        _caret = p;
        if (!extend) _anchor = p;
        if (!keepDesiredX) _desiredX = -1;
        _pending = null;
        RefreshSelectedImage();
        _lastEditKind = null;
        _caretOn = true;
        ScrollToCaret();
        OnSelectionChanged();
    }

    /// <summary>The selection is exactly one inline image → that image is "selected" (resize handles, Picture tab).</summary>
    private void RefreshSelectedImage()
    {
        _selImage = null;
        if (SelStart.Para != SelEnd.Para || SelEnd.Offset - SelStart.Offset != 1) return;
        var para = _doc.Paragraphs[SelStart.Para];
        var (ri, off) = para.Locate(SelStart.Offset);
        if (ri < para.Runs.Count && off == 0 && para.Runs[ri].Object is ImageObject img) _selImage = img;
        else if (ri < para.Runs.Count && off == para.Runs[ri].Length && ri + 1 < para.Runs.Count && para.Runs[ri + 1].Object is ImageObject img2) _selImage = img2;
    }

    public void Select(DocPos a, DocPos b)
    {
        _anchor = Editing.Clamp(_doc, a);
        SetCaret(b, true);
    }

    public void SelectAll() => Select(new DocPos(0, 0), Editing.End(_doc));

    private void OnSelectionChanged()
    {
        Invalidate();
        SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    public string SelectedText => HasSelection ? Editing.GetText(_doc, SelStart, SelEnd) : "";

    // ------------------------------------------------------------------ editing commands

    public CharProps TypingProps
    {
        get
        {
            if (_pending != null) return _pending;
            var p = _doc.Paragraphs[_caret.Para];
            if (p.Length == 0) return p.MarkProps;
            var cp = _caret.Offset == 0 ? p.Runs[0].Props : p.PropsAt(_caret.Offset);
            if (cp.Link != null && (_caret.Offset == p.Length || _caret.Offset == 0)) cp = cp with { Link = null, Underline = UnderlineStyle.None, Color = null };
            return cp;
        }
    }

    public void InsertText(string text, string kind = "type")
    {
        var props = TypingProps;
        Edit(kind, () =>
        {
            var pos = HasSelection ? Editing.Delete(_doc, SelStart, SelEnd) : _caret;
            pos = Editing.InsertText(_doc, pos, text, props);
            _anchor = _caret = pos;
            _pending = null;
        });
    }

    public void InsertObject(InlineObject obj)
    {
        var props = TypingProps with { Link = null };
        Edit("object", () =>
        {
            var pos = HasSelection ? Editing.Delete(_doc, SelStart, SelEnd) : _caret;
            pos = Editing.InsertObject(_doc, pos, obj, props);
            _anchor = _caret = pos;
        });
    }

    public void InsertParagraphBreak()
    {
        var para = CurrentParagraph;
        if (!HasSelection && para.Length == 0 && para.Props.List != ListType.None)
        {
            // Enter on an empty list item ends the list (like Word)
            Edit("list", () => { para.Props = para.Props with { List = ListType.None, IndentLeft = 0, FirstLine = 0 }; para.Touch(); });
            return;
        }
        Edit("enter", () =>
        {
            var pos = HasSelection ? Editing.Delete(_doc, SelStart, SelEnd) : _caret;
            pos = Editing.SplitParagraph(_doc, pos);
            _anchor = _caret = pos;
        });
    }

    public void DeleteSelection()
    {
        if (!HasSelection) return;
        Edit("delete", () => { _anchor = _caret = Editing.Delete(_doc, SelStart, SelEnd); });
    }

    public void Backspace(bool word)
    {
        if (HasSelection) { DeleteSelection(); return; }
        var p = _caret;
        var para = _doc.Paragraphs[p.Para];
        if (p.Offset == 0)
        {
            if (para.Props.List != ListType.None)
            {
                Edit("list", () => { para.Props = para.Props with { List = ListType.None, IndentLeft = 0, FirstLine = 0 }; para.Touch(); });
                return;
            }
            if (para.Props.IndentLeft > 0 || para.Props.FirstLine > 0)
            {
                Edit("indent", () => { para.Props = para.Props with { FirstLine = para.Props.FirstLine > 0 ? 0 : para.Props.FirstLine, IndentLeft = para.Props.FirstLine > 0 ? para.Props.IndentLeft : Math.Max(0, para.Props.IndentLeft - 36) }; para.Touch(); });
                return;
            }
            if (p.Para == 0) return;
            var prev = _doc.Paragraphs[p.Para - 1];
            int len = prev.Length;
            if (prev.Cell != para.Cell) return; // table boundary
            Edit("join", () =>
            {
                if (Editing.JoinWithNext(_doc, p.Para - 1)) _anchor = _caret = new DocPos(p.Para - 1, len);
            });
            return;
        }
        int from = word ? Editing.PrevWord(para.Text, p.Offset) : p.Offset - 1;
        if (!word && from > 0 && char.IsLowSurrogate(para.Text[from]) && char.IsHighSurrogate(para.Text[from - 1])) from--;
        Edit(word ? "delete" : "back", () => { _anchor = _caret = Editing.Delete(_doc, new DocPos(p.Para, from), p); });
    }

    public void DeleteForward(bool word)
    {
        if (HasSelection) { DeleteSelection(); return; }
        var p = _caret;
        var para = _doc.Paragraphs[p.Para];
        if (p.Offset >= para.Length)
        {
            if (p.Para + 1 >= _doc.Paragraphs.Count) return;
            Edit("join", () => Editing.JoinWithNext(_doc, p.Para));
            return;
        }
        int to = word ? Editing.NextWord(para.Text, p.Offset) : p.Offset + 1;
        if (!word && to < para.Length && char.IsLowSurrogate(para.Text[to])) to++;
        Edit("delete", () => { _anchor = _caret = Editing.Delete(_doc, p, new DocPos(p.Para, to)); });
    }

    /// <summary>Applies character formatting to the selection, the word at the caret, or the typing format.</summary>
    public void ApplyChar(Func<CharProps, CharProps> f)
    {
        if (!HasSelection)
        {
            var para = CurrentParagraph;
            var (ws, we) = Editing.WordAt(para, _caret.Offset);
            bool inside = we > ws && _caret.Offset > ws && _caret.Offset < we && Editing.IsWordChar(para.Text[Math.Min(_caret.Offset, para.Length - 1)]);
            if (inside)
            {
                var t = para.Text;
                while (we > ws && t[we - 1] == ' ') we--;
                var a = new DocPos(_caret.Para, ws); var b = new DocPos(_caret.Para, we);
                Edit("format", () => Editing.ApplyChar(_doc, a, b, f));
                return;
            }
            _pending = f(TypingProps);
            if (para.Length == 0) { para.MarkProps = _pending; }
            OnSelectionChanged();
            return;
        }
        var s = SelStart; var e = SelEnd;
        Edit("format", () => Editing.ApplyChar(_doc, s, e, f));
    }

    public void ApplyPara(Func<ParaProps, ParaProps> f)
    {
        var s = SelStart; var e = SelEnd;
        Edit("para", () => Editing.ApplyPara(_doc, s, e, f));
    }

    public void ApplyStyle(string id)
    {
        var s = SelStart; var e = SelEnd;
        Edit("style", () => Editing.ApplyStyle(_doc, s, e, id));
    }

    public void ToggleList(ListType t)
    {
        var s = SelStart; var e = SelEnd;
        Edit("list", () => Editing.ToggleList(_doc, s, e, t));
    }

    public void ChangeIndent(int delta)
    {
        var s = SelStart; var e = SelEnd;
        Edit("indent", () => Editing.ChangeListLevel(_doc, s, e, delta));
    }

    public void InsertTable(int rows, int cols)
    {
        Edit("table", () =>
        {
            var pos = HasSelection ? Editing.Delete(_doc, SelStart, SelEnd) : _caret;
            if (_doc.Paragraphs[pos.Para].Cell != null) return;
            _anchor = _caret = Editing.InsertTable(_doc, pos, rows, cols);
        });
    }

    /// <summary>Runs a structural change of the current table; the caret is restored to the same paragraph object when it survives.</summary>
    public void TableEdit(Action<TableCell> op)
    {
        var cell = CurrentCell;
        if (cell == null) return;
        var keep = CurrentParagraph;
        int off = _caret.Offset;
        Edit("table", () =>
        {
            op(cell);
            _doc.StructureChanged();
            int i = _doc.IndexOf(keep);
            if (i < 0)
            {
                // the paragraph was deleted: go to the first paragraph of a remaining cell or after the table
                i = Math.Min(_caret.Para, _doc.Paragraphs.Count - 1);
                off = 0;
            }
            _anchor = _caret = Editing.Clamp(_doc, new DocPos(i, off));
        });
    }

    public (TableCell? A, TableCell? B) SelectedCells()
    {
        var a = _doc.Paragraphs[SelStart.Para].Cell;
        var b = _doc.Paragraphs[SelEnd.Para].Cell;
        return (a, b);
    }

    public void TransformSelection(Func<string, string> f)
    {
        if (!HasSelection) return;
        var s = SelStart; var e = SelEnd;
        Edit("case", () => Editing.TransformText(_doc, s, e, f));
    }

    /// <summary>Generic change of document-level properties (page setup, header, watermark...).</summary>
    public void DocumentEdit(Action<Document> change) => Edit("doc", () => change(_doc));

    public void ReplaceRange(DocPos a, DocPos b, string text)
    {
        var props = Editing.CharAt(_doc, new DocPos(a.Para, Math.Min(a.Offset + 1, _doc.Paragraphs[a.Para].Length)));
        Edit("replace", () =>
        {
            var p = Editing.Delete(_doc, a, b);
            p = Editing.InsertText(_doc, p, text, props);
            _anchor = a; _caret = p;
        });
    }

    // ------------------------------------------------------------------ clipboard

    private static List<Block>? _clipBlocks;
    private static string? _clipToken;
    private const string ClipFormat = "MicroadsWordFragment";

    public void Copy()
    {
        if (!HasSelection) return;
        var blocks = Editing.CopyRange(_doc, SelStart, SelEnd);
        _clipBlocks = blocks;
        _clipToken = Guid.NewGuid().ToString("N");
        var frag = new Document { Blocks = blocks.Select(b => b.CloneBlock()).ToList(), Styles = _doc.Styles };
        var data = new DataObject();
        // a picture's placeholder character must never reach other programs as "text" (Excel, chat apps would paste it instead of the picture)
        var text = SelectedText.Replace(Run.ObjectChar.ToString(), "").Replace("\n", Environment.NewLine);
        if (text.Length > 0) data.SetText(text, TextDataFormat.UnicodeText);
        try { data.SetData(DataFormats.Rtf, RtfWriter.Write(frag)); } catch { }
        try { data.SetData(DataFormats.Html, HtmlWriter.ClipboardHtml(frag)); } catch { }
        data.SetData(ClipFormat, _clipToken);
        if (_selImage != null) DataTransfer.PutImage(data, _selImage.Image);
        DataTransfer.WriteClipboard(data);
    }

    public void Cut()
    {
        if (!HasSelection) return;
        Copy();
        DeleteSelection();
    }

    /// <summary>Ctrl+V: Microads Word content, RTF, files from Explorer, a picture (screenshot, copied image) or text.</summary>
    public void Paste(bool textOnly = false) => InsertData(DataTransfer.ReadClipboard(), textOnly, fromClipboard: true);

    /// <summary>
    /// Inserts copied or dragged data at the caret as ONE undo step. Order: our own fragment, files and folders (pictures are
    /// inserted, other files and folders become links), formatted text (RTF), a picture, plain text. Returns false when
    /// there was nothing usable.
    /// </summary>
    public bool InsertData(IDataObject? data, bool textOnly = false, bool fromClipboard = false)
    {
        if (data == null) return false;
        if (fromClipboard && !textOnly && _clipBlocks != null && data.GetDataPresent(ClipFormat) && data.GetData(ClipFormat) as string == _clipToken)
        {
            PasteBlocksEdit(_clipBlocks);
            return true;
        }
        if (!textOnly)
        {
            var files = DataTransfer.Files(data);
            if (files.Length > 0) { InsertFiles(files); return true; }
            if (data.GetDataPresent(DataFormats.Rtf) && data.GetData(DataFormats.Rtf) is string rtf && rtf.Length > 0)
            {
                try
                {
                    var blocks = RtfReader.Read(rtf).Blocks;
                    // an RTF without anything in it (programs that copy only a picture) falls through to the picture
                    if (blocks.Any(b => b is Table || (b is Paragraph p && p.Length > 0))) { PasteBlocksEdit(blocks); return true; }
                }
                catch { }
            }
            if (DataTransfer.Picture(data) is { } pic) { InsertImage(pic.Bytes, pic.Ext); return true; }
        }
        var text = DataTransfer.Text(data);
        if (string.IsNullOrEmpty(text)) return false;
        InsertText(text, "paste");
        return true;
    }

    private void PasteBlocksEdit(List<Block> blocks) => Edit("paste", () =>
    {
        var pos = HasSelection ? Editing.Delete(_doc, SelStart, SelEnd) : _caret;
        _anchor = _caret = Editing.PasteBlocks(_doc, pos, blocks);
    });

    /// <summary>Files and folders from Explorer as ONE undo step: pictures are inserted, every other file or folder becomes a link (Ctrl+click shows it in Explorer).</summary>
    public void InsertFiles(IReadOnlyList<string> files)
    {
        var items = new List<object>();
        foreach (var f in files)
        {
            if (DataTransfer.IsImageFile(f) && DataTransfer.ImageFiles([f]).FirstOrDefault() is { } img) items.Add(FitImage(new ImageObject(img.Bytes, img.Ext)));
            else items.Add(f);
        }
        if (items.Count == 0) return;
        var props = TypingProps with { Link = null };
        Edit("paste", () =>
        {
            var pos = HasSelection ? Editing.Delete(_doc, SelStart, SelEnd) : _caret;
            for (int i = 0; i < items.Count; i++)
            {
                if (i > 0) pos = Editing.SplitParagraph(_doc, pos);
                pos = items[i] is ImageObject io
                    ? Editing.InsertObject(_doc, pos, io, props)
                    : Editing.InsertText(_doc, pos, (string)items[i], props with { Link = new Uri((string)items[i]).AbsoluteUri, Color = Color.FromArgb(5, 99, 193), Underline = UnderlineStyle.Single });
            }
            _anchor = _caret = pos;
            _pending = null;
        });
    }

    /// <summary>Moves the caret under the mouse while something is dragged over the page (shows where it will land).</summary>
    public void ShowDropCaret(Point client)
    {
        var (page, pt) = ClientToPage(client);
        var pos = _layout.HitTest(page, pt);
        if (pos != _caret || HasSelection) Select(pos, pos);
    }

    /// <summary>Drops dragged data (text, formatted text, a picture, files) at the point under the mouse.</summary>
    public bool DropData(IDataObject? data, Point client)
    {
        ShowDropCaret(client);
        Focus();
        return InsertData(data);
    }

    public static bool IsImageFile(string f) => DataTransfer.IsImageFile(f);

    public void InsertImageFile(string path)
    {
        if (DataTransfer.ImageFiles([path]).FirstOrDefault() is { } img) InsertImage(img.Bytes, img.Ext);
    }

    public void InsertImage(byte[] data, string ext) => InsertObject(FitImage(new ImageObject(data, ext)));

    /// <summary>Scales a new picture down to the column / page it goes into.</summary>
    private ImageObject FitImage(ImageObject img)
    {
        float maxW = _doc.Page.ContentWidth;
        var cell = CurrentCell;
        if (cell?.Row?.Table is Table t) maxW = Math.Max(30, t.ColumnWidths.Skip(t.GridColumnOf(cell.Row, cell)).Take(Math.Max(1, cell.GridSpan)).Sum() - 2 * DocLayout.CellPad);
        if (img.Width > maxW) { float k = maxW / img.Width; img.Width *= k; img.Height *= k; }
        if (img.Height > _doc.Page.ContentHeight * 0.9f) { float k = _doc.Page.ContentHeight * 0.9f / img.Height; img.Width *= k; img.Height *= k; }
        return img;
    }

    public void ResizeSelectedImage(float w, float h)
    {
        var img = _selImage;
        if (img == null) return;
        Edit("image", () => { img.Width = Math.Max(4, w); img.Height = Math.Max(4, h); CurrentParagraph.Touch(); });
    }

    // ------------------------------------------------------------------ keyboard

    protected override bool IsInputKey(Keys keyData)
    {
        var k = keyData & Keys.KeyCode;
        if (k is Keys.Left or Keys.Right or Keys.Up or Keys.Down or Keys.Tab or Keys.Enter or Keys.Home or Keys.End or Keys.PageUp or Keys.PageDown or Keys.Escape) return true;
        return base.IsInputKey(keyData);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.Handled) return;
        bool shift = e.Shift, ctrl = e.Control;
        var para = CurrentParagraph;
        switch (e.KeyCode)
        {
            case Keys.Left:
                if (!shift && HasSelection) SetCaret(SelStart, false);
                else SetCaret(MoveCaret(_caret, -1, ctrl), shift);
                break;
            case Keys.Right:
                if (!shift && HasSelection) SetCaret(SelEnd, false);
                else SetCaret(MoveCaret(_caret, 1, ctrl), shift);
                break;
            case Keys.Up:
            case Keys.Down:
            {
                int dir = e.KeyCode == Keys.Up ? -1 : 1;
                if (ctrl)
                {
                    var np = dir < 0 ? (_caret.Offset > 0 ? new DocPos(_caret.Para, 0) : new DocPos(Math.Max(0, _caret.Para - 1), 0)) : new DocPos(Math.Min(_doc.Paragraphs.Count - 1, _caret.Para + 1), 0);
                    SetCaret(np, shift);
                    break;
                }
                var (_, cx, _, _) = _layout.CaretAt(_caret);
                if (_desiredX < 0) _desiredX = cx;
                var target = _layout.Vertical(_caret, _desiredX, dir);
                if (target is DocPos t) SetCaret(t, shift, keepDesiredX: true);
                else SetCaret(dir < 0 ? new DocPos(0, 0) : Editing.End(_doc), shift, true);
                break;
            }
            case Keys.Home:
                if (ctrl) SetCaret(new DocPos(0, 0), shift);
                else { var l = _layout.LineOf(_caret); SetCaret(new DocPos(_caret.Para, l?.Line.Start ?? 0), shift); }
                break;
            case Keys.End:
                if (ctrl) SetCaret(Editing.End(_doc), shift);
                else
                {
                    var l = _layout.LineOf(_caret);
                    int end = l?.Line.End ?? para.Length;
                    if (l != null && !l.Line.IsLast && end > l.Line.Start && para.Text[end - 1] == ' ') end--;
                    if (l != null && l.Line.Segs.Count > 0 && l.Line.Segs[^1].Kind == SegKind.Break) end = l.Line.Segs[^1].Start;
                    SetCaret(new DocPos(_caret.Para, end), shift);
                }
                break;
            case Keys.PageUp:
            case Keys.PageDown:
            {
                int dir = e.KeyCode == Keys.PageUp ? -1 : 1;
                var r = CaretClientRect();
                var pt = new Point((int)r.X, (int)(r.Y + dir * ClientSize.Height * 0.9f));
                AutoScrollPosition = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y + (int)(dir * ClientSize.Height * 0.9f));
                pt = new Point((int)r.X, (int)Math.Clamp(r.Y, 0, ClientSize.Height - 1));
                var (pg, ppt) = ClientToPage(pt);
                SetCaret(_layout.HitTest(pg, ppt), shift);
                break;
            }
            case Keys.Back: Backspace(ctrl); break;
            case Keys.Delete: DeleteForward(ctrl); break;
            case Keys.Enter:
                if (ctrl) InsertText(Run.PageBreak.ToString(), "break");
                else if (shift) InsertText(Run.LineBreak.ToString(), "break");
                else InsertParagraphBreak();
                break;
            case Keys.Tab:
                HandleTab(shift);
                break;
            case Keys.Escape:
                if (HasSelection) SetCaret(_caret, false);
                break;
            default:
                if (ctrl && HandleShortcut(e)) break;
                return;
        }
        e.Handled = true;
        e.SuppressKeyPress = true;
    }

    private bool HandleShortcut(KeyEventArgs e)
    {
        bool shift = e.Shift;
        switch (e.KeyCode)
        {
            case Keys.A: SelectAll(); return true;
            case Keys.C: case Keys.Insert: Copy(); return true;
            case Keys.X: Cut(); return true;
            case Keys.V: Paste(shift); return true;
            case Keys.Z: Undo(); return true;
            case Keys.Y: Redo(); return true;
            case Keys.B: { bool v = !CurrentCharProps.Bold; ApplyChar(p => p with { Bold = v }); return true; }
            case Keys.I: { bool v = !CurrentCharProps.Italic; ApplyChar(p => p with { Italic = v }); return true; }
            case Keys.U: { var v = CurrentCharProps.Underline == UnderlineStyle.None ? (shift ? UnderlineStyle.Double : UnderlineStyle.Single) : UnderlineStyle.None; ApplyChar(p => p with { Underline = v }); return true; }
            case Keys.L: ApplyPara(p => p with { Align = ParaAlign.Left }); return true;
            case Keys.E: ApplyPara(p => p with { Align = ParaAlign.Center }); return true;
            case Keys.R: ApplyPara(p => p with { Align = ParaAlign.Right }); return true;
            case Keys.J: ApplyPara(p => p with { Align = ParaAlign.Justify }); return true;
            case Keys.OemCloseBrackets: ApplyChar(p => p with { Size = Math.Min(1638, p.Size + 1) }); return true;
            case Keys.OemOpenBrackets: ApplyChar(p => p with { Size = Math.Max(1, p.Size - 1) }); return true;
            case Keys.Oemplus:
            {
                var target = shift ? VertAlign.Superscript : VertAlign.Subscript;
                var v = CurrentCharProps.VertAlign == target ? VertAlign.Baseline : target;
                ApplyChar(p => p with { VertAlign = v });
                return true;
            }
            case Keys.OemPeriod when shift: ApplyChar(p => p with { Size = FontSizes.Next(p.Size, true) }); return true;
            case Keys.Oemcomma when shift: ApplyChar(p => p with { Size = FontSizes.Next(p.Size, false) }); return true;
            case Keys.Space: ApplyChar(p => CharProps.Default with { FontName = p.FontName, Size = p.Size }); return true;
            case Keys.M: ChangeIndent(shift ? -1 : 1); return true;
        }
        return false;
    }

    private void HandleTab(bool shift)
    {
        var cell = CurrentCell;
        if (cell != null)
        {
            var t = cell.Row!.Table!;
            var cells = t.Rows.SelectMany(r => r.Cells).Where(c => c.VMerge != VMerge.Continue).ToList();
            int i = cells.IndexOf(cell);
            if (!shift && i == cells.Count - 1)
            {
                // Tab in the last cell adds a row
                TableEdit(c => Editing.InsertRow(_doc, c, true));
                cells = t.Rows.SelectMany(r => r.Cells).ToList();
                var nc = t.Rows[^1].Cells[0];
                SelectCell(nc);
                return;
            }
            int j = Math.Clamp(i + (shift ? -1 : 1), 0, cells.Count - 1);
            SelectCell(cells[j]);
            return;
        }
        var para = CurrentParagraph;
        if (para.Props.List != ListType.None && _caret.Offset == 0 && !HasSelection) { ChangeIndent(shift ? -1 : 1); return; }
        if (shift) return;
        InsertText("\t");
    }

    private void SelectCell(TableCell c)
    {
        var first = c.Paragraphs[0]; var last = c.Paragraphs[^1];
        int a = _doc.IndexOf(first), b = _doc.IndexOf(last);
        if (a < 0) return;
        Select(new DocPos(a, 0), new DocPos(b, last.Length));
    }

    private DocPos MoveCaret(DocPos p, int dir, bool word)
    {
        var ps = _doc.Paragraphs;
        var t = ps[p.Para].Text;
        if (dir < 0)
        {
            if (p.Offset == 0) return p.Para > 0 ? new DocPos(p.Para - 1, ps[p.Para - 1].Length) : p;
            int o = word ? Editing.PrevWord(t, p.Offset) : p.Offset - 1;
            if (!word && o > 0 && char.IsLowSurrogate(t[o])) o--;
            return new DocPos(p.Para, o);
        }
        if (p.Offset >= t.Length) return p.Para + 1 < ps.Count ? new DocPos(p.Para + 1, 0) : p;
        int n = word ? Editing.NextWord(t, p.Offset) : p.Offset + 1;
        if (!word && n < t.Length && char.IsLowSurrogate(t[n])) n++;
        return new DocPos(p.Para, n);
    }

    protected override void OnKeyPress(KeyPressEventArgs e)
    {
        base.OnKeyPress(e);
        if (e.Handled) return;
        char c = e.KeyChar;
        if (c < 32 || c == 127) return;
        if ((ModifierKeys & Keys.Control) != 0 && (ModifierKeys & Keys.Alt) == 0) return;
        InsertText(c.ToString());
        e.Handled = true;
    }

    private const int WM_IME_CHAR = 0x0286, WM_IME_STARTCOMPOSITION = 0x010D;

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_IME_CHAR)
        {
            char c = (char)(long)m.WParam;
            if (c >= 32) InsertText(c.ToString());
            return;
        }
        if (m.Msg == WM_IME_STARTCOMPOSITION) UpdateImeWindow();
        base.WndProc(ref m);
    }

    [StructLayout(LayoutKind.Sequential)] private struct COMPOSITIONFORM { public int dwStyle; public Point ptCurrentPos; public Rectangle rcArea; }
    [DllImport("imm32.dll")] private static extern IntPtr ImmGetContext(IntPtr hWnd);
    [DllImport("imm32.dll")] private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
    [DllImport("imm32.dll")] private static extern bool ImmSetCompositionWindow(IntPtr hIMC, ref COMPOSITIONFORM lpCompForm);

    private void UpdateImeWindow()
    {
        if (!IsHandleCreated || !Focused) return;
        try
        {
            var ctx = ImmGetContext(Handle);
            if (ctx == IntPtr.Zero) return;
            var r = CaretClientRect();
            var cf = new COMPOSITIONFORM { dwStyle = 2, ptCurrentPos = new Point((int)r.X, (int)r.Y) };
            ImmSetCompositionWindow(ctx, ref cf);
            ImmReleaseContext(Handle, ctx);
        }
        catch { }
    }

    // ------------------------------------------------------------------ mouse

    private RectangleF? ImageRectClient(ImageObject img)
    {
        foreach (var pl in _layout.AllLines)
        {
            foreach (var s in pl.Line.Segs)
            {
                if (s.Obj != img) continue;
                float top = PageTop(pl.Page), left = PageLeft;
                var r = new RectangleF(pl.X + s.X, pl.Y + pl.Line.Baseline - img.Height, img.Width, img.Height);
                return new RectangleF(left + r.X * PxScale, top + r.Y * PxScale, r.Width * PxScale, r.Height * PxScale);
            }
        }
        return null;
    }

    private int HandleAt(Point p)
    {
        if (_selImage == null) return -1;
        var r = ImageRectClient(_selImage);
        if (r == null) return -1;
        var hs = PageRenderer.Handles(r.Value);
        for (int i = 0; i < hs.Length; i++) if (RectangleF.Inflate(hs[i], 4, 4).Contains(p)) return i;
        return -1;
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        Focus();
        if (e.Button == MouseButtons.Left)
        {
            int h = HandleAt(e.Location);
            if (h >= 0 && _selImage != null)
            {
                _resizing = (_selImage, h, new RectangleF(0, 0, _selImage.Width, _selImage.Height), e.Location);
                return;
            }
            var (page, pt) = ClientToPage(e.Location);
            var pos = _layout.HitTest(page, pt);
            var now = DateTime.Now;
            if ((now - _lastClick).TotalMilliseconds < SystemInformation.DoubleClickTime && Math.Abs(e.X - _lastClickPt.X) < 5 && Math.Abs(e.Y - _lastClickPt.Y) < 5) _clickCount++;
            else _clickCount = 1;
            _lastClick = now; _lastClickPt = e.Location;

            // Ctrl+click on a link opens it
            var cp = _doc.Paragraphs[pos.Para].PropsAt(Math.Max(1, pos.Offset));
            if ((ModifierKeys & Keys.Control) != 0 && cp.Link != null)
            {
                OpenLink(cp.Link);
                return;
            }
            // click on an image selects it
            var img = ImageAt(e.Location, out var imgPos);
            if (img != null && _clickCount == 1)
            {
                Select(imgPos, new DocPos(imgPos.Para, imgPos.Offset + 1));
                return;
            }
            if (_clickCount == 2)
            {
                var (s, en) = Editing.WordAt(_doc.Paragraphs[pos.Para], pos.Offset);
                Select(new DocPos(pos.Para, s), new DocPos(pos.Para, en));
                return;
            }
            if (_clickCount >= 3)
            {
                Select(new DocPos(pos.Para, 0), new DocPos(pos.Para, _doc.Paragraphs[pos.Para].Length));
                return;
            }
            SetCaret(pos, (ModifierKeys & Keys.Shift) != 0);
            _mouseSelecting = true;
            Capture = true;
        }
        else if (e.Button == MouseButtons.Right)
        {
            var (page, pt) = ClientToPage(e.Location);
            var pos = _layout.HitTest(page, pt);
            if (!HasSelection || pos < SelStart || pos > SelEnd) SetCaret(pos, false);
        }
    }

    private ImageObject? ImageAt(Point p, out DocPos pos)
    {
        pos = default;
        foreach (var pl in _layout.AllLines)
        {
            float top = PageTop(pl.Page);
            if (top > p.Y || top + PageHpx < p.Y) continue;
            foreach (var s in pl.Line.Segs)
            {
                if (s.Obj is not ImageObject img) continue;
                var r = new RectangleF(PageLeft + (pl.X + s.X) * PxScale, top + (pl.Y + pl.Line.Baseline - img.Height) * PxScale, img.Width * PxScale, img.Height * PxScale);
                if (r.Contains(p)) { pos = _layout.PosOf(pl, s.Start); return img; }
            }
        }
        return null;
    }

    public static void OpenLink(string url)
    {
        // links to files / folders (pasted from Explorer): show them in Explorer; a document link never runs a program
        if (url.StartsWith("file:", StringComparison.OrdinalIgnoreCase) && Uri.TryCreate(url, UriKind.Absolute, out var fu) && fu.IsFile)
        {
            try
            {
                if (Directory.Exists(fu.LocalPath)) Process.Start(new ProcessStartInfo("explorer.exe", "\"" + fu.LocalPath + "\"") { UseShellExecute = false });
                else if (File.Exists(fu.LocalPath)) Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + fu.LocalPath + "\"") { UseShellExecute = false });
            }
            catch { }
            return;
        }
        if (!(url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) || url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))) return;
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        if (_resizing is { } rz)
        {
            float dx = (e.X - rz.Mouse.X) / PxScale, dy = (e.Y - rz.Mouse.Y) / PxScale;
            float w = rz.Start.Width, h = rz.Start.Height;
            int hi = rz.Handle;
            if (hi is 2 or 4 or 7) w += dx;
            if (hi is 0 or 3 or 5) w -= dx;
            if (hi is 5 or 6 or 7) h += dy;
            if (hi is 0 or 1 or 2) h -= dy;
            if (hi is 0 or 2 or 5 or 7) h = w * rz.Start.Height / Math.Max(1, rz.Start.Width); // corners keep the aspect ratio
            rz.Img.Width = Math.Max(8, w);
            rz.Img.Height = Math.Max(8, h);
            CurrentParagraph.Touch();
            Relayout();
            return;
        }
        if (_mouseSelecting && e.Button == MouseButtons.Left)
        {
            var (page, pt) = ClientToPage(e.Location);
            var pos = _layout.HitTest(page, pt);
            if (_clickCount <= 1 && pos != _caret) { _caret = pos; _selImage = null; OnSelectionChanged(); }
            if (e.Y < 0) AutoScrollPosition = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y - 20);
            else if (e.Y > ClientSize.Height) AutoScrollPosition = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y + 20);
            return;
        }
        int hdl = HandleAt(e.Location);
        if (hdl >= 0) Cursor = hdl is 0 or 7 ? Cursors.SizeNWSE : hdl is 2 or 5 ? Cursors.SizeNESW : hdl is 1 or 6 ? Cursors.SizeNS : Cursors.SizeWE;
        else
        {
            var (page, pt) = ClientToPage(e.Location);
            bool onPage = pt.X >= 0 && pt.X <= _doc.Page.Width && pt.Y >= 0 && pt.Y <= _doc.Page.Height;
            if (onPage && (ModifierKeys & Keys.Control) != 0)
            {
                var pos = _layout.HitTest(page, pt);
                Cursor = _doc.Paragraphs[pos.Para].PropsAt(Math.Max(1, pos.Offset)).Link != null ? Cursors.Hand : Cursors.IBeam;
            }
            else Cursor = onPage ? (ImageAt(e.Location, out _) != null ? Cursors.Default : Cursors.IBeam) : Cursors.Default;
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (_resizing is { } rz)
        {
            _resizing = null;
            float w = rz.Img.Width, h = rz.Img.Height;
            rz.Img.Width = rz.Start.Width; rz.Img.Height = rz.Start.Height;
            var img = rz.Img;
            Edit("image", () => { img.Width = w; img.Height = h; CurrentParagraph.Touch(); });
            _selImage = img;
            Invalidate();
            return;
        }
        if (_mouseSelecting)
        {
            _mouseSelecting = false;
            Capture = false;
            SetCaret(_caret, true);
            MouseSelectionDone?.Invoke(this, EventArgs.Empty);
        }
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Control) != 0)
        {
            Zoom = (float)Math.Round(_zoom + (e.Delta > 0 ? 0.1f : -0.1f), 2);
            return;
        }
        base.OnMouseWheel(e);
        Invalidate();
    }

    protected override void OnScroll(ScrollEventArgs se) { base.OnScroll(se); Invalidate(); }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { _blink.Dispose(); _m.Dispose(); }
        base.Dispose(disposing);
    }

    /// <summary>Selection highlight for find results.</summary>
    public void SelectRange(DocPos a, DocPos b) => Select(a, b);
}

public static class FontSizes
{
    public static readonly float[] Standard = [8, 9, 10, 10.5f, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

    public static float Next(float s, bool up)
    {
        if (up) { foreach (var v in Standard) if (v > s + 0.01f) return v; return Math.Min(1638, s + 10); }
        for (int i = Standard.Length - 1; i >= 0; i--) if (Standard[i] < s - 0.01f) return Standard[i];
        return Math.Max(1, s - 1);
    }
}
