using System.Drawing.Drawing2D;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.UI;

/// <summary>
/// Office-style ribbon with grouped commands (used by Word and Excel).
/// <code>
/// var home = ribbon.AddPage("Home", Loc.T("Trang chủ","Home"));
/// var g = home.AddGroup("Font", Loc.T("Phông chữ","Font"));
/// g.Row(fontCombo, sizeCombo);                 // up to 3 rows form one column block
/// g.Row(boldBtn, italicBtn, underlineBtn);
/// g.Large(pasteBtn);                            // full-height button
/// g.Stack(cutBtn, copyBtn, painterBtn);         // column of up to 3 small labelled buttons
/// </code>
/// Groups collapse into a single drop-down button (right to left) when the window is too narrow.
/// Ribbon buttons are not focusable, so clicking them never steals the caret from the document.
/// </summary>
public sealed class RibbonBar : UserControl
{
    private readonly ToolStrip _header;
    private readonly Panel _body;
    private readonly List<RibbonPage> _pages = new();
    private readonly ToolStripButton _fileButton;
    private int _qatCount;
    private bool _minimized;

    public RibbonPage? SelectedPage { get; private set; }
    public event EventHandler? FileClicked;
    public event EventHandler? SelectedPageChanged;
    /// <summary>
    /// Raised before any ribbon command runs or a tab is chosen. Apps close their File (backstage) view here,
    /// so commands always act on the visible document, like Office.
    /// </summary>
    public event EventHandler? CommandInvoking;
    internal void RaiseCommandInvoking() => CommandInvoking?.Invoke(this, EventArgs.Empty);
    public ToolTip ToolTip { get; } = new() { InitialDelay = 500, ReshowDelay = 100, AutoPopDelay = 12000, ShowAlways = true };

    public static int BodyHeight => Dpi.Scale(98);

    public RibbonBar()
    {
        DoubleBuffered = true;
        BackColor = Theme.RibbonBack;
        Dock = DockStyle.Top;
        Name = "ribbon";
        _header = new ToolStrip
        {
            Name = "ribbonTabs",
            GripStyle = ToolStripGripStyle.Hidden,
            Dock = DockStyle.Top,
            BackColor = Theme.RibbonBack,
            Renderer = new RibbonRenderer(header: true),
            Padding = new Padding(Dpi.Scale(4), Dpi.Scale(2), Dpi.Scale(4), 0),
            Font = new Font("Segoe UI", 9.5f),
            CanOverflow = false,
            ImageScalingSize = new Size(Dpi.Scale(16), Dpi.Scale(16)),
            AccessibleName = Loc.T("Thẻ Ribbon", "Ribbon tabs")
        };
        _fileButton = new ToolStripButton(Loc.T("Tệp", "File"))
        {
            Name = "btnFileMenu",
            ForeColor = Color.White,
            BackColor = Theme.Accent,
            Padding = new Padding(Dpi.Scale(12), Dpi.Scale(2), Dpi.Scale(12), Dpi.Scale(2)),
            Font = new Font("Segoe UI Semibold", 9.5f),
            AccessibleName = Loc.T("Tệp", "File")
        };
        _fileButton.Click += (_, _) => FileClicked?.Invoke(this, EventArgs.Empty);
        _header.Items.Add(_fileButton);
        _header.Items[0].Margin = new Padding(0, 0, Dpi.Scale(6), 0);
        _body = new Panel { Dock = DockStyle.Top, Height = BodyHeight, BackColor = Theme.RibbonBack, Name = "ribbonBody" };
        _body.Paint += (_, e) =>
        {
            using var p = new Pen(Theme.Border);
            e.Graphics.DrawLine(p, 0, _body.Height - 1, _body.Width, _body.Height - 1);
        };
        Controls.Add(_body);
        Controls.Add(_header);
        Height = _header.PreferredSize.Height + BodyHeight;
        _header.SizeChanged += (_, _) => UpdateHeight();
    }

    private void UpdateHeight() => Height = _header.Height + (_minimized ? 0 : BodyHeight);

    /// <summary>Quick access button (Save, Undo, Redo) placed before the File tab.</summary>
    public ToolStripButton AddQuickAccess(string name, Image icon, string tooltip, Action onClick)
    {
        var b = new ToolStripButton(icon)
        {
            Name = name, DisplayStyle = ToolStripItemDisplayStyle.Image, ToolTipText = tooltip, AccessibleName = tooltip,
            ImageScaling = ToolStripItemImageScaling.None, Margin = new Padding(0, 0, Dpi.Scale(1), 0)
        };
        b.Click += (_, _) => onClick();
        _header.Items.Insert(_qatCount++, b);
        if (_qatCount == 1) _header.Items.Insert(_qatCount++, new ToolStripSeparator());
        else
        {
            // keep the separator right after the last quick access button
            var sep = _header.Items.OfType<ToolStripSeparator>().First();
            _header.Items.Remove(sep);
            _header.Items.Insert(_qatCount - 1, sep);
        }
        return b;
    }

    public RibbonPage AddPage(string name, string text, bool contextual = false)
    {
        var page = new RibbonPage(this, name, text, contextual);
        page.Header.Click += (_, _) => { RaiseCommandInvoking(); if (_minimized) SetMinimized(false); Select(page); };
        page.Header.DoubleClickEnabled = true;
        page.Header.DoubleClick += (_, _) => SetMinimized(!_minimized);
        // insert before right-aligned items
        int idx = _header.Items.Cast<ToolStripItem>().TakeWhile(i => i.Alignment == ToolStripItemAlignment.Left).Count();
        _header.Items.Insert(idx, page.Header);
        _body.Controls.Add(page);
        _pages.Add(page);
        if (contextual) page.Header.Visible = false;
        if (SelectedPage == null && !contextual) Select(page);
        return page;
    }

    public IReadOnlyList<RibbonPage> Pages => _pages;

    public void Select(RibbonPage page)
    {
        if (!page.Header.Visible) return;
        foreach (var p in _pages)
        {
            p.Visible = p == page;
            p.Header.Checked = p == page;
        }
        SelectedPage = page;
        page.PerformLayout();
        SelectedPageChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Shows or hides a contextual page (e.g. Table / Picture / Chart tools).</summary>
    public void SetContextual(RibbonPage page, bool visible, bool select = false)
    {
        if (page.Header.Visible == visible) { if (visible && select) Select(page); return; }
        page.Header.Visible = visible;
        if (visible && select) Select(page);
        else if (!visible && SelectedPage == page) Select(_pages.First(p => !p.Contextual));
    }

    public void AddHeaderRight(params ToolStripItem[] items)
    {
        foreach (var item in items)
        {
            item.Alignment = ToolStripItemAlignment.Right;
            _header.Items.Add(item);
        }
    }

    public bool Minimized => _minimized;

    /// <summary>Collapses the ribbon to its tab row (Ctrl+F1 / double-click a tab).</summary>
    public void SetMinimized(bool value)
    {
        _minimized = value;
        _body.Visible = !value;
        UpdateHeight();
    }

    internal ToolStrip HeaderStrip => _header;
}

/// <summary>One ribbon tab: a horizontal row of groups.</summary>
public sealed class RibbonPage : Panel
{
    private readonly RibbonBar _bar;
    private readonly List<RibbonGroup> _groups = new();
    public ToolStripButton Header { get; }
    public bool Contextual { get; }
    public string Title { get; }

    internal RibbonPage(RibbonBar bar, string name, string text, bool contextual)
    {
        _bar = bar;
        Name = "page" + name;
        Title = text;
        Contextual = contextual;
        Dock = DockStyle.Fill;
        Visible = false;
        BackColor = Theme.RibbonBack;
        DoubleBuffered = true;
        AccessibleName = text;
        Header = new ToolStripButton(text)
        {
            Name = "tab" + name,
            Padding = new Padding(Dpi.Scale(8), Dpi.Scale(2), Dpi.Scale(8), Dpi.Scale(2)),
            AccessibleName = Loc.T("Thẻ ", "Tab ") + text.Replace("&&", "&"),
            ForeColor = contextual ? Theme.AccentDark : Theme.Ink
        };
        if (contextual) Header.Font = new Font("Segoe UI Semibold", 9.5f);
    }

    public RibbonBar Bar => _bar;

    public RibbonGroup AddGroup(string name, string caption)
    {
        var g = new RibbonGroup(this, name, caption);
        _groups.Add(g);
        Controls.Add(g);
        return g;
    }

    public IReadOnlyList<RibbonGroup> Groups => _groups;

    protected override void OnLayout(LayoutEventArgs e)
    {
        if (_groups.Count == 0) return;
        int avail = ClientSize.Width - Dpi.Scale(8);
        foreach (var g in _groups) g.Collapsed = false;
        int total = _groups.Sum(g => g.ExpandedWidth);
        for (int i = _groups.Count - 1; i >= 0 && total > avail; i--)
        {
            if (!_groups[i].CanCollapse) continue;
            total -= _groups[i].ExpandedWidth - _groups[i].CollapsedWidth;
            _groups[i].Collapsed = true;
        }
        int x = Dpi.Scale(4);
        foreach (var g in _groups)
        {
            int w = g.Collapsed ? g.CollapsedWidth : g.ExpandedWidth;
            g.SetBounds(x, 0, w, Height - 1);
            x += w;
        }
    }

    protected override void OnResize(EventArgs eventargs) { base.OnResize(eventargs); PerformLayout(); }
}

/// <summary>A ribbon group: items arranged in columns plus a caption at the bottom.</summary>
public sealed class RibbonGroup : Panel
{
    private abstract record Block;
    private sealed record LargeBlock(Control Item) : Block;
    private sealed record StackBlock(List<Control> Items) : Block;
    private sealed record RowsBlock(List<List<Control>> Rows) : Block;

    private readonly RibbonPage _page;
    private readonly List<Block> _blocks = new();
    private readonly Panel _content;
    private readonly RibbonButton _collapsedButton;
    private RowsBlock? _openRows;
    private Button? _launcher;
    private bool _collapsed;
    private ToolStripDropDown? _popup;

    public string Caption { get; }
    public bool CanCollapse { get; set; } = true;
    /// <summary>Icon for the collapsed form of the group (defaults to the first large item's icon).</summary>
    public Image? CollapsedIcon { get; set; }

    internal RibbonGroup(RibbonPage page, string name, string caption)
    {
        _page = page;
        Name = "grp" + name;
        Caption = caption;
        AccessibleName = caption;
        DoubleBuffered = true;
        BackColor = Theme.RibbonBack;
        _content = new Panel { BackColor = Theme.RibbonBack, Name = "grpContent" + name };
        Controls.Add(_content);
        _collapsedButton = new RibbonButton(caption, Icons.Large(Icons.More), RibbonButtonSize.Large)
        {
            Name = "grpCollapsed" + name, Visible = false, DropDownArrow = true, ToolTipText = caption
        };
        _collapsedButton.Click += (_, _) => ShowPopup();
        Controls.Add(_collapsedButton);
    }

    private static int RowH => Dpi.Scale(24);
    private static int ContentH => RowH * 3;
    private static int CaptionH => Dpi.Scale(18);
    private static int Pad => Dpi.Scale(4);

    public T Large<T>(T item) where T : Control
    {
        _openRows = null;
        Prepare(item, RibbonButtonSize.Large);
        _blocks.Add(new LargeBlock(item));
        _expandedWidth = -1;
        _content.Controls.Add(item);
        CollapsedIcon ??= (item as RibbonButton)?.Image;
        return item;
    }

    /// <summary>A column of up to three small (labelled) items.</summary>
    public void Stack(params Control[] items)
    {
        _openRows = null;
        foreach (var it in items) { Prepare(it, RibbonButtonSize.Small); _content.Controls.Add(it); }
        _blocks.Add(new StackBlock(items.ToList()));
        _expandedWidth = -1;
        CollapsedIcon ??= (items.FirstOrDefault() as RibbonButton)?.Image;
    }

    /// <summary>Adds a horizontal row; consecutive rows (max 3) form one block.</summary>
    public void Row(params Control[] items)
    {
        if (_openRows == null || _openRows.Rows.Count >= 3)
        {
            _openRows = new RowsBlock(new List<List<Control>>());
            _blocks.Add(_openRows);
        }
        foreach (var it in items) { Prepare(it, RibbonButtonSize.Icon); _content.Controls.Add(it); }
        _openRows.Rows.Add(items.ToList());
        _expandedWidth = -1;
        CollapsedIcon ??= (items.FirstOrDefault(i => i is RibbonButton) as RibbonButton)?.Image;
    }

    /// <summary>Ends the current row block so the next Row() starts a new column.</summary>
    public void Break() => _openRows = null;

    /// <summary>Small ↘ button in the caption that opens the full dialog of the group.</summary>
    public void SetDialogLauncher(string tooltip, Action onClick)
    {
        _launcher = new Button
        {
            Name = Name + "Launcher", FlatStyle = FlatStyle.Flat, Size = new Size(Dpi.Scale(16), Dpi.Scale(16)), TabStop = false,
            Image = Icons.Get(Icons.Launcher, Dpi.Scale(10), Theme.SubtleInk), AccessibleName = tooltip, Cursor = Cursors.Hand
        };
        _launcher.FlatAppearance.BorderSize = 0;
        _launcher.FlatAppearance.MouseOverBackColor = Theme.Hover;
        _launcher.Click += (_, _) => onClick();
        _page.Bar.ToolTip.SetToolTip(_launcher, tooltip);
        _content.Controls.Add(_launcher);
    }

    private void Prepare(Control c, RibbonButtonSize size)
    {
        if (c is RibbonButton rb)
        {
            if (rb.SizeMode == RibbonButtonSize.Auto) rb.SizeMode = size;
            if (!string.IsNullOrEmpty(rb.ToolTipText)) _page.Bar.ToolTip.SetToolTip(rb, rb.ToolTipText);
        }
        else if (c is ComboBox cb)
        {
            cb.Height = RowH - Dpi.Scale(2);
            cb.Font = Theme.UiFont;
        }
        c.Margin = Padding.Empty;
    }

    internal bool Collapsed
    {
        get => _collapsed;
        set
        {
            if (_collapsed == value) return;
            _collapsed = value;
            _collapsedButton.Visible = value;
            if (value) _collapsedButton.Image = CollapsedIcon ?? Icons.Large(Icons.More);
            if (_popup == null) _content.Visible = !value;
            Invalidate();
        }
    }

    internal int CollapsedWidth => Math.Max(Dpi.Scale(64), TextRenderer.MeasureText(Caption, Theme.UiFont).Width + Dpi.Scale(18));

    private int _expandedWidth = -1;
    internal int ExpandedWidth
    {
        get
        {
            if (_expandedWidth < 0) _expandedWidth = Math.Max(ArrangeContent(apply: false) + Pad * 2 + 1, TextRenderer.MeasureText(Caption, Theme.SmallFont).Width + Dpi.Scale(28));
            return _expandedWidth;
        }
    }

    /// <summary>Call after changing the text of items so the group re-measures.</summary>
    public void Remeasure() { _expandedWidth = -1; _page.PerformLayout(); }

    private int ArrangeContent(bool apply)
    {
        int x = 0;
        int gap = Dpi.Scale(2);
        foreach (var b in _blocks)
        {
            switch (b)
            {
                case LargeBlock lb:
                {
                    var w = Pref(lb.Item, true).Width;
                    if (apply) lb.Item.SetBounds(x, 0, w, ContentH);
                    x += w + gap;
                    break;
                }
                case StackBlock sb:
                {
                    int w = sb.Items.Max(i => Pref(i, false).Width);
                    int y = 0;
                    foreach (var it in sb.Items)
                    {
                        if (apply) it.SetBounds(x, y, w, RowH);
                        y += RowH;
                    }
                    x += w + gap;
                    break;
                }
                case RowsBlock rb:
                {
                    int maxW = 0, y = 0;
                    int extra = rb.Rows.Count < 3 ? (ContentH - rb.Rows.Count * RowH) / (rb.Rows.Count + 1) : 0;
                    foreach (var row in rb.Rows)
                    {
                        y += extra;
                        int rx = x;
                        foreach (var it in row)
                        {
                            var sz = Pref(it, false);
                            if (apply) it.SetBounds(rx, y + (RowH - sz.Height) / 2, sz.Width, sz.Height);
                            rx += sz.Width + (it is ComboBox ? Dpi.Scale(3) : 0);
                        }
                        maxW = Math.Max(maxW, rx - x);
                        y += RowH;
                    }
                    x += maxW + gap + Dpi.Scale(2);
                    break;
                }
            }
        }
        return Math.Max(0, x - gap);
    }

    private static Size Pref(Control c, bool large)
    {
        if (c is RibbonButton rb) return rb.GetPreferredSize(Size.Empty);
        if (c is ComboBox cb) return new Size(cb.Width, Dpi.Scale(22));
        return new Size(c.Width, Math.Min(c.Height, large ? ContentH : RowH));
    }

    protected override void OnLayout(LayoutEventArgs levent)
    {
        int w = ArrangeContent(apply: true);
        if (_popup == null)
        {
            _content.SetBounds(Pad, Dpi.Scale(3), Math.Max(w, Width - Pad * 2), ContentH);
        }
        _collapsedButton?.SetBounds(Dpi.Scale(3), Dpi.Scale(3), Width - Dpi.Scale(7), ContentH); // null while the constructor adds _content
        if (_launcher != null && _popup == null)
        {
            _launcher.Parent = this;
            _launcher.Location = new Point(Width - _launcher.Width - Dpi.Scale(4), Height - CaptionH + Dpi.Scale(1));
            _launcher.Visible = !_collapsed;
            _launcher.BringToFront();
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        if (!_collapsed)
        {
            var cap = new Rectangle(0, Height - CaptionH - Dpi.Scale(1), Width - Dpi.Scale(18), CaptionH);
            TextRenderer.DrawText(g, Caption, Theme.SmallFont, new Rectangle(0, cap.Y, Width, CaptionH), Theme.SubtleInk,
                TextFormatFlags.NoPrefix | TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
        using var p = new Pen(Theme.Border);
        g.DrawLine(p, Width - 1, Dpi.Scale(8), Width - 1, Height - Dpi.Scale(8));
    }

    private void ShowPopup()
    {
        if (_popup != null) return;
        int w = ArrangeContent(apply: true);
        var holder = new Panel { BackColor = Theme.RibbonBack, Size = new Size(w + Pad * 2, ContentH + CaptionH + Dpi.Scale(6)) };
        _content.Parent = holder;
        _content.SetBounds(Pad, Dpi.Scale(3), w, ContentH);
        _content.Visible = true;
        var cap = new Label { Text = Caption, Dock = DockStyle.Bottom, Height = CaptionH, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Theme.SubtleInk, Font = Theme.SmallFont };
        holder.Controls.Add(cap);
        var host = new ToolStripControlHost(holder) { Margin = Padding.Empty, Padding = Padding.Empty, AutoSize = false, Size = holder.Size };
        _popup = new ToolStripDropDown { Padding = Padding.Empty, AutoClose = true, Name = Name + "Popup" };
        _popup.Items.Add(host);
        _popup.Closed += (_, _) =>
        {
            var pp = _popup;
            _popup = null;
            _content.Parent = this;
            _content.Visible = !_collapsed;
            holder.Controls.Remove(cap);
            BeginInvoke(() => { pp?.Dispose(); PerformLayout(); });
        };
        foreach (var rb in _content.Controls.OfType<RibbonButton>()) rb.Click += ClosePopupOnClick;
        _popup.Show(this, new Point(0, Height));
    }

    private void ClosePopupOnClick(object? sender, EventArgs e)
    {
        if (sender is RibbonButton rb) rb.Click -= ClosePopupOnClick;
        if (sender is RibbonButton { DropDown: not null } or RibbonButton { DropDownArrow: true }) return;
        _popup?.Close();
    }
}

public enum RibbonButtonSize { Auto, Large, Small, Icon }

/// <summary>
/// Owner-drawn, non-focusable ribbon button. Supports toggle state (Checked), drop-down menus,
/// split buttons (main part + arrow) and a colour swatch under the icon (font / fill colour).
/// </summary>
public class RibbonButton : Button
{
    private bool _hover, _pressed, _hoverArrow, _checked;
    private Color? _swatch;
    private Image? _image;

    public RibbonButtonSize SizeMode { get; set; }
    public string ToolTipText { get; set; } = "";
    /// <summary>Menu shown when the button (or the arrow part of a split button) is clicked.</summary>
    public ContextMenuStrip? DropDown { get; set; }
    /// <summary>Popup shown instead of <see cref="DropDown"/> (e.g. colour picker, table grid).</summary>
    public Func<ToolStripDropDown>? DropDownFactory { get; set; }
    /// <summary>When true the main part runs Click and only the arrow opens the drop-down.</summary>
    public bool Split { get; set; }
    public bool DropDownArrow { get; set; }
    public new Image? Image { get => _image; set { _image = value; Invalidate(); } }
    public bool Checked { get => _checked; set { if (_checked == value) return; _checked = value; Invalidate(); AccessibleDescription = value ? "checked" : ""; } }
    public Color? SwatchColor { get => _swatch; set { _swatch = value; Invalidate(); } }
    /// <summary>Raised when the main part of a split button is clicked (Click fires too).</summary>
    public event EventHandler? DropDownOpening;

    public RibbonButton(string text, Image? image, RibbonButtonSize size = RibbonButtonSize.Auto)
    {
        Text = text;
        _image = image;
        SizeMode = size;
        AccessibleName = text.Replace("\n", " ");
        SetStyle(ControlStyles.Selectable, false);
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        TabStop = false;
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        BackColor = Theme.RibbonBack;
        Font = Theme.UiFont;
        Cursor = Cursors.Default;
    }

    /// <summary>Convenience factory used by the apps.</summary>
    public static RibbonButton Create(string name, string text, Image? image, string tooltip, Action? onClick, RibbonButtonSize size = RibbonButtonSize.Auto)
    {
        var b = new RibbonButton(text, image, size) { Name = name, ToolTipText = tooltip };
        if (onClick != null) b.Click += (_, _) => onClick();
        if (string.IsNullOrEmpty(text)) b.AccessibleName = tooltip.Split('\n')[0].Split(" (")[0];
        return b;
    }

    private bool HasMenu => DropDown != null || DropDownFactory != null;
    private bool ShowArrow => HasMenu || DropDownArrow;

    private int ArrowW => Dpi.Scale(12);

    public override Size GetPreferredSize(Size proposedSize)
    {
        var text = Text ?? "";
        switch (SizeMode)
        {
            case RibbonButtonSize.Large:
            {
                var lines = text.Split('\n');
                int tw = lines.Max(l => TextRenderer.MeasureText(l, Font, Size.Empty, TextFormatFlags.NoPrefix).Width) + (ShowArrow && lines.Length == 1 ? ArrowW : 0);
                return new Size(Math.Max(Dpi.Scale(44), tw + Dpi.Scale(10)), Dpi.Scale(72));
            }
            case RibbonButtonSize.Small:
            {
                int tw = string.IsNullOrEmpty(text) ? 0 : TextRenderer.MeasureText(text.Replace("\n", " "), Font).Width + Dpi.Scale(4);
                return new Size(Dpi.Scale(24) + tw + (ShowArrow ? ArrowW : 0) + Dpi.Scale(4), Dpi.Scale(24));
            }
            default:
                return new Size(Dpi.Scale(24) + (ShowArrow ? ArrowW : 0), Dpi.Scale(24));
        }
    }

    private Rectangle ArrowRect => SizeMode == RibbonButtonSize.Large
        ? new Rectangle(0, Height - Dpi.Scale(26), Width, Dpi.Scale(26))
        : new Rectangle(Width - ArrowW - Dpi.Scale(2), 0, ArrowW + Dpi.Scale(2), Height);

    protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { _hover = false; _hoverArrow = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseMove(MouseEventArgs e)
    {
        bool a = Split && ArrowRect.Contains(e.Location);
        if (a != _hoverArrow) { _hoverArrow = a; Invalidate(); }
        base.OnMouseMove(e);
    }
    protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
    protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }

    private bool _arrowClick;

    protected override void OnMouseClick(MouseEventArgs e)
    {
        _arrowClick = Split && ArrowRect.Contains(e.Location);
        base.OnMouseClick(e);
    }

    protected override void OnClick(EventArgs e)
    {
        if (!Enabled) return;
        bool openMenu = HasMenu && (!Split || _arrowClick);
        _arrowClick = false;
        FindBar()?.RaiseCommandInvoking();
        if (openMenu)
        {
            DropDownOpening?.Invoke(this, EventArgs.Empty);
            var pt = new Point(0, Height);
            if (DropDownFactory != null)
            {
                var dd = DropDownFactory();
                dd.Show(this, pt);
            }
            else DropDown!.Show(this, pt);
            return;
        }
        base.OnClick(e);
    }

    /// <summary>Clicks the button from code, UI Automation (Invoke) or screen readers (the base implementation
    /// does nothing because ribbon buttons are deliberately not focusable).</summary>
    public new void PerformClick()
    {
        if (!Enabled || !Visible) return;
        OnClick(EventArgs.Empty);
    }

    private RibbonBar? FindBar()
    {
        for (Control? c = Parent; c != null; c = c.Parent) if (c is RibbonBar rb) return rb;
        return null;
    }

    /// <summary>Opens the drop-down programmatically (also used by UI tests through the arrow).</summary>
    public void OpenDropDown()
    {
        if (!HasMenu) return;
        DropDownOpening?.Invoke(this, EventArgs.Empty);
        if (DropDownFactory != null) DropDownFactory().Show(this, new Point(0, Height));
        else DropDown!.Show(this, new Point(0, Height));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? Theme.RibbonBack);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        if (Enabled && (_hover || _pressed || _checked))
        {
            var col = _pressed || _checked ? Theme.Pressed : Theme.Hover;
            using var path = BrandRound(r, Dpi.Scale(3));
            using var b = new SolidBrush(col);
            g.FillPath(b, path);
            if (_checked)
            {
                using var pen = new Pen(Theme.PressedBorder);
                g.DrawPath(pen, path);
            }
            if (Split && _hover)
            {
                using var sep = new Pen(Theme.PressedBorder);
                var ar = ArrowRect;
                if (SizeMode == RibbonButtonSize.Large) g.DrawLine(sep, 2, ar.Top, Width - 3, ar.Top);
                else g.DrawLine(sep, ar.Left, 3, ar.Left, Height - 4);
                if (_hoverArrow)
                {
                    using var hb = new SolidBrush(Theme.Pressed);
                    g.FillRectangle(hb, SizeMode == RibbonButtonSize.Large ? new Rectangle(1, ar.Top + 1, Width - 3, ar.Height - 2) : new Rectangle(ar.Left + 1, 1, ar.Width - 3, Height - 3));
                }
            }
        }
        var ink = Enabled ? Theme.Ink : Color.FromArgb(160, 160, 165);
        int iconSize = SizeMode == RibbonButtonSize.Large ? Dpi.Scale(32) : Dpi.Scale(16);
        Rectangle iconRect;
        if (SizeMode == RibbonButtonSize.Large)
            iconRect = new Rectangle((Width - iconSize) / 2, Dpi.Scale(4), iconSize, iconSize);
        else
            iconRect = new Rectangle(Dpi.Scale(4), (Height - iconSize) / 2, iconSize, iconSize);
        if (_image != null)
        {
            if (Enabled) g.DrawImage(_image, iconRect);
            else ControlPaint.DrawImageDisabled(g, _image, iconRect.X, iconRect.Y, Color.Transparent);
        }
        if (_swatch is Color sw)
        {
            var sr = SizeMode == RibbonButtonSize.Large
                ? new Rectangle(iconRect.X + Dpi.Scale(2), iconRect.Bottom - Dpi.Scale(5), iconRect.Width - Dpi.Scale(4), Dpi.Scale(5))
                : new Rectangle(iconRect.X, iconRect.Bottom - Dpi.Scale(3), iconRect.Width, Dpi.Scale(4));
            if (sw == Color.Transparent || sw.A == 0)
            {
                using var pen = new Pen(Theme.Border);
                g.DrawRectangle(pen, sr);
            }
            else
            {
                using var sb = new SolidBrush(sw);
                g.FillRectangle(sb, sr);
            }
        }
        var text = Text ?? "";
        if (SizeMode == RibbonButtonSize.Large)
        {
            var lines = text.Split('\n');
            int y = iconRect.Bottom + Dpi.Scale(3);
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                bool last = i == lines.Length - 1;
                var sz = TextRenderer.MeasureText(line, Font, Size.Empty, TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
                int extra = last && ShowArrow ? ArrowW : 0;
                int tx = (Width - sz.Width - extra) / 2;
                TextRenderer.DrawText(g, line, Font, new Point(tx, y), ink, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
                if (last && ShowArrow) DrawArrow(g, new Point(tx + sz.Width + Dpi.Scale(5), y + sz.Height / 2), ink);
                y += sz.Height - Dpi.Scale(1);
            }
            if (lines.Length == 1 && ShowArrow && string.IsNullOrEmpty(text)) DrawArrow(g, new Point(Width / 2, Height - Dpi.Scale(10)), ink);
        }
        else
        {
            int tx = iconRect.Right + Dpi.Scale(4);
            if (SizeMode == RibbonButtonSize.Small && !string.IsNullOrEmpty(text))
                TextRenderer.DrawText(g, text.Replace("\n", " "), Font, new Rectangle(tx, 0, Width - tx, Height), ink, TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            if (ShowArrow) DrawArrow(g, new Point(Width - ArrowW / 2 - Dpi.Scale(2), Height / 2), ink);
        }
    }

    private static void DrawArrow(Graphics g, Point c, Color color)
    {
        int s = Dpi.Scale(3);
        using var b = new SolidBrush(color);
        g.FillPolygon(b, [new Point(c.X - s, c.Y - s / 2), new Point(c.X + s + 1, c.Y - s / 2), new Point(c.X, c.Y + s / 2 + 1)]);
    }

    private static GraphicsPath BrandRound(Rectangle r, int rad) => MicroadsOffice.Core.Branding.BrandArt.RoundRect(r, rad);

    protected override AccessibleObject CreateAccessibilityInstance() => new RibbonButtonAccessible(this);

    private sealed class RibbonButtonAccessible(RibbonButton owner) : ButtonBaseAccessibleObject(owner)
    {
        // Return to the UI Automation client at once (a command may open a modal dialog), then click.
        public override void DoDefaultAction() { if (owner.IsHandleCreated) owner.BeginInvoke(owner.PerformClick); }

        public override AccessibleStates State
        {
            get
            {
                var s = base.State;
                if (owner.Checked) s |= AccessibleStates.Checked;
                return s;
            }
        }
    }
}

/// <summary>Office-like colour palette drop-down (theme colours, standard colours, automatic/no colour, more colours).</summary>
public sealed class ColorPickerDropDown : ToolStripDropDown
{
    public static readonly Color[] ThemeRow =
    [
        Color.White, Color.Black, Color.FromArgb(231, 230, 230), Color.FromArgb(68, 84, 106), Color.FromArgb(68, 114, 196),
        Color.FromArgb(237, 125, 49), Color.FromArgb(165, 165, 165), Color.FromArgb(255, 192, 0), Color.FromArgb(91, 155, 213), Color.FromArgb(112, 173, 71)
    ];

    public static readonly Color[] StandardRow =
    [
        Color.FromArgb(192, 0, 0), Color.FromArgb(255, 0, 0), Color.FromArgb(255, 192, 0), Color.FromArgb(255, 255, 0), Color.FromArgb(146, 208, 80),
        Color.FromArgb(0, 176, 80), Color.FromArgb(0, 176, 240), Color.FromArgb(0, 112, 192), Color.FromArgb(0, 32, 96), Color.FromArgb(112, 48, 160)
    ];

    /// <param name="emptyText">Text of the first option ("Automatic" / "No fill"); null hides it.</param>
    public ColorPickerDropDown(Action<Color?> onPick, string? emptyText, Color? current = null)
    {
        Name = "colorPicker";
        var panel = new ColorPalettePanel(c => { Close(); onPick(c); }, emptyText, current);
        var host = new ToolStripControlHost(panel) { AutoSize = false, Size = panel.Size, Margin = Padding.Empty, Padding = Padding.Empty };
        Padding = new Padding(2);
        Items.Add(host);
    }

    private sealed class ColorPalettePanel : Panel
    {
        private readonly Action<Color?> _pick;
        private readonly string? _empty;
        private readonly List<(Rectangle R, Color? C, string Name)> _cells = new();
        private int _hover = -1;

        public ColorPalettePanel(Action<Color?> pick, string? emptyText, Color? current)
        {
            _pick = pick; _empty = emptyText;
            DoubleBuffered = true;
            BackColor = Color.White;
            int cs = Dpi.Scale(18), gap = Dpi.Scale(4), x0 = Dpi.Scale(8);
            int y = Dpi.Scale(6);
            if (emptyText != null)
            {
                _cells.Add((new Rectangle(x0, y, 10 * (cs + gap) - gap, Dpi.Scale(24)), null, emptyText));
                y += Dpi.Scale(30);
            }
            y += Dpi.Scale(16); // caption "Theme colours"
            for (int col = 0; col < 10; col++) _cells.Add((new Rectangle(x0 + col * (cs + gap), y, cs, cs), ThemeRow[col], Hex(ThemeRow[col])));
            y += cs + Dpi.Scale(4);
            float[] tints = [0.8f, 0.6f, 0.4f, -0.25f, -0.5f];
            foreach (var t in tints)
            {
                for (int col = 0; col < 10; col++)
                {
                    var c = t > 0 ? Theme.Blend(ThemeRow[col], Color.White, t) : Theme.Blend(ThemeRow[col], Color.Black, -t);
                    _cells.Add((new Rectangle(x0 + col * (cs + gap), y, cs, cs), c, Hex(c)));
                }
                y += cs;
            }
            y += Dpi.Scale(22);
            for (int col = 0; col < 10; col++) _cells.Add((new Rectangle(x0 + col * (cs + gap), y, cs, cs), StandardRow[col], Hex(StandardRow[col])));
            y += cs + Dpi.Scale(8);
            _cells.Add((new Rectangle(x0, y, 10 * (cs + gap) - gap, Dpi.Scale(24)), Color.Empty, Loc.T("Màu khác...", "More colours...")));
            y += Dpi.Scale(30);
            Size = new Size(x0 * 2 + 10 * (cs + gap) - gap, y);
            AccessibleName = Loc.T("Bảng màu", "Colour palette");
        }

        private static string Hex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int h = _cells.FindIndex(c => c.R.Contains(e.Location));
            if (h != _hover) { _hover = h; Invalidate(); }
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            int i = _cells.FindIndex(c => c.R.Contains(e.Location));
            if (i < 0) return;
            var cell = _cells[i];
            if (cell.C == Color.Empty)
            {
                using var d = new ColorDialog { FullOpen = true, AnyColor = true };
                if (d.ShowDialog(FindForm()) == DialogResult.OK) _pick(d.Color);
                return;
            }
            _pick(cell.C);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            using var border = new Pen(Color.FromArgb(200, 200, 205));
            using var hot = new Pen(Theme.AccentDark, 2);
            int captionY = _empty != null ? Dpi.Scale(36) : Dpi.Scale(6);
            TextRenderer.DrawText(g, Loc.T("Màu chủ đề", "Theme colours"), Theme.SmallFont, new Point(Dpi.Scale(6), captionY), Theme.SubtleInk);
            for (int i = 0; i < _cells.Count; i++)
            {
                var (r, c, name) = _cells[i];
                if (c == null || c == Color.Empty)
                {
                    if (i == _hover) { using var hb = new SolidBrush(Theme.Hover); g.FillRectangle(hb, r); }
                    TextRenderer.DrawText(g, name, Theme.UiFont, r, Theme.Ink, TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
                    continue;
                }
                using var b = new SolidBrush(c.Value);
                g.FillRectangle(b, r);
                g.DrawRectangle(i == _hover ? hot : border, r);
            }
            var std = _cells.FindIndex(x => x.C == StandardRow[0]);
            if (std >= 0) TextRenderer.DrawText(g, Loc.T("Màu chuẩn", "Standard colours"), Theme.SmallFont, new Point(Dpi.Scale(6), _cells[std].R.Top - Dpi.Scale(18)), Theme.SubtleInk);
        }
    }
}
