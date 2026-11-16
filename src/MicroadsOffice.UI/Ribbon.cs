using System.Drawing.Drawing2D;

namespace MicroadsOffice.UI;

/// <summary>Office-style ribbon: a row of tab headers and one ToolStrip of large buttons per tab.</summary>
public sealed class Ribbon : UserControl
{
    private readonly ToolStrip _header;
    private readonly Panel _body;
    private readonly List<RibbonTab> _tabs = new();
    public ToolStripDropDownButton FileButton { get; }
    public RibbonTab? SelectedTab { get; private set; }

    public Ribbon()
    {
        DoubleBuffered = true;
        BackColor = Theme.RibbonBack;
        Dock = DockStyle.Top;
        _header = new ToolStrip
        {
            Name = "ribbonTabs",
            GripStyle = ToolStripGripStyle.Hidden,
            Dock = DockStyle.Top,
            BackColor = Theme.RibbonBack,
            Renderer = new RibbonRenderer(header: true),
            Padding = new Padding(Dpi.Scale(4), Dpi.Scale(2), 0, 0),
            Font = new Font("Segoe UI", 9.5f),
            CanOverflow = false,
            AccessibleName = MicroadsOffice.Core.Office.Loc.T("Thẻ Ribbon", "Ribbon tabs")
        };
        FileButton = new ToolStripDropDownButton("Tệp")
        {
            Name = "btnFileMenu",
            ShowDropDownArrow = false,
            ForeColor = Color.White,
            BackColor = Theme.Accent,
            Padding = new Padding(Dpi.Scale(10), Dpi.Scale(2), Dpi.Scale(10), Dpi.Scale(2)),
            Font = new Font("Segoe UI Semibold", 9.5f),
            AccessibleName = "Tệp"
        };
        _header.Items.Add(FileButton);
        _body = new Panel { Dock = DockStyle.Fill, BackColor = Theme.RibbonBack };
        Controls.Add(_body);
        Controls.Add(_header);
        Height = Dpi.Scale(112);
    }

    public RibbonTab AddTab(string name, string text)
    {
        var strip = new ToolStrip
        {
            Name = "ribbon" + name,
            GripStyle = ToolStripGripStyle.Hidden,
            Dock = DockStyle.Fill,
            BackColor = Theme.RibbonBack,
            Renderer = new RibbonRenderer(header: false),
            ImageScalingSize = new Size(Dpi.Scale(28), Dpi.Scale(28)),
            Padding = new Padding(Dpi.Scale(6), Dpi.Scale(2), Dpi.Scale(6), Dpi.Scale(2)),
            AutoSize = false,
            Visible = false,
            AccessibleName = text
        };
        var header = new ToolStripButton(text) { Name = "tab" + name, Padding = new Padding(Dpi.Scale(8), Dpi.Scale(2), Dpi.Scale(8), Dpi.Scale(2)), AccessibleName = MicroadsOffice.Core.Office.Loc.T("Thẻ ", "Tab ") + text.Replace("&&", "&") };
        var tab = new RibbonTab(name, text, strip, header);
        header.Click += (_, _) => Select(tab);
        _header.Items.Add(header);
        _body.Controls.Add(strip);
        _tabs.Add(tab);
        if (SelectedTab == null) Select(tab);
        return tab;
    }

    public void Select(RibbonTab tab)
    {
        foreach (var t in _tabs)
        {
            t.Strip.Visible = t == tab;
            t.Header.Checked = t == tab;
        }
        SelectedTab = tab;
    }

    public void AddHeaderRight(ToolStripItem item)
    {
        item.Alignment = ToolStripItemAlignment.Right;
        _header.Items.Add(item);
    }
}

public sealed class RibbonTab
{
    public string Name { get; }
    public string Text { get; }
    public ToolStrip Strip { get; }
    public ToolStripButton Header { get; }

    public RibbonTab(string name, string text, ToolStrip strip, ToolStripButton header)
    {
        Name = name; Text = text; Strip = strip; Header = header;
    }

    public ToolStripButton Button(string name, string text, string glyph, EventHandler onClick, string? tooltip = null, bool mirror = false)
    {
        var b = new ToolStripButton(text, Icons.Large(glyph, null, mirror))
        {
            Name = name,
            TextImageRelation = TextImageRelation.ImageAboveText,
            ImageScaling = ToolStripItemImageScaling.None,
            AutoSize = true,
            Padding = new Padding(Dpi.Scale(4), Dpi.Scale(2), Dpi.Scale(4), Dpi.Scale(2)),
            ToolTipText = tooltip ?? text.Replace("\n", " ").Replace("&&", "&"),
            AccessibleName = text.Replace("\n", " ").Replace("&&", "&"),
            Margin = new Padding(Dpi.Scale(1), Dpi.Scale(2), Dpi.Scale(1), Dpi.Scale(2))
        };
        b.Click += onClick;
        Strip.Items.Add(b);
        return b;
    }

    public ToolStripDropDownButton DropDown(string name, string text, string glyph, string? tooltip = null)
    {
        var b = new ToolStripDropDownButton(text, Icons.Large(glyph))
        {
            Name = name,
            TextImageRelation = TextImageRelation.ImageAboveText,
            ImageScaling = ToolStripItemImageScaling.None,
            ToolTipText = tooltip ?? text,
            AccessibleName = text.Replace("\n", " ").Replace("&&", "&"),
            Padding = new Padding(Dpi.Scale(4), Dpi.Scale(2), Dpi.Scale(4), Dpi.Scale(2))
        };
        Strip.Items.Add(b);
        return b;
    }

    public void Separator() => Strip.Items.Add(new ToolStripSeparator { Margin = new Padding(Dpi.Scale(4), Dpi.Scale(6), Dpi.Scale(4), Dpi.Scale(6)) });

    public T Add<T>(T item) where T : ToolStripItem
    {
        Strip.Items.Add(item);
        return item;
    }
}

public sealed class RibbonRenderer : ToolStripProfessionalRenderer
{
    private readonly bool _header;

    public RibbonRenderer(bool header) : base(new RibbonColors()) { _header = header; RoundedEdges = false; }

    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
    {
        if (!_header)
        {
            using var p = new Pen(Theme.Border);
            e.Graphics.DrawLine(p, 0, e.AffectedBounds.Bottom - 1, e.AffectedBounds.Right, e.AffectedBounds.Bottom - 1);
        }
    }

    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
    {
        using var b = new SolidBrush(Theme.RibbonBack);
        e.Graphics.FillRectangle(b, e.AffectedBounds);
    }

    protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
    {
        var item = e.Item;
        var g = e.Graphics;
        var r = new Rectangle(0, 0, item.Width - 1, item.Height - 1);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        if (_header && item.Name == "btnFileMenu")
        {
            using var fb = new SolidBrush(item.Selected || item.Pressed ? Theme.AccentDark : Theme.Accent);
            g.FillRectangle(fb, new Rectangle(0, 0, item.Width, item.Height));
            return;
        }
        if (_header && item is ToolStripButton hb)
        {
            if (hb.Checked)
            {
                using var accent = new SolidBrush(Theme.Accent);
                g.FillRectangle(accent, Dpi.Scale(6), item.Height - Dpi.Scale(3), item.Width - Dpi.Scale(12), Dpi.Scale(3));
            }
            else if (item.Selected)
            {
                using var hover = new SolidBrush(Theme.Hover);
                g.FillRectangle(hover, r);
            }
            return;
        }
        bool isChecked = item is ToolStripButton { Checked: true };
        if (item.Pressed || isChecked || item.Selected)
        {
            var color = item.Pressed || isChecked ? Theme.Pressed : Theme.Hover;
            using var path = Rounded(r, Dpi.Scale(4));
            using var b = new SolidBrush(color);
            g.FillPath(b, path);
            if (isChecked)
            {
                using var pen = new Pen(Theme.PressedBorder);
                g.DrawPath(pen, path);
            }
        }
    }

    protected override void OnRenderDropDownButtonBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item.Name == "btnFileMenu")
        {
            using var b = new SolidBrush(e.Item.Selected || e.Item.Pressed ? Theme.AccentDark : Theme.Accent);
            e.Graphics.FillRectangle(b, new Rectangle(0, 0, e.Item.Width, e.Item.Height));
            return;
        }
        OnRenderButtonBackground(e);
    }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    {
        if (_header && e.Item is ToolStripButton { Checked: true })
        {
            e.TextColor = Theme.AccentDark;
            e.TextFont = new Font(e.TextFont ?? Theme.UiFont, FontStyle.Bold);
        }
        else if (e.Item.Name != "btnFileMenu" && e.Item.Enabled) e.TextColor = Theme.Ink;
        base.OnRenderItemText(e);
    }

    private static GraphicsPath Rounded(Rectangle r, int rad)
    {
        var p = new GraphicsPath();
        int d = rad * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}

public sealed class RibbonColors : ProfessionalColorTable
{
    public override Color MenuItemSelected => Theme.Hover;
    public override Color MenuItemSelectedGradientBegin => Theme.Hover;
    public override Color MenuItemSelectedGradientEnd => Theme.Hover;
    public override Color MenuItemBorder => Theme.PressedBorder;
    public override Color MenuBorder => Theme.Border;
    public override Color ToolStripDropDownBackground => Color.White;
    public override Color ImageMarginGradientBegin => Color.White;
    public override Color ImageMarginGradientMiddle => Color.White;
    public override Color ImageMarginGradientEnd => Color.White;
    public override Color SeparatorDark => Theme.Border;
    public override Color SeparatorLight => Color.White;
    public override Color ButtonSelectedHighlight => Theme.Hover;
    public override Color ButtonCheckedHighlight => Theme.Pressed;
    public override Color StatusStripGradientBegin => Theme.TabStripBack;
    public override Color StatusStripGradientEnd => Theme.TabStripBack;
}
