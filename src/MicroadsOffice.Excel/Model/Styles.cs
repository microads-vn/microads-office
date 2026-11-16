using System.Drawing;

namespace MicroadsOffice.Excel.Model;

public enum HAlign : byte { General, Left, Center, Right, Justify, CenterAcross }
public enum VAlign : byte { Bottom, Center, Top }
public enum LineStyle : byte { None, Thin, Medium, Thick, Dashed, Dotted, Double, Hair }

public readonly record struct BorderLine(LineStyle Style, int Argb)
{
    public static readonly BorderLine None = new(LineStyle.None, 0);
    public static BorderLine Thin(Color? c = null) => new(LineStyle.Thin, (c ?? Color.Black).ToArgb());
    public bool IsNone => Style == LineStyle.None;
    public Color Color => Color.FromArgb(Argb);
}

/// <summary>Immutable cell format. Equal formats share one entry in <see cref="StyleTable"/>.</summary>
public sealed record CellStyle
{
    public const string DefaultFont = "Calibri";
    public const float DefaultSize = 11f;

    public string FontName { get; init; } = DefaultFont;
    public float FontSize { get; init; } = DefaultSize;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    /// <summary>ARGB; 0 = automatic (black).</summary>
    public int FontColor { get; init; }
    /// <summary>ARGB; 0 = no fill.</summary>
    public int Fill { get; init; }
    public HAlign HAlign { get; init; }
    public VAlign VAlign { get; init; }
    public bool Wrap { get; init; }
    public int Indent { get; init; }
    public string NumberFormat { get; init; } = "General";
    public BorderLine Left { get; init; }
    public BorderLine Right { get; init; }
    public BorderLine Top { get; init; }
    public BorderLine Bottom { get; init; }
    public bool Locked { get; init; } = true;
    public bool HiddenFormula { get; init; }

    public static readonly CellStyle Default = new();

    public Color FontColorValue => FontColor == 0 ? Color.Black : Color.FromArgb(FontColor);
    public bool HasFill => Fill != 0;
}

public sealed class StyleTable
{
    private readonly List<CellStyle> _list = new() { CellStyle.Default };
    private readonly Dictionary<CellStyle, int> _index = new() { [CellStyle.Default] = 0 };

    public CellStyle this[int id] => id >= 0 && id < _list.Count ? _list[id] : CellStyle.Default;
    public int Count => _list.Count;

    public int Intern(CellStyle s)
    {
        if (_index.TryGetValue(s, out var id)) return id;
        _list.Add(s);
        _index[s] = _list.Count - 1;
        return _list.Count - 1;
    }

    public IReadOnlyList<CellStyle> All => _list;

    public StyleTable Clone()
    {
        var t = new StyleTable();
        for (int i = 1; i < _list.Count; i++) t.Intern(_list[i]);
        return t;
    }
}
