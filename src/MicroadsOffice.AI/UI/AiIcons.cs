using System.Drawing.Drawing2D;
using MicroadsOffice.UI;

namespace MicroadsOffice.AI.UI;

/// <summary>The assistant's ribbon icon: a four-point sparkle with a small companion star (drawn in code, cached).</summary>
public static class AiIcons
{
    public static readonly Color Ink = Color.FromArgb(124, 58, 237);

    public static Image Sparkle(bool large) => Icons.Painted("ai-sparkle", Dpi.Scale(large ? 32 : 16), Ink, (g, s, c) =>
    {
        using var b = new LinearGradientBrush(new RectangleF(0, 0, s, s), Color.FromArgb(37, 99, 235), Color.FromArgb(168, 85, 247), 45f);
        Star(g, b, s * 0.42f, s * 0.56f, s * 0.40f);
        Star(g, b, s * 0.78f, s * 0.22f, s * 0.18f);
    });

    static void Star(Graphics g, Brush b, float cx, float cy, float r)
    {
        float k = r * 0.28f;
        using var p = new GraphicsPath();
        p.AddPolygon([
            new PointF(cx, cy - r), new PointF(cx + k, cy - k), new PointF(cx + r, cy), new PointF(cx + k, cy + k),
            new PointF(cx, cy + r), new PointF(cx - k, cy + k), new PointF(cx - r, cy), new PointF(cx - k, cy - k),
        ]);
        g.FillPath(b, p);
    }
}
