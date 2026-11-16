using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using MicroadsOffice.Core.Pdf;
using MicroadsOffice.Core.Office;

namespace MicroadsOffice.Core.Signing;

/// <summary>
/// Lays out a visible signature (graphic + text) and renders it either as a PDF Form XObject
/// (text converted to vector outlines so Vietnamese diacritics are always correct without font embedding)
/// or as a GDI+ bitmap for on-screen preview. Both outputs come from the same layout, so the preview is exact.
/// </summary>
public static class AppearanceRenderer
{
    public const string PlatformName = "Microads Office";
    public const string PlatformVendor = "Microads.vn";
    public static string PlatformVersion { get; set; } = "1.0.0";
    public static string PlatformFullName => $"{PlatformName} ({PlatformVendor})";

    private abstract record Op;
    private sealed record ImageOp(RectangleF Rect, Image Image, float Opacity) : Op;
    private sealed record PathOp(GraphicsPath Path, Color Color) : Op;
    private sealed record BorderOp(RectangleF Rect, Color Color, float Width) : Op;

    /// <summary>Builds the text lines shown in the signature according to the style.</summary>
    public static List<string> BuildLines(SignatureStyle s, SignatureAppearanceInfo info)
    {
        var lines = new List<string>();
        string L(string label, string value) => s.ShowLabels ? $"{label}: {value}" : value;
        if (s.ShowName && !string.IsNullOrWhiteSpace(info.SignerName)) lines.Add(L(Loc.T("Ký bởi", "Signed by"), info.SignerName));
        if (s.ShowDate)
        {
            var fmt = string.IsNullOrWhiteSpace(s.DateFormat) ? "dd/MM/yyyy HH:mm:ss" : s.DateFormat;
            string date;
            try { date = info.SigningTime.ToString(fmt, CultureInfo.InvariantCulture); }
            catch (FormatException) { date = info.SigningTime.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture); }
            var off = info.SigningTime.Offset;
            date += $" {(off < TimeSpan.Zero ? "-" : "+")}{off:hh\\:mm}";
            lines.Add(L(Loc.T("Ngày ký", "Date"), date));
        }
        if (s.ShowReason && !string.IsNullOrWhiteSpace(info.Reason)) lines.Add(L(Loc.T("Lý do", "Reason"), info.Reason));
        if (s.ShowLocation && !string.IsNullOrWhiteSpace(info.Location)) lines.Add(L(Loc.T("Vị trí", "Location"), info.Location));
        if (s.ShowContact && !string.IsNullOrWhiteSpace(info.Contact)) lines.Add(L(Loc.T("Liên hệ", "Contact"), info.Contact));
        if (s.ShowDistinguishedName && !string.IsNullOrWhiteSpace(info.DistinguishedName)) lines.Add(L("DN", info.DistinguishedName));
        if (!string.IsNullOrWhiteSpace(s.CustomText))
            lines.AddRange(s.CustomText.Replace("\r", "").Split('\n').Where(x => x.Length > 0));
        if (s.ShowPlatform) lines.Add(L(Loc.T("Nền tảng", "Platform"), $"{PlatformFullName} v{PlatformVersion}"));
        return lines;
    }

    // ------------------------------------------------------------------ layout

    private static List<Op> Layout(SignatureStyle style, SignatureAppearanceInfo info, float w, float h)
    {
        var ops = new List<Op>();
        float pad = Math.Clamp(Math.Min(w, h) * 0.04f, 1f, 4f);
        var full = new RectangleF(pad, pad, Math.Max(1, w - 2 * pad), Math.Max(1, h - 2 * pad));
        var lines = BuildLines(style, info);

        bool hasGraphic = style.Graphic switch
        {
            SignatureGraphic.Image or SignatureGraphic.Drawn => info.Graphic != null,
            SignatureGraphic.Name => !string.IsNullOrWhiteSpace(info.SignerName),
            _ => false
        };
        bool hasText = lines.Count > 0;

        RectangleF gRect = full, tRect = full;
        if (hasGraphic && hasText)
        {
            switch (style.Layout)
            {
                case SignatureLayout.GraphicLeft:
                    gRect = new RectangleF(full.X, full.Y, full.Width * 0.42f, full.Height);
                    tRect = RectangleF.FromLTRB(gRect.Right + pad, full.Y, full.Right, full.Bottom);
                    break;
                case SignatureLayout.GraphicRight:
                    tRect = new RectangleF(full.X, full.Y, full.Width * 0.58f - pad, full.Height);
                    gRect = RectangleF.FromLTRB(tRect.Right + pad, full.Y, full.Right, full.Bottom);
                    break;
                case SignatureLayout.GraphicTop:
                    gRect = new RectangleF(full.X, full.Y, full.Width, full.Height * 0.5f);
                    tRect = RectangleF.FromLTRB(full.X, gRect.Bottom + pad * 0.5f, full.Right, full.Bottom);
                    break;
                case SignatureLayout.GraphicBackground:
                    gRect = full; tRect = full;
                    break;
            }
        }

        if (hasGraphic)
        {
            if (style.Graphic == SignatureGraphic.Name)
            {
                var nameOps = LayoutText([info.SignerName], gRect, style.FontFamily.Length > 0 ? PickScriptFont() : "Arial",
                    FontStyle.Italic | FontStyle.Bold, 0, Color.FromArgb(style.TextColorArgb), center: true);
                ops.AddRange(nameOps);
            }
            else if (info.Graphic != null)
            {
                var img = info.Graphic;
                float ar = (float)img.Width / Math.Max(1, img.Height);
                float gw = gRect.Width, gh = gRect.Width / ar;
                if (gh > gRect.Height) { gh = gRect.Height; gw = gh * ar; }
                var r = new RectangleF(gRect.X + (gRect.Width - gw) / 2, gRect.Y + (gRect.Height - gh) / 2, gw, gh);
                ops.Add(new ImageOp(r, img, Math.Clamp(style.ImageOpacity, 0.05f, 1f)));
            }
        }
        if (hasText)
        {
            var fs = style.Bold ? FontStyle.Bold : FontStyle.Regular;
            ops.AddRange(LayoutText(lines, tRect, string.IsNullOrWhiteSpace(style.FontFamily) ? "Arial" : style.FontFamily,
                fs, style.FontSize, Color.FromArgb(style.TextColorArgb), center: false));
        }
        if (style.Border)
            ops.Add(new BorderOp(new RectangleF(0.5f, 0.5f, w - 1, h - 1), Color.FromArgb(style.BorderColorArgb), 1f));
        return ops;
    }

    private static string? _scriptFont;
    private static string PickScriptFont()
    {
        if (_scriptFont != null) return _scriptFont;
        using var fonts = new InstalledFontCollection();
        var families = fonts.Families;
        try
        {
            foreach (var name in new[] { "Segoe Script", "Lucida Handwriting", "Segoe Print", "Times New Roman" })
                if (families.Any(f => f.Name == name)) return _scriptFont = name;
            return _scriptFont = "Arial";
        }
        finally { foreach (var f in families) f.Dispose(); }
    }

    private static FontFamily ResolveFamily(string name)
    {
        try { return new FontFamily(name); }
        catch { return new FontFamily("Arial"); }
    }

    private static float MeasureWidth(string text, FontFamily fam, FontStyle style, float size)
    {
        if (string.IsNullOrEmpty(text)) return 0;
        using var p = new GraphicsPath();
        p.AddString(text, fam, (int)style, size, PointF.Empty, StringFormat.GenericTypographic);
        var b = p.GetBounds();
        return b.Width <= 0 ? 0 : b.Right;
    }

    private static List<Op> LayoutText(IReadOnlyList<string> lines, RectangleF area, string familyName, FontStyle style,
        float fixedSize, Color color, bool center)
    {
        var ops = new List<Op>();
        using var fam = ResolveFamily(familyName);
        if (!fam.IsStyleAvailable(style)) style = FontStyle.Regular;
        const float refSize = 10f;
        float lineFactor = (float)fam.GetLineSpacing(style) / fam.GetEmHeight(style);
        var wordWidth = new Dictionary<string, float>();
        float W(string word)
        {
            if (!wordWidth.TryGetValue(word, out var ww)) wordWidth[word] = ww = MeasureWidth(word, fam, style, refSize);
            return ww;
        }
        float space = Math.Max(MeasureWidth("a a", fam, style, refSize) - MeasureWidth("aa", fam, style, refSize), refSize * 0.25f);

        List<string> Wrap(float size)
        {
            float k = size / refSize, maxW = area.Width;
            var result = new List<string>();
            foreach (var line in lines)
            {
                var words = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var cur = new StringBuilder();
                float curW = 0;
                foreach (var word in words)
                {
                    float ww = W(word) * k;
                    if (ww > maxW)
                    {
                        // break long word into chunks
                        if (cur.Length > 0) { result.Add(cur.ToString()); cur.Clear(); curW = 0; }
                        var chunk = new StringBuilder();
                        foreach (var ch in word)
                        {
                            if (chunk.Length > 0 && MeasureWidth(chunk.ToString() + ch, fam, style, size) > maxW)
                            { result.Add(chunk.ToString()); chunk.Clear(); }
                            chunk.Append(ch);
                        }
                        cur.Append(chunk); curW = MeasureWidth(chunk.ToString(), fam, style, size);
                        continue;
                    }
                    float add = (cur.Length > 0 ? space * k : 0) + ww;
                    if (cur.Length > 0 && curW + add > maxW)
                    {
                        result.Add(cur.ToString()); cur.Clear(); curW = 0; add = ww;
                    }
                    if (cur.Length > 0) cur.Append(' ');
                    cur.Append(word);
                    curW += add;
                }
                if (cur.Length > 0) result.Add(cur.ToString());
            }
            return result;
        }

        float size;
        List<string> wrapped;
        if (fixedSize > 0)
        {
            size = fixedSize;
            wrapped = Wrap(size);
        }
        else
        {
            float lo = 2f, hi = Math.Min(center ? 200f : 28f, area.Height);
            for (int i = 0; i < 18; i++)
            {
                float mid = (lo + hi) / 2;
                var wl = Wrap(mid);
                if (wl.Count * mid * lineFactor <= area.Height + 0.01f) lo = mid; else hi = mid;
            }
            size = lo;
            wrapped = Wrap(size);
        }

        float lineH = size * lineFactor;
        float total = wrapped.Count * lineH;
        float y = area.Y + Math.Max(0, (area.Height - total) / 2);
        foreach (var line in wrapped)
        {
            var path = new GraphicsPath(FillMode.Winding);
            path.AddString(line, fam, (int)style, size, PointF.Empty, StringFormat.GenericTypographic);
            var b = path.GetBounds();
            float x = area.X;
            if (center) x = area.X + (area.Width - (b.Right)) / 2;
            using (var m = new Matrix()) { m.Translate(x, y); path.Transform(m); }
            ops.Add(new PathOp(path, color));
            y += lineH;
        }
        return ops;
    }

    // ------------------------------------------------------------------ bitmap preview

    /// <summary>Renders the appearance into a bitmap (w,h are in PDF points).</summary>
    public static Bitmap RenderBitmap(SignatureStyle style, SignatureAppearanceInfo info, float w, float h, float scale, bool transparent = true)
    {
        int pw = Math.Max(1, (int)Math.Ceiling(w * scale)), ph = Math.Max(1, (int)Math.Ceiling(h * scale));
        var bmp = new Bitmap(pw, ph, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(transparent ? Color.Transparent : Color.White);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.ScaleTransform(scale, scale);
        foreach (var op in Layout(style, info, w, h))
        {
            switch (op)
            {
                case ImageOp io:
                {
                    using var ia = new ImageAttributes();
                    var cm = new ColorMatrix { Matrix33 = io.Opacity };
                    ia.SetColorMatrix(cm);
                    var r = io.Rect;
                    g.DrawImage(io.Image, new[] { new PointF(r.Left, r.Top), new PointF(r.Right, r.Top), new PointF(r.Left, r.Bottom) },
                        new RectangleF(0, 0, io.Image.Width, io.Image.Height), GraphicsUnit.Pixel, ia);
                    break;
                }
                case PathOp po:
                {
                    using var br = new SolidBrush(po.Color);
                    g.FillPath(br, po.Path);
                    po.Path.Dispose();
                    break;
                }
                case BorderOp bo:
                {
                    using var pen = new Pen(bo.Color, bo.Width);
                    g.DrawRectangle(pen, bo.Rect.X, bo.Rect.Y, bo.Rect.Width, bo.Rect.Height);
                    break;
                }
            }
        }
        return bmp;
    }

    // ------------------------------------------------------------------ PDF form XObject

    /// <summary>
    /// Creates the Form XObject (plus image XObjects) for the appearance and registers the objects in the update.
    /// w,h are the visual (upright) dimensions; pageRotation is the page /Rotate value in degrees.
    /// </summary>
    public static PdfReference BuildFormXObject(IncrementalUpdate upd, SignatureStyle style, SignatureAppearanceInfo info,
        float w, float h, int pageRotation)
    {
        var ops = Layout(style, info, w, h);
        var content = new StringBuilder();
        var xobjects = new PdfDictionary();
        var extg = new PdfDictionary();
        int imgIndex = 0, gsIndex = 0;
        string F(double v) => PdfNumber.Format(Math.Round(v, 3));
        foreach (var op in ops)
        {
            switch (op)
            {
                case ImageOp io:
                {
                    var imgRef = AddImage(upd, io.Image);
                    string name = "Img" + imgIndex++;
                    xobjects[name] = imgRef;
                    content.Append("q\n");
                    if (io.Opacity < 0.999f)
                    {
                        string gs = "GS" + gsIndex++;
                        extg[gs] = new PdfDictionary().Set("Type", "ExtGState").Set("ca", io.Opacity).Set("CA", io.Opacity);
                        content.Append('/').Append(gs).Append(" gs\n");
                    }
                    var r = io.Rect;
                    content.Append($"{F(r.Width)} 0 0 {F(r.Height)} {F(r.X)} {F(h - r.Bottom)} cm /{name} Do\nQ\n");
                    break;
                }
                case PathOp po:
                {
                    content.Append($"q {F(po.Color.R / 255.0)} {F(po.Color.G / 255.0)} {F(po.Color.B / 255.0)} rg\n");
                    AppendPath(content, po.Path, h, F);
                    content.Append("f\nQ\n");
                    po.Path.Dispose();
                    break;
                }
                case BorderOp bo:
                {
                    var r = bo.Rect;
                    content.Append($"q {F(bo.Color.R / 255.0)} {F(bo.Color.G / 255.0)} {F(bo.Color.B / 255.0)} RG {F(bo.Width)} w " +
                                   $"{F(r.X)} {F(h - r.Bottom)} {F(r.Width)} {F(r.Height)} re S Q\n");
                    break;
                }
            }
        }
        var resources = new PdfDictionary();
        if (xobjects.Count > 0) resources["XObject"] = xobjects;
        if (extg.Count > 0) resources["ExtGState"] = extg;
        resources["ProcSet"] = new PdfArray([new PdfName("PDF"), new PdfName("ImageC")]);
        var dict = new PdfDictionary()
            .Set("Type", "XObject").Set("Subtype", "Form")
            .Set("BBox", PdfArray.Of(0, 0, w, h))
            .Set("Resources", resources)
            .Set("Filter", "FlateDecode");
        var matrix = RotationMatrix(pageRotation, w, h);
        if (matrix != null) dict["Matrix"] = matrix;
        return upd.Add(new PdfStream(dict, Deflate(Encoding.ASCII.GetBytes(content.ToString()))));
    }

    /// <summary>Matrix that keeps the appearance upright on rotated pages.</summary>
    public static PdfArray? RotationMatrix(int pageRotation, float w, float h) => (((pageRotation % 360) + 360) % 360) switch
    {
        90 => PdfArray.Of(0, 1, -1, 0, h, 0),
        180 => PdfArray.Of(-1, 0, 0, -1, w, h),
        270 => PdfArray.Of(0, -1, 1, 0, 0, w),
        _ => null
    };

    private static void AppendPath(StringBuilder sb, GraphicsPath path, float h, Func<double, string> F)
    {
        if (path.PointCount == 0) return;
        var pts = path.PathPoints;
        var types = path.PathTypes;
        int i = 0;
        while (i < pts.Length)
        {
            int t = types[i] & 0x07;
            if (t == 0)
            {
                sb.Append(F(pts[i].X)).Append(' ').Append(F(h - pts[i].Y)).Append(" m\n");
            }
            else if (t == 1)
            {
                sb.Append(F(pts[i].X)).Append(' ').Append(F(h - pts[i].Y)).Append(" l\n");
            }
            else if (t == 3 && i + 2 < pts.Length)
            {
                sb.Append(F(pts[i].X)).Append(' ').Append(F(h - pts[i].Y)).Append(' ')
                  .Append(F(pts[i + 1].X)).Append(' ').Append(F(h - pts[i + 1].Y)).Append(' ')
                  .Append(F(pts[i + 2].X)).Append(' ').Append(F(h - pts[i + 2].Y)).Append(" c\n");
                if ((types[i + 2] & 0x80) != 0) sb.Append("h\n");
                i += 3;
                continue;
            }
            if ((types[i] & 0x80) != 0) sb.Append("h\n");
            i++;
        }
    }

    public static byte[] Deflate(byte[] data)
    {
        using var ms = new MemoryStream();
        using (var z = new ZLibStream(ms, CompressionLevel.Optimal, leaveOpen: true)) z.Write(data, 0, data.Length);
        return ms.ToArray();
    }

    /// <summary>Adds an RGB image XObject (with /SMask for transparency) to the update.</summary>
    public static PdfReference AddImage(IncrementalUpdate upd, Image source, int maxSide = 1600)
    {
        int w = source.Width, h = source.Height;
        float k = Math.Min(1f, (float)maxSide / Math.Max(w, h));
        w = Math.Max(1, (int)(w * k)); h = Math.Max(1, (int)(h * k));
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Transparent);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawImage(source, new Rectangle(0, 0, w, h));
        }
        var bd = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        var raw = new byte[bd.Stride * h];
        System.Runtime.InteropServices.Marshal.Copy(bd.Scan0, raw, 0, raw.Length);
        bmp.UnlockBits(bd);
        var rgb = new byte[w * h * 3];
        var alpha = new byte[w * h];
        bool hasAlpha = false;
        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                int s = y * bd.Stride + x * 4, p = y * w + x;
                byte a = raw[s + 3];
                // un-premultiply is not needed (Format32bppArgb is straight alpha)
                rgb[p * 3] = raw[s + 2]; rgb[p * 3 + 1] = raw[s + 1]; rgb[p * 3 + 2] = raw[s];
                alpha[p] = a;
                if (a != 255) hasAlpha = true;
            }
        }
        var dict = new PdfDictionary()
            .Set("Type", "XObject").Set("Subtype", "Image")
            .Set("Width", w).Set("Height", h)
            .Set("ColorSpace", "DeviceRGB").Set("BitsPerComponent", 8)
            .Set("Filter", "FlateDecode");
        if (hasAlpha)
        {
            var smask = new PdfDictionary()
                .Set("Type", "XObject").Set("Subtype", "Image")
                .Set("Width", w).Set("Height", h)
                .Set("ColorSpace", "DeviceGray").Set("BitsPerComponent", 8)
                .Set("Filter", "FlateDecode");
            dict["SMask"] = upd.Add(new PdfStream(smask, Deflate(alpha)));
        }
        return upd.Add(new PdfStream(dict, Deflate(rgb)));
    }
}
