using MicroadsOffice.Core.Office;
using System.Text.Json.Serialization;

namespace MicroadsOffice.Core.Signing;

public enum SignatureGraphic
{
    None,
    Image,      // imported picture: company seal / scanned signature (PNG with transparency recommended)
    Name,       // signer name rendered as large graphic text
    Drawn       // hand-drawn signature (stored as PNG image)
}

public enum SignatureLayout
{
    GraphicLeft,        // graphic on the left, text on the right (Foxit default)
    GraphicTop,         // graphic on top, text below
    GraphicBackground,  // seal behind the text (common Vietnamese style: text over the red seal)
    GraphicRight
}

public enum SignatureStandard
{
    PAdES,   // /SubFilter /ETSI.CAdES.detached (ETSI EN 319 142)
    Pkcs7    // /SubFilter /adbe.pkcs7.detached
}

/// <summary>A reusable signature appearance style (like Foxit's "Signature Style").</summary>
public sealed class SignatureStyle
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Title { get; set; } = Loc.T("Mẫu chữ ký", "Signature style");

    public SignatureGraphic Graphic { get; set; } = SignatureGraphic.None;
    /// <summary>File name of the graphic image inside the styles folder (for Image / Drawn).</summary>
    public string? ImageFile { get; set; }
    public SignatureLayout Layout { get; set; } = SignatureLayout.GraphicLeft;
    public float ImageOpacity { get; set; } = 1f;

    public bool ShowName { get; set; } = true;
    public bool ShowDate { get; set; } = true;
    public bool ShowReason { get; set; } = true;
    public bool ShowLocation { get; set; } = true;
    public bool ShowDistinguishedName { get; set; }
    public bool ShowContact { get; set; }
    public bool ShowPlatform { get; set; } = true;
    public bool ShowLabels { get; set; } = true;
    public string CustomText { get; set; } = "";

    public string FontFamily { get; set; } = "Arial";
    public bool Bold { get; set; }
    /// <summary>0 = automatic (fit to the box).</summary>
    public float FontSize { get; set; }
    public int TextColorArgb { get; set; } = unchecked((int)0xFF1F2A44);
    public bool Border { get; set; }
    public int BorderColorArgb { get; set; } = unchecked((int)0xFF2B579A);
    public string DateFormat { get; set; } = "dd/MM/yyyy HH:mm:ss";

    public SignatureStyle Clone()
    {
        var c = (SignatureStyle)MemberwiseClone();
        return c;
    }

    public static SignatureStyle CreateDefault() => new() { Id = "default", Title = Loc.T("Mặc định (chỉ thông tin)", "Default (details only)") };
}

/// <summary>Information printed inside the visible signature.</summary>
public sealed class SignatureAppearanceInfo
{
    public string SignerName { get; set; } = "";
    public string DistinguishedName { get; set; } = "";
    public DateTimeOffset SigningTime { get; set; } = DateTimeOffset.Now;
    public string Reason { get; set; } = "";
    public string Location { get; set; } = "";
    public string Contact { get; set; } = "";
    [JsonIgnore] public System.Drawing.Image? Graphic { get; set; }
}
