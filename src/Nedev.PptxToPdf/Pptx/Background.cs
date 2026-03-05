using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class Background
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public Color? Color { get; }
    public FillType FillType { get; }

    public Background(XElement element)
    {
        _element = element;

        var solidFill = element.Descendants(A + "solidFill").FirstOrDefault();
        if (solidFill != null)
        {
            var srgbClr = solidFill.Element(A + "srgbClr");
            if (srgbClr != null)
            {
                var val = srgbClr.Attribute("val")?.Value;
                if (val != null && val.Length == 6)
                {
                    if (byte.TryParse(val.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                        byte.TryParse(val.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                        byte.TryParse(val.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
                    {
                        Color = new Color(r, g, b);
                        FillType = FillType.Solid;
                    }
                }
            }
        }
    }
}
