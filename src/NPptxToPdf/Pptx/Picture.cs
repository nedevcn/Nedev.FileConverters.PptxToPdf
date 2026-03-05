using System.Xml.Linq;

namespace NPptxToPdf.Pptx;

public class Picture
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public string? Id { get; }
    public string? Name { get; }
    public Rect Bounds { get; }
    public string? ImageRelationshipId { get; }

    public Picture(XElement element)
    {
        _element = element;

        var nvPicPr = element.Element(P + "nvPicPr");
        if (nvPicPr != null)
        {
            var cNvPr = nvPicPr.Element(P + "cNvPr");
            Id = cNvPr?.Attribute("id")?.Value;
            Name = cNvPr?.Attribute("name")?.Value;
        }

        var blipFill = element.Element(P + "blipFill");
        if (blipFill != null)
        {
            var blip = blipFill.Element(A + "blip");
            ImageRelationshipId = blip?.Attribute(R + "embed")?.Value;
        }

        var spPr = element.Element(P + "spPr");
        if (spPr != null)
        {
            Bounds = ParseBounds(spPr);
        }
        else
        {
            Bounds = new Rect();
        }
    }

    private static Rect ParseBounds(XElement spPr)
    {
        var xfrm = spPr.Element(A + "xfrm");
        if (xfrm == null) return new Rect();

        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");

        if (off == null || ext == null) return new Rect();

        return new Rect
        {
            X = long.TryParse(off.Attribute("x")?.Value, out var x) ? x : 0,
            Y = long.TryParse(off.Attribute("y")?.Value, out var y) ? y : 0,
            Width = long.TryParse(ext.Attribute("cx")?.Value, out var w) ? w : 0,
            Height = long.TryParse(ext.Attribute("cy")?.Value, out var h) ? h : 0
        };
    }
}
