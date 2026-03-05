using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class Presentation
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public int SlideWidth { get; }
    public int SlideHeight { get; }
    public List<string> SlideIds { get; } = new();

    public Presentation(XElement element)
    {
        _element = element;

        var sldSz = element.Descendants(P + "sldSz").FirstOrDefault();
        if (sldSz != null)
        {
            SlideWidth = int.TryParse(sldSz.Attribute("cx")?.Value, out var w) ? w : 9144000;
            SlideHeight = int.TryParse(sldSz.Attribute("cy")?.Value, out var h) ? h : 6858000;
        }
        else
        {
            SlideWidth = 9144000;
            SlideHeight = 6858000;
        }

        var sldIdLst = element.Descendants(P + "sldIdLst").FirstOrDefault();
        if (sldIdLst != null)
        {
            foreach (var sldId in sldIdLst.Elements(P + "sldId"))
            {
                var id = sldId.Attribute(R + "id")?.Value;
                if (id != null)
                    SlideIds.Add(id);
            }
        }
    }
}
