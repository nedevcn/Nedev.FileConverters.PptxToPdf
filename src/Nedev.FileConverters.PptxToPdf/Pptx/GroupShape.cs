using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class GroupShape
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? Id { get; }
    public string? Name { get; }
    public Rect Bounds { get; }
    public List<Shape> Shapes { get; } = new();
    public List<Picture> Pictures { get; } = new();
    public List<GroupShape> ChildGroups { get; } = new();
    public List<Table> Tables { get; } = new();

    public GroupShape(XElement element)
    {
        _element = element;

        var nvGrpSpPr = element.Element(P + "nvGrpSpPr");
        if (nvGrpSpPr != null)
        {
            var cNvPr = nvGrpSpPr.Element(P + "cNvPr");
            Id = cNvPr?.Attribute("id")?.Value;
            Name = cNvPr?.Attribute("name")?.Value;
        }

        var grpSpPr = element.Element(P + "grpSpPr");
        if (grpSpPr != null)
        {
            Bounds = ParseBounds(grpSpPr);
        }

        var spTree = element.Element(P + "spTree");
        if (spTree != null)
        {
            ParseShapeTree(spTree);
        }
        else
        {
            // Direct children
            ParseShapeTree(element);
        }
    }

    private static Rect ParseBounds(XElement grpSpPr)
    {
        var xfrm = grpSpPr.Element(A + "xfrm");
        if (xfrm == null) return new Rect();

        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");
        var chOff = xfrm.Element(A + "chOff");
        var chExt = xfrm.Element(A + "chExt");

        if (off == null || ext == null) return new Rect();

        long.TryParse(off.Attribute("x")?.Value, out var x);
        long.TryParse(off.Attribute("y")?.Value, out var y);
        long.TryParse(ext.Attribute("cx")?.Value, out var w);
        long.TryParse(ext.Attribute("cy")?.Value, out var h);

        return new Rect(x, y, w, h);
    }

    private void ParseShapeTree(XElement parent)
    {
        foreach (var sp in parent.Elements(P + "sp"))
        {
            var shape = new Shape(sp);
            if (!shape.IsPlaceholder || shape.HasText)
            {
                Shapes.Add(shape);
            }
        }

        foreach (var pic in parent.Elements(P + "pic"))
        {
            Pictures.Add(new Picture(pic));
        }

        foreach (var grpSp in parent.Elements(P + "grpSp"))
        {
            ChildGroups.Add(new GroupShape(grpSp));
        }

        foreach (var graphicFrame in parent.Elements(P + "graphicFrame"))
        {
            // Check if it's a table
            var tbl = graphicFrame.Element(A + "graphic")?.
                                  Element(A + "graphicData")?.
                                  Element(A + "tbl");
            if (tbl != null)
            {
                Tables.Add(new Table(graphicFrame));
            }
        }
    }
}
