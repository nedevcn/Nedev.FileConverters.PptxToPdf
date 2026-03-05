using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class Slide
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public string Id { get; }
    public List<Shape> Shapes { get; } = new();
    public List<Picture> Pictures { get; } = new();
    public List<GroupShape> GroupShapes { get; } = new();
    public List<Table> Tables { get; } = new();
    public List<Connector> Connectors { get; } = new();
    public List<Chart> Charts { get; } = new();
    public List<SmartArt> SmartArts { get; } = new();
    public Background? Background { get; private set; }
    public SlideLayout? Layout { get; set; }
    public SlideTransition? Transition { get; private set; }
    public SlideTiming? Timing { get; private set; }

    public Slide(XElement element, string id)
    {
        _element = element;
        Id = id;
        Parse();
    }

    private void Parse()
    {
        var cSld = _element.Element(P + "cSld");
        if (cSld == null) return;

        // Parse background
        var bg = cSld.Element(P + "bg");
        if (bg != null)
        {
            Background = new Background(bg);
        }

        // Parse shape tree
        var spTree = cSld.Element(P + "spTree");
        if (spTree == null) return;

        // Parse shapes
        foreach (var sp in spTree.Elements(P + "sp"))
        {
            var shape = new Shape(sp);
            if (!shape.IsPlaceholder || shape.HasText)
            {
                Shapes.Add(shape);
            }
        }

        // Parse pictures
        foreach (var pic in spTree.Elements(P + "pic"))
        {
            Pictures.Add(new Picture(pic));
        }

        // Parse group shapes
        foreach (var grpSp in spTree.Elements(P + "grpSp"))
        {
            GroupShapes.Add(new GroupShape(grpSp));
        }

        // Parse connectors
        foreach (var cxnSp in spTree.Elements(P + "cxnSp"))
        {
            Connectors.Add(new Connector(cxnSp));
        }

        // Parse graphic frames (tables, charts, etc.)
        foreach (var graphicFrame in spTree.Elements(P + "graphicFrame"))
        {
            ParseGraphicFrame(graphicFrame);
        }

        // Parse slide transition
        var transition = _element.Element(P + "transition");
        if (transition != null)
        {
            Transition = new SlideTransition(transition);
        }

        // Parse slide timing
        var timing = _element.Element(P + "timing");
        if (timing != null)
        {
            Timing = new SlideTiming(timing);
        }
    }

    private void ParseGraphicFrame(XElement graphicFrame)
    {
        var graphic = graphicFrame.Element(A + "graphic");
        if (graphic == null) return;

        var graphicData = graphic.Element(A + "graphicData");
        if (graphicData == null) return;

        var uri = graphicData.Attribute("uri")?.Value;

        // Table
        var tbl = graphicData.Element(A + "tbl");
        if (tbl != null)
        {
            Tables.Add(new Table(graphicFrame));
            return;
        }

        // Chart
        var chart = graphicData.Element(A + "chart");
        if (chart != null || uri?.Contains("chart") == true)
        {
            // Get chart bounds from graphic frame
            var xfrm = graphicFrame.Element(P + "xfrm");
            var bounds = ParseBounds(xfrm);

            // For now, store chart reference with bounds
            // Actual chart data will be loaded from relationships
            var chartRef = chart?.Attribute(R + "id")?.Value;
            if (chartRef != null)
            {
                // Chart will be loaded by PptxDocument using the relationship
                // For now, create a placeholder chart
                Charts.Add(new Chart(graphicData, bounds));
            }
            return;
        }

        // SmartArt
        var smartArt = graphicData.Descendants().FirstOrDefault(e => e.Name.LocalName.Contains("diagram"));
        if (smartArt != null || uri?.Contains("diagram") == true)
        {
            // Get SmartArt bounds from graphic frame
            var xfrm = graphicFrame.Element(P + "xfrm");
            var bounds = ParseBounds(xfrm);

            // Parse SmartArt
            var diagram = graphicData.Element(A + "diagramData") ?? smartArt;
            if (diagram != null)
            {
                SmartArts.Add(new SmartArt(diagram, bounds));
            }
            return;
        }

        // OLE object
        var oleObj = graphicData.Element(A + "oleObj");
        if (oleObj != null)
        {
            // OLE object will be handled separately
            return;
        }
    }

    private static Rect ParseBounds(XElement? xfrm)
    {
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

public class Connector
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? Id { get; }
    public string? Name { get; }
    public ShapeType ConnectorType { get; }
    public Rect Bounds { get; }
    public Outline? Outline { get; }
    public ConnectionPoint? StartConnection { get; }
    public ConnectionPoint? EndConnection { get; }

    public Connector(XElement element)
    {
        _element = element;

        var nvCxnSpPr = element.Element(P + "nvCxnSpPr");
        if (nvCxnSpPr != null)
        {
            var cNvPr = nvCxnSpPr.Element(P + "cNvPr");
            Id = cNvPr?.Attribute("id")?.Value;
            Name = cNvPr?.Attribute("name")?.Value;
        }

        var spPr = element.Element(P + "spPr");
        if (spPr != null)
        {
            ConnectorType = ParseConnectorType(spPr);
            Bounds = ParseBounds(spPr);
            Outline = ParseOutline(spPr);
        }

        // Parse connection points
        var stCxn = element.Element(P + "stCxn");
        if (stCxn != null)
        {
            StartConnection = new ConnectionPoint
            {
                Id = stCxn.Attribute("id")?.Value,
                Index = int.TryParse(stCxn.Attribute("idx")?.Value, out var idx) ? idx : 0
            };
        }

        var endCxn = element.Element(P + "endCxn");
        if (endCxn != null)
        {
            EndConnection = new ConnectionPoint
            {
                Id = endCxn.Attribute("id")?.Value,
                Index = int.TryParse(endCxn.Attribute("idx")?.Value, out var idx) ? idx : 0
            };
        }
    }

    private static ShapeType ParseConnectorType(XElement spPr)
    {
        var prstGeom = spPr.Element(A + "prstGeom");
        if (prstGeom != null)
        {
            var prst = prstGeom.Attribute("prst")?.Value;
            return prst switch
            {
                "straightConnector1" => ShapeType.StraightConnector,
                "elbowConnector" => ShapeType.ElbowConnector,
                "curvedConnector" => ShapeType.CurvedConnector,
                _ => ShapeType.Line
            };
        }

        return ShapeType.Line;
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

    private static Outline? ParseOutline(XElement spPr)
    {
        var ln = spPr.Element(A + "ln");
        if (ln == null) return null;

        var noFill = ln.Element(A + "noFill");
        if (noFill != null)
            return new Outline { Width = 0 };

        var width = int.TryParse(ln.Attribute("w")?.Value, out var w) ? w : 12700;

        var solidFill = ln.Element(A + "solidFill");
        Color? color = null;
        if (solidFill != null)
        {
            color = Shape.ParseColor(solidFill);
        }

        return new Outline { Width = width, Color = color };
    }
}

public class ConnectionPoint
{
    public string? Id { get; set; }
    public int Index { get; set; }
}

public class SlideTransition
{
    private readonly XElement _element;

    public TransitionType Type { get; }
    public int Duration { get; }
    public bool AdvanceOnClick { get; }
    public int AdvanceAfterTime { get; }

    public SlideTransition(XElement element)
    {
        _element = element;

        // Parse transition type
        var type = element.Elements().FirstOrDefault()?.Name.LocalName;
        Type = type switch
        {
            "cut" => TransitionType.Cut,
            "fade" => TransitionType.Fade,
            "push" => TransitionType.Push,
            "wipe" => TransitionType.Wipe,
            "split" => TransitionType.Split,
            "reveal" => TransitionType.Reveal,
            "randomBar" => TransitionType.RandomBars,
            "cover" => TransitionType.Cover,
            "uncover" => TransitionType.Uncover,
            "clock" => TransitionType.Clock,
            "zoom" => TransitionType.Zoom,
            "morph" => TransitionType.Morph,
            _ => TransitionType.None
        };

        // Parse duration
        var spd = element.Attribute("spd")?.Value;
        Duration = spd switch
        {
            "slow" => 2000,
            "med" => 1000,
            "fast" => 500,
            _ => 1000
        };

        // Advance on click
        AdvanceOnClick = element.Attribute("advClick")?.Value != "0";

        // Advance after time
        var advTm = element.Attribute("advTm");
        if (advTm != null && int.TryParse(advTm.Value, out var time))
        {
            AdvanceAfterTime = time;
        }
    }
}

public class SlideTiming
{
    private readonly XElement _element;

    public SlideTiming(XElement element)
    {
        _element = element;
        // Parse timing information if needed
    }
}


