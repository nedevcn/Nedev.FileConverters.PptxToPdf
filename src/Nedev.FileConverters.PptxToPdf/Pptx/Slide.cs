using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class Slide
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
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
    internal List<ChartReference> ChartReferences { get; } = new();
    internal List<SmartArtReference> SmartArtReferences { get; } = new();
    public Background? Background { get; private set; }
    public ColorMap? ColorMap { get; private set; }
    public string SourcePath { get; }
    public SlideLayout? Layout { get; set; }
    public SlideTransition? Transition { get; private set; }
    public SlideTiming? Timing { get; private set; }

    public Slide(XElement element, string id, string sourcePath)
    {
        _element = element;
        Id = id;
        SourcePath = sourcePath;
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
            Background = new Background(bg, SourcePath);
        }

        ColorMap = ColorMap.FromOverride(_element.Element(P + "clrMapOvr"));

        // Parse shape tree
        var spTree = cSld.Element(P + "spTree");
        if (spTree == null) return;

        // Parse shapes
        foreach (var sp in spTree.Elements(P + "sp"))
        {
            var shape = new Shape(sp, SourcePath);
            if (!shape.IsPlaceholder || shape.HasText)
            {
                Shapes.Add(shape);
            }
        }

        // Parse pictures
        foreach (var pic in spTree.Elements(P + "pic"))
        {
            Pictures.Add(new Picture(pic, SourcePath));
        }

        // Parse group shapes
        foreach (var grpSp in spTree.Elements(P + "grpSp"))
        {
            GroupShapes.Add(new GroupShape(grpSp, SourcePath));
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
        var chart = graphicData.Element(C + "chart") ?? graphicData.Elements().FirstOrDefault(element => element.Name.LocalName == "chart");
        if (chart != null || uri?.Contains("chart") == true)
        {
            // Get chart bounds from graphic frame
            var xfrm = graphicFrame.Element(P + "xfrm");
            var bounds = ParseBounds(xfrm);

            var chartRef = chart?.Attribute(R + "id")?.Value;
            if (chartRef != null)
            {
                ChartReferences.Add(new ChartReference(chartRef, bounds));
            }
            return;
        }

        // SmartArt
        var smartArtRelIds = graphicData.Element(D + "relIds") ?? graphicData.Elements().FirstOrDefault(element => element.Name.LocalName == "relIds");
        var smartArt = graphicData.Descendants().FirstOrDefault(element => element.Name.LocalName is "diagramData" or "dataModel" or "layoutDef");
        if (smartArtRelIds != null || smartArt != null || uri?.Contains("diagram") == true)
        {
            // Get SmartArt bounds from graphic frame
            var xfrm = graphicFrame.Element(P + "xfrm");
            var bounds = ParseBounds(xfrm);

            var dataModelRef = smartArtRelIds?.Attribute(R + "dm")?.Value;
            var layoutRef = smartArtRelIds?.Attribute(R + "lo")?.Value;
            if (!string.IsNullOrEmpty(dataModelRef) || !string.IsNullOrEmpty(layoutRef))
            {
                SmartArtReferences.Add(new SmartArtReference(dataModelRef, layoutRef, bounds));
                return;
            }

            // Parse inline SmartArt when diagram parts are embedded directly in the slide.
            var diagram = graphicData.Elements().FirstOrDefault(element => element.Name.LocalName == "diagramData") ?? smartArt;
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

        long.TryParse(off.Attribute("x")?.Value, out var x);
        long.TryParse(off.Attribute("y")?.Value, out var y);
        long.TryParse(ext.Attribute("cx")?.Value, out var w);
        long.TryParse(ext.Attribute("cy")?.Value, out var h);

        return new Rect(x, y, w, h);
    }

    public Background? GetEffectiveBackground()
    {
        if (Background?.IsDefined == true)
            return Background;

        if (Layout?.Background?.IsDefined == true)
            return Layout.Background;

        if (Layout?.Master?.Background?.IsDefined == true)
            return Layout.Master.Background;

        return null;
    }

    public Theme? GetEffectiveTheme(Theme? fallback = null)
    {
        return Layout?.Master?.Theme ?? fallback;
    }

    public ColorMap? GetEffectiveColorMap()
    {
        return ColorMap ?? Layout?.ColorMap ?? Layout?.Master?.ColorMap;
    }

    public IEnumerable<Shape> GetRenderableShapes()
    {
        foreach (var shape in GetInheritedRenderableShapes())
        {
            yield return shape;
        }

        foreach (var shape in Shapes)
        {
            var renderableShape = ResolveRenderableShape(shape);
            if (renderableShape != null)
            {
                yield return renderableShape;
            }
        }
    }

    public IEnumerable<Picture> GetRenderablePictures()
    {
        foreach (var picture in Layout?.Master?.Pictures ?? Enumerable.Empty<Picture>())
        {
            yield return picture;
        }

        foreach (var picture in Layout?.Pictures ?? Enumerable.Empty<Picture>())
        {
            yield return picture;
        }

        foreach (var picture in Pictures)
        {
            yield return picture;
        }
    }

    public IEnumerable<GroupShape> GetRenderableGroupShapes()
    {
        foreach (var group in Layout?.Master?.GroupShapes ?? Enumerable.Empty<GroupShape>())
        {
            yield return group;
        }

        foreach (var group in Layout?.GroupShapes ?? Enumerable.Empty<GroupShape>())
        {
            yield return group;
        }

        foreach (var group in GroupShapes)
        {
            yield return group;
        }
    }

    private IEnumerable<Shape> GetInheritedRenderableShapes()
    {
        foreach (var shape in Layout?.Master?.Shapes ?? Enumerable.Empty<Shape>())
        {
            if (IsOverriddenBySlidePlaceholder(shape))
                continue;

            var renderableShape = ResolveRenderableInheritedShape(shape, null);
            if (renderableShape != null)
            {
                yield return renderableShape;
            }
        }

        foreach (var shape in Layout?.Shapes ?? Enumerable.Empty<Shape>())
        {
            if (IsOverriddenBySlidePlaceholder(shape))
                continue;

            var masterPlaceholder = FindMatchingPlaceholder(Layout?.Master?.Shapes, shape);
            var renderableShape = ResolveRenderableInheritedShape(shape, masterPlaceholder);
            if (renderableShape != null)
            {
                yield return renderableShape;
            }
        }
    }

    private Shape? ResolveRenderableShape(Shape shape)
    {
        if (!shape.IsPlaceholder)
            return shape;

        if (!shape.HasText)
            return null;

        var layoutPlaceholder = FindMatchingPlaceholder(Layout?.Shapes, shape);
        var masterPlaceholder = FindMatchingPlaceholder(Layout?.Master?.Shapes, layoutPlaceholder ?? shape);
        var placeholderBase = layoutPlaceholder?.ResolvePlaceholder(masterPlaceholder, Layout?.Master?.TextStyles) ?? masterPlaceholder;
        return shape.ResolvePlaceholder(placeholderBase, Layout?.Master?.TextStyles);
    }

    private Shape? ResolveRenderableInheritedShape(Shape shape, Shape? placeholderBase)
    {
        if (!shape.IsPlaceholder)
            return shape;

        if (!shape.HasText)
            return null;

        return shape.ResolvePlaceholder(placeholderBase, Layout?.Master?.TextStyles);
    }

    private static Shape? FindMatchingPlaceholder(IEnumerable<Shape>? candidates, Shape placeholder)
    {
        if (!placeholder.IsPlaceholder || candidates == null)
            return null;

        return candidates.FirstOrDefault(candidate => candidate.IsPlaceholder && placeholder.MatchesPlaceholder(candidate))
            ?? candidates.FirstOrDefault(candidate =>
                candidate.IsPlaceholder &&
                placeholder.PlaceholderType != PlaceholderType.None &&
                candidate.PlaceholderType == placeholder.PlaceholderType);
    }

    private bool IsOverriddenBySlidePlaceholder(Shape shape)
    {
        return shape.IsPlaceholder && Shapes.Any(candidate => candidate.IsPlaceholder && candidate.MatchesPlaceholder(shape));
    }
}

internal sealed class ChartReference
{
    public string RelationshipId { get; }
    public Rect Bounds { get; }

    public ChartReference(string relationshipId, Rect bounds)
    {
        RelationshipId = relationshipId;
        Bounds = bounds;
    }
}

internal sealed class SmartArtReference
{
    public string? DataModelRelationshipId { get; }
    public string? LayoutRelationshipId { get; }
    public Rect Bounds { get; }

    public SmartArtReference(string? dataModelRelationshipId, string? layoutRelationshipId, Rect bounds)
    {
        DataModelRelationshipId = dataModelRelationshipId;
        LayoutRelationshipId = layoutRelationshipId;
        Bounds = bounds;
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

        long.TryParse(off.Attribute("x")?.Value, out var x);
        long.TryParse(off.Attribute("y")?.Value, out var y);
        long.TryParse(ext.Attribute("cx")?.Value, out var w);
        long.TryParse(ext.Attribute("cy")?.Value, out var h);

        return new Rect(x, y, w, h);
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


