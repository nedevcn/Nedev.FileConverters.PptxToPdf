using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class SmartArt
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    public string? Type { get; }
    public SmartArtType ResolvedType { get; }
    public string? DisplayName { get; }
    public Rect Bounds { get; }
    public List<SmartArtNode> Nodes { get; } = new();
    public List<SmartArtConnection> Connections { get; } = new();
    public SmartArtLayout? Layout { get; }

    public SmartArt(XElement element, Rect bounds)
        : this(ResolveDataModel(element), ResolveLayoutDef(element), bounds)
    {
    }

    public SmartArt(XElement? dataModelElement, XElement? layoutDefElement, Rect bounds)
    {
        _element = dataModelElement ?? layoutDefElement ?? new XElement(D + "smartArt");
        Bounds = bounds;

        var dataModel = ResolveDataModel(dataModelElement);
        if (dataModel != null)
        {
            ParseDataModel(dataModel);
        }

        var layoutDef = ResolveLayoutDef(layoutDefElement);
        if (layoutDef != null)
        {
            Layout = new SmartArtLayout(layoutDef);
        }

        Type = layoutDef?.Attribute("uniqueId")?.Value
            ?? layoutDefElement?.Attribute("uniqueId")?.Value
            ?? dataModelElement?.Attribute("uniqueId")?.Value;
        ResolvedType = ResolveSmartArtType(Type, Layout?.Name, Layout?.Description, Layout?.Category);
        DisplayName = Layout?.Name
            ?? (ResolvedType != SmartArtType.Unknown ? GetDisplayName(ResolvedType) : Type);
    }

    private void ParseDataModel(XElement dataModel)
    {
        var ptLst = dataModel.Element(D + "ptLst");
        if (ptLst != null)
        {
            foreach (var pt in ptLst.Elements(D + "pt"))
            {
                Nodes.Add(new SmartArtNode(pt));
            }
        }

        var cxnLst = dataModel.Element(D + "cxnLst");
        if (cxnLst != null)
        {
            foreach (var cxn in cxnLst.Elements(D + "cxn"))
            {
                Connections.Add(new SmartArtConnection(cxn));
            }
        }
    }

    private static XElement? ResolveDataModel(XElement? element)
    {
        if (element == null)
            return null;

        return element.Name == D + "dataModel"
            ? element
            : element.Element(D + "dataModel");
    }

    private static XElement? ResolveLayoutDef(XElement? element)
    {
        if (element == null)
            return null;

        return element.Name == D + "layoutDef"
            ? element
            : element.Element(D + "layoutDef");
    }

    public static SmartArtType ResolveSmartArtType(params string?[] candidates)
    {
        foreach (var candidate in candidates)
        {
            var resolved = GetSmartArtType(candidate);
            if (resolved != SmartArtType.Unknown)
                return resolved;
        }

        return SmartArtType.Unknown;
    }

    public static SmartArtType GetSmartArtType(string? uniqueId)
    {
        if (string.IsNullOrWhiteSpace(uniqueId))
            return SmartArtType.Unknown;

        return uniqueId switch
        {
            // List types
            var s when s.Contains("VerticalBulletList", StringComparison.OrdinalIgnoreCase) => SmartArtType.VerticalBulletList,
            var s when s.Contains("HorizontalBulletList", StringComparison.OrdinalIgnoreCase) => SmartArtType.HorizontalBulletList,
            var s when s.Contains("List", StringComparison.OrdinalIgnoreCase) => SmartArtType.List,

            // Process types
            var s when s.Contains("ContinuousBlockProcess", StringComparison.OrdinalIgnoreCase) => SmartArtType.ContinuousBlockProcess,
            var s when s.Contains("BasicProcess", StringComparison.OrdinalIgnoreCase) => SmartArtType.BasicProcess,
            var s when s.Contains("Process", StringComparison.OrdinalIgnoreCase) => SmartArtType.Process,

            // Cycle types
            var s when s.Contains("BasicCycle", StringComparison.OrdinalIgnoreCase) => SmartArtType.BasicCycle,
            var s when s.Contains("Cycle", StringComparison.OrdinalIgnoreCase) => SmartArtType.Cycle,

            // Hierarchy types
            var s when s.Contains("OrganizationChart", StringComparison.OrdinalIgnoreCase) => SmartArtType.OrganizationChart,
            var s when s.Contains("Hierarchy", StringComparison.OrdinalIgnoreCase) => SmartArtType.Hierarchy,

            // Relationship types
            var s when s.Contains("BasicTarget", StringComparison.OrdinalIgnoreCase) => SmartArtType.BasicTarget,
            var s when s.Contains("Relationship", StringComparison.OrdinalIgnoreCase) => SmartArtType.Relationship,

            // Matrix types
            var s when s.Contains("Matrix", StringComparison.OrdinalIgnoreCase) => SmartArtType.Matrix,

            // Pyramid types
            var s when s.Contains("Pyramid", StringComparison.OrdinalIgnoreCase) => SmartArtType.Pyramid,

            // Picture types
            var s when s.Contains("Picture", StringComparison.OrdinalIgnoreCase) => SmartArtType.Picture,

            _ => SmartArtType.Unknown
        };
    }

    public static string GetDisplayName(SmartArtType type)
    {
        return type switch
        {
            SmartArtType.VerticalBulletList => "Vertical Bullet List",
            SmartArtType.HorizontalBulletList => "Horizontal Bullet List",
            SmartArtType.BasicProcess => "Basic Process",
            SmartArtType.ContinuousBlockProcess => "Continuous Block Process",
            SmartArtType.OrganizationChart => "Organization Chart",
            SmartArtType.BasicCycle => "Basic Cycle",
            SmartArtType.BasicTarget => "Basic Target",
            SmartArtType.Unknown => "SmartArt",
            _ => type.ToString()
        };
    }
}

public enum SmartArtType
{
    Unknown,
    List,
    VerticalBulletList,
    HorizontalBulletList,
    Process,
    BasicProcess,
    ContinuousBlockProcess,
    Cycle,
    BasicCycle,
    Hierarchy,
    OrganizationChart,
    Relationship,
    BasicTarget,
    Matrix,
    Pyramid,
    Picture
}

public class SmartArtNode
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    public string? Id { get; }
    public string? Type { get; }
    public string? ModelId { get; }
    public SmartArtNodeProperties? Properties { get; }
    public string? Text { get; private set; }
    public List<SmartArtTextRun> TextRuns { get; } = new();

    public SmartArtNode(XElement element)
    {
        _element = element;

        Id = element.Attribute("id")?.Value;
        Type = element.Attribute("type")?.Value;
        ModelId = element.Attribute("modelId")?.Value;

        // Parse properties
        var prSet = element.Element(D + "prSet");
        if (prSet != null)
        {
            Properties = new SmartArtNodeProperties(prSet);
        }

        // Parse text
        var txBody = element.Element(D + "txBody");
        if (txBody != null)
        {
            ParseTextWithFormatting(txBody);
        }
        else
        {
            // Try to get text from t element
            var t = element.Element(D + "t");
            if (t != null)
            {
                Text = t.Value;
                TextRuns.Add(new SmartArtTextRun(t.Value, false, false, false, 12, "000000"));
            }
        }
    }

    private void ParseTextWithFormatting(XElement txBody)
    {
        var paragraphs = txBody.Elements(A + "p").ToList();
        if (paragraphs.Count == 0)
        {
            paragraphs = txBody.Descendants(A + "p").ToList();
        }

        if (paragraphs.Count == 0) return;

        var text = "";
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
        {
            var p = paragraphs[paragraphIndex];
            if (paragraphIndex > 0 && text.Length > 0)
            {
                text += "\n";
                TextRuns.Add(SmartArtTextRun.LineBreak());
            }

            foreach (var child in p.Elements())
            {
                if (child.Name == A + "br")
                {
                    text += "\n";
                    TextRuns.Add(SmartArtTextRun.LineBreak());
                    continue;
                }

                if (child.Name != A + "r")
                    continue;

                var t = child.Element(A + "t");
                if (t == null)
                    continue;

                var textValue = t.Value;
                text += textValue;

                var rPr = child.Element(A + "rPr");
                var bold = rPr?.Attribute("b")?.Value == "1";
                var italic = rPr?.Attribute("i")?.Value == "1";
                var underline = rPr?.Attribute("u")?.Value == "single";
                var fontSize = int.TryParse(rPr?.Attribute("sz")?.Value, out var size) ? size / 100 : 12;
                var color = "000000";

                var solidFill = rPr?.Element(A + "solidFill");
                if (solidFill != null)
                {
                    var srgbClr = solidFill.Element(A + "srgbClr");
                    if (srgbClr != null)
                    {
                        color = srgbClr.Attribute("val")?.Value ?? "000000";
                    }
                }

                TextRuns.Add(new SmartArtTextRun(textValue, bold, italic, underline, fontSize, color));
            }
        }

        Text = string.IsNullOrEmpty(text) ? null : text;
    }
}

public class SmartArtTextRun
{
    public string Text { get; }
    public bool IsLineBreak => Text == "\n";
    public bool Bold { get; }
    public bool Italic { get; }
    public bool Underline { get; }
    public int FontSize { get; }
    public string Color { get; }

    public SmartArtTextRun(string text, bool bold, bool italic, bool underline, int fontSize, string color)
    {
        Text = text;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        FontSize = fontSize;
        Color = color;
    }

    public static SmartArtTextRun LineBreak(int fontSize = 12, string color = "000000")
    {
        return new SmartArtTextRun("\n", false, false, false, fontSize, color);
    }
}

public class SmartArtNodeProperties
{
    private readonly XElement _element;

    public int? PresId { get; }
    public int? PresAssocId { get; }
    public int? PresName { get; }
    public int? PresStyleIdx { get; }
    public int? PresStyleCnt { get; }
    public int? LoTypeId { get; }
    public int? LoCatId { get; }
    public int? QsTypeId { get; }
    public int? QsCatId { get; }
    public int? CsTypeId { get; }
    public int? CsCatId { get; }
    public int? Coherent3DOff { get; }
    public int? Phldr { get; }
    public int? PhldrT { get; }

    public SmartArtNodeProperties(XElement element)
    {
        _element = element;

        if (int.TryParse(element.Attribute("presId")?.Value, out var presId))
            PresId = presId;

        if (int.TryParse(element.Attribute("presAssocId")?.Value, out var presAssocId))
            PresAssocId = presAssocId;

        if (int.TryParse(element.Attribute("presName")?.Value, out var presName))
            PresName = presName;

        if (int.TryParse(element.Attribute("presStyleIdx")?.Value, out var presStyleIdx))
            PresStyleIdx = presStyleIdx;

        if (int.TryParse(element.Attribute("presStyleCnt")?.Value, out var presStyleCnt))
            PresStyleCnt = presStyleCnt;

        if (int.TryParse(element.Attribute("loTypeId")?.Value, out var loTypeId))
            LoTypeId = loTypeId;

        if (int.TryParse(element.Attribute("loCatId")?.Value, out var loCatId))
            LoCatId = loCatId;

        if (int.TryParse(element.Attribute("qsTypeId")?.Value, out var qsTypeId))
            QsTypeId = qsTypeId;

        if (int.TryParse(element.Attribute("qsCatId")?.Value, out var qsCatId))
            QsCatId = qsCatId;

        if (int.TryParse(element.Attribute("csTypeId")?.Value, out var csTypeId))
            CsTypeId = csTypeId;

        if (int.TryParse(element.Attribute("csCatId")?.Value, out var csCatId))
            CsCatId = csCatId;

        if (int.TryParse(element.Attribute("coherent3DOff")?.Value, out var coherent3DOff))
            Coherent3DOff = coherent3DOff;

        if (int.TryParse(element.Attribute("phldr")?.Value, out var phldr))
            Phldr = phldr;

        if (int.TryParse(element.Attribute("phldrT")?.Value, out var phldrT))
            PhldrT = phldrT;
    }
}

public class SmartArtConnection
{
    private readonly XElement _element;
    private static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    public string? ModelId { get; }
    public string? SourceId { get; }
    public string? DestinationId { get; }
    public string? SourcePoint { get; }
    public string? DestinationPoint { get; }
    public int? StartId { get; }
    public int? EndId { get; }
    public int? Count { get; }
    public bool IsBidirectional { get; }

    public SmartArtConnection(XElement element)
    {
        _element = element;

        ModelId = element.Attribute("modelId")?.Value;
        SourceId = element.Attribute("srcId")?.Value;
        DestinationId = element.Attribute("destId")?.Value;
        SourcePoint = element.Attribute("srcOrd")?.Value;
        DestinationPoint = element.Attribute("destOrd")?.Value;

        if (int.TryParse(element.Attribute("sId")?.Value, out var sId))
            StartId = sId;

        if (int.TryParse(element.Attribute("eId")?.Value, out var eId))
            EndId = eId;

        if (int.TryParse(element.Attribute("cnt")?.Value, out var cnt))
            Count = cnt;

        IsBidirectional = element.Attribute("parTrans")?.Value == "bi";
    }
}

public class SmartArtLayout
{
    private readonly XElement _element;
    private static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    public string? UniqueId { get; }
    public string? Description { get; }
    public string? Category { get; }
    public string? Name { get; }

    public SmartArtLayout(XElement element)
    {
        _element = element;

        UniqueId = element.Attribute("uniqueId")?.Value;
        Description = element.Attribute("desc")?.Value;
        Category = element.Attribute("cat")?.Value;
        Name = element.Attribute("name")?.Value;
    }
}

public class SmartArtDrawing
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    public List<SmartArtShape> Shapes { get; } = new();
    public List<SmartArtConnector> Connectors { get; } = new();

    public SmartArtDrawing(XElement element)
    {
        _element = element;

        // Parse shapes
        foreach (var sp in element.Descendants(A + "sp"))
        {
            Shapes.Add(new SmartArtShape(sp));
        }

        // Parse connectors
        foreach (var cxnSp in element.Descendants(A + "cxnSp"))
        {
            Connectors.Add(new SmartArtConnector(cxnSp));
        }
    }
}

public class SmartArtShape
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? ModelId { get; }
    public Rect Bounds { get; }
    public ShapeType Type { get; }
    public Fill? Fill { get; }
    public Outline? Outline { get; }
    public XElement? TextBody { get; }

    public SmartArtShape(XElement element)
    {
        _element = element;

        // Parse model ID from non-visual properties
        var nvSpPr = element.Element(A + "nvSpPr");
        var cNvPr = nvSpPr?.Element(A + "cNvPr");
        ModelId = cNvPr?.Attribute("id")?.Value;

        // Parse shape properties
        var spPr = element.Element(A + "spPr");
        if (spPr != null)
        {
            Bounds = ParseBounds(spPr);
            Type = ParseShapeType(spPr);
            Fill = Shape.ParseFill(spPr);
            Outline = ParseOutline(spPr);
        }

        // Parse text
        var txBody = element.Element(A + "txBody");
        if (txBody != null)
        {
            TextBody = txBody;
        }
    }

    private static Rect ParseBounds(XElement spPr)
    {
        var xfrm = spPr.Element(spPr.Name.Namespace + "xfrm");
        if (xfrm == null) return new Rect();

        var off = xfrm.Element(spPr.Name.Namespace + "off");
        var ext = xfrm.Element(spPr.Name.Namespace + "ext");

        if (off == null || ext == null) return new Rect();

        long.TryParse(off.Attribute("x")?.Value, out var x);
        long.TryParse(off.Attribute("y")?.Value, out var y);
        long.TryParse(ext.Attribute("cx")?.Value, out var w);
        long.TryParse(ext.Attribute("cy")?.Value, out var h);

        return new Rect(x, y, w, h);
    }

    private static ShapeType ParseShapeType(XElement spPr)
    {
        var prstGeom = spPr.Element(spPr.Name.Namespace + "prstGeom");
        if (prstGeom != null)
        {
            var prst = prstGeom.Attribute("prst")?.Value;
            return ShapeTypeMapping.FromPresetGeometry(prst);
        }

        return ShapeType.Rectangle;
    }

    private static Outline? ParseOutline(XElement spPr)
    {
        var ln = spPr.Element(spPr.Name.Namespace + "ln");
        if (ln == null) return null;

        var noFill = ln.Element(spPr.Name.Namespace + "noFill");
        if (noFill != null)
            return new Outline { Width = 0 };

        var width = int.TryParse(ln.Attribute("w")?.Value, out var w) ? w : 12700;

        var solidFill = ln.Element(spPr.Name.Namespace + "solidFill");
        Color? color = null;
        if (solidFill != null)
        {
            color = Shape.ParseColor(solidFill);
        }

        return new Outline { Width = width, Color = color };
    }
}

public class SmartArtConnector
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? ModelId { get; }
    public string? SourceId { get; }
    public string? DestinationId { get; }
    public Outline? Outline { get; }

    public SmartArtConnector(XElement element)
    {
        _element = element;

        // Parse model ID
        var nvCxnSpPr = element.Element(A + "nvCxnSpPr");
        var cNvPr = nvCxnSpPr?.Element(A + "cNvPr");
        ModelId = cNvPr?.Attribute("id")?.Value;

        // Parse connection points
        var stCxn = element.Element(A + "stCxn");
        if (stCxn != null)
        {
            SourceId = stCxn.Attribute("id")?.Value;
        }

        var endCxn = element.Element(A + "endCxn");
        if (endCxn != null)
        {
            DestinationId = endCxn.Attribute("id")?.Value;
        }

        // Parse outline
        var spPr = element.Element(A + "spPr");
        if (spPr != null)
        {
            var ln = spPr.Element(A + "ln");
            if (ln != null)
            {
                var noFill = ln.Element(A + "noFill");
                if (noFill != null)
                {
                    Outline = new Outline { Width = 0 };
                }
                else
                {
                    var width = int.TryParse(ln.Attribute("w")?.Value, out var w) ? w : 12700;
                    var solidFill = ln.Element(A + "solidFill");
                    Color? color = null;
                    if (solidFill != null)
                    {
                        color = Shape.ParseColor(solidFill);
                    }
                    Outline = new Outline { Width = width, Color = color };
                }
            }
        }
    }
}
