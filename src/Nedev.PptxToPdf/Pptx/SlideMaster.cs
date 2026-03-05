using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class SlideMaster
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public List<SlideLayout> Layouts { get; } = new();
    public List<Shape> Shapes { get; } = new();
    public List<Picture> Pictures { get; } = new();
    public List<GroupShape> GroupShapes { get; } = new();
    public Background? Background { get; private set; }
    public ColorMap? ColorMap { get; private set; }
    public TextStyles? TextStyles { get; private set; }
    public long Width { get; }
    public long Height { get; }

    public SlideMaster(XElement element)
    {
        _element = element;

        // Parse slide size
        var sldSz = element.Parent?.Element(P + "sldSz");
        if (sldSz != null)
        {
            Width = long.TryParse(sldSz.Attribute("cx")?.Value, out var w) ? w : 9144000;
            Height = long.TryParse(sldSz.Attribute("cy")?.Value, out var h) ? h : 6858000;
        }
        else
        {
            Width = 9144000; // 10 inches in EMUs
            Height = 6858000; // 7.5 inches in EMUs
        }

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

        // Parse color map
        var clrMap = _element.Element(P + "clrMap");
        if (clrMap != null)
        {
            ColorMap = new ColorMap(clrMap);
        }

        // Parse shape tree
        var spTree = cSld.Element(P + "spTree");
        if (spTree == null) return;

        // Parse shapes
        foreach (var sp in spTree.Elements(P + "sp"))
        {
            Shapes.Add(new Shape(sp));
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

        // Parse text styles
        var txStyles = _element.Element(P + "txStyles");
        if (txStyles != null)
        {
            TextStyles = new TextStyles(txStyles);
        }
    }
}

public class SlideLayout
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? Name { get; }
    public SlideLayoutType Type { get; }
    public List<Shape> Shapes { get; } = new();
    public List<Picture> Pictures { get; } = new();
    public List<GroupShape> GroupShapes { get; } = new();
    public Background? Background { get; private set; }
    public ColorMap? ColorMap { get; private set; }

    public SlideLayout(XElement element)
    {
        _element = element;
        Name = element.Attribute("name")?.Value;
        Type = ParseLayoutType(element.Attribute("type")?.Value);

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

        // Parse color map
        var clrMapOvr = _element.Element(P + "clrMapOvr");
        if (clrMapOvr != null)
        {
            var clrMap = clrMapOvr.Element(A + "clrMap");
            if (clrMap != null)
            {
                ColorMap = new ColorMap(clrMap);
            }
        }

        // Parse shape tree
        var spTree = cSld.Element(P + "spTree");
        if (spTree == null) return;

        // Parse shapes
        foreach (var sp in spTree.Elements(P + "sp"))
        {
            Shapes.Add(new Shape(sp));
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
    }

    private static SlideLayoutType ParseLayoutType(string? type)
    {
        return type switch
        {
            "title" => SlideLayoutType.Title,
            "tx" => SlideLayoutType.Text,
            "twoColTx" => SlideLayoutType.TwoColumnText,
            "tbl" => SlideLayoutType.Table,
            "txAndChart" => SlideLayoutType.TextAndChart,
            "chartAndTx" => SlideLayoutType.ChartAndText,
            "dgm" => SlideLayoutType.Diagram,
            "chart" => SlideLayoutType.Chart,
            "txAndClipArt" => SlideLayoutType.TextAndClipArt,
            "clipArtAndTx" => SlideLayoutType.ClipArtAndText,
            "titleOnly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "txAndObj" => SlideLayoutType.TextAndObject,
            "objAndTx" => SlideLayoutType.ObjectAndText,
            "objOnly" => SlideLayoutType.LargeObject,
            "obj" => SlideLayoutType.Object,
            "titleSlide" => SlideLayoutType.TitleSlide,
            "titleAndObj" => SlideLayoutType.TitleAndObject,
            "titleAndMedia" => SlideLayoutType.TitleAndMedia,
            "mediaAndTitle" => SlideLayoutType.MediaAndTitle,
            "objOverTx" => SlideLayoutType.ObjectOverText,
            "txOverObj" => SlideLayoutType.TextOverObject,
            "txAndTwoObj" => SlideLayoutType.TextAndTwoObjects,
            "twoObjAndTx" => SlideLayoutType.TwoObjectsAndText,
            "twoObjOverTx" => SlideLayoutType.TwoObjectsOverText,
            "fourObj" => SlideLayoutType.FourObjects,
            "vertTitleAndTx" => SlideLayoutType.VerticalTitleAndText,
            "vertTwoColTx" => SlideLayoutType.VerticalTwoColumnText,
            _ => SlideLayoutType.Custom
        };
    }
}

public class ColorMap
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public Dictionary<string, string> Mappings { get; } = new();

    public ColorMap(XElement element)
    {
        _element = element;
        ParseMappings();
    }

    private void ParseMappings()
    {
        // Background colors
        Mappings["bg1"] = _element.Attribute("bg1")?.Value ?? "lt1";
        Mappings["bg2"] = _element.Attribute("bg2")?.Value ?? "lt2";

        // Text colors
        Mappings["tx1"] = _element.Attribute("tx1")?.Value ?? "dk1";
        Mappings["tx2"] = _element.Attribute("tx2")?.Value ?? "dk2";

        // Accent colors
        Mappings["accent1"] = _element.Attribute("accent1")?.Value ?? "accent1";
        Mappings["accent2"] = _element.Attribute("accent2")?.Value ?? "accent2";
        Mappings["accent3"] = _element.Attribute("accent3")?.Value ?? "accent3";
        Mappings["accent4"] = _element.Attribute("accent4")?.Value ?? "accent4";
        Mappings["accent5"] = _element.Attribute("accent5")?.Value ?? "accent5";
        Mappings["accent6"] = _element.Attribute("accent6")?.Value ?? "accent6";

        // Hyperlink colors
        Mappings["hlink"] = _element.Attribute("hlink")?.Value ?? "hlink";
        Mappings["folHlink"] = _element.Attribute("folHlink")?.Value ?? "folHlink";
    }

    public string GetMapping(string colorName)
    {
        return Mappings.TryGetValue(colorName, out var mapping) ? mapping : colorName;
    }
}

public class TextStyles
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public TextStyle? TitleStyle { get; }
    public TextStyle? BodyStyle { get; }
    public TextStyle? OtherStyle { get; }

    public TextStyles(XElement element)
    {
        _element = element;

        var titleStyle = element.Element(P + "titleStyle");
        if (titleStyle != null)
        {
            TitleStyle = new TextStyle(titleStyle);
        }

        var bodyStyle = element.Element(P + "bodyStyle");
        if (bodyStyle != null)
        {
            BodyStyle = new TextStyle(bodyStyle);
        }

        var otherStyle = element.Element(P + "otherStyle");
        if (otherStyle != null)
        {
            OtherStyle = new TextStyle(otherStyle);
        }
    }
}

public class TextStyle
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public List<LevelStyle> Levels { get; } = new();

    public TextStyle(XElement element)
    {
        _element = element;

        for (int i = 1; i <= 9; i++)
        {
            var lvl = element.Element(A + $"lvl{i}pPr");
            if (lvl != null)
            {
                Levels.Add(new LevelStyle(lvl, i));
            }
        }
    }
}

public class LevelStyle
{
    public int Level { get; }
    public TextParagraphProperties? Properties { get; }

    public LevelStyle(XElement element, int level)
    {
        Level = level;
        Properties = new TextParagraphProperties(element);
    }
}

public class TextParagraphProperties
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public TextAlignment? Alignment { get; }
    public double? LeftMargin { get; }
    public double? RightMargin { get; }
    public double? Indent { get; }
    public double? DefaultTabSize { get; }
    public int? Level { get; }
    public BulletStyle? Bullet { get; }
    public TextRunProperties? DefaultRunProperties { get; }

    public TextParagraphProperties(XElement element)
    {
        _element = element;

        // Parse alignment
        var algn = element.Attribute("algn")?.Value;
        Alignment = algn switch
        {
            "l" => TextAlignment.Left,
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            "just" => TextAlignment.Justify,
            "justLow" => TextAlignment.Justify,
            "dist" => TextAlignment.Distributed,
            _ => null
        };

        // Parse margins
        if (long.TryParse(element.Attribute("marL")?.Value, out var marL))
            LeftMargin = marL / 12700.0;

        if (long.TryParse(element.Attribute("marR")?.Value, out var marR))
            RightMargin = marR / 12700.0;

        if (long.TryParse(element.Attribute("indent")?.Value, out var indent))
            Indent = indent / 12700.0;

        if (long.TryParse(element.Attribute("defTabSz")?.Value, out var defTabSz))
            DefaultTabSize = defTabSz / 12700.0;

        if (int.TryParse(element.Attribute("lvl")?.Value, out var lvl))
            Level = lvl;

        // Parse bullet
        var buNone = element.Element(A + "buNone");
        if (buNone != null)
        {
            Bullet = new BulletStyle { Type = BulletType.None };
        }
        else
        {
            var buAutoNum = element.Element(A + "buAutoNum");
            if (buAutoNum != null)
            {
                Bullet = new BulletStyle
                {
                    Type = BulletType.AutoNumber,
                    AutoNumberType = buAutoNum.Attribute("type")?.Value
                };
            }

            var buChar = element.Element(A + "buChar");
            if (buChar != null)
            {
                Bullet = new BulletStyle
                {
                    Type = BulletType.Char,
                    Character = buChar.Attribute("char")?.Value
                };
            }

            var buBlip = element.Element(A + "buBlip");
            if (buBlip != null)
            {
                Bullet = new BulletStyle { Type = BulletType.Blip };
            }
        }

        // Parse default run properties
        var defRPr = element.Element(A + "defRPr");
        if (defRPr != null)
        {
            DefaultRunProperties = new TextRunProperties(defRPr);
        }
    }
}

public class TextRunProperties
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Language { get; }
    public string? AlternativeLanguage { get; }
    public double? FontSize { get; }
    public bool? Bold { get; }
    public bool? Italic { get; }
    public UnderlineType? Underline { get; }
    public StrikeType? Strike { get; }
    public double? Kerning { get; }
    public CapsType? Caps { get; }
    public int? Spacing { get; }
    public string? Typeface { get; }
    public Color? Color { get; }

    public TextRunProperties(XElement element)
    {
        _element = element;

        Language = element.Attribute("lang")?.Value;
        AlternativeLanguage = element.Attribute("altLang")?.Value;

        if (int.TryParse(element.Attribute("sz")?.Value, out var sz))
            FontSize = sz / 100.0;

        if (element.Attribute("b")?.Value is string b)
            Bold = b == "1";

        if (element.Attribute("i")?.Value is string i)
            Italic = i == "1";

        var u = element.Attribute("u")?.Value;
        Underline = u switch
        {
            "sng" => UnderlineType.Single,
            "dbl" => UnderlineType.Double,
            "sngAccounting" => UnderlineType.SingleAccounting,
            "dblAccounting" => UnderlineType.DoubleAccounting,
            "words" => UnderlineType.Words,
            "none" => UnderlineType.None,
            _ => null
        };

        var strike = element.Attribute("strike")?.Value;
        Strike = strike switch
        {
            "sngStrike" => StrikeType.Single,
            "dblStrike" => StrikeType.Double,
            "noStrike" => StrikeType.None,
            _ => null
        };

        if (int.TryParse(element.Attribute("kern")?.Value, out var kern))
            Kerning = kern / 100.0;

        var cap = element.Attribute("cap")?.Value;
        Caps = cap switch
        {
            "small" => CapsType.Small,
            "all" => CapsType.All,
            "none" => CapsType.None,
            _ => null
        };

        if (int.TryParse(element.Attribute("spc")?.Value, out var spc))
            Spacing = spc;

        // Parse font
        var latin = element.Element(A + "latin");
        if (latin != null)
        {
            Typeface = latin.Attribute("typeface")?.Value;
        }

        // Parse color
        var solidFill = element.Element(A + "solidFill");
        if (solidFill != null)
        {
            Color = Shape.ParseColor(solidFill);
        }
    }
}

public class BulletStyle
{
    public BulletType Type { get; set; }
    public string? Character { get; set; }
    public string? AutoNumberType { get; set; }
    public int? StartAt { get; set; }
    public Color? Color { get; set; }
    public double? Size { get; set; }
    public string? Font { get; set; }
}
