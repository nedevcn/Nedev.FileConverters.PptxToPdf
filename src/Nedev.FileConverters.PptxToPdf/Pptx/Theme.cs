using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class Theme
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Name { get; }
    public string? SourcePath { get; }
    public ColorScheme ColorScheme { get; }
    public FontScheme FontScheme { get; }
    public EffectScheme EffectScheme { get; }
    public FormatScheme FormatScheme { get; }

    public Theme(XElement element, string? sourcePath = null)
    {
        _element = element;
        Name = element.Attribute("name")?.Value;
        SourcePath = sourcePath;

        var themeElements = element.Element(A + "themeElements");
        if (themeElements != null)
        {
            ColorScheme = new ColorScheme(themeElements.Element(A + "clrScheme"));
            FontScheme = new FontScheme(themeElements.Element(A + "fontScheme"));
            using var _ = Shape.UseSchemeColorResolver(ColorScheme.GetColor);
            EffectScheme = new EffectScheme(themeElements.Element(A + "effectScheme"));
            FormatScheme = new FormatScheme(themeElements.Element(A + "fmtScheme"));
        }
        else
        {
            ColorScheme = new ColorScheme(null);
            FontScheme = new FontScheme(null);
            EffectScheme = new EffectScheme(null);
            FormatScheme = new FormatScheme(null);
        }
    }

    public Color GetSchemeColor(SchemeColor schemeColor)
    {
        return ColorScheme.GetColor(schemeColor);
    }
}

public class ColorScheme
{
    private readonly XElement? _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Name { get; }

    private readonly Dictionary<SchemeColor, Color> _colors = new();

    public ColorScheme(XElement? element)
    {
        _element = element;
        Name = element?.Attribute("name")?.Value;
        ParseColors();
    }

    private void ParseColors()
    {
        if (_element == null)
        {
            SetDefaultColors();
            return;
        }

        // Parse dk1 (Dark 1)
        _colors[SchemeColor.Dark1] = ParseColorElement(_element.Element(A + "dk1"));
        // Parse lt1 (Light 1)
        _colors[SchemeColor.Light1] = ParseColorElement(_element.Element(A + "lt1"));
        // Parse dk2 (Dark 2)
        _colors[SchemeColor.Dark2] = ParseColorElement(_element.Element(A + "dk2"));
        // Parse lt2 (Light 2)
        _colors[SchemeColor.Light2] = ParseColorElement(_element.Element(A + "lt2"));
        // Parse accent1-6
        _colors[SchemeColor.Accent1] = ParseColorElement(_element.Element(A + "accent1"));
        _colors[SchemeColor.Accent2] = ParseColorElement(_element.Element(A + "accent2"));
        _colors[SchemeColor.Accent3] = ParseColorElement(_element.Element(A + "accent3"));
        _colors[SchemeColor.Accent4] = ParseColorElement(_element.Element(A + "accent4"));
        _colors[SchemeColor.Accent5] = ParseColorElement(_element.Element(A + "accent5"));
        _colors[SchemeColor.Accent6] = ParseColorElement(_element.Element(A + "accent6"));
        // Parse hlink (Hyperlink)
        _colors[SchemeColor.Hyperlink] = ParseColorElement(_element.Element(A + "hlink"));
        // Parse folHlink (Followed Hyperlink)
        _colors[SchemeColor.FollowedHyperlink] = ParseColorElement(_element.Element(A + "folHlink"));

        // Background1 and Text1 are typically mapped from lt1 and dk1
        if (_colors.ContainsKey(SchemeColor.Light1))
            _colors[SchemeColor.Background1] = _colors[SchemeColor.Light1];
        if (_colors.ContainsKey(SchemeColor.Dark1))
            _colors[SchemeColor.Text1] = _colors[SchemeColor.Dark1];
        if (_colors.ContainsKey(SchemeColor.Light2))
            _colors[SchemeColor.Background2] = _colors[SchemeColor.Light2];
        if (_colors.ContainsKey(SchemeColor.Dark2))
            _colors[SchemeColor.Text2] = _colors[SchemeColor.Dark2];
    }

    private Color ParseColorElement(XElement? colorElement)
    {
        if (colorElement == null) return Color.Black;

        // Try srgbClr
        var srgbClr = colorElement.Element(A + "srgbClr");
        if (srgbClr != null)
        {
            var val = srgbClr.Attribute("val")?.Value;
            if (val != null && val.Length == 6)
            {
                var r = Convert.ToByte(val.Substring(0, 2), 16);
                var g = Convert.ToByte(val.Substring(2, 2), 16);
                var b = Convert.ToByte(val.Substring(4, 2), 16);
                return new Color(r, g, b);
            }
        }

        // Try sysClr (system color)
        var sysClr = colorElement.Element(A + "sysClr");
        if (sysClr != null)
        {
            var lastClr = sysClr.Attribute("lastClr")?.Value;
            if (lastClr != null && lastClr.Length == 6)
            {
                var r = Convert.ToByte(lastClr.Substring(0, 2), 16);
                var g = Convert.ToByte(lastClr.Substring(2, 2), 16);
                var b = Convert.ToByte(lastClr.Substring(4, 2), 16);
                return new Color(r, g, b);
            }
        }

        return Color.Black;
    }

    private void SetDefaultColors()
    {
        _colors[SchemeColor.Dark1] = new Color(0, 0, 0);
        _colors[SchemeColor.Light1] = new Color(255, 255, 255);
        _colors[SchemeColor.Dark2] = new Color(64, 64, 64);
        _colors[SchemeColor.Light2] = new Color(240, 240, 240);
        _colors[SchemeColor.Accent1] = new Color(68, 114, 196);
        _colors[SchemeColor.Accent2] = new Color(237, 125, 49);
        _colors[SchemeColor.Accent3] = new Color(165, 165, 165);
        _colors[SchemeColor.Accent4] = new Color(255, 192, 0);
        _colors[SchemeColor.Accent5] = new Color(91, 155, 213);
        _colors[SchemeColor.Accent6] = new Color(112, 173, 71);
        _colors[SchemeColor.Hyperlink] = new Color(5, 99, 193);
        _colors[SchemeColor.FollowedHyperlink] = new Color(149, 79, 114);
        _colors[SchemeColor.Background1] = new Color(255, 255, 255);
        _colors[SchemeColor.Text1] = new Color(0, 0, 0);
        _colors[SchemeColor.Background2] = new Color(240, 240, 240);
        _colors[SchemeColor.Text2] = new Color(64, 64, 64);
    }

    public Color GetColor(SchemeColor schemeColor)
    {
        return _colors.TryGetValue(schemeColor, out var color) ? color : Color.Black;
    }
}

public class FontScheme
{
    private readonly XElement? _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Name { get; }
    public FontCollection MajorFont { get; }
    public FontCollection MinorFont { get; }

    public FontScheme(XElement? element)
    {
        _element = element;
        Name = element?.Attribute("name")?.Value;

        var majorFont = element?.Element(A + "majorFont");
        MajorFont = new FontCollection(majorFont);

        var minorFont = element?.Element(A + "minorFont");
        MinorFont = new FontCollection(minorFont);
    }
}

public class FontCollection
{
    private readonly XElement? _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? LatinFont { get; }
    public string? EastAsianFont { get; }
    public string? ComplexScriptFont { get; }

    public FontCollection(XElement? element)
    {
        _element = element;

        LatinFont = element?.Element(A + "latin")?.Attribute("typeface")?.Value;
        EastAsianFont = element?.Element(A + "ea")?.Attribute("typeface")?.Value;
        ComplexScriptFont = element?.Element(A + "cs")?.Attribute("typeface")?.Value;
    }
}

public class EffectScheme
{
    private readonly XElement? _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Name { get; }
    public List<EffectStyle> EffectStyles { get; } = new();

    public EffectScheme(XElement? element)
    {
        _element = element;
        Name = element?.Attribute("name")?.Value;

        var effectLst = element?.Element(A + "effectLst");
        if (effectLst != null)
        {
            foreach (var effect in effectLst.Elements())
            {
                EffectStyles.Add(new EffectStyle(effect));
            }
        }
    }
}

public class EffectStyle
{
    private readonly XElement _element;

    public EffectType Type { get; }
    public EffectProperties Properties { get; }

    public EffectStyle(XElement element)
    {
        _element = element;
        Type = ParseEffectType(element.Name.LocalName);
        Properties = new EffectProperties(element);
    }

    private static EffectType ParseEffectType(string name)
    {
        return name switch
        {
            "outerShdw" => EffectType.OuterShadow,
            "innerShdw" => EffectType.InnerShadow,
            "glow" => EffectType.Glow,
            "reflection" => EffectType.Reflection,
            "softEdge" => EffectType.SoftEdge,
            "blur" => EffectType.Blur,
            _ => EffectType.None
        };
    }
}

public enum EffectType
{
    None,
    OuterShadow,
    InnerShadow,
    Glow,
    Reflection,
    SoftEdge,
    Blur
}

public class EffectProperties
{
    public double BlurRadius { get; }
    public double Distance { get; }
    public double Direction { get; }
    public Color? Color { get; }
    public double Size { get; }

    public EffectProperties(XElement element)
    {
        if (double.TryParse(element.Attribute("blurRad")?.Value, out var blurRad))
            BlurRadius = blurRad / 12700.0; // EMUs to points

        if (double.TryParse(element.Attribute("dist")?.Value, out var dist))
            Distance = dist / 12700.0;

        if (double.TryParse(element.Attribute("dir")?.Value, out var dir))
            Direction = dir / 60000.0; // Angle in degrees

        if (double.TryParse(element.Attribute("sz")?.Value, out var sz))
            Size = sz / 1000.0;

        // Parse color
        var colorElement = element.Element(element.Name.Namespace + "srgbClr")
            ?? element.Element(element.Name.Namespace + "schemeClr");
        if (colorElement != null)
        {
            Color = Shape.ParseColor(colorElement);
        }
    }
}

public class FormatScheme
{
    private readonly XElement? _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Name { get; }
    public List<FillStyle> FillStyles { get; } = new();
    public List<LineStyle> LineStyles { get; } = new();
    public List<EffectStyle> EffectStyles { get; } = new();
    public List<BackgroundFillStyle> BackgroundFillStyles { get; } = new();

    public FormatScheme(XElement? element)
    {
        _element = element;
        Name = element?.Attribute("name")?.Value;

        // Parse fill styles
        var fillStyleLst = element?.Element(A + "fillStyleLst");
        if (fillStyleLst != null)
        {
            foreach (var fill in fillStyleLst.Elements())
            {
                FillStyles.Add(new FillStyle(fill));
            }
        }

        // Parse line styles
        var lnStyleLst = element?.Element(A + "lnStyleLst");
        if (lnStyleLst != null)
        {
            foreach (var line in lnStyleLst.Elements())
            {
                LineStyles.Add(new LineStyle(line));
            }
        }

        // Parse effect styles
        var effectStyleLst = element?.Element(A + "effectStyleLst");
        if (effectStyleLst != null)
        {
            foreach (var effect in effectStyleLst.Elements())
            {
                EffectStyles.Add(new EffectStyle(effect));
            }
        }

        // Parse background fill styles
        var bgFillStyleLst = element?.Element(A + "bgFillStyleLst");
        if (bgFillStyleLst != null)
        {
            foreach (var bgFill in bgFillStyleLst.Elements())
            {
                BackgroundFillStyles.Add(new BackgroundFillStyle(bgFill));
            }
        }
    }
}

public class FillStyle
{
    private readonly XElement _element;

    public FillType Type { get; }
    public Fill? Fill { get; }

    public FillStyle(XElement element)
    {
        _element = element;
        Type = ParseFillType(element.Name.LocalName);
        Fill = ParseFillElement(element);
    }

    private static FillType ParseFillType(string name)
    {
        return name switch
        {
            "solidFill" => FillType.Solid,
            "gradFill" => FillType.Gradient,
            "pattFill" => FillType.Pattern,
            "blipFill" => FillType.Picture,
            "noFill" => FillType.None,
            _ => FillType.None
        };
    }

    private static Fill? ParseFillElement(XElement element)
    {
        var wrapper = new XElement(element.Name.Namespace + "spPr", new XElement(element));
        return Shape.ParseFill(wrapper);
    }
}

public class LineStyle
{
    private readonly XElement _element;

    public double Width { get; }
    public LineCap Cap { get; }
    public LineJoin Join { get; }
    public Outline? Outline { get; }

    public LineStyle(XElement element)
    {
        _element = element;

        if (double.TryParse(element.Attribute("w")?.Value, out var w))
            Width = w / 12700.0;

        Cap = element.Attribute("cap")?.Value switch
        {
            "rnd" => LineCap.Round,
            "sq" => LineCap.Square,
            _ => LineCap.Flat
        };

        Join = element.Attribute("algn")?.Value switch
        {
            "bevel" => LineJoin.Bevel,
            "round" => LineJoin.Round,
            _ => LineJoin.Miter
        };

        Outline = ParseOutline(element);
    }

    private static Outline? ParseOutline(XElement element)
    {
        var solidFill = element.Element(element.Name.Namespace + "solidFill");
        if (solidFill == null) return null;

        var color = Shape.ParseColor(solidFill);
        if (color == null) return null;

        return new Outline
        {
            Width = int.TryParse(element.Attribute("w")?.Value, out var w) ? w : 12700,
            Color = color
        };
    }
}

public class BackgroundFillStyle
{
    private readonly XElement _element;

    public FillType Type { get; }
    public Fill? Fill { get; }

    public BackgroundFillStyle(XElement element)
    {
        _element = element;
        Type = ParseFillType(element.Name.LocalName);
        Fill = ParseFillElement(element);
    }

    private static FillType ParseFillType(string name)
    {
        return name switch
        {
            "solidFill" => FillType.Solid,
            "gradFill" => FillType.Gradient,
            "pattFill" => FillType.Pattern,
            "blipFill" => FillType.Picture,
            "noFill" => FillType.None,
            _ => FillType.None
        };
    }

    private static Fill? ParseFillElement(XElement element)
    {
        var wrapper = new XElement(element.Name.Namespace + "spPr", new XElement(element));
        return Shape.ParseFill(wrapper);
    }
}
