namespace Nedev.FileConverters.PptxToPdf;

public class Paragraph
{
    public TextAlignment? Alignment { get; set; }
    public int Level { get; set; }
    public long DefaultTabSize { get; set; } = 914400;
    public bool RightToLeft { get; set; }
    public bool EastAsianLineBreak { get; set; } = true;
    public bool LatinLineBreak { get; set; }
    public bool HangingPunctuation { get; set; }

    // Indentation
    public long MarginLeft { get; set; }
    public long MarginRight { get; set; }
    public long Indent { get; set; }

    // Spacing
    public Spacing? LineSpacing { get; set; }
    public Spacing? SpaceBefore { get; set; }
    public Spacing? SpaceAfter { get; set; }

    // Bullet
    public BulletType BulletType { get; set; } = BulletType.None;
    public bool HasExplicitBulletDefinition { get; set; }
    public string? BulletChar { get; set; }
    public string? BulletAutoNumberType { get; set; }
    public int BulletStartAt { get; set; } = 1;
    public double BulletSize { get; set; } = 1;
    public string? BulletFont { get; set; }
    public Color? BulletColor { get; set; }

    // Runs
    public List<Run> Runs { get; set; } = new();

    public string GetFullText()
    {
        return string.Join("", Runs.Select(r => r.Text));
    }
}

public class Run
{
    public string? Text { get; set; }
    public RunProperties? Properties { get; set; }
}

public class RunProperties
{
    public int? FontSize { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public UnderlineType Underline { get; set; } = UnderlineType.None;
    public StrikeType Strike { get; set; } = StrikeType.None;
    public CapsType Caps { get; set; } = CapsType.None;
    public string? FontFamily { get; set; }
    public string? EastAsianFont { get; set; }
    public string? ComplexScriptFont { get; set; }
    public Color? Color { get; set; }
    public Color? HighlightColor { get; set; }
    public string? Language { get; set; }
    public double BaselineOffset { get; set; }
}

public class Spacing
{
    public double? Percent { get; set; }
    public double? Points { get; set; }
}
