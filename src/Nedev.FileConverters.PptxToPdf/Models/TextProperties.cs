namespace Nedev.FileConverters.PptxToPdf;

public class TextProperties
{
    public int? FontSize { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? FontFamily { get; set; }
    public Color? Color { get; set; }
    public TextAlignment? Alignment { get; set; }

    // Extended properties
    public TextDirection TextDirection { get; set; } = TextDirection.Horizontal;
    public TextAnchor Anchor { get; set; } = TextAnchor.Middle;
    public bool WrapText { get; set; } = true;

    // Margins (in inches)
    public double LeftInset { get; set; } = 0.1;
    public double TopInset { get; set; } = 0.05;
    public double RightInset { get; set; } = 0.1;
    public double BottomInset { get; set; } = 0.05;

    // Auto-fit
    public TextAutoFit AutoFit { get; set; } = TextAutoFit.None;
    public double FontScale { get; set; } = 1;
    public double LineSpaceReduction { get; set; }
}

public enum TextAutoFit
{
    None,
    Normal,
    Shape
}
