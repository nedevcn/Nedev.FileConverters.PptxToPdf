namespace NPptxToPdf;

public class Fill
{
    public FillType Type { get; set; }
    public Color Color { get; set; }

    // Gradient
    public List<GradientStop>? GradientStops { get; set; }
    public GradientType GradientType { get; set; }
    public double GradientAngle { get; set; }

    // Pattern
    public PatternType PatternType { get; set; }
    public Color PatternForegroundColor { get; set; }
    public Color PatternBackgroundColor { get; set; }

    // Picture
    public string? PictureRelationshipId { get; set; }
    public PictureFillMode PictureFillMode { get; set; }
    public PictureFill? PictureFill { get; set; }
}

public class PictureFill
{
    public Blip? Blip { get; set; }
}

public class Blip
{
    public byte[]? Data { get; set; }
}

public enum PictureFillMode
{
    Stretch,
    Tile
}
