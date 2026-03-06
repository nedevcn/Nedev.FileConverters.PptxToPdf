namespace Nedev.FileConverters.PptxToPdf;

public class Outline
{
    public int Width { get; set; }
    public Color? Color { get; set; }
    public Fill? GradientFill { get; set; }
    public LineCap LineCap { get; set; } = LineCap.Flat;
    public LineJoin LineJoin { get; set; } = LineJoin.Miter;
    public double MiterLimit { get; set; } = 8;
    public CompoundType CompoundType { get; set; } = CompoundType.Single;
    public LineAlignment Alignment { get; set; } = LineAlignment.Outside;
    public LineDashType DashType { get; set; } = LineDashType.Solid;
}

public enum CompoundType
{
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple
}

public enum LineAlignment
{
    Outside,
    Center,
    Inside
}
