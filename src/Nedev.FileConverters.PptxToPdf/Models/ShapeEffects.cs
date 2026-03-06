namespace Nedev.FileConverters.PptxToPdf;

public class ShapeEffects
{
    public ShadowEffect? Shadow { get; set; }
    public ReflectionEffect? Reflection { get; set; }
    public GlowEffect? Glow { get; set; }
    public SoftEdgeEffect? SoftEdge { get; set; }
    public BlurEffect? Blur { get; set; }
}

public class ShadowEffect
{
    public ShadowType Type { get; set; }
    public double BlurRadius { get; set; }
    public double Distance { get; set; }
    public double Direction { get; set; }
    public Color? Color { get; set; }
}

public enum ShadowType
{
    Outer,
    Inner,
    Preset
}

public class ReflectionEffect
{
    public double BlurRadius { get; set; }
    public double Distance { get; set; }
    public double Direction { get; set; }
    public double FadeDirection { get; set; }
    public double StartOpacity { get; set; } = 1;
    public double EndOpacity { get; set; }
}

public class GlowEffect
{
    public double Radius { get; set; }
    public Color? Color { get; set; }
}

public class SoftEdgeEffect
{
    public double Radius { get; set; }
}

public class BlurEffect
{
    public double Radius { get; set; }
}
