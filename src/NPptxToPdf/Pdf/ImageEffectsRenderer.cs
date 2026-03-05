using System.Text;

namespace NPptxToPdf.Pdf;

/// <summary>
/// Renders image effects like shadows, reflections, glow, and soft edges to PDF
/// </summary>
public class ImageEffectsRenderer
{
    /// <summary>
    /// Renders a drop shadow effect
    /// </summary>
    public static string RenderDropShadow(
        double x, double y, double width, double height,
        ShadowEffect shadow)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Calculate shadow offset
        var offsetX = shadow.Distance * Math.Cos(shadow.Direction * Math.PI / 180);
        var offsetY = shadow.Distance * Math.Sin(shadow.Direction * Math.PI / 180);

        // Set shadow color with transparency
        var alpha = shadow.Transparency;
        sb.AppendLine($"/GS{{/Type /ExtGState /CA {alpha} /ca {alpha}}} gs");
        sb.AppendLine($"{shadow.Color.R / 255.0:F4} {shadow.Color.G / 255.0:F4} {shadow.Color.B / 255.0:F4} rg");

        // Draw shadow rectangle
        sb.AppendLine($"{x + offsetX} {y + offsetY} {width} {height} re f");

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders an outer shadow with blur
    /// </summary>
    public static string RenderOuterShadow(
        double x, double y, double width, double height,
        ShadowEffect shadow)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Calculate shadow offset
        var offsetX = shadow.Distance * Math.Cos(shadow.Direction * Math.PI / 180);
        var offsetY = shadow.Distance * Math.Sin(shadow.Direction * Math.PI / 180);

        // For blurred shadow, we use multiple layers with decreasing opacity
        var layers = 5;
        for (int i = 0; i < layers; i++)
        {
            var layerAlpha = shadow.Transparency * (1 - i / (double)layers);
            var blur = shadow.BlurRadius * (i + 1) / layers;

            sb.AppendLine($"/GS{{/Type /ExtGState /CA {layerAlpha:F3} /ca {layerAlpha:F3}}} gs");
            sb.AppendLine($"{shadow.Color.R / 255.0:F4} {shadow.Color.G / 255.0:F4} {shadow.Color.B / 255.0:F4} rg");

            // Draw expanded shadow rectangle
            sb.AppendLine($"{x + offsetX - blur} {y + offsetY - blur} {width + blur * 2} {height + blur * 2} re f");
        }

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders an inner shadow effect
    /// </summary>
    public static string RenderInnerShadow(
        double x, double y, double width, double height,
        ShadowEffect shadow)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Create clipping path for the shape
        sb.AppendLine($"{x} {y} {width} {height} re W n");

        // Calculate shadow direction (opposite of outer shadow)
        var offsetX = -shadow.Distance * Math.Cos(shadow.Direction * Math.PI / 180);
        var offsetY = -shadow.Distance * Math.Sin(shadow.Direction * Math.PI / 180);

        // Set shadow color with transparency
        var alpha = shadow.Transparency;
        sb.AppendLine($"/GS{{/Type /ExtGState /CA {alpha} /ca {alpha}}} gs");
        sb.AppendLine($"{shadow.Color.R / 255.0:F4} {shadow.Color.G / 255.0:F4} {shadow.Color.B / 255.0:F4} rg");

        // Draw shadow offset from the edge
        sb.AppendLine($"{x + offsetX} {y + offsetY} {width} {height} re f");

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a reflection effect
    /// </summary>
    public static string RenderReflection(
        double x, double y, double width, double height,
        ReflectionEffect reflection)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Calculate reflection position (below the image)
        var reflectionY = y - height * reflection.Size - reflection.Distance;

        // Create gradient mask for reflection fade
        var gradientHeight = height * reflection.Size;

        // Draw reflection with gradient transparency
        for (int i = 0; i < 10; i++)
        {
            var t = i / 10.0;
            var alpha = reflection.Transparency * (1 - t);
            var stripY = reflectionY + gradientHeight * t;
            var stripHeight = gradientHeight / 10;

            sb.AppendLine($"/GS{{/Type /ExtGState /CA {alpha:F3} /ca {alpha:F3}}} gs");

            // In a real implementation, we would render the flipped image here
            // For now, just draw a placeholder rectangle
            sb.AppendLine($"{x} {stripY} {width} {stripHeight} re f");
        }

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a glow effect
    /// </summary>
    public static string RenderGlow(
        double x, double y, double width, double height,
        GlowEffect glow)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Draw multiple layers for glow effect
        var layers = 8;
        for (int i = layers; i >= 0; i--)
        {
            var t = i / (double)layers;
            var alpha = glow.Transparency * (1 - t * t); // Quadratic falloff
            var expansion = glow.Radius * t;

            sb.AppendLine($"/GS{{/Type /ExtGState /CA {alpha:F3} /ca {alpha:F3}}} gs");
            sb.AppendLine($"{glow.Color.R / 255.0:F4} {glow.Color.G / 255.0:F4} {glow.Color.B / 255.0:F4} rg");

            // Draw expanded rectangle
            sb.AppendLine($"{x - expansion} {y - expansion} {width + expansion * 2} {height + expansion * 2} re f");
        }

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders soft edges effect
    /// </summary>
    public static string RenderSoftEdges(
        double x, double y, double width, double height,
        double radius)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Draw multiple layers with decreasing opacity for soft edge effect
        var layers = (int)Math.Max(5, radius);
        for (int i = layers; i >= 0; i--)
        {
            var t = i / (double)layers;
            var alpha = 1 - t * t;
            var expansion = radius * (1 - t);

            sb.AppendLine($"/GS{{/Type /ExtGState /CA {alpha:F3} /ca {alpha:F3}}} gs");

            // Draw expanded rectangle
            sb.AppendLine($"{x - expansion} {y - expansion} {width + expansion * 2} {height + expansion * 2} re f");
        }

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a bevel effect
    /// </summary>
    public static string RenderBevel(
        double x, double y, double width, double height,
        BevelEffect bevel)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Draw highlight (top-left)
        sb.AppendLine($"/GS{{/Type /ExtGState /CA 0.5 /ca 0.5}} gs");
        sb.AppendLine($"{bevel.HighlightColor.R / 255.0:F4} {bevel.HighlightColor.G / 255.0:F4} {bevel.HighlightColor.B / 255.0:F4} rg");
        sb.AppendLine($"{x} {y + height - bevel.Width} {bevel.Width} {bevel.Width} re f"); // Top-left corner
        sb.AppendLine($"{x} {y + bevel.Width} {bevel.Width} {height - bevel.Width * 2} re f"); // Left edge
        sb.AppendLine($"{x + bevel.Width} {y + height - bevel.Width} {width - bevel.Width * 2} {bevel.Width} re f"); // Top edge

        // Draw shadow (bottom-right)
        sb.AppendLine($"{bevel.ShadowColor.R / 255.0:F4} {bevel.ShadowColor.G / 255.0:F4} {bevel.ShadowColor.B / 255.0:F4} rg");
        sb.AppendLine($"{x + width - bevel.Width} {y} {bevel.Width} {bevel.Width} re f"); // Bottom-right corner
        sb.AppendLine($"{x + width - bevel.Width} {y + bevel.Width} {bevel.Width} {height - bevel.Width * 2} re f"); // Right edge
        sb.AppendLine($"{x + bevel.Width} {y} {width - bevel.Width * 2} {bevel.Width} re f"); // Bottom edge

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a 3D rotation effect (simulated)
    /// </summary>
    public static string Render3DRotation(
        double x, double y, double width, double height,
        Rotation3DEffect rotation)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Apply transformation matrix for 3D effect
        // This is a simplified approximation
        var cosX = Math.Cos(rotation.X * Math.PI / 180);
        var cosY = Math.Cos(rotation.Y * Math.PI / 180);

        // Scale based on rotation
        var scaleX = cosY;
        var scaleY = cosX;

        var centerX = x + width / 2;
        var centerY = y + height / 2;

        // Translate to center, scale, translate back
        sb.AppendLine($"1 0 0 1 {centerX} {centerY} cm");
        sb.AppendLine($"{scaleX:F4} 0 0 {scaleY:F4} 0 0 cm");
        sb.AppendLine($"1 0 0 1 {-centerX} {-centerY} cm");

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Creates an ExtGState dictionary for transparency
    /// </summary>
    public static string CreateTransparencyState(double alpha)
    {
        return $"<< /Type /ExtGState /CA {alpha:F4} /ca {alpha:F4} >>";
    }

    /// <summary>
    /// Creates a blend mode ExtGState
    /// </summary>
    public static string CreateBlendModeState(string blendMode)
    {
        var validModes = new[] { "Normal", "Multiply", "Screen", "Overlay", "Darken", "Lighten", "ColorDodge", "ColorBurn", "HardLight", "SoftLight", "Difference", "Exclusion" };

        if (!validModes.Contains(blendMode))
            blendMode = "Normal";

        return $"<< /Type /ExtGState /BM /{blendMode} >>";
    }
}

/// <summary>
/// Shadow effect properties
/// </summary>
public class ShadowEffect
{
    public ShadowType Type { get; set; } = ShadowType.Outer;
    public Color Color { get; set; } = Color.Black;
    public double Transparency { get; set; } = 0.5;
    public double Distance { get; set; } = 4;
    public double Direction { get; set; } = 45; // degrees
    public double BlurRadius { get; set; } = 4;
    public double Angle { get; set; } = 45;
    public double Spread { get; set; } = 0;
}

public enum ShadowType
{
    Outer,
    Inner,
    Perspective
}

/// <summary>
/// Reflection effect properties
/// </summary>
public class ReflectionEffect
{
    public double Transparency { get; set; } = 0.5;
    public double Size { get; set; } = 0.5; // 0.0 to 1.0
    public double Distance { get; set; } = 0;
    public double Blur { get; set; } = 0;
    public int Direction { get; set; } = 0; // 0 = bottom, 1 = top, 2 = left, 3 = right
}

/// <summary>
/// Glow effect properties
/// </summary>
public class GlowEffect
{
    public Color Color { get; set; } = Color.White;
    public double Transparency { get; set; } = 0.5;
    public double Radius { get; set; } = 10;
}

/// <summary>
/// Bevel effect properties
/// </summary>
public class BevelEffect
{
    public double Width { get; set; } = 6;
    public double Height { get; set; } = 6;
    public Color HighlightColor { get; set; } = Color.White;
    public Color ShadowColor { get; set; } = Color.Black;
    public double HighlightTransparency { get; set; } = 0.5;
    public double ShadowTransparency { get; set; } = 0.5;
}

/// <summary>
/// 3D rotation effect properties
/// </summary>
public class Rotation3DEffect
{
    public double X { get; set; } = 0; // degrees
    public double Y { get; set; } = 0; // degrees
    public double Z { get; set; } = 0; // degrees
    public double Perspective { get; set; } = 0;
}

/// <summary>
/// Image effect collection
/// </summary>
public class ImageEffects
{
    public ShadowEffect? Shadow { get; set; }
    public ReflectionEffect? Reflection { get; set; }
    public GlowEffect? Glow { get; set; }
    public double SoftEdgeRadius { get; set; }
    public BevelEffect? Bevel { get; set; }
    public Rotation3DEffect? Rotation3D { get; set; }

    public bool HasEffects =>
        Shadow != null ||
        Reflection != null ||
        Glow != null ||
        SoftEdgeRadius > 0 ||
        Bevel != null ||
        Rotation3D != null;
}
