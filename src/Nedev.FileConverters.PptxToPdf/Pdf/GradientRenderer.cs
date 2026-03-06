using System.Text;

namespace Nedev.FileConverters.PptxToPdf.Pdf;

public class GradientRenderer
{
    /// <summary>
    /// Renders a linear gradient to PDF shading pattern
    /// </summary>
    public static string RenderLinearGradient(
        double x, double y, double width, double height,
        Gradient gradient,
        double angle = 0)
    {
        var sb = new StringBuilder();

        // Calculate gradient vector based on angle
        var radians = angle * Math.PI / 180;
        var dx = Math.Cos(radians) * width;
        var dy = Math.Sin(radians) * height;

        // Create shading dictionary
        sb.AppendLine("<<");
        sb.AppendLine("  /ShadingType 2"); // Axial shading
        sb.AppendLine("  /ColorSpace /DeviceRGB");
        sb.AppendLine($"  /Coords [{x} {y} {x + dx} {y + dy}]");
        sb.AppendLine("  /Function <<");
        sb.AppendLine("    /FunctionType 2"); // Exponential interpolation
        sb.AppendLine("    /Domain [0 1]");
        sb.AppendLine("    /C0 [" + FormatColor(gradient.Stops.First().Color) + "]");
        sb.AppendLine("    /C1 [" + FormatColor(gradient.Stops.Last().Color) + "]");
        sb.AppendLine("    /N 1");
        sb.AppendLine("  >>");
        sb.AppendLine("  /Extend [true true]");
        sb.AppendLine(">>");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a radial gradient to PDF shading pattern
    /// </summary>
    public static string RenderRadialGradient(
        double x, double y, double width, double height,
        Gradient gradient,
        double centerX = 0.5, double centerY = 0.5)
    {
        var sb = new StringBuilder();

        // Calculate center and radius
        var cx = x + width * centerX;
        var cy = y + height * centerY;
        var r = Math.Min(width, height) / 2;

        // Create shading dictionary
        sb.AppendLine("<<");
        sb.AppendLine("  /ShadingType 3"); // Radial shading
        sb.AppendLine("  /ColorSpace /DeviceRGB");
        sb.AppendLine($"  /Coords [{cx} {cy} 0 {cx} {cy} {r}]");
        sb.AppendLine("  /Function <<");
        sb.AppendLine("    /FunctionType 2");
        sb.AppendLine("    /Domain [0 1]");
        sb.AppendLine("    /C0 [" + FormatColor(gradient.Stops.First().Color) + "]");
        sb.AppendLine("    /C1 [" + FormatColor(gradient.Stops.Last().Color) + "]");
        sb.AppendLine("    /N 1");
        sb.AppendLine("  >>");
        sb.AppendLine("  /Extend [true true]");
        sb.AppendLine(">>");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a multi-stop gradient using stitching functions
    /// </summary>
    public static string RenderMultiStopGradient(
        double x, double y, double width, double height,
        Gradient gradient,
        bool isRadial = false)
    {
        var sb = new StringBuilder();
        var stops = gradient.Stops.OrderBy(s => s.Position).ToList();

        if (stops.Count < 2)
        {
            // Fallback to solid color
            return "";
        }

        // Create stitching function for multiple stops
        sb.AppendLine("<<");
        sb.AppendLine("  /FunctionType 3"); // Stitching function
        sb.AppendLine("  /Domain [0 1]");
        sb.Append("  /Functions [");

        // Create a function for each segment
        for (int i = 0; i < stops.Count - 1; i++)
        {
            sb.Append("<<");
            sb.Append("/FunctionType 2 ");
            sb.Append("/Domain [0 1] ");
            sb.Append($"/C0 [{FormatColor(stops[i].Color)}] ");
            sb.Append($"/C1 [{FormatColor(stops[i + 1].Color)}] ");
            sb.Append("/N 1");
            sb.Append(">> ");
        }

        sb.AppendLine("]");

        // Calculate bounds for each segment
        sb.Append("  /Bounds [");
        for (int i = 1; i < stops.Count - 1; i++)
        {
            sb.Append(stops[i].Position);
            if (i < stops.Count - 2) sb.Append(" ");
        }
        sb.AppendLine("]");

        // Encode values
        sb.Append("  /Encode [");
        for (int i = 0; i < stops.Count - 1; i++)
        {
            sb.Append("0 1 ");
        }
        sb.AppendLine("]");
        sb.AppendLine(">>");

        return sb.ToString();
    }

    /// <summary>
    /// Creates a PDF pattern for complex gradients
    /// </summary>
    public static string CreateGradientPattern(
        string shadingName,
        double x, double y, double width, double height)
    {
        var sb = new StringBuilder();

        sb.AppendLine("<<");
        sb.AppendLine("  /Type /Pattern");
        sb.AppendLine("  /PatternType 2"); // Shading pattern
        sb.AppendLine($"  /Shading {shadingName}");
        sb.AppendLine("  /Matrix [1 0 0 1 0 0]");
        sb.AppendLine(">>");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a rectangular gradient (special case of linear)
    /// </summary>
    public static string RenderRectangularGradient(
        double x, double y, double width, double height,
        Gradient gradient)
    {
        // Rectangular gradient can be approximated with multiple linear gradients
        // or using a custom Type 4 or Type 5 shading
        var sb = new StringBuilder();

        sb.AppendLine("<<");
        sb.AppendLine("  /ShadingType 6"); // Coons patch mesh
        sb.AppendLine("  /ColorSpace /DeviceRGB");
        sb.AppendLine("  /BitsPerCoordinate 32");
        sb.AppendLine("  /BitsPerComponent 8");
        sb.AppendLine("  /BitsPerFlag 8");
        sb.AppendLine("  /Decode [" +
            $"{x} {x + width} " + // X range
            $"{y} {y + height} " + // Y range
            "0 1 0 1 0 1]"); // RGB ranges
        sb.AppendLine("  /Data <<");
        // Simplified: just use corner colors
        var corners = new[]
        {
            gradient.Stops.First().Color,
            gradient.Stops.First().Color,
            gradient.Stops.Last().Color,
            gradient.Stops.Last().Color
        };
        sb.AppendLine("    // Corner colors would be encoded here");
        sb.AppendLine("  >>");
        sb.AppendLine(">>");

        return sb.ToString();
    }

    /// <summary>
    /// Renders a path gradient (conical gradient approximation)
    /// </summary>
    public static string RenderPathGradient(
        double x, double y, double width, double height,
        Gradient gradient,
        List<(double X, double Y)> pathPoints)
    {
        // Path gradient is complex in PDF, approximate with multiple radial gradients
        var sb = new StringBuilder();

        var centerX = x + width / 2;
        var centerY = y + height / 2;

        sb.AppendLine("<<");
        sb.AppendLine("  /ShadingType 7"); // Tensor product patch mesh
        sb.AppendLine("  /ColorSpace /DeviceRGB");
        sb.AppendLine("  /BitsPerCoordinate 32");
        sb.AppendLine("  /BitsPerComponent 8");
        sb.AppendLine("  /BitsPerFlag 8");
        sb.AppendLine($"  /Decode [{x} {x + width} {y} {y + height} 0 1 0 1 0 1]");
        sb.AppendLine(">>");

        return sb.ToString();
    }

    private static string FormatColor(Color color)
    {
        // Convert 0-255 to 0-1 range
        var r = color.R / 255.0;
        var g = color.G / 255.0;
        var b = color.B / 255.0;
        return $"{r:F4} {g:F4} {b:F4}";
    }

    /// <summary>
    /// Creates PDF content stream commands for applying a gradient fill
    /// </summary>
    public static string CreateGradientFillCommands(
        string patternName,
        double x, double y, double width, double height)
    {
        var sb = new StringBuilder();

        // Save graphics state
        sb.AppendLine("q");

        // Create clipping path for the shape
        sb.AppendLine($"{x} {y} {width} {height} re"); // Rectangle path
        sb.AppendLine("W"); // Set clipping path
        sb.AppendLine("n"); // End path without filling

        // Set pattern color space and fill
        sb.AppendLine("/Pattern cs");
        sb.AppendLine($"/P{patternName} scn");

        // Fill the entire bounding box
        sb.AppendLine($"{x} {y} {width} {height} re f");

        // Restore graphics state
        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Creates a gradient approximation using multiple stripes (fallback method)
    /// </summary>
    public static string CreateStripedGradient(
        double x, double y, double width, double height,
        Gradient gradient,
        int stripeCount = 20,
        bool horizontal = true)
    {
        var sb = new StringBuilder();
        var stops = gradient.Stops.OrderBy(s => s.Position).ToList();

        if (stops.Count < 2) return "";

        sb.AppendLine("q");

        for (int i = 0; i < stripeCount; i++)
        {
            var t = i / (double)stripeCount;
            var color = InterpolateColor(stops, t);

            // Set color
            sb.AppendLine($"{color.R / 255.0:F4} {color.G / 255.0:F4} {color.B / 255.0:F4} rg");

            // Draw stripe
            if (horizontal)
            {
                var stripeWidth = width / stripeCount;
                var stripeX = x + i * stripeWidth;
                sb.AppendLine($"{stripeX} {y} {stripeWidth} {height} re f");
            }
            else
            {
                var stripeHeight = height / stripeCount;
                var stripeY = y + i * stripeHeight;
                sb.AppendLine($"{x} {stripeY} {width} {stripeHeight} re f");
            }
        }

        sb.AppendLine("Q");

        return sb.ToString();
    }

    private static Color InterpolateColor(List<GradientStop> stops, double position)
    {
        // Find the two stops that bracket the position
        GradientStop? lower = null;
        GradientStop? upper = null;

        foreach (var stop in stops)
        {
            if (stop.Position <= position)
                lower = stop;
            if (stop.Position >= position && upper == null)
            {
                upper = stop;
                break;
            }
        }

        if (lower == null) return upper?.Color ?? Color.Black;
        if (upper == null) return lower.Color;
        if (lower.Position == upper.Position) return lower.Color;

        // Interpolate
        var t = (position - lower.Position) / (upper.Position - lower.Position);
        return new Color(
            (byte)(lower.Color.R + t * (upper.Color.R - lower.Color.R)),
            (byte)(lower.Color.G + t * (upper.Color.G - lower.Color.G)),
            (byte)(lower.Color.B + t * (upper.Color.B - lower.Color.B)),
            (byte)(lower.Color.A + t * (upper.Color.A - lower.Color.A))
        );
    }
}

public class Gradient
{
    public GradientType Type { get; set; }
    public List<GradientStop> Stops { get; set; } = new();
    public double Angle { get; set; }
    public double CenterX { get; set; } = 0.5;
    public double CenterY { get; set; } = 0.5;
    public double FocusX { get; set; } = 0.5;
    public double FocusY { get; set; } = 0.5;
    public double Radius { get; set; } = 0.5;
    public bool TileFlipX { get; set; }
    public bool TileFlipY { get; set; }
}

public class GradientStop
{
    public double Position { get; set; } // 0.0 to 1.0
    public Color Color { get; set; }
}
