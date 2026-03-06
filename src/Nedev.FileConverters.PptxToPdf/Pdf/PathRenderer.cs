using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pdf;

/// <summary>
/// Renders custom path shapes from OpenXML to PDF path commands
/// </summary>
public class PathRenderer
{
    /// <summary>
    /// Converts OpenXML path data to PDF path commands
    /// </summary>
    public static string ConvertPathToPdf(
        XElement pathElement,
        double x, double y,
        double width, double height)
    {
        var sb = new StringBuilder();

        // Parse path data
        var pathData = pathElement.Attribute("path")?.Value;
        if (string.IsNullOrEmpty(pathData))
        {
            // Try to get from child elements
            var pathLst = pathElement.Element(pathElement.Name.Namespace + "pathLst");
            if (pathLst != null)
            {
                var firstPath = pathLst.Element(pathElement.Name.Namespace + "path");
                if (firstPath != null)
                {
                    pathData = firstPath.Attribute("path")?.Value;
                }
            }
        }

        if (string.IsNullOrEmpty(pathData))
            return "";

        // Parse the path data
        var commands = ParsePathData(pathData);

        // Convert to PDF path commands with coordinate transformation
        sb.AppendLine("q");

        foreach (var cmd in commands)
        {
            var pdfCmd = ConvertCommand(cmd, x, y, width, height);
            if (!string.IsNullOrEmpty(pdfCmd))
            {
                sb.AppendLine(pdfCmd);
            }
        }

        sb.AppendLine("Q");

        return sb.ToString();
    }

    /// <summary>
    /// Parses OpenXML path data string into path commands
    /// </summary>
    private static List<PathCommand> ParsePathData(string pathData)
    {
        var commands = new List<PathCommand>();
        var tokens = Tokenize(pathData);

        int i = 0;
        while (i < tokens.Count)
        {
            var token = tokens[i];

            switch (token.ToUpper())
            {
                case "M": // Move to
                    i++;
                    if (i + 1 < tokens.Count)
                    {
                        commands.Add(new PathCommand
                        {
                            Type = PathCommandType.MoveTo,
                            X = ParseCoordinate(tokens[i]),
                            Y = ParseCoordinate(tokens[i + 1])
                        });
                        i += 2;
                    }
                    break;

                case "L": // Line to
                    i++;
                    if (i + 1 < tokens.Count)
                    {
                        commands.Add(new PathCommand
                        {
                            Type = PathCommandType.LineTo,
                            X = ParseCoordinate(tokens[i]),
                            Y = ParseCoordinate(tokens[i + 1])
                        });
                        i += 2;
                    }
                    break;

                case "C": // Cubic bezier
                    i++;
                    if (i + 5 < tokens.Count)
                    {
                        commands.Add(new PathCommand
                        {
                            Type = PathCommandType.CurveTo,
                            X1 = ParseCoordinate(tokens[i]),
                            Y1 = ParseCoordinate(tokens[i + 1]),
                            X2 = ParseCoordinate(tokens[i + 2]),
                            Y2 = ParseCoordinate(tokens[i + 3]),
                            X = ParseCoordinate(tokens[i + 4]),
                            Y = ParseCoordinate(tokens[i + 5])
                        });
                        i += 6;
                    }
                    break;

                case "Q": // Quadratic bezier
                    i++;
                    if (i + 3 < tokens.Count)
                    {
                        commands.Add(new PathCommand
                        {
                            Type = PathCommandType.QuadTo,
                            X1 = ParseCoordinate(tokens[i]),
                            Y1 = ParseCoordinate(tokens[i + 1]),
                            X = ParseCoordinate(tokens[i + 2]),
                            Y = ParseCoordinate(tokens[i + 3])
                        });
                        i += 4;
                    }
                    break;

                case "A": // Arc
                    i++;
                    if (i + 6 < tokens.Count)
                    {
                        commands.Add(new PathCommand
                        {
                            Type = PathCommandType.Arc,
                            RX = ParseCoordinate(tokens[i]),
                            RY = ParseCoordinate(tokens[i + 1]),
                            Rotation = ParseCoordinate(tokens[i + 2]),
                            LargeArcFlag = tokens[i + 3] == "1",
                            SweepFlag = tokens[i + 4] == "1",
                            X = ParseCoordinate(tokens[i + 5]),
                            Y = ParseCoordinate(tokens[i + 6])
                        });
                        i += 7;
                    }
                    break;

                case "Z": // Close path
                    commands.Add(new PathCommand { Type = PathCommandType.ClosePath });
                    i++;
                    break;

                default:
                    i++;
                    break;
            }
        }

        return commands;
    }

    /// <summary>
    /// Tokenizes path data string
    /// </summary>
    private static List<string> Tokenize(string pathData)
    {
        var tokens = new List<string>();
        var sb = new StringBuilder();

        for (int i = 0; i < pathData.Length; i++)
        {
            var c = pathData[i];

            if (char.IsLetter(c))
            {
                if (sb.Length > 0)
                {
                    tokens.Add(sb.ToString());
                    sb.Clear();
                }
                tokens.Add(c.ToString());
            }
            else if (char.IsDigit(c) || c == '.' || c == '-' || c == '+')
            {
                sb.Append(c);
            }
            else if (char.IsWhiteSpace(c) || c == ',')
            {
                if (sb.Length > 0)
                {
                    tokens.Add(sb.ToString());
                    sb.Clear();
                }
            }
        }

        if (sb.Length > 0)
        {
            tokens.Add(sb.ToString());
        }

        return tokens;
    }

    /// <summary>
    /// Parses a coordinate value (handles percentages and formulas)
    /// </summary>
    private static double ParseCoordinate(string token)
    {
        if (token.EndsWith("%"))
        {
            // Percentage value (0-100)
            if (double.TryParse(token.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct))
            {
                return pct / 100.0;
            }
        }
        else if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }

        return 0;
    }

    /// <summary>
    /// Converts a path command to PDF path command
    /// </summary>
    private static string ConvertCommand(PathCommand cmd, double offsetX, double offsetY, double scaleX, double scaleY)
    {
        switch (cmd.Type)
        {
            case PathCommandType.MoveTo:
                return $"{TransformX(cmd.X, offsetX, scaleX):F4} {TransformY(cmd.Y, offsetY, scaleY):F4} m";

            case PathCommandType.LineTo:
                return $"{TransformX(cmd.X, offsetX, scaleX):F4} {TransformY(cmd.Y, offsetY, scaleY):F4} l";

            case PathCommandType.CurveTo:
                return $"{TransformX(cmd.X1, offsetX, scaleX):F4} {TransformY(cmd.Y1, offsetY, scaleY):F4} " +
                       $"{TransformX(cmd.X2, offsetX, scaleX):F4} {TransformY(cmd.Y2, offsetY, scaleY):F4} " +
                       $"{TransformX(cmd.X, offsetX, scaleX):F4} {TransformY(cmd.Y, offsetY, scaleY):F4} c";

            case PathCommandType.QuadTo:
                // Convert quadratic bezier to cubic bezier
                var cp1x = TransformX(cmd.X1, offsetX, scaleX) + 2.0 / 3.0 * (TransformX(cmd.X, offsetX, scaleX) - TransformX(cmd.X1, offsetX, scaleX));
                var cp1y = TransformY(cmd.Y1, offsetY, scaleY) + 2.0 / 3.0 * (TransformY(cmd.Y, offsetY, scaleY) - TransformY(cmd.Y1, offsetY, scaleY));
                var cp2x = TransformX(cmd.X, offsetX, scaleX) + 2.0 / 3.0 * (TransformX(cmd.X1, offsetX, scaleX) - TransformX(cmd.X, offsetX, scaleX));
                var cp2y = TransformY(cmd.Y, offsetY, scaleY) + 2.0 / 3.0 * (TransformY(cmd.Y1, offsetY, scaleY) - TransformY(cmd.Y, offsetY, scaleY));
                return $"{cp1x:F4} {cp1y:F4} {cp2x:F4} {cp2y:F4} {TransformX(cmd.X, offsetX, scaleX):F4} {TransformY(cmd.Y, offsetY, scaleY):F4} c";

            case PathCommandType.Arc:
                // Convert arc to cubic bezier curves
                return ConvertArcToBezier(cmd, offsetX, offsetY, scaleX, scaleY);

            case PathCommandType.ClosePath:
                return "h";

            default:
                return "";
        }
    }

    /// <summary>
    /// Converts an arc to cubic bezier curves
    /// </summary>
    private static string ConvertArcToBezier(PathCommand cmd, double offsetX, double offsetY, double scaleX, double scaleY)
    {
        // Simplified arc to bezier conversion
        // For a complete implementation, use the standard SVG arc-to-bezier algorithm
        var sb = new StringBuilder();

        // Just draw a line to the end point as a simplified approximation
        sb.AppendLine($"{TransformX(cmd.X, offsetX, scaleX):F4} {TransformY(cmd.Y, offsetY, scaleY):F4} l");

        return sb.ToString();
    }

    private static double TransformX(double x, double offsetX, double scale)
    {
        return offsetX + x * scale;
    }

    private static double TransformY(double y, double offsetY, double scale)
    {
        return offsetY + y * scale;
    }

    /// <summary>
    /// Renders a freeform shape from path points
    /// </summary>
    public static string RenderFreeformPath(
        List<(double X, double Y)> points,
        bool closed = true)
    {
        if (points.Count < 2)
            return "";

        var sb = new StringBuilder();

        // Move to first point
        sb.AppendLine($"{points[0].X:F4} {points[0].Y:F4} m");

        // Line to remaining points
        for (int i = 1; i < points.Count; i++)
        {
            sb.AppendLine($"{points[i].X:F4} {points[i].Y:F4} l");
        }

        // Close path if needed
        if (closed)
        {
            sb.AppendLine("h");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Renders a smooth curve through points using Catmull-Rom splines
    /// </summary>
    public static string RenderSmoothPath(
        List<(double X, double Y)> points,
        bool closed = false)
    {
        if (points.Count < 2)
            return "";

        var sb = new StringBuilder();

        // Convert Catmull-Rom to cubic bezier
        var bezierPoints = CatmullRomToBezier(points, closed);

        // Move to first point
        sb.AppendLine($"{bezierPoints[0].X:F4} {bezierPoints[0].Y:F4} m");

        // Draw cubic bezier segments
        for (int i = 1; i < bezierPoints.Count; i += 3)
        {
            if (i + 2 < bezierPoints.Count)
            {
                sb.AppendLine($"{bezierPoints[i].X:F4} {bezierPoints[i].Y:F4} " +
                             $"{bezierPoints[i + 1].X:F4} {bezierPoints[i + 1].Y:F4} " +
                             $"{bezierPoints[i + 2].X:F4} {bezierPoints[i + 2].Y:F4} c");
            }
        }

        if (closed)
        {
            sb.AppendLine("h");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Converts Catmull-Rom spline points to cubic bezier control points
    /// </summary>
    private static List<(double X, double Y)> CatmullRomToBezier(
        List<(double X, double Y)> points,
        bool closed)
    {
        var result = new List<(double X, double Y)>();
        var tension = 0.5; // Catmull-Rom tension parameter

        int count = points.Count;
        for (int i = 0; i < count; i++)
        {
            var p0 = points[(i - 1 + count) % count];
            var p1 = points[i];
            var p2 = points[(i + 1) % count];
            var p3 = points[(i + 2) % count];

            if (i == 0 && !closed)
            {
                p0 = p1;
            }
            if (i == count - 1 && !closed)
            {
                p3 = p2;
            }

            // Calculate control points
            var cp1x = p1.X + (p2.X - p0.X) * tension / 3.0;
            var cp1y = p1.Y + (p2.Y - p0.Y) * tension / 3.0;
            var cp2x = p2.X - (p3.X - p1.X) * tension / 3.0;
            var cp2y = p2.Y - (p3.Y - p1.Y) * tension / 3.0;

            if (i == 0)
            {
                result.Add(p1);
            }

            result.Add((cp1x, cp1y));
            result.Add((cp2x, cp2y));
            result.Add(p2);
        }

        return result;
    }

    /// <summary>
    /// Creates a path from geometric adjustments
    /// </summary>
    public static string CreateAdjustedPath(
        ShapeType baseShape,
        List<GeometryAdjustment> adjustments,
        double x, double y,
        double width, double height)
    {
        // Apply adjustments to base shape
        var adjustedPath = ApplyAdjustments(baseShape, adjustments);

        // Scale and translate to target position
        var scaledPath = adjustedPath.Select(p => (
            X: x + p.X * width,
            Y: y + p.Y * height
        )).ToList();

        return RenderFreeformPath(scaledPath, true);
    }

    /// <summary>
    /// Applies geometry adjustments to a base shape
    /// </summary>
    private static List<(double X, double Y)> ApplyAdjustments(
        ShapeType shape,
        List<GeometryAdjustment> adjustments)
    {
        // This would contain logic for each shape type
        // For now, return default shape points
        return GetDefaultShapePoints(shape);
    }

    /// <summary>
    /// Gets default points for a shape type
    /// </summary>
    private static List<(double X, double Y)> GetDefaultShapePoints(ShapeType shape)
    {
        return shape switch
        {
            ShapeType.Rectangle => new List<(double X, double Y)>
            {
                (0, 0), (1, 0), (1, 1), (0, 1)
            },
            ShapeType.Ellipse => CreateEllipsePoints(),
            ShapeType.Triangle => new List<(double X, double Y)>
            {
                (0.5, 0), (1, 1), (0, 1)
            },
            ShapeType.Diamond => new List<(double X, double Y)>
            {
                (0.5, 0), (1, 0.5), (0.5, 1), (0, 0.5)
            },
            ShapeType.Pentagon => CreatePolygonPoints(5),
            ShapeType.Hexagon => CreatePolygonPoints(6),
            ShapeType.Octagon => CreatePolygonPoints(8),
            ShapeType.Star5 => CreateStarPoints(5, 0.5),
            ShapeType.Star6 => CreateStarPoints(6, 0.5),
            _ => new List<(double X, double Y)> { (0, 0), (1, 0), (1, 1), (0, 1) }
        };
    }

    private static List<(double X, double Y)> CreateEllipsePoints()
    {
        var points = new List<(double X, double Y)>();
        const int segments = 32;

        for (int i = 0; i < segments; i++)
        {
            var angle = 2 * Math.PI * i / segments;
            var x = 0.5 + 0.5 * Math.Cos(angle);
            var y = 0.5 + 0.5 * Math.Sin(angle);
            points.Add((x, y));
        }

        return points;
    }

    private static List<(double X, double Y)> CreatePolygonPoints(int sides)
    {
        var points = new List<(double X, double Y)>();

        for (int i = 0; i < sides; i++)
        {
            var angle = 2 * Math.PI * i / sides - Math.PI / 2;
            var x = 0.5 + 0.5 * Math.Cos(angle);
            var y = 0.5 + 0.5 * Math.Sin(angle);
            points.Add((x, y));
        }

        return points;
    }

    private static List<(double X, double Y)> CreateStarPoints(int points, double innerRadius)
    {
        var result = new List<(double X, double Y)>();
        var totalPoints = points * 2;

        for (int i = 0; i < totalPoints; i++)
        {
            var angle = 2 * Math.PI * i / totalPoints - Math.PI / 2;
            var radius = (i % 2 == 0) ? 0.5 : 0.5 * innerRadius;
            var x = 0.5 + radius * Math.Cos(angle);
            var y = 0.5 + radius * Math.Sin(angle);
            result.Add((x, y));
        }

        return result;
    }
}

/// <summary>
/// Represents a path command
/// </summary>
public class PathCommand
{
    public PathCommandType Type { get; set; }
    public double X { get; set; }
    public double Y { get; set; }
    public double X1 { get; set; }
    public double Y1 { get; set; }
    public double X2 { get; set; }
    public double Y2 { get; set; }
    public double RX { get; set; }
    public double RY { get; set; }
    public double Rotation { get; set; }
    public bool LargeArcFlag { get; set; }
    public bool SweepFlag { get; set; }
}

public enum PathCommandType
{
    MoveTo,
    LineTo,
    CurveTo,
    QuadTo,
    Arc,
    ClosePath
}

/// <summary>
/// Geometry adjustment for custom shapes
/// </summary>
public class GeometryAdjustment
{
    public string Name { get; set; } = "";
    public double Value { get; set; }
    public double Min { get; set; }
    public double Max { get; set; }
}
