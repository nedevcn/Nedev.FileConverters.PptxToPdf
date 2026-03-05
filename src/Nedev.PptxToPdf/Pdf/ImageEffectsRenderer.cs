using System.Globalization;
using System.Text;

namespace Nedev.PptxToPdf.Pdf;

/// <summary>
/// Renders shape/image effects (shadows, reflections, glow, soft edges) to PDF content streams.
/// All transparency is handled via properly registered ExtGState resources on the page.
/// </summary>
public class ImageEffectsRenderer
{
    private readonly PdfDocument _document;
    private readonly PdfPage _page;
    private int _gsCounter;

    public ImageEffectsRenderer(PdfDocument document, PdfPage page)
    {
        _document = document;
        _page = page;
    }

    // ---------------------------------------------------------------------------
    //  Public entry point 鈥?call this before rendering the shape itself
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Renders all applicable pre-shape effects (outer shadow, glow) that must
    /// appear behind the shape/image.
    /// </summary>
    public string RenderPreEffects(
        double x, double y, double width, double height,
        ShapeEffects effects, ShapeType shapeType = ShapeType.Rectangle)
    {
        var sb = new StringBuilder();

        // Glow is drawn first (furthest out)
        if (effects.Glow is { Radius: > 0 } glow)
        {
            RenderGlow(sb, x, y, width, height, glow, shapeType);
        }

        // Outer shadow drawn behind the shape
        if (effects.Shadow is { Type: ShadowType.Outer } shadow)
        {
            RenderOuterShadow(sb, x, y, width, height, shadow, shapeType);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Renders all applicable post-shape effects (inner shadow, soft edge, reflection)
    /// that must appear on top of / clipped to the shape.
    /// </summary>
    public string RenderPostEffects(
        double x, double y, double width, double height,
        ShapeEffects effects, ShapeType shapeType = ShapeType.Rectangle)
    {
        var sb = new StringBuilder();

        // Inner shadow (clipped inside the shape)
        if (effects.Shadow is { Type: ShadowType.Inner } innerShadow)
        {
            RenderInnerShadow(sb, x, y, width, height, innerShadow, shapeType);
        }

        // Soft edge
        if (effects.SoftEdge is { Radius: > 0 } softEdge)
        {
            RenderSoftEdges(sb, x, y, width, height, softEdge.Radius, shapeType);
        }

        // Reflection (drawn below the shape)
        if (effects.Reflection is { } reflection &&
            (reflection.StartOpacity > 0 || reflection.EndOpacity > 0))
        {
            RenderReflection(sb, x, y, width, height, reflection, shapeType);
        }

        return sb.ToString();
    }

    // ---------------------------------------------------------------------------
    //  Outer Shadow
    // ---------------------------------------------------------------------------

    private void RenderOuterShadow(
        StringBuilder sb, double x, double y, double w, double h,
        ShadowEffect shadow, ShapeType shapeType)
    {
        // Direction: OOXML uses 60000ths of a degree; parser already converts to degrees.
        // In PDF, positive Y is up, so we negate the Y component.
        var dirRad = shadow.Direction * Math.PI / 180.0;
        var offsetX = shadow.Distance * Math.Cos(dirRad);
        var offsetY = -shadow.Distance * Math.Sin(dirRad); // negate for PDF coords

        var color = shadow.Color ?? Color.Black;
        var blur = shadow.BlurRadius;

        // Multi-layer blur approximation: draw expanding copies at decreasing opacity
        int layers = blur > 0 ? Math.Max(4, (int)Math.Ceiling(blur)) : 1;
        layers = Math.Min(layers, 12); // cap for perf

        for (int i = layers - 1; i >= 0; i--)
        {
            double t = layers == 1 ? 0 : (double)i / (layers - 1);
            double layerAlpha = 0.40 * (1.0 - t * t); // quadratic falloff, peak ~0.40
            double expansion = blur * t;

            sb.AppendLine("q");
            EmitExtGState(sb, layerAlpha);
            SetFillColor(sb, color);
            EmitShapePath(sb, shapeType,
                x + offsetX - expansion,
                y + offsetY - expansion,
                w + expansion * 2,
                h + expansion * 2,
                fill: true);
            sb.AppendLine("Q");
        }
    }

    // ---------------------------------------------------------------------------
    //  Inner Shadow
    // ---------------------------------------------------------------------------

    private void RenderInnerShadow(
        StringBuilder sb, double x, double y, double w, double h,
        ShadowEffect shadow, ShapeType shapeType)
    {
        var dirRad = shadow.Direction * Math.PI / 180.0;
        var offsetX = shadow.Distance * Math.Cos(dirRad);
        var offsetY = -shadow.Distance * Math.Sin(dirRad);
        var color = shadow.Color ?? Color.Black;
        var blur = Math.Max(shadow.BlurRadius, 2);

        // We draw the inner shadow by:
        // 1. Setting a clip to the shape boundary
        // 2. Drawing a filled "frame" (large rect minus inset shape) at the offset

        sb.AppendLine("q");

        // Clip to the shape boundary
        EmitShapePath(sb, shapeType, x, y, w, h, fill: false);
        sb.AppendLine("W n"); // set clipping, no-paint

        int layers = Math.Max(3, Math.Min((int)Math.Ceiling(blur), 8));
        for (int i = 0; i < layers; i++)
        {
            double t = (double)i / layers;
            double layerAlpha = 0.30 * (1.0 - t);
            double inset = blur * t;

            sb.AppendLine("q");
            EmitExtGState(sb, layerAlpha);
            SetFillColor(sb, color);

            // Draw a large rectangle covering the whole clip area...
            sb.AppendLine($"{(x - blur):F2} {(y - blur):F2} {(w + blur * 2):F2} {(h + blur * 2):F2} re f");

            // ...then "punch out" an inset shape using even-odd rule in a separate step
            // (simplification: we draw the frame as a series of edge strips)
            sb.AppendLine("Q");
        }

        // Draw the core inset shadow with clip
        sb.AppendLine("q");
        EmitExtGState(sb, 0.25);
        SetFillColor(sb, color);

        // Top strip
        sb.AppendLine($"{(x + offsetX):F2} {(y + offsetY + h - blur):F2} {w:F2} {blur:F2} re f");
        // Bottom strip
        sb.AppendLine($"{(x + offsetX):F2} {(y + offsetY):F2} {w:F2} {blur:F2} re f");
        // Left strip
        sb.AppendLine($"{(x + offsetX):F2} {(y + offsetY):F2} {blur:F2} {h:F2} re f");
        // Right strip
        sb.AppendLine($"{(x + offsetX + w - blur):F2} {(y + offsetY):F2} {blur:F2} {h:F2} re f");
        sb.AppendLine("Q");

        sb.AppendLine("Q"); // end outer clip
    }

    // ---------------------------------------------------------------------------
    //  Glow
    // ---------------------------------------------------------------------------

    private void RenderGlow(
        StringBuilder sb, double x, double y, double w, double h,
        GlowEffect glow, ShapeType shapeType)
    {
        var color = glow.Color ?? new Color(255, 215, 0); // default gold-ish
        var radius = glow.Radius;

        // Draw expanding halos from outermost (most transparent) to innermost
        int layers = Math.Max(5, Math.Min((int)Math.Ceiling(radius * 1.5), 16));

        for (int i = layers; i >= 0; i--)
        {
            double t = (double)i / layers;
            double alpha = 0.35 * (1.0 - t * t); // quadratic falloff
            double expansion = radius * t;

            sb.AppendLine("q");
            EmitExtGState(sb, alpha);
            SetFillColor(sb, color);
            EmitShapePath(sb, shapeType,
                x - expansion, y - expansion,
                w + expansion * 2, h + expansion * 2,
                fill: true);
            sb.AppendLine("Q");
        }
    }

    // ---------------------------------------------------------------------------
    //  Soft Edges
    // ---------------------------------------------------------------------------

    private void RenderSoftEdges(
        StringBuilder sb, double x, double y, double w, double h,
        double radius, ShapeType shapeType)
    {
        // Approximate soft edges with semi-transparent strips along all four edges
        int strips = Math.Max(3, Math.Min((int)Math.Ceiling(radius), 10));
        double stripSize = radius / strips;

        // Use page background color (white) to "fade" edges
        var fadeColor = Color.White;

        for (int i = strips - 1; i >= 0; i--)
        {
            double t = (double)i / strips;
            double alpha = t * t * 0.85; // 0 at edge 鈫?higher inward
            double d = stripSize * (strips - i);

            sb.AppendLine("q");
            EmitExtGState(sb, alpha);
            SetFillColor(sb, fadeColor);

            // Top edge
            sb.AppendLine($"{x:F2} {(y + h - d):F2} {w:F2} {stripSize:F2} re f");
            // Bottom edge
            sb.AppendLine($"{x:F2} {(y + d - stripSize):F2} {w:F2} {stripSize:F2} re f");
            // Left edge
            sb.AppendLine($"{(x + d - stripSize):F2} {y:F2} {stripSize:F2} {h:F2} re f");
            // Right edge
            sb.AppendLine($"{(x + w - d):F2} {y:F2} {stripSize:F2} {h:F2} re f");

            sb.AppendLine("Q");
        }
    }

    // ---------------------------------------------------------------------------
    //  Reflection
    // ---------------------------------------------------------------------------

    private void RenderReflection(
        StringBuilder sb, double x, double y, double w, double h,
        ReflectionEffect reflection, ShapeType shapeType)
    {
        // Reflection is drawn below the shape.
        // OOXML reflection has startOpacity (top of reflection) 鈫?endOpacity (bottom).
        var gap = reflection.Distance;
        var reflHeight = h * 0.45; // default 45% of shape height
        if (reflHeight < 2) return;

        double reflY = y - gap - reflHeight; // below the shape in PDF coords

        int strips = Math.Max(6, (int)Math.Ceiling(reflHeight / 2));
        double stripH = reflHeight / strips;

        for (int i = 0; i < strips; i++)
        {
            double t = (double)i / strips; // 0 = near shape, 1 = far from shape
            double alpha = reflection.StartOpacity * (1.0 - t) + reflection.EndOpacity * t;
            alpha = Math.Clamp(alpha, 0, 1);
            if (alpha < 0.01) continue;

            double sy = reflY + reflHeight - stripH * (i + 1); // bottom 鈫?top in PDF

            sb.AppendLine("q");
            EmitExtGState(sb, alpha * 0.6); // scale down a bit for realism
            // Use a slightly muted version of white to simulate reflection of content
            SetFillColor(sb, new Color(200, 200, 200));
            sb.AppendLine($"{x:F2} {sy:F2} {w:F2} {stripH:F2} re f");
            sb.AppendLine("Q");
        }
    }

    // ---------------------------------------------------------------------------
    //  Helpers
    // ---------------------------------------------------------------------------

    /// <summary>
    /// Emits a reference to a properly registered ExtGState for the given alpha.
    /// The GS object is added to both the document and the page resources.
    /// </summary>
    private void EmitExtGState(StringBuilder sb, double alpha)
    {
        alpha = Math.Clamp(alpha, 0, 1);
        var gsName = $"GS{_page.Number}_{_gsCounter++}";

        var gsObj = new PdfExtGState(_document.GetNextObjectNumber(), alpha);
        _document.AddObject(gsObj);
        _page.ExtGStates[gsName] = gsObj;

        sb.AppendLine($"/{gsName} gs");
    }

    private static void SetFillColor(StringBuilder sb, Color c)
    {
        sb.AppendLine($"{c.R / 255.0:F4} {c.G / 255.0:F4} {c.B / 255.0:F4} rg");
    }

    /// <summary>
    /// Emits the path for the given shape type (fill or stroke).
    /// For simple shapes we emit directly; for others we fall back to a rectangle.
    /// </summary>
    private static void EmitShapePath(
        StringBuilder sb, ShapeType shapeType,
        double x, double y, double w, double h, bool fill)
    {
        switch (shapeType)
        {
            case ShapeType.Ellipse:
                EmitEllipsePath(sb, x, y, w, h);
                break;

            case ShapeType.RoundRectangle:
                var r = Math.Min(w, h) * 0.1;
                EmitRoundRectPath(sb, x, y, w, h, r);
                break;

            case ShapeType.Triangle:
                sb.AppendLine($"{x + w * 0.5:F2} {y + h:F2} m");
                sb.AppendLine($"{x + w:F2} {y:F2} l");
                sb.AppendLine($"{x:F2} {y:F2} l h");
                break;

            case ShapeType.Diamond:
                sb.AppendLine($"{x + w * 0.5:F2} {y + h:F2} m");
                sb.AppendLine($"{x + w:F2} {y + h * 0.5:F2} l");
                sb.AppendLine($"{x + w * 0.5:F2} {y:F2} l");
                sb.AppendLine($"{x:F2} {y + h * 0.5:F2} l h");
                break;

            default:
                // Rectangle fallback
                sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re");
                break;
        }

        if (fill)
            sb.AppendLine("f");
    }

    private static void EmitEllipsePath(StringBuilder sb, double x, double y, double w, double h)
    {
        const double kappa = 0.5522847498;
        var ox = w * kappa * 0.5;
        var oy = h * kappa * 0.5;
        var cx = x + w * 0.5;
        var cy = y + h * 0.5;
        var ex = x + w;
        var ey = y + h;

        sb.AppendLine($"{cx:F2} {ey:F2} m");
        sb.AppendLine($"{cx + ox:F2} {ey:F2} {ex:F2} {cy + oy:F2} {ex:F2} {cy:F2} c");
        sb.AppendLine($"{ex:F2} {cy - oy:F2} {cx + ox:F2} {y:F2} {cx:F2} {y:F2} c");
        sb.AppendLine($"{cx - ox:F2} {y:F2} {x:F2} {cy - oy:F2} {x:F2} {cy:F2} c");
        sb.AppendLine($"{x:F2} {cy + oy:F2} {cx - ox:F2} {ey:F2} {cx:F2} {ey:F2} c");
    }

    private static void EmitRoundRectPath(StringBuilder sb, double x, double y, double w, double h, double r)
    {
        r = Math.Min(r, Math.Min(w, h) * 0.5);
        sb.AppendLine($"{x + r:F2} {y:F2} m");
        sb.AppendLine($"{x + w - r:F2} {y:F2} l");
        sb.AppendLine($"{x + w:F2} {y:F2} {x + w:F2} {y + r:F2} {x + w:F2} {y + r:F2} c");
        sb.AppendLine($"{x + w:F2} {y + h - r:F2} l");
        sb.AppendLine($"{x + w:F2} {y + h:F2} {x + w - r:F2} {y + h:F2} {x + w - r:F2} {y + h:F2} c");
        sb.AppendLine($"{x + r:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h:F2} {x:F2} {y + h - r:F2} {x:F2} {y + h - r:F2} c");
        sb.AppendLine($"{x:F2} {y + r:F2} l");
        sb.AppendLine($"{x:F2} {y:F2} {x + r:F2} {y:F2} {x + r:F2} {y:F2} c");
    }
}

// ---------------------------------------------------------------------------
//  PDF ExtGState object 鈥?represents a Graphics State with alpha transparency
// ---------------------------------------------------------------------------

public class PdfExtGState : PdfObject
{
    public double FillAlpha { get; }
    public double StrokeAlpha { get; }

    public PdfExtGState(int number, double alpha) : base(number)
    {
        FillAlpha = alpha;
        StrokeAlpha = alpha;
    }

    public PdfExtGState(int number, double fillAlpha, double strokeAlpha) : base(number)
    {
        FillAlpha = fillAlpha;
        StrokeAlpha = strokeAlpha;
    }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, "/Type /ExtGState");
        WriteLine(stream, $"/CA {StrokeAlpha:F4}");
        WriteLine(stream, $"/ca {FillAlpha:F4}");
        WriteLine(stream, ">>");
    }
}
