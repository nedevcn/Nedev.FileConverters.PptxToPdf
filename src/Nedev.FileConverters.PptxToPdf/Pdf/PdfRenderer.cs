using System.Globalization;
using System.Text;
using Nedev.FileConverters.PptxToPdf;
using Nedev.FileConverters.PptxToPdf.Image;
using Nedev.FileConverters.PptxToPdf.Pptx;

namespace Nedev.FileConverters.PptxToPdf.Pdf;

public class PdfRenderer
{
    private readonly PdfDocument _document;
    private readonly Dictionary<string, PdfFont> _fonts = new();
    private readonly FontEmbedder _fontEmbedder;
    private PdfPage? _currentPage;

    public PdfRenderer(PdfDocument document)
    {
        _document = document;
        _fontEmbedder = new FontEmbedder(document);
    }

    public void RenderSlide(PdfPage page, Slide slide, PptxDocument pptx)
    {
        _currentPage = page;
        try
        {
            var content = new PdfContent(_document.GetNextObjectNumber());
            _document.AddObject(content);
            page.Content = content;

            var sb = new StringBuilder();

            sb.AppendLine("q");

            var slideTheme = slide.GetEffectiveTheme(pptx.Theme);
            var background = slide.GetEffectiveBackground();
            var backgroundFill = background?.ResolveFill(slideTheme);
            var backgroundSourcePath = background?.ResolveSourcePath(slideTheme) ?? slide.SourcePath;
            if (backgroundFill != null)
            {
                RenderBackground(sb, backgroundFill, backgroundSourcePath, page.Width, page.Height, pptx, page);
            }

            // Render connectors first (behind shapes)
            foreach (var connector in slide.Connectors)
            {
                RenderConnector(sb, connector, page.Height, page);
            }

            // Render shapes
            foreach (var shape in slide.GetRenderableShapes())
            {
                RenderShape(sb, shape, page.Height, slide, pptx, page);
            }

            // Render pictures
            foreach (var picture in slide.GetRenderablePictures())
            {
                RenderPicture(sb, picture, slide, page.Height, pptx, page);
            }

            // Render charts
            foreach (var chart in slide.Charts)
            {
                RenderChart(sb, chart, page.Height, page);
            }

            // Render SmartArt
            foreach (var smartArt in slide.SmartArts)
            {
                RenderSmartArt(sb, smartArt, page.Height);
            }

            // Render group shapes
            foreach (var group in slide.GetRenderableGroupShapes())
            {
                RenderGroupShape(sb, group, slide, page.Height, pptx, page);
            }

            // Render tables
            foreach (var table in slide.Tables)
            {
                RenderTable(sb, table, page.Height, page);
            }

            sb.AppendLine("Q");

            var operations = sb.ToString();
            var bytes = Encoding.UTF8.GetBytes(operations);
            content.Stream.Write(bytes, 0, bytes.Length);
        }
        finally
        {
            _currentPage = null;
        }
    }

    private void RenderBackground(StringBuilder sb, Fill fill, string sourcePartPath, double pageWidth, double pageHeight, PptxDocument pptx, PdfPage page)
    {
        switch (fill.Type)
        {
            case FillType.Solid:
                sb.AppendLine("q");
                SetColor(sb, page, fill.Color);
                sb.AppendLine($"0 0 {pageWidth:F2} {pageHeight:F2} re f");
                sb.AppendLine("Q");
                break;

            case FillType.Gradient:
                RenderGradientFill(sb, fill, 0, 0, pageWidth, pageHeight, ShapeType.Rectangle, page);
                break;

            case FillType.Pattern:
                sb.AppendLine("q");
                SetColor(sb, page, fill.PatternForegroundColor);
                sb.AppendLine($"0 0 {pageWidth:F2} {pageHeight:F2} re f");
                sb.AppendLine("Q");
                break;

            case FillType.Picture:
                RenderPictureFill(sb, fill, 0, 0, pageWidth, pageHeight, ShapeType.Rectangle, sourcePartPath, pptx, page);
                break;
        }
    }

    private void RenderConnector(StringBuilder sb, Connector connector, double pageHeight, PdfPage page)
    {
        if (connector.Outline == null || connector.Outline.Width <= 0) return;

        var x1 = connector.Bounds.XPoints;
        var y1 = pageHeight - connector.Bounds.YPoints;
        var x2 = x1 + connector.Bounds.WidthPoints;
        var y2 = y1 - connector.Bounds.HeightPoints;

        sb.AppendLine("q");
        var strokeColor = connector.Outline.Color ?? Color.Black;
        SetStrokeColor(sb, page, strokeColor);
        sb.AppendLine($"{connector.Outline.Width / 12700.0 * 72:F2} w");

        // Set dash pattern if needed
        if (connector.Outline.DashType != LineDashType.Solid)
        {
            var dashArray = GetDashArray(connector.Outline.DashType, connector.Outline.Width);
            sb.AppendLine($"[{dashArray}] 0 d");
        }

        sb.AppendLine($"{x1:F2} {y1:F2} m");
        sb.AppendLine($"{x2:F2} {y2:F2} l S");

        // Reset dash pattern
        if (connector.Outline.DashType != LineDashType.Solid)
        {
            sb.AppendLine("[] 0 d");
        }
        sb.AppendLine("Q");
    }

    private void RenderShape(StringBuilder sb, Shape shape, double pageHeight, Slide slide, PptxDocument pptx, PdfPage page)
    {
        if (shape.ShapeType == ShapeType.Line)
        {
            RenderLine(sb, shape, pageHeight, page);
            return;
        }

        var x = shape.Bounds.XPoints;
        var y = pageHeight - shape.Bounds.YPoints - shape.Bounds.HeightPoints;
        var w = shape.Bounds.WidthPoints;
        var h = shape.Bounds.HeightPoints;

        // Apply transform if present
        if (shape.Transform != null)
        {
            sb.AppendLine("q");
            ApplyTransform(sb, shape.Transform, x + w / 2, y + h / 2);
        }

        // Render pre-shape effects (like outer shadow/glow)
        if (shape.Effects != null)
        {
            var effectsRenderer = new ImageEffectsRenderer(_document, page);
            var pre = effectsRenderer.RenderPreEffects(x, y, w, h, shape.Effects, shape.ShapeType);
            if (!string.IsNullOrEmpty(pre)) sb.Append(pre);
        }

        // Render fill
        RenderShapeFill(sb, shape, x, y, w, h, slide, pptx, page);

        // Render outline
        RenderShapeOutline(sb, shape, x, y, w, h, page);

        // Render post-shape effects (like inner shadow, soft edges, reflection)
        if (shape.Effects != null)
        {
            var effectsRenderer = new ImageEffectsRenderer(_document, page);
            var post = effectsRenderer.RenderPostEffects(x, y, w, h, shape.Effects, shape.ShapeType);
            if (!string.IsNullOrEmpty(post)) sb.Append(post);
        }

        // Render text
        if (!string.IsNullOrEmpty(shape.Text))
        {
            RenderText(sb, shape, pageHeight, page);
        }

        // Render hyperlink
        if (shape.Hyperlink != null && shape.Hyperlink.IsExternal && shape.Hyperlink.Target != null)
        {
            var linkX = shape.Bounds.XPoints;
            var linkY = pageHeight - shape.Bounds.YPoints - shape.Bounds.HeightPoints;
            var linkW = shape.Bounds.WidthPoints;
            var linkH = shape.Bounds.HeightPoints;
            
            // Create link annotation
            var action = new PdfAction(_document.GetNextObjectNumber())
            {
                Type = "/Action",
                S = "/URI",
                URI = shape.Hyperlink.Target
            };
            
            var annotation = new PdfAnnotation(_document.GetNextObjectNumber())
            {
                Type = "/Annot",
                Subtype = "/Link",
                Rect = new[] { linkX, linkY, linkX + linkW, linkY + linkH },
                Action = action
            };
            
            _document.AddObject(action);
            _document.AddObject(annotation);
            page.Annotations.Add(annotation);
        }

        if (shape.Transform != null)
        {
            sb.AppendLine("Q");
        }
    }

    private void RenderShapeFill(StringBuilder sb, Shape shape, double x, double y, double w, double h, Slide slide, PptxDocument pptx, PdfPage page)
    {
        if (shape.Fill == null) return;

        switch (shape.Fill.Type)
        {
            case FillType.Solid:
                sb.AppendLine("q");
                SetColor(sb, page, shape.Fill.Color);
                RenderShapePath(sb, shape.ShapeType, x, y, w, h, true);
                sb.AppendLine("Q");
                break;

            case FillType.Gradient:
                // Render gradient fill using striped approximation
                if (shape.Fill.GradientStops?.Any() == true)
                {
                    RenderGradientFill(sb, shape.Fill, x, y, w, h, shape.ShapeType, page);
                }
                break;

            case FillType.Pattern:
                // Pattern fill - use foreground color
                sb.AppendLine("q");
                SetColor(sb, page, shape.Fill.PatternForegroundColor);
                RenderShapePath(sb, shape.ShapeType, x, y, w, h, true);
                sb.AppendLine("Q");
                break;

            case FillType.Picture:
                // Picture fill
                if (!string.IsNullOrEmpty(shape.Fill.PictureRelationshipId))
                {
                    RenderPictureFill(sb, shape.Fill, x, y, w, h, shape.ShapeType, shape.SourcePath ?? slide.SourcePath, pptx, page);
                }
                break;

            case FillType.None:
                // No fill
                break;
        }
    }

    private void RenderShapeOutline(StringBuilder sb, Shape shape, double x, double y, double w, double h, PdfPage page)
    {
        if (shape.Outline == null || shape.Outline.Width <= 0) return;

        sb.AppendLine("q");
        var strokeColor = shape.Outline.Color ?? Color.Black;
        SetStrokeColor(sb, page, strokeColor);
        sb.AppendLine($"{shape.Outline.Width / 12700.0 * 72:F2} w");

        // Set line cap
        var lineCap = shape.Outline.LineCap switch
        {
            LineCap.Round => "1",
            LineCap.Square => "2",
            _ => "0"
        };
        sb.AppendLine($"{lineCap} J");

        // Set line join
        var lineJoin = shape.Outline.LineJoin switch
        {
            LineJoin.Round => "1",
            LineJoin.Bevel => "2",
            _ => "0"
        };
        sb.AppendLine($"{lineJoin} j");

        // Set dash pattern
        if (shape.Outline.DashType != LineDashType.Solid)
        {
            var dashArray = GetDashArray(shape.Outline.DashType, shape.Outline.Width);
            sb.AppendLine($"[{dashArray}] 0 d");
        }

        RenderShapePath(sb, shape.ShapeType, x, y, w, h, false);
        sb.AppendLine("S");

        // Reset dash pattern
        if (shape.Outline.DashType != LineDashType.Solid)
        {
            sb.AppendLine("[] 0 d");
        }
        sb.AppendLine("Q");
    }

    private void RenderShapePath(StringBuilder sb, ShapeType shapeType, double x, double y, double w, double h, bool fill)
    {
        switch (shapeType)
        {
            case ShapeType.Rectangle:
            case ShapeType.AutoShape:
            case ShapeType.TextBox:
                sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re");
                if (fill) sb.AppendLine("f");
                break;

            case ShapeType.Ellipse:
                RenderEllipse(sb, x, y, w, h, fill);
                break;

            case ShapeType.RoundRectangle:
                var r = Math.Min(w, h) * 0.1;
                RenderRoundRect(sb, x, y, w, h, r, fill);
                break;

            case ShapeType.Triangle:
                RenderTriangle(sb, x, y, w, h, fill);
                break;

            case ShapeType.Diamond:
                RenderDiamond(sb, x, y, w, h, fill);
                break;

            case ShapeType.Pentagon:
                RenderPolygon(sb, x, y, w, h, 5, fill);
                break;

            case ShapeType.Hexagon:
                RenderPolygon(sb, x, y, w, h, 6, fill);
                break;

            case ShapeType.Octagon:
                RenderPolygon(sb, x, y, w, h, 8, fill);
                break;

            case ShapeType.Star5:
                RenderStar(sb, x, y, w, h, 5, fill);
                break;

            case ShapeType.Star6:
                RenderStar(sb, x, y, w, h, 6, fill);
                break;

            case ShapeType.Star8:
                RenderStar(sb, x, y, w, h, 8, fill);
                break;

            case ShapeType.Parallelogram:
                RenderParallelogram(sb, x, y, w, h, fill);
                break;

            case ShapeType.Trapezoid:
                RenderTrapezoid(sb, x, y, w, h, fill);
                break;

            case ShapeType.RightArrow:
                RenderRightArrow(sb, x, y, w, h, fill);
                break;

            case ShapeType.LeftArrow:
                RenderLeftArrow(sb, x, y, w, h, fill);
                break;

            case ShapeType.UpArrow:
                RenderUpArrow(sb, x, y, w, h, fill);
                break;

            case ShapeType.DownArrow:
                RenderDownArrow(sb, x, y, w, h, fill);
                break;



            case ShapeType.Heart:
                RenderHeart(sb, x, y, w, h, fill);
                break;

            case ShapeType.Cloud:
                RenderCloud(sb, x, y, w, h, fill);
                break;

            default:
                // Default to rectangle for unsupported shapes
                sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re");
                if (fill) sb.AppendLine("f");
                break;
        }
    }

    private void ApplyTransform(StringBuilder sb, Transform2D transform, double cx, double cy)
    {
        // Translate to center
        sb.AppendLine($"1 0 0 1 {cx:F2} {cy:F2} cm");

        // Rotate
        if (transform.Rotation != 0)
        {
            var angle = transform.Rotation * Math.PI / 180;
            var cos = Math.Cos(angle);
            var sin = Math.Sin(angle);
            sb.AppendLine($"{cos:F6} {sin:F6} {-sin:F6} {cos:F6} 0 0 cm");
        }

        // Flip
        if (transform.FlipHorizontal)
        {
            sb.AppendLine("-1 0 0 1 0 0 cm");
        }
        if (transform.FlipVertical)
        {
            sb.AppendLine("1 0 0 -1 0 0 cm");
        }

        // Translate back
        sb.AppendLine($"1 0 0 1 {-cx:F2} {-cy:F2} cm");
    }

    private void RenderLine(StringBuilder sb, Shape shape, double pageHeight, PdfPage page)
    {
        if (shape.Outline == null || shape.Outline.Width <= 0) return;

        var x1 = shape.Bounds.XPoints;
        var y1 = pageHeight - shape.Bounds.YPoints;
        var x2 = x1 + shape.Bounds.WidthPoints;
        var y2 = y1 - shape.Bounds.HeightPoints;

        sb.AppendLine("q");
        var strokeColor = shape.Outline.Color ?? Color.Black;
        SetStrokeColor(sb, page, strokeColor);
        sb.AppendLine($"{shape.Outline.Width / 12700.0 * 72:F2} w");

        sb.AppendLine($"{x1:F2} {y1:F2} m");
        sb.AppendLine($"{x2:F2} {y2:F2} l S");
        sb.AppendLine("Q");
    }

    private void RenderEllipse(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var kappa = 0.5522847498;
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

        if (fill)
            sb.AppendLine("f");
    }

    private void RenderRoundRect(StringBuilder sb, double x, double y, double w, double h, double r, bool fill)
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

        if (fill)
            sb.AppendLine("f");
    }

    private void RenderTriangle(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        sb.AppendLine($"{x + w * 0.5:F2} {y:F2} m");
        sb.AppendLine($"{x + w:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h:F2} l");
        sb.AppendLine("h");

        if (fill)
            sb.AppendLine("f");
    }

    private void RenderDiamond(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        sb.AppendLine($"{x + w * 0.5:F2} {y:F2} m");
        sb.AppendLine($"{x + w:F2} {y + h * 0.5:F2} l");
        sb.AppendLine($"{x + w * 0.5:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h * 0.5:F2} l");
        sb.AppendLine("h");

        if (fill)
            sb.AppendLine("f");
    }

    private void RenderPolygon(StringBuilder sb, double x, double y, double w, double h, int sides, bool fill)
    {
        var cx = x + w * 0.5;
        var cy = y + h * 0.5;
        var rx = w * 0.5;
        var ry = h * 0.5;

        for (int i = 0; i < sides; i++)
        {
            var angle = (i * 2 * Math.PI / sides) - Math.PI / 2;
            var px = cx + rx * Math.Cos(angle);
            var py = cy + ry * Math.Sin(angle);

            if (i == 0)
                sb.AppendLine($"{px:F2} {py:F2} m");
            else
                sb.AppendLine($"{px:F2} {py:F2} l");
        }

        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderStar(StringBuilder sb, double x, double y, double w, double h, int points, bool fill)
    {
        var cx = x + w * 0.5;
        var cy = y + h * 0.5;
        var outerR = Math.Min(w, h) * 0.5;
        var innerR = outerR * 0.4;

        for (int i = 0; i < points * 2; i++)
        {
            var angle = (i * Math.PI / points) - Math.PI / 2;
            var r = (i % 2 == 0) ? outerR : innerR;
            var px = cx + r * Math.Cos(angle);
            var py = cy + r * Math.Sin(angle);

            if (i == 0)
                sb.AppendLine($"{px:F2} {py:F2} m");
            else
                sb.AppendLine($"{px:F2} {py:F2} l");
        }

        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderParallelogram(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var skew = w * 0.2;
        sb.AppendLine($"{x + skew:F2} {y:F2} m");
        sb.AppendLine($"{x + w:F2} {y:F2} l");
        sb.AppendLine($"{x + w - skew:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h:F2} l");
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderTrapezoid(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var topWidth = w * 0.6;
        var offset = (w - topWidth) * 0.5;
        sb.AppendLine($"{x + offset:F2} {y:F2} m");
        sb.AppendLine($"{x + w - offset:F2} {y:F2} l");
        sb.AppendLine($"{x + w:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h:F2} l");
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderRightArrow(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var arrowWidth = w * 0.25;
        var arrowHeight = h * 0.5;
        sb.AppendLine($"{x:F2} {y + h * 0.3:F2} m");
        sb.AppendLine($"{x + w - arrowWidth:F2} {y + h * 0.3:F2} l");
        sb.AppendLine($"{x + w - arrowWidth:F2} {y:F2} l");
        sb.AppendLine($"{x + w:F2} {y + h * 0.5:F2} l");
        sb.AppendLine($"{x + w - arrowWidth:F2} {y + h:F2} l");
        sb.AppendLine($"{x + w - arrowWidth:F2} {y + h * 0.7:F2} l");
        sb.AppendLine($"{x:F2} {y + h * 0.7:F2} l");
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderLeftArrow(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var arrowWidth = w * 0.25;
        var arrowHeight = h * 0.5;
        sb.AppendLine($"{x + arrowWidth:F2} {y + h * 0.3:F2} m");
        sb.AppendLine($"{x + w:F2} {y + h * 0.3:F2} l");
        sb.AppendLine($"{x + w:F2} {y + h * 0.7:F2} l");
        sb.AppendLine($"{x + arrowWidth:F2} {y + h * 0.7:F2} l");
        sb.AppendLine($"{x + arrowWidth:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h * 0.5:F2} l");
        sb.AppendLine($"{x + arrowWidth:F2} {y:F2} l");
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderUpArrow(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var arrowHeight = h * 0.25;
        var arrowWidth = w * 0.5;
        sb.AppendLine($"{x + w * 0.3:F2} {y + arrowHeight:F2} m");
        sb.AppendLine($"{x + w * 0.3:F2} {y + h:F2} l");
        sb.AppendLine($"{x + w * 0.7:F2} {y + h:F2} l");
        sb.AppendLine($"{x + w * 0.7:F2} {y + arrowHeight:F2} l");
        sb.AppendLine($"{x + w:F2} {y + arrowHeight:F2} l");
        sb.AppendLine($"{x + w * 0.5:F2} {y:F2} l");
        sb.AppendLine($"{x:F2} {y + arrowHeight:F2} l");
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderDownArrow(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var arrowHeight = h * 0.25;
        var arrowWidth = w * 0.5;
        sb.AppendLine($"{x + w * 0.3:F2} {y:F2} l");
        sb.AppendLine($"{x + w * 0.7:F2} {y:F2} l");
        sb.AppendLine($"{x + w * 0.7:F2} {y + h - arrowHeight:F2} l");
        sb.AppendLine($"{x + w:F2} {y + h - arrowHeight:F2} l");
        sb.AppendLine($"{x + w * 0.5:F2} {y + h:F2} l");
        sb.AppendLine($"{x:F2} {y + h - arrowHeight:F2} l");
        sb.AppendLine($"{x + w * 0.3:F2} {y + h - arrowHeight:F2} l");
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderHeart(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var cx = x + w * 0.5;
        var cy = y + h * 0.5;
        var r = Math.Min(w, h) * 0.4;
        
        // Draw left half of heart
        sb.AppendLine($"{cx:F2} {y + h * 0.2:F2} m");
        sb.AppendLine($"{cx - r:F2} {y + h * 0.6:F2} l");
        sb.AppendLine($"{cx - r * 0.7:F2} {y + h * 0.8:F2} l");
        sb.AppendLine($"{cx:F2} {y + h * 0.6:F2} l");
        
        // Draw right half of heart
        sb.AppendLine($"{cx + r * 0.7:F2} {y + h * 0.8:F2} l");
        sb.AppendLine($"{cx + r:F2} {y + h * 0.6:F2} l");
        sb.AppendLine($"{cx:F2} {y + h * 0.2:F2} l");
        
        sb.AppendLine("h");
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderCloud(StringBuilder sb, double x, double y, double w, double h, bool fill)
    {
        var cx = x + w * 0.5;
        var cy = y + h * 0.5;
        var r = Math.Min(w, h) * 0.25;
        
        // Draw multiple overlapping circles to form a cloud
        sb.AppendLine($"{cx:F2} {y + r:F2} m");
        sb.AppendLine($"{cx + r * 1.5:F2} {y + r:F2} l");
        sb.AppendLine($"{cx + r * 1.8:F2} {y + r * 1.5:F2} l");
        sb.AppendLine($"{cx + r * 1.5:F2} {y + r * 2:F2} l");
        sb.AppendLine($"{cx + r:F2} {y + r * 2.2:F2} l");
        sb.AppendLine($"{cx:F2} {y + r * 2:F2} l");
        sb.AppendLine($"{cx - r:F2} {y + r * 2.2:F2} l");
        sb.AppendLine($"{cx - r * 1.5:F2} {y + r * 2:F2} l");
        sb.AppendLine($"{cx - r * 1.8:F2} {y + r * 1.5:F2} l");
        sb.AppendLine($"{cx - r * 1.5:F2} {y + r:F2} l");
        sb.AppendLine($"{cx:F2} {y + r:F2} l");
        
        if (fill)
            sb.AppendLine("f");
    }

    private void RenderText(StringBuilder sb, Shape shape, double pageHeight, PdfPage page)
    {
        if (shape.Paragraphs == null || !shape.Paragraphs.Any()) return;

        var textProperties = shape.TextProperties;
        var leftInset = (textProperties?.LeftInset ?? 0.1) * 72;
        var topInset = (textProperties?.TopInset ?? 0.05) * 72;
        var rightInset = (textProperties?.RightInset ?? 0.1) * 72;
        var bottomInset = (textProperties?.BottomInset ?? 0.05) * 72;
        var x = shape.Bounds.XPoints + leftInset;
        var textTop = pageHeight - shape.Bounds.YPoints - topInset;
        var width = Math.Max(1, shape.Bounds.WidthPoints - leftInset - rightInset);
        var height = Math.Max(0, shape.Bounds.HeightPoints - topInset - bottomInset);
        var defaultFontSize = textProperties?.FontSize ?? 18;
        var contentHeight = EstimateTextContentHeight(shape.Paragraphs, width, defaultFontSize, textProperties);
        double currentY = ResolveTextStartY(textTop, height, contentHeight, textProperties?.Anchor ?? TextAnchor.Top);
        var autoNumberState = new Dictionary<int, int>();

        foreach (var paragraph in shape.Paragraphs)
        {
            var paragraphLineHeight = GetParagraphLineHeight(paragraph, defaultFontSize, textProperties);
            currentY -= ResolveParagraphSpacing(paragraph.SpaceBefore, paragraphLineHeight);

            double firstLineX = GetParagraphFirstLineX(x, paragraph);
            double continuationX = GetParagraphContinuationX(x, paragraph);
            double firstLineWidth = GetParagraphFirstLineWidth(paragraph, width);
            double continuationWidth = GetParagraphContinuationWidth(paragraph, width);
            double bulletX = GetParagraphBulletX(x, paragraph);
            var paragraphAlignment = paragraph.Alignment ?? textProperties?.Alignment ?? TextAlignment.Left;
            var bulletMarker = ResolveBulletMarker(paragraph, autoNumberState);
            var paragraphLineIndex = 0;

            foreach (var run in paragraph.Runs)
            {
                if (string.IsNullOrEmpty(run.Text)) continue;

                var runFontName = run.Properties?.FontFamily ?? textProperties?.FontFamily ?? "Arial";
                var runFontSize = ResolveEffectiveFontSize(run.Properties?.FontSize ?? defaultFontSize, textProperties);
                var runFontColor = run.Properties?.Color ?? textProperties?.Color ?? Color.Black;
                var lineHeight = ResolveLineHeight(runFontSize, textProperties);

                // Handle baseline offset (superscript/subscript)
                double baselineOffset = 0;
                var baselineOffsetValue = run.Properties?.BaselineOffset;
                if (baselineOffsetValue is double offsetValue && offsetValue != 0)
                {
                    baselineOffset = offsetValue * runFontSize;
                }

                // Handle text wrapping
                var wrappedText = WrapText(
                    run.Text,
                    paragraphLineIndex == 0 ? firstLineWidth : continuationWidth,
                    continuationWidth,
                    runFontSize);
                foreach (var line in wrappedText)
                {
                    currentY -= lineHeight;

                    if (paragraphLineIndex == 0 && !string.IsNullOrEmpty(bulletMarker))
                    {
                        var bulletFontSize = Math.Max(1, (int)Math.Round(runFontSize * ResolveBulletScale(paragraph)));
                        var bulletColor = paragraph.BulletColor ?? runFontColor;
                        var bulletFontName = paragraph.BulletFont ?? runFontName;
                        RenderTextFragment(sb, page, bulletMarker, bulletFontName, bulletFontSize, bulletColor, bulletX, currentY + baselineOffset);
                    }

                    var lineStartX = paragraphLineIndex == 0 ? firstLineX : continuationX;
                    var lineWidth = paragraphLineIndex == 0 ? firstLineWidth : continuationWidth;
                    var lineX = ResolveAlignedTextX(lineStartX, lineWidth, line, runFontSize, paragraphAlignment);
                    RenderTextFragment(sb, page, line, runFontName, runFontSize, runFontColor, lineX, currentY + baselineOffset);
                    paragraphLineIndex++;
                }
            }

            currentY -= ResolveParagraphSpacing(paragraph.SpaceAfter, paragraphLineHeight);
            currentY -= ResolveLineSpacingAdjustment(paragraph.LineSpacing, paragraphLineHeight);
        }
    }

    private static int ResolveEffectiveFontSize(int fontSize, TextProperties? textProperties)
    {
        if (textProperties?.AutoFit != TextAutoFit.Normal)
            return Math.Max(1, fontSize);

        var fontScale = textProperties.FontScale > 0 ? textProperties.FontScale : 1;
        return Math.Max(1, (int)Math.Round(fontSize * fontScale, MidpointRounding.AwayFromZero));
    }

    private static double ResolveLineHeight(double fontSize, TextProperties? textProperties)
    {
        var lineHeight = fontSize * 1.2;
        if (textProperties?.AutoFit != TextAutoFit.Normal)
            return lineHeight;

        var reduction = Math.Clamp(textProperties.LineSpaceReduction, 0, 1);
        return Math.Max(1, lineHeight * (1 - reduction));
    }

    private List<string> WrapText(string text, double width, double fontSize, bool allowWrap = true)
    {
        return WrapText(text, width, width, fontSize, allowWrap);
    }

    private List<string> WrapText(string text, double firstLineWidth, double continuationWidth, double fontSize, bool allowWrap = true)
    {
        if (string.IsNullOrEmpty(text))
            return new List<string> { string.Empty };

        if (firstLineWidth <= 0 && continuationWidth <= 0)
            return new List<string> { text };

        var normalizedText = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        if (!allowWrap)
            return normalizedText.Split('\n').ToList();

        var lines = new List<string>();
        var firstVisualLine = true;

        foreach (var rawLine in normalizedText.Split('\n'))
        {
            WrapRawLine(rawLine, firstLineWidth, continuationWidth, fontSize, lines, ref firstVisualLine);
        }

        return lines.Count > 0 ? lines : new List<string> { string.Empty };
    }

    private void WrapRawLine(
        string rawLine,
        double firstLineWidth,
        double continuationWidth,
        double fontSize,
        List<string> lines,
        ref bool firstVisualLine)
    {
        if (rawLine.Length == 0)
        {
            lines.Add(string.Empty);
            firstVisualLine = false;
            return;
        }

        var currentLine = new StringBuilder();
        var currentWidth = firstVisualLine ? firstLineWidth : continuationWidth;
        foreach (var token in TokenizeForWrapping(rawLine))
        {
            AppendWrappedToken(currentLine, token, fontSize, lines, continuationWidth, ref currentWidth, ref firstVisualLine);
        }

        if (currentLine.Length > 0)
        {
            lines.Add(currentLine.ToString().TrimEnd());
            firstVisualLine = false;
        }
    }

    private IEnumerable<string> TokenizeForWrapping(string text)
    {
        var token = new StringBuilder();
        bool? currentWhitespace = null;

        foreach (var character in text)
        {
            var isWhitespace = character == ' ' || character == '\t';
            if (currentWhitespace.HasValue && currentWhitespace.Value != isWhitespace)
            {
                yield return token.ToString();
                token.Clear();
            }

            token.Append(character);
            currentWhitespace = isWhitespace;
        }

        if (token.Length > 0)
        {
            yield return token.ToString();
        }
    }

    private void AppendWrappedToken(
        StringBuilder currentLine,
        string token,
        double fontSize,
        List<string> lines,
        double continuationWidth,
        ref double currentWidth,
        ref bool firstVisualLine)
    {
        if (token.Length == 0)
            return;

        if (currentLine.Length == 0 && string.IsNullOrWhiteSpace(token))
            return;

        var candidate = currentLine.Length == 0 ? token.TrimStart() : currentLine.ToString() + token;
        if (candidate.Length > 0 && CanFitWrappedText(candidate, fontSize, currentWidth))
        {
            if (currentLine.Length == 0)
            {
                currentLine.Append(candidate);
            }
            else
            {
                currentLine.Append(token);
            }

            return;
        }

        if (currentLine.Length > 0)
        {
            lines.Add(currentLine.ToString().TrimEnd());
            currentLine.Clear();
            currentWidth = continuationWidth;
            firstVisualLine = false;
        }

        AppendTokenByCharacter(currentLine, token.TrimStart(), fontSize, lines, continuationWidth, ref currentWidth, ref firstVisualLine);
    }

    private void AppendTokenByCharacter(
        StringBuilder currentLine,
        string token,
        double fontSize,
        List<string> lines,
        double continuationWidth,
        ref double currentWidth,
        ref bool firstVisualLine)
    {
        foreach (var character in token)
        {
            if (currentLine.Length == 0 && char.IsWhiteSpace(character))
                continue;

            var candidate = currentLine.ToString() + character;
            if (currentLine.Length == 0 || CanFitWrappedText(candidate, fontSize, currentWidth))
            {
                currentLine.Append(character);
                continue;
            }

            lines.Add(currentLine.ToString().TrimEnd());
            currentLine.Clear();
            currentWidth = continuationWidth;
            firstVisualLine = false;
            if (!char.IsWhiteSpace(character))
            {
                currentLine.Append(character);
            }
        }
    }

    private double EstimateTextContentHeight(
        IEnumerable<Paragraph> paragraphs,
        double width,
        int defaultFontSize,
        TextProperties? textProperties = null,
        bool allowWrap = true)
    {
        double contentHeight = 0;

        foreach (var paragraph in paragraphs)
        {
            var paragraphLineHeight = GetParagraphLineHeight(paragraph, defaultFontSize, textProperties);
            contentHeight += ResolveParagraphSpacing(paragraph.SpaceBefore, paragraphLineHeight);

            var firstLineWidth = GetParagraphFirstLineWidth(paragraph, width);
            var continuationWidth = GetParagraphContinuationWidth(paragraph, width);
            var paragraphLineIndex = 0;
            foreach (var run in paragraph.Runs.Where(run => !string.IsNullOrEmpty(run.Text)))
            {
                var runFontSize = ResolveEffectiveFontSize(run.Properties?.FontSize ?? defaultFontSize, textProperties);
                var lineHeight = ResolveLineHeight(runFontSize, textProperties);
                var wrappedLines = WrapText(
                    run.Text!,
                    paragraphLineIndex == 0 ? firstLineWidth : continuationWidth,
                    continuationWidth,
                    runFontSize,
                    allowWrap);
                contentHeight += wrappedLines.Count * lineHeight;
                paragraphLineIndex += wrappedLines.Count;
            }

            contentHeight += ResolveParagraphSpacing(paragraph.SpaceAfter, paragraphLineHeight);
            contentHeight += ResolveLineSpacingAdjustment(paragraph.LineSpacing, paragraphLineHeight);
        }

        return contentHeight;
    }

    private static double ResolveTextStartY(double textTop, double availableHeight, double contentHeight, TextAnchor anchor)
    {
        var offset = anchor switch
        {
            TextAnchor.Bottom or TextAnchor.BottomCentered => Math.Max(0, availableHeight - contentHeight),
            TextAnchor.Middle or TextAnchor.MiddleCentered => Math.Max(0, (availableHeight - contentHeight) / 2),
            _ => 0
        };

        return textTop - offset;
    }

    private static double ResolveParagraphSpacing(Spacing? spacing, double referenceLineHeight)
    {
        if (spacing == null)
            return 0;

        if (spacing.Points.HasValue)
            return spacing.Points.Value;

        if (spacing.Percent.HasValue)
            return referenceLineHeight * (spacing.Percent.Value / 100.0);

        return 0;
    }

    private static double ResolveLineSpacingAdjustment(Spacing? spacing, double referenceLineHeight)
    {
        if (spacing == null)
            return 0;

        if (spacing.Points.HasValue)
            return Math.Max(0, spacing.Points.Value - referenceLineHeight);

        if (spacing.Percent.HasValue)
            return Math.Max(0, referenceLineHeight * (spacing.Percent.Value / 100.0 - 1));

        return 0;
    }

    private static double GetParagraphFirstLineX(double baseX, Paragraph paragraph)
    {
        var indent = paragraph.Indent / 914400.0 * 72;
        return GetParagraphContinuationX(baseX, paragraph) + Math.Max(0, indent);
    }

    private static double GetParagraphContinuationX(double baseX, Paragraph paragraph)
    {
        var leftMargin = paragraph.MarginLeft / 914400.0 * 72;
        return baseX + leftMargin;
    }

    private static double GetParagraphFirstLineWidth(Paragraph paragraph, double totalWidth)
    {
        var indent = paragraph.Indent / 914400.0 * 72;
        return Math.Max(1, GetParagraphContinuationWidth(paragraph, totalWidth) - Math.Max(0, indent));
    }

    private static double GetParagraphContinuationWidth(Paragraph paragraph, double totalWidth)
    {
        var leftMargin = paragraph.MarginLeft / 914400.0 * 72;
        var rightMargin = paragraph.MarginRight / 914400.0 * 72;
        return Math.Max(1, totalWidth - leftMargin - rightMargin);
    }

    private static double GetParagraphStartX(double baseX, Paragraph paragraph)
    {
        return GetParagraphFirstLineX(baseX, paragraph);
    }

    private static double GetParagraphAvailableWidth(Paragraph paragraph, double totalWidth)
    {
        return GetParagraphContinuationWidth(paragraph, totalWidth);
    }

    private static double GetParagraphBulletX(double baseX, Paragraph paragraph)
    {
        var marginLeft = paragraph.MarginLeft / 914400.0 * 72;
        var indent = paragraph.Indent / 914400.0 * 72;
        return baseX + marginLeft + Math.Min(0, indent);
    }

    private static double GetParagraphLineHeight(Paragraph paragraph, int defaultFontSize, TextProperties? textProperties = null)
    {
        var fontSize = paragraph.Runs
            .Where(run => !string.IsNullOrEmpty(run.Text))
            .Select(run => run.Properties?.FontSize ?? defaultFontSize)
            .DefaultIfEmpty(defaultFontSize)
            .Max();

        return ResolveLineHeight(ResolveEffectiveFontSize(fontSize, textProperties), textProperties);
    }

    private double ResolveAlignedTextX(double startX, double width, string text, double fontSize, TextAlignment alignment)
    {
        var textWidth = EstimateTextWidth(text, fontSize);
        return alignment switch
        {
            TextAlignment.Center => startX + Math.Max(0, (width - textWidth) / 2),
            TextAlignment.Right => startX + Math.Max(0, width - textWidth),
            _ => startX
        };
    }

    private void RenderTextFragment(StringBuilder sb, PdfPage page, string text, string fontName, int fontSize, Color color, double x, double y)
    {
        bool isChinese = ContainsChineseCharacters(text);
        PdfFont font = isChinese
            ? GetOrCreateFont("MicrosoftYaHei", page, true)
            : GetOrCreateFont(fontName, page);

        sb.AppendLine("q");
        SetTextColor(sb, page, color);
        sb.AppendLine("BT");
        sb.AppendLine($"/F{font.Number} {fontSize} Tf");
        sb.AppendLine($"1 0 0 1 {x:F2} {y:F2} Tm");
        sb.AppendLine($"({EscapeText(text, isChinese)}) Tj");
        sb.AppendLine("ET");
        sb.AppendLine("Q");
    }

    private void RenderSimpleText(StringBuilder sb, string text, string fontName, int fontSize, Color color, double x, double y)
    {
        var page = _currentPage ?? throw new InvalidOperationException("Text rendering requires an active page context.");
        RenderTextFragment(sb, page, text, fontName, fontSize, color, x, y);
    }

    private static double ResolveBulletScale(Paragraph paragraph)
    {
        return paragraph.BulletSize > 0 ? paragraph.BulletSize : 1;
    }

    private static string? ResolveBulletMarker(Paragraph paragraph, IDictionary<int, int> autoNumberState)
    {
        if (paragraph.BulletType != BulletType.AutoNumber)
        {
            ResetAutoNumberState(autoNumberState, paragraph.Level);
            return paragraph.BulletType switch
            {
                BulletType.Char when !string.IsNullOrEmpty(paragraph.BulletChar) => paragraph.BulletChar,
                BulletType.Blip => "*",
                _ => null
            };
        }

        ResetDeeperAutoNumberState(autoNumberState, paragraph.Level);
        if (!autoNumberState.TryGetValue(paragraph.Level, out var current))
        {
            current = Math.Max(1, paragraph.BulletStartAt);
        }

        autoNumberState[paragraph.Level] = current + 1;
        return FormatAutoNumber(current, paragraph.BulletAutoNumberType);
    }

    private static void ResetAutoNumberState(IDictionary<int, int> autoNumberState, int level)
    {
        foreach (var key in autoNumberState.Keys.Where(key => key >= level).ToList())
        {
            autoNumberState.Remove(key);
        }
    }

    private static void ResetDeeperAutoNumberState(IDictionary<int, int> autoNumberState, int level)
    {
        foreach (var key in autoNumberState.Keys.Where(key => key > level).ToList())
        {
            autoNumberState.Remove(key);
        }
    }

    private static string FormatAutoNumber(int value, string? type)
    {
        return type switch
        {
            "alphaLcParenBoth" => $"({ToAlphabetic(value, false)})",
            "alphaLcParenR" => $"{ToAlphabetic(value, false)})",
            "alphaLcPeriod" => $"{ToAlphabetic(value, false)}.",
            "alphaUcParenBoth" => $"({ToAlphabetic(value, true)})",
            "alphaUcParenR" => $"{ToAlphabetic(value, true)})",
            "alphaUcPeriod" => $"{ToAlphabetic(value, true)}.",
            "romanLcParenBoth" => $"({ToRoman(value, false)})",
            "romanLcParenR" => $"{ToRoman(value, false)})",
            "romanLcPeriod" => $"{ToRoman(value, false)}.",
            "romanUcParenBoth" => $"({ToRoman(value, true)})",
            "romanUcParenR" => $"{ToRoman(value, true)})",
            "romanUcPeriod" => $"{ToRoman(value, true)}.",
            "arabicParenBoth" => $"({value.ToString(CultureInfo.InvariantCulture)})",
            "arabicParenR" => $"{value.ToString(CultureInfo.InvariantCulture)})",
            "arabicPlain" => value.ToString(CultureInfo.InvariantCulture),
            _ => $"{value.ToString(CultureInfo.InvariantCulture)}."
        };
    }

    private static string ToAlphabetic(int value, bool upperCase)
    {
        if (value <= 0)
            return upperCase ? "A" : "a";

        var builder = new StringBuilder();
        var current = value;
        while (current > 0)
        {
            current--;
            var offset = current % 26;
            builder.Insert(0, (char)((upperCase ? 'A' : 'a') + offset));
            current /= 26;
        }

        return builder.ToString();
    }

    private static string ToRoman(int value, bool upperCase)
    {
        if (value <= 0)
            return upperCase ? "I" : "i";

        var numerals = new (int Value, string Symbol)[]
        {
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        };

        var builder = new StringBuilder();
        var current = value;
        foreach (var (numeralValue, symbol) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(symbol);
                current -= numeralValue;
            }
        }

        return upperCase ? builder.ToString() : builder.ToString().ToLowerInvariant();
    }

    private double EstimateTextWidth(string text, double fontSize)
    {
        double widthUnits = 0;

        foreach (var character in text)
        {
            widthUnits += character switch
            {
                ' ' => 0.28,
                '\t' => 1.12,
                _ when IsWideCharacter(character) => 1.0,
                _ when "ilIjtf".IndexOf(character) >= 0 => 0.32,
                _ when "mwMW".IndexOf(character) >= 0 => char.IsUpper(character) ? 0.78 : 0.72,
                _ when char.IsUpper(character) => 0.58,
                _ when char.IsDigit(character) => 0.52,
                _ when ".,'`!:;|".IndexOf(character) >= 0 => 0.24,
                _ when "()[]{}".IndexOf(character) >= 0 => 0.30,
                _ when char.IsPunctuation(character) => 0.32,
                _ => 0.50
            };
        }

        return widthUnits * fontSize;
    }

    private bool CanFitWrappedText(string text, double fontSize, double availableWidth)
    {
        const double wrappingTolerance = 1.03;
        return EstimateTextWidth(text, fontSize) <= availableWidth * wrappingTolerance;
    }

    private static bool IsWideCharacter(char character)
    {
        return character >= 0x2E80 ||
               (character >= 0x1100 && character <= 0x11FF) ||
               (character >= 0x3040 && character <= 0x30FF) ||
               (character >= 0xAC00 && character <= 0xD7AF);
    }

    private bool TryPreparePdfImage(byte[] imageData, out PreparedImageData? preparedImage)
    {
        preparedImage = null;

        try
        {
            preparedImage = ImageConverter.PrepareForPdf(imageData);
            return true;
        }
        catch (NotSupportedException)
        {
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing image for PDF: {ex.Message}");
            return false;
        }
    }

    private void RenderPicture(StringBuilder sb, Picture picture, Slide slide, double pageHeight, PptxDocument pptx, PdfPage page)
    {
        if (string.IsNullOrEmpty(picture.ImageRelationshipId)) return;

        var imagePath = pptx.GetImagePathFromRId(picture.SourcePath ?? slide.SourcePath, picture.ImageRelationshipId);
        if (imagePath == null) return;

        var imageData = pptx.GetImageData(imagePath);
        if (imageData == null) return;

        try
        {
            if (!TryPreparePdfImage(imageData, out var preparedImage) || preparedImage == null)
                return;

            var pdfImage = _document.AddImage(preparedImage);
            page.Images.Add(pdfImage);

            var x = picture.Bounds.XPoints;
            var y = pageHeight - picture.Bounds.YPoints - picture.Bounds.HeightPoints;
            var w = picture.Bounds.WidthPoints;
            var h = picture.Bounds.HeightPoints;

            // Render pre-picture effects
            if (picture.Effects != null)
            {
                var effectsRenderer = new ImageEffectsRenderer(_document, page);
                var pre = effectsRenderer.RenderPreEffects(x, y, w, h, picture.Effects, ShapeType.Rectangle);
                if (!string.IsNullOrEmpty(pre)) sb.Append(pre);
            }

            sb.AppendLine("q");
            sb.AppendLine($"{w:F2} 0 0 {h:F2} {x:F2} {y:F2} cm");
            sb.AppendLine($"/Im{pdfImage.Number} Do");
            sb.AppendLine("Q");

            // Render post-picture effects
            if (picture.Effects != null)
            {
                var effectsRenderer = new ImageEffectsRenderer(_document, page);
                var post = effectsRenderer.RenderPostEffects(x, y, w, h, picture.Effects, ShapeType.Rectangle);
                if (!string.IsNullOrEmpty(post)) sb.Append(post);
            }
        }
        catch (Exception ex)
        {
            // Log error or handle unsupported image format
            Console.WriteLine($"Error rendering image: {ex.Message}");
        }
    }

    private void RenderGroupShape(StringBuilder sb, GroupShape group, Slide slide, double pageHeight, PptxDocument pptx, PdfPage page)
    {
        // Render grouped shapes
        foreach (var shape in group.Shapes)
        {
            RenderShape(sb, shape, pageHeight, slide, pptx, page);
        }

        foreach (var picture in group.Pictures)
        {
            RenderPicture(sb, picture, slide, pageHeight, pptx, page);
        }

        // Recursively render nested groups
        foreach (var childGroup in group.ChildGroups)
        {
            RenderGroupShape(sb, childGroup, slide, pageHeight, pptx, page);
        }

        foreach (var table in group.Tables)
        {
            RenderTable(sb, table, pageHeight, page);
        }
    }

    private void RenderTable(StringBuilder sb, Table table, double pageHeight, PdfPage page)
    {
        var tableX = table.Bounds.XPoints;
        var tableY = pageHeight - table.Bounds.YPoints - table.Bounds.HeightPoints;

        // Calculate column widths and row heights
        var colWidths = table.Columns.Select(c => c.Width / 914400.0 * 72).ToList();
        var rowHeights = table.Rows.Select(r => r.Height / 914400.0 * 72).ToList();

        var currentY = tableY;

        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            var row = table.Rows[rowIdx];
            var rowHeight = rowHeights[rowIdx];
            var currentX = tableX;

            for (int colIdx = 0; colIdx < row.Cells.Count; colIdx++)
            {
                var cell = row.Cells[colIdx];
                var colWidth = colWidths[colIdx];

                // Skip merged cells
                if (cell.HorizontalMerge || cell.VerticalMerge)
                {
                    currentX += colWidth;
                    continue;
                }

                // Calculate cell dimensions considering spans
                var cellWidth = colWidth;
                for (int i = 1; i < cell.ColumnSpan && colIdx + i < colWidths.Count; i++)
                {
                    cellWidth += colWidths[colIdx + i];
                }

                var cellHeight = rowHeight;
                for (int i = 1; i < cell.RowSpan && rowIdx + i < rowHeights.Count; i++)
                {
                    cellHeight += rowHeights[rowIdx + i];
                }

                // Render cell background
                if (cell.Properties?.Fill != null)
                {
                    sb.AppendLine("q");
                    SetColor(sb, page, cell.Properties.Fill.Color);
                    sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                    sb.AppendLine("Q");
                }
                // Apply table banding styles if no explicit fill
                else if (table.Properties != null)
                {
                    if (table.Properties.BandRows && rowIdx % 2 == 1)
                    {
                        // Light gray for banded rows
                        sb.AppendLine("q");
                        SetColor(sb, page, new Color(230, 230, 230));
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                        sb.AppendLine("Q");
                    }
                    else if (table.Properties.BandColumns && colIdx % 2 == 1)
                    {
                        // Light gray for banded columns
                        sb.AppendLine("q");
                        SetColor(sb, page, new Color(242, 242, 242));
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                        sb.AppendLine("Q");
                    }
                    else if (table.Properties.HasHeaderRow && rowIdx == 0)
                    {
                        // Header row style
                        sb.AppendLine("q");
                        SetColor(sb, page, new Color(204, 204, 255));
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                        sb.AppendLine("Q");
                    }
                    else if (table.Properties.HasHeaderColumn && colIdx == 0)
                    {
                        // Header column style
                        sb.AppendLine("q");
                        SetColor(sb, page, new Color(204, 204, 255));
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                        sb.AppendLine("Q");
                    }
                }

                // Render cell borders
                if (cell.Properties?.Borders != null)
                {
                    RenderCellBorders(sb, cell.Properties.Borders, currentX, currentY, cellWidth, cellHeight, page);
                }

                // Render cell text
                if (cell.Paragraphs.Any())
                {
                    RenderCellText(sb, cell, currentX, currentY, cellWidth, cellHeight, page);
                }

                currentX += colWidth;
            }

            currentY += rowHeight;
        }
    }

    private void RenderCellBorders(StringBuilder sb, CellBorders borders, double x, double y, double w, double h, PdfPage page)
    {
        // Top border
        if (borders.Top != null)
        {
            RenderBorder(sb, borders.Top, x, y + h, x + w, y + h, page);
        }

        // Bottom border
        if (borders.Bottom != null)
        {
            RenderBorder(sb, borders.Bottom, x, y, x + w, y, page);
        }

        // Left border
        if (borders.Left != null)
        {
            RenderBorder(sb, borders.Left, x, y, x, y + h, page);
        }

        // Right border
        if (borders.Right != null)
        {
            RenderBorder(sb, borders.Right, x + w, y, x + w, y + h, page);
        }
    }

    private void RenderBorder(StringBuilder sb, CellBorder border, double x1, double y1, double x2, double y2, PdfPage page)
    {
        sb.AppendLine("q");
        if (border.Color.HasValue)
        {
            SetStrokeColor(sb, page, border.Color.Value);
        }
        sb.AppendLine($"{border.Width / 12700.0 * 72:F2} w");

        if (border.DashType != LineDashType.Solid)
        {
            var dashArray = GetDashArray(border.DashType, border.Width);
            sb.AppendLine($"[{dashArray}] 0 d");
        }

        sb.AppendLine($"{x1:F2} {y1:F2} m");
        sb.AppendLine($"{x2:F2} {y2:F2} l S");

        if (border.DashType != LineDashType.Solid)
        {
            sb.AppendLine("[] 0 d");
        }
        sb.AppendLine("Q");
    }

    private void RenderCellText(StringBuilder sb, TableCell cell, double x, double y, double w, double h, PdfPage page)
    {
        if (!cell.Paragraphs.Any()) return;

        // Calculate margins
        double leftMargin = 5;
        double rightMargin = 5;
        double topMargin = 5;
        double bottomMargin = 5;

        if (cell.Properties != null)
        {
            leftMargin = cell.Properties.LeftMargin / 914400.0 * 72;
            rightMargin = cell.Properties.RightMargin / 914400.0 * 72;
            topMargin = cell.Properties.TopMargin / 914400.0 * 72;
            bottomMargin = cell.Properties.BottomMargin / 914400.0 * 72;
        }

        var availableWidth = w - leftMargin - rightMargin;
        var availableHeight = h - topMargin - bottomMargin;
        var textTop = y + h - topMargin;
        var allowWrap = cell.Properties?.NoWrap != true;
        var defaultFontSize = cell.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Select(run => run.Properties?.FontSize ?? 12)
            .DefaultIfEmpty(12)
            .Max();
        var contentHeight = EstimateTextContentHeight(cell.Paragraphs, availableWidth, defaultFontSize, allowWrap: allowWrap);
        var currentY = ResolveTextStartY(textTop, availableHeight, contentHeight, cell.Properties?.Anchor ?? TextAnchor.Top);
        var autoNumberState = new Dictionary<int, int>();

        foreach (var paragraph in cell.Paragraphs)
        {
            var paragraphLineHeight = GetParagraphLineHeight(paragraph, defaultFontSize);
            currentY -= ResolveParagraphSpacing(paragraph.SpaceBefore, paragraphLineHeight);

            double firstLineX = GetParagraphFirstLineX(x + leftMargin, paragraph);
            double continuationX = GetParagraphContinuationX(x + leftMargin, paragraph);
            double firstLineWidth = GetParagraphFirstLineWidth(paragraph, availableWidth);
            double continuationWidth = GetParagraphContinuationWidth(paragraph, availableWidth);
            double bulletX = GetParagraphBulletX(x + leftMargin, paragraph);
            var paragraphAlignment = paragraph.Alignment ?? TextAlignment.Left;
            var bulletMarker = ResolveBulletMarker(paragraph, autoNumberState);
            var paragraphLineIndex = 0;

            foreach (var run in paragraph.Runs)
            {
                if (string.IsNullOrEmpty(run.Text)) continue;

                var runFontName = run.Properties?.FontFamily ?? "Arial";
                var runFontSize = ResolveEffectiveFontSize(run.Properties?.FontSize ?? defaultFontSize, null);
                var runFontColor = run.Properties?.Color ?? Color.Black;
                var lineHeight = ResolveLineHeight(runFontSize, null);

                double baselineOffset = 0;
                var baselineOffsetValue = run.Properties?.BaselineOffset;
                if (baselineOffsetValue is double offsetValue && offsetValue != 0)
                {
                    baselineOffset = offsetValue * runFontSize;
                }

                var wrappedText = WrapText(
                    run.Text,
                    paragraphLineIndex == 0 ? firstLineWidth : continuationWidth,
                    continuationWidth,
                    runFontSize,
                    allowWrap);
                foreach (var line in wrappedText)
                {
                    currentY -= lineHeight;

                    if (paragraphLineIndex == 0 && !string.IsNullOrEmpty(bulletMarker))
                    {
                        var bulletFontSize = Math.Max(1, (int)Math.Round(runFontSize * ResolveBulletScale(paragraph)));
                        var bulletColor = paragraph.BulletColor ?? runFontColor;
                        var bulletFontName = paragraph.BulletFont ?? runFontName;
                        RenderTextFragment(sb, page, bulletMarker, bulletFontName, bulletFontSize, bulletColor, bulletX, currentY + baselineOffset);
                    }

                    var lineStartX = paragraphLineIndex == 0 ? firstLineX : continuationX;
                    var lineWidth = paragraphLineIndex == 0 ? firstLineWidth : continuationWidth;
                    var lineX = ResolveAlignedTextX(lineStartX, lineWidth, line, runFontSize, paragraphAlignment);
                    RenderTextFragment(sb, page, line, runFontName, runFontSize, runFontColor, lineX, currentY + baselineOffset);
                    paragraphLineIndex++;
                }
            }

            currentY -= ResolveParagraphSpacing(paragraph.SpaceAfter, paragraphLineHeight);
            currentY -= ResolveLineSpacingAdjustment(paragraph.LineSpacing, paragraphLineHeight);
        }
    }

    private static string GetDashArray(LineDashType dashType, int width)
    {
        var w = Math.Max(1, width / 12700.0 * 72);
        return dashType switch
        {
            LineDashType.Dot => $"{w} {w * 2}",
            LineDashType.Dash => $"{w * 4} {w * 2}",
            LineDashType.DashDot => $"{w * 4} {w * 2} {w} {w * 2}",
            LineDashType.DashDotDot => $"{w * 4} {w * 2} {w} {w * 2} {w} {w * 2}",
            _ => ""
        };
    }

    private PdfFont GetOrCreateFont(string fontName, PdfPage? page = null, bool isChinese = false)
    {
        // Use different cache key for Chinese fonts
        string cacheKey = isChinese ? $"{fontName}_Chinese" : fontName;
        
        if (_fonts.TryGetValue(cacheKey, out var existingFont))
        {
            // Add existing font to page if not already added
            if (page != null && !page.Fonts.Contains(existingFont))
            {
                page.Fonts.Add(existingFont);
            }
            return existingFont;
        }

        string pdfFontName;
        int fontObjectNumber;
        
        // Check if font name contains Chinese characters or is a known Chinese font
        if (isChinese || fontName.Contains("???") || fontName.Contains("SimSun") || fontName.Contains("STSong") || fontName.Contains("??????") || fontName.Contains("Microsoft YaHei") || fontName.Contains("???") || fontName.Contains("SimHei"))
        {
            // Use Adobe standard CJK font names
            // These are the standard names that PDF readers should support
            pdfFontName = "AdobeSongStd-Light";
        }
        else
        {
            pdfFontName = fontName switch
            {
                "Arial" => "Helvetica",
                "Times New Roman" => "Times-Roman",
                "Courier New" => "Courier",
                "Helvetica" => "Helvetica",
                "Times" => "Times-Roman",
                "Courier" => "Courier",
                _ => "Helvetica"
            };
        }

        fontObjectNumber = _document.GetNextObjectNumber();
        var font = new PdfFont(fontObjectNumber, pdfFontName);
        _document.AddObject(font);
        _fonts[cacheKey] = font;
        
        // Add font to page if provided
        if (page != null)
        {
            page.Fonts.Add(font);
        }

        return font;
    }

    // Helper method to check if text contains Chinese characters
    private static bool ContainsChineseCharacters(string text)
    {
        foreach (char c in text)
        {
            if (c >= 0x4E00 && c <= 0x9FFF)
            {
                return true;
            }
        }
        return false;
    }

    private void SetColor(StringBuilder sb, PdfPage page, Color color)
    {
        ApplyExtGState(sb, page, color.A / 255.0, 1);
        sb.AppendLine($"{color.R / 255.0:F3} {color.G / 255.0:F3} {color.B / 255.0:F3} rg");
    }

    private void SetStrokeColor(StringBuilder sb, PdfPage page, Color color)
    {
        ApplyExtGState(sb, page, 1, color.A / 255.0);
        sb.AppendLine($"{color.R / 255.0:F3} {color.G / 255.0:F3} {color.B / 255.0:F3} RG");
    }

    private void SetTextColor(StringBuilder sb, PdfPage page, Color color)
    {
        ApplyExtGState(sb, page, color.A / 255.0, 1);
        sb.AppendLine($"{color.R / 255.0:F3} {color.G / 255.0:F3} {color.B / 255.0:F3} rg");
    }

    private void ApplyExtGState(StringBuilder sb, PdfPage page, double fillAlpha, double strokeAlpha)
    {
        fillAlpha = Math.Clamp(fillAlpha, 0, 1);
        strokeAlpha = Math.Clamp(strokeAlpha, 0, 1);

        var gsName = $"GS_{(int)Math.Round(fillAlpha * 1000):D4}_{(int)Math.Round(strokeAlpha * 1000):D4}";
        if (!page.ExtGStates.TryGetValue(gsName, out var gsObject))
        {
            gsObject = new PdfExtGState(_document.GetNextObjectNumber(), fillAlpha, strokeAlpha);
            _document.AddObject(gsObject);
            page.ExtGStates[gsName] = gsObject;
        }

        sb.AppendLine($"/{gsName} gs");
    }

    private static string EscapeText(string text, bool isChinese = false)
    {
        // Check if text contains non-ASCII characters (including Chinese)
        bool containsUnicode = false;
        foreach (char c in text)
        {
            if (c > 127)
            {
                containsUnicode = true;
                break;
            }
        }
        
        // If contains Unicode or is Chinese text, use hex string format with UTF-16BE encoding
        if (containsUnicode || isChinese)
        {
            // For Type0 fonts with Identity-H encoding, use UTF-16BE
            var bytes = System.Text.Encoding.BigEndianUnicode.GetBytes(text);
            var hex = BitConverter.ToString(bytes).Replace("-", "");
            return $"<{hex}>";
        }
        
        // Otherwise use literal string with escaping
        return text
            .Replace("\\", "\\\\")
            .Replace("(", "\\(")
            .Replace(")", "\\)")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
    }

    private void RenderGradientFill(StringBuilder sb, Fill fill, double x, double y, double w, double h, ShapeType shapeType, PdfPage page)
    {
        if (fill.GradientStops == null || fill.GradientStops.Count < 2) return;

        var stops = fill.GradientStops.OrderBy(s => s.Position).ToList();
        var stripeCount = 20; // Number of gradient stripes

        sb.AppendLine("q");

        // Create clipping path for the shape
        RenderShapePathForClipping(sb, shapeType, x, y, w, h);
        sb.AppendLine("W n");

        // Render gradient stripes
        for (int i = 0; i < stripeCount; i++)
        {
            var t = i / (double)stripeCount;
            var color = InterpolateGradientColor(stops, t);

            SetColor(sb, page, color);

            // Calculate stripe position based on gradient type
            double sx, sy, sw, sh;
            switch (fill.GradientType)
            {
                case GradientType.Linear:
                case GradientType.Rectangular:
                default:
                    // Horizontal gradient
                    sx = x + w * t;
                    sy = y;
                    sw = w / stripeCount + 1; // Slight overlap to avoid gaps
                    sh = h;
                    break;

                case GradientType.Radial:
                case GradientType.Path:
                    // Radial gradient approximation - concentric rectangles
                    var maxDim = Math.Max(w, h);
                    var radius = maxDim * (1 - t) * 0.5;
                    sx = x + w * 0.5 - radius;
                    sy = y + h * 0.5 - radius;
                    sw = radius * 2;
                    sh = radius * 2;
                    break;
            }

            sb.AppendLine($"{sx:F2} {sy:F2} {sw:F2} {sh:F2} re f");
        }

        sb.AppendLine("Q");
    }

    private void RenderPictureFill(StringBuilder sb, Fill fill, double x, double y, double w, double h, ShapeType shapeType, string sourcePartPath, PptxDocument pptx, PdfPage page)
    {
        if (string.IsNullOrEmpty(fill.PictureRelationshipId))
            return;

        try
        {
            var imagePath = pptx.GetImagePathFromRId(sourcePartPath, fill.PictureRelationshipId);
            if (imagePath == null)
                return;

            var imageData = pptx.GetImageData(imagePath);
            if (imageData == null || !TryPreparePdfImage(imageData, out var preparedImage) || preparedImage == null)
                return;

            var pdfImage = _document.AddImage(preparedImage);
            page.Images.Add(pdfImage);

            sb.AppendLine("q");

            RenderShapePathForClipping(sb, shapeType, x, y, w, h);
            sb.AppendLine("W n");

            sb.AppendLine($"{w:F2} 0 0 {h:F2} {x:F2} {y:F2} cm");
            sb.AppendLine($"/Im{pdfImage.Number} Do");

            sb.AppendLine("Q");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error rendering picture fill: {ex.Message}");
        }
    }

    private void RenderChart(StringBuilder sb, Chart chart, double pageHeight, PdfPage page)
    {
        if (chart == null) return;

        var x = chart.Bounds.XPoints;
        var y = pageHeight - chart.Bounds.YPoints - chart.Bounds.HeightPoints;
        var w = chart.Bounds.WidthPoints;
        var h = chart.Bounds.HeightPoints;

        // Draw chart background
        sb.AppendLine("0.95 0.95 0.95 rg");
        sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re f");

        // Draw chart border
        sb.AppendLine("0.7 0.7 0.7 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re S");

        // Draw chart title if exists
        if (!string.IsNullOrEmpty(chart.Title))
        {
            RenderSimpleText(sb, chart.Title, "Helvetica", 12, Color.Black, x + w / 2, y + h - 20);
        }

        // Calculate plot area
        var plotX = x + 40;
        var plotY = y + 40;
        var plotW = w - 80;
        var plotH = h - 80;

        // Draw plot area
        sb.AppendLine("0.98 0.98 0.98 rg");
        sb.AppendLine($"{plotX:F2} {plotY:F2} {plotW:F2} {plotH:F2} re f");
        sb.AppendLine("0.8 0.8 0.8 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{plotX:F2} {plotY:F2} {plotW:F2} {plotH:F2} re S");

        // Render based on chart type
        switch (chart.Type)
        {
            case ChartType.Bar:
                RenderBarChart(sb, chart, plotX, plotY, plotW, plotH, page);
                break;
            case ChartType.Column:
                RenderColumnChart(sb, chart, plotX, plotY, plotW, plotH, page);
                break;
            case ChartType.Line:
                RenderLineChart(sb, chart, plotX, plotY, plotW, plotH, page);
                break;
            case ChartType.Pie:
                RenderPieChart(sb, chart, plotX, plotY, plotW, plotH, page);
                break;
            default:
                // Draw placeholder for unsupported chart types
                RenderSimpleText(sb, $"{chart.Type} Chart", "Helvetica", 10, new Color(128, 128, 128), plotX + plotW / 2, plotY + plotH / 2);
                break;
        }

        // Draw legend if exists
        if (chart.Legend != null)
        {
            RenderChartLegend(sb, chart, x, y, w, h, page);
        }
    }

    private void RenderBarChart(StringBuilder sb, Chart chart, double x, double y, double w, double h, PdfPage page)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count == 0) return;

        var maxValue = dataPoints.Max(dp => dp.Value);
        if (maxValue <= 0) return;

        var labelWidth = Math.Min(50.0, w * 0.2);
        var valueWidth = 24.0;
        var barAreaX = x + labelWidth + 6;
        var barAreaWidth = Math.Max(10, w - labelWidth - valueWidth - 10);
        var slotHeight = h / dataPoints.Count;
        var barHeight = Math.Max(8, slotHeight * 0.6);
        var fillColor = GetChartSeriesFillColor(series, 0);
        var strokeColor = GetChartSeriesStrokeColor(series, 0);

        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var barWidth = (dataPoint.Value / maxValue) * barAreaWidth;
            var barY = y + h - ((i + 1) * slotHeight) + (slotHeight - barHeight) / 2;

            // Set bar color
            sb.AppendLine("q");
            SetColor(sb, page, fillColor);
            sb.AppendLine($"{barAreaX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re f");
            sb.AppendLine("Q");

            // Draw bar border
            sb.AppendLine("q");
            SetStrokeColor(sb, page, strokeColor);
            sb.AppendLine("1 w");
            sb.AppendLine($"{barAreaX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re S");
            sb.AppendLine("Q");

            // Draw data label
            if (!string.IsNullOrEmpty(dataPoint.Category))
            {
                RenderSimpleText(sb, dataPoint.Category, "Helvetica", 8, Color.Black, x + 2, barY + barHeight / 2 - 2);
            }

            RenderSimpleText(sb, $"{dataPoint.Value:0.##}", "Helvetica", 8, Color.Black, barAreaX + barWidth + 4, barY + barHeight / 2 - 2);
        }
    }

    private void RenderColumnChart(StringBuilder sb, Chart chart, double x, double y, double w, double h, PdfPage page)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count == 0) return;

        var barWidth = w / (dataPoints.Count * 1.5);
        var maxValue = dataPoints.Max(dp => dp.Value);
        if (maxValue <= 0) return;
        var fillColor = GetChartSeriesFillColor(series, 0);
        var strokeColor = GetChartSeriesStrokeColor(series, 0);

        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var barHeight = (dataPoint.Value / maxValue) * h;
            var barX = x + i * barWidth * 1.5 + barWidth * 0.25;
            var barY = y + h - barHeight;

            // Set bar color
            sb.AppendLine("q");
            SetColor(sb, page, fillColor);
            sb.AppendLine($"{barX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re f");
            sb.AppendLine("Q");

            // Draw bar border
            sb.AppendLine("q");
            SetStrokeColor(sb, page, strokeColor);
            sb.AppendLine("1 w");
            sb.AppendLine($"{barX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re S");
            sb.AppendLine("Q");

            if (!string.IsNullOrEmpty(dataPoint.Category))
            {
                RenderSimpleText(sb, dataPoint.Category, "Helvetica", 8, Color.Black, barX + barWidth / 2, y - 10);
            }
        }
    }

    private void RenderLineChart(StringBuilder sb, Chart chart, double x, double y, double w, double h, PdfPage page)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count < 2) return;

        var maxValue = dataPoints.Max(dp => dp.Value);
        if (maxValue <= 0) return;
        var strokeColor = GetChartSeriesStrokeColor(series, 0);
        var fillColor = GetChartSeriesFillColor(series, 0);

        // Draw line
        sb.AppendLine("q");
        SetStrokeColor(sb, page, strokeColor);
        sb.AppendLine("2 w");

        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var pointX = x + (i / (double)(dataPoints.Count - 1)) * w;
            var pointY = y + h - (dataPoint.Value / maxValue) * h;

            if (i == 0)
                sb.AppendLine($"{pointX:F2} {pointY:F2} m");
            else
                sb.AppendLine($"{pointX:F2} {pointY:F2} l");
        }

        sb.AppendLine("S");
        sb.AppendLine("Q");

        // Draw data points
        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var pointX = x + (i / (double)(dataPoints.Count - 1)) * w;
            var pointY = y + h - (dataPoint.Value / maxValue) * h;
            sb.AppendLine("q");
            SetColor(sb, page, fillColor);
            sb.AppendLine($"{pointX - 3:F2} {pointY - 3:F2} 6 0 360 re f");
            sb.AppendLine("Q");

            if (!string.IsNullOrEmpty(dataPoint.Category))
            {
                RenderSimpleText(sb, dataPoint.Category, "Helvetica", 8, Color.Black, pointX, y - 10);
            }
        }
    }

    private void RenderPieChart(StringBuilder sb, Chart chart, double x, double y, double w, double h, PdfPage page)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count == 0) return;

        var centerX = x + w / 2;
        var centerY = y + h / 2;
        var radius = Math.Min(w, h) / 2;

        var totalValue = dataPoints.Sum(dp => dp.Value);
        var currentAngle = 0.0;

        // Define colors for pie slices
        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var sliceAngle = (dataPoint.Value / totalValue) * 360;

            // Set slice color
            sb.AppendLine("q");
            SetColor(sb, page, GetDefaultChartColor(i));

            // Draw pie slice
            sb.AppendLine($"{centerX:F2} {centerY:F2} {radius:F2} {currentAngle:F2} {currentAngle + sliceAngle:F2} ar cn");
            sb.AppendLine("Q");

            currentAngle += sliceAngle;
        }

        // Draw pie border
        sb.AppendLine("0 0 0 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{centerX:F2} {centerY:F2} {radius:F2} 0 360 ar S");
    }

    private void RenderChartLegend(StringBuilder sb, Chart chart, double x, double y, double w, double h, PdfPage page)
    {
        var legendWidth = 110.0;
        var legendItemHeight = 14.0;
        var legendHeight = chart.Series.Count * legendItemHeight;
        var (legendX, legendTop) = GetLegendLayout(chart.Legend?.Position ?? LegendPosition.Right, x, y, w, h, legendWidth, legendHeight);

        for (int i = 0; i < chart.Series.Count; i++)
        {
            var series = chart.Series[i];
            var legendItemY = legendTop - ((i + 1) * legendItemHeight);
            var fillColor = GetChartSeriesFillColor(series, i);

            // Draw legend color box
            sb.AppendLine("q");
            SetColor(sb, page, fillColor);
            sb.AppendLine($"{legendX:F2} {legendItemY:F2} 10 10 re f");
            sb.AppendLine("Q");

            // Draw legend text
            var seriesName = string.IsNullOrWhiteSpace(series.Name) ? $"Series {i + 1}" : series.Name;
            RenderSimpleText(sb, seriesName, "Helvetica", 8, Color.Black, legendX + 15, legendItemY + 2);
        }
    }

    private static Color GetChartSeriesFillColor(ChartSeries series, int index)
    {
        return series.Fill?.Color ?? GetDefaultChartColor(index);
    }

    private static Color GetChartSeriesStrokeColor(ChartSeries series, int index)
    {
        return series.Outline?.Color ?? Darken(GetChartSeriesFillColor(series, index), 0.75);
    }

    private static Color GetDefaultChartColor(int index)
    {
        return (index % 6) switch
        {
            0 => new Color(51, 102, 204),
            1 => new Color(153, 51, 102),
            2 => new Color(46, 184, 92),
            3 => new Color(214, 143, 48),
            4 => new Color(94, 84, 163),
            _ => new Color(44, 162, 173)
        };
    }

    private static Color Darken(Color color, double factor)
    {
        factor = Math.Clamp(factor, 0, 1);
        return new Color(
            (byte)Math.Clamp((int)Math.Round(color.R * factor), 0, 255),
            (byte)Math.Clamp((int)Math.Round(color.G * factor), 0, 255),
            (byte)Math.Clamp((int)Math.Round(color.B * factor), 0, 255),
            color.A);
    }

    private static (double X, double Top) GetLegendLayout(LegendPosition position, double x, double y, double w, double h, double legendWidth, double legendHeight)
    {
        return position switch
        {
            LegendPosition.Left => (x + 12, y + h - 24),
            LegendPosition.Top => (x + (w - legendWidth) / 2, y + h - 12),
            LegendPosition.Bottom => (x + (w - legendWidth) / 2, y + legendHeight + 12),
            LegendPosition.TopRight => (x + w - legendWidth - 12, y + h - 12),
            _ => (x + w - legendWidth - 12, y + h - 24)
        };
    }

    private void RenderSmartArt(StringBuilder sb, SmartArt smartArt, double pageHeight)
    {
        if (smartArt == null) return;

        var x = smartArt.Bounds.XPoints;
        var y = pageHeight - smartArt.Bounds.YPoints - smartArt.Bounds.HeightPoints;
        var w = smartArt.Bounds.WidthPoints;
        var h = smartArt.Bounds.HeightPoints;
        var smartArtTitle = ResolveSmartArtTitle(smartArt);

        // Draw SmartArt background
        sb.AppendLine("0.95 0.95 0.95 rg");
        sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re f");

        // Draw SmartArt border
        sb.AppendLine("0.7 0.7 0.7 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re S");

        // Draw SmartArt title if exists
        if (!string.IsNullOrEmpty(smartArtTitle))
        {
            RenderSimpleText(sb, smartArtTitle, "Helvetica", 10, Color.Black, x + 10, y + h - 15);
        }

        // Calculate content area
        var contentX = x + 20;
        var contentY = y + 20;
        var contentW = w - 40;
        var contentH = h - 40;

        // Render based on SmartArt type
        switch (smartArt.ResolvedType)
        {
            case SmartArtType.List:
            case SmartArtType.VerticalBulletList:
                RenderVerticalListSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.HorizontalBulletList:
                RenderHorizontalListSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.Process:
            case SmartArtType.BasicProcess:
            case SmartArtType.ContinuousBlockProcess:
                RenderProcessSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.Hierarchy:
            case SmartArtType.OrganizationChart:
                RenderHierarchySmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.Cycle:
            case SmartArtType.BasicCycle:
                RenderCycleSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.Matrix:
                RenderMatrixSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.Pyramid:
                RenderPyramidSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            case SmartArtType.Relationship:
            case SmartArtType.BasicTarget:
                RenderRelationshipSmartArt(sb, smartArt, contentX, contentY, contentW, contentH);
                break;
            default:
                // Draw placeholder for unsupported SmartArt types
                RenderSimpleText(sb, ResolveSmartArtFallbackLabel(smartArt), "Helvetica", 10, new Color(128, 128, 128), contentX + contentW / 2, contentY + contentH / 2);
                break;
        }
    }

    private static string? ResolveSmartArtTitle(SmartArt smartArt)
    {
        if (!string.IsNullOrWhiteSpace(smartArt.DisplayName))
            return smartArt.DisplayName;

        if (smartArt.ResolvedType != SmartArtType.Unknown)
            return SmartArt.GetDisplayName(smartArt.ResolvedType);

        return smartArt.Type;
    }

    private static string ResolveSmartArtFallbackLabel(SmartArt smartArt)
    {
        if (!string.IsNullOrWhiteSpace(smartArt.DisplayName))
            return $"{smartArt.DisplayName} SmartArt";

        if (smartArt.ResolvedType != SmartArtType.Unknown)
            return $"{SmartArt.GetDisplayName(smartArt.ResolvedType)} SmartArt";

        return "SmartArt";
    }

    private void RenderSmartArtTextRuns(StringBuilder sb, IReadOnlyList<SmartArtTextRun> runs, double x, double y)
    {
        var normalizedRuns = ExpandSmartArtTextRuns(runs).ToList();
        if (normalizedRuns.Count == 0)
            return;

        var currentX = x;
        var currentY = y;
        var currentFontSize = 12;

        foreach (var run in normalizedRuns)
        {
            if (run.IsLineBreak)
            {
                currentY -= Math.Max(1, currentFontSize) * 1.2;
                currentX = x;
                continue;
            }

            currentFontSize = Math.Max(1, run.FontSize);
            var (r, g, b) = ResolveSmartArtTextColor(run.Color);
            var color = new Color(
                (byte)Math.Clamp((int)Math.Round(r * 255), 0, 255),
                (byte)Math.Clamp((int)Math.Round(g * 255), 0, 255),
                (byte)Math.Clamp((int)Math.Round(b * 255), 0, 255));
            RenderSimpleText(sb, run.Text, "Helvetica", currentFontSize, color, currentX, currentY);
            currentX += EstimateTextWidth(run.Text, currentFontSize);
        }
    }

    private static IEnumerable<SmartArtTextRun> ExpandSmartArtTextRuns(IEnumerable<SmartArtTextRun> runs)
    {
        foreach (var run in runs)
        {
            if (run.IsLineBreak)
            {
                yield return run;
                continue;
            }

            var normalizedText = run.Text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
            var segments = normalizedText.Split('\n');
            for (var i = 0; i < segments.Length; i++)
            {
                if (segments[i].Length > 0)
                {
                    yield return new SmartArtTextRun(segments[i], run.Bold, run.Italic, run.Underline, run.FontSize, run.Color);
                }

                if (i < segments.Length - 1)
                {
                    yield return SmartArtTextRun.LineBreak(run.FontSize, run.Color);
                }
            }
        }
    }

    private static (double R, double G, double B) ResolveSmartArtTextColor(string? color)
    {
        if (!string.IsNullOrWhiteSpace(color) &&
            color.Length >= 6 &&
            int.TryParse(color.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r) &&
            int.TryParse(color.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g) &&
            int.TryParse(color.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
        {
            return (r / 255.0, g / 255.0, b / 255.0);
        }

        return (0, 0, 0);
    }

    private void RenderVerticalListSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        var nodeHeight = h / smartArt.Nodes.Count;
        var nodeWidth = w * 0.9;
        var nodeX = x + (w - nodeWidth) / 2;

        for (int i = 0; i < smartArt.Nodes.Count; i++)
        {
            var node = smartArt.Nodes[i];
            var nodeY = y + i * nodeHeight;
            var nodeH = nodeHeight - 10;

            // Draw node box with rounded corners
            sb.AppendLine("0.4 0.6 0.8 rg");
            var cornerRadius = 8;
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + nodeH - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY + nodeH:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeH:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeH - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h f");

            // Draw node border
            sb.AppendLine("0.2 0.4 0.6 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + nodeH - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY + nodeH:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeH:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeH - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h S");

            // Draw node text with formatting
            RenderSmartArtTextRuns(sb, node.TextRuns, nodeX + 15, nodeY + nodeH / 2);

            // Draw bullet point
            sb.AppendLine("0.2 0.4 0.6 rg");
            sb.AppendLine($"{nodeX - 10:F2} {nodeY + nodeH/2 - 3:F2} 6 6 re f");
        }
    }

    private void RenderHorizontalListSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        var nodeWidth = w / smartArt.Nodes.Count;
        var nodeHeight = h * 0.8;
        var nodeY = y + (h - nodeHeight) / 2;

        for (int i = 0; i < smartArt.Nodes.Count; i++)
        {
            var node = smartArt.Nodes[i];
            var nodeX = x + i * nodeWidth;
            var nodeW = nodeWidth - 10;

            // Draw node box with rounded corners
            sb.AppendLine("0.6 0.4 0.8 rg");
            var cornerRadius = 8;
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h f");

            // Draw node border
            sb.AppendLine("0.4 0.2 0.6 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h S");

            // Draw node text with formatting
            RenderSmartArtTextRuns(sb, node.TextRuns, nodeX + 10, nodeY + nodeHeight / 2);

            // Draw connector to next node
            if (i < smartArt.Nodes.Count - 1)
            {
                var nextNodeX = x + (i + 1) * nodeWidth;
                sb.AppendLine("0.4 0.2 0.6 RG");
                sb.AppendLine("2 w");
                sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + nodeHeight/2:F2} m");
                sb.AppendLine($"{nextNodeX:F2} {nodeY + nodeHeight/2:F2} l S");

                // Draw arrowhead
                sb.AppendLine($"{nextNodeX - 10:F2} {nodeY + nodeHeight/2 - 5:F2} m");
                sb.AppendLine($"{nextNodeX:F2} {nodeY + nodeHeight/2:F2} l");
                sb.AppendLine($"{nextNodeX - 10:F2} {nodeY + nodeHeight/2 + 5:F2} l S");
            }
        }
    }

    private void RenderProcessSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        var nodeWidth = w / smartArt.Nodes.Count;
        var nodeHeight = h * 0.8;
        var nodeY = y + (h - nodeHeight) / 2;

        for (int i = 0; i < smartArt.Nodes.Count; i++)
        {
            var node = smartArt.Nodes[i];
            var nodeX = x + i * nodeWidth;
            var nodeW = nodeWidth - 10;

            // Draw node box with rounded corners
            sb.AppendLine("0.8 0.4 0.2 rg");
            var cornerRadius = 8;
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h f");

            // Draw node border
            sb.AppendLine("0.6 0.2 0 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeW - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h S");

            // Draw node text with formatting
            RenderSmartArtTextRuns(sb, node.TextRuns, nodeX + 10, nodeY + nodeHeight / 2);

            // Draw connector to next node
            if (i < smartArt.Nodes.Count - 1)
            {
                var nextNodeX = x + (i + 1) * nodeWidth;
                sb.AppendLine("0.6 0.2 0 RG");
                sb.AppendLine("2 w");
                sb.AppendLine($"{nodeX + nodeW:F2} {nodeY + nodeHeight/2:F2} m");
                sb.AppendLine($"{nextNodeX:F2} {nodeY + nodeHeight/2:F2} l S");

                // Draw arrowhead
                sb.AppendLine($"{nextNodeX - 10:F2} {nodeY + nodeHeight/2 - 5:F2} m");
                sb.AppendLine($"{nextNodeX:F2} {nodeY + nodeHeight/2:F2} l");
                sb.AppendLine($"{nextNodeX - 10:F2} {nodeY + nodeHeight/2 + 5:F2} l S");
            }
        }
    }

    private void RenderHierarchySmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        // Simple hierarchy rendering
        var rootNode = smartArt.Nodes[0];
        var rootWidth = w * 0.7;
        var rootHeight = h * 0.2;
        var rootX = x + (w - rootWidth) / 2;
        var rootY = y;

        // Draw root node with rounded corners
        sb.AppendLine("0.2 0.6 0.8 rg");
        var cornerRadius = 8;
        sb.AppendLine($"{rootX + cornerRadius:F2} {rootY:F2} m");
        sb.AppendLine($"{rootX + rootWidth - cornerRadius:F2} {rootY:F2} l");
        sb.AppendLine($"{rootX + rootWidth:F2} {rootY + cornerRadius:F2} l");
        sb.AppendLine($"{rootX + rootWidth:F2} {rootY + rootHeight - cornerRadius:F2} l");
        sb.AppendLine($"{rootX + rootWidth - cornerRadius:F2} {rootY + rootHeight:F2} l");
        sb.AppendLine($"{rootX + cornerRadius:F2} {rootY + rootHeight:F2} l");
        sb.AppendLine($"{rootX:F2} {rootY + rootHeight - cornerRadius:F2} l");
        sb.AppendLine($"{rootX:F2} {rootY + cornerRadius:F2} l h f");

        // Draw root node border
        sb.AppendLine("0.1 0.4 0.6 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{rootX + cornerRadius:F2} {rootY:F2} m");
        sb.AppendLine($"{rootX + rootWidth - cornerRadius:F2} {rootY:F2} l");
        sb.AppendLine($"{rootX + rootWidth:F2} {rootY + cornerRadius:F2} l");
        sb.AppendLine($"{rootX + rootWidth:F2} {rootY + rootHeight - cornerRadius:F2} l");
        sb.AppendLine($"{rootX + rootWidth - cornerRadius:F2} {rootY + rootHeight:F2} l");
        sb.AppendLine($"{rootX + cornerRadius:F2} {rootY + rootHeight:F2} l");
        sb.AppendLine($"{rootX:F2} {rootY + rootHeight - cornerRadius:F2} l");
        sb.AppendLine($"{rootX:F2} {rootY + cornerRadius:F2} l h S");

        // Draw root node text with formatting
        RenderSmartArtTextRuns(sb, rootNode.TextRuns, rootX + rootWidth / 2, rootY + rootHeight / 2);

        // Draw child nodes
        if (smartArt.Nodes.Count > 1)
        {
            var childCount = smartArt.Nodes.Count - 1;
            var childWidth = w / childCount * 0.8;
            var childHeight = h * 0.2;
            var childY = y + rootHeight + 30;

            for (int i = 1; i < smartArt.Nodes.Count; i++)
            {
                var node = smartArt.Nodes[i];
                var childX = x + (i - 1) * (w / childCount) + (w / childCount - childWidth) / 2;

                // Draw child node with rounded corners
                sb.AppendLine("0.4 0.8 0.6 rg");
                sb.AppendLine($"{childX + cornerRadius:F2} {childY:F2} m");
                sb.AppendLine($"{childX + childWidth - cornerRadius:F2} {childY:F2} l");
                sb.AppendLine($"{childX + childWidth:F2} {childY + cornerRadius:F2} l");
                sb.AppendLine($"{childX + childWidth:F2} {childY + childHeight - cornerRadius:F2} l");
                sb.AppendLine($"{childX + childWidth - cornerRadius:F2} {childY + childHeight:F2} l");
                sb.AppendLine($"{childX + cornerRadius:F2} {childY + childHeight:F2} l");
                sb.AppendLine($"{childX:F2} {childY + childHeight - cornerRadius:F2} l");
                sb.AppendLine($"{childX:F2} {childY + cornerRadius:F2} l h f");

                // Draw child node border
                sb.AppendLine("0.2 0.6 0.4 RG");
                sb.AppendLine("1 w");
                sb.AppendLine($"{childX + cornerRadius:F2} {childY:F2} m");
                sb.AppendLine($"{childX + childWidth - cornerRadius:F2} {childY:F2} l");
                sb.AppendLine($"{childX + childWidth:F2} {childY + cornerRadius:F2} l");
                sb.AppendLine($"{childX + childWidth:F2} {childY + childHeight - cornerRadius:F2} l");
                sb.AppendLine($"{childX + childWidth - cornerRadius:F2} {childY + childHeight:F2} l");
                sb.AppendLine($"{childX + cornerRadius:F2} {childY + childHeight:F2} l");
                sb.AppendLine($"{childX:F2} {childY + childHeight - cornerRadius:F2} l");
                sb.AppendLine($"{childX:F2} {childY + cornerRadius:F2} l h S");

                // Draw child node text with formatting
                RenderSmartArtTextRuns(sb, node.TextRuns, childX + childWidth / 2, childY + childHeight / 2);

                // Draw connector from root to child
                sb.AppendLine("0.1 0.4 0.6 RG");
                sb.AppendLine("1.5 w");
                sb.AppendLine($"{rootX + rootWidth/2:F2} {rootY + rootHeight:F2} m");
                sb.AppendLine($"{rootX + rootWidth/2:F2} {childY:F2} l S");
                sb.AppendLine($"{rootX + rootWidth/2:F2} {childY:F2} m");
                sb.AppendLine($"{childX + childWidth/2:F2} {childY:F2} l S");
            }
        }
    }

    private void RenderCycleSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        var centerX = x + w / 2;
        var centerY = y + h / 2;
        var radius = Math.Min(w, h) * 0.4;

        // Draw circle background
        sb.AppendLine("0.9 0.9 0.9 rg");
        sb.AppendLine($"{centerX:F2} {centerY:F2} {radius:F2} 0 360 ar f");
        sb.AppendLine("0.7 0.7 0.7 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{centerX:F2} {centerY:F2} {radius:F2} 0 360 ar S");

        // Draw nodes around the circle
        var angleStep = 360.0 / smartArt.Nodes.Count;
        var nodeWidth = radius * 0.4;
        var nodeHeight = radius * 0.25;

        for (int i = 0; i < smartArt.Nodes.Count; i++)
        {
            var node = smartArt.Nodes[i];
            var angle = i * angleStep;
            var radians = angle * Math.PI / 180;
            var nodeX = centerX + Math.Cos(radians) * (radius - nodeWidth/2) - nodeWidth/2;
            var nodeY = centerY + Math.Sin(radians) * (radius - nodeHeight/2) - nodeHeight/2;

            // Draw node box with rounded corners
            sb.AppendLine("0.3 0.6 0.9 rg");
            var cornerRadius = 6;
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h f");

            // Draw node border
            sb.AppendLine("0.1 0.4 0.7 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
            sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h S");

            // Draw node text with formatting
            RenderSmartArtTextRuns(sb, node.TextRuns, nodeX + nodeWidth / 2, nodeY + nodeHeight / 2);

            // Draw connector to next node
            var nextAngle = (i + 1) * angleStep;
            var nextRadians = nextAngle * Math.PI / 180;
            var nextNodeX = centerX + Math.Cos(nextRadians) * (radius - nodeWidth/2) - nodeWidth/2;
            var nextNodeY = centerY + Math.Sin(nextRadians) * (radius - nodeHeight/2) - nodeHeight/2;

            sb.AppendLine("0.1 0.4 0.7 RG");
            sb.AppendLine("1.5 w");
            sb.AppendLine($"{nodeX + nodeWidth/2:F2} {nodeY + nodeHeight/2:F2} m");
            sb.AppendLine($"{nextNodeX + nodeWidth/2:F2} {nextNodeY + nodeHeight/2:F2} l S");

            // Draw arrowhead
            var arrowX = nextNodeX + nodeWidth/2;
            var arrowY = nextNodeY + nodeHeight/2;
            var arrowAngle = Math.Atan2(arrowY - (nodeY + nodeHeight/2), arrowX - (nodeX + nodeWidth/2));
            sb.AppendLine($"{arrowX - 10*Math.Cos(arrowAngle - Math.PI/6):F2} {arrowY - 10*Math.Sin(arrowAngle - Math.PI/6):F2} m");
            sb.AppendLine($"{arrowX:F2} {arrowY:F2} l");
            sb.AppendLine($"{arrowX - 10*Math.Cos(arrowAngle + Math.PI/6):F2} {arrowY - 10*Math.Sin(arrowAngle + Math.PI/6):F2} l S");
        }
    }

    private void RenderMatrixSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        // Simple 2x2 matrix
        var rows = 2;
        var cols = 2;
        var cellWidth = w / cols;
        var cellHeight = h / rows;

        for (int i = 0; i < Math.Min(smartArt.Nodes.Count, rows * cols); i++)
        {
            var node = smartArt.Nodes[i];
            var row = i / cols;
            var col = i % cols;
            var cellX = x + col * cellWidth;
            var cellY = y + row * cellHeight;

            // Draw cell box with rounded corners
            sb.AppendLine("0.7 0.7 0.9 rg");
            var cornerRadius = 6;
            sb.AppendLine($"{cellX + cornerRadius:F2} {cellY:F2} m");
            sb.AppendLine($"{cellX + cellWidth - cornerRadius:F2} {cellY:F2} l");
            sb.AppendLine($"{cellX + cellWidth:F2} {cellY + cornerRadius:F2} l");
            sb.AppendLine($"{cellX + cellWidth:F2} {cellY + cellHeight - cornerRadius:F2} l");
            sb.AppendLine($"{cellX + cellWidth - cornerRadius:F2} {cellY + cellHeight:F2} l");
            sb.AppendLine($"{cellX + cornerRadius:F2} {cellY + cellHeight:F2} l");
            sb.AppendLine($"{cellX:F2} {cellY + cellHeight - cornerRadius:F2} l");
            sb.AppendLine($"{cellX:F2} {cellY + cornerRadius:F2} l h f");

            // Draw cell border
            sb.AppendLine("0.5 0.5 0.8 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{cellX + cornerRadius:F2} {cellY:F2} m");
            sb.AppendLine($"{cellX + cellWidth - cornerRadius:F2} {cellY:F2} l");
            sb.AppendLine($"{cellX + cellWidth:F2} {cellY + cornerRadius:F2} l");
            sb.AppendLine($"{cellX + cellWidth:F2} {cellY + cellHeight - cornerRadius:F2} l");
            sb.AppendLine($"{cellX + cellWidth - cornerRadius:F2} {cellY + cellHeight:F2} l");
            sb.AppendLine($"{cellX + cornerRadius:F2} {cellY + cellHeight:F2} l");
            sb.AppendLine($"{cellX:F2} {cellY + cellHeight - cornerRadius:F2} l");
            sb.AppendLine($"{cellX:F2} {cellY + cornerRadius:F2} l h S");

            // Draw cell text with formatting
            RenderSmartArtTextRuns(sb, node.TextRuns, cellX + cellWidth / 2, cellY + cellHeight / 2);
        }
    }

    private void RenderPyramidSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        var layerHeight = h / smartArt.Nodes.Count;

        for (int i = 0; i < smartArt.Nodes.Count; i++)
        {
            var node = smartArt.Nodes[i];
            var layerWidth = w * (1 - i / (double)smartArt.Nodes.Count * 0.6);
            var layerX = x + (w - layerWidth) / 2;
            var layerY = y + i * layerHeight;

            // Draw pyramid layer with rounded top corners
            sb.AppendLine("0.8 0.6 0.4 rg");
            var cornerRadius = 6;
            sb.AppendLine($"{layerX + cornerRadius:F2} {layerY:F2} m");
            sb.AppendLine($"{layerX + layerWidth - cornerRadius:F2} {layerY:F2} l");
            sb.AppendLine($"{layerX + layerWidth:F2} {layerY + cornerRadius:F2} l");
            sb.AppendLine($"{layerX + layerWidth:F2} {layerY + layerHeight:F2} l");
            sb.AppendLine($"{layerX:F2} {layerY + layerHeight:F2} l");
            sb.AppendLine($"{layerX:F2} {layerY + cornerRadius:F2} l h f");

            // Draw layer border
            sb.AppendLine("0.6 0.4 0.2 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{layerX + cornerRadius:F2} {layerY:F2} m");
            sb.AppendLine($"{layerX + layerWidth - cornerRadius:F2} {layerY:F2} l");
            sb.AppendLine($"{layerX + layerWidth:F2} {layerY + cornerRadius:F2} l");
            sb.AppendLine($"{layerX + layerWidth:F2} {layerY + layerHeight:F2} l");
            sb.AppendLine($"{layerX:F2} {layerY + layerHeight:F2} l");
            sb.AppendLine($"{layerX:F2} {layerY + cornerRadius:F2} l h S");

            // Draw layer text with formatting
            RenderSmartArtTextRuns(sb, node.TextRuns, layerX + layerWidth / 2, layerY + layerHeight / 2);
        }
    }

    private void RenderRelationshipSmartArt(StringBuilder sb, SmartArt smartArt, double x, double y, double w, double h)
    {
        if (smartArt.Nodes.Count == 0) return;

        // Draw center circle
        var centerX = x + w / 2;
        var centerY = y + h / 2;
        var centerRadius = Math.Min(w, h) * 0.2;

        sb.AppendLine("0.2 0.7 0.5 rg");
        sb.AppendLine($"{centerX:F2} {centerY:F2} {centerRadius:F2} 0 360 ar f");
        sb.AppendLine("0.1 0.5 0.3 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{centerX:F2} {centerY:F2} {centerRadius:F2} 0 360 ar S");

        // Draw center text
        if (smartArt.Nodes.Count > 0)
        {
            RenderSmartArtTextRuns(sb, smartArt.Nodes[0].TextRuns, centerX, centerY);
        }

        // Draw surrounding nodes
        if (smartArt.Nodes.Count > 1)
        {
            var nodeCount = smartArt.Nodes.Count - 1;
            var angleStep = 360.0 / nodeCount;
            var nodeRadius = Math.Min(w, h) * 0.35;
            var nodeWidth = centerRadius * 1.5;
            var nodeHeight = centerRadius;

            for (int i = 1; i < smartArt.Nodes.Count; i++)
            {
                var node = smartArt.Nodes[i];
                var angle = (i - 1) * angleStep;
                var radians = angle * Math.PI / 180;
                var nodeX = centerX + Math.Cos(radians) * nodeRadius - nodeWidth/2;
                var nodeY = centerY + Math.Sin(radians) * nodeRadius - nodeHeight/2;

                // Draw node box with rounded corners
                sb.AppendLine("0.6 0.8 0.6 rg");
                var cornerRadius = 6;
                sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
                sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY:F2} l");
                sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + cornerRadius:F2} l");
                sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
                sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
                sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
                sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
                sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h f");

                // Draw node border
                sb.AppendLine("0.4 0.6 0.4 RG");
                sb.AppendLine("1 w");
                sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY:F2} m");
                sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY:F2} l");
                sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + cornerRadius:F2} l");
                sb.AppendLine($"{nodeX + nodeWidth:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
                sb.AppendLine($"{nodeX + nodeWidth - cornerRadius:F2} {nodeY + nodeHeight:F2} l");
                sb.AppendLine($"{nodeX + cornerRadius:F2} {nodeY + nodeHeight:F2} l");
                sb.AppendLine($"{nodeX:F2} {nodeY + nodeHeight - cornerRadius:F2} l");
                sb.AppendLine($"{nodeX:F2} {nodeY + cornerRadius:F2} l h S");

                // Draw node text with formatting
                RenderSmartArtTextRuns(sb, node.TextRuns, nodeX + nodeWidth / 2, nodeY + nodeHeight / 2);

                // Draw connector to center
                sb.AppendLine("0.4 0.6 0.4 RG");
                sb.AppendLine("1.5 w");
                sb.AppendLine($"{centerX:F2} {centerY:F2} m");
                sb.AppendLine($"{nodeX + nodeWidth/2:F2} {nodeY + nodeHeight/2:F2} l S");
            }
        }
    }

    private void RenderShapePathForClipping(StringBuilder sb, ShapeType shapeType, double x, double y, double w, double h)
    {
        switch (shapeType)
        {
            case ShapeType.Rectangle:
            case ShapeType.AutoShape:
            case ShapeType.TextBox:
                sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re");
                break;

            case ShapeType.Ellipse:
                RenderEllipsePath(sb, x, y, w, h);
                break;

            default:
                // Default to rectangle
                sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re");
                break;
        }
    }

    private void RenderEllipsePath(StringBuilder sb, double x, double y, double w, double h)
    {
        var cx = x + w / 2;
        var cy = y + h / 2;
        var rx = w / 2;
        var ry = h / 2;

        // Approximate ellipse with bezier curves
        const double kappa = 0.5522847498; // 4/3 * (sqrt(2) - 1)

        var ox = rx * kappa; // Control point offset x
        var oy = ry * kappa; // Control point offset y

        sb.AppendLine($"{cx + rx:F2} {cy:F2} m");
        sb.AppendLine($"{cx + rx:F2} {cy + oy:F2} {cx + ox:F2} {cy + ry:F2} {cx:F2} {cy + ry:F2} c");
        sb.AppendLine($"{cx - ox:F2} {cy + ry:F2} {cx - rx:F2} {cy + oy:F2} {cx - rx:F2} {cy:F2} c");
        sb.AppendLine($"{cx - rx:F2} {cy - oy:F2} {cx - ox:F2} {cy - ry:F2} {cx:F2} {cy - ry:F2} c");
        sb.AppendLine($"{cx + ox:F2} {cy - ry:F2} {cx + rx:F2} {cy - oy:F2} {cx + rx:F2} {cy:F2} c");
        sb.AppendLine("h");
    }

    private Color InterpolateGradientColor(List<Nedev.FileConverters.PptxToPdf.GradientStop> stops, double position)
    {
        // Find the two stops that bracket the position
        Nedev.FileConverters.PptxToPdf.GradientStop? lower = null;
        Nedev.FileConverters.PptxToPdf.GradientStop? upper = null;

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
