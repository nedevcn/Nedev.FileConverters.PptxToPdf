using System.Globalization;
using System.Text;
using NPptxToPdf.Image;
using NPptxToPdf.Pptx;

namespace NPptxToPdf.Pdf;

public class PdfRenderer
{
    private readonly PdfDocument _document;
    private readonly Dictionary<string, PdfFont> _fonts = new();
    private readonly FontEmbedder _fontEmbedder;
    private int _imageCounter;

    public PdfRenderer(PdfDocument document)
    {
        _document = document;
        _fontEmbedder = new FontEmbedder(document);
    }

    public void RenderSlide(PdfPage page, Slide slide, PptxDocument pptx)
    {
        var content = new PdfContent(_document.GetNextObjectNumber());
        _document.AddObject(content);
        page.Content = content;

        var sb = new StringBuilder();

        sb.AppendLine("q");

        if (slide.Background?.Color != null)
        {
            RenderBackground(sb, slide.Background.Color.Value, page.Width, page.Height);
        }

        // Render connectors first (behind shapes)
        foreach (var connector in slide.Connectors)
        {
            RenderConnector(sb, connector, page.Height);
        }

        // Render shapes
        foreach (var shape in slide.Shapes)
        {
            RenderShape(sb, shape, page.Height, page);
        }

        // Render pictures
        foreach (var picture in slide.Pictures)
        {
            RenderPicture(sb, picture, page.Height, pptx, page);
        }

        // Render charts
        foreach (var chart in slide.Charts)
        {
            RenderChart(sb, chart, page.Height);
        }

        // Render SmartArt
        foreach (var smartArt in slide.SmartArts)
        {
            RenderSmartArt(sb, smartArt, page.Height);
        }

        // Render group shapes
        foreach (var group in slide.GroupShapes)
        {
            RenderGroupShape(sb, group, page.Height, pptx, page);
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

    private void RenderBackground(StringBuilder sb, Color color, double pageWidth, double pageHeight)
    {
        SetColor(sb, color);
        sb.AppendLine($"0 0 {pageWidth:F2} {pageHeight:F2} re f");
    }

    private void RenderConnector(StringBuilder sb, Connector connector, double pageHeight)
    {
        if (connector.Outline == null || connector.Outline.Width <= 0) return;

        var x1 = connector.Bounds.XPoints;
        var y1 = pageHeight - connector.Bounds.YPoints;
        var x2 = x1 + connector.Bounds.WidthPoints;
        var y2 = y1 - connector.Bounds.HeightPoints;

        var strokeColor = connector.Outline.Color ?? Color.Black;
        SetStrokeColor(sb, strokeColor);
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
    }

    private void RenderShape(StringBuilder sb, Shape shape, double pageHeight, PdfPage page)
    {
        if (shape.ShapeType == ShapeType.Line)
        {
            RenderLine(sb, shape, pageHeight);
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

        // Render fill
        RenderShapeFill(sb, shape, x, y, w, h);

        // Render outline
        RenderShapeOutline(sb, shape, x, y, w, h);

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

    private void RenderShapeFill(StringBuilder sb, Shape shape, double x, double y, double w, double h)
    {
        if (shape.Fill == null) return;

        switch (shape.Fill.Type)
        {
            case FillType.Solid:
                SetColor(sb, shape.Fill.Color);
                RenderShapePath(sb, shape.ShapeType, x, y, w, h, true);
                break;

            case FillType.Gradient:
                // Render gradient fill using striped approximation
                if (shape.Fill.GradientStops?.Any() == true)
                {
                    RenderGradientFill(sb, shape.Fill, x, y, w, h, shape.ShapeType);
                }
                break;

            case FillType.Pattern:
                // Pattern fill - use foreground color
                SetColor(sb, shape.Fill.PatternForegroundColor);
                RenderShapePath(sb, shape.ShapeType, x, y, w, h, true);
                break;

            case FillType.Picture:
                // Picture fill
                if (shape.Fill.PictureFill != null && shape.Fill.PictureFill.Blip != null)
                {
                    RenderPictureFill(sb, shape.Fill, x, y, w, h, shape.ShapeType);
                }
                break;

            case FillType.None:
                // No fill
                break;
        }
    }

    private void RenderShapeOutline(StringBuilder sb, Shape shape, double x, double y, double w, double h)
    {
        if (shape.Outline == null || shape.Outline.Width <= 0) return;

        var strokeColor = shape.Outline.Color ?? Color.Black;
        SetStrokeColor(sb, strokeColor);
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

    private void RenderLine(StringBuilder sb, Shape shape, double pageHeight)
    {
        if (shape.Outline == null || shape.Outline.Width <= 0) return;

        var x1 = shape.Bounds.XPoints;
        var y1 = pageHeight - shape.Bounds.YPoints;
        var x2 = x1 + shape.Bounds.WidthPoints;
        var y2 = y1 - shape.Bounds.HeightPoints;

        var strokeColor = shape.Outline.Color ?? Color.Black;
        SetStrokeColor(sb, strokeColor);
        sb.AppendLine($"{shape.Outline.Width / 12700.0 * 72:F2} w");

        sb.AppendLine($"{x1:F2} {y1:F2} m");
        sb.AppendLine($"{x2:F2} {y2:F2} l S");
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

        var x = shape.Bounds.XPoints + (shape.TextProperties?.LeftInset ?? 0.1) * 72;
        var y = pageHeight - shape.Bounds.YPoints - shape.Bounds.HeightPoints + (shape.TextProperties?.TopInset ?? 0.05) * 72;
        var width = shape.Bounds.WidthPoints - (shape.TextProperties?.LeftInset ?? 0.1) * 72 - (shape.TextProperties?.RightInset ?? 0.1) * 72;
        var height = shape.Bounds.HeightPoints - (shape.TextProperties?.TopInset ?? 0.05) * 72 - (shape.TextProperties?.BottomInset ?? 0.05) * 72;

        double currentY = y;

        foreach (var paragraph in shape.Paragraphs)
        {
            // Apply space before
            if (paragraph.SpaceBefore != null)
            {
                if (paragraph.SpaceBefore.Percent.HasValue)
                    currentY += 12 * (paragraph.SpaceBefore.Percent.Value / 100);
                else if (paragraph.SpaceBefore.Points.HasValue)
                    currentY += paragraph.SpaceBefore.Points.Value;
            }

            // Calculate paragraph indent
            double paragraphX = x + (paragraph.MarginLeft / 914400.0 * 72) + (paragraph.Indent / 914400.0 * 72);

            // Render runs with formatting
            double currentX = paragraphX;
            double lineHeight = 18; // Default line height

            foreach (var run in paragraph.Runs)
            {
                if (string.IsNullOrEmpty(run.Text)) continue;

                var runFontName = run.Properties?.FontFamily ?? shape.TextProperties?.FontFamily ?? "Arial";
                var runFontSize = run.Properties?.FontSize ?? shape.TextProperties?.FontSize ?? 18;
                var runFontColor = run.Properties?.Color ?? shape.TextProperties?.Color ?? Color.Black;

                lineHeight = runFontSize * 1.2;

                // Check if text contains Chinese characters
                bool isChinese = ContainsChineseCharacters(run.Text);
                
                // Get or create font
                PdfFont font;
                if (isChinese)
                {
                    // Use a widely supported Chinese font
                    // Try using system font if available
                    font = GetOrCreateFont("MicrosoftYaHei", page, true);
                }
                else
                {
                    font = GetOrCreateFont(runFontName, page);
                }

                // Apply text color
                SetTextColor(sb, runFontColor);

                // Start text object
                sb.AppendLine("BT");

                // Set font and size
                string fontName = $"/F{font.Number}";
                string fontStyle = "";
                
                // Apply font styles
                if (run.Properties?.Bold == true && run.Properties?.Italic == true)
                {
                    fontStyle = "BoldItalic";
                    // Note: We'd need to handle bold/italic fonts properly
                }
                else if (run.Properties?.Bold == true)
                {
                    fontStyle = "Bold";
                }
                else if (run.Properties?.Italic == true)
                {
                    fontStyle = "Italic";
                }

                sb.AppendLine($"{fontName} {runFontSize} Tf");

                // Handle text alignment
                string alignment = paragraph.Alignment switch
                {
                    TextAlignment.Center => "1 0 0 1 0 0 Tm 0 -1 1 0 0 0 Tm",
                    TextAlignment.Right => $"1 0 0 1 {paragraphX + width} 0 Tm 0 -1 1 0 0 0 Tm",
                    _ => ""
                };

                if (!string.IsNullOrEmpty(alignment))
                {
                    sb.AppendLine(alignment);
                }

                // Handle underline
                if (run.Properties?.Underline != UnderlineType.None)
                {
                    // PDF doesn't support underline directly in text state
                    // We'll need to draw a line under the text
                }

                // Handle strikethrough
                if (run.Properties?.Strike != StrikeType.None)
                {
                    // PDF doesn't support strikethrough directly
                    // We'll need to draw a line through the text
                }

                // Handle baseline offset (superscript/subscript)
                if (run.Properties?.BaselineOffset != 0)
                {
                    double offset = run.Properties.BaselineOffset * runFontSize;
                    sb.AppendLine($"1 0 0 1 0 {offset} Tm");
                }

                // Handle text wrapping
                var wrappedText = WrapText(run.Text, width, runFontSize);
                foreach (var line in wrappedText)
                {
                    sb.AppendLine($"{currentX:F2} {currentY:F2} Td");
                    sb.AppendLine($"({EscapeText(line, isChinese)}) Tj");
                    currentY += lineHeight;
                    currentX = paragraphX;
                }

                // End text object
                sb.AppendLine("ET");
            }

            // Apply space after
            if (paragraph.SpaceAfter != null)
            {
                if (paragraph.SpaceAfter.Percent.HasValue)
                    currentY += 12 * (paragraph.SpaceAfter.Percent.Value / 100);
                else if (paragraph.SpaceAfter.Points.HasValue)
                    currentY += paragraph.SpaceAfter.Points.Value;
            }

            // Apply line spacing
            if (paragraph.LineSpacing != null)
            {
                if (paragraph.LineSpacing.Percent.HasValue)
                    currentY += lineHeight * (paragraph.LineSpacing.Percent.Value / 100 - 1);
                else if (paragraph.LineSpacing.Points.HasValue)
                    currentY += paragraph.LineSpacing.Points.Value - lineHeight;
            }
        }
    }

    private List<string> WrapText(string text, double width, double fontSize)
    {
        var lines = new List<string>();
        var words = text.Split(' ');
        var currentLine = new System.Text.StringBuilder();

        foreach (var word in words)
        {
            // Simple word wrapping - estimate width based on average character width
            double currentWidth = (currentLine.Length + word.Length + 1) * (fontSize * 0.5);
            
            if (currentWidth <= width)
            {
                if (currentLine.Length > 0)
                    currentLine.Append(' ');
                currentLine.Append(word);
            }
            else
            {
                lines.Add(currentLine.ToString());
                currentLine.Clear();
                currentLine.Append(word);
            }
        }

        if (currentLine.Length > 0)
            lines.Add(currentLine.ToString());

        return lines;
    }

    private void RenderPicture(StringBuilder sb, Picture picture, double pageHeight, PptxDocument pptx, PdfPage page)
    {
        if (string.IsNullOrEmpty(picture.ImageRelationshipId)) return;

        var imagePath = pptx.GetImagePathFromRId(picture.ImageRelationshipId);
        if (imagePath == null) return;

        var imageData = pptx.GetImageData(imagePath);
        if (imageData == null) return;

        try
        {
            var imageInfo = ImageDecoder.Decode(imageData);

            // For PDF, we need JPEG format
            // If it's already JPEG, use directly; otherwise we would need conversion
            bool isJpeg = imageInfo.Format == ImageFormat.Jpeg;

            if (!isJpeg)
            {
                // Convert non-JPEG formats to JPEG
                try
                {
                    imageData = ImageConverter.EnsureJpegFormat(imageData);
                    isJpeg = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting image: {ex.Message}");
                    return;
                }
            }

            var pdfImage = _document.AddImage(imageData, imageInfo.Width, imageInfo.Height, true);
            page.Images.Add(pdfImage);

            var x = picture.Bounds.XPoints;
            var y = pageHeight - picture.Bounds.YPoints - picture.Bounds.HeightPoints;
            var w = picture.Bounds.WidthPoints;
            var h = picture.Bounds.HeightPoints;

            sb.AppendLine("q");
            sb.AppendLine($"{w:F2} 0 0 {h:F2} {x:F2} {y:F2} cm");
            sb.AppendLine($"/Im{pdfImage.Number} Do");
            sb.AppendLine("Q");
        }
        catch (Exception ex)
        {
            // Log error or handle unsupported image format
            Console.WriteLine($"Error rendering image: {ex.Message}");
        }
    }

    private void RenderGroupShape(StringBuilder sb, GroupShape group, double pageHeight, PptxDocument pptx, PdfPage page)
    {
        // Render grouped shapes
        foreach (var shape in group.Shapes)
        {
            RenderShape(sb, shape, pageHeight, page);
        }

        foreach (var picture in group.Pictures)
        {
            RenderPicture(sb, picture, pageHeight, pptx, page);
        }

        // Recursively render nested groups
        foreach (var childGroup in group.ChildGroups)
        {
            RenderGroupShape(sb, childGroup, pageHeight, pptx, page);
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
                    SetColor(sb, cell.Properties.Fill.Color);
                    sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                }
                // Apply table banding styles if no explicit fill
                else if (table.Properties != null)
                {
                    if (table.Properties.BandRows && rowIdx % 2 == 1)
                    {
                        // Light gray for banded rows
                        sb.AppendLine("0.9 0.9 0.9 rg");
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                    }
                    else if (table.Properties.BandColumns && colIdx % 2 == 1)
                    {
                        // Light gray for banded columns
                        sb.AppendLine("0.95 0.95 0.95 rg");
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                    }
                    else if (table.Properties.HasHeaderRow && rowIdx == 0)
                    {
                        // Header row style
                        sb.AppendLine("0.8 0.8 1 rg");
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                    }
                    else if (table.Properties.HasHeaderColumn && colIdx == 0)
                    {
                        // Header column style
                        sb.AppendLine("0.8 0.8 1 rg");
                        sb.AppendLine($"{currentX:F2} {currentY:F2} {cellWidth:F2} {cellHeight:F2} re f");
                    }
                }

                // Render cell borders
                if (cell.Properties?.Borders != null)
                {
                    RenderCellBorders(sb, cell.Properties.Borders, currentX, currentY, cellWidth, cellHeight);
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

    private void RenderCellBorders(StringBuilder sb, CellBorders borders, double x, double y, double w, double h)
    {
        // Top border
        if (borders.Top != null)
        {
            RenderBorder(sb, borders.Top, x, y + h, x + w, y + h);
        }

        // Bottom border
        if (borders.Bottom != null)
        {
            RenderBorder(sb, borders.Bottom, x, y, x + w, y);
        }

        // Left border
        if (borders.Left != null)
        {
            RenderBorder(sb, borders.Left, x, y, x, y + h);
        }

        // Right border
        if (borders.Right != null)
        {
            RenderBorder(sb, borders.Right, x + w, y, x + w, y + h);
        }
    }

    private void RenderBorder(StringBuilder sb, CellBorder border, double x1, double y1, double x2, double y2)
    {
        if (border.Color.HasValue)
        {
            SetStrokeColor(sb, border.Color.Value);
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

        var currentY = y + h - topMargin;

        foreach (var paragraph in cell.Paragraphs)
        {
            var text = paragraph.GetFullText();
            if (string.IsNullOrEmpty(text)) continue;

            // Check if text contains Chinese characters
            bool isChinese = ContainsChineseCharacters(text);

            // Get font and style from the first run
            var firstRun = paragraph.Runs.FirstOrDefault();
            var fontFamily = firstRun?.Properties?.FontFamily ?? "Arial";
            var fontSize = firstRun?.Properties?.FontSize ?? 12;
            var color = firstRun?.Properties?.Color ?? Color.Black;

            // Calculate line height
            var lineHeight = fontSize * 1.2;

            // Get or create font
            PdfFont font;
            if (isChinese)
            {
                font = GetOrCreateFont("MicrosoftYaHei", page, true);
            }
            else
            {
                font = GetOrCreateFont(fontFamily, page);
            }

            SetTextColor(sb, color);
            sb.AppendLine("BT");
            sb.AppendLine($"/F{font.Number} {fontSize} Tf");

            // Handle text wrapping
            var wrappedText = WrapText(text, availableWidth, fontSize);

            foreach (var line in wrappedText)
            {
                // Calculate text position based on alignment and cell anchor
                var textX = x + leftMargin;
                var textY = currentY;

                // Apply cell anchor
                if (cell.Properties != null)
                {
                    switch (cell.Properties.Anchor)
                    {
                        case TextAnchor.Bottom:
                            textY = y + bottomMargin + (availableHeight - lineHeight * wrappedText.Count);
                            break;
                        case TextAnchor.Middle:
                            textY = y + h * 0.5 - lineHeight * wrappedText.Count * 0.5;
                            break;
                    }
                }

                // Apply paragraph alignment
                if (paragraph.Alignment == TextAlignment.Center)
                {
                    // Approximate center alignment
                    textX = x + w * 0.5 - line.Length * fontSize * 0.25;
                }
                else if (paragraph.Alignment == TextAlignment.Right)
                {
                    // Approximate right alignment
                    textX = x + w - rightMargin - line.Length * fontSize * 0.5;
                }

                sb.AppendLine($"{textX:F2} {textY:F2} Td");
                sb.AppendLine($"({EscapeText(line, isChinese)}) Tj");
                sb.AppendLine("ET");

                currentY -= lineHeight;
            }
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
        if (isChinese || fontName.Contains("宋体") || fontName.Contains("SimSun") || fontName.Contains("STSong") || fontName.Contains("微软雅黑") || fontName.Contains("Microsoft YaHei") || fontName.Contains("黑体") || fontName.Contains("SimHei"))
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

    private static void SetColor(StringBuilder sb, Color color)
    {
        sb.AppendLine($"{color.R / 255.0:F3} {color.G / 255.0:F3} {color.B / 255.0:F3} rg");
    }

    private static void SetStrokeColor(StringBuilder sb, Color color)
    {
        sb.AppendLine($"{color.R / 255.0:F3} {color.G / 255.0:F3} {color.B / 255.0:F3} RG");
    }

    private static void SetTextColor(StringBuilder sb, Color color)
    {
        sb.AppendLine($"{color.R / 255.0:F3} {color.G / 255.0:F3} {color.B / 255.0:F3} rg");
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

    private void RenderGradientFill(StringBuilder sb, Fill fill, double x, double y, double w, double h, ShapeType shapeType)
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

            SetColor(sb, color);

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

    private void RenderPictureFill(StringBuilder sb, Fill fill, double x, double y, double w, double h, ShapeType shapeType)
    {
        if (fill.PictureFill == null || fill.PictureFill.Blip == null || fill.PictureFill.Blip.Data == null)
            return;

        try
        {
            // Get image data
            var imageData = fill.PictureFill.Blip.Data;
            var base64Image = Convert.ToBase64String(imageData);

            // Create image object in PDF
            var imageId = _imageCounter++;
            sb.AppendLine($"{imageId} 0 obj");
            sb.AppendLine($"<< /Type /XObject /Subtype /Image /Width 100 /Height 100 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode >>");
            sb.AppendLine($"stream");
            sb.AppendLine(base64Image);
            sb.AppendLine("endstream");
            sb.AppendLine(">>");
            sb.AppendLine("endobj");

            // Use the image as fill
            sb.AppendLine("q");

            // Create clipping path for the shape
            RenderShapePathForClipping(sb, shapeType, x, y, w, h);
            sb.AppendLine("W n");

            // Draw the image
            sb.AppendLine($"1 0 0 1 {x:F2} {y:F2} cm");
            sb.AppendLine($"/{imageId} Do");

            sb.AppendLine("Q");
        }
        catch (Exception ex)
        {
            // Fallback to solid fill if image processing fails
            SetColor(sb, fill.Color);
            RenderShapePath(sb, shapeType, x, y, w, h, true);
        }
    }

    private void RenderChart(StringBuilder sb, Chart chart, double pageHeight)
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
            sb.AppendLine("BT");
            sb.AppendLine("/F1 12 Tf");
            sb.AppendLine("0 0 0 rg");
            sb.AppendLine($"{x + w/2:F2} {y + h - 20:F2} Td");
            sb.AppendLine($"({chart.Title}) Tj");
            sb.AppendLine("ET");
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
                RenderBarChart(sb, chart, plotX, plotY, plotW, plotH);
                break;
            case ChartType.Column:
                RenderColumnChart(sb, chart, plotX, plotY, plotW, plotH);
                break;
            case ChartType.Line:
                RenderLineChart(sb, chart, plotX, plotY, plotW, plotH);
                break;
            case ChartType.Pie:
                RenderPieChart(sb, chart, plotX, plotY, plotW, plotH);
                break;
            default:
                // Draw placeholder for unsupported chart types
                sb.AppendLine("BT");
                sb.AppendLine("/F1 10 Tf");
                sb.AppendLine("0.5 0.5 0.5 rg");
                sb.AppendLine($"{plotX + plotW/2:F2} {plotY + plotH/2:F2} Td");
                sb.AppendLine($"({chart.Type} Chart) Tj");
                sb.AppendLine("ET");
                break;
        }

        // Draw legend if exists
        if (chart.Legend != null)
        {
            RenderChartLegend(sb, chart, x, y, w, h);
        }
    }

    private void RenderBarChart(StringBuilder sb, Chart chart, double x, double y, double w, double h)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count == 0) return;

        var barWidth = w / (dataPoints.Count * 1.5);
        var maxValue = dataPoints.Max(dp => dp.Value);

        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var barHeight = (dataPoint.Value / maxValue) * h;
            var barX = x + i * barWidth * 1.5 + barWidth * 0.25;
            var barY = y + h - barHeight;

            // Set bar color
            sb.AppendLine("0.4 0.6 0.8 rg");
            sb.AppendLine($"{barX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re f");

            // Draw bar border
            sb.AppendLine("0.2 0.4 0.6 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{barX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re S");

            // Draw data label
            if (!string.IsNullOrEmpty(dataPoint.Category))
            {
                sb.AppendLine("BT");
                sb.AppendLine("/F1 8 Tf");
                sb.AppendLine("0 0 0 rg");
                sb.AppendLine($"{barX + barWidth/2:F2} {y - 10:F2} Td");
                sb.AppendLine($"({dataPoint.Category}) Tj");
                sb.AppendLine("ET");
            }
        }
    }

    private void RenderColumnChart(StringBuilder sb, Chart chart, double x, double y, double w, double h)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count == 0) return;

        var barWidth = w / (dataPoints.Count * 1.5);
        var maxValue = dataPoints.Max(dp => dp.Value);

        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var barHeight = (dataPoint.Value / maxValue) * h;
            var barX = x + i * barWidth * 1.5 + barWidth * 0.25;
            var barY = y + h - barHeight;

            // Set bar color
            sb.AppendLine("0.6 0.4 0.8 rg");
            sb.AppendLine($"{barX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re f");

            // Draw bar border
            sb.AppendLine("0.4 0.2 0.6 RG");
            sb.AppendLine("1 w");
            sb.AppendLine($"{barX:F2} {barY:F2} {barWidth:F2} {barHeight:F2} re S");

            // Draw data label
            if (!string.IsNullOrEmpty(dataPoint.Category))
            {
                sb.AppendLine("BT");
                sb.AppendLine("/F1 8 Tf");
                sb.AppendLine("0 0 0 rg");
                sb.AppendLine($"{barX + barWidth/2:F2} {y - 10:F2} Td");
                sb.AppendLine($"({dataPoint.Category}) Tj");
                sb.AppendLine("ET");
            }
        }
    }

    private void RenderLineChart(StringBuilder sb, Chart chart, double x, double y, double w, double h)
    {
        if (chart.Series.Count == 0) return;

        var series = chart.Series[0];
        var dataPoints = series.DataPoints;
        if (dataPoints.Count < 2) return;

        var maxValue = dataPoints.Max(dp => dp.Value);

        // Draw line
        sb.AppendLine("0.8 0.4 0.2 RG");
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

        // Draw data points
        sb.AppendLine("0.8 0.4 0.2 rg");
        foreach (var dataPoint in dataPoints)
        {
            var pointX = x + (dataPoints.IndexOf(dataPoint) / (double)(dataPoints.Count - 1)) * w;
            var pointY = y + h - (dataPoint.Value / maxValue) * h;
            sb.AppendLine($"{pointX - 3:F2} {pointY - 3:F2} 6 0 360 re f");
        }
    }

    private void RenderPieChart(StringBuilder sb, Chart chart, double x, double y, double w, double h)
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
        var colors = new[] {
            "0.8 0.2 0.2", "0.2 0.8 0.2", "0.2 0.2 0.8",
            "0.8 0.8 0.2", "0.8 0.2 0.8", "0.2 0.8 0.8"
        };

        for (int i = 0; i < dataPoints.Count; i++)
        {
            var dataPoint = dataPoints[i];
            var sliceAngle = (dataPoint.Value / totalValue) * 360;

            // Set slice color
            sb.AppendLine($"{colors[i % colors.Length]} rg");

            // Draw pie slice
            sb.AppendLine($"{centerX:F2} {centerY:F2} {radius:F2} {currentAngle:F2} {currentAngle + sliceAngle:F2} ar cn");

            currentAngle += sliceAngle;
        }

        // Draw pie border
        sb.AppendLine("0 0 0 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{centerX:F2} {centerY:F2} {radius:F2} 0 360 ar S");
    }

    private void RenderChartLegend(StringBuilder sb, Chart chart, double x, double y, double w, double h)
    {
        var legendX = x + 20;
        var legendY = y + 10;
        var legendItemHeight = 15;

        sb.AppendLine("BT");
        sb.AppendLine("/F1 8 Tf");
        sb.AppendLine("0 0 0 rg");

        for (int i = 0; i < chart.Series.Count; i++)
        {
            var series = chart.Series[i];
            var legendItemY = legendY + i * legendItemHeight;

            // Draw legend color box
            sb.AppendLine("Q");
            sb.AppendLine("0.4 0.6 0.8 rg");
            sb.AppendLine($"{legendX:F2} {legendItemY:F2} 10 10 re f");
            sb.AppendLine("BT");
            sb.AppendLine("/F1 8 Tf");
            sb.AppendLine("0 0 0 rg");

            // Draw legend text
            sb.AppendLine($"{legendX + 15:F2} {legendItemY + 2:F2} Td");
            sb.AppendLine($"({series.Name ?? $"Series {i + 1}"}) Tj");
        }

        sb.AppendLine("ET");
    }

    private void RenderSmartArt(StringBuilder sb, SmartArt smartArt, double pageHeight)
    {
        if (smartArt == null) return;

        var x = smartArt.Bounds.XPoints;
        var y = pageHeight - smartArt.Bounds.YPoints - smartArt.Bounds.HeightPoints;
        var w = smartArt.Bounds.WidthPoints;
        var h = smartArt.Bounds.HeightPoints;

        // Draw SmartArt background
        sb.AppendLine("0.95 0.95 0.95 rg");
        sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re f");

        // Draw SmartArt border
        sb.AppendLine("0.7 0.7 0.7 RG");
        sb.AppendLine("1 w");
        sb.AppendLine($"{x:F2} {y:F2} {w:F2} {h:F2} re S");

        // Draw SmartArt title if exists
        if (!string.IsNullOrEmpty(smartArt.Type))
        {
            sb.AppendLine("BT");
            sb.AppendLine("/F1 10 Tf");
            sb.AppendLine("0 0 0 rg");
            sb.AppendLine($"{x + 10:F2} {y + h - 15:F2} Td");
            sb.AppendLine($"({smartArt.Type}) Tj");
            sb.AppendLine("ET");
        }

        // Calculate content area
        var contentX = x + 20;
        var contentY = y + 20;
        var contentW = w - 40;
        var contentH = h - 40;

        // Render based on SmartArt type
        switch (SmartArt.GetSmartArtType(smartArt.Type))
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
                sb.AppendLine("BT");
                sb.AppendLine("/F1 10 Tf");
                sb.AppendLine("0.5 0.5 0.5 rg");
                sb.AppendLine($"{contentX + contentW/2:F2} {contentY + contentH/2:F2} Td");
                sb.AppendLine($"({SmartArt.GetSmartArtType(smartArt.Type)} SmartArt) Tj");
                sb.AppendLine("ET");
                break;
        }
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
            if (node.TextRuns.Count > 0)
            {
                sb.AppendLine("BT");
                sb.AppendLine($"{nodeX + 15:F2} {nodeY + nodeH/2:F2} Td");
                
                foreach (var run in node.TextRuns)
                {
                    // Set font size
                    sb.AppendLine($"/F1 {run.FontSize} Tf");
                    
                    // Set text color
                    var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                    
                    // Draw text
                    sb.AppendLine($"({run.Text}) Tj");
                }
                
                sb.AppendLine("ET");
            }

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
            if (node.TextRuns.Count > 0)
            {
                sb.AppendLine("BT");
                sb.AppendLine($"{nodeX + 10:F2} {nodeY + nodeHeight/2:F2} Td");
                
                foreach (var run in node.TextRuns)
                {
                    // Set font size
                    sb.AppendLine($"/F1 {run.FontSize} Tf");
                    
                    // Set text color
                    var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                    
                    // Draw text
                    sb.AppendLine($"({run.Text}) Tj");
                }
                
                sb.AppendLine("ET");
            }

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
            if (node.TextRuns.Count > 0)
            {
                sb.AppendLine("BT");
                sb.AppendLine($"{nodeX + 10:F2} {nodeY + nodeHeight/2:F2} Td");
                
                foreach (var run in node.TextRuns)
                {
                    // Set font size
                    sb.AppendLine($"/F1 {run.FontSize} Tf");
                    
                    // Set text color
                    var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                    
                    // Draw text
                    sb.AppendLine($"({run.Text}) Tj");
                }
                
                sb.AppendLine("ET");
            }

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
        if (rootNode.TextRuns.Count > 0)
        {
            sb.AppendLine("BT");
            sb.AppendLine($"{rootX + rootWidth/2:F2} {rootY + rootHeight/2:F2} Td");
            
            foreach (var run in rootNode.TextRuns)
            {
                // Set font size
                sb.AppendLine($"/F1 {run.FontSize} Tf");
                
                // Set text color
                var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                
                // Draw text
                sb.AppendLine($"({run.Text}) Tj");
            }
            
            sb.AppendLine("ET");
        }

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
                if (node.TextRuns.Count > 0)
                {
                    sb.AppendLine("BT");
                    sb.AppendLine($"{childX + childWidth/2:F2} {childY + childHeight/2:F2} Td");
                    
                    foreach (var run in node.TextRuns)
                    {
                        // Set font size
                        sb.AppendLine($"/F1 {run.FontSize} Tf");
                        
                        // Set text color
                        var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                        var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                        var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                        sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                        
                        // Draw text
                        sb.AppendLine($"({run.Text}) Tj");
                    }
                    
                    sb.AppendLine("ET");
                }

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
            if (node.TextRuns.Count > 0)
            {
                sb.AppendLine("BT");
                sb.AppendLine($"{nodeX + nodeWidth/2:F2} {nodeY + nodeHeight/2:F2} Td");
                
                foreach (var run in node.TextRuns)
                {
                    // Set font size
                    sb.AppendLine($"/F1 {run.FontSize} Tf");
                    
                    // Set text color
                    var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                    
                    // Draw text
                    sb.AppendLine($"({run.Text}) Tj");
                }
                
                sb.AppendLine("ET");
            }

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
            if (node.TextRuns.Count > 0)
            {
                sb.AppendLine("BT");
                sb.AppendLine($"{cellX + cellWidth/2:F2} {cellY + cellHeight/2:F2} Td");
                
                foreach (var run in node.TextRuns)
                {
                    // Set font size
                    sb.AppendLine($"/F1 {run.FontSize} Tf");
                    
                    // Set text color
                    var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                    
                    // Draw text
                    sb.AppendLine($"({run.Text}) Tj");
                }
                
                sb.AppendLine("ET");
            }
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
            if (node.TextRuns.Count > 0)
            {
                sb.AppendLine("BT");
                sb.AppendLine($"{layerX + layerWidth/2:F2} {layerY + layerHeight/2:F2} Td");
                
                foreach (var run in node.TextRuns)
                {
                    // Set font size
                    sb.AppendLine($"/F1 {run.FontSize} Tf");
                    
                    // Set text color
                    var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                    sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                    
                    // Draw text
                    sb.AppendLine($"({run.Text}) Tj");
                }
                
                sb.AppendLine("ET");
            }
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
        if (smartArt.Nodes.Count > 0 && smartArt.Nodes[0].TextRuns.Count > 0)
        {
            sb.AppendLine("BT");
            sb.AppendLine($"{centerX:F2} {centerY:F2} Td");
            
            foreach (var run in smartArt.Nodes[0].TextRuns)
            {
                // Set font size
                sb.AppendLine($"/F1 {run.FontSize} Tf");
                
                // Set text color
                var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                
                // Draw text
                sb.AppendLine($"({run.Text}) Tj");
            }
            
            sb.AppendLine("ET");
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
                if (node.TextRuns.Count > 0)
                {
                    sb.AppendLine("BT");
                    sb.AppendLine($"{nodeX + nodeWidth/2:F2} {nodeY + nodeHeight/2:F2} Td");
                    
                    foreach (var run in node.TextRuns)
                    {
                        // Set font size
                        sb.AppendLine($"/F1 {run.FontSize} Tf");
                        
                        // Set text color
                        var r = int.Parse(run.Color.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                        var g = int.Parse(run.Color.Substring(2, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                        var b = int.Parse(run.Color.Substring(4, 2), System.Globalization.NumberStyles.HexNumber) / 255.0;
                        sb.AppendLine($"{r:F2} {g:F2} {b:F2} rg");
                        
                        // Draw text
                        sb.AppendLine($"({run.Text}) Tj");
                    }
                    
                    sb.AppendLine("ET");
                }

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

    private Color InterpolateGradientColor(List<NPptxToPdf.GradientStop> stops, double position)
    {
        // Find the two stops that bracket the position
        NPptxToPdf.GradientStop? lower = null;
        NPptxToPdf.GradientStop? upper = null;

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
