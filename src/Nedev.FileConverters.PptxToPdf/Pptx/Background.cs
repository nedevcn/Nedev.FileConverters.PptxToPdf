using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class Background
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? SourcePath { get; }
    public Fill? Fill { get; }
    public int? StyleReferenceIndex { get; }
    public FillType FillType { get; }
    public Color? Color => Fill?.Type == FillType.Solid ? Fill.Color : null;
    public bool IsDefined => Fill != null || StyleReferenceIndex.HasValue;

    public Background(XElement element, string? sourcePath = null)
    {
        _element = element;
        SourcePath = sourcePath;

        var bgPr = element.Element(P + "bgPr");
        if (bgPr != null)
        {
            Fill = Shape.ParseFill(bgPr) ?? ParseLegacySolidFill(bgPr);
            FillType = Fill?.Type ?? FillType.None;
            return;
        }

        var bgRef = element.Element(P + "bgRef");
        if (bgRef != null)
        {
            StyleReferenceIndex = int.TryParse(bgRef.Attribute("idx")?.Value, out var idx) ? idx : null;
            FillType = FillType.None;
            return;
        }

        Fill = ParseLegacySolidFill(element);
        FillType = Fill?.Type ?? FillType.None;
    }

    public Fill? ResolveFill(Theme? theme)
    {
        if (Fill != null)
            return CloneFill(Fill);

        if (theme == null || !StyleReferenceIndex.HasValue)
            return null;

        var styleIndex = StyleReferenceIndex.Value - 1001;
        if (styleIndex < 0 || styleIndex >= theme.FormatScheme.BackgroundFillStyles.Count)
            return null;

        return CloneFill(theme.FormatScheme.BackgroundFillStyles[styleIndex].Fill);
    }

    public string? ResolveSourcePath(Theme? theme)
    {
        if (Fill != null)
            return SourcePath;

        return theme?.SourcePath ?? SourcePath;
    }

    private static Fill? ParseLegacySolidFill(XElement element)
    {
        var solidFill = element.Descendants(A + "solidFill").FirstOrDefault();
        if (solidFill == null)
            return null;

        var color = Shape.ParseColor(solidFill);
        if (!color.HasValue)
            return null;

        return new Fill
        {
            Type = FillType.Solid,
            Color = color.Value
        };
    }

    private static Fill? CloneFill(Fill? fill)
    {
        if (fill == null)
            return null;

        return new Fill
        {
            Type = fill.Type,
            Color = fill.Color,
            GradientStops = fill.GradientStops?.Select(stop => new GradientStop
            {
                Position = stop.Position,
                Color = stop.Color
            }).ToList(),
            GradientType = fill.GradientType,
            GradientAngle = fill.GradientAngle,
            PatternType = fill.PatternType,
            PatternForegroundColor = fill.PatternForegroundColor,
            PatternBackgroundColor = fill.PatternBackgroundColor,
            PictureRelationshipId = fill.PictureRelationshipId,
            PictureFillMode = fill.PictureFillMode,
            PictureFill = fill.PictureFill == null
                ? null
                : new PictureFill
                {
                    Blip = fill.PictureFill.Blip == null
                        ? null
                        : new Blip
                        {
                            Data = fill.PictureFill.Blip.Data == null
                                ? null
                                : (byte[])fill.PictureFill.Blip.Data.Clone()
                        }
                }
        };
    }
}
