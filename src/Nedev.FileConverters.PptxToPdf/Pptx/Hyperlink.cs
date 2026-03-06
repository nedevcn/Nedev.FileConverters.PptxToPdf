using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class Hyperlink
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public string? RelationshipId { get; }
    public HyperlinkType Type { get; }
    public string? Target { get; private set; }
    public string? Tooltip { get; }
    public bool IsExternal { get; private set; }
    public int? SlideIndex { get; private set; }

    public Hyperlink(XElement element)
    {
        _element = element;

        RelationshipId = element.Attribute(R + "id")?.Value;
        Tooltip = element.Attribute("tooltip")?.Value;

        // Parse action
        var action = element.Attribute("action")?.Value;
        if (!string.IsNullOrEmpty(action))
        {
            Type = ParseActionType(action);
            IsExternal = Type == HyperlinkType.Url || Type == HyperlinkType.File;
        }
        else
        {
            Type = HyperlinkType.Unknown;
            IsExternal = false;
        }
    }

    private static HyperlinkType ParseActionType(string action)
    {
        if (action.StartsWith("ppaction://hlinkshowjump?jump="))
        {
            var jumpType = action.Substring("ppaction://hlinkshowjump?jump=".Length);
            return jumpType switch
            {
                "firstslide" => HyperlinkType.FirstSlide,
                "lastslide" => HyperlinkType.LastSlide,
                "nextslide" => HyperlinkType.NextSlide,
                "previousslide" => HyperlinkType.PreviousSlide,
                "endshow" => HyperlinkType.EndShow,
                _ => HyperlinkType.SlideReference
            };
        }

        if (action.StartsWith("ppaction://hlinksldjumpto?slide="))
        {
            return HyperlinkType.SlideReference;
        }

        if (action.StartsWith("ppaction://hlinkfile"))
        {
            return HyperlinkType.File;
        }

        if (action.StartsWith("ppaction://hlinkpres"))
        {
            return HyperlinkType.Presentation;
        }

        return HyperlinkType.Url;
    }

    public void ResolveTarget(PptxDocument document, string slidePath)
    {
        if (string.IsNullOrEmpty(RelationshipId)) return;

        // Get relationship from slide relationships
        var relsPath = slidePath.Replace(".xml", ".xml.rels")
                                .Replace("/slides/", "/slides/_rels/");

        if (!document.TryGetPart(relsPath, out var relsData)) return;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Id")?.Value == RelationshipId);

        if (rel == null) return;

        var target = rel.Attribute("Target")?.Value;
        var targetMode = rel.Attribute("TargetMode")?.Value;

        if (targetMode == "External")
        {
            Target = target;
            IsExternal = true;
        }
        else if (!string.IsNullOrEmpty(target))
        {
            // Internal link
            Target = target;
            IsExternal = false;

            // Try to parse slide index
            if (target.Contains("slide"))
            {
                var match = System.Text.RegularExpressions.Regex.Match(target, @"slide(\d+)");
                if (match.Success && int.TryParse(match.Groups[1].Value, out var slideNum))
                {
                    SlideIndex = slideNum - 1; // Convert to 0-based index
                }
            }
        }
    }
}

public enum HyperlinkType
{
    Unknown,
    Url,
    File,
    Presentation,
    SlideReference,
    FirstSlide,
    LastSlide,
    NextSlide,
    PreviousSlide,
    EndShow,
    Bookmark,
    Email
}

public class HyperlinkManager
{
    private readonly Dictionary<string, List<Hyperlink>> _slideHyperlinks = new();

    public void RegisterHyperlink(string slidePath, Hyperlink hyperlink)
    {
        if (!_slideHyperlinks.ContainsKey(slidePath))
        {
            _slideHyperlinks[slidePath] = new List<Hyperlink>();
        }
        _slideHyperlinks[slidePath].Add(hyperlink);
    }

    public List<Hyperlink> GetHyperlinksForSlide(string slidePath)
    {
        return _slideHyperlinks.TryGetValue(slidePath, out var links) ? links : new List<Hyperlink>();
    }

    public void ResolveAllHyperlinks(PptxDocument document)
    {
        foreach (var entry in _slideHyperlinks)
        {
            foreach (var link in entry.Value)
            {
                link.ResolveTarget(document, entry.Key);
            }
        }
    }
}
