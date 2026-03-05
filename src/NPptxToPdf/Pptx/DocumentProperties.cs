using System.Xml.Linq;

namespace NPptxToPdf.Pptx;

public class DocumentProperties
{
    private readonly XElement _element;
    private static readonly XNamespace CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DC = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DCTERMS = "http://purl.org/dc/terms/";

    public string? Title { get; }
    public string? Subject { get; }
    public string? Creator { get; }
    public string? Keywords { get; }
    public string? Description { get; }
    public string? LastModifiedBy { get; }
    public string? Revision { get; }
    public DateTime? Created { get; }
    public DateTime? Modified { get; }
    public string? Category { get; }
    public string? ContentStatus { get; }
    public string? Language { get; }
    public string? Version { get; }

    public DocumentProperties(XElement element)
    {
        _element = element;

        Title = element.Element(DC + "title")?.Value;
        Subject = element.Element(DC + "subject")?.Value;
        Creator = element.Element(DC + "creator")?.Value;
        Keywords = element.Element(CP + "keywords")?.Value;
        Description = element.Element(DC + "description")?.Value;
        LastModifiedBy = element.Element(CP + "lastModifiedBy")?.Value;
        Revision = element.Element(CP + "revision")?.Value;
        Category = element.Element(CP + "category")?.Value;
        ContentStatus = element.Element(CP + "contentStatus")?.Value;
        Language = element.Element(DC + "language")?.Value;
        Version = element.Element(CP + "version")?.Value;

        // Parse dates
        var createdStr = element.Element(DCTERMS + "created")?.Value;
        if (createdStr != null && DateTime.TryParse(createdStr, out var created))
        {
            Created = created;
        }

        var modifiedStr = element.Element(DCTERMS + "modified")?.Value;
        if (modifiedStr != null && DateTime.TryParse(modifiedStr, out var modified))
        {
            Modified = modified;
        }
    }
}
