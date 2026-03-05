using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class NotesSlide
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? SlideId { get; }
    public List<Shape> Shapes { get; } = new();
    public List<Picture> Pictures { get; } = new();
    public XElement? NotesText { get; private set; }

    public NotesSlide(XElement element, string? slideId = null)
    {
        _element = element;
        SlideId = slideId;
        Parse();
    }

    private void Parse()
    {
        var cSld = _element.Element(P + "cSld");
        if (cSld == null) return;

        // Parse shape tree
        var spTree = cSld.Element(P + "spTree");
        if (spTree == null) return;

        // Parse shapes
        foreach (var sp in spTree.Elements(P + "sp"))
        {
            var shape = new Shape(sp);

            // Check if this is the notes placeholder
            if (shape.PlaceholderType == PlaceholderType.Body)
            {
                // Get text body from shape
                var txBody = sp.Element(A + "txBody");
                if (txBody != null)
                {
                    NotesText = txBody;
                }
            }

            Shapes.Add(shape);
        }

        // Parse pictures
        foreach (var pic in spTree.Elements(P + "pic"))
        {
            Pictures.Add(new Picture(pic));
        }
    }
}

public class NotesMaster
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public List<Shape> Shapes { get; } = new();
    public List<Picture> Pictures { get; } = new();
    public Background? Background { get; private set; }
    public TextStyles? TextStyles { get; private set; }

    public NotesMaster(XElement element)
    {
        _element = element;
        Parse();
    }

    private void Parse()
    {
        var cSld = _element.Element(P + "cSld");
        if (cSld == null) return;

        // Parse background
        var bg = cSld.Element(P + "bg");
        if (bg != null)
        {
            Background = new Background(bg);
        }

        // Parse shape tree
        var spTree = cSld.Element(P + "spTree");
        if (spTree == null) return;

        // Parse shapes
        foreach (var sp in spTree.Elements(P + "sp"))
        {
            Shapes.Add(new Shape(sp));
        }

        // Parse pictures
        foreach (var pic in spTree.Elements(P + "pic"))
        {
            Pictures.Add(new Picture(pic));
        }

        // Parse text styles
        var txStyles = _element.Element(P + "txStyles");
        if (txStyles != null)
        {
            TextStyles = new TextStyles(txStyles);
        }
    }
}

public class Comment
{
    private readonly XElement _element;
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public int Id { get; }
    public int AuthorId { get; }
    public DateTime? Date { get; }
    public List<CommentText> TextElements { get; } = new();
    public long PositionX { get; }
    public long PositionY { get; }

    public Comment(XElement element)
    {
        _element = element;

        // Parse attributes
        if (int.TryParse(element.Attribute("id")?.Value, out var id))
            Id = id;

        if (int.TryParse(element.Attribute("authorId")?.Value, out var authorId))
            AuthorId = authorId;

        if (DateTime.TryParse(element.Attribute("dt")?.Value, out var dt))
            Date = dt;

        // Parse position
        if (long.TryParse(element.Attribute("x")?.Value, out var x))
            PositionX = x;

        if (long.TryParse(element.Attribute("y")?.Value, out var y))
            PositionY = y;

        // Parse text
        var textLst = element.Element(P + "text");
        if (textLst != null)
        {
            foreach (var r in textLst.Elements(P + "r"))
            {
                TextElements.Add(new CommentText(r));
            }
        }
    }

    public string GetFullText()
    {
        return string.Join("", TextElements.Select(t => t.Text));
    }
}

public class CommentText
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? Text { get; }
    public TextRunProperties? Properties { get; }

    public CommentText(XElement element)
    {
        _element = element;

        // Parse text content
        var t = element.Element(A + "t");
        Text = t?.Value;

        // Parse run properties
        var rPr = element.Element(A + "rPr");
        if (rPr != null)
        {
            Properties = new TextRunProperties(rPr);
        }
    }
}

public class CommentAuthor
{
    private readonly XElement _element;

    public int Id { get; }
    public string? Name { get; }
    public string? Initials { get; }
    public int ColorIndex { get; }

    public CommentAuthor(XElement element)
    {
        _element = element;

        if (int.TryParse(element.Attribute("id")?.Value, out var id))
            Id = id;

        Name = element.Attribute("name")?.Value;
        Initials = element.Attribute("initials")?.Value;

        if (int.TryParse(element.Attribute("colorIdx")?.Value, out var colorIdx))
            ColorIndex = colorIdx;
    }
}

public class CommentList
{
    public List<Comment> Comments { get; } = new();
    public List<CommentAuthor> Authors { get; } = new();

    public void AddComment(Comment comment)
    {
        Comments.Add(comment);
    }

    public void AddAuthor(CommentAuthor author)
    {
        Authors.Add(author);
    }

    public CommentAuthor? GetAuthor(int authorId)
    {
        return Authors.FirstOrDefault(a => a.Id == authorId);
    }

    public List<Comment> GetCommentsForSlide(string slideId)
    {
        // Comments are associated with slides via the commentAuthors.xml and comments/*.xml structure
        // This would need to be implemented based on the specific PPTX structure
        return Comments;
    }
}
