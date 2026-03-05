using System.IO.Compression;
using System.Text;

namespace Nedev.PptxToPdf.Pdf;

public abstract class PdfObject
{
    public int Number { get; }

    protected PdfObject(int number)
    {
        Number = number;
    }

    public abstract void WriteTo(Stream stream);

    protected static void Write(Stream stream, string text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    protected static void WriteLine(Stream stream, string text)
    {
        Write(stream, text + "\n");
    }
}

public class PdfCatalog : PdfObject
{
    public PdfPages? Pages { get; set; }

    public PdfCatalog(int number) : base(number) { }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, "/Type /Catalog");
        if (Pages != null)
            WriteLine(stream, $"/Pages {Pages.Number} 0 R");
        WriteLine(stream, ">>");
    }
}

public class PdfPages : PdfObject
{
    private readonly List<PdfPage> _pages = new();

    public PdfPages(int number) : base(number) { }

    public void AddPage(PdfPage page)
    {
        _pages.Add(page);
    }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, "/Type /Pages");
        WriteLine(stream, $"/Count {_pages.Count}");
        Write(stream, "/Kids [");
        foreach (var page in _pages)
        {
            Write(stream, $"{page.Number} 0 R ");
        }
        WriteLine(stream, "]");
        WriteLine(stream, ">>");
    }
}

public class PdfPage : PdfObject
{
    public PdfPages? Parent { get; set; }
    public double Width { get; }
    public double Height { get; }
    public PdfContent? Content { get; set; }
    public List<PdfFont> Fonts { get; } = new();
    public List<PdfImage> Images { get; } = new();
    public List<PdfAnnotation> Annotations { get; } = new();
    public Dictionary<string, PdfExtGState> ExtGStates { get; } = new();

    public PdfPage(int number, double width, double height) : base(number)
    {
        Width = width;
        Height = height;
    }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, "/Type /Page");
        if (Parent != null)
            WriteLine(stream, $"/Parent {Parent.Number} 0 R");
        WriteLine(stream, $"/MediaBox [0 0 {Width:F2} {Height:F2}]");

        if (Fonts.Any() || Images.Any() || ExtGStates.Any())
        {
            WriteLine(stream, "/Resources <<");
            if (Fonts.Any())
            {
                WriteLine(stream, "/Font <<");
                foreach (var font in Fonts)
                {
                    WriteLine(stream, $"/F{font.Number} {font.Number} 0 R");
                }
                WriteLine(stream, ">>");
            }
            if (Images.Any())
            {
                WriteLine(stream, "/XObject <<");
                foreach (var image in Images)
                {
                    WriteLine(stream, $"/Im{image.Number} {image.Number} 0 R");
                }
                WriteLine(stream, ">>");
            }
            if (ExtGStates.Any())
            {
                WriteLine(stream, "/ExtGState <<");
                foreach (var kvp in ExtGStates)
                {
                    WriteLine(stream, $"/{kvp.Key} {kvp.Value.Number} 0 R");
                }
                WriteLine(stream, ">>");
            }
            WriteLine(stream, ">>");
        }

        if (Content != null)
            WriteLine(stream, $"/Contents {Content.Number} 0 R");

        if (Annotations.Any())
        {
            Write(stream, "/Annots [");
            foreach (var annotation in Annotations)
            {
                Write(stream, $"{annotation.Number} 0 R ");
            }
            WriteLine(stream, "]");
        }

        WriteLine(stream, ">>");
    }
}

public class PdfFont : PdfObject
{
    public string BaseFont { get; }
    public string Subtype { get; }
    public string Encoding { get; }
    public bool IsExternallyDefined { get; set; } // True if font is defined externally (e.g., embedded font)

    public PdfFont(int number, string baseFont, string subtype = "Type1", string encoding = "WinAnsiEncoding") : base(number)
    {
        BaseFont = baseFont;
        Subtype = subtype;
        // Use Unicode encoding for Chinese fonts
        if (IsChineseFont(baseFont))
        {
            Encoding = "Identity-H";
            Subtype = "Type0"; // Type0 font for composite fonts (Unicode)
        }
        else
        {
            Encoding = encoding;
        }
    }

    private static bool IsChineseFont(string fontName)
    {
        var chineseFontNames = new[]
        {
            "SimSun", "SimHei", "SimKai", "SimFang", "SimLi",
            "STSong", "STHeiti", "STKaiti", "STFangSong",
            "Adobe", "Song", "Hei", "Kai", "Fang",
            "YaHei", "Microsoft", "Chinese",
            "瀹嬩綋", "榛戜綋", "妤蜂綋", "浠垮畫", "闅朵功"
        };
        
        foreach (var name in chineseFontNames)
        {
            if (fontName.Contains(name, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    public override void WriteTo(Stream stream)
    {
        // If font is externally defined (e.g., embedded font), don't write anything
        // The font definition is written by EmbeddedFontManager
        if (IsExternallyDefined)
        {
            return;
        }

        WriteLine(stream, "<<");
        WriteLine(stream, $"/Type /Font");
        WriteLine(stream, $"/Subtype /{Subtype}");
        WriteLine(stream, $"/BaseFont /{BaseFont}");
        
        if (Subtype == "Type1")
        {
            WriteLine(stream, $"/Encoding /{Encoding}");
        }
        else if (Subtype == "Type0")
        {
            // For Type0 fonts (composite fonts for Unicode), we need to reference a descendant CID font
            WriteLine(stream, "/Encoding /Identity-H");
            WriteLine(stream, "/DescendantFonts [");
            WriteLine(stream, "<<");
            WriteLine(stream, "/Type /Font");
            WriteLine(stream, "/Subtype /CIDFontType2");
            WriteLine(stream, $"/BaseFont /{BaseFont}");
            WriteLine(stream, "/CIDSystemInfo <<");
            WriteLine(stream, "/Registry (Adobe)");
            WriteLine(stream, "/Ordering (Identity)");
            WriteLine(stream, "/Supplement 0");
            WriteLine(stream, ">>");
            WriteLine(stream, "/W [0 [500]]");
            WriteLine(stream, ">>");
            WriteLine(stream, "]");
        }
        
        WriteLine(stream, ">>");
    }
}

public class PdfContent : PdfObject
{
    private readonly MemoryStream _content = new();

    public PdfContent(int number) : base(number) { }

    public Stream Stream => _content;

    public void AddOperation(string operation)
    {
        var bytes = Encoding.UTF8.GetBytes(operation + "\n");
        _content.Write(bytes, 0, bytes.Length);
    }

    public override void WriteTo(Stream stream)
    {
        var data = _content.ToArray();

        WriteLine(stream, "<<");
        WriteLine(stream, $"/Length {data.Length}");
        WriteLine(stream, ">>");
        WriteLine(stream, "stream");
        stream.Write(data, 0, data.Length);
        WriteLine(stream, "");
        WriteLine(stream, "endstream");
    }
}

public class PdfImage : PdfObject
{
    public byte[] Data { get; }
    public int Width { get; }
    public int Height { get; }
    public bool IsJpeg { get; }

    public PdfImage(int number, byte[] data, int width, int height, bool isJpeg) : base(number)
    {
        Data = data;
        Width = width;
        Height = height;
        IsJpeg = isJpeg;
    }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, "/Type /XObject");
        WriteLine(stream, "/Subtype /Image");
        WriteLine(stream, $"/Width {Width}");
        WriteLine(stream, $"/Height {Height}");
        WriteLine(stream, "/ColorSpace /DeviceRGB");
        WriteLine(stream, "/BitsPerComponent 8");

        if (IsJpeg)
        {
            WriteLine(stream, "/Filter /DCTDecode");
        }

        WriteLine(stream, $"/Length {Data.Length}");
        WriteLine(stream, ">>");
        WriteLine(stream, "stream");
        stream.Write(Data, 0, Data.Length);
        WriteLine(stream, "");
        WriteLine(stream, "endstream");
    }
}

public class PdfAnnotation : PdfObject
{
    public string Type { get; set; }
    public string Subtype { get; set; }
    public double[] Rect { get; set; }
    public PdfAction? Action { get; set; }

    public PdfAnnotation(int number) : base(number)
    {
    }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, $"/Type {Type}");
        WriteLine(stream, $"/Subtype {Subtype}");
        Write(stream, $"/Rect [");
        for (int i = 0; i < Rect.Length; i++)
        {
            Write(stream, $"{Rect[i]:F2}");
            if (i < Rect.Length - 1) Write(stream, " ");
        }
        WriteLine(stream, "]");
        if (Action != null)
        {
            WriteLine(stream, $"/A {Action.Number} 0 R");
        }
        WriteLine(stream, ">>");
    }
}

public class PdfAction : PdfObject
{
    public string Type { get; set; }
    public string S { get; set; }
    public string URI { get; set; }

    public PdfAction(int number = 0) : base(number)
    {
    }

    public override void WriteTo(Stream stream)
    {
        WriteLine(stream, "<<");
        WriteLine(stream, $"/Type {Type}");
        WriteLine(stream, $"/S {S}");
        WriteLine(stream, $"/URI ({URI})");
        WriteLine(stream, ">>");
    }
}
