using System.Text;

namespace Nedev.PptxToPdf.Pdf;

public class FontEmbedder
{
    private readonly Dictionary<string, EmbeddedFont> _embeddedFonts = new();
    private readonly PdfDocument _document;

    public FontEmbedder(PdfDocument document)
    {
        _document = document;
    }

    public EmbeddedFont? EmbedFont(string fontName, string fontFilePath)
    {
        if (_embeddedFonts.TryGetValue(fontName, out var existingFont))
            return existingFont;

        if (!File.Exists(fontFilePath))
            return null;

        try
        {
            var fontData = File.ReadAllBytes(fontFilePath);
            var embeddedFont = new EmbeddedFont
            {
                FontName = fontName,
                FontData = fontData,
                ObjectNumber = _document.GetNextObjectNumber()
            };

            _embeddedFonts[fontName] = embeddedFont;
            return embeddedFont;
        }
        catch
        {
            return null;
        }
    }

    public EmbeddedFont? GetEmbeddedFont(string fontName)
    {
        _embeddedFonts.TryGetValue(fontName, out var font);
        return font;
    }

    public void WriteEmbeddedFonts(Stream stream)
    {
        foreach (var font in _embeddedFonts.Values)
        {
            WriteFontObject(stream, font);
        }
    }

    private void WriteFontObject(Stream stream, EmbeddedFont font)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{font.ObjectNumber} 0 obj");
        sb.AppendLine("<<");
        sb.AppendLine("/Type /Font");
        sb.AppendLine("/Subtype /Type0");
        sb.AppendLine($"/BaseFont /{font.FontName}");
        sb.AppendLine("/Encoding /Identity-H");
        sb.AppendLine("/DescendantFonts [");
        sb.AppendLine($"{font.ObjectNumber + 1} 0 R");
        sb.AppendLine("]");
        sb.AppendLine(">>");
        sb.AppendLine("endobj");

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        stream.Write(bytes, 0, bytes.Length);

        // Write CIDFont
        WriteCIDFont(stream, font);

        // Write FontDescriptor
        WriteFontDescriptor(stream, font);

        // Write FontFile
        WriteFontFile(stream, font);
    }

    private void WriteCIDFont(Stream stream, EmbeddedFont font)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{font.ObjectNumber + 1} 0 obj");
        sb.AppendLine("<<");
        sb.AppendLine("/Type /Font");
        sb.AppendLine("/Subtype /CIDFontType2");
        sb.AppendLine($"/BaseFont /{font.FontName}");
        sb.AppendLine("/CIDSystemInfo <<");
        sb.AppendLine("/Registry (Adobe)");
        sb.AppendLine("/Ordering (Identity)");
        sb.AppendLine("/Supplement 0");
        sb.AppendLine(">>");
        sb.AppendLine($"/FontDescriptor {font.ObjectNumber + 2} 0 R");
        sb.AppendLine($"/W {font.ObjectNumber + 4} 0 R");
        sb.AppendLine(">>");
        sb.AppendLine("endobj");

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        stream.Write(bytes, 0, bytes.Length);
    }

    private void WriteFontDescriptor(Stream stream, EmbeddedFont font)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{font.ObjectNumber + 2} 0 obj");
        sb.AppendLine("<<");
        sb.AppendLine("/Type /FontDescriptor");
        sb.AppendLine($"/FontName /{font.FontName}");
        sb.AppendLine("/Flags 4");
        sb.AppendLine("/FontBBox [-500 -300 1000 800]");
        sb.AppendLine("/ItalicAngle 0");
        sb.AppendLine("/Ascent 800");
        sb.AppendLine("/Descent -200");
        sb.AppendLine("/CapHeight 700");
        sb.AppendLine("/StemV 80");
        sb.AppendLine($"/FontFile2 {font.ObjectNumber + 3} 0 R");
        sb.AppendLine(">>");
        sb.AppendLine("endobj");

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        stream.Write(bytes, 0, bytes.Length);
    }

    private void WriteFontFile(Stream stream, EmbeddedFont font)
    {
        // Write FontFile2 object
        var sb = new StringBuilder();
        sb.AppendLine($"{font.ObjectNumber + 3} 0 obj");
        sb.AppendLine("<<");
        sb.AppendLine($"/Length {font.FontData.Length}");
        sb.AppendLine("/Length1 " + font.FontData.Length);
        sb.AppendLine(">>");
        sb.AppendLine("stream");

        var headerBytes = Encoding.UTF8.GetBytes(sb.ToString());
        stream.Write(headerBytes, 0, headerBytes.Length);
        stream.Write(font.FontData, 0, font.FontData.Length);

        var endBytes = Encoding.UTF8.GetBytes("\nendstream\nendobj\n");
        stream.Write(endBytes, 0, endBytes.Length);

        // Write Widths object
        WriteWidths(stream, font);
    }

    private void WriteWidths(Stream stream, EmbeddedFont font)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"{font.ObjectNumber + 4} 0 obj");
        sb.AppendLine("[");
        // Simplified width array - all characters have width 500
        sb.AppendLine("0 [500]");
        sb.AppendLine("]");
        sb.AppendLine("endobj");

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        stream.Write(bytes, 0, bytes.Length);
    }
}

public class EmbeddedFont
{
    public string FontName { get; set; } = string.Empty;
    public byte[] FontData { get; set; } = Array.Empty<byte>();
    public int ObjectNumber { get; set; }
}
