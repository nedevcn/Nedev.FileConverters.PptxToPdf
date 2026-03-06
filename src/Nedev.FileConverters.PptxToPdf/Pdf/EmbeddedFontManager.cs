using System.Text;

namespace Nedev.FileConverters.PptxToPdf.Pdf;

public class EmbeddedFontManager
{
    private readonly PdfDocument _document;
    private readonly Dictionary<string, EmbeddedFontInfo> _embeddedFonts = new();
    private int _nextFontObjectNumber = 1000; // Start from 1000 to avoid conflicts with regular objects

    public EmbeddedFontManager(PdfDocument document)
    {
        _document = document;
    }

    public EmbeddedFontInfo? EmbedSystemFont(string fontName)
    {
        if (_embeddedFonts.TryGetValue(fontName, out var existingFont))
            return existingFont;

        // Map font names to system font files
        string? fontPath = GetSystemFontPath(fontName);
        if (fontPath == null)
            return null;

        try
        {
            var fontData = File.ReadAllBytes(fontPath);
            var fontInfo = new EmbeddedFontInfo
            {
                FontName = fontName,
                FontData = fontData,
                // Object numbers will be assigned when writing
                Type0ObjectNumber = -1,
                CIDFontObjectNumber = -1,
                FontDescriptorObjectNumber = -1,
                FontFileObjectNumber = -1,
                WidthsObjectNumber = -1,
                ToUnicodeObjectNumber = -1
            };

            _embeddedFonts[fontName] = fontInfo;
            return fontInfo;
        }
        catch
        {
            return null;
        }
    }

    private string? GetSystemFontPath(string fontName)
    {
        // Prioritize TTF fonts over TTC for better compatibility
        var fontPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["MicrosoftYaHei"] = @"C:\Windows\Fonts\msyh.ttc",
            ["MicrosoftYaHeiBold"] = @"C:\Windows\Fonts\msyhbd.ttc",
            ["MicrosoftYaHeiLight"] = @"C:\Windows\Fonts\msyhl.ttc",
            ["SimSun"] = @"C:\Windows\Fonts\simsun.ttc",
            ["SimHei"] = @"C:\Windows\Fonts\simhei.ttf",
            ["SimKai"] = @"C:\Windows\Fonts\simkai.ttf",
            ["SimFang"] = @"C:\Windows\Fonts\simfang.ttf"
        };

        if (fontPaths.TryGetValue(fontName, out var path))
            return path;

        // Try to find by partial match
        foreach (var kvp in fontPaths)
        {
            if (fontName.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase) ||
                kvp.Key.Contains(fontName, StringComparison.OrdinalIgnoreCase))
                return kvp.Value;
        }

        // Check if file exists for common Chinese fonts
        // Prioritize TTF fonts for better embedding compatibility
        if (fontName.Contains("???") || fontName.Contains("SimHei"))
            return @"C:\Windows\Fonts\simhei.ttf";
        if (fontName.Contains("???") || fontName.Contains("SimKai"))
            return @"C:\Windows\Fonts\simkai.ttf";
        if (fontName.Contains("???") || fontName.Contains("SimFang"))
            return @"C:\Windows\Fonts\simfang.ttf";
        if (fontName.Contains("???") || fontName.Contains("SimSun"))
            return @"C:\Windows\Fonts\simsun.ttc";

        return null;
    }

    public void WriteEmbeddedFonts(PdfDocument document, Stream stream)
    {
        foreach (var font in _embeddedFonts.Values)
        {
            // Object numbers are already assigned by GetType0ObjectNumber during rendering
            // Just write the font objects
            WriteType0Font(document, stream, font);
            WriteCIDFont(document, stream, font);
            WriteFontDescriptor(document, stream, font);
            WriteFontFile(document, stream, font);
            WriteWidths(document, stream, font);
            WriteToUnicode(document, stream, font);
        }
    }
    
    // Get the Type0 object number for a font (assigns if not already assigned)
    public int GetType0ObjectNumber(string fontName)
    {
        if (_embeddedFonts.TryGetValue(fontName, out var font))
        {
            if (font.Type0ObjectNumber < 0)
            {
                // Assign object numbers from reserved range to avoid conflicts
                font.Type0ObjectNumber = _nextFontObjectNumber++;
                font.CIDFontObjectNumber = _nextFontObjectNumber++;
                font.FontDescriptorObjectNumber = _nextFontObjectNumber++;
                font.FontFileObjectNumber = _nextFontObjectNumber++;
                font.WidthsObjectNumber = _nextFontObjectNumber++;
                font.ToUnicodeObjectNumber = _nextFontObjectNumber++;
            }
            return font.Type0ObjectNumber;
        }
        return -1;
    }

    private void WriteType0Font(PdfDocument document, Stream stream, EmbeddedFontInfo font)
    {
        document.WriteObjectDirect(font.Type0ObjectNumber, s =>
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{font.Type0ObjectNumber} 0 obj");
            sb.AppendLine("<<");
            sb.AppendLine("/Type /Font");
            sb.AppendLine("/Subtype /Type0");
            sb.AppendLine($"/BaseFont /{font.FontName}");
            sb.AppendLine("/Encoding /Identity-H");
            sb.AppendLine($"/DescendantFonts [{font.CIDFontObjectNumber} 0 R]");
            sb.AppendLine($"/ToUnicode {font.ToUnicodeObjectNumber} 0 R");
            sb.AppendLine(">>");
            sb.AppendLine("endobj");

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            s.Write(bytes, 0, bytes.Length);
        });
    }

    private void WriteCIDFont(PdfDocument document, Stream stream, EmbeddedFontInfo font)
    {
        document.WriteObjectDirect(font.CIDFontObjectNumber, s =>
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{font.CIDFontObjectNumber} 0 obj");
            sb.AppendLine("<<");
            sb.AppendLine("/Type /Font");
            sb.AppendLine("/Subtype /CIDFontType2");
            sb.AppendLine($"/BaseFont /{font.FontName}");
            sb.AppendLine("/CIDSystemInfo <<");
            sb.AppendLine("/Registry (Adobe)");
            sb.AppendLine("/Ordering (Identity)");
            sb.AppendLine("/Supplement 0");
            sb.AppendLine(">>");
            sb.AppendLine($"/FontDescriptor {font.FontDescriptorObjectNumber} 0 R");
            sb.AppendLine($"/W {font.WidthsObjectNumber} 0 R");
            sb.AppendLine(">>");
            sb.AppendLine("endobj");

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            s.Write(bytes, 0, bytes.Length);
        });
    }

    private void WriteFontDescriptor(PdfDocument document, Stream stream, EmbeddedFontInfo font)
    {
        document.WriteObjectDirect(font.FontDescriptorObjectNumber, s =>
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{font.FontDescriptorObjectNumber} 0 obj");
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
            sb.AppendLine($"/FontFile2 {font.FontFileObjectNumber} 0 R");
            sb.AppendLine(">>");
            sb.AppendLine("endobj");

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            s.Write(bytes, 0, bytes.Length);
        });
    }

    private void WriteFontFile(PdfDocument document, Stream stream, EmbeddedFontInfo font)
    {
        document.WriteObjectDirect(font.FontFileObjectNumber, s =>
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{font.FontFileObjectNumber} 0 obj");
            sb.AppendLine("<<");
            sb.AppendLine($"/Length {font.FontData.Length}");
            sb.AppendLine($"/Length1 {font.FontData.Length}");
            sb.AppendLine(">>");
            sb.AppendLine("stream");

            var headerBytes = Encoding.UTF8.GetBytes(sb.ToString());
            s.Write(headerBytes, 0, headerBytes.Length);
            s.Write(font.FontData, 0, font.FontData.Length);

            var endBytes = Encoding.UTF8.GetBytes("\nendstream\nendobj\n");
            s.Write(endBytes, 0, endBytes.Length);
        });
    }

    private void WriteWidths(PdfDocument document, Stream stream, EmbeddedFontInfo font)
    {
        document.WriteObjectDirect(font.WidthsObjectNumber, s =>
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{font.WidthsObjectNumber} 0 obj");
            sb.AppendLine("[");
            // Simplified width array for CJK characters
            // Define widths for common CJK character ranges
            sb.AppendLine("0 [500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500 500]");
            sb.AppendLine("]");
            sb.AppendLine("endobj");

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            s.Write(bytes, 0, bytes.Length);
        });
    }

    private void WriteToUnicode(PdfDocument document, Stream stream, EmbeddedFontInfo font)
    {
        document.WriteObjectDirect(font.ToUnicodeObjectNumber, s =>
        {
            // Create a simple ToUnicode CMap that maps character codes to Unicode
            var sb = new StringBuilder();
            sb.AppendLine($"{font.ToUnicodeObjectNumber} 0 obj");
            sb.AppendLine("<<");
            sb.AppendLine("/Type /CMap");
            sb.AppendLine("/CMapName /Adobe-Identity-UCS");
            sb.AppendLine("/CIDSystemInfo <<");
            sb.AppendLine("/Registry (Adobe)");
            sb.AppendLine("/Ordering (UCS)");
            sb.AppendLine("/Supplement 0");
            sb.AppendLine(">>");
            
            // CMap stream
            var cmapContent = @"/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
/CIDSystemInfo <<
/Registry (Adobe)
/Ordering (UCS)
/Supplement 0
>> def
/CMapName /Adobe-Identity-UCS def
/CMapType 2 def
1 begincodespacerange
<0000> <FFFF>
endcodespacerange
1 beginbfrange
<0000> <FFFF> <0000>
endbfrange
endcmap
CMapName currentdict /CMap defineresource pop
end
end";
            
            var cmapBytes = Encoding.ASCII.GetBytes(cmapContent);
            sb.AppendLine($"/Length {cmapBytes.Length}");
            sb.AppendLine(">>");
            sb.AppendLine("stream");
            
            var headerBytes = Encoding.UTF8.GetBytes(sb.ToString());
            s.Write(headerBytes, 0, headerBytes.Length);
            s.Write(cmapBytes, 0, cmapBytes.Length);
            
            var endBytes = Encoding.UTF8.GetBytes("\nendstream\nendobj\n");
            s.Write(endBytes, 0, endBytes.Length);
        });
    }

    public IReadOnlyDictionary<string, EmbeddedFontInfo> EmbeddedFonts => _embeddedFonts;
}

public class EmbeddedFontInfo
{
    public string FontName { get; set; } = string.Empty;
    public byte[] FontData { get; set; } = Array.Empty<byte>();
    public int Type0ObjectNumber { get; set; }
    public int CIDFontObjectNumber { get; set; }
    public int FontDescriptorObjectNumber { get; set; }
    public int FontFileObjectNumber { get; set; }
    public int WidthsObjectNumber { get; set; }
    public int ToUnicodeObjectNumber { get; set; }
}
