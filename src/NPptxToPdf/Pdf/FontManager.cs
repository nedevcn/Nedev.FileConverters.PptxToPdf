using System.Text;

namespace NPptxToPdf.Pdf;

public class FontManager
{
    // Standard PDF fonts that don't need embedding
    private static readonly HashSet<string> StandardPdfFonts = new(StringComparer.OrdinalIgnoreCase)
    {
        "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique",
        "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
        "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
        "Symbol", "ZapfDingbats"
    };

    // Common font name mappings
    private static readonly Dictionary<string, string> FontNameMappings = new(StringComparer.OrdinalIgnoreCase)
    {
        // Arial family
        ["Arial"] = "Helvetica",
        ["Arial Black"] = "Helvetica-Bold",
        ["Arial Bold"] = "Helvetica-Bold",
        ["Arial Italic"] = "Helvetica-Oblique",
        ["Arial Bold Italic"] = "Helvetica-BoldOblique",

        // Times family
        ["Times New Roman"] = "Times-Roman",
        ["Times"] = "Times-Roman",
        ["Times New Roman Bold"] = "Times-Bold",
        ["Times New Roman Italic"] = "Times-Italic",
        ["Times New Roman Bold Italic"] = "Times-BoldItalic",

        // Courier family
        ["Courier New"] = "Courier",
        ["Courier New Bold"] = "Courier-Bold",
        ["Courier New Italic"] = "Courier-Oblique",
        ["Courier New Bold Italic"] = "Courier-BoldOblique",

        // Common sans-serif fonts
        ["Verdana"] = "Helvetica",
        ["Tahoma"] = "Helvetica",
        ["Trebuchet MS"] = "Helvetica",
        ["Geneva"] = "Helvetica",
        ["Calibri"] = "Helvetica",
        ["Segoe UI"] = "Helvetica",
        ["Roboto"] = "Helvetica",
        ["Open Sans"] = "Helvetica",
        ["Lato"] = "Helvetica",
        ["Montserrat"] = "Helvetica",

        // Common serif fonts
        ["Georgia"] = "Times-Roman",
        ["Garamond"] = "Times-Roman",
        ["Palatino"] = "Times-Roman",
        ["Book Antiqua"] = "Times-Roman",
        ["Cambria"] = "Times-Roman",

        // Monospace fonts
        ["Consolas"] = "Courier",
        ["Monaco"] = "Courier",
        ["Lucida Console"] = "Courier",
        ["Monospace"] = "Courier",

        // Chinese fonts - map to standard fonts
        ["SimSun"] = "Helvetica",
        ["SimHei"] = "Helvetica",
        ["Microsoft YaHei"] = "Helvetica",
        ["Microsoft JhengHei"] = "Helvetica",
        ["NSimSun"] = "Helvetica",
        ["FangSong"] = "Helvetica",
        ["KaiTi"] = "Helvetica",
        ["LiSu"] = "Helvetica",
        ["YouYuan"] = "Helvetica",
        ["STSong"] = "Helvetica",
        ["STHeiti"] = "Helvetica",
        ["STKaiti"] = "Helvetica",
        ["STFangsong"] = "Helvetica",

        // Japanese fonts
        ["MS Mincho"] = "Times-Roman",
        ["MS Gothic"] = "Helvetica",
        ["Meiryo"] = "Helvetica",
        ["Yu Gothic"] = "Helvetica",
        ["Hiragino"] = "Helvetica",

        // Korean fonts
        ["Batang"] = "Times-Roman",
        ["Gulim"] = "Helvetica",
        ["Dotum"] = "Helvetica",
        ["Malgun Gothic"] = "Helvetica",
    };

    private readonly Dictionary<string, PdfFont> _fontCache = new();
    private int _fontIdCounter = 1;

    public string GetPdfFontName(string? fontName)
    {
        if (string.IsNullOrEmpty(fontName))
            return "Helvetica";

        // Check if it's already a standard PDF font
        if (StandardPdfFonts.Contains(fontName))
            return fontName;

        // Try direct mapping
        if (FontNameMappings.TryGetValue(fontName, out var mappedFont))
            return mappedFont;

        // Try to extract base font name (remove style suffixes)
        var baseName = ExtractBaseFontName(fontName);
        if (FontNameMappings.TryGetValue(baseName, out var baseMappedFont))
            return baseMappedFont;

        // Check if it contains style indicators
        if (fontName.Contains("Bold", StringComparison.OrdinalIgnoreCase) &&
            fontName.Contains("Italic", StringComparison.OrdinalIgnoreCase))
        {
            return "Helvetica-BoldOblique";
        }
        if (fontName.Contains("Bold", StringComparison.OrdinalIgnoreCase))
        {
            return "Helvetica-Bold";
        }
        if (fontName.Contains("Italic", StringComparison.OrdinalIgnoreCase) ||
            fontName.Contains("Oblique", StringComparison.OrdinalIgnoreCase))
        {
            return "Helvetica-Oblique";
        }

        // Default to Helvetica
        return "Helvetica";
    }

    private static string ExtractBaseFontName(string fontName)
    {
        // Remove common suffixes
        var suffixes = new[] { " Bold", " Italic", " Oblique", " Regular", " Normal" };
        var result = fontName;

        foreach (var suffix in suffixes)
        {
            if (result.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                result = result.Substring(0, result.Length - suffix.Length);
            }
        }

        return result.Trim();
    }

    public bool IsStandardFont(string fontName)
    {
        return StandardPdfFonts.Contains(GetPdfFontName(fontName));
    }

    public PdfFont GetOrCreateFont(string? fontName)
    {
        var pdfFontName = GetPdfFontName(fontName);

        if (_fontCache.TryGetValue(pdfFontName, out var cachedFont))
            return cachedFont;

        var font = new PdfFont(_fontIdCounter++, pdfFontName);
        _fontCache[pdfFontName] = font;
        return font;
    }

    public IEnumerable<PdfFont> GetAllFonts()
    {
        return _fontCache.Values;
    }

    public string EncodeText(string text, string fontName)
    {
        // For standard PDF fonts, we need to handle encoding properly
        // WinAnsiEncoding supports most Western European characters
        var sb = new StringBuilder();

        foreach (var c in text)
        {
            // Check if character is in WinAnsiEncoding range
            if (c <= 255)
            {
                // Escape special PDF characters
                switch (c)
                {
                    case '(':
                    case ')':
                    case '\\':
                        sb.Append('\\');
                        sb.Append(c);
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            else
            {
                // Character not in WinAnsiEncoding, try to find replacement or use default
                var replacement = GetCharacterReplacement(c);
                sb.Append(replacement);
            }
        }

        return sb.ToString();
    }

    private static char GetCharacterReplacement(char c)
    {
        // Map common Unicode characters to WinAnsiEncoding equivalents
        return c switch
        {
            '\u2013' => '\u2013', // en-dash
            '\u2014' => '\u2014', // em-dash
            '\u2018' => '\'',    // left single quote
            '\u2019' => '\'',    // right single quote
            '\u201C' => '"',     // left double quote
            '\u201D' => '"',     // right double quote
            '\u2022' => '\u2022', // bullet
            '\u2026' => '.',     // ellipsis
            '\u20AC' => 'E',     // Euro sign -> E
            '\u2122' => 'T',     // trademark -> T
            '\u00A9' => 'C',     // copyright -> C
            '\u00AE' => 'R',     // registered -> R
            _ => '?'             // unknown character
        };
    }

    public double GetFontWidth(string fontName, double fontSize, char c)
    {
        // Approximate character widths for standard fonts
        // These are rough estimates based on average character widths
        var baseWidth = fontName switch
        {
            var s when s.Contains("Courier") => 0.6, // Monospace
            var s when s.Contains("Times") => 0.45,  // Serif
            _ => 0.5                                  // Sans-serif (Helvetica)
        };

        return baseWidth * fontSize;
    }

    public double MeasureText(string text, string fontName, double fontSize)
    {
        var pdfFontName = GetPdfFontName(fontName);
        var totalWidth = 0.0;

        foreach (var c in text)
        {
            totalWidth += GetFontWidth(pdfFontName, fontSize, c);
        }

        return totalWidth;
    }
}
