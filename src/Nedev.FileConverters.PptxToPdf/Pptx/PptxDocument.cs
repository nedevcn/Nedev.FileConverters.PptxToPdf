using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class PptxDocument : IDisposable
{
    private readonly ZipArchive _archive;
    private readonly Dictionary<string, byte[]> _parts = new();
    private readonly Dictionary<string, Theme> _themesByPath = new(StringComparer.OrdinalIgnoreCase);
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public Presentation? Presentation { get; private set; }
    public List<Slide> Slides { get; } = new();
    public List<SlideMaster> SlideMasters { get; } = new();
    public List<SlideLayout> SlideLayouts { get; } = new();
    public Theme? Theme { get; private set; }
    public DocumentProperties? Properties { get; private set; }

    private PptxDocument(ZipArchive archive)
    {
        _archive = archive;
    }

    public static PptxDocument Open(string filePath)
    {
        var stream = File.OpenRead(filePath);
        var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var doc = new PptxDocument(archive);
        doc.Load();
        return doc;
    }

    public static PptxDocument Open(Stream stream)
    {
        var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var doc = new PptxDocument(archive);
        doc.Load();
        return doc;
    }

    private void Load()
    {
        try
        {
            LoadParts();
            LoadPresentation();
            LoadThemes();
            LoadSlideMasters();
            LoadSlideLayouts();
            LoadSlides();
            LoadDocumentProperties();
        }
        catch (Exception ex)
        {
            // Log error but continue with partial loading
            Console.WriteLine($"Error loading PPTX document: {ex.Message}");
        }
    }

    private void LoadParts()
    {
        try
        {
            foreach (var entry in _archive.Entries)
            {
                try
                {
                    if (entry.FullName.EndsWith(".xml") || entry.FullName.EndsWith(".rels"))
                    {
                        using var stream = entry.Open();
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        _parts[entry.FullName] = ms.ToArray();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading part {entry.FullName}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading parts: {ex.Message}");
        }
    }

    private void LoadPresentation()
    {
        try
        {
            if (_parts.TryGetValue("ppt/presentation.xml", out var data))
            {
                var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                Presentation = new Presentation(xml.Root!);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading presentation: {ex.Message}");
        }
    }

    private void LoadSlideMasters()
    {
        try
        {
            // Find all slide master files
            var masterFiles = _parts.Keys
                .Where(k => k.StartsWith("ppt/slideMasters/slideMaster") && k.EndsWith(".xml"))
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var masterFile in masterFiles)
            {
                try
                {
                    if (_parts.TryGetValue(masterFile, out var data))
                    {
                        var theme = GetThemeForMaster(masterFile);
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var colorMap = ColorMap.FromElement(xml.Root?.Element(P + "clrMap"));
                        using var _ = Shape.UseSchemeColorResolver(CreateSchemeColorResolver(theme, colorMap));
                        var master = new SlideMaster(xml.Root!, masterFile)
                        {
                            Theme = theme
                        };
                        SlideMasters.Add(master);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading slide master {masterFile}: {ex.Message}");
                }
            }

            // If no masters loaded, create a default one
            if (SlideMasters.Count == 0)
            {
                try
                {
                    if (_parts.TryGetValue("ppt/slideMasters/slideMaster1.xml", out var data))
                    {
                        var theme = GetThemeForMaster("ppt/slideMasters/slideMaster1.xml");
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var colorMap = ColorMap.FromElement(xml.Root?.Element(P + "clrMap"));
                        using var _ = Shape.UseSchemeColorResolver(CreateSchemeColorResolver(theme, colorMap));
                        var master = new SlideMaster(xml.Root!, "ppt/slideMasters/slideMaster1.xml")
                        {
                            Theme = theme
                        };
                        SlideMasters.Add(master);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading default slide master: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading slide masters: {ex.Message}");
        }
    }

    private void LoadSlideLayouts()
    {
        try
        {
            // Find all slide layout files
            var layoutFiles = _parts.Keys
                .Where(k => k.StartsWith("ppt/slideLayouts/slideLayout") && k.EndsWith(".xml"))
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var layoutFile in layoutFiles)
            {
                try
                {
                    var masterPath = GetLayoutMasterPath(layoutFile);
                    var master = masterPath == null
                        ? null
                        : SlideMasters.FirstOrDefault(candidate =>
                            string.Equals(candidate.SourcePath, masterPath, StringComparison.OrdinalIgnoreCase));

                    if (_parts.TryGetValue(layoutFile, out var data))
                    {
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var colorMap = ColorMap.FromOverride(xml.Root?.Element(P + "clrMapOvr")) ?? master?.ColorMap;
                        using var _ = Shape.UseSchemeColorResolver(CreateSchemeColorResolver(master?.Theme ?? Theme, colorMap));
                        var layout = new SlideLayout(xml.Root!, layoutFile)
                        {
                            Master = master
                        };
                        SlideLayouts.Add(layout);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading slide layout {layoutFile}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading slide layouts: {ex.Message}");
        }
    }

    private string? GetLayoutMasterPath(string layoutFile)
    {
        var relsFile = GetRelationshipsPathForPart(layoutFile);
        if (!_parts.TryGetValue(relsFile, out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Type")?.Value.Contains("slideMaster") == true);

        var target = rel?.Attribute("Target")?.Value;
        if (string.IsNullOrEmpty(target))
            return null;

        return ResolvePartPath(layoutFile, target);
    }

    private void LoadThemes()
    {
        try
        {
            foreach (var path in _parts.Keys
                         .Where(path => path.StartsWith("ppt/theme/theme", StringComparison.OrdinalIgnoreCase) &&
                                        path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                         .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    LoadThemePart(path);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading theme from {path}: {ex.Message}");
                }
            }

            foreach (var relsPath in _parts.Keys
                         .Where(path => path.StartsWith("ppt/slideMasters/_rels/", StringComparison.OrdinalIgnoreCase) &&
                                        path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                         .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    if (!_parts.TryGetValue(relsPath, out var data))
                        continue;

                    var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                    XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

                    var themeRel = relsXml.Root?.Elements(ns + "Relationship")
                        .FirstOrDefault(e => e.Attribute("Type")?.Value.Contains("theme", StringComparison.OrdinalIgnoreCase) == true);

                    var target = themeRel?.Attribute("Target")?.Value;
                    if (string.IsNullOrEmpty(target))
                        continue;

                    var sourcePartPath = relsPath.Replace("/_rels/", "/").Replace(".rels", string.Empty, StringComparison.OrdinalIgnoreCase);
                    var themePath = ResolvePartPath(sourcePartPath, target);
                    LoadThemePart(themePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading theme from {relsPath}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading theme: {ex.Message}");
        }
    }

    private void LoadSlides()
    {
        try
        {
            if (Presentation == null) return;

            foreach (var slideId in Presentation.SlideIds)
            {
                try
                {
                    var slidePath = GetSlidePathFromRId(slideId);
                    if (slidePath != null && _parts.TryGetValue(slidePath, out var data))
                    {
                        SlideLayout? layout = null;
                        var layoutRId = GetSlideLayoutRId(slidePath);
                        if (layoutRId != null)
                        {
                            var layoutPath = GetLayoutPathFromRId(layoutRId, slidePath);
                            if (layoutPath != null)
                            {
                                layout = SlideLayouts.FirstOrDefault(candidate =>
                                    string.Equals(candidate.SourcePath, layoutPath, StringComparison.OrdinalIgnoreCase));
                            }
                        }

                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var colorMap = ColorMap.FromOverride(xml.Root?.Element(P + "clrMapOvr"))
                            ?? layout?.ColorMap
                            ?? layout?.Master?.ColorMap;
                        using var _ = Shape.UseSchemeColorResolver(CreateSchemeColorResolver(layout?.Master?.Theme ?? Theme, colorMap));
                        var slide = new Slide(xml.Root!, slideId, slidePath)
                        {
                            Layout = layout
                        };
                        LoadCharts(slide, layout?.Master?.Theme ?? Theme, colorMap);
                        LoadSmartArts(slide);

                        Slides.Add(slide);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading slide {slideId}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading slides: {ex.Message}");
        }
    }

    private string? GetSlideLayoutRId(string slideFile)
    {
        var relsFile = slideFile.Replace(".xml", ".xml.rels").Replace("/slides/", "/slides/_rels/");
        if (!_parts.TryGetValue(relsFile, out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Type")?.Value.Contains("slideLayout") == true);

        return rel?.Attribute("Id")?.Value;
    }

    private string? GetLayoutPathFromRId(string rId, string slidePath)
    {
        return GetRelatedPartPath(slidePath, rId);
    }

    private void LoadDocumentProperties()
    {
        try
        {
            if (_parts.TryGetValue("docProps/core.xml", out var data))
            {
                var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                Properties = new DocumentProperties(xml.Root!);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading document properties: {ex.Message}");
        }
    }

    private string? GetSlidePathFromRId(string rId)
    {
        return GetRelatedPartPath("ppt/presentation.xml", rId);
    }

    private void LoadCharts(Slide slide, Theme? theme, ColorMap? colorMap)
    {
        if (slide.ChartReferences.Count == 0)
            return;

        foreach (var chartReference in slide.ChartReferences)
        {
            try
            {
                var chartPath = GetRelatedPartPath(slide.SourcePath, chartReference.RelationshipId);
                if (chartPath == null || !_parts.TryGetValue(chartPath, out var chartData))
                    continue;

                var chartXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(chartData));
                if (chartXml.Root == null)
                    continue;

                using var _ = Shape.UseSchemeColorResolver(CreateSchemeColorResolver(theme, colorMap));
                slide.Charts.Add(new Chart(chartXml.Root, chartReference.Bounds));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading chart {chartReference.RelationshipId} from {slide.SourcePath}: {ex.Message}");
            }
        }
    }

    private void LoadSmartArts(Slide slide)
    {
        if (slide.SmartArtReferences.Count == 0)
            return;

        foreach (var smartArtReference in slide.SmartArtReferences)
        {
            try
            {
                var dataModelRoot = LoadPartRoot(
                    smartArtReference.DataModelRelationshipId == null
                        ? null
                        : GetRelatedPartPath(slide.SourcePath, smartArtReference.DataModelRelationshipId));
                var layoutDefRoot = LoadPartRoot(
                    smartArtReference.LayoutRelationshipId == null
                        ? null
                        : GetRelatedPartPath(slide.SourcePath, smartArtReference.LayoutRelationshipId));
                if (dataModelRoot == null && layoutDefRoot == null)
                    continue;

                slide.SmartArts.Add(new SmartArt(dataModelRoot, layoutDefRoot, smartArtReference.Bounds));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading SmartArt from {slide.SourcePath}: {ex.Message}");
            }
        }
    }

    private XElement? LoadPartRoot(string? partPath)
    {
        if (string.IsNullOrEmpty(partPath) || !_parts.TryGetValue(partPath, out var data))
            return null;

        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
        return xml.Root;
    }

    public byte[]? GetImageData(string imagePath)
    {
        var fullPath = imagePath.Replace('\\', '/').TrimStart('/');
        if (!fullPath.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase))
        {
            fullPath = $"ppt/{fullPath}";
        }

        if (_parts.TryGetValue(fullPath, out var data))
            return data;

        var entry = _archive.GetEntry(fullPath);
        if (entry != null)
        {
            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        return null;
    }

    public string? GetImagePathFromRId(string sourcePartPath, string rId)
    {
        return GetRelatedPartPath(sourcePartPath, rId);
    }

    private Theme? GetThemeForMaster(string masterFile)
    {
        var themePath = GetMasterThemePath(masterFile);
        if (themePath != null)
            return LoadThemePart(themePath) ?? Theme;

        return Theme;
    }

    private string? GetMasterThemePath(string masterFile)
    {
        var relsFile = GetRelationshipsPathForPart(masterFile);
        if (!_parts.TryGetValue(relsFile, out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Type")?.Value.Contains("theme", StringComparison.OrdinalIgnoreCase) == true);

        var target = rel?.Attribute("Target")?.Value;
        if (string.IsNullOrEmpty(target))
            return null;

        return ResolvePartPath(masterFile, target);
    }

    private Theme? LoadThemePart(string path)
    {
        if (_themesByPath.TryGetValue(path, out var existingTheme))
            return existingTheme;

        if (!_parts.TryGetValue(path, out var data))
            return null;

        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
        var theme = new Theme(xml.Root!, path);
        _themesByPath[path] = theme;
        Theme ??= theme;
        return theme;
    }

    private static Func<SchemeColor, Color>? CreateSchemeColorResolver(Theme? theme, ColorMap? colorMap = null)
    {
        if (theme is null && colorMap is null)
            return null;

        return schemeColor =>
        {
            var resolvedSchemeColor = colorMap?.ResolveSchemeColor(schemeColor) ?? schemeColor;
            return theme?.GetSchemeColor(resolvedSchemeColor) ?? Color.FromSchemeColor(resolvedSchemeColor);
        };
    }

    private string? GetRelatedPartPath(string sourcePartPath, string relationshipId)
    {
        var relsFile = GetRelationshipsPathForPart(sourcePartPath);
        if (!_parts.TryGetValue(relsFile, out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Id")?.Value == relationshipId);

        var target = rel?.Attribute("Target")?.Value;
        if (string.IsNullOrEmpty(target))
            return null;

        return ResolvePartPath(sourcePartPath, target);
    }

    private static string GetRelationshipsPathForPart(string partPath)
    {
        var normalizedPath = partPath.Replace('\\', '/');
        var fileName = Path.GetFileName(normalizedPath);
        var directory = Path.GetDirectoryName(normalizedPath)?.Replace('\\', '/');

        return string.IsNullOrEmpty(directory)
            ? $"_rels/{fileName}.rels"
            : $"{directory}/_rels/{fileName}.rels";
    }

    private static string ResolvePartPath(string sourcePartPath, string target)
    {
        var normalizedTarget = target.Replace('\\', '/');
        if (normalizedTarget.StartsWith("/"))
            return normalizedTarget.TrimStart('/');

        var segments = sourcePartPath.Replace('\\', '/')
            .Split('/', StringSplitOptions.RemoveEmptyEntries)
            .ToList();

        if (segments.Count > 0)
            segments.RemoveAt(segments.Count - 1);

        foreach (var segment in normalizedTarget.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
                continue;

            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);

                continue;
            }

            segments.Add(segment);
        }

        return string.Join("/", segments);
    }

    public bool TryGetPart(string path, out byte[] data)
    {
        return _parts.TryGetValue(path, out data!);
    }

    public void Dispose()
    {
        _archive.Dispose();
        _parts.Clear();
    }
}
