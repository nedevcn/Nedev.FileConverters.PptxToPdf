using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class PptxDocument : IDisposable
{
    private readonly ZipArchive _archive;
    private readonly Dictionary<string, byte[]> _parts = new();

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
            LoadSlideMasters();
            LoadSlideLayouts();
            LoadTheme();
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
            var masterFiles = _parts.Keys.Where(k => k.StartsWith("ppt/slideMasters/slideMaster") && k.EndsWith(".xml")).ToList();

            foreach (var masterFile in masterFiles)
            {
                try
                {
                    if (_parts.TryGetValue(masterFile, out var data))
                    {
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var master = new SlideMaster(xml.Root!);
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
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var master = new SlideMaster(xml.Root!);
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
            var layoutFiles = _parts.Keys.Where(k => k.StartsWith("ppt/slideLayouts/slideLayout") && k.EndsWith(".xml")).ToList();

            foreach (var layoutFile in layoutFiles)
            {
                try
                {
                    if (_parts.TryGetValue(layoutFile, out var data))
                    {
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var layout = new SlideLayout(xml.Root!);
                        SlideLayouts.Add(layout);

                        // Associate with master if possible
                        var masterRId = GetLayoutMasterRId(layoutFile);
                        if (masterRId != null)
                        {
                            var masterPath = GetMasterPathFromRId(masterRId);
                            // Could store this relationship if needed
                        }
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

    private string? GetLayoutMasterRId(string layoutFile)
    {
        var relsFile = layoutFile.Replace(".xml", ".xml.rels").Replace("/slideLayouts/", "/slideLayouts/_rels/");
        if (!_parts.TryGetValue(relsFile, out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Type")?.Value.Contains("slideMaster") == true);

        return rel?.Attribute("Id")?.Value;
    }

    private string? GetMasterPathFromRId(string rId)
    {
        // Check presentation relationships
        if (!_parts.TryGetValue("ppt/_rels/presentation.xml.rels", out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Id")?.Value == rId);

        if (rel == null) return null;

        var target = rel.Attribute("Target")?.Value;
        if (target == null) return null;

        return target.StartsWith("/") ? target.TrimStart('/') : $"ppt/{target}";
    }

    private void LoadTheme()
    {
        try
        {
            // Try to load theme from various locations
            var themePaths = new[]
            {
                "ppt/theme/theme1.xml",
                "ppt/theme/theme2.xml",
                "ppt/slideMasters/_rels/slideMaster1.xml.rels"
            };

            foreach (var path in themePaths)
            {
                try
                {
                    if (_parts.TryGetValue(path, out var data))
                    {
                        if (path.EndsWith(".rels"))
                        {
                            // Parse relationships to find theme
                            var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

                            var themeRel = relsXml.Root?.Elements(ns + "Relationship")
                                .FirstOrDefault(e => e.Attribute("Type")?.Value.Contains("theme") == true);

                            if (themeRel != null)
                            {
                                var target = themeRel.Attribute("Target")?.Value;
                                if (target != null)
                                {
                                    var themePath = target.StartsWith("/")
                                        ? target.TrimStart('/')
                                        : $"ppt/slideMasters/{target}";

                                    if (_parts.TryGetValue(themePath, out var themeData))
                                    {
                                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(themeData));
                                        Theme = new Theme(xml.Root!);
                                        return;
                                    }
                                }
                            }
                        }
                        else
                        {
                            var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                            Theme = new Theme(xml.Root!);
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading theme from {path}: {ex.Message}");
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
                        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(data));
                        var slide = new Slide(xml.Root!, slideId);

                        // Get layout for this slide
                        var layoutRId = GetSlideLayoutRId(slidePath);
                        if (layoutRId != null)
                        {
                            var layoutPath = GetLayoutPathFromRId(layoutRId, slidePath);
                            if (layoutPath != null)
                            {
                                var layout = SlideLayouts.FirstOrDefault(l =>
                                    _parts.Any(p => p.Key == layoutPath && p.Value != null));
                                if (layout != null)
                                {
                                    slide.Layout = layout;
                                }
                            }
                        }

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
        var relsFile = slidePath.Replace(".xml", ".xml.rels").Replace("/slides/", "/slides/_rels/");
        if (!_parts.TryGetValue(relsFile, out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Id")?.Value == rId);

        if (rel == null) return null;

        var target = rel.Attribute("Target")?.Value;
        if (target == null) return null;

        return target.StartsWith("/") ? target.TrimStart('/') : $"ppt/{target}";
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
        if (!_parts.TryGetValue("ppt/_rels/presentation.xml.rels", out var relsData))
            return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Id")?.Value == rId);

        if (rel == null) return null;

        var target = rel.Attribute("Target")?.Value;
        if (target == null) return null;

        return target.StartsWith("/") ? target.TrimStart('/') : $"ppt/{target}";
    }

    public byte[]? GetImageData(string imagePath)
    {
        var fullPath = imagePath.StartsWith("/") ? imagePath.TrimStart('/') : $"ppt/{imagePath}";

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

    public string? GetImagePathFromRId(string rId)
    {
        if (!_parts.TryGetValue("ppt/slides/_rels/slide1.xml.rels", out var relsData))
        {
            foreach (var part in _parts.Keys.Where(k => k.Contains("/_rels/") && k.Contains("slide")))
            {
                relsData = _parts[part];
                break;
            }
        }

        if (relsData == null) return null;

        var relsXml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(relsData));
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var rel = relsXml.Root?.Elements(ns + "Relationship")
            .FirstOrDefault(e => e.Attribute("Id")?.Value == rId);

        if (rel == null) return null;

        var target = rel.Attribute("Target")?.Value;
        if (target == null) return null;

        return target.StartsWith("/") ? target.TrimStart('/') : target;
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
