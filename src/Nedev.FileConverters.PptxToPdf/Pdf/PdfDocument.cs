using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

namespace Nedev.FileConverters.PptxToPdf.Pdf;

public class PdfDocument : IDisposable
{
    private readonly Stream _stream;
    private readonly List<PdfObject> _objects = new();
    private readonly Dictionary<int, long> _offsets = new();
    private int _objectNumber = 1;
    private PdfCatalog? _catalog;
    private PdfPages? _pages;
    private long _xrefOffset;
    private EmbeddedFontManager? _embeddedFontManager;

    public PdfDocument(string filePath)
    {
        _stream = File.Create(filePath);
    }

    public PdfDocument(Stream stream)
    {
        _stream = stream;
    }

    public EmbeddedFontManager GetEmbeddedFontManager()
    {
        _embeddedFontManager ??= new EmbeddedFontManager(this);
        return _embeddedFontManager;
    }

    public void Initialize()
    {
        WriteHeader();

        _catalog = new PdfCatalog(_objectNumber++);
        _pages = new PdfPages(_objectNumber++);
        _catalog.Pages = _pages;

        _objects.Add(_catalog);
        _objects.Add(_pages);
    }

    public PdfPage AddPage(double width, double height)
    {
        if (_catalog == null || _pages == null)
            throw new InvalidOperationException("Document not initialized");

        var page = new PdfPage(_objectNumber++, width, height);
        page.Parent = _pages;
        _pages.AddPage(page);
        _objects.Add(page);

        return page;
    }

    public PdfImage AddImage(byte[] imageData, int width, int height, bool isJpeg = false)
    {
        var image = new PdfImage(_objectNumber++, imageData, width, height, isJpeg);
        _objects.Add(image);
        return image;
    }

    public int GetNextObjectNumber()
    {
        return _objectNumber++;
    }

    public void AddObject(PdfObject obj)
    {
        if (!_objects.Contains(obj))
        {
            _objects.Add(obj);
        }
    }

    public void Save()
    {
        if (_catalog == null) return;

        foreach (var obj in _objects)
        {
            WriteObject(obj);
        }

        // Write embedded fonts if any
        _embeddedFontManager?.WriteEmbeddedFonts(this, _stream);

        WriteXref();
        WriteTrailer();
        _stream.Flush();
    }

    public void WriteObjectDirect(int objectNumber, Action<Stream> writeAction)
    {
        _offsets[objectNumber] = _stream.Position;
        writeAction(_stream);
    }

    private void WriteHeader()
    {
        WriteLine("%PDF-1.4");
        WriteLine("%\xE2\xE3\xCF\xD3");
    }

    private void WriteObject(PdfObject obj)
    {
        _offsets[obj.Number] = _stream.Position;
        WriteLine($"{obj.Number} 0 obj");
        obj.WriteTo(_stream);
        WriteLine("endobj");
    }

    private void WriteXref()
    {
        _xrefOffset = _stream.Position;
        WriteLine("xref");

        // Group offsets by consecutive ranges
        var sortedOffsets = _offsets.OrderBy(x => x.Key).ToList();
        var ranges = new List<(int start, int count, List<long> offsets)>();

        if (sortedOffsets.Count > 0)
        {
            var currentRange = (start: sortedOffsets[0].Key, count: 1, offsets: new List<long> { sortedOffsets[0].Value });

            for (int i = 1; i < sortedOffsets.Count; i++)
            {
                if (sortedOffsets[i].Key == sortedOffsets[i - 1].Key + 1)
                {
                    // Consecutive
                    currentRange.count++;
                    currentRange.offsets.Add(sortedOffsets[i].Value);
                }
                else
                {
                    // Non-consecutive, start new range
                    ranges.Add(currentRange);
                    currentRange = (sortedOffsets[i].Key, 1, new List<long> { sortedOffsets[i].Value });
                }
            }
            ranges.Add(currentRange);
        }

        // Write xref sections for each range
        // First section includes object 0 (free entry)
        if (ranges.Count > 0 && ranges[0].start == 1)
        {
            // First range starts at 1, so we need to include object 0
            WriteLine($"0 {ranges[0].count + 1}");
            WriteLine("0000000000 65535 f ");
            foreach (var offset in ranges[0].offsets)
            {
                WriteLine($"{offset:D10} 00000 n ");
            }
            ranges.RemoveAt(0);
        }
        else if (ranges.Count > 0 && ranges[0].start > 1)
        {
            // First range starts after 1, write object 0 separately
            WriteLine("0 1");
            WriteLine("0000000000 65535 f ");
        }

        // Write remaining ranges
        foreach (var range in ranges)
        {
            WriteLine($"{range.start} {range.count}");
            foreach (var offset in range.offsets)
            {
                WriteLine($"{offset:D10} 00000 n ");
            }
        }
    }

    private void WriteTrailer()
    {
        WriteLine("trailer");
        WriteLine("<<");
        // Size should be the highest object number + 1
        int maxObjectNumber = _offsets.Count > 0 ? _offsets.Keys.Max() : 0;
        WriteLine($"/Size {maxObjectNumber + 1}");
        WriteLine($"/Root {_catalog!.Number} 0 R");
        WriteLine(">>");
        WriteLine("startxref");
        WriteLine(_xrefOffset.ToString());
        WriteLine("%%EOF");
    }

    private void WriteLine(string text)
    {
        var bytes = Encoding.UTF8.GetBytes(text + "\n");
        _stream.Write(bytes, 0, bytes.Length);
    }

    public void Dispose()
    {
        _stream.Dispose();
    }
}
