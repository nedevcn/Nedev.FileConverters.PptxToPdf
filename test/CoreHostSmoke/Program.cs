using System.IO.Compression;
using System.Text;
using Nedev.FileConverters;
using Nedev.FileConverters.PptxToPdf;

var tempRoot = Path.Combine(Path.GetTempPath(), "Nedev.FileConverters.PptxToPdf", "CoreHostSmoke", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(tempRoot);

try
{
    var pptxPath = Path.Combine(tempRoot, "blank-smoke.pptx");
    var directPdfPath = Path.Combine(tempRoot, "direct.pdf");
    CreateMinimalBlankPptx(pptxPath);

    var directConverter = new PptxToPdfConverter();
    directConverter.Convert(pptxPath, directPdfPath);
    var directPdfBytes = File.ReadAllBytes(directPdfPath);
    AssertPdfBytes(directPdfBytes, "Direct API should produce a valid PDF.");

    using var adapterInput = File.OpenRead(pptxPath);
    using var adapterOutput = new PptxToPdfFileConverter().Convert(adapterInput);
    var adapterPdfBytes = ReadAllBytes(adapterOutput);
    AssertPdfBytes(adapterPdfBytes, "Adapter path should produce a valid PDF.");

    PptxToPdfCoreRegistration.EnsureRegistered();
    PptxToPdfCoreRegistration.EnsureRegistered();

    using var coreInput1 = File.OpenRead(pptxPath);
    using var coreOutput1 = Converter.Convert(coreInput1, "pptx", "pdf");
    var corePdfBytes1 = ReadAllBytes(coreOutput1);
    AssertPdfBytes(corePdfBytes1, "Core entry point should produce a valid PDF after registration.");

    using var coreInput2 = File.OpenRead(pptxPath);
    using var coreOutput2 = Converter.Convert(coreInput2, "pptx", "pdf");
    var corePdfBytes2 = ReadAllBytes(coreOutput2);
    AssertPdfBytes(corePdfBytes2, "Repeated Core entry conversion should remain stable.");

    Assert(directPdfBytes.SequenceEqual(adapterPdfBytes), "Direct API and adapter path should emit identical PDF bytes for the blank smoke sample.");
    Assert(directPdfBytes.SequenceEqual(corePdfBytes1), "Direct API and Core entry point should emit identical PDF bytes for the blank smoke sample.");
    Assert(corePdfBytes1.SequenceEqual(corePdfBytes2), "Repeated Core entry conversions should be byte-stable for the blank smoke sample.");

    Console.WriteLine("Core host smoke passed.");
}
finally
{
    try
    {
        Directory.Delete(tempRoot, recursive: true);
    }
    catch
    {
        // Keep the temp output when cleanup fails so the smoke result remains inspectable.
    }
}

static void CreateMinimalBlankPptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddTextEntry(archive, "[Content_Types].xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
          <Default Extension="xml" ContentType="application/xml" />
          <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" />
          <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" />
        </Types>
        """);
    AddTextEntry(archive, "_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml" />
        </Relationships>
        """);
    AddTextEntry(archive, "ppt/presentation.xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldMasterIdLst />
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
          <p:notesSz cx="6858000" cy="9144000" />
        </p:presentation>
        """);
    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);
    AddTextEntry(archive, "ppt/slides/slide1.xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld>
            <p:spTree />
          </p:cSld>
        </p:sld>
        """);
}

static void AddTextEntry(ZipArchive archive, string path, string content)
{
    var entry = archive.CreateEntry(path);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    writer.Write(content);
}

static byte[] ReadAllBytes(Stream stream)
{
    if (stream is MemoryStream memoryStream)
        return memoryStream.ToArray();

    using var output = new MemoryStream();
    stream.Position = 0;
    stream.CopyTo(output);
    return output.ToArray();
}

static void AssertPdfBytes(byte[] pdfBytes, string message)
{
    Assert(pdfBytes.Length > 0, message);
    Assert(Encoding.ASCII.GetString(pdfBytes, 0, Math.Min(5, pdfBytes.Length)) == "%PDF-", message);
}

static void Assert(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}
