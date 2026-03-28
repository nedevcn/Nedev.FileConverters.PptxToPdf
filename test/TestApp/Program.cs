using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Nedev.FileConverters.PptxToPdf;
using Nedev.FileConverters.PptxToPdf.Image;
using Nedev.FileConverters.PptxToPdf.Pptx;

var tempRoot = Path.Combine(Path.GetTempPath(), "Nedev.FileConverters.PptxToPdf", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(tempRoot);

try
{
    AssertColor(
        Shape.ParseColor(XElement.Parse("""<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:schemeClr val="bg1" /></a:solidFill>""")),
        255, 255, 255, 255,
        "bg1 should map to Background1.");

    var transformedColor = Shape.ParseColor(XElement.Parse("""
        <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:srgbClr val="808080">
            <a:alpha val="50000" />
            <a:tint val="50000" />
          </a:srgbClr>
        </a:solidFill>
        """));
    Assert(transformedColor.HasValue && transformedColor.Value.A == 128, "Color alpha should honor OOXML percentage values.");
    Assert(transformedColor.HasValue && transformedColor.Value.R > 128, "Tint should lighten RGB channels.");

    var image1Png = CreateRgbaPng(2, 1, Px(255, 0, 0, 255), Px(0, 255, 0, 255));
    var image2Png = CreateRgbaPng(2, 1, Px(0, 0, 255, 255), Px(255, 255, 0, 96));
    var image3Png = CreateRgbaPng(1, 1, Px(255, 0, 255, 255));

    var preparedOpaquePng = ImageConverter.PrepareForPdf(image1Png);
    Assert(!preparedOpaquePng.IsJpeg, "PNG pictures should stay lossless when prepared for PDF.");
    Assert(preparedOpaquePng.AlphaMaskData == null, "Fully opaque PNGs should not allocate a soft mask.");
    Assert(InflateZlib(preparedOpaquePng.Data).SequenceEqual(new byte[] { 255, 0, 0, 0, 255, 0 }), "Opaque PNG RGB bytes should be preserved.");

    var preparedTransparentPng = ImageConverter.PrepareForPdf(image2Png);
    Assert(!preparedTransparentPng.IsJpeg, "Transparent PNGs should stay lossless when prepared for PDF.");
    Assert(preparedTransparentPng.AlphaMaskData != null, "Transparent PNGs should emit a separate alpha mask.");
    Assert(InflateZlib(preparedTransparentPng.Data).SequenceEqual(new byte[] { 0, 0, 255, 255, 255, 0 }), "Transparent PNG RGB bytes should be preserved.");
    Assert(InflateZlib(preparedTransparentPng.AlphaMaskData!).SequenceEqual(new byte[] { 255, 96 }), "Transparent PNG alpha bytes should be preserved.");

    var image4Bmp = CreateBmp32(2, 1, Px(10, 20, 30, 255), Px(200, 210, 220, 64));
    var preparedBmp = ImageConverter.PrepareForPdf(image4Bmp);
    Assert(!preparedBmp.IsJpeg, "32-bit BMPs should use the lossless PDF image path.");
    Assert(preparedBmp.AlphaMaskData != null, "32-bit BMPs should emit an alpha mask when transparency is present.");
    Assert(InflateZlib(preparedBmp.Data).SequenceEqual(new byte[] { 10, 20, 30, 200, 210, 220 }), "BMP RGB bytes should be preserved.");
    Assert(InflateZlib(preparedBmp.AlphaMaskData!).SequenceEqual(new byte[] { 255, 64 }), "BMP alpha bytes should be preserved.");

    var jpegBytes = ImageConverter.ConvertToJpeg(image1Png);
    var preparedJpeg = ImageConverter.PrepareForPdf(jpegBytes);
    Assert(preparedJpeg.IsJpeg, "JPEG inputs should keep the passthrough path.");
    Assert(preparedJpeg.AlphaMaskData == null, "JPEG passthrough should not allocate a soft mask.");
    Assert(preparedJpeg.Data.SequenceEqual(jpegBytes), "JPEG passthrough should keep the original bytes.");

    AssertThrows<NotSupportedException>(() => ImageConverter.PrepareForPdf(CreateMinimalGif()), "GIF should fail explicitly instead of producing placeholder pixels.");
    AssertThrows<NotSupportedException>(() => ImageConverter.PrepareForPdf(CreateMinimalTiff()), "TIFF should fail explicitly instead of producing placeholder pixels.");
    AssertThrows<NotSupportedException>(() => ImageConverter.EnsureJpegFormat(CreateMinimalGif()), "EnsureJpegFormat should fail explicitly for unsupported GIF input.");

    var pptxPath = Path.Combine(tempRoot, "sample.pptx");
    var pdfPath = Path.Combine(tempRoot, "sample.pdf");

    CreateSamplePptx(pptxPath, image1Png, image2Png, image3Png);

    using (var input = File.OpenRead(pptxPath))
    using (var document = PptxDocument.Open(input))
    {
        Assert(document.Slides.Count == 2, "Expected two slides to be loaded.");
        Assert(document.Slides[0].Layout?.SourcePath == "ppt/slideLayouts/slideLayout1.xml", "Slide 1 should resolve layout1.");
        Assert(document.Slides[1].Layout?.SourcePath == "ppt/slideLayouts/slideLayout2.xml", "Slide 2 should resolve layout2.");
        Assert(document.Slides[0].Layout?.Master?.Theme?.SourcePath == "ppt/theme/theme1.xml", "Slide 1 should inherit theme1 from master1.");
        Assert(document.Slides[1].Layout?.Master?.Theme?.SourcePath == "ppt/theme/theme2.xml", "Slide 2 should inherit theme2 from master2.");
        Assert(document.GetImagePathFromRId(document.Slides[0].SourcePath, "rIdImg") == "ppt/media/image1.png", "Slide 1 image relationship should resolve to image1.");
        Assert(document.GetImagePathFromRId(document.Slides[1].SourcePath, "rIdImg2") == "ppt/media/image2.png", "Slide 2 image relationship should resolve to image2.");
        var themedBackground = document.Slides[0].GetEffectiveBackground()?.ResolveFill(document.Slides[0].GetEffectiveTheme(document.Theme));
        Assert(themedBackground?.Type == FillType.Solid, "Slide 1 should inherit a solid layout background from theme bgRef.");
        AssertColor(themedBackground?.Color, 0xDF, 0xDB, 0xC2, 0xFF, "Theme bgRef should resolve through the theme color scheme.");
        Assert(document.Slides[1].GetEffectiveBackground()?.ResolveFill(document.Slides[1].GetEffectiveTheme(document.Theme))?.Type == FillType.Gradient, "Slide 2 should inherit a gradient layout background.");
        Assert(document.Slides[0].Shapes.Count == 1, "Slide 1 should include the themed rectangle.");
        AssertColor(document.Slides[0].Shapes[0].Fill?.Color, 0xD9, 0x7A, 0x2B, 0xFF, "Layout clrMapOvr should remap accent1 to accent2 before theme resolution.");
        Assert(document.Slides[1].Shapes.Count == 2, "Slide 2 should include the themed rectangle plus the picture-fill shape.");
        AssertColor(document.Slides[1].Shapes[0].Fill?.Color, 0x6A, 0x8F, 0x63, 0xFF, "Slide clrMapOvr should win over layout/master mapping and resolve accent1 as accent3.");
        var slide1RenderableShapes = document.Slides[0].GetRenderableShapes().ToList();
        Assert(slide1RenderableShapes.Count == 2, "Slide 1 should include the inherited master shape plus its own shape.");
        Assert(slide1RenderableShapes[0].Fill?.Color.A == 128, "Inherited master shape should preserve alpha from OOXML color transforms.");
        Assert(document.Slides[1].Pictures.Count == 0, "Slide 2 should not declare direct pictures.");
        Assert(document.Slides[1].GetRenderablePictures().Count() == 1, "Slide 2 should inherit one layout picture.");
    }

    var converter = new PptxToPdfConverter();
    converter.Convert(pptxPath, pdfPath, useParallelProcessing: true);

    var pdfBytes = File.ReadAllBytes(pdfPath);
    var pdfText = Encoding.ASCII.GetString(pdfBytes);
    Assert(pdfBytes.Length > 0, "PDF output should not be empty.");
    Assert(CountOccurrences(pdfText, "/Subtype /Image") >= 4, "PDF should embed slide images plus the PNG soft mask.");
    Assert(pdfText.Contains("0 0 720.00 540.00 re f", StringComparison.Ordinal), "PDF should render inherited full-slide background fill.");
    Assert(pdfText.Contains("0.549 0.357 0.667 rg", StringComparison.Ordinal), "PDF should render inherited master shapes.");
    Assert(pdfText.Contains("/ca 0.5020", StringComparison.Ordinal), "PDF should emit fill alpha extgstates for semi-transparent fills.");
    Assert(pdfText.Contains("/Filter /FlateDecode", StringComparison.Ordinal), "PNG images should be embedded as Flate streams.");
    Assert(pdfText.Contains("/SMask", StringComparison.Ordinal), "Transparent PNGs should emit a soft mask.");
    Assert(!pdfText.Contains("/DCTDecode", StringComparison.Ordinal), "PNG-only smoke sample should not fall back to JPEG.");

    using var pptxStream = File.OpenRead(pptxPath);
    using var outputStream = new MemoryStream();
    converter.Convert(pptxStream, outputStream);
    Assert(outputStream.CanRead, "Output stream should remain open after conversion.");
    Assert(outputStream.Length > 0, "Stream-based conversion should write PDF bytes.");

    var coreFileConverter = new PptxToPdfFileConverter();

    using var coreInputStream = File.OpenRead(pptxPath);
    using var coreConvertedStream = coreFileConverter.Convert(coreInputStream);
    Assert(coreConvertedStream.CanRead, "Core-compatible file converter should return a readable output stream.");
    Assert(coreConvertedStream.Length > 0, "Core-compatible file converter should write PDF bytes.");

    using var coreAsyncInputStream = File.OpenRead(pptxPath);
    using var coreConvertedAsyncStream = await coreFileConverter.ConvertAsync(coreAsyncInputStream);
    Assert(coreConvertedAsyncStream.CanRead, "Core-compatible async file converter should return a readable output stream.");
    Assert(coreConvertedAsyncStream.Length > 0, "Core-compatible async file converter should write PDF bytes.");

    PptxToPdfCoreRegistration.EnsureRegistered();
    using var coreRegistryInputStream = File.OpenRead(pptxPath);
    using var coreRegistryConvertedStream = Nedev.FileConverters.Converter.Convert(coreRegistryInputStream, "pptx", "pdf");
    Assert(coreRegistryConvertedStream.CanRead, "Core registration helper should enable the Core Converter entry point.");
    Assert(coreRegistryConvertedStream.Length > 0, "Core Converter entry point should write PDF bytes after registration.");

    var bmpPptxPath = Path.Combine(tempRoot, "bmp-sample.pptx");
    var bmpPdfPath = Path.Combine(tempRoot, "bmp-sample.pdf");
    CreateSingleImagePptx(bmpPptxPath, "ppt/media/image1.bmp", image4Bmp);
    converter.Convert(bmpPptxPath, bmpPdfPath);
    var bmpPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(bmpPdfPath));
    Assert(CountOccurrences(bmpPdfText, "/Subtype /Image") >= 2, "BMP sample should embed the image plus its alpha mask.");
    Assert(bmpPdfText.Contains("/Filter /FlateDecode", StringComparison.Ordinal), "BMP sample should use the Flate image path.");
    Assert(bmpPdfText.Contains("/SMask", StringComparison.Ordinal), "BMP sample should emit a soft mask.");

    var gifPptxPath = Path.Combine(tempRoot, "gif-skip-sample.pptx");
    var gifPdfPath = Path.Combine(tempRoot, "gif-skip-sample.pdf");
    CreateSingleImagePptx(gifPptxPath, "ppt/media/image1.gif", CreateMinimalGif());
    converter.Convert(gifPptxPath, gifPdfPath);
    var gifPdfBytes = File.ReadAllBytes(gifPdfPath);
    var gifPdfText = Encoding.ASCII.GetString(gifPdfBytes);
    Assert(gifPdfBytes.Length > 0, "Unsupported-image sample should still produce a PDF.");
    Assert(CountOccurrences(gifPdfText, "/Subtype /Image") == 0, "Unsupported GIF input should be skipped instead of embedding placeholder image data.");

    var placeholderPptxPath = Path.Combine(tempRoot, "placeholder-sample.pptx");
    var placeholderPdfPath = Path.Combine(tempRoot, "placeholder-sample.pdf");
    CreatePlaceholderInheritancePptx(placeholderPptxPath);

    using (var placeholderInput = File.OpenRead(placeholderPptxPath))
    using (var placeholderDocument = PptxDocument.Open(placeholderInput))
    {
        Assert(placeholderDocument.Slides.Count == 1, "Placeholder sample should contain one slide.");
        var renderableShapes = placeholderDocument.Slides[0].GetRenderableShapes().ToList();
        Assert(renderableShapes.Count == 2, "Placeholder sample should resolve title/body placeholders into renderable slide shapes.");

        var titleShape = renderableShapes.First(shape => shape.PlaceholderType == PlaceholderType.Title);
        Assert(titleShape.Bounds.X == 914400 && titleShape.Bounds.Width == 7315200, "Title placeholder should inherit layout bounds.");
        Assert(titleShape.Paragraphs[0].Alignment == TextAlignment.Center, "Title placeholder should inherit centered paragraph alignment from master text styles.");
        Assert(titleShape.Paragraphs[0].Runs[0].Properties?.FontSize == 26, "Title placeholder should inherit master default font size.");
        AssertColor(titleShape.Paragraphs[0].Runs[0].Properties?.Color, 0xB9, 0x44, 0x41, 0xFF, "Title placeholder should inherit master default run color.");

        var bodyShape = renderableShapes.First(shape => shape.PlaceholderType == PlaceholderType.Body);
        Assert(bodyShape.Bounds.Y == 1828800 && bodyShape.Bounds.Height == 2743200, "Body placeholder should inherit layout bounds.");
        Assert(bodyShape.Paragraphs[0].MarginLeft == 457200, "Body placeholder should inherit paragraph left margin from master body style.");
        Assert(bodyShape.Paragraphs[0].Runs[0].Properties?.FontSize == 18, "Body placeholder should inherit body text default font size.");
        AssertColor(bodyShape.Paragraphs[0].Runs[0].Properties?.Color, 0x35, 0x5C, 0x7D, 0xFF, "Body placeholder should inherit body text default run color.");
    }

    converter.Convert(placeholderPptxPath, placeholderPdfPath);
    var placeholderPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(placeholderPdfPath));
    Assert(placeholderPdfText.Contains("Inherited Title", StringComparison.Ordinal), "Placeholder sample PDF should include title text.");
    Assert(placeholderPdfText.Contains("Inherited body copy", StringComparison.Ordinal), "Placeholder sample PDF should include body text.");
    Assert(placeholderPdfText.Contains("Second line", StringComparison.Ordinal), "Placeholder sample PDF should preserve explicit newlines as separate rendered lines.");
    Assert(!placeholderPdfText.Contains("Layout body prompt", StringComparison.Ordinal), "Layout placeholder prompt text should not render once a slide placeholder overrides it.");
    Assert(placeholderPdfText.Contains("26 Tf", StringComparison.Ordinal), "Placeholder sample PDF should use inherited title font size.");
    Assert(placeholderPdfText.Contains("18 Tf", StringComparison.Ordinal), "Placeholder sample PDF should use inherited body font size.");
    Assert(placeholderPdfText.Contains("0.725 0.267 0.255 rg", StringComparison.Ordinal), "Placeholder sample PDF should use inherited title text color.");
    Assert(placeholderPdfText.Contains("0.208 0.361 0.490 rg", StringComparison.Ordinal), "Placeholder sample PDF should use inherited body text color.");
    Assert(CountOccurrences(placeholderPdfText, " Tm") == 3, "Placeholder sample should render exactly three text matrices: one title line plus two body lines.");
    Assert(placeholderPdfText.Contains("469.20 Tm", StringComparison.Ordinal), "Placeholder title text should start near the top of its text box.");
    Assert(placeholderPdfText.Contains("1 0 0 1 115.20 370.80 Tm", StringComparison.Ordinal), "Placeholder body text should honor inherited paragraph margin and top-down layout.");
    Assert(placeholderPdfText.Contains("1 0 0 1 115.20 349.20 Tm", StringComparison.Ordinal), "Placeholder body second line should continue downward with stable line spacing.");

    var bulletPptxPath = Path.Combine(tempRoot, "bullet-sample.pptx");
    var bulletPdfPath = Path.Combine(tempRoot, "bullet-sample.pdf");
    CreateBulletInheritancePptx(bulletPptxPath);

    using (var bulletInput = File.OpenRead(bulletPptxPath))
    using (var bulletDocument = PptxDocument.Open(bulletInput))
    {
        Assert(bulletDocument.Slides.Count == 1, "Bullet sample should contain one slide.");
        var bulletShapes = bulletDocument.Slides[0].GetRenderableShapes().ToList();
        Assert(bulletShapes.Count == 1, "Bullet sample should resolve one body placeholder shape.");

        var bulletShape = bulletShapes[0];
        Assert(bulletShape.Paragraphs.Count == 4, "Bullet sample should contain four paragraphs.");
        Assert(bulletShape.Paragraphs[0].BulletType == BulletType.Char, "First paragraph should inherit a character bullet from master body style.");
        Assert(bulletShape.Paragraphs[0].BulletChar == "*", "Inherited bullet character should come from master body style.");
        Assert(bulletShape.Paragraphs[0].MarginLeft == 457200, "Inherited bullet paragraph should keep master left margin.");
        Assert(bulletShape.Paragraphs[0].Indent == -228600, "Inherited bullet paragraph should keep master hanging indent.");
        Assert(bulletShape.Paragraphs[1].BulletType == BulletType.AutoNumber, "Second paragraph should keep its explicit auto-number bullet.");
        Assert(bulletShape.Paragraphs[1].BulletStartAt == 3, "Explicit auto-number start should be preserved.");
        Assert(bulletShape.Paragraphs[3].BulletType == BulletType.None, "Fourth paragraph should preserve buNone instead of inheriting the master bullet.");
        Assert(bulletShape.Paragraphs[3].HasExplicitBulletDefinition, "Fourth paragraph should mark buNone as an explicit bullet override.");
    }

    converter.Convert(bulletPptxPath, bulletPdfPath);
    var bulletPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(bulletPdfPath));
    Assert(bulletPdfText.Contains("Inherited bullet", StringComparison.Ordinal), "Bullet sample PDF should include inherited bullet text.");
    Assert(bulletPdfText.Contains("Numbered item three", StringComparison.Ordinal), "Bullet sample PDF should include explicit auto-number item text.");
    Assert(bulletPdfText.Contains("Plain paragraph", StringComparison.Ordinal), "Bullet sample PDF should include buNone paragraph text.");
    Assert(CountOccurrences(bulletPdfText, "(*) Tj") == 1, "Bullet sample PDF should render exactly one inherited character bullet marker.");
    Assert(bulletPdfText.Contains("(3.) Tj", StringComparison.Ordinal), "Bullet sample PDF should render explicit auto-number start.");
    Assert(bulletPdfText.Contains("(4.) Tj", StringComparison.Ordinal), "Bullet sample PDF should continue auto-numbering across following paragraphs.");
    Assert(!bulletPdfText.Contains("(1.) Tj", StringComparison.Ordinal), "Bullet sample PDF should not invent numbering for the inherited character bullet paragraph.");

    var autofitPptxPath = Path.Combine(tempRoot, "autofit-sample.pptx");
    var autofitPdfPath = Path.Combine(tempRoot, "autofit-sample.pdf");
    CreateAutofitPptx(autofitPptxPath);

    using (var autofitInput = File.OpenRead(autofitPptxPath))
    using (var autofitDocument = PptxDocument.Open(autofitInput))
    {
        Assert(autofitDocument.Slides.Count == 1, "Autofit sample should contain one slide.");
        Assert(autofitDocument.Slides[0].Shapes.Count == 1, "Autofit sample should contain one text shape.");

        var autofitShape = autofitDocument.Slides[0].Shapes[0];
        Assert(autofitShape.TextProperties?.AutoFit == TextAutoFit.Normal, "Autofit sample should parse normAutofit from bodyPr.");
        Assert(Math.Abs((autofitShape.TextProperties?.FontScale ?? 0) - 0.65) < 0.0001, "Autofit sample should normalize fontScale as a fractional percentage.");
        Assert(Math.Abs((autofitShape.TextProperties?.LineSpaceReduction ?? 0) - 0.20) < 0.0001, "Autofit sample should normalize line space reduction as a fractional percentage.");
        Assert(autofitShape.Paragraphs.Count == 2, "Autofit sample should preserve two text paragraphs.");
        Assert(autofitShape.Paragraphs.All(paragraph => paragraph.Runs[0].Properties?.FontSize == 20), "Autofit sample paragraphs should keep the original run font size before render-time scaling.");
    }

    converter.Convert(autofitPptxPath, autofitPdfPath);
    var autofitPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(autofitPdfPath));
    Assert(CountOccurrences(autofitPdfText, "13 Tf") == 2, "Autofit sample PDF should scale 20pt text down to 13pt using fontScale.");
    Assert(autofitPdfText.Contains("1 0 0 1 79.20 451.92 Tm", StringComparison.Ordinal), "Autofit sample first line should use the reduced line height.");
    Assert(autofitPdfText.Contains("1 0 0 1 79.20 439.44 Tm", StringComparison.Ordinal), "Autofit sample second line should reflect line space reduction in its Y position.");

    var autofitIndentPptxPath = Path.Combine(tempRoot, "autofit-indent-sample.pptx");
    var autofitIndentPdfPath = Path.Combine(tempRoot, "autofit-indent-sample.pdf");
    CreateAutofitIndentPptx(autofitIndentPptxPath);

    using (var autofitIndentInput = File.OpenRead(autofitIndentPptxPath))
    using (var autofitIndentDocument = PptxDocument.Open(autofitIndentInput))
    {
        Assert(autofitIndentDocument.Slides.Count == 1, "Autofit indent sample should contain one slide.");
        Assert(autofitIndentDocument.Slides[0].Shapes.Count == 1, "Autofit indent sample should contain one text shape.");

        var autofitIndentParagraph = autofitIndentDocument.Slides[0].Shapes[0].Paragraphs[0];
        Assert(autofitIndentParagraph.MarginLeft == 457200, "Autofit indent sample should preserve paragraph left margin.");
        Assert(autofitIndentParagraph.Indent == 228600, "Autofit indent sample should preserve positive first-line indent.");
    }

    converter.Convert(autofitIndentPptxPath, autofitIndentPdfPath);
    var autofitIndentPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(autofitIndentPdfPath));
    Assert(autofitIndentPdfText.Contains("1 0 0 1 133.20", StringComparison.Ordinal), "Autofit indent sample first visual line should include the first-line indent.");
    Assert(autofitIndentPdfText.Contains("1 0 0 1 115.20", StringComparison.Ordinal), "Autofit indent sample continuation line should return to the paragraph margin.");

    var autofitBulletPptxPath = Path.Combine(tempRoot, "autofit-bullet-sample.pptx");
    var autofitBulletPdfPath = Path.Combine(tempRoot, "autofit-bullet-sample.pdf");
    CreateAutofitBulletWrapPptx(autofitBulletPptxPath);

    using (var autofitBulletInput = File.OpenRead(autofitBulletPptxPath))
    using (var autofitBulletDocument = PptxDocument.Open(autofitBulletInput))
    {
        Assert(autofitBulletDocument.Slides.Count == 1, "Autofit bullet sample should contain one slide.");
        Assert(autofitBulletDocument.Slides[0].Shapes.Count == 1, "Autofit bullet sample should contain one text shape.");

        var autofitBulletParagraph = autofitBulletDocument.Slides[0].Shapes[0].Paragraphs[0];
        Assert(autofitBulletParagraph.BulletType == BulletType.Char, "Autofit bullet sample should preserve the explicit character bullet.");
        Assert(autofitBulletParagraph.MarginLeft == 457200, "Autofit bullet sample should preserve the paragraph margin.");
        Assert(autofitBulletParagraph.Indent == -228600, "Autofit bullet sample should preserve the hanging indent.");
    }

    converter.Convert(autofitBulletPptxPath, autofitBulletPdfPath);
    var autofitBulletPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(autofitBulletPdfPath));
    Assert(CountOccurrences(autofitBulletPdfText, "(*) Tj") == 1, "Autofit bullet sample should render exactly one bullet marker.");
    Assert(autofitBulletPdfText.Contains("1 0 0 1 97.20", StringComparison.Ordinal), "Autofit bullet marker should stay in the hanging bullet slot.");
    Assert(CountOccurrences(autofitBulletPdfText, "1 0 0 1 115.20") >= 2, "Autofit bullet sample should align wrapped text lines to the paragraph margin.");

    var tablePptxPath = Path.Combine(tempRoot, "table-text-sample.pptx");
    var tablePdfPath = Path.Combine(tempRoot, "table-text-sample.pdf");
    CreateTableTextPptx(tablePptxPath);

    using (var tableInput = File.OpenRead(tablePptxPath))
    using (var tableDocument = PptxDocument.Open(tableInput))
    {
        Assert(tableDocument.Slides.Count == 1, "Table text sample should contain one slide.");
        Assert(tableDocument.Slides[0].Tables.Count == 1, "Table text sample should resolve one table from the graphic frame.");

        var table = tableDocument.Slides[0].Tables[0];
        Assert(table.Rows.Count == 1, "Table text sample should contain one row.");
        Assert(table.Rows[0].Cells.Count == 1, "Table text sample should contain one cell.");

        var cell = table.Rows[0].Cells[0];
        Assert(cell.Paragraphs.Count == 1, "Table text sample should preserve the cell paragraph.");
        Assert(cell.Paragraphs[0].Runs.Count == 2, "Table text sample should preserve multiple runs inside the cell paragraph.");
        Assert(cell.Paragraphs[0].BulletType == BulletType.Char, "Table text sample should preserve the cell bullet definition.");
        AssertColor(cell.Paragraphs[0].Runs[0].Properties?.Color, 0xCC, 0x33, 0x00, 0xFF, "Table text sample should preserve the first run color.");
        AssertColor(cell.Paragraphs[0].Runs[1].Properties?.Color, 0x00, 0x66, 0xCC, 0xFF, "Table text sample should preserve the second run color.");
    }

    converter.Convert(tablePptxPath, tablePdfPath);
    var tablePdfText = Encoding.ASCII.GetString(File.ReadAllBytes(tablePdfPath));
    Assert(CountOccurrences(tablePdfText, "(*) Tj") == 1, "Table text sample should render exactly one bullet marker.");
    Assert(tablePdfText.Contains("Cell", StringComparison.Ordinal), "Table text sample PDF should include the first text run.");
    Assert(tablePdfText.Contains("detail", StringComparison.Ordinal), "Table text sample PDF should include the second text run.");
    Assert(tablePdfText.Contains("0.800 0.200 0.000 rg", StringComparison.Ordinal), "Table text sample PDF should use the first run color.");
    Assert(tablePdfText.Contains("0.000 0.400 0.800 rg", StringComparison.Ordinal), "Table text sample PDF should use the second run color.");
    Assert(CountOccurrences(tablePdfText, "1 0 0 1 115.20") >= 2, "Table text sample should align wrapped cell text to the paragraph margin.");

    var chartFontPptxPath = Path.Combine(tempRoot, "chart-font-resource-sample.pptx");
    var chartFontPdfPath = Path.Combine(tempRoot, "chart-font-resource-sample.pdf");
    CreateChartPptx(chartFontPptxPath, CreateBarChartXml(chartTitle: "Font Registration Chart"));

    converter.Convert(chartFontPptxPath, chartFontPdfPath);
    var chartFontPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(chartFontPdfPath));
    Assert(chartFontPdfText.Contains("Font Registration Chart", StringComparison.Ordinal), "Chart font-resource sample PDF should include the chart title.");
    AssertPdfUsesRegisteredFontResources(chartFontPdfText, "Chart font-resource sample PDF should declare registered font resources for chart text.");

    var smartArtFontPptxPath = Path.Combine(tempRoot, "smartart-font-resource-sample.pptx");
    var smartArtFontPdfPath = Path.Combine(tempRoot, "smartart-font-resource-sample.pdf");
    CreateSmartArtPptx(
        smartArtFontPptxPath,
        CreateSmartArtDataModelXml("Font registration node"),
        CreateSmartArtLayoutXml());

    converter.Convert(smartArtFontPptxPath, smartArtFontPdfPath);
    var smartArtFontPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(smartArtFontPdfPath));
    Assert(smartArtFontPdfText.Contains("Vertical Bullet List", StringComparison.Ordinal), "SmartArt font-resource sample PDF should use the layout display name as the title.");
    Assert(smartArtFontPdfText.Contains("Font registration node", StringComparison.Ordinal), "SmartArt font-resource sample PDF should include node text.");
    AssertPdfUsesRegisteredFontResources(smartArtFontPdfText, "SmartArt font-resource sample PDF should declare registered font resources for SmartArt text.");

    var tableLayoutPptxPath = Path.Combine(tempRoot, "table-layout-sample.pptx");
    var tableLayoutPdfPath = Path.Combine(tempRoot, "table-layout-sample.pdf");
    CreateTableLayoutPptx(tableLayoutPptxPath);

    using (var tableLayoutInput = File.OpenRead(tableLayoutPptxPath))
    using (var tableLayoutDocument = PptxDocument.Open(tableLayoutInput))
    {
        Assert(tableLayoutDocument.Slides.Count == 1, "Table layout sample should contain one slide.");
        Assert(tableLayoutDocument.Slides[0].Tables.Count == 2, "Table layout sample should resolve two tables.");

        var mergedTable = tableLayoutDocument.Slides[0].Tables[0];
        Assert(mergedTable.Rows[0].Cells[0].ColumnSpan == 2, "Merged table sample should preserve the leading cell grid span.");
        Assert(mergedTable.Rows[0].Cells[1].HorizontalMerge, "Merged table sample should preserve the trailing hMerge cell.");

        var centeredTable = tableLayoutDocument.Slides[0].Tables[1];
        Assert(centeredTable.Rows[0].Cells[0].Properties?.NoWrap == true, "Centered table sample should preserve cell noWrap.");
        Assert(centeredTable.Rows[0].Cells[0].Properties?.Anchor == TextAnchor.Middle, "Centered table sample should preserve vertical center anchor.");
    }

    converter.Convert(tableLayoutPptxPath, tableLayoutPdfPath);
    var tableLayoutPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(tableLayoutPdfPath));
    Assert(CountOccurrences(tableLayoutPdfText, "Merged text should stay on one line") == 1, "Merged table sample PDF should render the merged-cell text exactly once.");
    Assert(CountOccurrences(tableLayoutPdfText, " Tm") >= 2, "Table layout sample PDF should emit separate text matrices for the merged and centered cell samples.");
    Assert(tableLayoutPdfText.Contains("1 0 0 1 79.20 424.80 Tm", StringComparison.Ordinal), "Centered table sample should vertically center the noWrap line inside the cell.");
    Assert(tableLayoutPdfText.Contains("Centered no-wrap text should remain on a single line", StringComparison.Ordinal), "Centered table sample PDF should keep the noWrap text on one line.");

    var chartPptxPath = Path.Combine(tempRoot, "chart-sample.pptx");
    var chartPdfPath = Path.Combine(tempRoot, "chart-sample.pdf");
    CreateChartPptx(chartPptxPath, CreateBarChartXml());

    using (var chartInput = File.OpenRead(chartPptxPath))
    using (var chartDocument = PptxDocument.Open(chartInput))
    {
        Assert(chartDocument.Slides.Count == 1, "Chart sample should contain one slide.");
        Assert(chartDocument.Slides[0].Charts.Count == 1, "Chart sample should resolve one chart from slide relationships.");

        var chart = chartDocument.Slides[0].Charts[0];
        Assert(chart.Title == "Revenue by Quarter", "Chart title should be loaded from the chart part.");
        Assert(chart.Type == ChartType.Bar, "Bar chart sample should parse as a bar chart.");
        Assert(chart.Series.Count == 1, "Bar chart sample should include one series.");
        Assert(chart.Series[0].Name == "North Region", "Series name should be loaded from chart XML.");
        AssertColor(chart.Series[0].Fill?.Color, 0x33, 0x66, 0xCC, 0xFF, "Bar chart sample should preserve series fill color.");
        Assert(chart.Series[0].DataPoints.Count == 3, "Bar chart sample should include three data points.");
        Assert(chart.Series[0].DataPoints[0].Category == "Q1", "First chart category should be read from cached chart data.");
        Assert(chart.Series[0].DataPoints[1].Value == 18, "Second chart value should be read from cached chart data.");
    }

    converter.Convert(chartPptxPath, chartPdfPath);
    var chartPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(chartPdfPath));
    Assert(chartPdfText.Contains("Revenue by Quarter", StringComparison.Ordinal), "Chart sample PDF should include the chart title.");
    Assert(chartPdfText.Contains("Q1", StringComparison.Ordinal), "Chart sample PDF should include category labels from chart data.");
    Assert(chartPdfText.Contains("Q3", StringComparison.Ordinal), "Chart sample PDF should include all rendered bar labels.");
    Assert(chartPdfText.Contains("0.200 0.400 0.800 rg", StringComparison.Ordinal), "Bar chart sample PDF should use the series fill color.");
    AssertPdfUsesRegisteredFontResources(chartPdfText, "Chart sample PDF should declare registered font resources for chart text.");
    Assert(!chartPdfText.Contains("Unknown Chart", StringComparison.Ordinal), "Resolved chart sample should not fall back to the unknown chart placeholder.");

    var columnChartPptxPath = Path.Combine(tempRoot, "column-chart-sample.pptx");
    var columnChartPdfPath = Path.Combine(tempRoot, "column-chart-sample.pdf");
    CreateChartPptx(columnChartPptxPath, CreateBarChartXml(
        chartTitle: "Pipeline by Stage",
        barDir: "col",
        seriesName: "South Region",
        seriesColor: "993366",
        includeLegend: true));

    using (var columnChartInput = File.OpenRead(columnChartPptxPath))
    using (var columnChartDocument = PptxDocument.Open(columnChartInput))
    {
        Assert(columnChartDocument.Slides.Count == 1, "Column chart sample should contain one slide.");
        Assert(columnChartDocument.Slides[0].Charts.Count == 1, "Column chart sample should resolve one chart from slide relationships.");

        var columnChart = columnChartDocument.Slides[0].Charts[0];
        Assert(columnChart.Title == "Pipeline by Stage", "Column chart title should be loaded from the chart part.");
        Assert(columnChart.Type == ChartType.Column, "barDir=col should parse as a column chart.");
        Assert(columnChart.Legend != null, "Column chart sample should parse legend metadata.");
        Assert(columnChart.Series.Count == 1, "Column chart sample should include one series.");
        Assert(columnChart.Series[0].Name == "South Region", "Column chart series name should be loaded from chart XML.");
        AssertColor(columnChart.Series[0].Fill?.Color, 0x99, 0x33, 0x66, 0xFF, "Column chart sample should preserve series fill color.");
    }

    converter.Convert(columnChartPptxPath, columnChartPdfPath);
    var columnChartPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(columnChartPdfPath));
    Assert(columnChartPdfText.Contains("Pipeline by Stage", StringComparison.Ordinal), "Column chart sample PDF should include the chart title.");
    Assert(columnChartPdfText.Contains("South Region", StringComparison.Ordinal), "Column chart sample PDF should include legend text.");
    Assert(columnChartPdfText.Contains("0.600 0.200 0.400 rg", StringComparison.Ordinal), "Column chart sample PDF should use the series fill color.");
    Assert(!columnChartPdfText.Contains("Unknown Chart", StringComparison.Ordinal), "Column chart sample should not fall back to the unknown chart placeholder.");

    var areaChartPptxPath = Path.Combine(tempRoot, "area-chart-sample.pptx");
    var areaChartPdfPath = Path.Combine(tempRoot, "area-chart-sample.pdf");
    CreateChartPptx(areaChartPptxPath, CreateAreaChartXml());

    using (var areaChartInput = File.OpenRead(areaChartPptxPath))
    using (var areaChartDocument = PptxDocument.Open(areaChartInput))
    {
        Assert(areaChartDocument.Slides.Count == 1, "Area chart sample should contain one slide.");
        Assert(areaChartDocument.Slides[0].Charts.Count == 1, "Area chart sample should resolve one chart from slide relationships.");

        var areaChart = areaChartDocument.Slides[0].Charts[0];
        Assert(areaChart.Title == "Coverage Trend", "Area chart title should be loaded from the chart part.");
        Assert(areaChart.Type == ChartType.Area, "Area chart sample should parse as an area chart.");
        Assert(areaChart.Series.Count == 1, "Area chart sample should include one series.");
    }

    converter.Convert(areaChartPptxPath, areaChartPdfPath);
    var areaChartPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(areaChartPdfPath));
    Assert(areaChartPdfText.Contains("Coverage Trend", StringComparison.Ordinal), "Area chart sample PDF should still include the chart title.");
    Assert(areaChartPdfText.Contains("Area Chart", StringComparison.Ordinal), "Unsupported area chart sample should render the stable fallback placeholder.");
    Assert(!areaChartPdfText.Contains("Unknown Chart", StringComparison.Ordinal), "Unsupported area chart sample should preserve its parsed chart type in fallback output.");

    var smartArtPptxPath = Path.Combine(tempRoot, "smartart-sample.pptx");
    var smartArtPdfPath = Path.Combine(tempRoot, "smartart-sample.pdf");
    CreateSmartArtPptx(
        smartArtPptxPath,
        CreateSmartArtDataModelXml("Discover needs", "Deliver safely"),
        CreateSmartArtLayoutXml());

    using (var smartArtInput = File.OpenRead(smartArtPptxPath))
    using (var smartArtDocument = PptxDocument.Open(smartArtInput))
    {
        Assert(smartArtDocument.Slides.Count == 1, "SmartArt sample should contain one slide.");
        Assert(smartArtDocument.Slides[0].SmartArts.Count == 1, "SmartArt sample should resolve one SmartArt from slide relationships.");

        var smartArt = smartArtDocument.Slides[0].SmartArts[0];
        Assert(smartArt.Layout != null, "SmartArt sample should load its layout part.");
        Assert(smartArt.Type?.Contains("VerticalBulletList", StringComparison.Ordinal) == true, "SmartArt sample should preserve the layout uniqueId from the related layout part.");
        Assert(smartArt.ResolvedType == SmartArtType.VerticalBulletList, "SmartArt sample should resolve its renderer type from the layout metadata.");
        Assert(smartArt.DisplayName == "Vertical Bullet List", "SmartArt sample should surface the layout display name.");
        Assert(smartArt.Nodes.Count == 2, "SmartArt sample should load node data from the related data model part.");
        Assert(smartArt.Nodes[0].Text == "Discover needs", "SmartArt sample should load the first node text from the data model part.");
        Assert(smartArt.Nodes[1].Text == "Deliver safely", "SmartArt sample should load the second node text from the data model part.");
    }

    converter.Convert(smartArtPptxPath, smartArtPdfPath);
    var smartArtPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(smartArtPdfPath));
    Assert(smartArtPdfText.Contains("Vertical Bullet List", StringComparison.Ordinal), "SmartArt sample PDF should use the layout display name as the title.");
    Assert(smartArtPdfText.Contains("Discover needs", StringComparison.Ordinal), "SmartArt sample PDF should include the first node text.");
    Assert(smartArtPdfText.Contains("Deliver safely", StringComparison.Ordinal), "SmartArt sample PDF should include the second node text.");
    AssertPdfUsesRegisteredFontResources(smartArtPdfText, "SmartArt sample PDF should declare registered font resources for SmartArt text.");
    Assert(!smartArtPdfText.Contains("Unknown SmartArt", StringComparison.Ordinal), "SmartArt sample PDF should not fall back to the unknown SmartArt placeholder.");
    Assert(!smartArtPdfText.Contains("urn:microsoft.com/office/officeart/2005/8/layout/VerticalBulletList", StringComparison.Ordinal), "SmartArt sample PDF should not expose the raw layout uniqueId as the title.");

    var processSmartArtPptxPath = Path.Combine(tempRoot, "smartart-process-sample.pptx");
    var processSmartArtPdfPath = Path.Combine(tempRoot, "smartart-process-sample.pdf");
    CreateSmartArtPptx(
        processSmartArtPptxPath,
        CreateSmartArtTxBodyDataModelXml("Plan", "Build", "Launch"),
        CreateSmartArtLayoutXml(
            uniqueId: "urn:microsoft.com/office/officeart/2005/8/layout/BasicProcess",
            name: "Basic Process",
            description: "Process SmartArt smoke test"));

    using (var processSmartArtInput = File.OpenRead(processSmartArtPptxPath))
    using (var processSmartArtDocument = PptxDocument.Open(processSmartArtInput))
    {
        Assert(processSmartArtDocument.Slides.Count == 1, "Process SmartArt sample should contain one slide.");
        Assert(processSmartArtDocument.Slides[0].SmartArts.Count == 1, "Process SmartArt sample should resolve one SmartArt from slide relationships.");

        var processSmartArt = processSmartArtDocument.Slides[0].SmartArts[0];
        Assert(processSmartArt.ResolvedType == SmartArtType.BasicProcess, "Process SmartArt sample should resolve to the BasicProcess renderer branch.");
        Assert(processSmartArt.DisplayName == "Basic Process", "Process SmartArt sample should surface the layout display name.");
        Assert(processSmartArt.Nodes.Count == 3, "Process SmartArt sample should parse all nodes from the data model part.");
        Assert(processSmartArt.Nodes[0].Text == "Plan", "Process SmartArt sample should parse txBody text for the first node.");
        Assert(processSmartArt.Nodes[2].Text == "Launch", "Process SmartArt sample should parse txBody text for the last node.");
    }

    converter.Convert(processSmartArtPptxPath, processSmartArtPdfPath);
    var processSmartArtPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(processSmartArtPdfPath));
    Assert(processSmartArtPdfText.Contains("Basic Process", StringComparison.Ordinal), "Process SmartArt sample PDF should use the resolved display name as the title.");
    Assert(processSmartArtPdfText.Contains("Plan", StringComparison.Ordinal), "Process SmartArt sample PDF should include the first process node text.");
    Assert(processSmartArtPdfText.Contains("Launch", StringComparison.Ordinal), "Process SmartArt sample PDF should include the last process node text.");
    Assert(!processSmartArtPdfText.Contains("Unknown SmartArt", StringComparison.Ordinal), "Process SmartArt sample PDF should not fall back to the unknown placeholder.");

    var unknownSmartArtPptxPath = Path.Combine(tempRoot, "smartart-unknown-sample.pptx");
    var unknownSmartArtPdfPath = Path.Combine(tempRoot, "smartart-unknown-sample.pdf");
    CreateSmartArtPptx(
        unknownSmartArtPptxPath,
        CreateSmartArtDataModelXml("Sync backlog"),
        CreateSmartArtLayoutXml(
            uniqueId: "urn:example:layout/TeamSyncCanvas",
            name: "Team Sync Canvas",
            description: "Unknown SmartArt layout smoke test"));

    using (var unknownSmartArtInput = File.OpenRead(unknownSmartArtPptxPath))
    using (var unknownSmartArtDocument = PptxDocument.Open(unknownSmartArtInput))
    {
        Assert(unknownSmartArtDocument.Slides.Count == 1, "Unknown SmartArt sample should contain one slide.");
        Assert(unknownSmartArtDocument.Slides[0].SmartArts.Count == 1, "Unknown SmartArt sample should resolve one SmartArt from slide relationships.");

        var unknownSmartArt = unknownSmartArtDocument.Slides[0].SmartArts[0];
        Assert(unknownSmartArt.ResolvedType == SmartArtType.Unknown, "Unknown SmartArt sample should preserve an unknown renderer type when the layout is not recognized.");
        Assert(unknownSmartArt.DisplayName == "Team Sync Canvas", "Unknown SmartArt sample should still surface the layout display name for fallback rendering.");
    }

    converter.Convert(unknownSmartArtPptxPath, unknownSmartArtPdfPath);
    var unknownSmartArtPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(unknownSmartArtPdfPath));
    Assert(unknownSmartArtPdfText.Contains("Team Sync Canvas", StringComparison.Ordinal), "Unknown SmartArt sample PDF should use the layout display name in fallback output.");
    Assert(unknownSmartArtPdfText.Contains("Team Sync Canvas SmartArt", StringComparison.Ordinal), "Unknown SmartArt sample PDF should use a stable named fallback instead of a raw enum placeholder.");
    Assert(!unknownSmartArtPdfText.Contains("Unknown SmartArt", StringComparison.Ordinal), "Unknown SmartArt sample PDF should not fall back to the generic unknown placeholder label.");

    var richSmartArtPptxPath = Path.Combine(tempRoot, "smartart-rich-text-sample.pptx");
    var richSmartArtPdfPath = Path.Combine(tempRoot, "smartart-rich-text-sample.pdf");
    CreateSmartArtPptx(
        richSmartArtPptxPath,
        CreateSmartArtTxBodyDataModelXml("Scope (A)\nAlign \\ Ship", "Review & Sign-off"),
        CreateSmartArtLayoutXml(
            uniqueId: "urn:microsoft.com/office/officeart/2005/8/layout/BasicProcess",
            name: "Basic Process",
            description: "SmartArt rich text smoke test"));

    using (var richSmartArtInput = File.OpenRead(richSmartArtPptxPath))
    using (var richSmartArtDocument = PptxDocument.Open(richSmartArtInput))
    {
        Assert(richSmartArtDocument.Slides.Count == 1, "Rich SmartArt sample should contain one slide.");
        Assert(richSmartArtDocument.Slides[0].SmartArts.Count == 1, "Rich SmartArt sample should resolve one SmartArt from slide relationships.");

        var richSmartArt = richSmartArtDocument.Slides[0].SmartArts[0];
        Assert(richSmartArt.Nodes[0].Text == "Scope (A)\nAlign \\ Ship", "Rich SmartArt sample should preserve paragraph breaks and backslashes in parsed node text.");
        Assert(richSmartArt.Nodes[1].Text == "Review & Sign-off", "Rich SmartArt sample should decode XML-escaped node text.");
    }

    converter.Convert(richSmartArtPptxPath, richSmartArtPdfPath);
    var richSmartArtPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(richSmartArtPdfPath));
    Assert(richSmartArtPdfText.Contains("Scope \\(A\\)", StringComparison.Ordinal), "Rich SmartArt sample PDF should escape parentheses in SmartArt text.");
    Assert(richSmartArtPdfText.Contains(@"Align \\ Ship", StringComparison.Ordinal), "Rich SmartArt sample PDF should escape backslashes in SmartArt text.");
    Assert(richSmartArtPdfText.Contains("Review & Sign-off", StringComparison.Ordinal), "Rich SmartArt sample PDF should include XML-decoded SmartArt text.");

    var partialSmartArtPptxPath = Path.Combine(tempRoot, "smartart-partial-sample.pptx");
    var partialSmartArtPdfPath = Path.Combine(tempRoot, "smartart-partial-sample.pdf");
    CreatePartialSmartArtPptx(
        partialSmartArtPptxPath,
        CreateSmartArtDataModelXml("Orphaned node"));

    using (var partialSmartArtInput = File.OpenRead(partialSmartArtPptxPath))
    using (var partialSmartArtDocument = PptxDocument.Open(partialSmartArtInput))
    {
        Assert(partialSmartArtDocument.Slides.Count == 1, "Partial SmartArt sample should contain one slide.");
        Assert(partialSmartArtDocument.Slides[0].SmartArts.Count == 1, "Partial SmartArt sample should still materialize one SmartArt when only the data part exists.");

        var partialSmartArt = partialSmartArtDocument.Slides[0].SmartArts[0];
        Assert(partialSmartArt.ResolvedType == SmartArtType.Unknown, "Partial SmartArt sample should fall back to Unknown when the layout part is missing.");
        Assert(partialSmartArt.Nodes.Count == 1, "Partial SmartArt sample should still parse data nodes without a layout part.");
    }

    converter.Convert(partialSmartArtPptxPath, partialSmartArtPdfPath);
    var partialSmartArtPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(partialSmartArtPdfPath));
    Assert(partialSmartArtPdfText.Contains("SmartArt", StringComparison.Ordinal), "Partial SmartArt sample PDF should render the generic stable fallback label.");
    Assert(!partialSmartArtPdfText.Contains("Unknown SmartArt", StringComparison.Ordinal), "Partial SmartArt sample PDF should not regress to the old unknown placeholder label.");

    var inlineSmartArtPptxPath = Path.Combine(tempRoot, "smartart-inline-sample.pptx");
    var inlineSmartArtPdfPath = Path.Combine(tempRoot, "smartart-inline-sample.pdf");
    CreateInlineSmartArtPptx(
        inlineSmartArtPptxPath,
        CreateSmartArtDataModelXml("Inline discovery", "Inline delivery"),
        CreateSmartArtLayoutXml());

    using (var inlineSmartArtInput = File.OpenRead(inlineSmartArtPptxPath))
    using (var inlineSmartArtDocument = PptxDocument.Open(inlineSmartArtInput))
    {
        Assert(inlineSmartArtDocument.Slides.Count == 1, "Inline SmartArt sample should contain one slide.");
        Assert(inlineSmartArtDocument.Slides[0].SmartArts.Count == 1, "Inline SmartArt sample should still parse the embedded SmartArt path.");

        var inlineSmartArt = inlineSmartArtDocument.Slides[0].SmartArts[0];
        Assert(inlineSmartArt.ResolvedType == SmartArtType.VerticalBulletList, "Inline SmartArt sample should resolve the embedded layout type.");
        Assert(inlineSmartArt.Nodes[0].Text == "Inline discovery", "Inline SmartArt sample should preserve the first embedded node text.");
        Assert(inlineSmartArt.Nodes[1].Text == "Inline delivery", "Inline SmartArt sample should preserve the second embedded node text.");
    }

    converter.Convert(inlineSmartArtPptxPath, inlineSmartArtPdfPath);
    var inlineSmartArtPdfText = Encoding.ASCII.GetString(File.ReadAllBytes(inlineSmartArtPdfPath));
    Assert(inlineSmartArtPdfText.Contains("Vertical Bullet List", StringComparison.Ordinal), "Inline SmartArt sample PDF should still use the resolved display name.");
    Assert(inlineSmartArtPdfText.Contains("Inline discovery", StringComparison.Ordinal), "Inline SmartArt sample PDF should include the first embedded node text.");
    Assert(inlineSmartArtPdfText.Contains("Inline delivery", StringComparison.Ordinal), "Inline SmartArt sample PDF should include the second embedded node text.");

    Console.WriteLine("Smoke tests passed.");
}
finally
{
    try
    {
        Directory.Delete(tempRoot, recursive: true);
    }
    catch
    {
        // Best effort cleanup only.
    }
}

static void CreateSamplePptx(string path, byte[] image1Png, byte[] image2Png, byte[] image3Png)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldMasterIdLst />
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
            <p:sldId id="257" r:id="rId2" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/theme/theme1.xml", CreateThemeXml("SmokeTheme1", "F8F3D8", "2A5CAA"));
    AddTextEntry(archive, "ppt/theme/theme2.xml", CreateThemeXml("SmokeTheme2", "E7F5EC", "C95A52"));

    AddTextEntry(archive, "ppt/slideMasters/slideMaster1.xml", CreateSlideMasterXml("Master 1"));
    AddTextEntry(archive, "ppt/slideMasters/slideMaster2.xml", CreateSlideMasterXml("Master 2"));

    AddTextEntry(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slideMasters/_rels/slideMaster2.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreatePictureSlideXml("rIdImg"));
    AddTextEntry(archive, "ppt/slides/slide2.xml", CreatePictureFillSlideXml("rIdImg2"));

    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml" />
          <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/_rels/slide2.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml" />
          <Relationship Id="rIdImg2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slideLayouts/slideLayout1.xml", CreateReferencedBackgroundLayoutXml("layout1"));
    AddTextEntry(archive, "ppt/slideLayouts/slideLayout2.xml", CreateGradientLayoutXml("layout2"));

    AddTextEntry(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slideLayouts/_rels/slideLayout2.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster2.xml" />
          <Relationship Id="rIdLayoutImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.png" />
        </Relationships>
        """);

    AddBinaryEntry(archive, "ppt/media/image1.png", image1Png);
    AddBinaryEntry(archive, "ppt/media/image2.png", image2Png);
    AddBinaryEntry(archive, "ppt/media/image3.png", image3Png);
}

static void CreateSingleImagePptx(string path, string mediaPath, byte[] imageData)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreatePictureOnlySlideXml("rIdImg"));
    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", $$"""
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{{Path.GetFileName(mediaPath)}}" />
        </Relationships>
        """);

    AddBinaryEntry(archive, mediaPath, imageData);
}

static void CreatePlaceholderInheritancePptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/theme/theme1.xml", CreateThemeXml("PlaceholderTheme", "FFFFFF", "4B6A88"));
    AddTextEntry(archive, "ppt/slideMasters/slideMaster1.xml", CreatePlaceholderSlideMasterXml());
    AddTextEntry(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slideLayouts/slideLayout1.xml", CreatePlaceholderLayoutXml());
    AddTextEntry(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreatePlaceholderSlideXml());
    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml" />
        </Relationships>
        """);
}

static void CreateBulletInheritancePptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/theme/theme1.xml", CreateThemeXml("BulletTheme", "FFFFFF", "4B6A88"));
    AddTextEntry(archive, "ppt/slideMasters/slideMaster1.xml", CreateBulletSlideMasterXml());
    AddTextEntry(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slideLayouts/slideLayout1.xml", CreateBulletLayoutXml());
    AddTextEntry(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateBulletSlideXml());
    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml" />
        </Relationships>
        """);
}

static void CreateChartPptx(string path, string chartXml)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateChartSlideXml());
    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/charts/chart1.xml", chartXml);
}

static void CreateSmartArtPptx(string path, string dataXml, string layoutXml)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateSmartArtSlideXml());
    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdDm1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml" />
          <Relationship Id="rIdLo1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/diagrams/data1.xml", dataXml);
    AddTextEntry(archive, "ppt/diagrams/layout1.xml", layoutXml);
}

static void CreatePartialSmartArtPptx(string path, string dataXml)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreatePartialSmartArtSlideXml());
    AddTextEntry(archive, "ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdDm1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/diagrams/data1.xml", dataXml);
}

static void CreateInlineSmartArtPptx(string path, string dataXml, string layoutXml)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateInlineSmartArtSlideXml(dataXml, layoutXml));
}

static string CreateSmartArtSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="1" name="SmartArt 1" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="914400" />
              <a:ext cx="5486400" cy="3657600" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rIdDm1" r:lo="rIdLo1" />
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreatePartialSmartArtSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="1" name="SmartArt Partial" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="914400" />
              <a:ext cx="5486400" cy="3657600" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rIdDm1" />
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateInlineSmartArtSlideXml(string dataXml, string layoutXml) => $$"""
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="1" name="SmartArt Inline" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="914400" />
              <a:ext cx="5486400" cy="3657600" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:diagramData>
        {{dataXml}}
        {{layoutXml}}
                </dgm:diagramData>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static void CreateAutofitPptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateAutofitSlideXml());
}

static void CreateAutofitIndentPptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateAutofitIndentSlideXml());
}

static void CreateAutofitBulletWrapPptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateAutofitBulletWrapSlideXml());
}

static void CreateTableTextPptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateTableTextSlideXml());
}

static void CreateTableLayoutPptx(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

    AddTextEntry(archive, "ppt/presentation.xml", """
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldIdLst>
            <p:sldId id="256" r:id="rId1" />
          </p:sldIdLst>
          <p:sldSz cx="9144000" cy="6858000" />
        </p:presentation>
        """);

    AddTextEntry(archive, "ppt/_rels/presentation.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" />
        </Relationships>
        """);

    AddTextEntry(archive, "ppt/slides/slide1.xml", CreateTableLayoutSlideXml());
}

static string CreateAutofitSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Autofit Shape" />
              <p:cNvSpPr />
              <p:nvPr />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="914400" />
                <a:ext cx="3657600" cy="1828800" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr>
                <a:normAutofit fontScale="65000" lnSpcReduction="20000" />
              </a:bodyPr>
              <a:lstStyle />
              <a:p>
                <a:r>
                  <a:rPr sz="2000" />
                  <a:t>Scaled line one</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:r>
                  <a:rPr sz="2000" />
                  <a:t>Scaled line two</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateAutofitIndentSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Autofit Indent Shape" />
              <p:cNvSpPr />
              <p:nvPr />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="914400" />
                <a:ext cx="2286000" cy="1828800" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr>
                <a:normAutofit fontScale="65000" lnSpcReduction="20000" />
              </a:bodyPr>
              <a:lstStyle />
              <a:p>
                <a:pPr marL="457200" indent="228600" />
                <a:r>
                  <a:rPr sz="2000" />
                  <a:t>Indented wrapping should move the second line left.</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateAutofitBulletWrapSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Autofit Bullet Shape" />
              <p:cNvSpPr />
              <p:nvPr />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="914400" />
                <a:ext cx="2286000" cy="1828800" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
            </p:spPr>
            <p:txBody>
              <a:bodyPr>
                <a:normAutofit fontScale="65000" lnSpcReduction="20000" />
              </a:bodyPr>
              <a:lstStyle />
              <a:p>
                <a:pPr marL="457200" indent="-228600">
                  <a:buChar char="*" />
                </a:pPr>
                <a:r>
                  <a:rPr sz="2000" />
                  <a:t>Bullet wrapping should keep the continuation line aligned with the paragraph margin.</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateTableTextSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="1" name="Table 1" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="914400" />
              <a:ext cx="2286000" cy="914400" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl>
                  <a:tblPr firstRow="1" bandRow="1" />
                  <a:tblGrid>
                    <a:gridCol w="2286000" />
                  </a:tblGrid>
                  <a:tr h="914400">
                    <a:tc>
                      <a:txBody>
                        <a:bodyPr />
                        <a:lstStyle />
                        <a:p>
                          <a:pPr marL="457200" indent="-228600">
                            <a:buChar char="*" />
                          </a:pPr>
                          <a:r>
                            <a:rPr sz="1800">
                              <a:solidFill>
                                <a:srgbClr val="CC3300" />
                              </a:solidFill>
                            </a:rPr>
                            <a:t>Cell</a:t>
                          </a:r>
                          <a:r>
                            <a:rPr sz="1800">
                              <a:solidFill>
                                <a:srgbClr val="0066CC" />
                              </a:solidFill>
                            </a:rPr>
                            <a:t> detail should wrap inside the table cell</a:t>
                          </a:r>
                        </a:p>
                      </a:txBody>
                      <a:tcPr anchor="t" marL="91440" marR="91440" marT="45720" marB="45720" />
                    </a:tc>
                  </a:tr>
                </a:tbl>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateTableLayoutSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="1" name="Merged Table" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="2286000" />
              <a:ext cx="2743200" cy="914400" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl>
                  <a:tblPr />
                  <a:tblGrid>
                    <a:gridCol w="1371600" />
                    <a:gridCol w="1371600" />
                  </a:tblGrid>
                  <a:tr h="914400">
                    <a:tc gridSpan="2">
                      <a:txBody>
                        <a:bodyPr />
                        <a:lstStyle />
                        <a:p>
                          <a:r>
                            <a:rPr sz="1200" />
                            <a:t>Merged text should stay on one line</a:t>
                          </a:r>
                        </a:p>
                      </a:txBody>
                      <a:tcPr anchor="t" marL="91440" marR="91440" marT="45720" marB="45720" />
                    </a:tc>
                    <a:tc hMerge="1">
                      <a:txBody>
                        <a:bodyPr />
                        <a:lstStyle />
                        <a:p />
                      </a:txBody>
                      <a:tcPr />
                    </a:tc>
                  </a:tr>
                </a:tbl>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="2" name="Centered Table" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="914400" />
              <a:ext cx="2286000" cy="914400" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl>
                  <a:tblPr />
                  <a:tblGrid>
                    <a:gridCol w="2286000" />
                  </a:tblGrid>
                  <a:tr h="914400">
                    <a:tc>
                      <a:txBody>
                        <a:bodyPr />
                        <a:lstStyle />
                        <a:p>
                          <a:r>
                            <a:rPr sz="1200" />
                            <a:t>Centered no-wrap text should remain on a single line</a:t>
                          </a:r>
                        </a:p>
                      </a:txBody>
                      <a:tcPr anchor="ctr" noWrap="1" marL="91440" marR="91440" marT="45720" marB="45720" />
                    </a:tc>
                  </a:tr>
                </a:tbl>
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateSmartArtDataModelXml(params string[] nodeTexts)
{
    var nodes = new StringBuilder();
    var connections = new StringBuilder();

    for (var i = 0; i < nodeTexts.Length; i++)
    {
        var escapedText = System.Security.SecurityElement.Escape(nodeTexts[i]) ?? string.Empty;
        nodes.AppendLine($$"""
              <dgm:pt modelId="{{i}}" type="node">
                <dgm:t>{{escapedText}}</dgm:t>
              </dgm:pt>
            """);

        if (i > 0)
        {
            connections.AppendLine($$"""
                <dgm:cxn modelId="{{nodeTexts.Length + i}}" srcId="{{i - 1}}" destId="{{i}}" type="parOf" />
              """);
        }
    }

    return $$"""
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:ptLst>
        {{nodes.ToString().TrimEnd()}}
          </dgm:ptLst>
          <dgm:cxnLst>
        {{connections.ToString().TrimEnd()}}
          </dgm:cxnLst>
        </dgm:dataModel>
        """;
}

static string CreateSmartArtTxBodyDataModelXml(params string[] nodeTexts)
{
    var nodes = new StringBuilder();
    var connections = new StringBuilder();

    for (var i = 0; i < nodeTexts.Length; i++)
    {
        var paragraphs = new StringBuilder();
        var normalizedText = nodeTexts[i].Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        foreach (var paragraph in normalizedText.Split('\n'))
        {
            var escapedParagraph = System.Security.SecurityElement.Escape(paragraph) ?? string.Empty;
            paragraphs.AppendLine($$"""
                  <a:p>
                    <a:r>
                      <a:rPr sz="1400">
                        <a:solidFill>
                          <a:srgbClr val="2F5D7C" />
                        </a:solidFill>
                      </a:rPr>
                      <a:t>{{escapedParagraph}}</a:t>
                    </a:r>
                  </a:p>
                """);
        }

        nodes.AppendLine($$"""
              <dgm:pt modelId="{{i}}" type="node">
                <dgm:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:bodyPr />
                  <a:lstStyle />
            {{paragraphs.ToString().TrimEnd()}}
                </dgm:txBody>
              </dgm:pt>
            """);

        if (i > 0)
        {
            connections.AppendLine($$"""
                <dgm:cxn modelId="{{nodeTexts.Length + i}}" srcId="{{i - 1}}" destId="{{i}}" type="parOf" />
              """);
        }
    }

    return $$"""
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:ptLst>
        {{nodes.ToString().TrimEnd()}}
          </dgm:ptLst>
          <dgm:cxnLst>
        {{connections.ToString().TrimEnd()}}
          </dgm:cxnLst>
        </dgm:dataModel>
        """;
}

static string CreateSmartArtLayoutXml(
    string uniqueId = "urn:microsoft.com/office/officeart/2005/8/layout/VerticalBulletList",
    string name = "Vertical Bullet List",
    string description = "Minimal SmartArt layout for smoke tests") => $$"""
    <dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                   uniqueId="{{uniqueId}}"
                   name="{{name}}"
                   desc="{{description}}" />
    """;

static string CreatePictureSlideXml(string imageRelationshipId) => $$"""
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Theme Fill Shape" />
              <p:cNvSpPr />
              <p:nvPr />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="457200" />
                <a:ext cx="1828800" cy="914400" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:solidFill>
                <a:schemeClr val="accent1" />
              </a:solidFill>
            </p:spPr>
          </p:sp>
          <p:pic>
            <p:nvPicPr>
              <p:cNvPr id="2" name="Picture 1" />
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="{{imageRelationshipId}}" />
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="914400" />
                <a:ext cx="1828800" cy="1828800" />
              </a:xfrm>
            </p:spPr>
          </p:pic>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateChartSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr>
              <p:cNvPr id="1" name="Chart 1" />
              <p:cNvGraphicFramePr />
              <p:nvPr />
            </p:nvGraphicFramePr>
            <p:xfrm>
              <a:off x="914400" y="914400" />
              <a:ext cx="5486400" cy="3657600" />
            </p:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart1" />
              </a:graphicData>
            </a:graphic>
          </p:graphicFrame>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateBulletSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Bullet Body Placeholder" />
              <p:cNvSpPr />
              <p:nvPr>
                <p:ph type="body" idx="1" />
              </p:nvPr>
            </p:nvSpPr>
            <p:txBody>
              <a:bodyPr />
              <a:lstStyle />
              <a:p>
                <a:r>
                  <a:t>Inherited bullet</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr>
                  <a:buAutoNum type="arabicPeriod" startAt="3" />
                </a:pPr>
                <a:r>
                  <a:t>Numbered item three</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr>
                  <a:buAutoNum type="arabicPeriod" />
                </a:pPr>
                <a:r>
                  <a:t>Numbered item four</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr>
                  <a:buNone />
                </a:pPr>
                <a:r>
                  <a:t>Plain paragraph</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreatePictureFillSlideXml(string imageRelationshipId) => $$"""
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="4" name="Theme 2 Fill Shape" />
              <p:cNvSpPr />
              <p:nvPr />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="457200" y="2286000" />
                <a:ext cx="1828800" cy="914400" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:solidFill>
                <a:schemeClr val="accent1" />
              </a:solidFill>
            </p:spPr>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="Picture Fill" />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="3657600" y="914400" />
                <a:ext cx="1828800" cy="1828800" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:blipFill>
                <a:blip r:embed="{{imageRelationshipId}}" />
                <a:stretch>
                  <a:fillRect />
                </a:stretch>
              </a:blipFill>
            </p:spPr>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMapOvr>
        <a:overrideClrMapping bg1="bg1" tx1="tx1" bg2="bg2" tx2="tx2"
                              accent1="accent3" accent2="accent2" accent3="accent3"
                              accent4="accent4" accent5="accent5" accent6="accent6"
                              hlink="hlink" folHlink="folHlink" />
      </p:clrMapOvr>
    </p:sld>
    """;

static string CreateBarChartXml(
    string chartTitle = "Revenue by Quarter",
    string barDir = "bar",
    string seriesName = "North Region",
    string seriesColor = "3366CC",
    bool includeLegend = false)
{
    var legendXml = includeLegend
        ? "<c:legend><c:legendPos val=\"r\" /><c:overlay val=\"0\" /></c:legend>"
        : string.Empty;

    return $$"""
        <c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr />
                  <a:lstStyle />
                  <a:p>
                    <a:r>
                      <a:t>{{chartTitle}}</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
            </c:title>
            <c:plotArea>
              <c:barChart>
                <c:barDir val="{{barDir}}" />
                <c:grouping val="clustered" />
                <c:ser>
                  <c:idx val="0" />
                  <c:order val="0" />
                  <c:tx>
                    <c:v>{{seriesName}}</c:v>
                  </c:tx>
                  <c:spPr>
                    <a:solidFill>
                      <a:srgbClr val="{{seriesColor}}" />
                    </a:solidFill>
                    <a:ln w="12700">
                      <a:solidFill>
                        <a:srgbClr val="{{seriesColor}}" />
                      </a:solidFill>
                    </a:ln>
                  </c:spPr>
                  <c:cat>
                    <c:strLit>
                      <c:ptCount val="3" />
                      <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                      <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                      <c:pt idx="2"><c:v>Q3</c:v></c:pt>
                    </c:strLit>
                  </c:cat>
                  <c:val>
                    <c:numLit>
                      <c:ptCount val="3" />
                      <c:pt idx="0"><c:v>12</c:v></c:pt>
                      <c:pt idx="1"><c:v>18</c:v></c:pt>
                      <c:pt idx="2"><c:v>9</c:v></c:pt>
                    </c:numLit>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
            {{legendXml}}
          </c:chart>
        </c:chartSpace>
        """;
}

static string CreateAreaChartXml() => """
    <c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart>
        <c:title>
          <c:tx>
            <c:rich>
              <a:bodyPr />
              <a:lstStyle />
              <a:p>
                <a:r>
                  <a:t>Coverage Trend</a:t>
                </a:r>
              </a:p>
            </c:rich>
          </c:tx>
        </c:title>
        <c:plotArea>
          <c:areaChart>
            <c:grouping val="standard" />
            <c:ser>
              <c:idx val="0" />
              <c:order val="0" />
              <c:tx>
                <c:v>Open Accounts</c:v>
              </c:tx>
              <c:cat>
                <c:strLit>
                  <c:ptCount val="3" />
                  <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                  <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                  <c:pt idx="2"><c:v>Mar</c:v></c:pt>
                </c:strLit>
              </c:cat>
              <c:val>
                <c:numLit>
                  <c:ptCount val="3" />
                  <c:pt idx="0"><c:v>40</c:v></c:pt>
                  <c:pt idx="1"><c:v>48</c:v></c:pt>
                  <c:pt idx="2"><c:v>51</c:v></c:pt>
                </c:numLit>
              </c:val>
            </c:ser>
          </c:areaChart>
        </c:plotArea>
      </c:chart>
    </c:chartSpace>
    """;

static string CreatePictureOnlySlideXml(string imageRelationshipId) => $$"""
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld>
        <p:spTree>
          <p:pic>
            <p:nvPicPr>
              <p:cNvPr id="1" name="Standalone Picture" />
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="{{imageRelationshipId}}" />
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="914400" />
                <a:ext cx="2743200" cy="1828800" />
              </a:xfrm>
            </p:spPr>
          </p:pic>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateReferencedBackgroundLayoutXml(string layoutName) => $$"""
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 name="{{layoutName}}"
                 type="titleOnly">
      <p:cSld>
        <p:bg>
          <p:bgRef idx="1001">
            <a:schemeClr val="bg1" />
          </p:bgRef>
        </p:bg>
        <p:spTree />
      </p:cSld>
      <p:clrMapOvr>
        <a:overrideClrMapping bg1="bg1" tx1="tx1" bg2="bg2" tx2="tx2"
                              accent1="accent2" accent2="accent2" accent3="accent3"
                              accent4="accent4" accent5="accent5" accent6="accent6"
                              hlink="hlink" folHlink="folHlink" />
      </p:clrMapOvr>
    </p:sldLayout>
    """;

static string CreateGradientLayoutXml(string layoutName) => $$"""
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                 name="{{layoutName}}"
                 type="titleOnly">
      <p:cSld>
        <p:bg>
          <p:bgPr>
            <a:gradFill>
              <a:gsLst>
                <a:gs pos="0">
                  <a:srgbClr val="DDEEFF" />
                </a:gs>
                <a:gs pos="100000">
                  <a:srgbClr val="88AADD" />
                </a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="1" />
            </a:gradFill>
          </p:bgPr>
        </p:bg>
        <p:spTree>
          <p:pic>
            <p:nvPicPr>
              <p:cNvPr id="12" name="Layout Picture" />
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="rIdLayoutImg" />
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="6400800" y="457200" />
                <a:ext cx="457200" cy="457200" />
              </a:xfrm>
            </p:spPr>
          </p:pic>
        </p:spTree>
      </p:cSld>
    </p:sldLayout>
    """;

static string CreateSlideMasterXml(string name) => $$"""
    <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld name="{{name}}">
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="21" name="Master Accent Shape" />
              <p:cNvSpPr />
              <p:nvPr />
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="685800" y="5943600" />
                <a:ext cx="914400" cy="457200" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:solidFill>
                <a:schemeClr val="accent4">
                  <a:alpha val="50000" />
                </a:schemeClr>
              </a:solidFill>
            </p:spPr>
          </p:sp>
        </p:spTree>
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                accent1="accent1" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6"
                hlink="hlink" folHlink="folHlink" />
    </p:sldMaster>
    """;

static string CreatePlaceholderSlideMasterXml() => """
    <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld name="Placeholder Master">
        <p:spTree />
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                accent1="accent1" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6"
                hlink="hlink" folHlink="folHlink" />
      <p:txStyles>
        <p:titleStyle>
          <a:lvl1pPr algn="ctr">
            <a:defRPr sz="2600">
              <a:solidFill>
                <a:srgbClr val="B94441" />
              </a:solidFill>
            </a:defRPr>
          </a:lvl1pPr>
        </p:titleStyle>
        <p:bodyStyle>
          <a:lvl1pPr marL="457200" indent="0">
            <a:defRPr sz="1800">
              <a:solidFill>
                <a:srgbClr val="355C7D" />
              </a:solidFill>
            </a:defRPr>
          </a:lvl1pPr>
        </p:bodyStyle>
        <p:otherStyle>
          <a:lvl1pPr>
            <a:defRPr sz="1400">
              <a:solidFill>
                <a:srgbClr val="444444" />
              </a:solidFill>
            </a:defRPr>
          </a:lvl1pPr>
        </p:otherStyle>
      </p:txStyles>
    </p:sldMaster>
    """;

static string CreateBulletSlideMasterXml() => """
    <p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld name="Bullet Master">
        <p:spTree />
      </p:cSld>
      <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
                accent1="accent1" accent2="accent2" accent3="accent3"
                accent4="accent4" accent5="accent5" accent6="accent6"
                hlink="hlink" folHlink="folHlink" />
      <p:txStyles>
        <p:bodyStyle>
          <a:lvl1pPr marL="457200" indent="-228600">
            <a:buChar char="*" />
            <a:buClr>
              <a:srgbClr val="CC5500" />
            </a:buClr>
            <a:defRPr sz="1800">
              <a:solidFill>
                <a:srgbClr val="2F4858" />
              </a:solidFill>
            </a:defRPr>
          </a:lvl1pPr>
        </p:bodyStyle>
      </p:txStyles>
    </p:sldMaster>
    """;

static string CreatePlaceholderLayoutXml() => """
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 name="placeholderLayout"
                 type="titleAndObj">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Title Placeholder" />
              <p:cNvSpPr />
              <p:nvPr>
                <p:ph type="title" idx="1" />
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="457200" />
                <a:ext cx="7315200" cy="914400" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:noFill />
            </p:spPr>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Body Placeholder" />
              <p:cNvSpPr />
              <p:nvPr>
                <p:ph type="body" idx="2" />
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="1828800" />
                <a:ext cx="7315200" cy="2743200" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:noFill />
            </p:spPr>
            <p:txBody>
              <a:bodyPr />
              <a:lstStyle />
              <a:p>
                <a:r>
                  <a:t>Layout body prompt</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sldLayout>
    """;

static string CreateBulletLayoutXml() => """
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 name="bulletLayout"
                 type="titleAndObj">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Body Placeholder" />
              <p:cNvSpPr />
              <p:nvPr>
                <p:ph type="body" idx="1" />
              </p:nvPr>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="1371600" />
                <a:ext cx="7315200" cy="3657600" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
              <a:noFill />
            </p:spPr>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sldLayout>
    """;

static string CreatePlaceholderSlideXml() => """
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="1" name="Title Placeholder Instance" />
              <p:cNvSpPr />
              <p:nvPr>
                <p:ph type="title" idx="1" />
              </p:nvPr>
            </p:nvSpPr>
            <p:txBody>
              <a:bodyPr />
              <a:lstStyle />
              <a:p>
                <a:r>
                  <a:t>Inherited Title</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Body Placeholder Instance" />
              <p:cNvSpPr />
              <p:nvPr>
                <p:ph type="body" idx="2" />
              </p:nvPr>
            </p:nvSpPr>
            <p:txBody>
              <a:bodyPr />
              <a:lstStyle />
              <a:p>
                <a:r>
                  <a:t>Inherited body copy&#10;Second line</a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sld>
    """;

static string CreateThemeXml(string name, string light1, string accent1) => $$"""
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{{name}}">
      <a:themeElements>
        <a:clrScheme name="SmokeColors">
          <a:dk1>
            <a:srgbClr val="111111" />
          </a:dk1>
          <a:lt1>
            <a:srgbClr val="{{light1}}" />
          </a:lt1>
          <a:dk2>
            <a:srgbClr val="333333" />
          </a:dk2>
          <a:lt2>
            <a:srgbClr val="F4F7FB" />
          </a:lt2>
          <a:accent1>
            <a:srgbClr val="{{accent1}}" />
          </a:accent1>
          <a:accent2>
            <a:srgbClr val="D97A2B" />
          </a:accent2>
          <a:accent3>
            <a:srgbClr val="6A8F63" />
          </a:accent3>
          <a:accent4>
            <a:srgbClr val="8C5BAA" />
          </a:accent4>
          <a:accent5>
            <a:srgbClr val="3F8F9C" />
          </a:accent5>
          <a:accent6>
            <a:srgbClr val="C04B59" />
          </a:accent6>
          <a:hlink>
            <a:srgbClr val="1B66C9" />
          </a:hlink>
          <a:folHlink>
            <a:srgbClr val="7A5B99" />
          </a:folHlink>
        </a:clrScheme>
        <a:fmtScheme name="SmokeFormats">
          <a:bgFillStyleLst>
            <a:solidFill>
              <a:schemeClr val="lt1">
                <a:shade val="90000" />
              </a:schemeClr>
            </a:solidFill>
            <a:solidFill>
              <a:srgbClr val="E8EAF6" />
            </a:solidFill>
            <a:solidFill>
              <a:srgbClr val="FFFFFF" />
            </a:solidFill>
          </a:bgFillStyleLst>
        </a:fmtScheme>
      </a:themeElements>
    </a:theme>
    """;

static void AddTextEntry(ZipArchive archive, string path, string content)
{
    var entry = archive.CreateEntry(path);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    writer.Write(content);
}

static void AddBinaryEntry(ZipArchive archive, string path, byte[] content)
{
    var entry = archive.CreateEntry(path);
    using var stream = entry.Open();
    stream.Write(content, 0, content.Length);
}

static (byte R, byte G, byte B, byte A) Px(byte r, byte g, byte b, byte a) => (r, g, b, a);

static byte[] CreateRgbaPng(int width, int height, params (byte R, byte G, byte B, byte A)[] pixels)
{
    Assert(pixels.Length == width * height, "Pixel count should match the requested PNG dimensions.");

    var scanData = new byte[height * (1 + width * 4)];
    var offset = 0;
    for (var y = 0; y < height; y++)
    {
        scanData[offset++] = 0;
        for (var x = 0; x < width; x++)
        {
            var pixel = pixels[y * width + x];
            scanData[offset++] = pixel.R;
            scanData[offset++] = pixel.G;
            scanData[offset++] = pixel.B;
            scanData[offset++] = pixel.A;
        }
    }

    using var png = new MemoryStream();
    png.Write(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, 0, 8);

    using (var ihdr = new MemoryStream())
    {
        WriteUInt32BigEndian(ihdr, (uint)width);
        WriteUInt32BigEndian(ihdr, (uint)height);
        ihdr.WriteByte(8);
        ihdr.WriteByte(6);
        ihdr.WriteByte(0);
        ihdr.WriteByte(0);
        ihdr.WriteByte(0);
        WritePngChunk(png, "IHDR", ihdr.ToArray());
    }

    WritePngChunk(png, "IDAT", CompressZlib(scanData));
    WritePngChunk(png, "IEND", Array.Empty<byte>());
    return png.ToArray();
}

static byte[] CreateBmp32(int width, int height, params (byte R, byte G, byte B, byte A)[] pixels)
{
    Assert(pixels.Length == width * height, "Pixel count should match the requested BMP dimensions.");

    const int fileHeaderSize = 14;
    const int dibHeaderSize = 40;
    const int pixelOffset = fileHeaderSize + dibHeaderSize;
    const int bytesPerPixel = 4;
    var rowSize = width * bytesPerPixel;
    var imageSize = rowSize * height;
    var fileSize = pixelOffset + imageSize;
    var bmp = new byte[fileSize];

    bmp[0] = (byte)'B';
    bmp[1] = (byte)'M';
    WriteInt32LittleEndian(bmp, 2, fileSize);
    WriteInt32LittleEndian(bmp, 10, pixelOffset);
    WriteInt32LittleEndian(bmp, 14, dibHeaderSize);
    WriteInt32LittleEndian(bmp, 18, width);
    WriteInt32LittleEndian(bmp, 22, height);
    WriteUInt16LittleEndian(bmp, 26, 1);
    WriteUInt16LittleEndian(bmp, 28, 32);
    WriteInt32LittleEndian(bmp, 34, imageSize);

    for (var y = 0; y < height; y++)
    {
        var srcRow = y * width;
        var dstOffset = pixelOffset + (height - 1 - y) * rowSize;
        for (var x = 0; x < width; x++)
        {
            var pixel = pixels[srcRow + x];
            bmp[dstOffset++] = pixel.B;
            bmp[dstOffset++] = pixel.G;
            bmp[dstOffset++] = pixel.R;
            bmp[dstOffset++] = pixel.A;
        }
    }

    return bmp;
}

static byte[] CreateMinimalGif()
{
    return new byte[]
    {
        (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a',
        1, 0, 1, 0,
        0, 0, 0
    };
}

static byte[] CreateMinimalTiff()
{
    return new byte[]
    {
        0x49, 0x49, 0x2A, 0x00,
        0x08, 0x00, 0x00, 0x00,
        0x03, 0x00,
        0x00, 0x01, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
        0x01, 0x01, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
        0x02, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00
    };
}

static void WritePngChunk(Stream stream, string chunkType, byte[] data)
{
    var typeBytes = Encoding.ASCII.GetBytes(chunkType);
    WriteUInt32BigEndian(stream, (uint)data.Length);
    stream.Write(typeBytes, 0, typeBytes.Length);
    stream.Write(data, 0, data.Length);

    var crcInput = new byte[typeBytes.Length + data.Length];
    Buffer.BlockCopy(typeBytes, 0, crcInput, 0, typeBytes.Length);
    Buffer.BlockCopy(data, 0, crcInput, typeBytes.Length, data.Length);
    WriteUInt32BigEndian(stream, ComputeCrc32(crcInput));
}

static void WriteUInt32BigEndian(Stream stream, uint value)
{
    stream.WriteByte((byte)(value >> 24));
    stream.WriteByte((byte)(value >> 16));
    stream.WriteByte((byte)(value >> 8));
    stream.WriteByte((byte)value);
}

static void WriteUInt16LittleEndian(byte[] buffer, int offset, ushort value)
{
    buffer[offset + 0] = (byte)value;
    buffer[offset + 1] = (byte)(value >> 8);
}

static void WriteInt32LittleEndian(byte[] buffer, int offset, int value)
{
    buffer[offset + 0] = (byte)value;
    buffer[offset + 1] = (byte)(value >> 8);
    buffer[offset + 2] = (byte)(value >> 16);
    buffer[offset + 3] = (byte)(value >> 24);
}

static byte[] CompressZlib(byte[] data)
{
    using var output = new MemoryStream();
    output.WriteByte(0x78);
    output.WriteByte(0x9C);

    using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true))
    {
        deflate.Write(data, 0, data.Length);
    }

    var checksum = ComputeAdler32(data);
    WriteUInt32BigEndian(output, checksum);
    return output.ToArray();
}

static byte[] InflateZlib(byte[] data)
{
    using var input = new MemoryStream(data, 2, data.Length - 6, writable: false);
    using var deflate = new DeflateStream(input, CompressionMode.Decompress);
    using var output = new MemoryStream();
    deflate.CopyTo(output);
    return output.ToArray();
}

static uint ComputeAdler32(byte[] data)
{
    const uint mod = 65521;
    uint a = 1;
    uint b = 0;

    foreach (var value in data)
    {
        a = (a + value) % mod;
        b = (b + a) % mod;
    }

    return (b << 16) | a;
}

static uint ComputeCrc32(byte[] data)
{
    uint crc = 0xFFFFFFFF;
    foreach (var value in data)
    {
        crc ^= value;
        for (var bit = 0; bit < 8; bit++)
        {
            var mask = (crc & 1) == 1 ? 0xEDB88320u : 0u;
            crc = (crc >> 1) ^ mask;
        }
    }

    return ~crc;
}

static int CountOccurrences(string text, string value)
{
    int count = 0;
    int index = 0;

    while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
    {
        count++;
        index += value.Length;
    }

    return count;
}

static void AssertPdfUsesRegisteredFontResources(string pdfText, string message)
{
    Assert(
        System.Text.RegularExpressions.Regex.IsMatch(
            pdfText,
            @"/F\d+\s+\d+(?:\.\d+)?\s+Tf",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant),
        message);

    Assert(
        System.Text.RegularExpressions.Regex.IsMatch(
            pdfText,
            @"/Font\s*<<[\s\S]*?/F\d+\s+\d+\s+\d+\s+R",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant),
        message);
}

static void Assert(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

static void AssertThrows<TException>(Action action, string message) where TException : Exception
{
    try
    {
        action();
    }
    catch (TException)
    {
        return;
    }

    throw new InvalidOperationException(message);
}

static void AssertColor(Color? color, byte r, byte g, byte b, byte a, string message)
{
    Assert(color.HasValue, message);
    var value = color.GetValueOrDefault();
    Assert(value.R == r && value.G == g && value.B == b && value.A == a, message);
}
