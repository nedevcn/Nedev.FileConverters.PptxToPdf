using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.PptxToPdf;

[FileConverter("pptx", "pdf")]
public class PptxToPdfFileConverter : IFileConverter
{
    private readonly PptxToPdfConverter _converter;

    public PptxToPdfFileConverter()
    {
        _converter = new PptxToPdfConverter();
    }

    public Stream Convert(Stream input)
    {
        if (input == null)
            throw new ArgumentNullException(nameof(input));

        var outputStream = new MemoryStream();
        _converter.Convert(input, outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    public async Task<Stream> ConvertAsync(Stream input, CancellationToken cancellationToken = default)
    {
        if (input == null)
            throw new ArgumentNullException(nameof(input));

        var outputStream = new MemoryStream();
        await Task.Run(() => _converter.Convert(input, outputStream), cancellationToken);
        outputStream.Position = 0;
        return outputStream;
    }

    public PptxToPdfConverter Converter => _converter;
}