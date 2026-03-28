namespace Nedev.FileConverters.PptxToPdf.Image;

public sealed class PreparedImageData
{
    public PreparedImageData(byte[] data, int width, int height, bool isJpeg, byte[]? alphaMaskData = null)
    {
        Data = data;
        Width = width;
        Height = height;
        IsJpeg = isJpeg;
        AlphaMaskData = alphaMaskData;
    }

    public byte[] Data { get; }
    public int Width { get; }
    public int Height { get; }
    public bool IsJpeg { get; }
    public byte[]? AlphaMaskData { get; }
}
