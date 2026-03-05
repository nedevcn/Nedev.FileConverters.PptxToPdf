namespace NPptxToPdf.Image;

public static class ImageDecoder
{
    public static ImageInfo Decode(byte[] data)
    {
        if (data == null || data.Length < 8)
            throw new ArgumentException("Invalid image data");

        // Check magic numbers
        if (IsPng(data))
            return DecodePng(data);
        if (IsJpeg(data))
            return DecodeJpeg(data);
        if (IsGif(data))
            return DecodeGif(data);
        if (IsBmp(data))
            return DecodeBmp(data);
        if (IsTiff(data))
            return DecodeTiff(data);

        throw new NotSupportedException("Unsupported image format");
    }

    public static bool IsPng(byte[] data)
    {
        return data.Length >= 8 &&
               data[0] == 0x89 && data[1] == 0x50 &&
               data[2] == 0x4E && data[3] == 0x47 &&
               data[4] == 0x0D && data[5] == 0x0A &&
               data[6] == 0x1A && data[7] == 0x0A;
    }

    public static bool IsJpeg(byte[] data)
    {
        return data.Length >= 2 &&
               data[0] == 0xFF && data[1] == 0xD8;
    }

    public static bool IsGif(byte[] data)
    {
        return data.Length >= 6 &&
               ((data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46) || // GIF87a
                (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46));   // GIF89a
    }

    public static bool IsBmp(byte[] data)
    {
        return data.Length >= 2 &&
               data[0] == 0x42 && data[1] == 0x4D; // BM
    }

    public static bool IsTiff(byte[] data)
    {
        return data.Length >= 4 &&
               ((data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00) || // Little endian
                (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A));   // Big endian
    }

    private static ImageInfo DecodePng(byte[] data)
    {
        // PNG IHDR chunk starts at byte 16
        if (data.Length < 24)
            throw new ArgumentException("Invalid PNG data");

        var width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
        var height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
        var bitDepth = data[24];
        var colorType = data[25];

        return new ImageInfo
        {
            Width = width,
            Height = height,
            Format = ImageFormat.Png,
            BitDepth = bitDepth,
            HasAlpha = colorType == 4 || colorType == 6,
            RawData = data
        };
    }

    private static ImageInfo DecodeJpeg(byte[] data)
    {
        int i = 2;
        while (i < data.Length - 1)
        {
            if (data[i] != 0xFF)
            {
                i++;
                continue;
            }

            byte marker = data[i + 1];

            if (marker == 0xD9 || marker == 0xDA)
                break;

            if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2)
            {
                if (i + 9 < data.Length)
                {
                    int height = (data[i + 5] << 8) | data[i + 6];
                    int width = (data[i + 7] << 8) | data[i + 8];

                    return new ImageInfo
                    {
                        Width = width,
                        Height = height,
                        Format = ImageFormat.Jpeg,
                        BitDepth = 8,
                        HasAlpha = false,
                        RawData = data
                    };
                }
            }

            if (i + 3 < data.Length)
            {
                int length = (data[i + 2] << 8) | data[i + 3];
                i += 2 + length;
            }
            else
            {
                break;
            }
        }

        throw new ArgumentException("Could not decode JPEG dimensions");
    }

    private static ImageInfo DecodeGif(byte[] data)
    {
        if (data.Length < 10)
            throw new ArgumentException("Invalid GIF data");

        int width = data[6] | (data[7] << 8);
        int height = data[8] | (data[9] << 8);
        var packed = data[10];
        bool hasGlobalColorTable = (packed & 0x80) != 0;

        return new ImageInfo
        {
            Width = width,
            Height = height,
            Format = ImageFormat.Gif,
            BitDepth = ((packed & 0x07) + 1),
            HasAlpha = false, // GIF doesn't support alpha in the traditional sense
            RawData = data
        };
    }

    private static ImageInfo DecodeBmp(byte[] data)
    {
        if (data.Length < 26)
            throw new ArgumentException("Invalid BMP data");

        // BMP header
        int headerSize = data[14] | (data[15] << 8) | (data[16] << 16) | (data[17] << 24);

        int width, height;
        short bitsPerPixel;

        if (headerSize == 12) // BITMAPCOREHEADER
        {
            width = data[18] | (data[19] << 8);
            height = data[20] | (data[21] << 8);
            bitsPerPixel = (short)(data[24] | (data[25] << 8));
        }
        else // BITMAPINFOHEADER or later
        {
            width = data[18] | (data[19] << 8) | (data[20] << 16) | (data[21] << 24);
            height = data[22] | (data[23] << 8) | (data[24] << 16) | (data[25] << 24);
            bitsPerPixel = (short)(data[28] | (data[29] << 8));
        }

        return new ImageInfo
        {
            Width = width,
            Height = height,
            Format = ImageFormat.Bmp,
            BitDepth = bitsPerPixel,
            HasAlpha = bitsPerPixel == 32,
            RawData = data
        };
    }

    private static ImageInfo DecodeTiff(byte[] data)
    {
        bool littleEndian = data[0] == 0x49;

        int ifdOffset = ReadInt(data, 4, littleEndian);

        if (ifdOffset + 2 > data.Length)
            throw new ArgumentException("Invalid TIFF data");

        int numEntries = ReadShort(data, ifdOffset, littleEndian);
        int offset = ifdOffset + 2;

        int width = 0, height = 0, bitsPerSample = 8;

        for (int i = 0; i < numEntries && offset + 12 <= data.Length; i++)
        {
            int tag = ReadShort(data, offset, littleEndian);
            int type = ReadShort(data, offset + 2, littleEndian);
            int count = ReadInt(data, offset + 4, littleEndian);
            int value = ReadInt(data, offset + 8, littleEndian);

            switch (tag)
            {
                case 256: // ImageWidth
                    width = value;
                    break;
                case 257: // ImageLength
                    height = value;
                    break;
                case 258: // BitsPerSample
                    bitsPerSample = value;
                    break;
            }

            offset += 12;
        }

        if (width == 0 || height == 0)
            throw new ArgumentException("Could not decode TIFF dimensions");

        return new ImageInfo
        {
            Width = width,
            Height = height,
            Format = ImageFormat.Tiff,
            BitDepth = bitsPerSample,
            HasAlpha = false,
            RawData = data
        };
    }

    public static short ReadShort(byte[] data, int offset, bool littleEndian)
    {
        if (littleEndian)
            return (short)(data[offset] | (data[offset + 1] << 8));
        else
            return (short)((data[offset] << 8) | data[offset + 1]);
    }

    public static int ReadInt(byte[] data, int offset, bool littleEndian)
    {
        if (littleEndian)
            return data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
        else
            return (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
    }

    public static byte[] ConvertToJpeg(byte[] data)
    {
        var info = Decode(data);

        // If already JPEG, return as-is
        if (info.Format == ImageFormat.Jpeg)
            return data;

        // For other formats, we would need to implement conversion
        // For now, throw exception - in production, use System.Drawing or SkiaSharp
        throw new NotImplementedException($"Conversion from {info.Format} to JPEG not implemented. " +
            "Only JPEG images are supported for direct embedding in PDF.");
    }
}

public class ImageInfo
{
    public int Width { get; set; }
    public int Height { get; set; }
    public ImageFormat Format { get; set; }
    public int BitDepth { get; set; }
    public bool HasAlpha { get; set; }
    public byte[] RawData { get; set; } = Array.Empty<byte>();
}

public enum ImageFormat
{
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff
}
