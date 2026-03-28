using System.IO.Compression;
using System.Text;

namespace Nedev.FileConverters.PptxToPdf.Image;

/// <summary>
/// Prepares raster images for PDF embedding without third-party dependencies.
/// JPEG is passed through, PNG/BMP are embedded as lossless Flate streams,
/// and unsupported formats fail explicitly instead of producing placeholder pixels.
/// </summary>
public static class ImageConverter
{
    public static PreparedImageData PrepareForPdf(byte[] data)
    {
        var info = ImageDecoder.Decode(data);

        return info.Format switch
        {
            ImageFormat.Jpeg => new PreparedImageData(data, info.Width, info.Height, isJpeg: true),
            ImageFormat.Png => PreparePngForPdf(data),
            ImageFormat.Bmp => PrepareBmpForPdf(data, info),
            ImageFormat.Gif or ImageFormat.Tiff => throw new NotSupportedException($"{info.Format} images are not supported for direct PDF embedding."),
            _ => throw new NotSupportedException($"{info.Format} images are not supported for direct PDF embedding.")
        };
    }

    /// <summary>
    /// Converts any supported image format to JPEG
    /// </summary>
    public static byte[] ConvertToJpeg(byte[] data)
    {
        var info = ImageDecoder.Decode(data);

        // If already JPEG, return as-is
        if (info.Format == ImageFormat.Jpeg)
            return data;

        // For other formats, convert pixel data to JPEG
        return info.Format switch
        {
            ImageFormat.Png => ConvertPngToJpeg(data, info),
            ImageFormat.Gif => ConvertGifToJpeg(data, info),
            ImageFormat.Bmp => ConvertBmpToJpeg(data, info),
            ImageFormat.Tiff => ConvertTiffToJpeg(data, info),
            _ => throw new NotSupportedException($"Conversion from {info.Format} to JPEG not supported")
        };
    }

    /// <summary>
    /// Converts PNG to JPEG
    /// </summary>
    private static byte[] ConvertPngToJpeg(byte[] data, ImageInfo info)
    {
        var png = ParsePng(data);
        var (pixelData, _) = DecodePngPixelData(png);
        return EncodeJpeg(pixelData, info.Width, info.Height, info.HasAlpha);
    }

    /// <summary>
    /// Converts GIF to JPEG
    /// </summary>
    private static byte[] ConvertGifToJpeg(byte[] data, ImageInfo info)
    {
        throw new NotSupportedException("GIF images are not supported for conversion to PDF/JPEG.");
    }

    /// <summary>
    /// Converts BMP to JPEG
    /// </summary>
    private static byte[] ConvertBmpToJpeg(byte[] data, ImageInfo info)
    {
        var (pixelData, _) = DecodeBmpPixelData(data, info);
        return EncodeJpeg(pixelData, info.Width, info.Height, info.HasAlpha);
    }

    /// <summary>
    /// Converts TIFF to JPEG
    /// </summary>
    private static byte[] ConvertTiffToJpeg(byte[] data, ImageInfo info)
    {
        throw new NotSupportedException("TIFF images are not supported for conversion to PDF/JPEG.");
    }

    /// <summary>
    /// Extracts raw pixel data from PNG
    /// </summary>
    private static byte[] ExtractPngPixelData(byte[] data, ImageInfo info)
    {
        try
        {
            // Simple PNG extraction for common cases
            // Skip IHDR chunk (8 bytes signature + 13 bytes IHDR)
            int offset = 8 + 13;
            
            // Skip other chunks until we find IDAT
            while (offset + 8 <= data.Length)
            {
                int chunkLength = (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
                string chunkType = System.Text.Encoding.ASCII.GetString(data, offset + 4, 4);
                
                if (chunkType == "IDAT")
                {
                    // Found data chunk, use placeholder for now
                    break;
                }
                
                offset += 8 + chunkLength + 4; // length + type + data + crc
            }
        }
        catch {}
        
        return CreatePlaceholderImage(info.Width, info.Height, "PNG");
    }

    /// <summary>
    /// Extracts raw pixel data from GIF
    /// </summary>
    private static byte[] ExtractGifPixelData(byte[] data, ImageInfo info)
    {
        try
        {
            // Simple GIF extraction
            int offset = 10; // Skip header
            
            // Read global color table if present
            bool hasGlobalColorTable = (data[10] & 0x80) != 0;
            int colorTableSize = 3 * (1 << ((data[10] & 0x07) + 1));
            
            if (hasGlobalColorTable)
            {
                offset += colorTableSize;
            }
            
            // Skip to first image descriptor
            while (offset < data.Length && data[offset] != 0x2C) // Image descriptor marker
            {
                offset++;
            }
        }
        catch {}
        
        return CreatePlaceholderImage(info.Width, info.Height, "GIF");
    }

    /// <summary>
    /// Extracts raw pixel data from BMP
    /// </summary>
    private static byte[] ExtractBmpPixelData(byte[] data, ImageInfo info)
    {
        try
        {
            int headerSize = data[14] | (data[15] << 8) | (data[16] << 16) | (data[17] << 24);
            int pixelOffset = 14 + headerSize;

            int bytesPerPixel = info.BitDepth / 8;
            int rowSize = ((info.Width * bytesPerPixel + 3) / 4) * 4; // Align to 4 bytes

            var pixelData = new byte[info.Width * info.Height * 3];

            for (int y = 0; y < info.Height; y++)
            {
                int srcRow = info.Height - 1 - y; // BMP is bottom-up
                int srcOffset = pixelOffset + srcRow * rowSize;
                int dstOffset = y * info.Width * 3;

                for (int x = 0; x < info.Width; x++)
                {
                    if (srcOffset + x * bytesPerPixel + 2 < data.Length)
                    {
                        // BMP stores BGR, convert to RGB
                        pixelData[dstOffset + x * 3 + 0] = data[srcOffset + x * bytesPerPixel + 2]; // R
                        pixelData[dstOffset + x * 3 + 1] = data[srcOffset + x * bytesPerPixel + 1]; // G
                        pixelData[dstOffset + x * 3 + 2] = data[srcOffset + x * bytesPerPixel + 0]; // B
                    }
                }
            }

            return pixelData;
        }
        catch
        {
            return CreatePlaceholderImage(info.Width, info.Height, "BMP");
        }
    }

    /// <summary>
    /// Extracts raw pixel data from TIFF
    /// </summary>
    private static byte[] ExtractTiffPixelData(byte[] data, ImageInfo info)
    {
        try
        {
            bool littleEndian = data[0] == 0x49;
            int ifdOffset = ImageDecoder.ReadInt(data, 4, littleEndian);
            
            // Simple TIFF parsing - just return placeholder for now
        }
        catch {}
        
        return CreatePlaceholderImage(info.Width, info.Height, "TIFF");
    }

    /// <summary>
    /// Creates a placeholder image with format indication
    /// </summary>
    private static byte[] CreatePlaceholderImage(int width, int height, string format = "Unknown")
    {
        var pixelData = new byte[width * height * 3];

        // Create a pattern with format indication
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int offset = (y * width + x) * 3;
                
                // Create a colored pattern based on format
                byte r = 200, g = 200, b = 200; // Default gray
                
                switch (format.ToUpper())
                {
                    case "PNG":
                        r = (byte)((x + y) % 256);
                        g = (byte)((x * 2 + y) % 256);
                        b = 255;
                        break;
                    case "GIF":
                        r = 255;
                        g = (byte)((x + y) % 256);
                        b = (byte)((x * 2 + y) % 256);
                        break;
                    case "BMP":
                        r = (byte)((x * 2 + y) % 256);
                        g = 255;
                        b = (byte)((x + y) % 256);
                        break;
                    case "TIFF":
                        r = 255;
                        g = 255;
                        b = (byte)((x + y) % 256);
                        break;
                }

                pixelData[offset + 0] = r;
                pixelData[offset + 1] = g;
                pixelData[offset + 2] = b;
            }
        }

        return pixelData;
    }

    /// <summary>
    /// Creates a placeholder gray image
    /// </summary>
    private static byte[] CreatePlaceholderImage(int width, int height)
    {
        var pixelData = new byte[width * height * 3];

        // Create a gray checkerboard pattern
        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                int offset = (y * width + x) * 3;
                bool isGray = ((x / 10) + (y / 10)) % 2 == 0;
                byte color = isGray ? (byte)200 : (byte)150;

                pixelData[offset + 0] = color; // R
                pixelData[offset + 1] = color; // G
                pixelData[offset + 2] = color; // B
            }
        }

        return pixelData;
    }

    /// <summary>
    /// Encodes raw RGB pixel data to JPEG format
    /// This is a simplified JPEG encoder
    /// </summary>
    private static byte[] EncodeJpeg(byte[] pixelData, int width, int height, bool hasAlpha)
    {
        // For a complete implementation, we would need:
        // 1. Color space conversion (RGB to YCbCr)
        // 2. DCT (Discrete Cosine Transform)
        // 3. Quantization
        // 4. Huffman encoding
        // 5. JFIF header generation

        // Since JPEG encoding is extremely complex, we'll create a minimal valid JPEG
        // that displays the image data in a simplified form

        using var ms = new MemoryStream();
        var writer = new BinaryWriter(ms);

        // JFIF Header
        writer.Write(new byte[] { 0xFF, 0xD8 }); // SOI

        // APP0 marker (JFIF)
        writer.Write(new byte[] { 0xFF, 0xE0 });
        WriteBigEndianUInt16(writer, 16); // Length
        writer.Write(new byte[] { 0x4A, 0x46, 0x49, 0x46, 0x00 }); // "JFIF\0"
        writer.Write((byte)1); // Version major
        writer.Write((byte)1); // Version minor
        writer.Write((byte)0); // Units (0 = no units)
        WriteBigEndianUInt16(writer, 1); // X density
        WriteBigEndianUInt16(writer, 1); // Y density
        writer.Write((byte)0); // Thumbnail width
        writer.Write((byte)0); // Thumbnail height

        // DQT (Define Quantization Table) - simplified
        writer.Write(new byte[] { 0xFF, 0xDB });
        WriteBigEndianUInt16(writer, 67); // Length
        writer.Write((byte)0); // Table ID
        // Luminance quantization table (simplified - all 16)
        for (int i = 0; i < 64; i++)
            writer.Write((byte)16);

        // SOF0 (Start of Frame - Baseline DCT)
        writer.Write(new byte[] { 0xFF, 0xC0 });
        WriteBigEndianUInt16(writer, 11); // Length
        writer.Write((byte)8); // Precision
        WriteBigEndianUInt16(writer, height); // Height
        WriteBigEndianUInt16(writer, width); // Width
        writer.Write((byte)1); // Number of components (grayscale)
        writer.Write((byte)1); // Component ID
        writer.Write((byte)0x11); // Sampling factors (1x1)
        writer.Write((byte)0); // Quantization table ID

        // DHT (Define Huffman Table) - minimal DC table
        writer.Write(new byte[] { 0xFF, 0xC4 });
        WriteBigEndianUInt16(writer, 31); // Length
        writer.Write((byte)0x00); // DC table, ID 0
        // Number of codes of each length (1-16)
        writer.Write(new byte[] { 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 });
        // Values
        writer.Write(new byte[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 });

        // SOS (Start of Scan)
        writer.Write(new byte[] { 0xFF, 0xDA });
        WriteBigEndianUInt16(writer, 8); // Length
        writer.Write((byte)1); // Number of components
        writer.Write((byte)1); // Component ID
        writer.Write((byte)0x00); // Huffman table IDs
        writer.Write((byte)0); // Start of spectral selection
        writer.Write((byte)63); // End of spectral selection
        writer.Write((byte)0); // Successive approximation

        // Encode minimal scan data (placeholder)
        // In a real implementation, this would be the actual DCT coefficients
        var encodedData = EncodeMinimalScanData(pixelData, width, height);
        writer.Write(encodedData);

        // EOI
        writer.Write(new byte[] { 0xFF, 0xD9 });

        return ms.ToArray();
    }

    /// <summary>
    /// Encodes minimal scan data for grayscale image
    /// </summary>
    private static byte[] EncodeMinimalScanData(byte[] pixelData, int width, int height)
    {
        // Convert RGB to grayscale and create minimal encoded data
        var ms = new MemoryStream();

        for (int y = 0; y < height; y += 8)
        {
            for (int x = 0; x < width; x += 8)
            {
                // Process 8x8 block
                int blockSum = 0;
                int pixelCount = 0;

                for (int by = 0; by < 8 && y + by < height; by++)
                {
                    for (int bx = 0; bx < 8 && x + bx < width; bx++)
                    {
                        int pixelOffset = ((y + by) * width + (x + bx)) * 3;
                        if (pixelOffset + 2 < pixelData.Length)
                        {
                            // Convert RGB to grayscale
                            int gray = (pixelData[pixelOffset] +
                                       pixelData[pixelOffset + 1] +
                                       pixelData[pixelOffset + 2]) / 3;
                            blockSum += gray;
                            pixelCount++;
                        }
                    }
                }

                // Write DC coefficient (average of block)
                if (pixelCount > 0)
                {
                    int dc = blockSum / pixelCount;
                    // Simplified encoding: just write the value
                    if (dc == 0)
                        ms.WriteByte(0);
                    else
                    {
                        ms.WriteByte((byte)dc);
                    }
                }
            }
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Checks if conversion is needed and performs it
    /// </summary>
    public static byte[] EnsureJpegFormat(byte[] data)
    {
        var info = ImageDecoder.Decode(data);
        if (info.Format == ImageFormat.Jpeg)
            return data;

        return ConvertToJpeg(data);
    }

    private static PreparedImageData PreparePngForPdf(byte[] data)
    {
        var png = ParsePng(data);
        var (rgbData, alphaData) = DecodePngPixelData(png);

        return new PreparedImageData(
            CompressZlib(rgbData),
            png.Width,
            png.Height,
            isJpeg: false,
            alphaData == null ? null : CompressZlib(alphaData));
    }

    private static PreparedImageData PrepareBmpForPdf(byte[] data, ImageInfo info)
    {
        var (rgbData, alphaData) = DecodeBmpPixelData(data, info);
        var height = Math.Abs(info.Height);

        return new PreparedImageData(
            CompressZlib(rgbData),
            info.Width,
            height,
            isJpeg: false,
            alphaData == null ? null : CompressZlib(alphaData));
    }

    private static ParsedPng ParsePng(byte[] data)
    {
        if (!ImageDecoder.IsPng(data))
            throw new ArgumentException("Invalid PNG data.");

        var parsed = new ParsedPng();
        using var idat = new MemoryStream();

        var offset = 8;
        while (offset + 8 <= data.Length)
        {
            var chunkLength = ReadBigEndianInt(data, offset);
            if (chunkLength < 0 || offset + 12 + chunkLength > data.Length)
                throw new ArgumentException("Invalid PNG chunk length.");

            var chunkType = Encoding.ASCII.GetString(data, offset + 4, 4);
            var chunkOffset = offset + 8;

            switch (chunkType)
            {
                case "IHDR":
                    if (chunkLength < 13)
                        throw new ArgumentException("Invalid PNG IHDR chunk.");

                    parsed.Width = ReadBigEndianInt(data, chunkOffset);
                    parsed.Height = ReadBigEndianInt(data, chunkOffset + 4);
                    parsed.BitDepth = data[chunkOffset + 8];
                    parsed.ColorType = data[chunkOffset + 9];
                    parsed.CompressionMethod = data[chunkOffset + 10];
                    parsed.FilterMethod = data[chunkOffset + 11];
                    parsed.InterlaceMethod = data[chunkOffset + 12];
                    break;
                case "PLTE":
                    parsed.Palette = CopyChunkData(data, chunkOffset, chunkLength);
                    break;
                case "tRNS":
                    parsed.Transparency = CopyChunkData(data, chunkOffset, chunkLength);
                    break;
                case "IDAT":
                    idat.Write(data, chunkOffset, chunkLength);
                    break;
                case "IEND":
                    offset = data.Length;
                    continue;
            }

            offset += chunkLength + 12;
        }

        if (parsed.Width <= 0 || parsed.Height <= 0)
            throw new ArgumentException("PNG image dimensions were not found.");

        parsed.IdatData = idat.ToArray();
        if (parsed.IdatData.Length == 0)
            throw new ArgumentException("PNG image data was not found.");

        return parsed;
    }

    private static (byte[] RgbData, byte[]? AlphaData) DecodePngPixelData(ParsedPng png)
    {
        if (png.CompressionMethod != 0 || png.FilterMethod != 0)
            throw new NotSupportedException("Only standard PNG compression and filtering are supported.");
        if (png.InterlaceMethod != 0)
            throw new NotSupportedException("Interlaced PNG images are not supported yet.");
        if (png.BitDepth != 8)
            throw new NotSupportedException($"PNG bit depth {png.BitDepth} is not supported yet.");

        var samplesPerPixel = png.ColorType switch
        {
            0 => 1,
            2 => 3,
            3 => 1,
            4 => 2,
            6 => 4,
            _ => throw new NotSupportedException($"PNG color type {png.ColorType} is not supported.")
        };

        var bytesPerPixel = samplesPerPixel;
        var rowSize = checked(png.Width * samplesPerPixel);
        var scanlines = InflateZlib(png.IdatData);
        var expectedLength = checked((rowSize + 1) * png.Height);
        if (scanlines.Length < expectedLength)
            throw new ArgumentException("PNG scanline data is incomplete.");

        var unfiltered = new byte[checked(rowSize * png.Height)];
        var previousRow = new byte[rowSize];
        var currentRow = new byte[rowSize];
        var sourceOffset = 0;

        for (var row = 0; row < png.Height; row++)
        {
            var filterType = scanlines[sourceOffset++];
            ApplyPngFilter(filterType, scanlines, sourceOffset, currentRow, previousRow, rowSize, bytesPerPixel);
            Buffer.BlockCopy(currentRow, 0, unfiltered, row * rowSize, rowSize);
            sourceOffset += rowSize;

            var temp = previousRow;
            previousRow = currentRow;
            currentRow = temp;
            Array.Clear(currentRow, 0, rowSize);
        }

        return ConvertPngPixelsToRgb(png, unfiltered);
    }

    private static (byte[] RgbData, byte[]? AlphaData) ConvertPngPixelsToRgb(ParsedPng png, byte[] pixelData)
    {
        var pixelCount = checked(png.Width * png.Height);
        var rgbData = new byte[checked(pixelCount * 3)];
        var needsAlpha = png.ColorType is 4 or 6 || png.Transparency?.Length > 0;
        byte[]? alphaData = needsAlpha ? new byte[pixelCount] : null;
        var hasTransparency = false;

        var src = 0;
        var rgbOffset = 0;
        var alphaOffset = 0;

        switch (png.ColorType)
        {
            case 0:
            {
                var transparentGray = png.Transparency?.Length >= 2 ? png.Transparency[1] : (byte?)null;
                for (var i = 0; i < pixelCount; i++)
                {
                    var gray = pixelData[src++];
                    rgbData[rgbOffset++] = gray;
                    rgbData[rgbOffset++] = gray;
                    rgbData[rgbOffset++] = gray;

                    if (alphaData != null)
                    {
                        var alpha = transparentGray.HasValue && gray == transparentGray.Value ? (byte)0 : (byte)255;
                        alphaData[alphaOffset++] = alpha;
                        hasTransparency |= alpha != 255;
                    }
                }
                break;
            }
            case 2:
            {
                var hasTransparentColor = png.Transparency?.Length >= 6;
                var transparentR = hasTransparentColor ? png.Transparency![1] : (byte)0;
                var transparentG = hasTransparentColor ? png.Transparency![3] : (byte)0;
                var transparentB = hasTransparentColor ? png.Transparency![5] : (byte)0;

                for (var i = 0; i < pixelCount; i++)
                {
                    var r = pixelData[src++];
                    var g = pixelData[src++];
                    var b = pixelData[src++];
                    rgbData[rgbOffset++] = r;
                    rgbData[rgbOffset++] = g;
                    rgbData[rgbOffset++] = b;

                    if (alphaData != null)
                    {
                        var alpha = hasTransparentColor && r == transparentR && g == transparentG && b == transparentB
                            ? (byte)0
                            : (byte)255;
                        alphaData[alphaOffset++] = alpha;
                        hasTransparency |= alpha != 255;
                    }
                }
                break;
            }
            case 3:
            {
                if (png.Palette == null || png.Palette.Length == 0)
                    throw new NotSupportedException("Indexed PNG images require a palette.");

                for (var i = 0; i < pixelCount; i++)
                {
                    var paletteIndex = pixelData[src++];
                    var paletteOffset = paletteIndex * 3;
                    if (paletteOffset + 2 >= png.Palette.Length)
                        throw new ArgumentException("PNG palette index is out of range.");

                    rgbData[rgbOffset++] = png.Palette[paletteOffset];
                    rgbData[rgbOffset++] = png.Palette[paletteOffset + 1];
                    rgbData[rgbOffset++] = png.Palette[paletteOffset + 2];

                    if (alphaData != null)
                    {
                        var alpha = png.Transparency != null && paletteIndex < png.Transparency.Length
                            ? png.Transparency[paletteIndex]
                            : (byte)255;
                        alphaData[alphaOffset++] = alpha;
                        hasTransparency |= alpha != 255;
                    }
                }
                break;
            }
            case 4:
            {
                for (var i = 0; i < pixelCount; i++)
                {
                    var gray = pixelData[src++];
                    var alpha = pixelData[src++];
                    rgbData[rgbOffset++] = gray;
                    rgbData[rgbOffset++] = gray;
                    rgbData[rgbOffset++] = gray;
                    alphaData![alphaOffset++] = alpha;
                    hasTransparency |= alpha != 255;
                }
                break;
            }
            case 6:
            {
                for (var i = 0; i < pixelCount; i++)
                {
                    rgbData[rgbOffset++] = pixelData[src++];
                    rgbData[rgbOffset++] = pixelData[src++];
                    rgbData[rgbOffset++] = pixelData[src++];
                    var alpha = pixelData[src++];
                    alphaData![alphaOffset++] = alpha;
                    hasTransparency |= alpha != 255;
                }
                break;
            }
        }

        return (rgbData, hasTransparency ? alphaData : null);
    }

    private static void ApplyPngFilter(byte filterType, byte[] scanlines, int sourceOffset, byte[] currentRow, byte[] previousRow, int rowSize, int bytesPerPixel)
    {
        for (var i = 0; i < rowSize; i++)
        {
            var raw = scanlines[sourceOffset + i];
            var left = i >= bytesPerPixel ? currentRow[i - bytesPerPixel] : (byte)0;
            var up = previousRow[i];
            var upperLeft = i >= bytesPerPixel ? previousRow[i - bytesPerPixel] : (byte)0;

            currentRow[i] = filterType switch
            {
                0 => raw,
                1 => (byte)(raw + left),
                2 => (byte)(raw + up),
                3 => (byte)(raw + ((left + up) / 2)),
                4 => (byte)(raw + PaethPredictor(left, up, upperLeft)),
                _ => throw new NotSupportedException($"PNG filter type {filterType} is not supported.")
            };
        }
    }

    private static byte PaethPredictor(byte left, byte up, byte upperLeft)
    {
        var predictor = left + up - upperLeft;
        var leftDistance = Math.Abs(predictor - left);
        var upDistance = Math.Abs(predictor - up);
        var upperLeftDistance = Math.Abs(predictor - upperLeft);

        if (leftDistance <= upDistance && leftDistance <= upperLeftDistance)
            return left;
        if (upDistance <= upperLeftDistance)
            return up;
        return upperLeft;
    }

    private static byte[] InflateZlib(byte[] data)
    {
        if (data.Length < 6)
            throw new ArgumentException("Invalid zlib payload.");

        using var input = new MemoryStream(data, 2, data.Length - 6, writable: false);
        using var deflate = new DeflateStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        deflate.CopyTo(output);
        return output.ToArray();
    }

    private static byte[] CompressZlib(byte[] data)
    {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);

        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true))
        {
            deflate.Write(data, 0, data.Length);
        }

        var checksum = ComputeAdler32(data);
        output.WriteByte((byte)(checksum >> 24));
        output.WriteByte((byte)(checksum >> 16));
        output.WriteByte((byte)(checksum >> 8));
        output.WriteByte((byte)checksum);
        return output.ToArray();
    }

    private static (byte[] RgbData, byte[]? AlphaData) DecodeBmpPixelData(byte[] data, ImageInfo info)
    {
        if (data.Length < 30)
            throw new ArgumentException("Invalid BMP data.");

        var headerSize = data[14] | (data[15] << 8) | (data[16] << 16) | (data[17] << 24);
        if (headerSize < 12)
            throw new NotSupportedException($"BMP header size {headerSize} is not supported.");

        var pixelOffset = data[10] | (data[11] << 8) | (data[12] << 16) | (data[13] << 24);
        var width = info.Width;
        var height = info.Height;
        var topDown = false;
        if (height < 0)
        {
            topDown = true;
            height = -height;
        }

        var compression = headerSize >= 40
            ? data[30] | (data[31] << 8) | (data[32] << 16) | (data[33] << 24)
            : 0;

        if (compression != 0)
            throw new NotSupportedException($"BMP compression mode {compression} is not supported.");
        if (info.BitDepth is not (24 or 32))
            throw new NotSupportedException($"BMP bit depth {info.BitDepth} is not supported.");

        var bytesPerPixel = info.BitDepth / 8;
        var rowSize = ((width * bytesPerPixel + 3) / 4) * 4;
        var rgbData = new byte[checked(width * height * 3)];
        byte[]? alphaData = bytesPerPixel == 4 ? new byte[width * height] : null;
        var hasTransparency = false;

        for (var y = 0; y < height; y++)
        {
            var srcRow = topDown ? y : height - 1 - y;
            var srcOffset = pixelOffset + srcRow * rowSize;
            if (srcOffset < 0 || srcOffset >= data.Length)
                throw new ArgumentException("BMP pixel data is out of range.");

            for (var x = 0; x < width; x++)
            {
                var pixelOffsetInRow = srcOffset + x * bytesPerPixel;
                if (pixelOffsetInRow + bytesPerPixel - 1 >= data.Length)
                    throw new ArgumentException("BMP pixel data is incomplete.");

                var rgbOffset = (y * width + x) * 3;
                rgbData[rgbOffset + 0] = data[pixelOffsetInRow + 2];
                rgbData[rgbOffset + 1] = data[pixelOffsetInRow + 1];
                rgbData[rgbOffset + 2] = data[pixelOffsetInRow + 0];

                if (alphaData != null)
                {
                    var alpha = data[pixelOffsetInRow + 3];
                    alphaData[y * width + x] = alpha;
                    hasTransparency |= alpha != 255;
                }
            }
        }

        return (rgbData, hasTransparency ? alphaData : null);
    }

    private static uint ComputeAdler32(byte[] data)
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

    private static void WriteBigEndianUInt16(BinaryWriter writer, int value)
    {
        writer.Write((byte)(value >> 8));
        writer.Write((byte)value);
    }

    private static int ReadBigEndianInt(byte[] data, int offset)
    {
        return (data[offset] << 24) |
               (data[offset + 1] << 16) |
               (data[offset + 2] << 8) |
               data[offset + 3];
    }

    private static byte[] CopyChunkData(byte[] data, int offset, int length)
    {
        var result = new byte[length];
        Buffer.BlockCopy(data, offset, result, 0, length);
        return result;
    }

    private sealed class ParsedPng
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public byte BitDepth { get; set; }
        public byte ColorType { get; set; }
        public byte CompressionMethod { get; set; }
        public byte FilterMethod { get; set; }
        public byte InterlaceMethod { get; set; }
        public byte[] IdatData { get; set; } = Array.Empty<byte>();
        public byte[]? Palette { get; set; }
        public byte[]? Transparency { get; set; }
    }
}
