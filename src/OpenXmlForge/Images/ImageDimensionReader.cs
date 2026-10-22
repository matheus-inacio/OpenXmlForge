using System.Buffers.Binary;

namespace OpenXmlForge.Images;

/// <summary>
/// Extracts width and height from image headers (PNG, JPEG, BMP, GIF, WebP) without decoding pixel data.
/// Identifies format by magic bytes rather than file extension.
/// </summary>
public static class ImageDimensionReader
{
    private const int MaxAcceptableDimension = 100_000;

    private static ReadOnlySpan<byte> PngSignature => [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    public static ImageDimensions? Read(ReadOnlySpan<byte> header, out string? detectedFormat)
    {
        if (TryReadPng(header, out var dimensions))
        {
            detectedFormat = "png";
            return dimensions;
        }

        if (TryReadJpeg(header, out dimensions))
        {
            detectedFormat = "jpeg";
            return dimensions;
        }

        if (TryReadBmp(header, out dimensions))
        {
            detectedFormat = "bmp";
            return dimensions;
        }

        if (TryReadGif(header, out dimensions))
        {
            detectedFormat = "gif";
            return dimensions;
        }

        if (TryReadWebp(header, out dimensions))
        {
            detectedFormat = "webp";
            return dimensions;
        }

        detectedFormat = DetectFormatOnly(header);
        return null;
    }

    public static ImageDimensions? Read(ReadOnlySpan<byte> header)
        => Read(header, out _);

    private static string? DetectFormatOnly(ReadOnlySpan<byte> header)
    {
        if (header.Length >= 8 && header[..8].SequenceEqual(PngSignature))
        {
            return "png";
        }

        if (header.Length >= 2 && header[0] is 0xFF && header[1] is 0xD8)
        {
            return "jpeg";
        }

        if (header.Length >= 2 && header[0] is (byte)'B' && header[1] is (byte)'M')
        {
            return "bmp";
        }

        if (header.Length >= 6 && (header[..6].SequenceEqual("GIF87a"u8) || header[..6].SequenceEqual("GIF89a"u8)))
        {
            return "gif";
        }

        if (header.Length >= 12 && header[..4].SequenceEqual("RIFF"u8) && header.Slice(8, 4).SequenceEqual("WEBP"u8))
        {
            return "webp";
        }

        return null;
    }

    private static bool TryReadPng(ReadOnlySpan<byte> header, out ImageDimensions dimensions)
    {
        dimensions = default;

        if (header.Length < 24 ||
            !header[..8].SequenceEqual(PngSignature) ||
            !header.Slice(12, 4).SequenceEqual("IHDR"u8))
        {
            return false;
        }

        dimensions = new ImageDimensions(
            BinaryPrimitives.ReadInt32BigEndian(header.Slice(16, 4)),
            BinaryPrimitives.ReadInt32BigEndian(header.Slice(20, 4)));

        return IsValid(dimensions);
    }

    private static bool TryReadJpeg(ReadOnlySpan<byte> header, out ImageDimensions dimensions)
    {
        dimensions = default;

        if (header.Length < 4 || header[0] is not 0xFF || header[1] is not 0xD8)
        {
            return false;
        }

        var position = 2;

        while (position + 1 < header.Length)
        {
            if (header[position] is not 0xFF)
            {
                return false;
            }

            while (position < header.Length && header[position] is 0xFF)
            {
                position++;
            }

            if (position >= header.Length)
            {
                return false;
            }

            var marker = header[position];
            position++;

            // Markers without payload: TEM (0x01), RSTn (0xD0-0xD7), SOI (0xD8), EOI (0xD9).
            if (marker is 0x01 or >= 0xD0 and <= 0xD9)
            {
                continue;
            }

            // SOS: compressed data starts here; SOF should have appeared before.
            if (marker is 0xDA)
            {
                return false;
            }

            if (position + 1 >= header.Length)
            {
                return false;
            }

            var segmentLength = BinaryPrimitives.ReadUInt16BigEndian(header.Slice(position, 2));

            if (segmentLength < 2)
            {
                return false;
            }

            if (IsSofMarker(marker))
            {
                if (position + 7 > header.Length)
                {
                    return false;
                }

                // SOF layout: length (2) + precision (1) + height (2) + width (2).
                dimensions = new ImageDimensions(
                    BinaryPrimitives.ReadUInt16BigEndian(header.Slice(position + 5, 2)),
                    BinaryPrimitives.ReadUInt16BigEndian(header.Slice(position + 3, 2)));

                return IsValid(dimensions);
            }

            position += segmentLength;
        }

        return false;
    }

    private static bool IsSofMarker(byte marker)
        => marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC;

    private static bool TryReadBmp(ReadOnlySpan<byte> header, out ImageDimensions dimensions)
    {
        dimensions = default;

        if (header.Length < 26 || header[0] is not (byte)'B' || header[1] is not (byte)'M')
        {
            return false;
        }

        var dibHeaderSize = BinaryPrimitives.ReadUInt32LittleEndian(header.Slice(14, 4));

        var (width, height) = dibHeaderSize is 12
            ? (BinaryPrimitives.ReadInt16LittleEndian(header.Slice(18, 2)),
               Math.Abs((long)BinaryPrimitives.ReadInt16LittleEndian(header.Slice(20, 2))))
            : (BinaryPrimitives.ReadInt32LittleEndian(header.Slice(18, 4)),
               Math.Abs((long)BinaryPrimitives.ReadInt32LittleEndian(header.Slice(22, 4))));

        if (width > MaxAcceptableDimension || height > MaxAcceptableDimension)
        {
            return false;
        }

        dimensions = new ImageDimensions(width, (int)height);
        return IsValid(dimensions);
    }

    private static bool TryReadGif(ReadOnlySpan<byte> header, out ImageDimensions dimensions)
    {
        dimensions = default;

        if (header.Length < 10 ||
            (!header[..6].SequenceEqual("GIF87a"u8) && !header[..6].SequenceEqual("GIF89a"u8)))
        {
            return false;
        }

        dimensions = new ImageDimensions(
            BinaryPrimitives.ReadUInt16LittleEndian(header.Slice(6, 2)),
            BinaryPrimitives.ReadUInt16LittleEndian(header.Slice(8, 2)));

        return IsValid(dimensions);
    }

    private static bool TryReadWebp(ReadOnlySpan<byte> header, out ImageDimensions dimensions)
    {
        dimensions = default;

        if (header.Length < 30 ||
            !header[..4].SequenceEqual("RIFF"u8) ||
            !header.Slice(8, 4).SequenceEqual("WEBP"u8))
        {
            return false;
        }

        var chunkType = header.Slice(12, 4);

        if (chunkType.SequenceEqual("VP8X"u8))
        {
            var width = 1 + (header[24] | (header[25] << 8) | (header[26] << 16));
            var height = 1 + (header[27] | (header[28] << 8) | (header[29] << 16));
            dimensions = new ImageDimensions(width, height);
            return IsValid(dimensions);
        }

        if (chunkType.SequenceEqual("VP8 "u8) &&
            header[23] is 0x9D && header[24] is 0x01 && header[25] is 0x2A)
        {
            var width = BinaryPrimitives.ReadUInt16LittleEndian(header.Slice(26, 2)) & 0x3FFF;
            var height = BinaryPrimitives.ReadUInt16LittleEndian(header.Slice(28, 2)) & 0x3FFF;
            dimensions = new ImageDimensions(width, height);
            return IsValid(dimensions);
        }

        if (chunkType.SequenceEqual("VP8L"u8) && header[20] is 0x2F)
        {
            var bits = BinaryPrimitives.ReadUInt32LittleEndian(header.Slice(21, 4));
            var width = (int)(bits & 0x3FFF) + 1;
            var height = (int)((bits >> 14) & 0x3FFF) + 1;
            dimensions = new ImageDimensions(width, height);
            return IsValid(dimensions);
        }

        return false;
    }

    private static bool IsValid(ImageDimensions dimensions)
        => dimensions.Width is > 0 and <= MaxAcceptableDimension &&
           dimensions.Height is > 0 and <= MaxAcceptableDimension;
}
