using System.Collections.Frozen;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlForge.Models;

public enum WordPictureWrappingType
{
    Inline,
    Anchor,
}

public sealed class WordPicture
{
    public const long DefaultDimensionPixels = 400;
    public const long EmusPerPixel = 9525;

    public static FrozenSet<string> SupportedExtensions { get; } = new[]
    {
        "bmp", "emf", "gif", "ico", "icon", "jpg", "jpeg", "pcx", "png", "svg", "tif", "tiff", "webp", "wmf",
    }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    private string? _extension;

    public WordPicture()
    {
    }

    public WordPicture(string path)
    {
        Path = path;
    }

    public WordPicture(string path, long width, long height)
    {
        Path = path;
        Width = width;
        Height = height;
    }

    public WordPicture(Stream stream, string? extension = null)
    {
        Stream = stream;
        Extension = extension;
    }

    public WordPicture(Stream stream, string? extension, long width, long height)
    {
        Stream = stream;
        Extension = extension;
        Width = width;
        Height = height;
    }

    public WordPicture(byte[] bytes, string? extension = null)
    {
        Bytes = bytes;
        Extension = extension;
    }

    public WordPicture(byte[] bytes, string? extension, long width, long height)
    {
        Bytes = bytes;
        Extension = extension;
        Width = width;
        Height = height;
    }

    public WordPicture(byte[] bytes, Extension extension, long width = DefaultDimensionPixels, long height = DefaultDimensionPixels)
    {
        Bytes = bytes;
        Extension = extension.ToString();
        Width = width;
        Height = height;
    }

    public WordPictureWrappingType WrappingType { get; set; } = WordPictureWrappingType.Inline;

    public bool BehindDoc { get; set; }

    public bool AllowOverlap { get; set; }

    public long HorizontalPositionOffset { get; set; }

    public long VerticalPositionOffset { get; set; }

    public string? Path { get; set; }

    public byte[]? Bytes { get; set; }

    public Stream? Stream { get; set; }

    /// <summary>
    /// When false (default), <see cref="Stream"/> is automatically disposed when template generation completes.
    /// </summary>
    public bool LeaveStreamOpen { get; set; }

    public string? Extension
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(_extension))
            {
                return NormalizeExtension(_extension);
            }

            if (!string.IsNullOrWhiteSpace(Path))
            {
                return NormalizeExtension(System.IO.Path.GetExtension(Path));
            }

            return string.Empty;
        }
        set => _extension = value;
    }

    /// <summary>
    /// Explicit width in pixels. When null or &lt;= 0, natural width from the image header is used
    /// (clamped to the containing table cell width if applicable).
    /// </summary>
    public long? Width { get; set; }

    /// <summary>
    /// Explicit height in pixels. When null or &lt;= 0, natural height from the image header is used
    /// (scaled proportionally to the containing table cell width if applicable).
    /// </summary>
    public long? Height { get; set; }

    internal bool HasExplicitDimensions => Width is > 0 && Height is > 0;

    internal static string NormalizeExtension(string? extension)
        => string.IsNullOrWhiteSpace(extension)
            ? string.Empty
            : extension.AsSpan().Trim().TrimStart('.').ToString().ToLowerInvariant();

    internal PartTypeInfo GetImagePartType(string? detectedExtension = null)
    {
        var ext = !string.IsNullOrWhiteSpace(Extension)
            ? Extension
            : NormalizeExtension(detectedExtension);

        return ext switch
        {
            "bmp" => ImagePartType.Bmp,
            "emf" => ImagePartType.Emf,
            "gif" => ImagePartType.Gif,
            "ico" or "icon" => ImagePartType.Icon,
            "jpg" or "jpeg" => ImagePartType.Jpeg,
            "pcx" => ImagePartType.Pcx,
            "png" => ImagePartType.Png,
            "svg" => ImagePartType.Svg,
            "tif" or "tiff" => ImagePartType.Tiff,
            "webp" => new PartTypeInfo("image/webp", ".webp"),
            "wmf" => ImagePartType.Wmf,
            _ => throw new NotSupportedException($"Image extension '{_extension ?? ext}' is not supported."),
        };
    }
}
