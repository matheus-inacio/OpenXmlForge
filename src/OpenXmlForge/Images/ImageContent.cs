using DocumentFormat.OpenXml.Packaging;
using OpenXmlForge.Models;

namespace OpenXmlForge.Images;

/// <summary>
/// Holds prepared image stream/bytes along with natural dimensions read from its header.
/// </summary>
internal sealed class ImageContent : IDisposable, IAsyncDisposable
{
    private const int HeaderBufferSize = 64 * 1024;

    private readonly Stream? _stream;
    private readonly byte[]? _bytes;
    private readonly bool _ownsStream;

    private ImageContent(
        Stream? stream,
        byte[]? bytes,
        ImageDimensions? dimensions,
        string? detectedFormat,
        bool ownsStream)
    {
        _stream = stream;
        _bytes = bytes;
        Dimensions = dimensions;
        DetectedFormat = detectedFormat;
        _ownsStream = ownsStream;
    }

    public ImageDimensions? Dimensions { get; }

    public string? DetectedFormat { get; }

    public static async ValueTask<ImageContent> PrepareAsync(
        WordPicture picture,
        CancellationToken cancellationToken = default)
    {
        if (picture.Bytes is { } bytes)
        {
            string? detectedFmt = null;
            var dims = picture.HasExplicitDimensions
                ? new ImageDimensions((int)picture.Width!.Value, (int)picture.Height!.Value)
                : ImageDimensionReader.Read(bytes, out detectedFmt);
            return new ImageContent(null, bytes, dims, detectedFmt, ownsStream: false);
        }

        if (!string.IsNullOrWhiteSpace(picture.Path))
        {
            var fileStream = new FileStream(
                picture.Path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite,
                bufferSize: 4096,
                useAsync: true);

            return await FromStreamAsync(
                fileStream,
                picture.HasExplicitDimensions
                    ? new ImageDimensions((int)picture.Width!.Value, (int)picture.Height!.Value)
                    : null,
                ownsStream: true,
                cancellationToken).ConfigureAwait(false);
        }

        if (picture.Stream is { } stream && stream != Stream.Null)
        {
            return await FromStreamAsync(
                stream,
                picture.HasExplicitDimensions
                    ? new ImageDimensions((int)picture.Width!.Value, (int)picture.Height!.Value)
                    : null,
                ownsStream: !picture.LeaveStreamOpen,
                cancellationToken).ConfigureAwait(false);
        }

        return new ImageContent(null, [], null, null, ownsStream: false);
    }

    public static ImageContent PrepareSync(WordPicture picture)
    {
        if (picture.Bytes is { } bytes)
        {
            string? detectedFmt = null;
            var dims = picture.HasExplicitDimensions
                ? new ImageDimensions((int)picture.Width!.Value, (int)picture.Height!.Value)
                : ImageDimensionReader.Read(bytes, out detectedFmt);
            return new ImageContent(null, bytes, dims, detectedFmt, ownsStream: false);
        }

        if (!string.IsNullOrWhiteSpace(picture.Path))
        {
            var fileStream = new FileStream(
                picture.Path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

            return FromStreamSync(
                fileStream,
                picture.HasExplicitDimensions
                    ? new ImageDimensions((int)picture.Width!.Value, (int)picture.Height!.Value)
                    : null,
                ownsStream: true);
        }

        if (picture.Stream is { } stream && stream != Stream.Null)
        {
            return FromStreamSync(
                stream,
                picture.HasExplicitDimensions
                    ? new ImageDimensions((int)picture.Width!.Value, (int)picture.Height!.Value)
                    : null,
                ownsStream: !picture.LeaveStreamOpen);
        }

        return new ImageContent(null, [], null, null, ownsStream: false);
    }

    public void WriteTo(ImagePart imagePart)
    {
        if (_bytes is not null)
        {
            using var ms = new MemoryStream(_bytes, writable: false);
            imagePart.FeedData(ms);
            return;
        }

        if (_stream is null)
        {
            return;
        }

        if (_stream.CanSeek)
        {
            _stream.Position = 0;
        }

        imagePart.FeedData(_stream);
    }

    public void Dispose()
    {
        if (_ownsStream)
        {
            _stream?.Dispose();
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_ownsStream && _stream is not null)
        {
            await _stream.DisposeAsync().ConfigureAwait(false);
        }
    }

    private static async ValueTask<ImageContent> FromStreamAsync(
        Stream stream,
        ImageDimensions? explicitDimensions,
        bool ownsStream,
        CancellationToken cancellationToken)
    {
        if (explicitDimensions is not null && stream.CanSeek)
        {
            stream.Position = 0;
            return new ImageContent(stream, null, explicitDimensions, null, ownsStream);
        }

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        var header = new byte[HeaderBufferSize];
        var bytesRead = await stream.ReadAtLeastAsync(
            header,
            header.Length,
            throwOnEndOfStream: false,
            cancellationToken).ConfigureAwait(false);

        var dims = ImageDimensionReader.Read(header.AsSpan(0, bytesRead), out var detectedFormat);
        if (explicitDimensions is not null)
        {
            dims = explicitDimensions;
        }

        if (stream.CanSeek)
        {
            stream.Position = 0;
            return new ImageContent(stream, null, dims, detectedFormat, ownsStream);
        }

        return new ImageContent(
            new PrefixedStream(header, bytesRead, stream),
            null,
            dims,
            detectedFormat,
            ownsStream);
    }

    private static ImageContent FromStreamSync(
        Stream stream,
        ImageDimensions? explicitDimensions,
        bool ownsStream)
    {
        if (explicitDimensions is not null && stream.CanSeek)
        {
            stream.Position = 0;
            return new ImageContent(stream, null, explicitDimensions, null, ownsStream);
        }

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        var header = new byte[HeaderBufferSize];
        var bytesRead = stream.ReadAtLeast(
            header,
            header.Length,
            throwOnEndOfStream: false);

        var dims = ImageDimensionReader.Read(header.AsSpan(0, bytesRead), out var detectedFormat);

        if (explicitDimensions is not null)
        {
            dims = explicitDimensions;
        }

        if (stream.CanSeek)
        {
            stream.Position = 0;
            return new ImageContent(stream, null, dims, detectedFormat, ownsStream);
        }

        return new ImageContent(
            new PrefixedStream(header, bytesRead, stream),
            null,
            dims,
            detectedFormat,
            ownsStream);
    }
}
