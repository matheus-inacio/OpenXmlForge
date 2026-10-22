namespace OpenXmlForge.Images;

/// <summary>
/// Read-only stream that yields an already-read in-memory header prefix followed by the remainder
/// of a non-seekable underlying stream.
/// </summary>
internal sealed class PrefixedStream(byte[] prefix, int prefixLength, Stream source) : Stream
{
    private int _prefixPosition;

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count)
        => Read(buffer.AsSpan(offset, count));

    public override int Read(Span<byte> buffer)
    {
        var readFromPrefix = ReadFromPrefix(buffer);
        return readFromPrefix > 0 ? readFromPrefix : source.Read(buffer);
    }

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        => ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();

    public override async ValueTask<int> ReadAsync(
        Memory<byte> buffer,
        CancellationToken cancellationToken = default)
    {
        var readFromPrefix = ReadFromPrefix(buffer.Span);

        return readFromPrefix > 0
            ? readFromPrefix
            : await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
    }

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override async ValueTask DisposeAsync()
    {
        await source.DisposeAsync().ConfigureAwait(false);
        await base.DisposeAsync().ConfigureAwait(false);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            source.Dispose();
        }

        base.Dispose(disposing);
    }

    private int ReadFromPrefix(Span<byte> buffer)
    {
        var remaining = prefixLength - _prefixPosition;

        if (remaining <= 0 || buffer.IsEmpty)
        {
            return 0;
        }

        var count = Math.Min(remaining, buffer.Length);
        prefix.AsSpan(_prefixPosition, count).CopyTo(buffer);
        _prefixPosition += count;

        return count;
    }
}
