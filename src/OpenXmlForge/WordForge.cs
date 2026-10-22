using DocumentFormat.OpenXml.Packaging;
using OpenXmlForge.Core;

namespace OpenXmlForge;

/// <summary>
/// High-performance .NET 10 Word (DOCX) template engine.
/// </summary>
public static class WordForge
{
    /// <summary>
    /// Fills a DOCX template stream with <paramref name="value"/> and returns a new positioned <see cref="MemoryStream"/>.
    /// Does not require <paramref name="templateStream"/> to be seekable.
    /// </summary>
    public static async Task<MemoryStream> GenerateStreamAsync(
        Stream templateStream,
        object? value,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(templateStream);

        var initialCapacity = templateStream.CanSeek ? (int)(templateStream.Length * 1.15) : 0;
        var outputStream = initialCapacity > 0 ? new MemoryStream(initialCapacity) : new MemoryStream();

        if (templateStream.CanSeek)
        {
            templateStream.Position = 0;
        }

        await templateStream.CopyToAsync(outputStream, cancellationToken).ConfigureAwait(false);
        outputStream.Position = 0;

        using (var docx = WordprocessingDocument.Open(outputStream, true))
        {
            await TemplateEngine.ProcessDocumentAsync(docx, value, cancellationToken).ConfigureAwait(false);
        }

        outputStream.Position = 0;
        return outputStream;
    }

    /// <summary>
    /// Synchronously fills a DOCX template stream with <paramref name="value"/>
    /// and returns a new positioned <see cref="MemoryStream"/>.
    /// </summary>
    public static MemoryStream GenerateStream(Stream templateStream, object? value)
    {
        ArgumentNullException.ThrowIfNull(templateStream);

        var initialCapacity = templateStream.CanSeek ? (int)(templateStream.Length * 1.15) : 0;
        var outputStream = initialCapacity > 0 ? new MemoryStream(initialCapacity) : new MemoryStream();

        if (templateStream.CanSeek)
        {
            templateStream.Position = 0;
        }

        templateStream.CopyTo(outputStream);
        outputStream.Position = 0;

        using (var docx = WordprocessingDocument.Open(outputStream, true))
        {
            TemplateEngine.ProcessDocumentSync(docx, value);
        }

        outputStream.Position = 0;
        return outputStream;
    }

    public static void SaveAsByTemplate(string path, string templatePath, object? value)
    {
        using var outputStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        SaveAsByTemplate(outputStream, templateStream, value);
    }

    public static void SaveAsByTemplate(string path, byte[] templateBytes, object? value)
    {
        using var outputStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        SaveAsByTemplate(outputStream, templateBytes, value);
    }

    public static void SaveAsByTemplate(string path, Stream templateStream, object? value)
    {
        using var outputStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        SaveAsByTemplate(outputStream, templateStream, value);
    }

    public static void SaveAsByTemplate(this Stream stream, string templatePath, object? value)
    {
        using var templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        SaveAsByTemplate(stream, templateStream, value);
    }

    public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object? value)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(templateBytes);

        if (stream is { CanSeek: true, CanRead: true, CanWrite: true })
        {
            var startPos = stream.Position;
            stream.Write(templateBytes, 0, templateBytes.Length);
            stream.Position = startPos;

            using var docx = WordprocessingDocument.Open(stream, true);
            TemplateEngine.ProcessDocumentSync(docx, value);
            return;
        }

        using var ms = new MemoryStream((int)(templateBytes.Length * 1.15));
        ms.Write(templateBytes, 0, templateBytes.Length);
        ms.Position = 0;

        using (var docx = WordprocessingDocument.Open(ms, true))
        {
            TemplateEngine.ProcessDocumentSync(docx, value);
        }

        ms.Position = 0;
        ms.CopyTo(stream);
    }

    public static void SaveAsByTemplate(this Stream stream, Stream templateStream, object? value)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(templateStream);

        if (templateStream.CanSeek)
        {
            templateStream.Position = 0;
        }

        if (stream.CanSeek && stream.CanRead && stream.CanWrite)
        {
            var startPos = stream.Position;
            templateStream.CopyTo(stream);
            stream.Position = startPos;

            using var docx = WordprocessingDocument.Open(stream, true);
            TemplateEngine.ProcessDocumentSync(docx, value);
            return;
        }

        using var generated = GenerateStream(templateStream, value);
        generated.CopyTo(stream);
    }

    public static async Task SaveAsByTemplateAsync(
        string path,
        string templatePath,
        object? value,
        CancellationToken cancellationToken = default)
    {
        await using var outputStream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.ReadWrite,
            FileShare.None,
            bufferSize: 4096,
            useAsync: true);
        await using var templateStream = new FileStream(
            templatePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite,
            bufferSize: 4096,
            useAsync: true);

        await SaveAsByTemplateAsync(outputStream, templateStream, value, cancellationToken).ConfigureAwait(false);
    }

    public static async Task SaveAsByTemplateAsync(
        string path,
        byte[] templateBytes,
        object? value,
        CancellationToken cancellationToken = default)
    {
        await using var outputStream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.ReadWrite,
            FileShare.None,
            bufferSize: 4096,
            useAsync: true);

        await SaveAsByTemplateAsync(outputStream, templateBytes, value, cancellationToken).ConfigureAwait(false);
    }

    public static async Task SaveAsByTemplateAsync(
        string path,
        Stream templateStream,
        object? value,
        CancellationToken cancellationToken = default)
    {
        await using var outputStream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.ReadWrite,
            FileShare.None,
            bufferSize: 4096,
            useAsync: true);

        await SaveAsByTemplateAsync(outputStream, templateStream, value, cancellationToken).ConfigureAwait(false);
    }

    public static async Task SaveAsByTemplateAsync(
        this Stream stream,
        string templatePath,
        object? value,
        CancellationToken cancellationToken = default)
    {
        await using var templateStream = new FileStream(
            templatePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite,
            bufferSize: 4096,
            useAsync: true);

        await SaveAsByTemplateAsync(stream, templateStream, value, cancellationToken).ConfigureAwait(false);
    }

    public static async Task SaveAsByTemplateAsync(
        this Stream stream,
        byte[] templateBytes,
        object? value,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(templateBytes);

        if (stream is { CanSeek: true, CanRead: true, CanWrite: true })
        {
            var startPos = stream.Position;
            await stream.WriteAsync(templateBytes, cancellationToken).ConfigureAwait(false);
            stream.Position = startPos;

            using var docx = WordprocessingDocument.Open(stream, true);
            await TemplateEngine.ProcessDocumentAsync(docx, value, cancellationToken).ConfigureAwait(false);
            return;
        }

        await using var ms = new MemoryStream((int)(templateBytes.Length * 1.15));
        await ms.WriteAsync(templateBytes, cancellationToken).ConfigureAwait(false);
        ms.Position = 0;

        using (var docx = WordprocessingDocument.Open(ms, true))
        {
            await TemplateEngine.ProcessDocumentAsync(docx, value, cancellationToken).ConfigureAwait(false);
        }

        ms.Position = 0;
        await ms.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    public static async Task SaveAsByTemplateAsync(
        this Stream stream,
        Stream templateStream,
        object? value,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(templateStream);

        if (templateStream.CanSeek)
        {
            templateStream.Position = 0;
        }

        if (stream is { CanSeek: true, CanRead: true, CanWrite: true })
        {
            var startPos = stream.Position;
            await templateStream.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            stream.Position = startPos;

            using var docx = WordprocessingDocument.Open(stream, true);
            await TemplateEngine.ProcessDocumentAsync(docx, value, cancellationToken).ConfigureAwait(false);
            return;
        }

        await using var generated = await GenerateStreamAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
        await generated.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }
}
