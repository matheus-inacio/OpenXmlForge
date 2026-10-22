using System.Buffers.Binary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlForge.Models;
using Xunit;

namespace OpenXmlForge.Tests;

public class MultipleTablesImageRepetitionTests
{
    private const string OutsideParagraphExpectedText = "Repeat collection image outside table.";

    private sealed class NonSeekableMemoryStream(byte[] data) : Stream
    {
        private int _position;

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
        {
            var remaining = data.Length - _position;
            if (remaining <= 0) return 0;
            var toRead = Math.Min(remaining, count);
            Buffer.BlockCopy(data, _position, buffer, offset, toRead);
            _position += toRead;
            return toRead;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private static byte[] CreatePngBytes(int width, int height)
    {
        var bytes = new byte[24];
        ReadOnlySpan<byte> sig = [137, 80, 78, 71, 13, 10, 26, 10];
        sig.CopyTo(bytes);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(8, 4), 13);
        "IHDR"u8.CopyTo(bytes.AsSpan(12, 4));
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), width);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20, 4), height);
        return bytes;
    }

    private static MemoryStream CreateTemplate()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();

            var table1 = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("{{Items.Image}}"))))));

            var table2 = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("{{Items.Image}}")))),
                    new TableCell(new Paragraph(new Run(new Text("Repeat collection image in table cell."))))));

            var pOutside = new Paragraph(new Run(new Text("{{Items.Image}}     Repeat collection image outside table.")));

            mainPart.Document = new Document(new Body(table1, table2, pOutside));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task TestWithSeekableStream_RendersImagesInBothTablesAndDeletesTagOutsideTable()
    {
        using var template = CreateTemplate();
        var pngBytes = CreatePngBytes(100, 100);

        var data = new
        {
            Items = new[]
            {
                new { Image = new WordPicture(new MemoryStream(pngBytes), "png") }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var tables = body.Elements<Table>().ToList();
        Assert.Equal(2, tables.Count);

        var t1Drawings = tables[0].Descendants<Drawing>().ToList();
        Assert.Single(t1Drawings);

        var t2Drawings = tables[1].Descendants<Drawing>().ToList();
        Assert.Single(t2Drawings);

        var paragraphs = body.Elements<Paragraph>().ToList();
        var p = Assert.Single(paragraphs);
        Assert.Empty(p.Descendants<Drawing>());
        Assert.DoesNotContain("{{Items.Image}}", p.InnerText);
        Assert.Equal(OutsideParagraphExpectedText, p.InnerText.Trim());

        var imageParts = doc.MainDocumentPart!.ImageParts.ToList();
        var imagePart = Assert.Single(imageParts);
        using var imageStream = imagePart.GetStream();
        Assert.True(imageStream.Length > 0, "Image part should not be empty");
    }

    [Fact]
    public async Task TestWithNonSeekableStream_RendersImagesInBothTablesAndDeletesTagOutsideTable()
    {
        using var template = CreateTemplate();
        var pngBytes = CreatePngBytes(100, 100);

        var data = new
        {
            Items = new[]
            {
                new { Image = new WordPicture(new NonSeekableMemoryStream(pngBytes), "png") }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var tables = body.Elements<Table>().ToList();
        Assert.Equal(2, tables.Count);

        var t1Drawings = tables[0].Descendants<Drawing>().ToList();
        Assert.Single(t1Drawings);

        var t2Drawings = tables[1].Descendants<Drawing>().ToList();
        Assert.Single(t2Drawings);

        var paragraphs = body.Elements<Paragraph>().ToList();
        var p = Assert.Single(paragraphs);
        Assert.Empty(p.Descendants<Drawing>());
        Assert.DoesNotContain("{{Items.Image}}", p.InnerText);
        Assert.Equal(OutsideParagraphExpectedText, p.InnerText.Trim());

        var imageParts = doc.MainDocumentPart!.ImageParts.ToList();
        var imagePart = Assert.Single(imageParts);
        using var imageStream = imagePart.GetStream();
        Assert.True(imageStream.Length > 0, "Image part should not be empty");
    }

    [Fact]
    public async Task TestWithEmptyCollection_RemovesTagOutsideTable()
    {
        using var template = CreateTemplate();

        var data = new
        {
            Items = Array.Empty<object>()
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var p = body.Elements<Paragraph>().Single();

        Assert.DoesNotContain("{{Items.Image}}", p.InnerText);
        Assert.Equal(OutsideParagraphExpectedText, p.InnerText.Trim());
    }

    [Fact]
    public async Task TestWithMultipleItems_RepeatsImagesInBothTablesAndDeletesTagOutsideTable()
    {
        using var template = CreateTemplate();
        var pngBytes1 = CreatePngBytes(100, 100);
        var pngBytes2 = CreatePngBytes(80, 80);

        var data = new
        {
            Items = new[]
            {
                new { Image = new WordPicture(new NonSeekableMemoryStream(pngBytes1), "png") },
                new { Image = new WordPicture(new NonSeekableMemoryStream(pngBytes2), "png") }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var tables = body.Elements<Table>().ToList();

        Assert.Equal(2, tables[0].Elements<TableRow>().Count());
        Assert.Equal(2, tables[0].Descendants<Drawing>().Count());

        Assert.Equal(2, tables[1].Elements<TableRow>().Count());
        Assert.Equal(2, tables[1].Descendants<Drawing>().Count());

        var p = body.Elements<Paragraph>().Single();
        Assert.Empty(p.Descendants<Drawing>());
        Assert.DoesNotContain("{{Items.Image}}", p.InnerText);
        Assert.Equal(OutsideParagraphExpectedText, p.InnerText.Trim());
    }

    [Fact]
    public async Task SingleNonSeekableImage_RendersWithoutReusableCopy()
    {
        using var template = CreateParagraphTemplate("{{Image}}");
        var pngBytes = CreatePngBytes(100, 100);
        var data = new { Image = new WordPicture(new NonSeekableMemoryStream(pngBytes), "png") };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var imagePart = Assert.Single(doc.MainDocumentPart!.ImageParts);
        using var s = imagePart.GetStream();
        Assert.True(s.Length > 0);
    }

    [Fact]
    public async Task SameNonSeekableImageInTwoForeachBlocks_RendersBothCopies()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("@foreach(item in Items)"))),
                new Paragraph(new Run(new Text("{{item.Image}}"))),
                new Paragraph(new Run(new Text("@endforeach"))),
                new Paragraph(new Run(new Text("@foreach(item in Items)"))),
                new Paragraph(new Run(new Text("{{item.Image}}"))),
                new Paragraph(new Run(new Text("@endforeach")))));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        var pngBytes = CreatePngBytes(100, 100);
        var data = new
        {
            Items = new[]
            {
                new { Image = new WordPicture(new NonSeekableMemoryStream(pngBytes), "png") }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, data);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var imageParts = outDoc.MainDocumentPart!.ImageParts.ToList();
        var imagePart = Assert.Single(imageParts);
        using var imageStream = imagePart.GetStream();
        Assert.True(imageStream.Length > 0);
    }

    [Fact]
    public async Task SameNonSeekableImageInHeaderAndBody_ReusesPackagePart()
    {
        var template = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var headerPartId = mainPart.GetIdOfPart(headerPart);

            headerPart.Header = new Header(new Paragraph(new Run(new Text("{{Image}}"))));
            headerPart.Header.Save();

            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{{Image}}"))),
                new SectionProperties(
                    new HeaderReference { Id = headerPartId, Type = HeaderFooterValues.Default })));
            mainPart.Document.Save();
        }

        template.Position = 0;
        var pngBytes = CreatePngBytes(100, 100);
        var data = new { Image = new WordPicture(new NonSeekableMemoryStream(pngBytes), "png") };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var mainPartOutput = outDoc.MainDocumentPart!;
        var headerPartOutput = Assert.Single(mainPartOutput.HeaderParts);
        var bodyImagePart = Assert.Single(mainPartOutput.ImageParts);
        var headerImagePart = Assert.Single(headerPartOutput.ImageParts);

        Assert.Equal(bodyImagePart.Uri, headerImagePart.Uri);
        Assert.Single(mainPartOutput.Document.Body!.Descendants<Drawing>());
        Assert.Single(headerPartOutput.Header!.Descendants<Drawing>());
        using var imageStream = bodyImagePart.GetStream();
        Assert.Equal(pngBytes.Length, imageStream.Length);
    }

    private static MemoryStream CreateParagraphTemplate(string text)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }
}
