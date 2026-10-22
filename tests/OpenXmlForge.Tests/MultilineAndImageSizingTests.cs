using System.Buffers.Binary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlForge.Images;
using OpenXmlForge.Models;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OpenXmlForge.Tests;

public class MultilineAndImageSizingTests
{
    private const long EmusPerPixelAt96Dpi = 9_525;

    [Fact]
    public async Task GenerateStreamAsync_MultilineTextInParagraph_ReplacesWithBreakElementsWithoutLeaking()
    {
        using var template = CreateParagraphTemplate("{{Company.LegalName}}");
        var data = new Dictionary<string, object?>
        {
            ["Company.LegalName"] = "Line 1\nLine 2\r\nLine 3"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        Assert.Equal("Line 1Line 2Line 3", body.InnerText);
        Assert.Equal(2, body.Descendants<Break>().Count());
    }

    [Fact]
    public async Task GenerateStreamAsync_MultilineTextInTextBoxWithAlternateContent_DoesNotDuplicateLinesIntoOuterRunOrAcrossBranches()
    {
        // Previously MiniWord's p.Descendants<Run>() matched runs inside TextBoxContent from the outer Paragraph
        // and appended <w:br/> + lines into the outer body Run, AND duplicated lines across mc:Choice + mc:Fallback.
        using var template = CreateTextBoxWithAlternateContentTemplate("{{Document.Notes}}");
        var data = new Dictionary<string, object?>
        {
            ["Document.Notes"] = "Line 1\nLine 2\nLine 3"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        using var doc = WordprocessingDocument.Open(result, false);
        var outerParagraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        var outerRun = outerParagraph.Elements<Run>().Single();

        // Outer run must NOT have direct Text or Break children leaked from the inner TextBoxContent!
        Assert.Empty(outerRun.Elements<Text>());
        Assert.Empty(outerRun.Elements<Break>());

        // Each TextBoxContent (Choice and Fallback) must have exactly 2 Breaks and "Line 1Line 2Line 3"
        var textBoxes = outerParagraph.Descendants<TextBoxContent>().ToList();
        Assert.Equal(2, textBoxes.Count);

        foreach (var textBox in textBoxes)
        {
            var innerParagraph = textBox.Elements<Paragraph>().Single();
            Assert.Equal(2, innerParagraph.Descendants<Break>().Count());
            Assert.Equal(["Line 1", "Line 2", "Line 3"], innerParagraph.Descendants<Text>().Select(t => t.Text));
        }
    }

    [Fact]
    public async Task GenerateStreamAsync_WithImageStreamInTableRow_EmbedsExactBytesAndDisposesStream()
    {
        // WordPicture natively accepts Stream, works with non-seekable streams, and disposes owned streams.
        using var template = CreateSingleTableCellTemplate("{{Products.Image}}", widthTwips: 2_880);
        byte[] imgBytes = [1, 2, 3, 4, 5, 6];
        var imgStream = new TrackingMemoryStream(imgBytes);

        var data = new Dictionary<string, object?>
        {
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Image"] = new WordPicture
                    {
                        Stream = imgStream,
                        Extension = "png",
                        Width = 80,
                        Height = 80
                    }
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        Assert.True(imgStream.IsDisposed);

        using var doc = WordprocessingDocument.Open(result, false);
        var imagePart = Assert.Single(doc.MainDocumentPart!.ImageParts);
        using var partStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        await partStream.CopyToAsync(copy);
        Assert.Equal(imgBytes, copy.ToArray());
    }

    [Fact]
    public async Task GenerateStreamAsync_WhenImageTagIsAbsentFromTemplate_StillDisposesOwnedStream()
    {
        using var template = CreateParagraphTemplate("{{Invoice.Number}}");
        var unusedStream = new TrackingMemoryStream(CreatePngBytes(120, 60));

        var data = new Dictionary<string, object?>
        {
            ["Invoice.Number"] = "INV-10",
            ["Company.Logo"] = new WordPicture
            {
                Stream = unusedStream,
                Extension = "png"
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        Assert.True(unusedStream.IsDisposed);
    }

    [Fact]
    public async Task GenerateStreamAsync_WithNonSeekableDownloadStreamLargerThanHeaderBuffer_PreservesAllBytesAndClampsToCellWidth()
    {
        // Non-seekable stream (like HTTP/object storage response stream) larger than 64KB header prefix
        // must have its dimensions read from the header AND the full payload written without corruption.
        using var template = CreateSingleTableCellTemplate("{{Products.Image}}", widthTwips: 1_656); // 1656 - 216 margins = 1440 twips = 96px
        var payload = CreatePngWithLargePayload(width: 400, height: 100, extraPayloadBytes: 96 * 1024);
        var downloadStream = new NonSeekableDownloadStream(payload);

        var data = new Dictionary<string, object?>
        {
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Image"] = new WordPicture
                    {
                        Stream = downloadStream,
                        Extension = "png"
                    }
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        Assert.True(downloadStream.IsDisposed);
        var (cx, cy) = ReadFirstDrawingExtent(result);
        Assert.Equal(96 * EmusPerPixelAt96Dpi, cx);
        Assert.Equal(24 * EmusPerPixelAt96Dpi, cy);

        result.Position = 0;
        using var doc = WordprocessingDocument.Open(result, false);
        var imagePart = Assert.Single(doc.MainDocumentPart!.ImageParts);
        using var partStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        await partStream.CopyToAsync(copy);
        Assert.Equal(payload, copy.ToArray());
    }

    [Fact]
    public async Task GenerateStreamAsync_WhenImageIsSmallerThanTableCell_PreservesNaturalDimensionsWithoutStretchingTo400x400()
    {
        // Old MiniWord defaulted Width/Height to 400x400px whenever not specified,
        // stretching small logos/icons into giant squares.
        using var template = CreateSingleTableCellTemplate("{{Products.Image}}", widthTwips: 2_880); // 192px
        var pngStream = new MemoryStream(CreatePngBytes(width: 80, height: 40));

        var data = new Dictionary<string, object?>
        {
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Image"] = new WordPicture
                    {
                        Stream = pngStream,
                        Extension = "png"
                    }
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        var (cx, cy) = ReadFirstDrawingExtent(result);
        Assert.Equal(80 * EmusPerPixelAt96Dpi, cx);
        Assert.Equal(40 * EmusPerPixelAt96Dpi, cy);
    }

    [Fact]
    public async Task GenerateStreamAsync_WhenImageHasFractionalGoogleDocsWidth_ClampsProportionally()
    {
        // Google Docs exports fractional dxa like "1656.0" and cell margins "72.0"
        using var template = CreateSingleTableCellTemplate(
            "{{Products.Image}}",
            widthRaw: "1656.0",
            widthType: TableWidthUnitValues.Dxa,
            leftMarginRaw: "72.0",
            rightMarginRaw: "72.0");

        var pngStream = new MemoryStream(CreatePngBytes(width: 500, height: 500));
        var data = new Dictionary<string, object?>
        {
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Image"] = new WordPicture
                    {
                        Stream = pngStream,
                        Extension = "png"
                    }
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        // 1656 - 72 - 72 = 1512 twips => 1512 / 15 = 100px
        var (cx, cy) = ReadFirstDrawingExtent(result);
        Assert.Equal(100 * EmusPerPixelAt96Dpi, cx);
        Assert.Equal(100 * EmusPerPixelAt96Dpi, cy);
    }

    [Fact]
    public async Task GenerateStreamAsync_WhenTableCellUsesPercentageWidth_ResolvesAgainstAvailablePageWidth()
    {
        // Page usable width = 9000 twips (600px), table 100% (5000 pct), cell 25% (1250 pct) => 2250 twips = 150px
        using var template = CreateSingleTableCellTemplate(
            "{{Products.Image}}",
            widthRaw: "1250",
            widthType: TableWidthUnitValues.Pct,
            leftMarginRaw: "0",
            rightMarginRaw: "0",
            tableWidthRaw: "5000",
            tableWidthType: TableWidthUnitValues.Pct,
            pageWidthTwips: 10_440,
            pageLeftMarginTwips: 720,
            pageRightMarginTwips: 720);

        var pngStream = new MemoryStream(CreatePngBytes(width: 600, height: 300));
        var data = new Dictionary<string, object?>
        {
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Image"] = new WordPicture
                    {
                        Stream = pngStream,
                        Extension = "png"
                    }
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, data);

        var (cx, cy) = ReadFirstDrawingExtent(result);
        Assert.Equal(150 * EmusPerPixelAt96Dpi, cx);
        Assert.Equal(75 * EmusPerPixelAt96Dpi, cy);
    }

    [Fact]
    public async Task GenerateStreamAsync_WhenTableCellReliesOnTableGridAndGridSpan_SumsSpannedColumns()
    {
        // First cell spans 1 column (900 twips), second cell has "{{Products.Image}}" and spans 2 columns (900 + 1200 = 2100 twips)
        // Minus default cell margins (108 + 108 = 216 twips) => 1884 twips => 125px
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var tableGrid = new TableGrid(
                new GridColumn { Width = "900" },
                new GridColumn { Width = "900" },
                new GridColumn { Width = "1200" });

            var firstCell = new TableCell(new Paragraph(new Run(new Text("Col 1"))));
            var secondCell = new TableCell(
                new TableCellProperties(new GridSpan { Val = 2 }),
                new Paragraph(new Run(new Text("{{Products.Image}}"))));

            var table = new Table(tableGrid, new TableRow(firstCell, secondCell));
            mainPart.Document = new Document(new Body(table));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        var pngStream = new MemoryStream(CreatePngBytes(width: 500, height: 250));
        var data = new Dictionary<string, object?>
        {
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Image"] = new WordPicture
                    {
                        Stream = pngStream,
                        Extension = "png"
                    }
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, data);

        var (cx, cy) = ReadFirstDrawingExtent(result);
        Assert.Equal(125 * EmusPerPixelAt96Dpi, cx);
        Assert.Equal(63 * EmusPerPixelAt96Dpi, cy);
    }

    [Fact]
    public async Task GenerateStreamAsync_WhenLogoIsInHeaderTableCell_ClampsToHeaderCellWidth()
    {
        // Header cell 1500 twips - 216 default margins = 1284 twips = 85px
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var headerPartId = mainPart.GetIdOfPart(headerPart);

            var cell = new TableCell(
                new TableCellProperties(new TableCellWidth { Width = "1500", Type = TableWidthUnitValues.Dxa }),
                new Paragraph(new Run(new Text("{{Company.Logo}}"))));
            headerPart.Header = new Header(new Table(new TableRow(cell)));
            headerPart.Header.Save();

            var sectionProps = new SectionProperties(new HeaderReference { Id = headerPartId, Type = HeaderFooterValues.Default });
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Document body"))), sectionProps));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        var data = new Dictionary<string, object?>
        {
            ["Company.Logo"] = new WordPicture
            {
                Stream = new MemoryStream(CreatePngBytes(width: 340, height: 170)),
                Extension = "png"
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, data);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var header = Assert.Single(outDoc.MainDocumentPart!.HeaderParts).Header;
        var extent = header.Descendants<Drawing>().Single().Descendants<A.Wordprocessing.Extent>().Single();
        Assert.Equal(85 * EmusPerPixelAt96Dpi, extent.Cx!.Value);
        Assert.Equal(43 * EmusPerPixelAt96Dpi, extent.Cy!.Value);
    }

    [Fact]
    public void ImageDimensionReader_ParsesPngJpegAndBmpIncludingIntMinValueHeight()
    {
        // PNG
        var pngDims = ImageDimensionReader.Read(CreatePngBytes(320, 180));
        Assert.Equal(new ImageDimensions(320, 180), pngDims);

        // JPEG with APP0 marker before SOF0
        var jpegBytes = new byte[]
        {
            0xFF, 0xD8,
            0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00,
            0xFF, 0xC0, 0x00, 0x11, 0x08,
            0x00, 0xB4, // height = 180
            0x01, 0x40, // width = 320
            0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
            0xFF, 0xD9
        };
        Assert.Equal(new ImageDimensions(320, 180), ImageDimensionReader.Read(jpegBytes));

        // BMP with int.MinValue height (must not throw OverflowException on Math.Abs(int.MinValue))
        var bmpInvalid = new byte[26];
        bmpInvalid[0] = (byte)'B';
        bmpInvalid[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(bmpInvalid.AsSpan(14, 4), 40);
        BinaryPrimitives.WriteInt32LittleEndian(bmpInvalid.AsSpan(18, 4), 120);
        BinaryPrimitives.WriteInt32LittleEndian(bmpInvalid.AsSpan(22, 4), int.MinValue);
        Assert.Null(ImageDimensionReader.Read(bmpInvalid));
    }

    private static (long Cx, long Cy) ReadFirstDrawingExtent(Stream docStream)
    {
        docStream.Position = 0;
        using var doc = WordprocessingDocument.Open(docStream, false);
        var extent = doc.MainDocumentPart!.Document.Body!
            .Descendants<Drawing>().Single()
            .Descendants<A.Wordprocessing.Extent>().Single();
        return (extent.Cx!.Value, extent.Cy!.Value);
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

    private static MemoryStream CreateTextBoxWithAlternateContentTemplate(string text)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var choice = new AlternateContentChoice(
                new Drawing(
                    new Dw.Inline(
                        new A.Graphic(
                            new A.GraphicData(
                                new Wps.WordprocessingShape(
                                    new Wps.TextBoxInfo2(
                                        new TextBoxContent(new Paragraph(new Run(new Text(text)))))))
                            {
                                Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                            }))))
            {
                Requires = "wps",
            };

            var fallback = new AlternateContentFallback(
                new Picture(
                    new V.Rectangle(
                        new V.TextBox(
                            new TextBoxContent(new Paragraph(new Run(new Text(text))))))));

            var alternateContent = new AlternateContent(choice, fallback);
            mainPart.Document = new Document(new Body(new Paragraph(new Run(alternateContent))));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateSingleTableCellTemplate(
        string text,
        int? widthTwips = null,
        string? widthRaw = null,
        TableWidthUnitValues? widthType = null,
        string? leftMarginRaw = null,
        string? rightMarginRaw = null,
        string? tableWidthRaw = null,
        TableWidthUnitValues? tableWidthType = null,
        uint? pageWidthTwips = null,
        int? pageLeftMarginTwips = null,
        int? pageRightMarginTwips = null)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var cellProps = new TableCellProperties();
            var resolvedWidth = widthRaw ?? widthTwips?.ToString();

            if (resolvedWidth is not null)
            {
                cellProps.Append(new TableCellWidth
                {
                    Width = resolvedWidth,
                    Type = widthType ?? TableWidthUnitValues.Dxa
                });
            }

            if (leftMarginRaw is not null || rightMarginRaw is not null)
            {
                var cellMargin = new TableCellMargin();
                if (leftMarginRaw is not null)
                {
                    cellMargin.LeftMargin = new LeftMargin { Width = leftMarginRaw, Type = TableWidthUnitValues.Dxa };
                }

                if (rightMarginRaw is not null)
                {
                    cellMargin.RightMargin = new RightMargin { Width = rightMarginRaw, Type = TableWidthUnitValues.Dxa };
                }

                cellProps.Append(cellMargin);
            }

            var cell = new TableCell(cellProps, new Paragraph(new Run(new Text(text))));
            var table = new Table();
            if (tableWidthRaw is not null)
            {
                table.Append(new TableProperties(new TableWidth
                {
                    Width = tableWidthRaw,
                    Type = tableWidthType ?? TableWidthUnitValues.Dxa
                }));
            }

            table.Append(new TableRow(cell));

            var body = new Body(table);
            if (pageWidthTwips.HasValue)
            {
                body.Append(new SectionProperties(
                    new PageSize { Width = pageWidthTwips.Value },
                    new PageMargin
                    {
                        Left = (uint)(pageLeftMarginTwips ?? 0),
                        Right = (uint)(pageRightMarginTwips ?? 0)
                    }));
            }

            mainPart.Document = new Document(body);
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
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

    private static byte[] CreatePngWithLargePayload(int width, int height, int extraPayloadBytes)
    {
        var header = CreatePngBytes(width, height);
        var full = new byte[header.Length + extraPayloadBytes];
        header.CopyTo(full, 0);
        for (var i = header.Length; i < full.Length; i++)
        {
            full[i] = (byte)(i % 251);
        }

        return full;
    }

    private sealed class TrackingMemoryStream(byte[] buffer) : MemoryStream(buffer)
    {
        public bool IsDisposed { get; private set; }

        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            base.Dispose(disposing);
        }
    }

    private sealed class NonSeekableDownloadStream(byte[] payload) : Stream
    {
        private readonly MemoryStream _inner = new(payload);

        public bool IsDisposed { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => throw new NotSupportedException();

        public override int Read(byte[] buffer, int offset, int count)
        {
            var chunk = Math.Min(count, 512);
            return _inner.Read(buffer, offset, chunk);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            var slice = buffer.Length > 512 ? buffer[..512] : buffer;
            return _inner.ReadAsync(slice, cancellationToken);
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            if (disposing)
            {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
