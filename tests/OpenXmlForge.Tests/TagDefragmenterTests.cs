using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenXmlForge.Tests;

public class TagDefragmenterTests
{
    [Fact]
    public async Task GenerateStreamAsync_FragmentedTagWithPrefixInFirstRun_ConsolidatesAndReplacesWhilePreservingFormatting()
    {
        // <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">$ {{</w:t></w:r>
        // <w:r><w:t>Services</w:t></w:r>
        // <w:r><w:t>.UnitPrice}}</w:t></w:r>
        // Old MiniWord's AvoidSplitTagText only checked text.InnerText.TrimStart().StartsWith("{")
        // so "$ {{" was completely missed!
        using var template = CreateParagraphWithRawRuns(
            new Run(
                new RunProperties(new Bold(), new Color { Val = "FF0000" }),
                new Text("$ {{") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text("Services")),
            new Run(new Text(".UnitPrice}}")));

        var value = new Dictionary<string, object?>
        {
            ["Services.UnitPrice"] = "150.00"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var p = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        Assert.Equal("$ 150.00", p.InnerText);

        var run = p.Elements<Run>().First(r => !string.IsNullOrEmpty(r.InnerText));
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.Equal("FF0000", run.RunProperties?.Color?.Val?.Value);
    }

    [Fact]
    public async Task GenerateStreamAsync_MultipleFragmentedTagsInSameParagraphWithProofErr_ReplacesAllCleanly()
    {
        // Mirrors Word inserting <w:proofErr w:type="spellStart"/> and splitting around '.'
        using var template = CreateParagraphWithRawElements(
            new Run(new Text("{{Order.")),
            new ProofError { Type = ProofingErrorValues.SpellStart },
            new Run(new Text("Number}}")),
            new ProofError { Type = ProofingErrorValues.SpellEnd },
            new Run(new Text(" - {{") { Space = SpaceProcessingModeValues.Preserve }),
            new ProofError { Type = ProofingErrorValues.GrammarStart },
            new Run(new Text("Customer")),
            new Run(new Text(".")),
            new Run(new Text("Name}}")),
            new ProofError { Type = ProofingErrorValues.GrammarEnd });

        var value = new Dictionary<string, object?>
        {
            ["Order.Number"] = "42",
            ["Customer.Name"] = "Jane"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var p = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        Assert.Equal("42 - Jane", p.InnerText);
        Assert.Empty(p.Descendants<ProofError>());
    }

    [Fact]
    public async Task GenerateStreamAsync_SplitOpeningAndClosingBracesAcrossRunsWithSpacesInsideTag_ConsolidatesCleanly()
    {
        // { in one run, { in next run, spaces around variable name, } in one run, } in next run
        using var template = CreateParagraphWithRawRuns(
            new Run(new Text("Total: {") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text("{  Products.TotalAmount  }")),
            new Run(new Text("} end")));

        var value = new Dictionary<string, object?>
        {
            ["Products.TotalAmount"] = "$ 500.00"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var p = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        Assert.Equal("Total: $ 500.00 end", p.InnerText);
    }

    [Fact]
    public async Task GenerateStreamAsync_FragmentedTagsInsideTableRowWith3PlusDottedVariables_WorksEndToEnd()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var row = new TableRow(
                new TableCell(new Paragraph(
                    new Run(new Text("No {{") { Space = SpaceProcessingModeValues.Preserve }),
                    new ProofError { Type = ProofingErrorValues.SpellStart },
                    new Run(new Text("Order")),
                    new Run(new Text(".Number}}")))),
                new TableCell(new Paragraph(
                    new Run(new Text("{{Customer.")),
                    new Run(new Text("LegalName}}")))),
                new TableCell(new Paragraph(
                    new Run(new Text("$ {{") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text("Services.")),
                    new Run(new Text("UnitPrice")),
                    new Run(new Text("}}")))),
                new TableCell(new Paragraph(
                    new Run(new Text("{{Services")),
                    new Run(new Text(".Description}}")))));

            mainPart.Document = new Document(new Body(new Table(row)));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        var value = new Dictionary<string, object?>
        {
            ["Order.Number"] = "100",
            ["Customer.LegalName"] = "Customer X",
            ["Services"] = new List<Dictionary<string, object?>>
            {
                new() { ["UnitPrice"] = "80.00", ["Description"] = "Installation" },
                new() { ["UnitPrice"] = "120.00", ["Description"] = "Maintenance" }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, value);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var rows = outDoc.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(["No 100", "Customer X", "$ 80.00", "Installation"], rows[0].Elements<TableCell>().Select(c => c.InnerText));
        Assert.Equal(["No 100", "Customer X", "$ 120.00", "Maintenance"], rows[1].Elements<TableCell>().Select(c => c.InnerText));
    }

    private static MemoryStream CreateParagraphWithRawRuns(params Run[] runs)
        => CreateParagraphWithRawElements([.. runs.Cast<OpenXmlElement>()]);

    private static MemoryStream CreateParagraphWithRawElements(params OpenXmlElement[] elements)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var p = new Paragraph();
            foreach (var el in elements)
            {
                p.Append((OpenXmlElement)el.CloneNode(true));
            }

            mainPart.Document = new Document(new Body(p));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }
}
