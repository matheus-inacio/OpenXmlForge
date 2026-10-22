using System.Dynamic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlForge.Models;
using Snapshooter.Xunit;
using Xunit;

namespace OpenXmlForge.Tests;

public class CompatibilityTests
{
    private static readonly string SamplesDir = Path.Combine(AppContext.BaseDirectory, "Samples");

    [Fact]
    public async Task GenerateStreamAsync_BlockIfAndEndIf_KeepsTrueBranchAndRemovesFalseBranch()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("@if(Total > 100)"))),
                    new Paragraph(new Run(new Text("Approved: {{Customer}}"))),
                    new Paragraph(new Run(new Text("@endif"))),
                    new Paragraph(new Run(new Text("@if(Total <= 100)"))),
                    new Paragraph(new Run(new Text("Rejected"))),
                    new Paragraph(new Run(new Text("@endif")))));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        var value = new Dictionary<string, object?>
        {
            ["Total"] = 250.50m,
            ["Customer"] = "Acme Corp"
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, value);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var paragraphs = outDoc.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Select(p => p.InnerText)
            .ToList();
        Assert.Equal(["Approved: Acme Corp"], paragraphs);
    }

    [Fact]
    public async Task GenerateStreamAsync_InlineIfStatement_EvaluatesBooleansAndComparisonsCorrectly()
    {
        // Note: old MiniWord had an inverted boolean bug where Eval("true == true") returned false
        // so {{if(1 == 1) YES endif}} was deleted instead of kept!
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text(
                        "Status: {{if(IsActive == true) ACTIVE endif}}{{if(IsActive == false) INACTIVE endif}}")))));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        await using var result = await WordForge.GenerateStreamAsync(stream, new { IsActive = true });

        using var outDoc = WordprocessingDocument.Open(result, false);
        var text = outDoc.MainDocumentPart!.Document.Body!.InnerText;
        Assert.Equal("Status: ACTIVE", text);
    }

    [Fact]
    public async Task GenerateStreamAsync_InlineForeachAndBlockForeach_ExpandsCollectionsCleanly()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document =new Document(new Body(
                    new Paragraph(new Run(new Text("Items: {{foreach item in Items}}[{{item.Name}}]{{endforeach}}"))),
                    new Paragraph(new Run(new Text("@foreach(item in Items)"))),
                    new Paragraph(new Run(new Text("- {{item.Name}}: {{item.Price}}"))),
                    new Paragraph(new Run(new Text("@endforeach")))));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        var data = new
        {
            Items = new[]
            {
                new { Name = "Bolt", Price = "1.50" },
                new { Name = "Nut", Price = "0.80" }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, data);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var paragraphs = outDoc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Select(p => p.InnerText).ToList();
        Assert.Equal(
            [
                "Items: [Bolt][Nut]",
                "- Bolt: 1.50",
                "- Nut: 0.80"
            ],
            paragraphs);
    }

    [Fact]
    public async Task GenerateStreamAsync_HyperLinkAndColorTextAndWordForeach_RenderExpectedElements()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("Link: {{Site}}"))),
                    new Paragraph(new Run(new Text("Badge: {{Status}}"))),
                    new Paragraph(new Run(new Text("Tags: {{Tags}} - {{Tags}}")))));
            mainPart.Document.Save();
        }

        stream.Position = 0;

        var data = new Dictionary<string, object?>
        {
            ["Site"] = new WordHyperLink
            {
                Url = "https://example.com",
                Text = "Customer Portal"
            },
            ["Status"] = new WordColorText
            {
                Text = "APPROVED",
                FontColor = "#FFFFFF",
                HighlightColor = "#00AA00"
            },
            ["Tags"] = new WordForeach
            {
                Value = new List<Dictionary<string, object?>>
                {
                    new() { ["Tags"] = "Urgent" },
                    new() { ["Tags"] = "VIP" }
                },
                Separator = " | "
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(stream, data);

        using var outDoc = WordprocessingDocument.Open(result, false);
        var body = outDoc.MainDocumentPart!.Document.Body!;
        var hyperlink = Assert.Single(body.Descendants<Hyperlink>());
        Assert.Equal("Customer Portal", hyperlink.InnerText);
        Assert.Contains(outDoc.MainDocumentPart.HyperlinkRelationships, r => r.Uri.ToString() == "https://example.com/");

        var shading = Assert.Single(body.Descendants<Shading>());
        Assert.Equal("00AA00", shading.Fill?.Value);

        var paragraphs = body.Elements<Paragraph>().Select(p => p.InnerText).ToList();
        Assert.Equal("Tags: Urgent | VIP - Urgent | VIP", paragraphs[2]);
    }

    [Fact]
    public async Task GenerateStreamAsync_VerticallyMergedTableRows_ClonesEntireRowGroupPerItem()
    {
        // Solves MiniWord Issue #69 / TestIssue47.docx where a single logical item spans 2+ rows via <w:vMerge/>
        var samplePath = Path.Combine(SamplesDir, "TestIssue47.docx");
        Assert.True(File.Exists(samplePath), $"Expected sample file at {samplePath}");

        var value = new Dictionary<string, object?>
        {
            ["line"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["num"] = "1",
                    ["spc"] = "Specification1",
                    ["desc"] = "Description1",
                    ["qty"] = "10",
                    ["up"] = "20.00",
                    ["tax_rate"] = "0.20",
                    ["tax_amount"] = "4.00"
                },
                new()
                {
                    ["num"] = "2",
                    ["spc"] = "Specification2",
                    ["desc"] = "Description2",
                    ["qty"] = "5",
                    ["up"] = "30.00",
                    ["tax_rate"] = "0.15",
                    ["tax_amount"] = "2.25"
                }
            }
        };

        await using var templateStream = File.OpenRead(samplePath);
        await using var result = await WordForge.GenerateStreamAsync(templateStream, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var text = doc.MainDocumentPart!.Document.Body!.InnerText;

        Assert.Contains("Specification1", text);
        Assert.Contains("Description1", text);
        Assert.Contains("Specification2", text);
        Assert.Contains("Description2", text);
        Assert.DoesNotContain("{{line.", text);

        DocxSnapshotHelper.Extract(result).MatchSnapshot();
    }

    [Fact]
    public async Task GenerateStreamAsync_OriginalMiniWordSampleTemplates_RenderExpectedContent()
    {
        // 1. TestIfStatement.docx
        var ifSample = Path.Combine(SamplesDir, "TestIfStatement.docx");
        Assert.True(File.Exists(ifSample));
        var ifData = new Dictionary<string, object?>
        {
            ["Name"] = new List<WordHyperLink>
            {
                new() { Url = "https://google.com", Text = "Link1" },
                new() { Url = "https://google1.com", Text = "Link2" }
            },
            ["Company_Name"] = "MiniSofteware",
            ["CreateDate"] = new DateTime(2021, 01, 01),
            ["VIP"] = true,
            ["Points"] = 123,
            ["APP"] = "Demo APP"
        };

        await using (var s = File.OpenRead(ifSample))
        await using (var outStream = await WordForge.GenerateStreamAsync(s, ifData))
        {
            using var doc = WordprocessingDocument.Open(outStream, false);
            var bodyText = doc.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("First if chosen: MiniSofteware", bodyText);
            Assert.DoesNotContain("Second if chosen: MaxiSoftware", bodyText);
            Assert.Contains("Points are greater than 100", bodyText);
            Assert.Contains("CreateDate is not less than 2021", bodyText);
            Assert.DoesNotContain("CreateDate is not greater than 2021", bodyText);
        }

        // 2. TestForeachInTablesWithIfStatementDemo.docx
        var foreachSample = Path.Combine(SamplesDir, "TestForeachInTablesWithIfStatementDemo.docx");
        Assert.True(File.Exists(foreachSample));
        var foreachData = new Dictionary<string, object?>
        {
            ["TripHs"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["sDate"] = DateTime.Parse("2022-09-08 08:30:00"),
                    ["eDate"] = DateTime.Parse("2022-09-08 15:00:00"),
                    ["How"] = "Discussion requirement part1",
                    ["Details"] = new List<WordForeach>
                    {
                        new() { Value = new Dictionary<string, object> { ["Text"] = "Air", ["Value"] = "Airplane" }, Separator = " | " },
                        new() { Value = new Dictionary<string, object> { ["Text"] = "Parking", ["Value"] = "Car" }, Separator = " / " },
                        new() { Value = new Dictionary<string, object> { ["Text"] = "Hotel", ["Value"] = "Room" }, Separator = ", " },
                        new() { Value = new Dictionary<string, object> { ["Text"] = "Food", ["Value"] = "Plate" }, Separator = "" }
                    }
                }
            }
        };

        await using (var s = File.OpenRead(foreachSample))
        await using (var outStream = await WordForge.GenerateStreamAsync(s, foreachData))
        {
            using var doc = WordprocessingDocument.Open(outStream, false);
            var bodyText = doc.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Discussion requirement part1", bodyText);
            Assert.Contains("Air way to the Airplane | Hotel way to the Room", bodyText);
        }
    }

    [Fact]
    public void SaveAsByTemplate_WithExpandoObjectAndStringArrays_ReplacesCleanly()
    {
        using var template = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document =
            new Document(new Body(
                    new Paragraph(new Run(new Text("Hello {{Name}}: {{Notes}}")))));
            mainPart.Document.Save();
        }

        dynamic expando = new ExpandoObject();
        expando.Name = "Alice";
        expando.Notes = new[] { "Item A", "Item B" };

        using var output = new MemoryStream();
        WordForge.SaveAsByTemplate(output, template.ToArray(), (object)expando);

        output.Position = 0;
        using var outDoc = WordprocessingDocument.Open(output, false);
        var body = outDoc.MainDocumentPart!.Document.Body!;
        Assert.Equal("Hello Alice: Item AItem B", body.InnerText);
        Assert.Single(body.Descendants<Break>());
    }
}
