using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OpenXmlForge.Tests;

public class TableDottedVariablesTests
{
    [Fact]
    public async Task GenerateStreamAsync_WithMoreThan3DottedScalarVariablesInSameTableRow_RendersAllCorrectly()
    {
        // Previously MiniWord threw NotSupportedException("MiniWord doesn't support more than 2 list in same row")
        // whenever a table row had >2 distinct prefixes before '.' (e.g. Order.Number, Customer.LegalName, Contact.Name, Address.City).
        using var template = CreateTableTemplate(
            [
                "{{Order.Number}}",
                "{{Customer.LegalName}}",
                "{{Contact.Name}}",
                "{{Address.City}}",
                "{{Company.LegalName}}",
                "{{Totals.NetAmount}}"
            ]);

        var value = new Dictionary<string, object?>
        {
            ["Order.Number"] = "ORD-2026-001",
            ["Customer.LegalName"] = "Acme Industrial Ltd",
            ["Contact.Name"] = "John Smith",
            ["Address.City"] = "Seattle - WA",
            ["Company.LegalName"] = "Contoso Systems",
            ["Totals.NetAmount"] = "$ 12,500.00"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var rows = doc.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Single(rows);

        var cellTexts = rows[0].Elements<TableCell>().Select(c => c.InnerText).ToList();
        Assert.Equal(
            [
                "ORD-2026-001",
                "Acme Industrial Ltd",
                "John Smith",
                "Seattle - WA",
                "Contoso Systems",
                "$ 12,500.00"
            ],
            cellTexts);
    }

    [Fact]
    public async Task GenerateStreamAsync_With10DottedVariablesInSingleCellOfTableRow_RendersAllCorrectly()
    {
        using var template = CreateTableTemplate(
            [
                "No: {{Order.Number}} | Date: {{Order.Date}} | ValidUntil: {{Order.ValidUntil}} | " +
                "Customer: {{Customer.Name}} ({{Customer.TaxId}}) | Contact: {{Contact.Name}} <{{Contact.Email}}> | " +
                "City: {{Address.City}}/{{Address.State}} | Rep: {{SalesRep.Name}} | Total: {{Totals.NetAmount}}"
            ]);

        var value = new Dictionary<string, object?>
        {
            ["Order.Number"] = "1042",
            ["Order.Date"] = "2026-10-21",
            ["Order.ValidUntil"] = "2026-10-30",
            ["Customer.Name"] = "Tech Corp",
            ["Customer.TaxId"] = "12-3456789",
            ["Contact.Name"] = "Mary",
            ["Contact.Email"] = "mary@tech.example.com",
            ["Address.City"] = "Austin",
            ["Address.State"] = "TX",
            ["SalesRep.Name"] = "Charles",
            ["Totals.NetAmount"] = "$ 9,800.00"
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var text = doc.MainDocumentPart!.Document.Body!.InnerText;
        Assert.Equal(
            "No: 1042 | Date: 2026-10-21 | ValidUntil: 2026-10-30 | " +
            "Customer: Tech Corp (12-3456789) | Contact: Mary <mary@tech.example.com> | " +
            "City: Austin/TX | Rep: Charles | Total: $ 9,800.00",
            text);
    }

    [Fact]
    public async Task GenerateStreamAsync_WithNestedPocoDottedPropertiesInSameRow_RendersAllCorrectly()
    {
        using var template = CreateTableTemplate(
            [
                "{{Order.Number}}",
                "{{Customer.Name}}",
                "{{Contact.Email}}",
                "{{Address.City}}",
                "{{Payment.Terms}}"
            ]);

        var model = new
        {
            Order = new { Number = "999" },
            Customer = new { Name = "Alpha Inc" },
            Contact = new { Email = "contact@alpha.example.com" },
            Address = new { City = "Denver" },
            Payment = new { Terms = "Net 30/60/90" }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, model);

        using var doc = WordprocessingDocument.Open(result, false);
        var cellTexts = doc.MainDocumentPart!.Document.Body!
            .Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().Select(c => c.InnerText).ToList();

        Assert.Equal(["999", "Alpha Inc", "contact@alpha.example.com", "Denver", "Net 30/60/90"], cellTexts);
    }

    [Fact]
    public async Task GenerateStreamAsync_RepeatingTableRowWith8ColumnsAndOuterDottedScalars_ExpandsRowsAndPreservesScalars()
    {
        using var template = CreateTableTemplate(
            [
                "{{Order.Number}}",
                "{{Customer.LegalName}}",
                "{{Company.Currency}}",
                "{{Products.Index}}",
                "{{Products.Sku}}",
                "{{Products.Description}}",
                "{{Products.Unit}}",
                "{{Products.Quantity}}",
                "{{Products.UnitPrice}}",
                "{{Products.Discount}}",
                "{{Products.TotalPrice}}"
            ]);

        var value = new Dictionary<string, object?>
        {
            ["Order.Number"] = "ORD-55",
            ["Customer.LegalName"] = "Beta Trading",
            ["Company.Currency"] = "USD",
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["Index"] = 1,
                    ["Sku"] = "P01",
                    ["Description"] = "Electric Motor",
                    ["Unit"] = "EA",
                    ["Quantity"] = 2,
                    ["UnitPrice"] = "1,000.00",
                    ["Discount"] = "50.00",
                    ["TotalPrice"] = "1,950.00"
                },
                new()
                {
                    ["Index"] = 2,
                    ["Sku"] = "P02",
                    ["Description"] = "Frequency Inverter",
                    ["Unit"] = "PC",
                    ["Quantity"] = 1,
                    ["UnitPrice"] = "800.00",
                    ["Discount"] = "0.00",
                    ["TotalPrice"] = "800.00"
                }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var rows = doc.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(2, rows.Count);

        var row0 = rows[0].Elements<TableCell>().Select(c => c.InnerText).ToList();
        var row1 = rows[1].Elements<TableCell>().Select(c => c.InnerText).ToList();

        Assert.Equal(
            ["ORD-55", "Beta Trading", "USD", "1", "P01", "Electric Motor", "EA", "2", "1,000.00", "50.00", "1,950.00"],
            row0);
        Assert.Equal(
            ["ORD-55", "Beta Trading", "USD", "2", "P02", "Frequency Inverter", "PC", "1", "800.00", "0.00", "800.00"],
            row1);
    }

    [Fact]
    public async Task GenerateStreamAsync_MultipleCollectionsInSameRow_ExpandsUpToMaxCountWithoutThrowing()
    {
        using var template = CreateTableTemplate(
            [
                "{{ColA.Name}}",
                "{{ColB.Name}}",
                "{{ColC.Name}}",
                "{{ColD.Name}}"
            ]);

        var value = new Dictionary<string, object?>
        {
            ["ColA"] = new List<Dictionary<string, object?>> { new() { ["Name"] = "A1" }, new() { ["Name"] = "A2" } },
            ["ColB"] = new List<Dictionary<string, object?>> { new() { ["Name"] = "B1" } },
            ["ColC"] = new List<Dictionary<string, object?>> { new() { ["Name"] = "C1" }, new() { ["Name"] = "C2" }, new() { ["Name"] = "C3" } },
            ["ColD"] = new List<Dictionary<string, object?>> { new() { ["Name"] = "D1" }, new() { ["Name"] = "D2" } }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var rows = doc.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count);

        Assert.Equal(["A1", "B1", "C1", "D1"], rows[0].Elements<TableCell>().Select(c => c.InnerText));
        Assert.Equal(["A2", "", "C2", "D2"], rows[1].Elements<TableCell>().Select(c => c.InnerText));
        Assert.Equal(["", "", "C3", ""], rows[2].Elements<TableCell>().Select(c => c.InnerText));
    }

    [Fact]
    public async Task GenerateStreamAsync_NonDottedTagFollowedByDottedTagInSameCell_DoesNotCorruptPrefixMatch()
    {
        // In old MiniWord, Regex "(?<={{).*?\\..*?(?=}})" matched across "Currency}} {{Products.Price"
        // extracting prefix "Currency}} {{Products" and failing!
        using var template = CreateTableTemplate(["{{Currency}} {{Products.Price}}"]);

        var value = new Dictionary<string, object?>
        {
            ["Currency"] = "$",
            ["Products"] = new List<Dictionary<string, object?>>
            {
                new() { ["Price"] = "10.00" },
                new() { ["Price"] = "25.00" }
            }
        };

        await using var result = await WordForge.GenerateStreamAsync(template, value);

        using var doc = WordprocessingDocument.Open(result, false);
        var rows = doc.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("$ 10.00", rows[0].InnerText);
        Assert.Equal("$ 25.00", rows[1].InnerText);
    }

    private static MemoryStream CreateTableTemplate(string[] cellTexts)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var row = new TableRow();
            foreach (var text in cellTexts)
            {
                row.Append(new TableCell(new Paragraph(new Run(new Text(text)))));
            }

            mainPart.Document = new Document(new Body(new Table(row)));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }
}
