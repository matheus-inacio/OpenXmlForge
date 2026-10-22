using System.Dynamic;
using OpenXmlForge.Models;
using Snapshooter.Xunit;
using Xunit;

namespace OpenXmlForge.Tests;

public class SampleTemplatesTests
{
    [Fact]
    public async Task TestForeachLoopInTables_SyncAndAsync_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestForeachInTablesDemo.docx");
        var value = CreateTripHoursWithForeachDetailsPayload();

        var syncPath = DocxSnapshotHelper.GetTempFilePath();
        var asyncPath = DocxSnapshotHelper.GetTempFilePath();
        try
        {
            WordForge.SaveAsByTemplate(syncPath, templatePath, value);
            await WordForge.SaveAsByTemplateAsync(asyncPath, templatePath, value);

            var xml = DocxSnapshotHelper.GetZipFileContent(syncPath, "word/document.xml");
            Assert.Contains("Discussion requirement part2 and development", xml);
            Assert.Contains("Discussion requirement part1", xml);
            Assert.Contains(
                "Air way to the Airplane | Parking way to the Car / Hotel way to the Room, Food way to the Plate",
                xml);

            var syncSnapshot = DocxSnapshotHelper.Extract(syncPath);
            var asyncSnapshot = DocxSnapshotHelper.Extract(asyncPath);
            Assert.Equal(syncSnapshot.DocumentXml, asyncSnapshot.DocumentXml);

            syncSnapshot.MatchSnapshot();
        }
        finally
        {
            DeleteIfExists(syncPath);
            DeleteIfExists(asyncPath);
        }
    }

    [Fact]
    public async Task TestForeachLoopInTablesWithIfStatement_SyncAndAsync_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestForeachInTablesWithIfStatementDemo.docx");
        var value = CreateTripHoursWithForeachDetailsPayload();

        var syncPath = DocxSnapshotHelper.GetTempFilePath();
        var asyncPath = DocxSnapshotHelper.GetTempFilePath();
        try
        {
            WordForge.SaveAsByTemplate(syncPath, templatePath, value);
            await WordForge.SaveAsByTemplateAsync(asyncPath, templatePath, value);

            var xml = DocxSnapshotHelper.GetZipFileContent(syncPath, "word/document.xml");
            Assert.Contains("Discussion requirement part2 and development", xml);
            Assert.Contains("Discussion requirement part1", xml);
            Assert.Contains("Air way to the Airplane | Hotel way to the Room", xml);

            var syncSnapshot = DocxSnapshotHelper.Extract(syncPath);
            var asyncSnapshot = DocxSnapshotHelper.Extract(asyncPath);
            Assert.Equal(syncSnapshot.DocumentXml, asyncSnapshot.DocumentXml);

            syncSnapshot.MatchSnapshot();
        }
        finally
        {
            DeleteIfExists(syncPath);
            DeleteIfExists(asyncPath);
        }
    }

    [Fact]
    public async Task IfStatement_FirstIf_SyncAndAsync_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestIfStatement.docx");
        var value = new Dictionary<string, object?>
        {
            ["Name"] = new List<WordHyperLink>
            {
                new() { Url = "https://google.com", Text = "測試連結22!!" },
                new() { Url = "https://google1.com", Text = "測試連結11!!" },
            },
            ["Company_Name"] = "MiniSofteware",
            ["CreateDate"] = new DateTime(2021, 01, 01),
            ["VIP"] = true,
            ["Points"] = 123,
            ["APP"] = "Demo APP",
        };

        var syncPath = DocxSnapshotHelper.GetTempFilePath();
        var asyncPath = DocxSnapshotHelper.GetTempFilePath();
        try
        {
            WordForge.SaveAsByTemplate(syncPath, templatePath, value);
            await WordForge.SaveAsByTemplateAsync(asyncPath, templatePath, value);

            var docXml = DocxSnapshotHelper.GetZipFileContent(syncPath, "word/document.xml");
            Assert.Contains("First if chosen: MiniSofteware", docXml);
            Assert.DoesNotContain("Second if chosen: MaxiSoftware", docXml);
            Assert.Contains("Points are greater than 100", docXml);
            Assert.Contains("CreateDate is not less than 2021", docXml);
            Assert.DoesNotContain("CreateDate is not greater than 2021", docXml);

            var syncSnapshot = DocxSnapshotHelper.Extract(syncPath);
            var asyncSnapshot = DocxSnapshotHelper.Extract(asyncPath);
            Assert.Equal(syncSnapshot.DocumentXml, asyncSnapshot.DocumentXml);

            syncSnapshot.MatchSnapshot();
        }
        finally
        {
            DeleteIfExists(syncPath);
            DeleteIfExists(asyncPath);
        }
    }

    [Fact]
    public async Task TestIssue69_TestIssue47Docx_SyncAndAsync_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestIssue47.docx");
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
                    ["tax_amount"] = "4.00",
                },
                new()
                {
                    ["num"] = "2",
                    ["spc"] = "Specification2",
                    ["desc"] = "Description2",
                    ["qty"] = "5",
                    ["up"] = "30.00",
                    ["tax_rate"] = "0.15",
                    ["tax_amount"] = "2.25",
                },
            },
        };

        var syncPath = DocxSnapshotHelper.GetTempFilePath();
        var asyncPath = DocxSnapshotHelper.GetTempFilePath();
        try
        {
            WordForge.SaveAsByTemplate(syncPath, templatePath, value);
            await WordForge.SaveAsByTemplateAsync(asyncPath, templatePath, value);

            var syncSnapshot = DocxSnapshotHelper.Extract(syncPath);
            var asyncSnapshot = DocxSnapshotHelper.Extract(asyncPath);
            Assert.Equal(syncSnapshot.DocumentXml, asyncSnapshot.DocumentXml);

            syncSnapshot.MatchSnapshot();
        }
        finally
        {
            DeleteIfExists(syncPath);
            DeleteIfExists(asyncPath);
        }
    }

    [Fact]
    public async Task TestDemo04_SplitTag_DictionaryAndAnonymous_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestDemo04.docx");
        var dictValue = new Dictionary<string, object?> { ["title"] = "Hello MiniWord" };
        var anonValue = new { title = "Hello MiniWord" };

        await using var s1 = File.OpenRead(templatePath);
        await using var outDict = await WordForge.GenerateStreamAsync(s1, dictValue);

        using var s2 = File.OpenRead(templatePath);
        using var outAnon = WordForge.GenerateStream(s2, anonValue);

        var xml = DocxSnapshotHelper.GetZipFileContent(outDict, "word/document.xml");
        Assert.Contains("MiniWord", xml);

        var dictSnapshot = DocxSnapshotHelper.Extract(outDict);
        var anonSnapshot = DocxSnapshotHelper.Extract(outAnon);
        Assert.Equal(dictSnapshot.DocumentXml, anonSnapshot.DocumentXml);

        dictSnapshot.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue37_PlainTextAndHyperlinks_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestIssue37.docx");

        var textValue = new Dictionary<string, object?>
        {
            ["Content"] = "Test",
            ["Content2"] = "Test2",
        };

        var linkValue = new Dictionary<string, object?>
        {
            ["Content"] = new WordHyperLink
            {
                Url = "https://google.com",
                Text = "Test1!!",
            },
            ["Content2"] = new WordHyperLink
            {
                Url = "https://google.com",
                Text = "Test2!!",
            },
        };

        await using var s1 = File.OpenRead(templatePath);
        await using var outText = await WordForge.GenerateStreamAsync(s1, textValue);

        await using var s2 = File.OpenRead(templatePath);
        await using var outLink = await WordForge.GenerateStreamAsync(s2, linkValue);

        var xmlText = DocxSnapshotHelper.GetZipFileContent(outText, "word/document.xml");
        Assert.Contains("Test", xmlText);
        Assert.Contains("Test2", xmlText);

        var xmlLink = DocxSnapshotHelper.GetZipFileContent(outLink, "word/document.xml");
        Assert.Contains("Test1!!", xmlLink);
        Assert.Contains("Test2!!", xmlLink);

        new
        {
            PlainTextVariant = DocxSnapshotHelper.Extract(outText),
            HyperlinkVariant = DocxSnapshotHelper.Extract(outLink),
        }.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue18_And_TestTemplateComplex_DictionaryAnonymousAndPoco_MatchesSnapshot()
    {
        var issue18Path = DocxSnapshotHelper.GetSamplePath("TestIssue18.docx");
        var complexPath = DocxSnapshotHelper.GetSamplePath("TestTemplateComplex.docx");

        var dictValue = new Dictionary<string, object?>
        {
            ["title"] = "FooCompany",
            ["managers"] = new List<Dictionary<string, object?>>
            {
                new() { ["name"] = "Jack", ["department"] = "HR" },
                new() { ["name"] = "Loan", ["department"] = "IT" },
            },
            ["employees"] = new List<Dictionary<string, object?>>
            {
                new() { ["name"] = "Wade", ["department"] = "HR" },
                new() { ["name"] = "Felix", ["department"] = "HR" },
                new() { ["name"] = "Eric", ["department"] = "IT" },
                new() { ["name"] = "Keaton", ["department"] = "IT" },
            },
        };

        var anonValue = new
        {
            title = "FooCompany",
            managers = new[]
            {
                new { name = "Jack", department = "HR" },
                new { name = "Loan", department = "IT" },
            },
            employees = new[]
            {
                new { name = "Wade", department = "HR" },
                new { name = "Felix", department = "HR" },
                new { name = "Eric", department = "IT" },
                new { name = "Keaton", department = "IT" },
            },
        };

        var pocoValue = new Foo
        {
            title = "FooCompany",
            managers =
            [
                new User { name = "Jack", department = "HR" },
                new User { name = "Loan", department = "IT" },
            ],
            employees =
            [
                new User { name = "Wade", department = "HR" },
                new User { name = "Felix", department = "HR" },
                new User { name = "Eric", department = "IT" },
                new User { name = "Keaton", department = "IT" },
            ],
        };

        await using var s1 = File.OpenRead(issue18Path);
        await using var outDict = await WordForge.GenerateStreamAsync(s1, dictValue);

        using var s2 = File.OpenRead(issue18Path);
        using var outAnon = WordForge.GenerateStream(s2, anonValue);

        await using var s3 = File.OpenRead(complexPath);
        await using var outPoco = await WordForge.GenerateStreamAsync(s3, pocoValue);

        var xml = DocxSnapshotHelper.GetZipFileContent(outDict, "word/document.xml");
        Assert.Contains("<w:t>Keaton", xml);
        Assert.Contains("<w:t>Eric", xml);

        var snapDict = DocxSnapshotHelper.Extract(outDict);
        var snapAnon = DocxSnapshotHelper.Extract(outAnon);
        var snapPoco = DocxSnapshotHelper.Extract(outPoco);

        Assert.Equal(snapDict.DocumentXml, snapAnon.DocumentXml);
        Assert.Equal(snapDict.DocumentXml, snapPoco.DocumentXml);

        snapDict.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue17_SplitTag_DictionaryAndAnonymous_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestIssue17.docx");
        var dictValue = new Dictionary<string, object?>
        {
            ["Content"] = "Test",
            ["Content2"] = "Test2",
        };
        var anonValue = new
        {
            Content = "Test",
            Content2 = "Test2",
        };

        await using var s1 = File.OpenRead(templatePath);
        await using var outDict = await WordForge.GenerateStreamAsync(s1, dictValue);

        using var s2 = File.OpenRead(templatePath);
        using var outAnon = WordForge.GenerateStream(s2, anonValue);

        var xml = DocxSnapshotHelper.GetZipFileContent(outDict, "word/document.xml");
        Assert.Contains("<w:t>Test", xml);
        Assert.Contains("<w:t>Test2", xml);

        var snapDict = DocxSnapshotHelper.Extract(outDict);
        var snapAnon = DocxSnapshotHelper.Extract(outAnon);
        Assert.Equal(snapDict.DocumentXml, snapAnon.DocumentXml);

        snapDict.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue13_And_TestDemo01TagText_TestExpenseDemoDocx_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestExpenseDemo.docx");
        var meeting02Path = DocxSnapshotHelper.GetSamplePath("DemoExpenseMeeting02.png");
        var meeting01Path = DocxSnapshotHelper.GetSamplePath("DemoExpenseMeeting01.png");
        var logoPath = DocxSnapshotHelper.GetSamplePath("DemoLogo.png");

        var value = new Dictionary<string, object?>
        {
            ["logo"] = new WordPicture(logoPath, 120, 120),
            ["Name"] = "Jack",
            ["Department"] = "IT Department",
            ["Purpose"] = "Shanghai site needs a new system to control HR system.",
            ["StartDate"] = DateTime.Parse("2022-09-07 08:30:00"),
            ["EndDate"] = DateTime.Parse("2022-09-15 15:30:00"),
            ["Approved"] = true,
            ["Total_Amount"] = 123456,
            ["TripHs"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["sDate"] = DateTime.Parse("2022-09-08 08:30:00"),
                    ["eDate"] = DateTime.Parse("2022-09-08 15:00:00"),
                    ["How"] = "Discussion requirement part1",
                    ["Photo"] = new WordPicture(meeting02Path, 160, 90),
                },
                new()
                {
                    ["sDate"] = DateTime.Parse("2022-09-09 08:30:00"),
                    ["eDate"] = DateTime.Parse("2022-09-09 17:00:00"),
                    ["How"] = "Discussion requirement part2 and development",
                    ["Photo"] = new WordPicture(meeting01Path, 160, 90),
                },
            },
        };

        await using var s = File.OpenRead(templatePath);
        await using var outStream = await WordForge.GenerateStreamAsync(s, value);

        var xml = DocxSnapshotHelper.GetZipFileContent(outStream, "word/document.xml");
        Assert.Contains("Discussion requirement part2 and development", xml);
        Assert.Contains("Discussion requirement part1", xml);
        Assert.Contains("Shanghai site needs a new system to control HR system.", xml);

        DocxSnapshotHelper.Extract(outStream).MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue12_MultilineBasicFill_DictionaryDynamicAndAnonymous_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestBasicFill.docx");

        var dictValue = new Dictionary<string, object?>
        {
            ["Company_Name"] = "MiniSofteware\n",
            ["Name"] = "Jack",
            ["CreateDate"] = new DateTime(2021, 01, 01),
            ["VIP"] = true,
            ["Points"] = 123,
            ["APP"] = "Demo APP\n",
        };

        dynamic expandoValue = new ExpandoObject();
        expandoValue.Company_Name = "MiniSofteware\n";
        expandoValue.Name = "Jack";
        expandoValue.CreateDate = new DateTime(2021, 01, 01);
        expandoValue.VIP = true;
        expandoValue.Points = 123;
        expandoValue.APP = "Demo APP\n";

        var anonValue = new
        {
            Company_Name = "MiniSofteware\n",
            Name = "Jack",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123,
            APP = "Demo APP\n",
        };

        await using var s1 = File.OpenRead(templatePath);
        await using var outDict = await WordForge.GenerateStreamAsync(s1, dictValue);

        await using var s2 = File.OpenRead(templatePath);
        await using var outExpando = await WordForge.GenerateStreamAsync(s2, (object)expandoValue);

        using var s3 = File.OpenRead(templatePath);
        using var outAnon = WordForge.GenerateStream(s3, anonValue);

        var xml = DocxSnapshotHelper.GetZipFileContent(outDict, "word/document.xml");
        Assert.Contains("<w:t>MiniSofteware", xml);
        Assert.Contains("<w:br />", xml);

        var snapDict = DocxSnapshotHelper.Extract(outDict);
        var snapExpando = DocxSnapshotHelper.Extract(outExpando);
        var snapAnon = DocxSnapshotHelper.Extract(outAnon);

        Assert.Equal(snapDict.DocumentXml, snapExpando.DocumentXml);
        Assert.Equal(snapDict.DocumentXml, snapAnon.DocumentXml);

        snapDict.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssueDemo03_TestDemo02Docx_DictionaryAndDynamic_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestDemo02.docx");
        var demo01Png = DocxSnapshotHelper.GetSamplePath("demo01.png");

        const string experiences = """
            # SENIOR UI/UX DEVELOPER & DESIGNER
            ◼ The Matrix Media Limited | From May 2013 to May 2015
            Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took.

            ◼ JUNIOR UI/UX DEVELOPER & DESIGNER
            Linux OS Interface Limited | From Jan 2010 to Feb 2013
            Lorem Ipsum has been the industry's standard dummy text 
            ever since the 1500s, when an unknown printer took.

            ◼ TEAM LEADER & CORE GRAPHIC DESIGNER
            Apple OS Interface Limited | From Jan 2008 to Feb 2010
            Lorem Ipsum has been the industry's standard dummy text 
            ever since the 1500s, when an unknown printer took.
            """;

        var dictValue = new Dictionary<string, object?>
        {
            ["FullName"] = "Julian Anderson",
            ["Title"] = "IT Manager",
            ["Phone"] = "+86 1234567890",
            ["Mail"] = "shps95100@gmail.com",
            ["Education"] = "Michigan State University | From Aug 2013 to May 2015",
            ["Major"] = "Computer Science",
            ["Favorites"] = "Music、Programing、Design",
            ["Skills"] = new[] { "- Photoshop", "- InDesign", "- MS Office", "- HTML 5", "- CSS 3" },
            ["Address"] = "1234, White Home, Road-12/ABC Street-13, New York, USA, 12345",
            ["AboutMe"] = "Hi, I’m Julian Anderson dummy text of the printing and typesetting industry.",
            ["Birthday"] = "1993-09-26",
            ["Experiences"] = experiences,
            ["Image"] = new WordPicture(demo01Png, 160, 90),
        };

        dynamic expandoValue = new ExpandoObject();
        expandoValue.FullName = "Julian Anderson";
        expandoValue.Title = "IT Manager";
        expandoValue.Phone = "+86 1234567890";
        expandoValue.Mail = "shps95100@gmail.com";
        expandoValue.Education = "Michigan State University | From Aug 2013 to May 2015";
        expandoValue.Major = "Computer Science";
        expandoValue.Favorites = "Music、Programing、Design";
        expandoValue.Skills = new[] { "- Photoshop", "- InDesign", "- MS Office", "- HTML 5", "- CSS 3" };
        expandoValue.Address = "1234, White Home, Road-12/ABC Street-13, New York, USA, 12345";
        expandoValue.AboutMe = "Hi, I’m Julian Anderson dummy text of the printing and typesetting industry.";
        expandoValue.Birthday = "1993-09-26";
        expandoValue.Experiences = experiences;
        expandoValue.Image = new WordPicture(demo01Png, 160, 90);

        await using var s1 = File.OpenRead(templatePath);
        await using var outDict = await WordForge.GenerateStreamAsync(s1, dictValue);

        using var s2 = File.OpenRead(templatePath);
        using var outExpando = WordForge.GenerateStream(s2, (object)expandoValue);

        var snapDict = DocxSnapshotHelper.Extract(outDict);
        var snapExpando = DocxSnapshotHelper.Extract(outExpando);
        Assert.Equal(snapDict.DocumentXml, snapExpando.DocumentXml);

        snapDict.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue11_StringArrayAndList_DictionaryAndAnonymous_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestIssue11.docx");

        var dictArray = new Dictionary<string, object?>
        {
            ["managers"] = new[] { "Jack", "Alan" },
            ["employees"] = new[] { "Mike", "Henry" },
        };

        var dictList = new Dictionary<string, object?>
        {
            ["managers"] = new List<string> { "Jack", "Alan" },
            ["employees"] = new List<string> { "Mike", "Henry" },
        };

        var anonArray = new
        {
            managers = new[] { "Jack", "Alan" },
            employees = new[] { "Mike", "Henry" },
        };

        var anonList = new
        {
            managers = new List<string> { "Jack", "Alan" },
            employees = new List<string> { "Mike", "Henry" },
        };

        await using var s1 = File.OpenRead(templatePath);
        await using var outDictArray = await WordForge.GenerateStreamAsync(s1, dictArray);

        await using var s2 = File.OpenRead(templatePath);
        await using var outDictList = await WordForge.GenerateStreamAsync(s2, dictList);

        using var s3 = File.OpenRead(templatePath);
        using var outAnonArray = WordForge.GenerateStream(s3, anonArray);

        using var s4 = File.OpenRead(templatePath);
        using var outAnonList = WordForge.GenerateStream(s4, anonList);

        var snap1 = DocxSnapshotHelper.Extract(outDictArray);
        var snap2 = DocxSnapshotHelper.Extract(outDictList);
        var snap3 = DocxSnapshotHelper.Extract(outAnonArray);
        var snap4 = DocxSnapshotHelper.Extract(outAnonList);

        Assert.Equal(snap1.DocumentXml, snap2.DocumentXml);
        Assert.Equal(snap1.DocumentXml, snap3.DocumentXml);
        Assert.Equal(snap1.DocumentXml, snap4.DocumentXml);

        snap1.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue3_TestBasicImageDocx_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestBasicImage.docx");
        var logoPath = DocxSnapshotHelper.GetSamplePath("DemoLogo.png");

        var value = new Dictionary<string, object?>
        {
            ["Logo"] = new WordPicture(logoPath, 180, 180),
        };

        await using var s = File.OpenRead(templatePath);
        await using var outStream = await WordForge.GenerateStreamAsync(s, value);

        var xml = DocxSnapshotHelper.GetZipFileContent(outStream, "word/document.xml");
        Assert.Contains("<w:drawing", xml);

        DocxSnapshotHelper.Extract(outStream).MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue5_And_HyperLinkListAndArray_TestBasicFillDocx_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestBasicFill.docx");

        var singleLinkValue = new
        {
            Name = new WordHyperLink
            {
                Url = "https://google.com",
                Text = "測試連結!!",
            },
            Company_Name = "MiniSofteware",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123,
            APP = "Demo APP",
        };

        var listLinkValue = new Dictionary<string, object?>
        {
            ["Name"] = new List<WordHyperLink>
            {
                new() { Url = "https://google.com", Text = "測試連結22!!" },
                new() { Url = "https://google1.com", Text = "測試連結11!!" },
            },
            ["Company_Name"] = "MiniSofteware",
            ["CreateDate"] = new DateTime(2021, 01, 01),
            ["VIP"] = true,
            ["Points"] = 123,
            ["APP"] = "Demo APP",
        };

        var arrayLinkValue = new Dictionary<string, object?>
        {
            ["Name"] = new[]
            {
                new WordHyperLink { Url = "https://google.com", Text = "測試連結22!!" },
                new WordHyperLink { Url = "https://google1.com", Text = "測試連結11!!" },
            },
            ["Company_Name"] = "MiniSofteware",
            ["CreateDate"] = new DateTime(2021, 01, 01),
            ["VIP"] = true,
            ["Points"] = 123,
            ["APP"] = "Demo APP",
        };

        await using var s1 = File.OpenRead(templatePath);
        await using var outSingle = await WordForge.GenerateStreamAsync(s1, singleLinkValue);

        await using var s2 = File.OpenRead(templatePath);
        await using var outList = await WordForge.GenerateStreamAsync(s2, listLinkValue);

        using var s3 = File.OpenRead(templatePath);
        using var outArray = WordForge.GenerateStream(s3, arrayLinkValue);

        var snapSingle = DocxSnapshotHelper.Extract(outSingle);
        var snapList = DocxSnapshotHelper.Extract(outList);
        var snapArray = DocxSnapshotHelper.Extract(outArray);

        Assert.Equal(snapList.DocumentXml, snapArray.DocumentXml);
        Assert.Equal(snapList.DocumentRelsXml, snapArray.DocumentRelsXml);

        new
        {
            SingleHyperLink = snapSingle,
            MultipleHyperLinks = snapList,
        }.MatchSnapshot();
    }

    [Fact]
    public async Task TestIssue4_And_ColorTextVariants_TestBasicFillDocx_MatchesSnapshot()
    {
        var templatePath = DocxSnapshotHelper.GetSamplePath("TestBasicFill.docx");

        var plainValue = new
        {
            Company_Name = "MiniSofteware",
            Name = "Jack",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123,
            APP = "Demo APP",
        };

        var singleColorValue = new
        {
            Company_Name = new WordColorText { Text = "MiniSofteware", FontColor = "#eb70AB" },
            Name = new WordColorText { Text = "Jack", HighlightColor = "#eb70AB" },
            CreateDate = new WordColorText
            {
                Text = "2021-01-01 00:00:00",
                HighlightColor = "#eb70AB",
                FontColor = "#ffffff",
            },
            VIP = true,
            Points = 123,
            APP = "Demo APP",
        };

        var multiColorValue = new
        {
            Company_Name = new WordColorText { Text = "MiniSofteware", FontColor = "#eb70AB" },
            Name = new[]
            {
                new WordColorText { Text = "Ja", HighlightColor = "#eb70AB" },
                new WordColorText { Text = "ck", HighlightColor = "#a56abe" },
            },
            CreateDate = new WordColorText
            {
                Text = "2021-01-01 00:00:00",
                HighlightColor = "#eb70AB",
                FontColor = "#ffffff",
            },
            VIP = true,
            Points = 123,
            APP = "Demo APP",
        };

        await using var s1 = File.OpenRead(templatePath);
        await using var outPlain = await WordForge.GenerateStreamAsync(s1, plainValue);

        await using var s2 = File.OpenRead(templatePath);
        await using var outSingleColor = await WordForge.GenerateStreamAsync(s2, singleColorValue);

        await using var s3 = File.OpenRead(templatePath);
        await using var outMultiColor = await WordForge.GenerateStreamAsync(s3, multiColorValue);

        new
        {
            Plain = DocxSnapshotHelper.Extract(outPlain),
            SingleColor = DocxSnapshotHelper.Extract(outSingleColor),
            MultiColor = DocxSnapshotHelper.Extract(outMultiColor),
        }.MatchSnapshot();
    }

    [Fact]
    public async Task TestDemo01_And_TestDemo03_And_TestIssue43_Samples_MatchSnapshot()
    {
        // 1. TestDemo01.docx (has tags inside TextBoxContent + header1.xml + body)
        var demo01Path = DocxSnapshotHelper.GetSamplePath("TestDemo01.docx");
        var logoPath = DocxSnapshotHelper.GetSamplePath("DemoLogo.png");
        var demo01Value = new Dictionary<string, object?>
        {
            ["Image"] = new WordPicture(logoPath, 90, 90),
            ["Name"] = "Jane Doe",
            ["Position Title"] = "Principal Architect",
            ["Date"] = "2026-10-21",
            ["Recipient"] = "Hiring Manager",
            ["Company"] = "Acme Corp",
            ["Content"] = "Experienced systems engineer.",
        };

        await using var s1 = File.OpenRead(demo01Path);
        await using var outDemo01 = await WordForge.GenerateStreamAsync(s1, demo01Value);

        // 2. TestIssue43.docx (split '{{title}}' with bookmark inside paragraph)
        var issue43Path = DocxSnapshotHelper.GetSamplePath("TestIssue43.docx");
        await using var s2 = File.OpenRead(issue43Path);
        await using var outIssue43 = await WordForge.GenerateStreamAsync(s2, new { title = "Title43" });

        // 3. TestDemo03.docx (static document with anchor drawings)
        var demo03Path = DocxSnapshotHelper.GetSamplePath("TestDemo03.docx");
        await using var s3 = File.OpenRead(demo03Path);
        await using var outDemo03 = await WordForge.GenerateStreamAsync(s3, new { });

        var snapIssue43 = DocxSnapshotHelper.Extract(outIssue43);
        Assert.Equal(["MyTitle43"], snapIssue43.Paragraphs);

        new
        {
            Demo01 = DocxSnapshotHelper.Extract(outDemo01),
            Issue43 = snapIssue43,
            Demo03Paragraphs = DocxSnapshotHelper.Extract(outDemo03).Paragraphs,
        }.MatchSnapshot();
    }

    private static Dictionary<string, object?> CreateTripHoursWithForeachDetailsPayload()
    {
        return new Dictionary<string, object?>
        {
            ["TripHs"] = new List<Dictionary<string, object?>>
            {
                new()
                {
                    ["sDate"] = DateTime.Parse("2022-09-08 08:30:00"),
                    ["eDate"] = DateTime.Parse("2022-09-08 15:00:00"),
                    ["How"] = "Discussion requirement part1",
                    ["Details"] = CreateDetailsList(),
                },
                new()
                {
                    ["sDate"] = DateTime.Parse("2022-09-09 08:30:00"),
                    ["eDate"] = DateTime.Parse("2022-09-09 17:00:00"),
                    ["How"] = "Discussion requirement part2 and development",
                    ["Details"] = CreateDetailsList(),
                },
            },
        };

        static List<WordForeach> CreateDetailsList() =>
        [
            new()
            {
                Value = new Dictionary<string, object?> { ["Text"] = "Air", ["Value"] = "Airplane" },
                Separator = " | ",
            },
            new()
            {
                Value = new Dictionary<string, object?> { ["Text"] = "Parking", ["Value"] = "Car" },
                Separator = " / ",
            },
            new()
            {
                Value = new Dictionary<string, object?> { ["Text"] = "Hotel", ["Value"] = "Room" },
                Separator = ", ",
            },
            new()
            {
                Value = new Dictionary<string, object?> { ["Text"] = "Food", ["Value"] = "Plate" },
                Separator = "",
            },
        ];
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Foo
    {
        public string? title { get; set; }
        public List<User>? managers { get; set; }
        public List<User>? employees { get; set; }
    }

    private sealed class User
    {
        public string? name { get; set; }
        public string? department { get; set; }
    }
}
