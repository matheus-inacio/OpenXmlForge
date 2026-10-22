using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlForge.Tests;

internal static partial class DocxSnapshotHelper
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    [GeneratedRegex(@"Picture [0-9a-fA-F-]{32,36}")]
    private static partial Regex PictureGuidRegex();

    [GeneratedRegex(@"Image [0-9a-fA-F-]{32,36}(?=\.)")]
    private static partial Regex ImageGuidRegex();

    [GeneratedRegex(@"uri=""\{[0-9a-fA-F-]{32,36}\}""")]
    private static partial Regex BlipUriGuidRegex();

    public static string GetSamplePath(string fileName)
        => Path.Combine(AppContext.BaseDirectory, "Samples", fileName);

    public static string GetTempFilePath(string extension = "docx")
        => Path.Combine(Path.GetTempPath(), $"OpenXmlForge_{Guid.NewGuid():N}.{extension}");

    public static string GetZipFileContent(string zipPath, string entryPath)
    {
        using var stream = File.OpenRead(zipPath);
        return GetZipFileContent(stream, entryPath);
    }

    public static string GetZipFileContent(Stream zipStream, string entryPath)
    {
        if (zipStream.CanSeek)
        {
            zipStream.Position = 0;
        }

        using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true, Encoding.UTF8);
        var entry = archive.Entries.Single(e => e.FullName == entryPath);
        using var entryStream = entry.Open();
        var doc = XDocument.Load(entryStream);
        return doc.ToString();
    }

    public static DocxSnapshot Extract(string docxPath)
    {
        using var fs = File.OpenRead(docxPath);
        return Extract(fs);
    }

    public static DocxSnapshot Extract(Stream docxStream)
    {
        using var ms = new MemoryStream();
        if (docxStream.CanSeek)
        {
            docxStream.Position = 0;
        }

        docxStream.CopyTo(ms);
        ms.Position = 0;

        List<string> paragraphs;
        List<List<List<string>>> tables;

        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            paragraphs = body?
                .Elements<Paragraph>()
                .Select(p => p.InnerText)
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList() ?? [];

            tables = body?
                .Descendants<Table>()
                .Select(t => t.Elements<TableRow>()
                    .Select(r => r.Elements<TableCell>()
                        .Select(c => c.InnerText)
                        .ToList())
                    .ToList())
                .ToList() ?? [];
        }

        ms.Position = 0;
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true, Encoding.UTF8);

        var mediaHashByPath = new Dictionary<string, MediaSnapshot>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries.Where(e => e.FullName.Contains("media/", StringComparison.OrdinalIgnoreCase)))
        {
            using var s = entry.Open();
            using var buf = new MemoryStream();
            s.CopyTo(buf);
            var bytes = buf.ToArray();
            var hash = Convert.ToHexString(SHA256.HashData(bytes));
            var snapshot = new MediaSnapshot(
                Path.GetExtension(entry.Name).ToLowerInvariant(),
                bytes.Length,
                hash);
            mediaHashByPath["/" + entry.FullName.TrimStart('/')] = snapshot;
            mediaHashByPath[entry.FullName.TrimStart('/')] = snapshot;
            mediaHashByPath[entry.Name] = snapshot;
        }

        var (documentXml, documentRelsXml, mediaList) = NormalizePartAndRels(
            archive,
            "word/document.xml",
            "word/_rels/document.xml.rels",
            mediaHashByPath);

        var headers = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries
                     .Where(e => e.FullName.StartsWith("word/header", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal))
                     .OrderBy(e => e.FullName, StringComparer.Ordinal))
        {
            var relsPath = $"word/_rels/{entry.Name}.rels";
            var (hdrXml, _, hdrMedia) = NormalizePartAndRels(archive, entry.FullName, relsPath, mediaHashByPath);
            headers[entry.Name] = hdrXml;
            mediaList.AddRange(hdrMedia);
        }

        var footers = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries
                     .Where(e => e.FullName.StartsWith("word/footer", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal))
                     .OrderBy(e => e.FullName, StringComparer.Ordinal))
        {
            var relsPath = $"word/_rels/{entry.Name}.rels";
            var (ftrXml, _, ftrMedia) = NormalizePartAndRels(archive, entry.FullName, relsPath, mediaHashByPath);
            footers[entry.Name] = ftrXml;
            mediaList.AddRange(ftrMedia);
        }

        return new DocxSnapshot(
            Paragraphs: paragraphs,
            Tables: tables,
            Media: mediaList,
            DocumentXml: documentXml,
            DocumentRelsXml: documentRelsXml,
            Headers: headers.Count > 0 ? headers : null,
            Footers: footers.Count > 0 ? footers : null);
    }

    private static (string PartXml, string? RelsXml, List<MediaSnapshot> ReferencedMedia) NormalizePartAndRels(
        ZipArchive archive,
        string partEntryPath,
        string relsEntryPath,
        Dictionary<string, MediaSnapshot> mediaByPath)
    {
        var partEntry = archive.GetEntry(partEntryPath)
            ?? throw new InvalidOperationException($"Missing entry '{partEntryPath}'.");

        XDocument partDoc;
        using (var s = partEntry.Open())
        {
            partDoc = XDocument.Load(s);
        }

        var partRaw = partDoc.ToString();
        partRaw = PictureGuidRegex().Replace(partRaw, "Picture {GUID}");
        partRaw = ImageGuidRegex().Replace(partRaw, "Image {GUID}");
        partRaw = BlipUriGuidRegex().Replace(partRaw, @"uri=""{GUID}""");
        partDoc = XDocument.Parse(partRaw);

        // Collect all relationship IDs referenced in this XML part in document order
        var referencedRelIdsInOrder = new List<string>();
        foreach (var attr in partDoc.Descendants().SelectMany(e => e.Attributes()))
        {
            if (attr.Name == R + "id" || attr.Name == R + "embed" || attr.Name == R + "link")
            {
                if (!referencedRelIdsInOrder.Contains(attr.Value, StringComparer.Ordinal))
                {
                    referencedRelIdsInOrder.Add(attr.Value);
                }
            }
        }

        var relsEntry = archive.GetEntry(relsEntryPath);
        if (relsEntry is null)
        {
            return (partDoc.ToString(), null, []);
        }

        XDocument relsDoc;
        using (var s = relsEntry.Open())
        {
            relsDoc = XDocument.Load(s);
        }

        var relElements = relsDoc.Root?.Elements(RelNs + "Relationship").ToList() ?? [];
        var idMap = new Dictionary<string, string>(StringComparer.Ordinal);
        var mediaTargetMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var referencedMedia = new List<MediaSnapshot>();

        var genRelCounter = 1;
        var genMediaCounter = 1;

        // First map referenced relationship IDs in exact order of appearance in partDoc
        foreach (var oldId in referencedRelIdsInOrder)
        {
            var relEl = relElements.FirstOrDefault(e => (string?)e.Attribute("Id") == oldId);
            if (relEl is null)
            {
                continue;
            }

            if (!Regex.IsMatch(oldId, @"^rId\d+$"))
            {
                idMap[oldId] = $"rId_gen_{genRelCounter++}";
            }

            var target = (string?)relEl.Attribute("Target");
            if (!string.IsNullOrEmpty(target) &&
                target.Contains("media/", StringComparison.OrdinalIgnoreCase))
            {
                var normalizedLookupKey = target.StartsWith('/')
                    ? target
                    : "/word/" + target.TrimStart('/');

                if (mediaByPath.TryGetValue(normalizedLookupKey, out var mediaInfo) ||
                    mediaByPath.TryGetValue(Path.GetFileName(target), out mediaInfo))
                {
                    referencedMedia.Add(mediaInfo);
                }

                if (!mediaTargetMap.ContainsKey(target))
                {
                    var ext = Path.GetExtension(target).ToLowerInvariant();
                    mediaTargetMap[target] = $"/word/media/image_gen_{genMediaCounter++}{ext}";
                }
            }
        }

        // Map any remaining non-standard relationship IDs deterministically by (Type, Target)
        foreach (var relEl in relElements
                     .OrderBy(e => (string?)e.Attribute("Type"), StringComparer.Ordinal)
                     .ThenBy(e => (string?)e.Attribute("Target"), StringComparer.Ordinal))
        {
            var oldId = (string?)relEl.Attribute("Id");
            if (oldId is not null && !idMap.ContainsKey(oldId) && !Regex.IsMatch(oldId, @"^rId\d+$"))
            {
                idMap[oldId] = $"rId_gen_{genRelCounter++}";
            }
        }

        // Apply relationship ID mapping to partDoc
        if (idMap.Count > 0)
        {
            foreach (var attr in partDoc.Descendants().SelectMany(e => e.Attributes()))
            {
                if ((attr.Name == R + "id" || attr.Name == R + "embed" || attr.Name == R + "link") &&
                    idMap.TryGetValue(attr.Value, out var newId))
                {
                    attr.Value = newId;
                }
            }
        }

        // Apply relationship ID and media target mapping to relsDoc, then sort relationships by Id
        foreach (var relEl in relElements)
        {
            var idAttr = relEl.Attribute("Id");
            if (idAttr is not null && idMap.TryGetValue(idAttr.Value, out var newId))
            {
                idAttr.Value = newId;
            }

            var targetAttr = relEl.Attribute("Target");
            if (targetAttr is not null && mediaTargetMap.TryGetValue(targetAttr.Value, out var newTarget))
            {
                targetAttr.Value = newTarget;
            }
        }

        if (relsDoc.Root is not null)
        {
            var sorted = relElements
                .OrderBy(e => (string?)e.Attribute("Id"), StringComparer.Ordinal)
                .ToList();
            relsDoc.Root.ReplaceNodes(sorted);
        }

        return (partDoc.ToString(), relsDoc.ToString(), referencedMedia);
    }
}

public sealed record DocxSnapshot(
    IReadOnlyList<string> Paragraphs,
    IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> Tables,
    IReadOnlyList<MediaSnapshot> Media,
    string DocumentXml,
    string? DocumentRelsXml,
    IReadOnlyDictionary<string, string>? Headers = null,
    IReadOnlyDictionary<string, string>? Footers = null);

public sealed record MediaSnapshot(
    string Extension,
    int ByteLength,
    string Sha256);
