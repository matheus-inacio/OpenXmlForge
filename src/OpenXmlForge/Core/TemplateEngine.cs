using System.Collections;
using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlForge.Images;
using OpenXmlForge.Layout;
using OpenXmlForge.Models;

namespace OpenXmlForge.Core;

internal static partial class TemplateEngine
{
    [GeneratedRegex(@"\{\{([^{}]+)\}\}")]
    private static partial Regex StandardTagRegex();

    [GeneratedRegex(@"(<[a-zA-Z/].*?>|\r\n|\n)")]
    private static partial Regex LineBreakSplitRegex();

    public static void ProcessDocumentSync(WordprocessingDocument docx, object? rawData)
    {
        var data = ValueResolver.NormalizeRoot(rawData);
        var pictures = new HashSet<WordPicture>(ReferenceEqualityComparer.Instance);
        ValueResolver.CollectPictures(data, pictures);

        var preparedImages = new Dictionary<WordPicture, ImageContent>(
            pictures.Count,
            ReferenceEqualityComparer.Instance);

        try
        {
            foreach (var pic in pictures)
            {
                var content = ImageContent.PrepareSync(pic);
                content.EnsureReusableForMultipleOccurrences();
                preparedImages[pic] = content;
            }

            ExecuteOnDocument(docx, data, preparedImages);
        }
        finally
        {
            foreach (var content in preparedImages.Values)
            {
                content.Dispose();
            }
        }
    }

    public static async Task ProcessDocumentAsync(
        WordprocessingDocument docx,
        object? rawData,
        CancellationToken cancellationToken = default)
    {
        var data = ValueResolver.NormalizeRoot(rawData);
        var pictures = new HashSet<WordPicture>(ReferenceEqualityComparer.Instance);
        ValueResolver.CollectPictures(data, pictures);

        var preparedImages = new ConcurrentDictionary<WordPicture, ImageContent>(
            ReferenceEqualityComparer.Instance);

        try
        {
            if (pictures.Count > 0)
            {
                var prepareTasks = pictures.Select(async pic =>
                {
                    var content = await ImageContent.PrepareAsync(pic, cancellationToken).ConfigureAwait(false);
                    await content.EnsureReusableForMultipleOccurrencesAsync(cancellationToken).ConfigureAwait(false);
                    preparedImages[pic] = content;
                });

                await Task.WhenAll(prepareTasks).ConfigureAwait(false);
            }

            ExecuteOnDocument(docx, data, preparedImages);
        }
        finally
        {
            foreach (var content in preparedImages.Values)
            {
                await content.DisposeAsync().ConfigureAwait(false);
            }
        }
    }

    private static void ExecuteOnDocument(
        WordprocessingDocument docx,
        Dictionary<string, object?> data,
        IReadOnlyDictionary<WordPicture, ImageContent> preparedImages)
    {
        if (docx.MainDocumentPart is not { } mainPart)
        {
            return;
        }

        var context = new RenderContext(
            preparedImages,
            new Dictionary<WordPicture, int>(ReferenceEqualityComparer.Instance));

        foreach (var headerPart in mainPart.HeaderParts.ToArray())
        {
            if (headerPart.Header is { } header)
            {
                GeneratePartRoot(header, headerPart, data, context);
            }
        }

        foreach (var footerPart in mainPart.FooterParts.ToArray())
        {
            if (footerPart.Footer is { } footer)
            {
                GeneratePartRoot(footer, footerPart, data, context);
            }
        }

        if (mainPart.Document?.Body is { } body)
        {
            GeneratePartRoot(body, mainPart, data, context);
        }

        docx.Save();
    }

    private static void GeneratePartRoot(
        OpenXmlElement root,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        // 1. Heal fragmented {{...}}, @foreach, @if tags across all paragraphs in this root
        TagDefragmenter.Defragment(root);

        // 2. Process TextBoxContent shapes from innermost to outermost so shape content
        //    never leaks into the outer paragraph or duplicates across Choice/Fallback branches.
        foreach (var textBox in root.Descendants<TextBoxContent>().Reverse().ToArray())
        {
            GenerateContainer(textBox, partContainer, scope, context);
        }

        // 3. Process the root container itself
        GenerateContainer(root, partContainer, scope, context);
    }

    private static void GenerateContainer(
        OpenXmlElement container,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        // 1. Block @foreach ... @endforeach statements
        ReplaceForeachBlocks(container, partContainer, scope, context);

        // 2. Tables belonging to this container (excluding tables inside nested TextBoxContent
        //    that were already processed, and handling nested tables cleanly)
        var topLevelTables = container
            .Descendants<Table>()
            .Where(t =>
                ReferenceEquals(GetContainerScopeRoot(t), container) &&
                !t.Ancestors<Table>().Any(parentTbl => ReferenceEquals(GetContainerScopeRoot(parentTbl), container)) &&
                t.InnerText.Contains("{{", StringComparison.Ordinal))
            .ToArray();

        foreach (var table in topLevelTables)
        {
            GenerateTable(table, partContainer, scope, context);
        }

        // 3. Block @if ... @endif statements
        ReplaceIfBlocks(container, scope);

        // 4. Paragraph text replacements for paragraphs belonging to this container ScopeRoot
        ReplaceTextInContainer(container, partContainer, scope, context);
    }

    /// <summary>
    /// Returns the nearest <see cref="TextBoxContent"/> ancestor if inside a shape,
    /// otherwise the root element (<see cref="Body"/>, <see cref="Header"/>, <see cref="Footer"/>).
    /// </summary>
    private static OpenXmlElement? GetContainerScopeRoot(OpenXmlElement element)
    {
        for (var parent = element.Parent; parent is not null; parent = parent.Parent)
        {
            if (parent is TextBoxContent or Body or Header or Footer)
            {
                return parent;
            }
        }

        return element is TextBoxContent or Body or Header or Footer
            ? element
            : element.Ancestors().LastOrDefault();
    }

    private static void GenerateTable(
        Table table,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        var rows = table.Elements<TableRow>().ToArray();
        var rowIndex = 0;

        while (rowIndex < rows.Length)
        {
            var tr = rows[rowIndex];
            var rowTags = ExtractOwnRowTags(tr);
            var rowCollections = FindRowCollectionsForTags(scope, rowTags);

            if (rowCollections.Count > 0)
            {
                // Group consecutive rows that continue a vertical merge (<w:vMerge/> without restart)
                // and bind to the same collection(s).
                var rowGroup = new List<TableRow> { tr };
                var primaryKey = rowCollections[0].Prefix;

                while (rowIndex + 1 < rows.Length && HasVerticalMergeContinuation(rows[rowIndex + 1]))
                {
                    var nextRow = rows[rowIndex + 1];
                    var nextTags = ExtractOwnRowTags(nextRow);
                    var nextCollections = FindRowCollectionsForTags(scope, nextTags);

                    if (nextCollections.Count == 0 ||
                        (nextCollections.Count == 1 && string.Equals(nextCollections[0].Prefix, primaryKey, StringComparison.Ordinal)))
                    {
                        rowGroup.Add(nextRow);
                        rowIndex++;
                    }
                    else
                    {
                        break;
                    }
                }

                var maxItemCount = 0;
                foreach (var rc in rowCollections)
                {
                    if (rc.Items.Count > maxItemCount)
                    {
                        maxItemCount = rc.Items.Count;
                    }
                }

                for (var itemIdx = 0; itemIdx < maxItemCount; itemIdx++)
                {
                    // Start from outer scope so scalar variables (e.g. {{Order.Number}} or {{Currency}})
                    // in the same row are also resolved!
                    var rowScope = new Dictionary<string, object?>(scope, StringComparer.Ordinal);

                    foreach (var rc in rowCollections)
                    {
                        if (itemIdx < rc.Items.Count)
                        {
                            BindRowCollectionItem(rowScope, rc.Prefix, rc.Items[itemIdx]);
                        }
                        else
                        {
                            // Exhausted collection in multi-collection row: bind empty string for its tags
                            rowScope[rc.Prefix] = null;
                            foreach (var tag in rowTags)
                            {
                                if (tag.StartsWith(rc.Prefix + ".", StringComparison.Ordinal))
                                {
                                    rowScope[tag] = string.Empty;
                                }
                            }
                        }
                    }

                    foreach (var templateRow in rowGroup)
                    {
                        var newTr = (TableRow)templateRow.CloneNode(true);

                        // Insert into the table BEFORE processing so newTr has its full parent chain
                        // (Table -> Body/Header/TextBoxContent -> SectionProperties) for scope & cell width measurement!
                        if (templateRow.Parent == table)
                        {
                            table.InsertBefore(newTr, rowGroup[0]);
                        }
                        else
                        {
                            table.Append(newTr);
                        }

                        ProcessNestedTablesInRow(newTr, table, partContainer, rowScope, context);
                        ReplaceIfBlocks(newTr, rowScope);
                        ReplaceTextInElement(newTr, partContainer, rowScope, context);
                    }
                }

                foreach (var templateRow in rowGroup)
                {
                    templateRow.Remove();
                }

                rowIndex++;
                continue;
            }

            // Static row (0 row collections — even if it has 1, 3, 5, or 50 dotted scalar variables!)
            ProcessNestedTablesInRow(tr, table, partContainer, scope, context);

            if (rowTags.Count > 0 || tr.InnerText.Contains("@if", StringComparison.Ordinal))
            {
                ReplaceIfBlocks(tr, scope);
                ReplaceTextInElement(tr, partContainer, scope, context);
            }

            rowIndex++;
        }
    }

    private static void ProcessNestedTablesInRow(
        TableRow row,
        Table parentTable,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        var nestedTables = row
            .Descendants<Table>()
            .Where(t =>
                ReferenceEquals(t.Ancestors<Table>().FirstOrDefault(), parentTable) &&
                t.InnerText.Contains("{{", StringComparison.Ordinal))
            .ToArray();

        foreach (var nestedTable in nestedTables)
        {
            GenerateTable(nestedTable, partContainer, scope, context);
        }
    }

    private static void BindRowCollectionItem(
        Dictionary<string, object?> rowScope,
        string collectionPrefix,
        object? item)
    {
        rowScope[collectionPrefix] = item;

        var itemDict = ValueResolver.ObjectToDictionary(item);
        foreach (var kvp in itemDict)
        {
            rowScope[$"{collectionPrefix}.{kvp.Key}"] = kvp.Value;
        }
    }

    private static List<string> ExtractOwnRowTags(TableRow row)
    {
        var tags = new List<string>();

        foreach (var paragraph in row.Descendants<Paragraph>())
        {
            if (!ReferenceEquals(paragraph.Ancestors<TableRow>().FirstOrDefault(), row))
            {
                continue;
            }

            var text = paragraph.InnerText;
            if (!text.Contains("{{", StringComparison.Ordinal))
            {
                continue;
            }

            var cleaned = text
                .Replace("{{foreach", "", StringComparison.Ordinal)
                .Replace("endforeach}}", "", StringComparison.Ordinal)
                .Replace("{{if(", "", StringComparison.Ordinal)
                .Replace(")if", "", StringComparison.Ordinal)
                .Replace("endif}}", "", StringComparison.Ordinal);

            foreach (Match match in StandardTagRegex().Matches(cleaned))
            {
                var tag = match.Groups[1].Value.Trim();
                if (tag.Length > 0 && !tags.Contains(tag, StringComparer.Ordinal))
                {
                    tags.Add(tag);
                }
            }
        }

        return tags;
    }

    private static List<RowCollectionBinding> FindRowCollectionsForTags(
        Dictionary<string, object?> scope,
        List<string> rowTags)
    {
        var bindings = new List<RowCollectionBinding>();

        foreach (var tag in rowTags)
        {
            // 1. If the full tag (even with dots, like "Order.Number" or "Customer.LegalName")
            //    exists directly in scope and is not a row collection, it is a scalar variable.
            if (scope.TryGetValue(tag, out var directVal) && !ValueResolver.IsRowCollection(directVal))
            {
                continue;
            }

            if (!tag.Contains('.'))
            {
                continue;
            }

            // 2. Check prefixes from shortest ("TripHs") to longest ("TripHs.Details")
            var dotPos = 0;
            while ((dotPos = tag.IndexOf('.', dotPos)) > 0)
            {
                var prefix = tag[..dotPos];
                dotPos++;

                if (bindings.Any(b => string.Equals(b.Prefix, prefix, StringComparison.Ordinal)))
                {
                    break;
                }

                if (ValueResolver.TryResolveValue(scope, prefix, out var prefixVal) &&
                    ValueResolver.IsRowCollection(prefixVal))
                {
                    var materialized = new List<object?>();
                    foreach (var element in (IEnumerable)prefixVal!)
                    {
                        materialized.Add(element);
                    }

                    bindings.Add(new RowCollectionBinding(prefix, materialized));
                    break;
                }
            }
        }

        return bindings;
    }

    private static bool HasVerticalMergeContinuation(TableRow row)
    {
        foreach (var cell in row.Elements<TableCell>())
        {
            var vMerge = cell.TableCellProperties?.VerticalMerge;
            if (vMerge is not null &&
                (vMerge.Val is null || vMerge.Val.Value == MergedCellValues.Continue))
            {
                return true;
            }
        }

        return false;
    }

    private static void ReplaceForeachBlocks(
        OpenXmlElement container,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        const string beginKey = "@foreach";
        const string endKey = "@endforeach";

        var betweenElements = GetBetweenElements(container, beginKey, endKey, out var beginParagraph, out var endParagraph);
        while (betweenElements is not null && beginParagraph is not null && endParagraph is not null)
        {
            string foreachDataKey;
            string? itemAlias = null;

            var match = StandardTagRegex().Match(beginParagraph.InnerText);
            if (match.Success)
            {
                foreachDataKey = match.Groups[1].Value.Trim();
            }
            else
            {
                // Also support "@foreach(item in Collection)" or "@foreach(Collection)"
                var tokens = beginParagraph.InnerText
                    .Replace("@foreach", "", StringComparison.Ordinal)
                    .Split([' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries);

                if (tokens.Length >= 3 && string.Equals(tokens[1], "in", StringComparison.OrdinalIgnoreCase))
                {
                    itemAlias = tokens[0];
                    foreachDataKey = tokens[2];
                }
                else if (tokens.Length >= 1)
                {
                    foreachDataKey = tokens[0];
                }
                else
                {
                    throw new InvalidOperationException("@foreach block is missing a valid collection variable.");
                }
            }

            beginParagraph.Remove();
            endParagraph.Remove();

            var foreachObj = ValueResolver.ResolveValueOrDefault(scope, foreachDataKey);
            var items = new List<object?>();

            if (foreachObj is IEnumerable enumerable and not string)
            {
                foreach (var item in enumerable)
                {
                    items.Add(item);
                }
            }

            if (items.Count > 0 && betweenElements.Count > 0)
            {
                var lastElementInLoop = betweenElements[^1];
                var templateClones = betweenElements.Select(e => e.CloneNode(true)).ToList();
                var currentLoopElements = new List<OpenXmlElement>(betweenElements);

                for (var i = 0; i < items.Count; i++)
                {
                    var itemScope = new Dictionary<string, object?>(scope, StringComparer.Ordinal);
                    var itemDict = ValueResolver.ObjectToDictionary(items[i]);

                    foreach (var kvp in itemDict)
                    {
                        itemScope[kvp.Key] = kvp.Value;
                        itemScope[$"{foreachDataKey}.{kvp.Key}"] = kvp.Value;
                        if (itemAlias is not null)
                        {
                            itemScope[$"{itemAlias}.{kvp.Key}"] = kvp.Value;
                        }
                    }

                    if (i > 0)
                    {
                        currentLoopElements.Clear();
                        foreach (var proto in templateClones)
                        {
                            var clone = proto.CloneNode(true);
                            lastElementInLoop.InsertAfterSelf(clone);
                            lastElementInLoop = clone;
                            currentLoopElements.Add(clone);
                        }
                    }

                    ReplaceIfBlocksInList(currentLoopElements, itemScope);

                    foreach (var ele in currentLoopElements)
                    {
                        if (ele.Parent is null)
                        {
                            continue;
                        }

                        if (ele is Table table)
                        {
                            GenerateTable(table, partContainer, itemScope, context);
                        }
                        else if (ele is Paragraph p)
                        {
                            ReplaceParagraphText(p, partContainer, itemScope, context);
                        }
                        else
                        {
                            ReplaceTextInElement(ele, partContainer, itemScope, context);
                        }
                    }
                }
            }
            else
            {
                foreach (var ele in betweenElements)
                {
                    ele.Remove();
                }
            }

            betweenElements = GetBetweenElements(container, beginKey, endKey, out beginParagraph, out endParagraph);
        }
    }

    private static List<OpenXmlElement>? GetBetweenElements(
        OpenXmlElement container,
        string beginKey,
        string endKey,
        out Paragraph? beginParagraph,
        out Paragraph? endParagraph)
    {
        beginParagraph = container
            .Elements<Paragraph>()
            .FirstOrDefault(p => p.InnerText.Contains(beginKey, StringComparison.Ordinal));

        endParagraph = null;

        if (beginParagraph is null)
        {
            return null;
        }

        var result = new List<OpenXmlElement>();
        for (var current = beginParagraph.NextSibling(); current is not null; current = current.NextSibling())
        {
            if (current is Paragraph p && p.InnerText.Contains(endKey, StringComparison.Ordinal))
            {
                endParagraph = p;
                return result;
            }

            result.Add(current);
        }

        return null;
    }

    private static void ReplaceIfBlocks(
        OpenXmlElement container,
        Dictionary<string, object?> scope)
    {
        var descendants = container.Descendants().ToList();
        ReplaceIfBlocksInList(descendants, scope);
    }

    private static void ReplaceIfBlocksInList(
        List<OpenXmlElement> elementList,
        Dictionary<string, object?> scope)
    {
        var paragraphs = elementList.OfType<Paragraph>().ToList();

        while (true)
        {
            var ifP = paragraphs.FirstOrDefault(p => p.InnerText.Contains("@if", StringComparison.Ordinal));
            if (ifP is null)
            {
                break;
            }

            var endIfP = paragraphs.FirstOrDefault(p => p.InnerText.Contains("@endif", StringComparison.Ordinal));
            if (endIfP is null)
            {
                paragraphs.Remove(ifP);
                continue;
            }

            var statement = ifP.InnerText.Trim().Split([' ', '(', ')'], StringSplitOptions.RemoveEmptyEntries);
            var isMatch = false;

            if (statement.Length >= 4)
            {
                var foundLeft = ValueResolver.TryResolveValue(scope, statement[1], out var leftVal);
                var foundRight = ValueResolver.TryResolveValue(scope, statement[3], out var rightVal);

                var leftStr = FormatStatementOperand(leftVal, foundLeft, statement[3]);
                var rightStr = foundRight ? FormatValueAsString(rightVal) : statement[3];

                isMatch = StatementEvaluator.EvaluateComparison(leftStr, statement[2], rightStr);
            }
            else if (statement.Length >= 2)
            {
                var rawVal = ValueResolver.ResolveValueOrDefault(scope, statement[1]);
                isMatch = rawVal is bool b ? b : bool.TryParse(rawVal?.ToString(), out var parsed) && parsed;
            }

            if (!isMatch)
            {
                var startIdx = elementList.IndexOf(ifP);
                var endIdx = elementList.IndexOf(endIfP);

                if (startIdx >= 0 && endIdx > startIdx)
                {
                    for (var i = startIdx + 1; i < endIdx; i++)
                    {
                        if (elementList[i].Parent is not null)
                        {
                            elementList[i].Remove();
                        }
                    }
                }
            }

            paragraphs.Remove(ifP);
            paragraphs.Remove(endIfP);

            if (ifP.Parent is not null)
            {
                ifP.Remove();
            }

            if (endIfP.Parent is not null)
            {
                endIfP.Remove();
            }
        }
    }

    private static string FormatStatementOperand(object? value, bool found, string otherRawOperand)
    {
        if (!found || value is null)
        {
            return "NULL";
        }

        return FormatValueAsString(value);
    }

    private static void ReplaceTextInContainer(
        OpenXmlElement container,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        if (container is Paragraph singleParagraph)
        {
            ReplaceParagraphText(singleParagraph, partContainer, scope, context);
            return;
        }

        var paragraphs = container
            .Descendants<Paragraph>()
            .Where(p => ReferenceEquals(GetContainerScopeRoot(p), container))
            .ToArray();

        foreach (var p in paragraphs)
        {
            ReplaceParagraphText(p, partContainer, scope, context);
        }
    }

    private static void ReplaceTextInElement(
        OpenXmlElement element,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        if (element is Paragraph singleParagraph)
        {
            ReplaceParagraphText(singleParagraph, partContainer, scope, context);
            return;
        }

        var elementScopeRoot = GetContainerScopeRoot(element);
        var paragraphs = element
            .Descendants<Paragraph>()
            .Where(p => ReferenceEquals(GetContainerScopeRoot(p), elementScopeRoot))
            .ToArray();

        foreach (var p in paragraphs)
        {
            ReplaceParagraphText(p, partContainer, scope, context);
        }
    }

    private static void ReplaceParagraphText(
        Paragraph paragraph,
        OpenXmlPartContainer partContainer,
        Dictionary<string, object?> scope,
        RenderContext context)
    {
        // Only process Runs whose nearest Paragraph ancestor is this exact paragraph!
        var runs = paragraph
            .Descendants<Run>()
            .Where(r => ReferenceEquals(r.Ancestors<Paragraph>().FirstOrDefault(), paragraph))
            .ToArray();

        foreach (var run in runs)
        {
            // Only process direct <w:t> children of this Run (never cross into nested shapes/drawings)
            var texts = run.Elements<Text>().ToArray();
            if (texts.Length == 0)
            {
                continue;
            }

            foreach (var t in texts)
            {
                var currentText = t.Text;
                if (string.IsNullOrEmpty(currentText) || !currentText.Contains("{{", StringComparison.Ordinal))
                {
                    continue;
                }

                // 1. Handle inline {{foreach ... endforeach}} (WordForeach)
                if (currentText.Contains("{{foreach", StringComparison.Ordinal) &&
                    currentText.Contains("endforeach}}", StringComparison.Ordinal))
                {
                    if (TryProcessInlineForeach(run, t, scope))
                    {
                        continue;
                    }
                }

                // 2. Find all {{key}} tags in t.Text
                var matches = StandardTagRegex().Matches(currentText);
                var removedTextNode = false;

                foreach (Match match in matches)
                {
                    var fullTag = match.Value;
                    var rawKey = match.Groups[1].Value;

                    if (!ValueResolver.TryResolveValue(scope, rawKey, out var value))
                    {
                        continue;
                    }

                    if (ValueResolver.IsStringEnumerable(value))
                    {
                        ExpandStringEnumerable(run, t, fullTag, (IEnumerable<string>)value!);
                        removedTextNode = true;
                        break;
                    }

                    if (value is WordHyperLink singleLink)
                    {
                        EmitHyperlinksInParagraph(paragraph, run, t, fullTag, partContainer, [singleLink]);
                        removedTextNode = true;
                        break;
                    }

                    if (value is IEnumerable<WordHyperLink> linkList)
                    {
                        EmitHyperlinksInParagraph(paragraph, run, t, fullTag, partContainer, linkList);
                        removedTextNode = true;
                        break;
                    }

                    if (value is WordColorText singleColor)
                    {
                        EmitRichContentWithSurroundingText(run, t, fullTag, r =>
                        {
                            OpenXmlElementBuilder.AppendColorTexts(r, [singleColor]);
                        });
                        removedTextNode = true;
                        break;
                    }

                    if (value is WordColorText[] colorArray)
                    {
                        EmitRichContentWithSurroundingText(run, t, fullTag, r =>
                        {
                            OpenXmlElementBuilder.AppendColorTexts(r, colorArray);
                        });
                        removedTextNode = true;
                        break;
                    }

                    if (value is IEnumerable<WordColorText> colorList)
                    {
                        EmitRichContentWithSurroundingText(run, t, fullTag, r =>
                        {
                            OpenXmlElementBuilder.AppendColorTexts(r, [.. colorList]);
                        });
                        removedTextNode = true;
                        break;
                    }

                    if (value is WordPicture singlePicture)
                    {
                        EmitRichContentWithSurroundingText(run, t, fullTag, r =>
                        {
                            AppendPictureToRun(r, partContainer, singlePicture, context);
                        });
                        removedTextNode = true;
                        break;
                    }

                    if (value is IEnumerable<WordPicture> pictureList)
                    {
                        EmitRichContentWithSurroundingText(run, t, fullTag, r =>
                        {
                            foreach (var pic in pictureList)
                            {
                                if (pic is not null)
                                {
                                    AppendPictureToRun(r, partContainer, pic, context);
                                    r.Append(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
                                }
                            }
                        });
                        removedTextNode = true;
                        break;
                    }

                    if (value is WordForeach singleForeach && singleForeach.Value is IEnumerable foreachEnumerable and not string)
                    {
                        var renderedItems = new List<string>();
                        foreach (var item in foreachEnumerable)
                        {
                            var dict = ValueResolver.ObjectToDictionary(item);
                            if (dict.TryGetValue(rawKey.Trim(), out var v))
                            {
                                renderedItems.Add(FormatValueAsString(v));
                            }
                            else if (dict.Count == 1)
                            {
                                renderedItems.Add(FormatValueAsString(dict.Values.First()));
                            }
                        }

                        var joined = string.Join(singleForeach.Separator ?? string.Empty, renderedItems);
                        t.Text = t.Text.Replace(fullTag, joined, StringComparison.Ordinal);
                        continue;
                    }

                    // Scalar value (or null when key is defined in scope)
                    var replacement = FormatValueAsString(value);
                    t.Text = t.Text.Replace(fullTag, replacement, StringComparison.Ordinal);
                }

                if (removedTextNode)
                {
                    continue;
                }

                // 3. Evaluate inline {{if(...)if...endif}} statements
                t.Text = StatementEvaluator.EvaluateInlineIfStatements(t.Text, scope);

                // 4. Split multiline strings (\r\n, \n, or <...> tags) into <w:t> + <w:br/> + <w:t>
                SplitMultilineTextInPlace(t);
            }
        }
    }

    private static bool TryProcessInlineForeach(
        Run run,
        Text t,
        Dictionary<string, object?> scope)
    {
        var rawText = t.Text;
        var startIdx = rawText.IndexOf("{{foreach", StringComparison.Ordinal);
        if (startIdx < 0)
        {
            return false;
        }

        var endTag = rawText.Contains("{{endforeach}}", StringComparison.Ordinal)
            ? "{{endforeach}}"
            : "endforeach}}";
        var endIdx = rawText.IndexOf(endTag, startIdx, StringComparison.Ordinal);
        if (endIdx < 0)
        {
            return false;
        }

        var prefixText = rawText[..startIdx];
        var suffixText = rawText[(endIdx + endTag.Length)..];
        var innerBlock = rawText[(startIdx + "{{foreach".Length)..endIdx];

        // Case A: "{{foreach item in Collection}}...{{endforeach}}"
        var closeHeaderIdx = innerBlock.IndexOf("}}", StringComparison.Ordinal);
        if (closeHeaderIdx > 0 && !innerBlock[..closeHeaderIdx].Contains("{{", StringComparison.Ordinal))
        {
            var headerTokens = innerBlock[..closeHeaderIdx].Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (headerTokens.Length >= 3 && string.Equals(headerTokens[1], "in", StringComparison.OrdinalIgnoreCase))
            {
                var itemAlias = headerTokens[0];
                var collectionKey = headerTokens[2];
                var bodyTemplate = innerBlock[(closeHeaderIdx + 2)..];

                if (ValueResolver.TryResolveValue(scope, collectionKey, out var colVal) &&
                    colVal is IEnumerable colEnum and not string)
                {
                    var sbAlias = new StringBuilder(prefixText);
                    foreach (var item in colEnum)
                    {
                        var itemDict = ValueResolver.ObjectToDictionary(item);
                        var itemText = bodyTemplate;
                        foreach (var kvp in itemDict)
                        {
                            var rep = FormatValueAsString(kvp.Value);
                            itemText = itemText
                                .Replace($"{{{{{itemAlias}.{kvp.Key}}}}}", rep, StringComparison.Ordinal)
                                .Replace($"{{{{{collectionKey}.{kvp.Key}}}}}", rep, StringComparison.Ordinal);
                        }

                        sbAlias.Append(StatementEvaluator.EvaluateInlineIfStatements(itemText, scope));
                    }

                    sbAlias.Append(suffixText);
                    var genText = new Text(sbAlias.ToString());
                    if (genText.Text.Length > 0 && (genText.Text[0] == ' ' || genText.Text[^1] == ' '))
                    {
                        genText.Space = SpaceProcessingModeValues.Preserve;
                    }

                    run.Append(genText);
                    t.Remove();
                    return true;
                }
            }
        }

        // Case B: Classic "{{foreach {{TripHs.Details.Text}} ... endforeach}}"
        var templateText = innerBlock;

        List<WordForeach>? foreachItems = null;
        string? matchedPrefix = null;

        foreach (Match m in StandardTagRegex().Matches(templateText))
        {
            var tagPath = m.Groups[1].Value.Trim();
            var lastDot = tagPath.LastIndexOf('.');
            if (lastDot <= 0)
            {
                continue;
            }

            var candidatePrefix = tagPath[..lastDot];
            if (ValueResolver.TryResolveValue(scope, candidatePrefix, out var resolved) &&
                resolved is IEnumerable<WordForeach> enumForeach)
            {
                foreachItems = [.. enumForeach];
                matchedPrefix = candidatePrefix;
                break;
            }
        }

        if (foreachItems is null || matchedPrefix is null)
        {
            return false;
        }

        var renderedSegments = new List<(int Index, string Text)>(foreachItems.Count);
        for (var i = 0; i < foreachItems.Count; i++)
        {
            var item = foreachItems[i];
            if (item?.Value is null)
            {
                continue;
            }

            var itemDict = ValueResolver.ObjectToDictionary(item.Value);
            var itemText = templateText;
            foreach (var kvp in itemDict)
            {
                var replacement = FormatValueAsString(kvp.Value);
                itemText = itemText.Replace($"{{{{{matchedPrefix}.{kvp.Key}}}}}", replacement, StringComparison.Ordinal);
            }

            itemText = StatementEvaluator.EvaluateInlineIfStatements(itemText, scope);
            if (!string.IsNullOrEmpty(itemText))
            {
                renderedSegments.Add((i, itemText));
            }
        }

        var sb = new StringBuilder(prefixText);
        for (var i = 0; i < renderedSegments.Count; i++)
        {
            sb.Append(renderedSegments[i].Text);
            if (i < renderedSegments.Count - 1)
            {
                sb.Append(foreachItems[renderedSegments[i].Index].Separator);
            }
        }
        sb.Append(suffixText);

        var generatedText = new Text(sb.ToString());
        if (generatedText.Text.Length > 0 && (generatedText.Text[0] == ' ' || generatedText.Text[^1] == ' '))
        {
            generatedText.Space = SpaceProcessingModeValues.Preserve;
        }

        run.Append(generatedText);
        t.Remove();
        return true;
    }

    private static void ExpandStringEnumerable(
        Run run,
        Text t,
        string fullTag,
        IEnumerable<string> values)
    {
        EmitRichContentWithSurroundingText(run, t, fullTag, r =>
        {
            var isFirst = true;
            foreach (var val in values)
            {
                var newT = (Text)t.CloneNode(true);
                newT.Text = StatementEvaluator.EvaluateInlineIfStatements(val ?? string.Empty);

                if (newT.Text.Length > 0 && (newT.Text[0] == ' ' || newT.Text[^1] == ' '))
                {
                    newT.Space = SpaceProcessingModeValues.Preserve;
                }

                if (isFirst)
                {
                    isFirst = false;
                }
                else
                {
                    r.Append(new Break());
                }

                r.Append(newT);
            }
        });
    }

    private static void EmitHyperlinksInParagraph(
        Paragraph paragraph,
        Run run,
        Text t,
        string fullTag,
        OpenXmlPartContainer partContainer,
        IEnumerable<WordHyperLink> links)
    {
        var text = t.Text;
        var tagIdx = text.IndexOf(fullTag, StringComparison.Ordinal);

        if (tagIdx > 0)
        {
            var prefix = text[..tagIdx];
            var prefixText = new Text(prefix);
            if (prefix[0] == ' ' || prefix[^1] == ' ')
            {
                prefixText.Space = SpaceProcessingModeValues.Preserve;
            }

            run.Append(prefixText);
        }

        OpenXmlElement anchor = run;
        var isFirst = true;
        foreach (var link in links)
        {
            if (link is null)
            {
                continue;
            }

            if (!isFirst)
            {
                var brRun = new Run(new Break());
                paragraph.InsertAfter(brRun, anchor);
                anchor = brRun;
            }

            var hyperlink = OpenXmlElementBuilder.CreateHyperlink(partContainer, link);
            paragraph.InsertAfter(hyperlink, anchor);
            anchor = hyperlink;
            isFirst = false;
        }

        var suffixStart = tagIdx + fullTag.Length;
        if (tagIdx >= 0 && suffixStart < text.Length)
        {
            var suffix = text[suffixStart..];
            var suffixText = new Text(suffix);
            if (suffix[0] == ' ' || suffix[^1] == ' ')
            {
                suffixText.Space = SpaceProcessingModeValues.Preserve;
            }

            var suffixRun = new Run();
            if (run.RunProperties is not null)
            {
                suffixRun.Append(run.RunProperties.CloneNode(true));
            }

            suffixRun.Append(suffixText);
            paragraph.InsertAfter(suffixRun, anchor);
        }

        t.Remove();
    }

    private static void EmitRichContentWithSurroundingText(
        Run run,
        Text t,
        string fullTag,
        Action<Run> emitAction)
    {
        var text = t.Text;
        var tagIdx = text.IndexOf(fullTag, StringComparison.Ordinal);

        if (tagIdx > 0)
        {
            var prefix = text[..tagIdx];
            var prefixText = new Text(prefix);
            if (prefix[0] == ' ' || prefix[^1] == ' ')
            {
                prefixText.Space = SpaceProcessingModeValues.Preserve;
            }

            run.Append(prefixText);
        }

        emitAction(run);

        var suffixStart = tagIdx + fullTag.Length;
        if (tagIdx >= 0 && suffixStart < text.Length)
        {
            var suffix = text[suffixStart..];
            var suffixText = new Text(suffix);
            if (suffix[0] == ' ' || suffix[^1] == ' ')
            {
                suffixText.Space = SpaceProcessingModeValues.Preserve;
            }

            run.Append(suffixText);
        }

        t.Remove();
    }

    private static void AppendPictureToRun(
        Run run,
        OpenXmlPartContainer partContainer,
        WordPicture picture,
        RenderContext context)
    {
        if (picture.Stream == Stream.Null && picture.Bytes is null && string.IsNullOrWhiteSpace(picture.Path))
        {
            return;
        }

        if (!context.PreparedImages.TryGetValue(picture, out var content))
        {
            content = ImageContent.PrepareSync(picture);
        }

        var (widthPx, heightPx) = ComputePictureDimensionsInPixels(run, picture, content);
        var cxEmus = widthPx * WordPicture.EmusPerPixel;
        var cyEmus = heightPx * WordPicture.EmusPerPixel;

        var partType = picture.GetImagePartType(content.DetectedFormat);
        var (imagePart, relationshipId) = AddImagePartToContainer(partContainer, partType);

        context.WriteCounts.TryGetValue(picture, out var writeCount);
        content.WriteTo(imagePart, ref writeCount);
        context.WriteCounts[picture] = writeCount;

        var drawing = picture.WrappingType == WordPictureWrappingType.Anchor
            ? OpenXmlElementBuilder.CreateAnchorPictureDrawing(relationshipId, cxEmus, cyEmus, picture)
            : OpenXmlElementBuilder.CreateInlinePictureDrawing(relationshipId, cxEmus, cyEmus, picture.Extension ?? string.Empty);

        run.Append(drawing);
    }

    private static (long WidthPx, long HeightPx) ComputePictureDimensionsInPixels(
        Run run,
        WordPicture picture,
        ImageContent content)
    {
        if (picture.HasExplicitDimensions)
        {
            return (picture.Width!.Value, picture.Height!.Value);
        }

        if (content.Dimensions is { } naturalDims)
        {
            // Find the innermost TableCell containing this run (handles nested tables automatically)
            var containingCell = run.Ancestors<TableCell>().FirstOrDefault();
            var cellWidthPx = TableCellWidthMeasurer.MeasureUsableWidthInPixels(containingCell);

            var scale = cellWidthPx is { } maxPx and > 0
                ? Math.Min(1.0, (double)maxPx / naturalDims.Width)
                : 1.0;

            return (
                Math.Max(1L, (long)Math.Round(naturalDims.Width * scale, MidpointRounding.AwayFromZero)),
                Math.Max(1L, (long)Math.Round(naturalDims.Height * scale, MidpointRounding.AwayFromZero)));
        }

        return (
            picture.Width is > 0 ? picture.Width.Value : WordPicture.DefaultDimensionPixels,
            picture.Height is > 0 ? picture.Height.Value : WordPicture.DefaultDimensionPixels);
    }

    private static (ImagePart Part, string RelationshipId) AddImagePartToContainer(
        OpenXmlPartContainer partContainer,
        PartTypeInfo partType)
    {
        return partContainer switch
        {
            MainDocumentPart main => CreatePart(main.AddImagePart(partType), main),
            HeaderPart header => CreatePart(header.AddImagePart(partType), header),
            FooterPart footer => CreatePart(footer.AddImagePart(partType), footer),
            _ => throw new InvalidOperationException("Unsupported OpenXmlPartContainer for ImagePart."),
        };

        static (ImagePart, string) CreatePart(ImagePart part, OpenXmlPartContainer container)
            => (part, container.GetIdOfPart(part));
    }

    private static void SplitMultilineTextInPlace(Text t)
    {
        var text = t.Text;
        if (string.IsNullOrEmpty(text) ||
            (!text.Contains('\n') && !text.Contains('<')))
        {
            if (!string.IsNullOrEmpty(text) && (text[0] == ' ' || text[^1] == ' '))
            {
                t.Space = SpaceProcessingModeValues.Preserve;
            }

            return;
        }

        var splits = LineBreakSplitRegex()
            .Split(text)
            .Where(s => s is not "\n" and not "\r\n")
            .ToArray();

        if (splits.Length <= 1)
        {
            if (text.Length > 0 && (text[0] == ' ' || text[^1] == ' '))
            {
                t.Space = SpaceProcessingModeValues.Preserve;
            }

            return;
        }

        var isFirst = true;
        foreach (var segment in splits)
        {
            var newT = (Text)t.CloneNode(true);
            newT.Text = segment ?? string.Empty;
            if (newT.Text.Length > 0 && (newT.Text[0] == ' ' || newT.Text[^1] == ' '))
            {
                newT.Space = SpaceProcessingModeValues.Preserve;
            }

            if (isFirst)
            {
                isFirst = false;
            }
            else
            {
                t.InsertBeforeSelf(new Break());
            }

            t.InsertBeforeSelf(newT);
        }

        t.Remove();
    }

    private static string FormatValueAsString(object? value) => value switch
    {
        null => string.Empty,
        string s => s,
        DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        DateTimeOffset dto => dto.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        _ when ValueResolver.IsRowCollection(value) => string.Empty,
        _ => value.ToString() ?? string.Empty,
    };

    private sealed record RowCollectionBinding(string Prefix, List<object?> Items);

    private sealed record RenderContext(
        IReadOnlyDictionary<WordPicture, ImageContent> PreparedImages,
        Dictionary<WordPicture, int> WriteCounts);
}
