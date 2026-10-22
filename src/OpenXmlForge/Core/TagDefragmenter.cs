using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlForge.Core;

/// <summary>
/// Consolidates template tags (<c>{{...}}</c>, <c>@foreach</c>, <c>@if</c>) that Word, Google Docs,
/// or LibreOffice split across multiple <see cref="Run"/> / <see cref="Text"/> elements and strips
/// <see cref="ProofError"/> markers, eliminating the need for external XML regex pre-cleaners.
/// </summary>
internal static class TagDefragmenter
{
    public static void Defragment(OpenXmlElement root)
    {
        if (root is Paragraph singleParagraph)
        {
            DefragmentParagraph(singleParagraph);
            return;
        }

        foreach (var paragraph in root.Descendants<Paragraph>().ToArray())
        {
            DefragmentParagraph(paragraph);
        }
    }

    public static void DefragmentParagraph(Paragraph paragraph)
    {
        // Remove spell/grammar proof errors belonging directly to this paragraph
        foreach (var proofErr in paragraph.Descendants<ProofError>()
                     .Where(pe => ReferenceEquals(pe.Ancestors<Paragraph>().FirstOrDefault(), paragraph))
                     .ToArray())
        {
            proofErr.Remove();
        }

        var textNodes = paragraph.Descendants<Text>()
            .Where(t => ReferenceEquals(t.Ancestors<Paragraph>().FirstOrDefault(), paragraph))
            .ToArray();

        if (textNodes.Length == 0)
        {
            return;
        }

        var hasBraceOrAt = false;
        var totalLength = 0;

        foreach (var t in textNodes)
        {
            var txt = t.Text;
            if (!string.IsNullOrEmpty(txt))
            {
                totalLength += txt.Length;
                if (!hasBraceOrAt && (txt.Contains('{') || txt.Contains('@')))
                {
                    hasBraceOrAt = true;
                }
            }
        }

        if (!hasBraceOrAt || totalLength == 0)
        {
            return;
        }

        var fullTextBuilder = new StringBuilder(totalLength);
        var charOwner = new int[totalLength];
        var pos = 0;

        for (var nodeIdx = 0; nodeIdx < textNodes.Length; nodeIdx++)
        {
            var txt = textNodes[nodeIdx].Text ?? string.Empty;
            fullTextBuilder.Append(txt);
            Array.Fill(charOwner, nodeIdx, pos, txt.Length);
            pos += txt.Length;
        }

        var fullText = fullTextBuilder.ToString();
        var trimmedStart = fullText.AsSpan().TrimStart();

        // Block directives (@foreach, @endforeach, @if, @endif) occupy the whole paragraph
        if (trimmedStart.StartsWith("@foreach", StringComparison.Ordinal) ||
            trimmedStart.StartsWith("@endforeach", StringComparison.Ordinal) ||
            trimmedStart.StartsWith("@if", StringComparison.Ordinal) ||
            trimmedStart.StartsWith("@endif", StringComparison.Ordinal))
        {
            if (textNodes.Length == 1)
            {
                return;
            }

            ApplySingleNodeConsolidation(textNodes, fullText);
            return;
        }

        // Scan for {{foreach...endforeach}}, {{if(...endif}}, and {{...}} spans
        var segments = new List<TagSegment>();
        var idx = 0;
        var modified = false;

        while (idx < fullText.Length)
        {
            if (fullText[idx] != '{')
            {
                idx++;
                continue;
            }

            // Try matching opening "{{" allowing optional whitespace between the two '{'
            if (!TryMatchDoubleOpenBrace(fullText, idx, out var afterOpenIdx, out var openHadSpaces))
            {
                idx++;
                continue;
            }

            var restAfterOpen = fullText.AsSpan(afterOpenIdx);

            // 1. Inline {{foreach ... endforeach}}
            if (restAfterOpen.StartsWith("foreach", StringComparison.Ordinal))
            {
                var endForeachIdx = fullText.IndexOf("endforeach}}", afterOpenIdx + 7, StringComparison.Ordinal);
                if (endForeachIdx >= 0)
                {
                    var matchEnd = endForeachIdx + "endforeach}}".Length;
                    var targetOwner = charOwner[idx];
                    for (var c = idx; c < matchEnd; c++)
                    {
                        if (charOwner[c] != targetOwner)
                        {
                            charOwner[c] = targetOwner;
                            modified = true;
                        }
                    }

                    idx = matchEnd;
                    continue;
                }
            }

            // 2. Inline {{if( ... endif}}
            if (restAfterOpen.StartsWith("if(", StringComparison.Ordinal))
            {
                var endIfIdx = fullText.IndexOf("endif}}", afterOpenIdx + 3, StringComparison.Ordinal);
                if (endIfIdx >= 0)
                {
                    var matchEnd = endIfIdx + "endif}}".Length;
                    var targetOwner = charOwner[idx];
                    for (var c = idx; c < matchEnd; c++)
                    {
                        if (charOwner[c] != targetOwner)
                        {
                            charOwner[c] = targetOwner;
                            modified = true;
                        }
                    }

                    idx = matchEnd;
                    continue;
                }
            }

            // 3. Standard variable tag {{ ... }} (allowing optional whitespace between the two '}')
            if (TryMatchStandardTagClose(fullText, afterOpenIdx, out var innerEndIdx, out var closeEndIdx, out var closeHadSpaces))
            {
                var rawInner = fullText[afterOpenIdx..innerEndIdx];
                var normalizedInner = NormalizeTagInner(rawInner);
                var normalizedTag = $"{{{{{normalizedInner}}}}}";

                var targetOwner = charOwner[idx];
                for (var c = idx; c < closeEndIdx; c++)
                {
                    if (charOwner[c] != targetOwner)
                    {
                        charOwner[c] = targetOwner;
                        modified = true;
                    }
                }

                if (openHadSpaces || closeHadSpaces || !string.Equals(rawInner, normalizedInner, StringComparison.Ordinal))
                {
                    modified = true;
                }

                segments.Add(new TagSegment(idx, closeEndIdx, targetOwner, normalizedTag));
                idx = closeEndIdx;
                continue;
            }

            idx++;
        }

        if (!modified)
        {
            return;
        }

        // Reconstruct text per Text node
        var perNodeBuilders = new StringBuilder[textNodes.Length];
        for (var i = 0; i < textNodes.Length; i++)
        {
            perNodeBuilders[i] = new StringBuilder();
        }

        var segIdx = 0;
        var charIdx = 0;

        while (charIdx < fullText.Length)
        {
            if (segIdx < segments.Count && charIdx == segments[segIdx].Start)
            {
                var seg = segments[segIdx];
                perNodeBuilders[seg.OwnerNodeIndex].Append(seg.NormalizedTag);
                charIdx = seg.End;
                segIdx++;
                continue;
            }

            perNodeBuilders[charOwner[charIdx]].Append(fullText[charIdx]);
            charIdx++;
        }

        for (var i = 0; i < textNodes.Length; i++)
        {
            var newText = perNodeBuilders[i].ToString();
            var textNode = textNodes[i];

            if (newText.Length == 0)
            {
                var parentRun = textNode.Parent as Run;
                textNode.Remove();

                if (parentRun is not null && !parentRun.Elements().Any(e => e is not RunProperties))
                {
                    parentRun.Remove();
                }
            }
            else
            {
                textNode.Text = newText;
                if (newText[0] == ' ' || newText[^1] == ' ')
                {
                    textNode.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }
    }

    private static void ApplySingleNodeConsolidation(Text[] textNodes, string fullText)
    {
        textNodes[0].Text = fullText;
        if (fullText[0] == ' ' || fullText[^1] == ' ')
        {
            textNodes[0].Space = SpaceProcessingModeValues.Preserve;
        }

        for (var i = 1; i < textNodes.Length; i++)
        {
            var parentRun = textNodes[i].Parent as Run;
            textNodes[i].Remove();

            if (parentRun is not null && !parentRun.Elements().Any(e => e is not RunProperties))
            {
                parentRun.Remove();
            }
        }
    }

    private static bool TryMatchDoubleOpenBrace(
        string text,
        int firstBraceIdx,
        out int afterSecondBraceIdx,
        out bool hadWhitespaceBetweenBraces)
    {
        afterSecondBraceIdx = -1;
        hadWhitespaceBetweenBraces = false;

        var p = firstBraceIdx + 1;
        while (p < text.Length && char.IsWhiteSpace(text[p]))
        {
            hadWhitespaceBetweenBraces = true;
            p++;
        }

        if (p < text.Length && text[p] == '{')
        {
            afterSecondBraceIdx = p + 1;
            return true;
        }

        return false;
    }

    private static bool TryMatchStandardTagClose(
        string text,
        int startIdx,
        out int innerEndIdx,
        out int afterSecondCloseBraceIdx,
        out bool hadWhitespaceBetweenBraces)
    {
        innerEndIdx = -1;
        afterSecondCloseBraceIdx = -1;
        hadWhitespaceBetweenBraces = false;

        var p = startIdx;
        while (p < text.Length)
        {
            var ch = text[p];
            if (ch == '{')
            {
                // Encountered another '{' before closing '}}' -> not a flat standard tag
                return false;
            }

            if (ch == '}')
            {
                var q = p + 1;
                var spaces = false;
                while (q < text.Length && char.IsWhiteSpace(text[q]))
                {
                    spaces = true;
                    q++;
                }

                if (q < text.Length && text[q] == '}' && p > startIdx)
                {
                    innerEndIdx = p;
                    afterSecondCloseBraceIdx = q + 1;
                    hadWhitespaceBetweenBraces = spaces;
                    return true;
                }

                return false;
            }

            p++;
        }

        return false;
    }

    private static string NormalizeTagInner(string rawInner)
    {
        var trimmed = rawInner.Trim();
        if (trimmed.Contains('.') && trimmed.AsSpan().IndexOfAny(" \t\r\n") >= 0)
        {
            Span<char> buf = stackalloc char[trimmed.Length];
            var len = 0;
            foreach (var c in trimmed)
            {
                if (!char.IsWhiteSpace(c))
                {
                    buf[len++] = c;
                }
            }

            return new string(buf[..len]);
        }

        return trimmed;
    }

    private readonly record struct TagSegment(
        int Start,
        int End,
        int OwnerNodeIndex,
        string NormalizedTag);
}
