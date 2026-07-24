using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Variables;

public class DocxVariableReplacer
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    /// <summary>
    /// Replaces {{variable}} and {{variable|default}} placeholders in the body and in all
    /// header/footer parts. Null-value semantics: a data entry whose value is null is
    /// treated as missing — the placeholder's default is used when specified, otherwise
    /// the placeholder is preserved as-is.
    /// </summary>
    public void Replace(WordprocessingDocument document, Dictionary<string, string> data)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            ReplaceInElement(body, data);
        }

        var headers = document.MainDocumentPart?.HeaderParts;
        if (headers != null)
        {
            foreach (var header in headers)
            {
                if (header.Header != null)
                {
                    ReplaceInElement(header.Header, data);
                }
            }
        }

        var footers = document.MainDocumentPart?.FooterParts;
        if (footers != null)
        {
            foreach (var footer in footers)
            {
                if (footer.Footer != null)
                {
                    ReplaceInElement(footer.Footer, data);
                }
            }
        }
    }

    private void ReplaceInElement(OpenXmlElement element, Dictionary<string, string> data)
    {
        var paragraphs = element.Descendants<Paragraph>();

        foreach (var paragraph in paragraphs)
        {
            // Descendants (not Elements) so runs nested inside hyperlinks are covered too.
            var runs = paragraph.Descendants<Run>().ToList();

            foreach (var run in runs)
            {
                var texts = run.Elements<Text>().ToList();

                foreach (var text in texts)
                {
                    if (text.Text.Contains("{{"))
                    {
                        var replaced = ReplaceVariablesInText(text.Text, data);
                        if (!string.Equals(replaced, text.Text, StringComparison.Ordinal))
                        {
                            text.Text = replaced;
                            // Replaced values may lead/trail with spaces; without preserve
                            // Word collapses them, matching the cross-run path below.
                            text.Space = SpaceProcessingModeValues.Preserve;
                        }
                    }
                }
            }

            ReplaceVariablesAcrossRuns(paragraph, data);
        }
    }

    private void ReplaceVariablesAcrossRuns(Paragraph paragraph, Dictionary<string, string> data)
    {
        var textNodes = paragraph.Descendants<Text>().ToList();
        if (textNodes.Count <= 1)
        {
            return;
        }

        var combinedText = string.Concat(textNodes.Select(t => t.Text));
        if (!combinedText.Contains("{{"))
        {
            return;
        }

        var matches = VariablePattern.Matches(combinedText);
        if (matches.Count == 0)
        {
            return;
        }

        var positions = new List<(Text text, int start, int length)>();
        int offset = 0;
        foreach (var t in textNodes)
        {
            positions.Add((t, offset, t.Text.Length));
            offset += t.Text.Length;
        }

        for (int m = matches.Count - 1; m >= 0; m--)
        {
            var match = matches[m];
            var replacement = GetReplacementValue(match, data);
            if (replacement == null)
            {
                continue;
            }

            int varStart = match.Index;
            int varEnd = match.Index + match.Length;

            var firstNode = default((Text text, int start, int length));
            int firstIdx = -1;
            for (int i = 0; i < positions.Count; i++)
            {
                if (positions[i].start + positions[i].length > varStart)
                {
                    firstNode = positions[i];
                    firstIdx = i;
                    break;
                }
            }

            if (firstIdx < 0)
            {
                continue;
            }

            if (firstNode.text == null || firstNode.text.Text == null)
            {
                continue;
            }

            int localVarStart = varStart - firstNode.start;
            int localVarEnd = varEnd - firstNode.start;

            string before = localVarStart > 0 ? firstNode.text.Text.Substring(0, localVarStart) : string.Empty;
            string after = localVarEnd < firstNode.text.Text.Length ? firstNode.text.Text.Substring(localVarEnd) : string.Empty;
            firstNode.text.Text = before + replacement + after;
            firstNode.text.Space = SpaceProcessingModeValues.Preserve;
            positions[firstIdx] = (firstNode.text, firstNode.start, firstNode.text.Text.Length);

            for (int i = firstIdx + 1; i < positions.Count; i++)
            {
                var node = positions[i];
                int nodeStart = node.start;
                int nodeEnd = nodeStart + node.length;

                if (nodeStart >= varEnd)
                {
                    break;
                }

                int keepStart = varStart - nodeStart;
                int keepEnd = varEnd - nodeStart;

                if (keepStart <= 0 && keepEnd >= node.length)
                {
                    node.text.Text = string.Empty;
                }
                else
                {
                    string keepBefore = keepStart > 0 ? node.text.Text.Substring(0, keepStart) : string.Empty;
                    string keepAfter = keepEnd < node.text.Text.Length ? node.text.Text.Substring(keepEnd) : string.Empty;
                    node.text.Text = keepBefore + keepAfter;
                }

                positions[i] = (node.text, positions[i].start, node.text.Text.Length);
            }

            int newOffset = 0;
            for (int i = 0; i < positions.Count; i++)
            {
                positions[i] = (positions[i].text, newOffset, positions[i].text.Text.Length);
                newOffset += positions[i].text.Text.Length;
            }
        }
    }

    private static string? GetReplacementValue(Match match, Dictionary<string, string> data)
    {
        var variableName = match.Groups[1].Value.Trim();
        var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;

        // Null data value == missing: fall through to the default, else preserve.
        if (data.TryGetValue(variableName, out var value) && value != null)
        {
            return value;
        }

        if (defaultValue != null)
        {
            return defaultValue;
        }

        return null;
    }

    private string ReplaceVariablesInText(string text, Dictionary<string, string> data)
    {
        return VariablePattern.Replace(text, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;

            // Null data value == missing, consistent with the cross-run path.
            if (data.TryGetValue(variableName, out var value) && value != null)
            {
                return value;
            }

            if (defaultValue != null)
            {
                return defaultValue;
            }

            return match.Value;
        });
    }
}
