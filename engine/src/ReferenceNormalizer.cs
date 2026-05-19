using System.Text;
using System.Text.RegularExpressions;

namespace DocxWeb;

public static class ReferenceNormalizer
{
    private const string BookmarkPrefix = "_Ref_ref_";
    private const string DefaultHangingChars = "200";
    private const string DefaultTabPosition = "480";

    private static readonly Regex ReferenceEntryPattern =
        new(@"^\s*\[(\d+)\](\s*)(.*)$", RegexOptions.Singleline);

    private static readonly Regex CitationPattern =
        new(@"\[(\d+(?:\s*[,\-]\s*\d+)*)\]");

    private static readonly Regex CitationTokenPattern =
        new(@"\d+|\s*,\s*|\s*-\s*");

    public static void Normalize(RenderSpec? spec)
    {
        if (spec == null) return;

        var refNumbers = new HashSet<int>();
        var nextId = 1000;

        foreach (var block in spec.Body)
        {
            if (!IsReferenceBlock(block)) continue;
            if (IsAlreadyNormalizedReference(block)) continue;
            NormalizeReferenceBlock(block, refNumbers, ref nextId);
        }

        if (refNumbers.Count == 0)
            return;

        foreach (var block in spec.Body)
        {
            if (IsReferenceBlock(block)) continue;
            NormalizeBodyCitations(block, refNumbers);
        }
    }

    private static bool IsReferenceBlock(RenderBodyBlock block)
    {
        if (string.Equals(block.Kind, "reference", StringComparison.OrdinalIgnoreCase))
            return true;

        var paragraph = block.Paragraph;
        if (paragraph == null) return false;

        if (string.Equals(paragraph.Style, "Reference", StringComparison.OrdinalIgnoreCase))
            return true;

        var text = ExtractParagraphText(paragraph);
        return !string.IsNullOrEmpty(text)
            && Regex.IsMatch(text, @"^\s*\[\d+\]");
    }

    private static bool IsAlreadyNormalizedReference(RenderBodyBlock block)
    {
        var runs = block.Paragraph?.Runs;
        if (runs == null) return false;
        return runs.Any(r => string.Equals(r.Kind, "bookmarkStart", StringComparison.OrdinalIgnoreCase));
    }

    private static string ExtractParagraphText(RenderParagraphSpec paragraph)
    {
        if (!string.IsNullOrEmpty(paragraph.Text)) return paragraph.Text;
        if (paragraph.Runs.Count == 0) return string.Empty;

        var sb = new StringBuilder();
        foreach (var run in paragraph.Runs)
        {
            if (!string.IsNullOrEmpty(run.Kind)) continue;
            if (run.Segments.Count > 0)
            {
                foreach (var seg in run.Segments) sb.Append(seg.Text ?? string.Empty);
            }
            else
            {
                sb.Append(run.Text ?? string.Empty);
            }
        }
        return sb.ToString();
    }

    private static void NormalizeReferenceBlock(RenderBodyBlock block, HashSet<int> refNumbers, ref int nextId)
    {
        var paragraph = block.Paragraph;
        if (paragraph == null) return;

        var text = ExtractParagraphText(paragraph);
        var match = ReferenceEntryPattern.Match(text);
        if (!match.Success) return;

        if (!int.TryParse(match.Groups[1].Value, out var refNumber)) return;
        var content = match.Groups[3].Value;
        refNumbers.Add(refNumber);

        var baseFormat = paragraph.Run ?? new RunFormatProfile();
        if (paragraph.Runs.Count > 0)
        {
            var firstWithText = paragraph.Runs
                .FirstOrDefault(r => string.IsNullOrEmpty(r.Kind)
                    && (!string.IsNullOrEmpty(r.Text) || r.Segments.Count > 0));
            if (firstWithText?.Format != null && HasAnyFormat(firstWithText.Format))
                baseFormat = firstWithText.Format;
        }

        var bookmarkId = (nextId++).ToString();
        var bookmarkName = BookmarkPrefix + refNumber;

        var newRuns = new List<InlineRunSummary>
        {
            new() { Kind = "bookmarkStart", BookmarkId = bookmarkId, BookmarkName = bookmarkName },
            new() { Text = $"[{refNumber}]", Format = CloneFormat(baseFormat) },
            new() { Kind = "bookmarkEnd", BookmarkId = bookmarkId },
            new() { Kind = "tab", Format = CloneFormat(baseFormat) }
        };

        if (!string.IsNullOrEmpty(content))
            newRuns.Add(new InlineRunSummary { Text = content, Format = CloneFormat(baseFormat) });

        paragraph.Text = null;
        paragraph.Runs = newRuns;

        var pFormat = paragraph.Paragraph ??= new ParagraphFormatProfile();
        if (string.IsNullOrWhiteSpace(pFormat.HangingChars)
            && string.IsNullOrWhiteSpace(pFormat.HangingIndent))
        {
            pFormat.HangingChars = DefaultHangingChars;
        }

        if (pFormat.TabStops.Count == 0)
        {
            pFormat.TabStops.Add(new TabStopSpec
            {
                Position = DefaultTabPosition,
                Alignment = "left"
            });
        }
    }

    private static void NormalizeBodyCitations(RenderBodyBlock block, HashSet<int> refNumbers)
    {
        var paragraph = block.Paragraph;
        if (paragraph == null || paragraph.Runs.Count == 0) return;

        var newRuns = new List<InlineRunSummary>();
        var changed = false;

        foreach (var run in paragraph.Runs)
        {
            if (!IsCandidateCitationRun(run))
            {
                newRuns.Add(run);
                continue;
            }

            var text = run.Text ?? string.Empty;
            if (!CitationPattern.IsMatch(text))
            {
                newRuns.Add(run);
                continue;
            }

            var split = SplitCitationRun(run, refNumbers);
            if (split == null)
            {
                newRuns.Add(run);
                continue;
            }

            newRuns.AddRange(split);
            changed = true;
        }

        if (changed)
            paragraph.Runs = newRuns;
    }

    private static bool IsCandidateCitationRun(InlineRunSummary run)
    {
        if (!string.IsNullOrEmpty(run.Kind)) return false;
        if (run.Segments.Count > 0) return false;
        return string.Equals(run.Format?.VerticalAlign, "superscript", StringComparison.OrdinalIgnoreCase);
    }

    private static List<InlineRunSummary>? SplitCitationRun(InlineRunSummary run, HashSet<int> refNumbers)
    {
        var text = run.Text ?? string.Empty;
        if (string.IsNullOrEmpty(text)) return null;

        var format = run.Format ?? new RunFormatProfile();
        var result = new List<InlineRunSummary>();
        var lastIndex = 0;
        var anyRef = false;

        foreach (Match m in CitationPattern.Matches(text))
        {
            if (m.Index > lastIndex)
            {
                result.Add(new InlineRunSummary
                {
                    Text = text.Substring(lastIndex, m.Index - lastIndex),
                    Format = CloneFormat(format)
                });
            }

            result.Add(new InlineRunSummary { Text = "[", Format = CloneFormat(format) });

            var inner = m.Groups[1].Value;
            foreach (Match tok in CitationTokenPattern.Matches(inner))
            {
                var tokText = tok.Value;
                if (int.TryParse(tokText.Trim(), out var n) && refNumbers.Contains(n))
                {
                    result.Add(new InlineRunSummary
                    {
                        Kind = "refField",
                        FieldInstruction = $" REF {BookmarkPrefix}{n} \\h ",
                        FieldDisplayText = n.ToString(),
                        Format = CloneFormat(format)
                    });
                    anyRef = true;
                }
                else
                {
                    result.Add(new InlineRunSummary
                    {
                        Text = tokText,
                        Format = CloneFormat(format)
                    });
                }
            }

            result.Add(new InlineRunSummary { Text = "]", Format = CloneFormat(format) });
            lastIndex = m.Index + m.Length;
        }

        if (lastIndex < text.Length)
        {
            result.Add(new InlineRunSummary
            {
                Text = text.Substring(lastIndex),
                Format = CloneFormat(format)
            });
        }

        return anyRef ? result : null;
    }

    private static RunFormatProfile CloneFormat(RunFormatProfile? source)
    {
        source ??= new RunFormatProfile();
        return new RunFormatProfile
        {
            FontSize = source.FontSize,
            AsciiFont = source.AsciiFont,
            EastAsiaFont = source.EastAsiaFont,
            Bold = source.Bold,
            Italic = source.Italic,
            Underline = source.Underline,
            FontColor = source.FontColor,
            Highlight = source.Highlight,
            VerticalAlign = source.VerticalAlign
        };
    }

    private static bool HasAnyFormat(RunFormatProfile format)
    {
        return !string.IsNullOrWhiteSpace(format.FontSize)
            || !string.IsNullOrWhiteSpace(format.AsciiFont)
            || !string.IsNullOrWhiteSpace(format.EastAsiaFont)
            || format.Bold
            || format.Italic
            || !string.IsNullOrWhiteSpace(format.Underline)
            || !string.IsNullOrWhiteSpace(format.FontColor)
            || !string.IsNullOrWhiteSpace(format.Highlight)
            || !string.IsNullOrWhiteSpace(format.VerticalAlign);
    }
}
