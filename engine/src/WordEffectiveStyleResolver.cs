using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxWeb;

internal sealed class WordEffectiveStyleResolver
{
    private readonly Dictionary<string, Style> _stylesById;
    private readonly EffectiveParagraphState _defaultParagraph;
    private readonly EffectiveRunState _defaultRun;
    private readonly Dictionary<string, ResolvedStyleState> _styleCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string?> _themeFonts;

    public WordEffectiveStyleResolver(MainDocumentPart mainPart)
    {
        var stylesRoot = mainPart.StyleDefinitionsPart?.Styles;
        _stylesById = stylesRoot?
            .Elements<Style>()
            .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, Style>(StringComparer.OrdinalIgnoreCase);

        var docDefaults = stylesRoot?.Elements().FirstOrDefault(x => x.LocalName == "docDefaults");
        var defaultParagraphProperties = docDefaults?
            .Elements()
            .FirstOrDefault(x => x.LocalName == "pPrDefault")?
            .Elements()
            .FirstOrDefault(x => x.LocalName == "pPr");
        var defaultRunProperties = docDefaults?
            .Elements()
            .FirstOrDefault(x => x.LocalName == "rPrDefault")?
            .Elements()
            .FirstOrDefault(x => x.LocalName == "rPr");

        _defaultParagraph = ExtractParagraphState(defaultParagraphProperties);
        _defaultRun = ExtractRunState(defaultRunProperties);
        _themeFonts = LoadThemeFonts(mainPart.ThemePart);
    }

    public ParagraphFormatProfile ResolveParagraphFormat(Paragraph paragraph, bool hasPageBreak)
    {
        var paragraphState = ResolveParagraphState(paragraph);
        if (hasPageBreak)
            paragraphState.PageBreakBefore = true;

        return ToParagraphFormatProfile(paragraphState);
    }

    public List<InlineRunSummary> ResolveInlineRuns(Paragraph paragraph, int paragraphIndex)
    {
        var result = new List<InlineRunSummary>();
        var runIndex = 0;

        foreach (var run in paragraph.Elements<Run>())
        {
            var text = NormalizeRunText(run);
            if (string.IsNullOrWhiteSpace(text))
                continue;

            runIndex++;
            var format = ResolveRunFormat(paragraph, run);
            var segments = TokenizeText(text)
                .Select((token, index) => new InlineTextSegment
                {
                    Path = $"/body/paragraph[{paragraphIndex}]/run[{runIndex}]/segment[{index + 1}]",
                    Text = token,
                    Format = CloneRunFormat(format)
                })
                .ToList();
            result.Add(new InlineRunSummary
            {
                Path = $"/body/paragraph[{paragraphIndex}]/run[{runIndex}]",
                Text = text,
                Format = CloneRunFormat(format),
                Segments = segments
            });
        }

        return result;
    }

    public TemplateStyleProfile? ResolveStyleProfile(Style? style)
    {
        if (style == null)
            return null;

        var styleId = style.StyleId?.Value;
        var resolved = ResolveStyleState(styleId);
        if (string.IsNullOrWhiteSpace(styleId))
        {
            MergeParagraphState(resolved.Paragraph, ExtractParagraphState(style.StyleParagraphProperties));
            MergeRunState(resolved.Run, ExtractRunState(style.StyleRunProperties));
        }

        return new TemplateStyleProfile
        {
            StyleName = styleId,
            Paragraph = ToParagraphFormatProfile(resolved.Paragraph),
            Run = ToRunFormatProfile(resolved.Run)
        };
    }

    private ResolvedStyleState ResolveStyleState(string? styleId)
    {
        return ResolveStyleState(styleId, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
    }

    private ResolvedStyleState ResolveStyleState(string? styleId, HashSet<string> visiting)
    {
        if (string.IsNullOrWhiteSpace(styleId))
            return CreateDefaultStateWithDefaults();

        if (_styleCache.TryGetValue(styleId, out var cached))
            return CloneState(cached);

        if (!_stylesById.TryGetValue(styleId, out var style))
            return CreateDefaultStateWithDefaults();

        if (!visiting.Add(styleId))
            return CreateDefaultStateWithDefaults();

        var baseStyleId = GetAttributeValue(style.Elements().FirstOrDefault(x => x.LocalName == "basedOn"), "val");
        var state = ResolveStyleState(baseStyleId, visiting);
        MergeParagraphState(state.Paragraph, ExtractParagraphState(style.StyleParagraphProperties));
        MergeRunState(state.Run, ExtractRunState(style.StyleRunProperties));

        visiting.Remove(styleId);
        _styleCache[styleId] = CloneState(state);
        return state;
    }

    private EffectiveParagraphState ResolveParagraphState(Paragraph paragraph)
    {
        return ResolveParagraphState(paragraph.ParagraphProperties);
    }

    private EffectiveParagraphState ResolveParagraphState(ParagraphProperties? properties)
    {
        var styleId = properties?.ParagraphStyleId?.Val?.Value;
        var state = ResolveStyleState(styleId).Paragraph;
        MergeParagraphState(state, ExtractParagraphState(properties));
        return state;
    }

    private RunFormatProfile ResolveRunFormat(Paragraph paragraph, Run run)
    {
        var paragraphStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var paragraphStyle = ResolveStyleState(paragraphStyleId);
        var runState = CloneRunState(paragraphStyle.Run);

        var runProperties = run.RunProperties;
        var runStyleId = GetAttributeValue(runProperties?.Elements().FirstOrDefault(x => x.LocalName == "rStyle"), "val");
        if (!string.IsNullOrWhiteSpace(runStyleId))
        {
            var runStyle = ResolveStyleState(runStyleId);
            MergeRunState(runState, runStyle.Run);
        }

        MergeRunState(runState, ExtractRunState(runProperties));
        return ToRunFormatProfile(runState);
    }

    private static EffectiveParagraphState ExtractParagraphState(OpenXmlElement? properties)
    {
        var justification = properties?.Elements().FirstOrDefault(x => x.LocalName == "jc");
        var indentation = properties?.Elements().FirstOrDefault(x => x.LocalName == "ind");
        var spacing = properties?.Elements().FirstOrDefault(x => x.LocalName == "spacing");
        var pageBreakBefore = properties?.Elements().FirstOrDefault(x => x.LocalName == "pageBreakBefore");

        return new EffectiveParagraphState
        {
            Align = GetAttributeValue(justification, "val"),
            PageBreakBefore = ParseOnOffValue(pageBreakBefore),
            LeftIndent = GetAttributeValue(indentation, "left"),
            RightIndent = GetAttributeValue(indentation, "right"),
            FirstLineIndent = GetAttributeValue(indentation, "firstLine"),
            FirstLineChars = GetAttributeValue(indentation, "firstLineChars"),
            HangingIndent = GetAttributeValue(indentation, "hanging"),
            HangingChars = GetAttributeValue(indentation, "hangingChars"),
            BeforeSpacing = GetAttributeValue(spacing, "before"),
            AfterSpacing = GetAttributeValue(spacing, "after"),
            LineSpacing = GetAttributeValue(spacing, "line"),
            LineSpacingRule = GetAttributeValue(spacing, "lineRule")
        };
    }

    private static EffectiveRunState ExtractRunState(OpenXmlElement? properties)
    {
        var runFonts = properties?.Elements().FirstOrDefault(x => x.LocalName == "rFonts");
        var fontSize = properties?.Elements().FirstOrDefault(x => x.LocalName == "sz");
        var fontSizeCs = properties?.Elements().FirstOrDefault(x => x.LocalName == "szCs");
        var bold = properties?.Elements().FirstOrDefault(x => x.LocalName == "b");
        var italic = properties?.Elements().FirstOrDefault(x => x.LocalName == "i");
        var underline = properties?.Elements().FirstOrDefault(x => x.LocalName == "u");
        var color = properties?.Elements().FirstOrDefault(x => x.LocalName == "color");
        var highlight = properties?.Elements().FirstOrDefault(x => x.LocalName == "highlight");
        var verticalAlign = properties?.Elements().FirstOrDefault(x => x.LocalName == "vertAlign");

        return new EffectiveRunState
        {
            FontSize = GetAttributeValue(fontSize, "val"),
            FontSizeComplexScript = GetAttributeValue(fontSizeCs, "val"),
            AsciiFont = GetAttributeValue(runFonts, "ascii"),
            HighAnsiFont = GetAttributeValue(runFonts, "hAnsi"),
            EastAsiaFont = GetAttributeValue(runFonts, "eastAsia"),
            ComplexScriptFont = GetAttributeValue(runFonts, "cs"),
            AsciiThemeFont = GetAttributeValue(runFonts, "asciiTheme"),
            HighAnsiThemeFont = GetAttributeValue(runFonts, "hAnsiTheme"),
            EastAsiaThemeFont = GetAttributeValue(runFonts, "eastAsiaTheme"),
            ComplexScriptThemeFont = GetAttributeValue(runFonts, "cstheme", "csTheme"),
            Bold = ParseOnOffValue(bold),
            Italic = ParseOnOffValue(italic),
            Underline = underline == null ? null : GetAttributeValue(underline, "val") ?? "single",
            FontColor = GetAttributeValue(color, "val"),
            Highlight = GetAttributeValue(highlight, "val"),
            VerticalAlign = GetAttributeValue(verticalAlign, "val")
        };
    }

    private RunFormatProfile ToRunFormatProfile(EffectiveRunState state)
    {
        return new RunFormatProfile
        {
            FontSize = FirstNonEmpty(state.FontSize, state.FontSizeComplexScript),
            AsciiFont = FirstNonEmpty(
                state.AsciiFont,
                state.HighAnsiFont,
                ResolveThemeFont(state.AsciiThemeFont),
                ResolveThemeFont(state.HighAnsiThemeFont),
                state.ComplexScriptFont,
                ResolveThemeFont(state.ComplexScriptThemeFont)),
            EastAsiaFont = FirstNonEmpty(
                state.EastAsiaFont,
                ResolveThemeFont(state.EastAsiaThemeFont)),
            Bold = state.Bold ?? false,
            Italic = state.Italic ?? false,
            Underline = state.Underline,
            FontColor = state.FontColor,
            Highlight = state.Highlight,
            VerticalAlign = state.VerticalAlign
        };
    }

    private static ParagraphFormatProfile ToParagraphFormatProfile(EffectiveParagraphState state)
    {
        return new ParagraphFormatProfile
        {
            Align = state.Align,
            PageBreakBefore = state.PageBreakBefore ?? false,
            LeftIndent = state.LeftIndent,
            RightIndent = state.RightIndent,
            FirstLineIndent = state.FirstLineIndent,
            FirstLineChars = state.FirstLineChars,
            HangingIndent = state.HangingIndent,
            HangingChars = state.HangingChars,
            BeforeSpacing = state.BeforeSpacing,
            AfterSpacing = state.AfterSpacing,
            LineSpacing = state.LineSpacing,
            LineSpacingRule = state.LineSpacingRule
        };
    }

    private static void MergeParagraphState(EffectiveParagraphState target, EffectiveParagraphState source)
    {
        target.Align = source.Align ?? target.Align;
        target.PageBreakBefore = source.PageBreakBefore ?? target.PageBreakBefore;
        target.LeftIndent = source.LeftIndent ?? target.LeftIndent;
        target.RightIndent = source.RightIndent ?? target.RightIndent;
        target.FirstLineIndent = source.FirstLineIndent ?? target.FirstLineIndent;
        target.FirstLineChars = source.FirstLineChars ?? target.FirstLineChars;
        target.HangingIndent = source.HangingIndent ?? target.HangingIndent;
        target.HangingChars = source.HangingChars ?? target.HangingChars;
        target.BeforeSpacing = source.BeforeSpacing ?? target.BeforeSpacing;
        target.AfterSpacing = source.AfterSpacing ?? target.AfterSpacing;
        target.LineSpacing = source.LineSpacing ?? target.LineSpacing;
        target.LineSpacingRule = source.LineSpacingRule ?? target.LineSpacingRule;
    }

    private static void MergeRunState(EffectiveRunState target, EffectiveRunState source)
    {
        target.FontSize = source.FontSize ?? target.FontSize;
        target.FontSizeComplexScript = source.FontSizeComplexScript ?? target.FontSizeComplexScript;
        target.AsciiFont = source.AsciiFont ?? target.AsciiFont;
        target.HighAnsiFont = source.HighAnsiFont ?? target.HighAnsiFont;
        target.EastAsiaFont = source.EastAsiaFont ?? target.EastAsiaFont;
        target.ComplexScriptFont = source.ComplexScriptFont ?? target.ComplexScriptFont;
        target.AsciiThemeFont = source.AsciiThemeFont ?? target.AsciiThemeFont;
        target.HighAnsiThemeFont = source.HighAnsiThemeFont ?? target.HighAnsiThemeFont;
        target.EastAsiaThemeFont = source.EastAsiaThemeFont ?? target.EastAsiaThemeFont;
        target.ComplexScriptThemeFont = source.ComplexScriptThemeFont ?? target.ComplexScriptThemeFont;
        target.Bold = source.Bold ?? target.Bold;
        target.Italic = source.Italic ?? target.Italic;
        target.Underline = source.Underline ?? target.Underline;
        target.FontColor = source.FontColor ?? target.FontColor;
        target.Highlight = source.Highlight ?? target.Highlight;
        target.VerticalAlign = source.VerticalAlign ?? target.VerticalAlign;
    }

    private ResolvedStyleState CreateDefaultStateWithDefaults()
    {
        return new ResolvedStyleState
        {
            Paragraph = CloneParagraphState(_defaultParagraph),
            Run = CloneRunState(_defaultRun)
        };
    }

    private static ResolvedStyleState CloneState(ResolvedStyleState source)
    {
        return new ResolvedStyleState
        {
            Paragraph = CloneParagraphState(source.Paragraph),
            Run = CloneRunState(source.Run)
        };
    }

    private static EffectiveParagraphState CloneParagraphState(EffectiveParagraphState source)
    {
        return new EffectiveParagraphState
        {
            Align = source.Align,
            PageBreakBefore = source.PageBreakBefore,
            LeftIndent = source.LeftIndent,
            RightIndent = source.RightIndent,
            FirstLineIndent = source.FirstLineIndent,
            FirstLineChars = source.FirstLineChars,
            HangingIndent = source.HangingIndent,
            HangingChars = source.HangingChars,
            BeforeSpacing = source.BeforeSpacing,
            AfterSpacing = source.AfterSpacing,
            LineSpacing = source.LineSpacing,
            LineSpacingRule = source.LineSpacingRule
        };
    }

    private static EffectiveRunState CloneRunState(EffectiveRunState source)
    {
        return new EffectiveRunState
        {
            FontSize = source.FontSize,
            FontSizeComplexScript = source.FontSizeComplexScript,
            AsciiFont = source.AsciiFont,
            HighAnsiFont = source.HighAnsiFont,
            EastAsiaFont = source.EastAsiaFont,
            ComplexScriptFont = source.ComplexScriptFont,
            AsciiThemeFont = source.AsciiThemeFont,
            HighAnsiThemeFont = source.HighAnsiThemeFont,
            EastAsiaThemeFont = source.EastAsiaThemeFont,
            ComplexScriptThemeFont = source.ComplexScriptThemeFont,
            Bold = source.Bold,
            Italic = source.Italic,
            Underline = source.Underline,
            FontColor = source.FontColor,
            Highlight = source.Highlight,
            VerticalAlign = source.VerticalAlign
        };
    }

    private string? ResolveThemeFont(string? themeKey)
    {
        if (string.IsNullOrWhiteSpace(themeKey))
            return null;

        return _themeFonts.TryGetValue(themeKey, out var value) ? value : null;
    }

    private static Dictionary<string, string?> LoadThemeFonts(ThemePart? themePart)
    {
        var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        if (themePart == null)
            return result;

        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        var document = XDocument.Load(stream);
        var fontScheme = document.Descendants().FirstOrDefault(x => x.Name.LocalName == "fontScheme");
        if (fontScheme == null)
            return result;

        AddThemeFontSet(result, fontScheme, "major", "major");
        AddThemeFontSet(result, fontScheme, "minor", "minor");
        return result;
    }

    private static void AddThemeFontSet(Dictionary<string, string?> target, XElement fontScheme, string elementName, string prefix)
    {
        var fontSet = fontScheme.Elements().FirstOrDefault(x => x.Name.LocalName == elementName + "Font");
        if (fontSet == null)
            return;

        var latin = GetThemeTypeface(fontSet, "latin");
        var eastAsia = GetThemeTypeface(fontSet, "ea");
        var complex = GetThemeTypeface(fontSet, "cs");

        target[prefix + "Ascii"] = latin;
        target[prefix + "HAnsi"] = latin;
        target[prefix + "EastAsia"] = eastAsia;
        target[prefix + "Bidi"] = complex;
        target[prefix + "Cs"] = complex;
    }

    private static string? GetThemeTypeface(XElement fontSet, string name)
    {
        return fontSet.Elements()
            .FirstOrDefault(x => x.Name.LocalName == name)?
            .Attribute("typeface")?
            .Value;
    }

    private static string? NormalizeRunText(Run run)
    {
        var value = string.Concat(run.Elements<Text>().Select(x => x.Text));
        if (string.IsNullOrEmpty(value))
            return null;

        value = value.Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Trim();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static List<string> TokenizeText(string value)
    {
        var tokens = new List<string>();
        var current = new System.Text.StringBuilder();
        var currentKind = TokenKind.None;

        foreach (var ch in value)
        {
            var kind = GetTokenKind(ch);
            if (kind == TokenKind.Whitespace)
            {
                Flush();
                tokens.Add(ch.ToString());
                continue;
            }

            if (currentKind != kind || kind == TokenKind.Other)
            {
                Flush();
                if (kind == TokenKind.Other)
                {
                    tokens.Add(ch.ToString());
                    continue;
                }
            }

            current.Append(ch);
            currentKind = kind;
        }

        Flush();
        return tokens.Where(token => token.Length > 0).ToList();

        void Flush()
        {
            if (current.Length == 0)
                return;

            tokens.Add(current.ToString());
            current.Clear();
            currentKind = TokenKind.None;
        }
    }

    private static TokenKind GetTokenKind(char ch)
    {
        if (char.IsWhiteSpace(ch))
            return TokenKind.Whitespace;
        if (char.IsLetter(ch))
            return TokenKind.Letter;
        if (char.IsDigit(ch))
            return TokenKind.Digit;
        return TokenKind.Other;
    }

    private static bool? ParseOnOffValue(OpenXmlElement? element)
    {
        if (element == null)
            return null;

        var value = GetAttributeValue(element, "val");
        if (string.IsNullOrWhiteSpace(value))
            return true;

        return value.Equals("0", StringComparison.OrdinalIgnoreCase)
            || value.Equals("false", StringComparison.OrdinalIgnoreCase)
            || value.Equals("off", StringComparison.OrdinalIgnoreCase)
            ? false
            : true;
    }

    private static string? GetAttributeValue(OpenXmlElement? element, params string[] names)
    {
        if (element == null)
            return null;

        foreach (var name in names)
        {
            var match = element.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(match.Value))
                return match.Value;
        }

        return null;
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    private static RunFormatProfile CloneRunFormat(RunFormatProfile? format)
    {
        format ??= new RunFormatProfile();
        return new RunFormatProfile
        {
            FontSize = format.FontSize,
            AsciiFont = format.AsciiFont,
            EastAsiaFont = format.EastAsiaFont,
            Bold = format.Bold,
            Italic = format.Italic,
            Underline = format.Underline,
            FontColor = format.FontColor,
            Highlight = format.Highlight,
            VerticalAlign = format.VerticalAlign
        };
    }

    private sealed class ResolvedStyleState
    {
        public EffectiveParagraphState Paragraph { get; set; } = new();
        public EffectiveRunState Run { get; set; } = new();
    }

    private sealed class EffectiveParagraphState
    {
        public string? Align { get; set; }
        public bool? PageBreakBefore { get; set; }
        public string? LeftIndent { get; set; }
        public string? RightIndent { get; set; }
        public string? FirstLineIndent { get; set; }
        public string? FirstLineChars { get; set; }
        public string? HangingIndent { get; set; }
        public string? HangingChars { get; set; }
        public string? BeforeSpacing { get; set; }
        public string? AfterSpacing { get; set; }
        public string? LineSpacing { get; set; }
        public string? LineSpacingRule { get; set; }
    }

    private sealed class EffectiveRunState
    {
        public string? FontSize { get; set; }
        public string? FontSizeComplexScript { get; set; }
        public string? AsciiFont { get; set; }
        public string? HighAnsiFont { get; set; }
        public string? EastAsiaFont { get; set; }
        public string? ComplexScriptFont { get; set; }
        public string? AsciiThemeFont { get; set; }
        public string? HighAnsiThemeFont { get; set; }
        public string? EastAsiaThemeFont { get; set; }
        public string? ComplexScriptThemeFont { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public string? Underline { get; set; }
        public string? FontColor { get; set; }
        public string? Highlight { get; set; }
        public string? VerticalAlign { get; set; }
    }

    private enum TokenKind
    {
        None,
        Whitespace,
        Letter,
        Digit,
        Other
    }
}
