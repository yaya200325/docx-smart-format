using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxWeb;

internal static class DocumentAnalysisFactory
{
    public static TemplateAnalysis BuildTemplateAnalysis(string templateDocId, AnalysisReport report, string? templatePath = null)
    {
        var styles = InferTemplateStyles(report, templatePath);
        return new TemplateAnalysis
        {
            TemplateDocId = templateDocId,
            Document = report.Document,
            Headers = report.Headers,
            Footers = report.Footers,
            Sections = report.Sections,
            Styles = styles,
            Rules = BuildTemplateRules(report, styles)
        };
    }

    public static TemplateOverview BuildTemplateOverview(TemplateAnalysis template)
    {
        var styles = new List<TemplateStyleOverview>();
        var heading1 = BuildStyleOverview("template.heading1", template.Styles.Heading1);
        var heading2 = BuildStyleOverview("template.heading2", template.Styles.Heading2);
        var heading3 = BuildStyleOverview("template.heading3", template.Styles.Heading3);
        var body = BuildStyleOverview("template.body", template.Styles.Body);
        var reference = BuildStyleOverview("template.reference", template.Styles.Reference);
        var figureCaption = BuildStyleOverview("template.figureCaption", template.Styles.FigureCaption);
        var tableCaption = BuildStyleOverview("template.tableCaption", template.Styles.TableCaption);
        var equation = BuildStyleOverview("template.equation", template.Styles.Equation);

        if (heading1 != null) styles.Add(heading1);
        if (heading2 != null) styles.Add(heading2);
        if (heading3 != null) styles.Add(heading3);
        if (body != null) styles.Add(body);
        if (reference != null) styles.Add(reference);
        if (figureCaption != null) styles.Add(figureCaption);
        if (tableCaption != null) styles.Add(tableCaption);
        if (equation != null) styles.Add(equation);

        var overview = new TemplateOverview
        {
            Styles = styles,
            Heading1 = heading1,
            Heading2 = heading2,
            Heading3 = heading3,
            Body = body,
            Reference = reference,
            FigureCaption = figureCaption,
            TableCaption = tableCaption,
            Equation = equation,
            Sections = template.Sections.Select((section, index) => new TemplateSectionOverview
            {
                Index = index + 1,
                Path = section.Path,
                Type = section.Type,
                TitlePage = section.TitlePage,
                PageStart = section.PageStart,
                PageNumFmt = section.PageNumFmt,
                HeaderType = FirstReferenceType(section.HeaderSourceRefs),
                FooterType = FirstReferenceType(section.FooterSourceRefs),
                HeaderText = FirstHeaderFooterText(template.Headers, section.HeaderSourceRefs),
                EvenHeaderText = HeaderFooterTextByType(template.Headers, section.HeaderSourceRefs, "even"),
                FooterText = FirstHeaderFooterText(template.Footers, section.FooterSourceRefs),
                HeaderSourcePath = FirstReferencePath(section.HeaderSourceRefs),
                FooterSourcePath = FirstReferencePath(section.FooterSourceRefs),
                HasEvenHeader = HasReferenceType(section.HeaderSourceRefs, "even"),
                HasEvenFooter = HasReferenceType(section.FooterSourceRefs, "even")
            }).ToList(),
            HeaderFooter = new TemplateHeaderFooterOverview
            {
                HasHeader = template.Headers.Count > 0,
                HasFooter = template.Footers.Count > 0,
                HasOddEvenHeader = template.Headers.Any(x => x.Type.Equals("even", StringComparison.OrdinalIgnoreCase) || x.Type.Equals("odd", StringComparison.OrdinalIgnoreCase)),
                HasOddEvenFooter = template.Footers.Any(x => x.Type.Equals("even", StringComparison.OrdinalIgnoreCase) || x.Type.Equals("odd", StringComparison.OrdinalIgnoreCase)),
                HasPageNumber = template.Rules.HeaderFooterRules.ContainsPageNumber || template.Rules.HeaderFooterRules.ContainsTotalPages,
                PageNumberScheme = template.Rules.DocumentRules.PageNumberFormats.FirstOrDefault(),
                HasTitlePage = template.Sections.Any(x => x.TitlePage)
            }
        };

        return overview;
    }

    public static SourceDocumentAnalysis BuildSourceAnalysis(string sourceDocId, AnalysisReport report)
    {
        var blocks = new List<SemanticBlock>();
        var index = 0;
        foreach (var block in report.Body)
        {
            index++;
            blocks.Add(new SemanticBlock
            {
                BlockId = $"b{index:D4}",
                Kind = block.Kind,
                Text = block.Text,
                Path = block.Path,
                SectionIndex = block.SectionIndex,
                Style = block.Style,
                HeadingLevel = block.HeadingLevel,
                Paragraph = CloneParagraphFormat(block.Paragraph),
                Run = CloneRunFormat(block.Run),
                Runs = block.Runs.Select(CloneInlineRun).ToList(),
                Image = block.Image,
                Table = block.HasTable ? new RenderTableSpec { Rows = block.TableRows } : null,
                Equation = block.Equation == null ? null : new EquationSpec
                {
                    Text = block.Equation.Text,
                    Xml = block.Equation.Xml,
                    DisplayMode = block.Equation.DisplayMode
                },
                Hints = BuildHints(block)
            });
        }

        return new SourceDocumentAnalysis
        {
            SourceDocId = sourceDocId,
            Document = report.Document,
            Blocks = blocks,
            Sections = report.Sections,
            Headers = report.Headers,
            Footers = report.Footers,
            Assets = report.Assets
        };
    }

    private static TemplateStyleCatalog InferTemplateStyles(AnalysisReport report, string? templatePath)
    {
        var styleCatalog = string.IsNullOrWhiteSpace(templatePath)
            ? new TemplateStyleCatalog()
            : LoadTemplateStylesFromDefinitions(templatePath);

        styleCatalog.Heading1 = MergeMissingStyleFields(styleCatalog.Heading1, BuildProfile(report.Body.FirstOrDefault(x => x.HeadingLevel == 1)));
        styleCatalog.Heading2 = MergeMissingStyleFields(styleCatalog.Heading2, BuildProfile(report.Body.FirstOrDefault(x => x.HeadingLevel == 2)));
        styleCatalog.Heading3 = MergeMissingStyleFields(styleCatalog.Heading3, BuildProfile(report.Body.FirstOrDefault(x => x.HeadingLevel == 3)));
        styleCatalog.Body = MergeMissingStyleFields(styleCatalog.Body, BuildProfile(report.Body.FirstOrDefault(x => x.Kind == "paragraph" && !x.HeadingLevel.HasValue)));
        styleCatalog.Reference = MergeMissingStyleFields(styleCatalog.Reference, BuildProfile(report.Body.FirstOrDefault(x => LooksLikeReferenceEntry(x.Text))));
        styleCatalog.FigureCaption = MergeMissingStyleFields(styleCatalog.FigureCaption, BuildProfile(report.Body.FirstOrDefault(x => LooksLikeFigureCaption(x.Text))));
        styleCatalog.TableCaption = MergeMissingStyleFields(styleCatalog.TableCaption, BuildProfile(report.Body.FirstOrDefault(x => LooksLikeTableCaption(x.Text))));
        styleCatalog.Equation = MergeMissingStyleFields(styleCatalog.Equation, BuildProfile(report.Body.FirstOrDefault(x => LooksLikeEquation(x.Text))));
        return styleCatalog;
    }

    private static TemplateRuleSet BuildTemplateRules(AnalysisReport report, TemplateStyleCatalog styles)
    {
        return new TemplateRuleSet
        {
            DocumentRules = new TemplateDocumentRules
            {
                PageWidth = report.Document.PageWidth,
                PageHeight = report.Document.PageHeight,
                Orientation = report.Document.Orientation,
                MarginTop = report.Document.MarginTop,
                MarginBottom = report.Document.MarginBottom,
                MarginLeft = report.Document.MarginLeft,
                MarginRight = report.Document.MarginRight,
                SectionTypes = report.Sections
                    .Select(x => x.Type)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Cast<string>()
                    .ToList(),
                PageNumberFormats = report.Sections
                    .Select(x => x.PageNumFmt)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Cast<string>()
                    .ToList()
            },
            HeaderFooterRules = new HeaderFooterRuleSet
            {
                HasHeader = report.Headers.Count > 0,
                HasFooter = report.Footers.Count > 0,
                HeaderTypes = report.Headers
                    .Select(x => x.Type)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                FooterTypes = report.Footers
                    .Select(x => x.Type)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                ContainsPageNumber = report.Headers.Any(x => x.ContainsPageField) || report.Footers.Any(x => x.ContainsPageField),
                ContainsTotalPages = report.Headers.Any(x => x.ContainsNumPagesField) || report.Footers.Any(x => x.ContainsNumPagesField)
            },
            StyleProfiles = styles,
            CaptionRules = new CaptionRuleSet
            {
                FigurePrefixes = ["Figure"],
                TablePrefixes = ["Table"],
                FigureCaptionStyleKey = styles.FigureCaption == null ? null : "template.figureCaption",
                TableCaptionStyleKey = styles.TableCaption == null ? null : "template.tableCaption",
                FigureCaptionPosition = "below",
                TableCaptionPosition = "above"
            },
            OutlineRules = new OutlineRuleSet
            {
                HeadingStyleKeys = ["template.heading1", "template.heading2", "template.heading3"],
                HeadingRecognitionHints = ["styleName", "headingLevel", "outlineLevel", "numberedPrefix"],
                MaxHeadingLevel = report.Body
                    .Where(x => x.HeadingLevel.HasValue)
                    .Select(x => x.HeadingLevel!.Value)
                    .DefaultIfEmpty(0)
                    .Max()
            },
            NormalizedRequirements = []
        };
    }

    private static TemplateStyleCatalog LoadTemplateStylesFromDefinitions(string templatePath)
    {
        if (!File.Exists(templatePath))
            return new TemplateStyleCatalog();

        using var document = WordprocessingDocument.Open(templatePath, false);
        var mainPart = document.MainDocumentPart;
        if (mainPart == null)
            return new TemplateStyleCatalog();

        var resolver = new WordEffectiveStyleResolver(mainPart);
        var styles = mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>().ToList() ?? [];
        if (styles.Count == 0)
            return new TemplateStyleCatalog();

        return new TemplateStyleCatalog
        {
            Heading1 = FindHeadingStyle(styles, resolver, 1),
            Heading2 = FindHeadingStyle(styles, resolver, 2),
            Heading3 = FindHeadingStyle(styles, resolver, 3),
            Body = FindBodyStyle(styles, resolver),
            Reference = FindReferenceStyle(styles, resolver),
            FigureCaption = FindCaptionStyle(styles, resolver, "figure"),
            TableCaption = FindCaptionStyle(styles, resolver, "table"),
            Equation = FindEquationStyle(styles, resolver)
        };
    }

    private static TemplateStyleProfile? FindHeadingStyle(List<Style> styles, WordEffectiveStyleResolver resolver, int level)
    {
        var headingCnA = $"鏍囬 {level}";
        var headingCnB = $"鏍囬{level}";
        var match = styles.FirstOrDefault(style =>
                IsParagraphStyle(style)
                && (MatchesStyleId(style, $"Heading{level}")
                    || MatchesStyleName(style, $"heading {level}")
                    || MatchesStyleName(style, headingCnA)
                    || MatchesStyleName(style, headingCnB)
                    || MatchesOutlineLevel(style, level - 1)));

        return resolver.ResolveStyleProfile(match);
    }

    private static TemplateStyleProfile? FindBodyStyle(List<Style> styles, WordEffectiveStyleResolver resolver)
    {
        var match = styles.FirstOrDefault(style =>
                IsParagraphStyle(style)
                && (style.Default?.Value ?? false))
            ?? styles.FirstOrDefault(style => IsParagraphStyle(style) && MatchesStyleId(style, "Normal"))
            ?? styles.FirstOrDefault(style => IsParagraphStyle(style) && MatchesStyleName(style, "姝ｆ枃"));

        return resolver.ResolveStyleProfile(match);
    }

    private static TemplateStyleProfile? FindReferenceStyle(List<Style> styles, WordEffectiveStyleResolver resolver)
    {
        var match = styles.FirstOrDefault(style =>
                IsParagraphStyle(style)
                && (MatchesStyleId(style, "Reference")
                    || MatchesStyleId(style, "Bibliography")
                    || MatchesStyleName(style, "reference")
                    || MatchesStyleName(style, "bibliography")
                    || MatchesStyleName(style, "\u53C2\u8003\u6587\u732E")
                    || MatchesStyleName(style, "\u6587\u732E")))
            ?? styles.FirstOrDefault(style =>
                IsParagraphStyle(style)
                && HasReferenceParagraphTraits(resolver.ResolveStyleProfile(style)));

        return resolver.ResolveStyleProfile(match);
    }

    private static TemplateStyleProfile? FindCaptionStyle(List<Style> styles, WordEffectiveStyleResolver resolver, string kind)
    {
        var match = styles.FirstOrDefault(style =>
                IsParagraphStyle(style)
                && (MatchesStyleId(style, "Caption")
                    || MatchesStyleName(style, "caption")
                    || MatchesStyleName(style, "棰樻敞")))
            ?? styles.FirstOrDefault(style =>
                IsParagraphStyle(style)
                && (kind == "figure"
                    ? MatchesStyleName(style, "\u56FE")
                    : MatchesStyleName(style, "\u8868")));

        return resolver.ResolveStyleProfile(match);
    }

    private static TemplateStyleProfile? FindEquationStyle(List<Style> styles, WordEffectiveStyleResolver resolver)
    {
        var match = styles.FirstOrDefault(style =>
            IsParagraphStyle(style)
            && (MatchesStyleName(style, "equation") || MatchesStyleName(style, "\u516C\u5F0F")));

        return resolver.ResolveStyleProfile(match);
    }

    private static TemplateStyleProfile? BuildProfile(BodyBlockSummary? block)
    {
        if (block == null)
            return null;

        return new TemplateStyleProfile
        {
            StyleName = block.Style,
            Paragraph = CloneParagraphFormat(block.Paragraph),
            Run = CloneRunFormat(block.Run)
        };
    }

    private static TemplateStyleProfile? MergeMissingStyleFields(TemplateStyleProfile? templateStyle, TemplateStyleProfile? extractedStyle)
    {
        if (templateStyle == null)
            return extractedStyle;

        if (extractedStyle == null)
            return templateStyle;

        templateStyle.StyleName ??= extractedStyle.StyleName;

        templateStyle.Run.FontSize = FirstNonEmpty(templateStyle.Run.FontSize, extractedStyle.Run.FontSize);
        templateStyle.Run.AsciiFont = FirstNonEmpty(templateStyle.Run.AsciiFont, extractedStyle.Run.AsciiFont);
        templateStyle.Run.EastAsiaFont = FirstNonEmpty(templateStyle.Run.EastAsiaFont, extractedStyle.Run.EastAsiaFont);
        templateStyle.Run.Underline = FirstNonEmpty(templateStyle.Run.Underline, extractedStyle.Run.Underline);
        templateStyle.Run.FontColor = FirstNonEmpty(templateStyle.Run.FontColor, extractedStyle.Run.FontColor);
        templateStyle.Run.Highlight = FirstNonEmpty(templateStyle.Run.Highlight, extractedStyle.Run.Highlight);

        templateStyle.Paragraph.Align = FirstNonEmpty(templateStyle.Paragraph.Align, extractedStyle.Paragraph.Align);
        templateStyle.Paragraph.LeftIndent = FirstNonEmpty(templateStyle.Paragraph.LeftIndent, extractedStyle.Paragraph.LeftIndent);
        templateStyle.Paragraph.RightIndent = FirstNonEmpty(templateStyle.Paragraph.RightIndent, extractedStyle.Paragraph.RightIndent);
        templateStyle.Paragraph.FirstLineIndent = FirstNonEmpty(templateStyle.Paragraph.FirstLineIndent, extractedStyle.Paragraph.FirstLineIndent);
        templateStyle.Paragraph.FirstLineChars = FirstNonEmpty(templateStyle.Paragraph.FirstLineChars, extractedStyle.Paragraph.FirstLineChars);
        templateStyle.Paragraph.HangingIndent = FirstNonEmpty(templateStyle.Paragraph.HangingIndent, extractedStyle.Paragraph.HangingIndent);
        templateStyle.Paragraph.HangingChars = FirstNonEmpty(templateStyle.Paragraph.HangingChars, extractedStyle.Paragraph.HangingChars);
        templateStyle.Paragraph.BeforeSpacing = FirstNonEmpty(templateStyle.Paragraph.BeforeSpacing, extractedStyle.Paragraph.BeforeSpacing);
        templateStyle.Paragraph.AfterSpacing = FirstNonEmpty(templateStyle.Paragraph.AfterSpacing, extractedStyle.Paragraph.AfterSpacing);
        templateStyle.Paragraph.LineSpacing = FirstNonEmpty(templateStyle.Paragraph.LineSpacing, extractedStyle.Paragraph.LineSpacing);
        templateStyle.Paragraph.LineSpacingRule = FirstNonEmpty(templateStyle.Paragraph.LineSpacingRule, extractedStyle.Paragraph.LineSpacingRule);

        return templateStyle;
    }

    private static bool LooksLikeReferenceEntry(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var normalized = NormalizeText(text);
        if (normalized.Length < 8)
            return false;

        var hasNumberPrefix =
            normalized.StartsWith("[")
            || normalized.StartsWith("\uFF3B")
            || normalized.StartsWith("(")
            || normalized.StartsWith("\uFF08")
            || char.IsDigit(normalized[0]);

        var signals = 0;
        if (normalized.Contains("[J]", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("[M]", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("[C]", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("[D]", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("[EB/OL]", StringComparison.OrdinalIgnoreCase))
        {
            signals++;
        }

        if (normalized.Contains("doi", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("vol.", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("no.", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("pp.", StringComparison.OrdinalIgnoreCase))
        {
            signals++;
        }

        if (normalized.Contains(":", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("\uFF1A", StringComparison.OrdinalIgnoreCase))
        {
            signals++;
        }

        if (normalized.Count(char.IsDigit) >= 4)
            signals++;

        return hasNumberPrefix && signals >= 1
            || signals >= 2 && normalized.Contains(".");
    }

    private static bool HasReferenceParagraphTraits(TemplateStyleProfile? profile)
    {
        if (profile == null)
            return false;

        var paragraph = profile.Paragraph;
        return (paragraph.FirstLineChars == "0" || paragraph.FirstLineIndent == "0")
            && (!string.IsNullOrWhiteSpace(paragraph.HangingChars) || !string.IsNullOrWhiteSpace(paragraph.HangingIndent));
    }

    private static string NormalizeText(string text)
    {
        return text.Replace(" ", "", StringComparison.Ordinal)
            .Replace("\t", "", StringComparison.Ordinal)
            .Replace("\r", "", StringComparison.Ordinal)
            .Replace("\n", "", StringComparison.Ordinal)
            .Trim();
    }

    private static string? FirstNonEmpty(string? primary, string? fallback)
    {
        return string.IsNullOrWhiteSpace(primary) ? fallback : primary;
    }

    private static TemplateStyleOverview? BuildStyleOverview(string key, TemplateStyleProfile? profile)
    {
        if (profile == null)
            return null;

        return new TemplateStyleOverview
        {
            Key = key,
            StyleName = profile.StyleName,
            AsciiFont = profile.Run.AsciiFont,
            EastAsiaFont = profile.Run.EastAsiaFont,
            FontFamily = profile.Run.EastAsiaFont ?? profile.Run.AsciiFont,
            FontSize = FormatFontSize(profile.Run.FontSize),
            Bold = profile.Run.Bold,
            Italic = profile.Run.Italic,
            PageBreakBefore = profile.Paragraph.PageBreakBefore,
            Align = profile.Paragraph.Align,
            LeftIndent = FormatTwips(profile.Paragraph.LeftIndent),
            RightIndent = FormatTwips(profile.Paragraph.RightIndent),
            FirstLineIndent = FormatIndent(profile.Paragraph.FirstLineChars, profile.Paragraph.FirstLineIndent),
            HangingIndent = FormatIndent(profile.Paragraph.HangingChars, profile.Paragraph.HangingIndent),
            BeforeSpacing = FormatTwips(profile.Paragraph.BeforeSpacing),
            AfterSpacing = FormatTwips(profile.Paragraph.AfterSpacing),
            LineSpacing = FormatLineSpacing(profile.Paragraph.LineSpacing, profile.Paragraph.LineSpacingRule),
            Underline = profile.Run.Underline,
            FontColor = profile.Run.FontColor,
            VerticalAlign = profile.Run.VerticalAlign
        };
    }

    private static string? FormatFontSize(string? halfPointValue)
    {
        if (string.IsNullOrWhiteSpace(halfPointValue))
            return null;

        if (!double.TryParse(halfPointValue, out var halfPoint))
            return halfPointValue;

        var point = halfPoint / 2d;
        var pointText = point % 1 == 0
            ? $"{point:0}pt"
            : $"{point:0.##}pt";
        var chineseSize = ToChineseFontSizeName(point);
        return string.IsNullOrWhiteSpace(chineseSize)
            ? pointText
            : $"{pointText}({chineseSize})";
    }

    private static string? FormatIndent(string? charsValue, string? twipsValue)
    {
        var charsText = FormatChars(charsValue);
        if (!string.IsNullOrWhiteSpace(charsText))
            return charsText;

        return FormatTwips(twipsValue);
    }

    private static string? FormatChars(string? charsValue)
    {
        if (string.IsNullOrWhiteSpace(charsValue))
            return null;

        if (!double.TryParse(charsValue, out var chars))
            return charsValue;

        var value = chars / 100d;
        return value % 1 == 0
            ? $"{value:0}瀛楃"
            : $"{value:0.##}瀛楃";
    }

    private static string? FormatTwips(string? twipsValue)
    {
        if (string.IsNullOrWhiteSpace(twipsValue))
            return null;

        if (!double.TryParse(twipsValue, out var twips))
            return twipsValue;

        var point = twips / 20d;
        return point % 1 == 0
            ? $"{point:0}pt"
            : $"{point:0.##}pt";
    }

    private static string? FormatLineSpacing(string? lineValue, string? lineRule)
    {
        if (string.IsNullOrWhiteSpace(lineValue))
            return null;

        if (!double.TryParse(lineValue, out var line))
            return lineValue;

        if (string.Equals(lineRule, "auto", StringComparison.OrdinalIgnoreCase))
        {
            var multiple = line / 240d;
            if (Math.Abs(multiple - 1d) < 0.01d)
                return "\u5355\u500D\u884C\u8DDD";
            if (Math.Abs(multiple - 1.5d) < 0.01d)
                return "1.5 \u500D\u884C\u8DDD";
            if (Math.Abs(multiple - 2d) < 0.01d)
                return "2 \u500D\u884C\u8DDD";

            return $"{multiple:0.##} \u500D\u884C\u8DDD";
        }

        var point = line / 20d;
        var pointText = point % 1 == 0
            ? $"{point:0}pt"
            : $"{point:0.##}pt";

        return string.Equals(lineRule, "exact", StringComparison.OrdinalIgnoreCase)
            ? $"\u56FA\u5B9A\u503C {pointText}"
            : pointText;
    }

    private static string? ToChineseFontSizeName(double point)
    {
        return point switch
        {
            42 => "\u521D\u53F7",
            36 => "\u5C0F\u521D",
            26 => "\u4E00\u53F7",
            24 => "\u5C0F\u4E00",
            22 => "\u4E8C\u53F7",
            18 => "\u5C0F\u4E8C",
            16 => "\u4E09\u53F7",
            15 => "\u5C0F\u4E09",
            14 => "\u56DB\u53F7",
            12 => "\u5C0F\u56DB",
            10.5 => "\u4E94\u53F7",
            9 => "\u5C0F\u4E94",
            7.5 => "\u516D\u53F7",
            6.5 => "\u5C0F\u516D",
            5.5 => "\u4E03\u53F7",
            5 => "\u516B\u53F7",
            _ => null
        };
    }

    private static string? FirstReferenceType(string? refs)
    {
        var parsed = ParseRefs(refs).FirstOrDefault();
        return parsed == null ? null : parsed.Type;
    }

    private static string? FirstReferencePath(string? refs)
    {
        var parsed = ParseRefs(refs).FirstOrDefault();
        return parsed == null ? null : parsed.Path;
    }

    private static string? FirstHeaderFooterText(IEnumerable<HeaderFooterSummary> items, string? refs)
    {
        var parsed = ParseRefs(refs).FirstOrDefault();
        if (parsed == null)
            return null;

        var match = items.FirstOrDefault(x =>
            string.Equals(x.Type, parsed.Type, StringComparison.OrdinalIgnoreCase)
            && (string.Equals(x.Path, parsed.Path, StringComparison.OrdinalIgnoreCase) || string.Equals(x.RelationshipId, parsed.Path, StringComparison.OrdinalIgnoreCase)));
        return match?.Paragraphs.FirstOrDefault()?.Text;
    }

    private static bool HasReferenceType(string? refs, string type)
    {
        return ParseRefs(refs).Any(x => string.Equals(x.Type, type, StringComparison.OrdinalIgnoreCase));
    }

    private static string? HeaderFooterTextByType(IEnumerable<HeaderFooterSummary> items, string? refs, string type)
    {
        var matchRef = ParseRefs(refs).FirstOrDefault(x => string.Equals(x.Type, type, StringComparison.OrdinalIgnoreCase));
        if (matchRef == null)
            return null;

        var match = items.FirstOrDefault(x =>
            string.Equals(x.Type, type, StringComparison.OrdinalIgnoreCase)
            && (string.Equals(x.Path, matchRef.Path, StringComparison.OrdinalIgnoreCase)
                || string.Equals(x.RelationshipId, matchRef.Path, StringComparison.OrdinalIgnoreCase)));
        return match?.Paragraphs.FirstOrDefault()?.Text;
    }

    private sealed class TemplateReference
    {
        public string Type { get; set; } = "";
        public string Path { get; set; } = "";
    }

    private static List<TemplateReference> ParseRefs(string? refs)
    {
        if (string.IsNullOrWhiteSpace(refs))
            return new List<TemplateReference>();

        return refs.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => x.Split(':', 2))
            .Where(x => x.Length == 2)
            .Select(x => new TemplateReference { Type = x[0], Path = x[1] })
            .ToList();
    }

    private static bool IsParagraphStyle(Style style)
    {
        return style.Type?.Value == StyleValues.Paragraph;
    }

    private static bool MatchesStyleId(Style style, string expected)
    {
        return string.Equals(style.StyleId?.Value, expected, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesStyleName(Style style, string expected)
    {
        return style.StyleName?.Val?.Value?.Contains(expected, StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool MatchesOutlineLevel(Style style, int outlineLevel)
    {
        return style.StyleParagraphProperties?.OutlineLevel?.Val?.Value == outlineLevel;
    }

    private static ParagraphFormatProfile CloneParagraphFormat(ParagraphFormatProfile? format)
    {
        format ??= new ParagraphFormatProfile();
        return new ParagraphFormatProfile
        {
            Align = format.Align,
            PageBreakBefore = format.PageBreakBefore,
            LeftIndent = format.LeftIndent,
            RightIndent = format.RightIndent,
            FirstLineIndent = format.FirstLineIndent,
            FirstLineChars = format.FirstLineChars,
            HangingIndent = format.HangingIndent,
            HangingChars = format.HangingChars,
            BeforeSpacing = format.BeforeSpacing,
            AfterSpacing = format.AfterSpacing,
            LineSpacing = format.LineSpacing,
            LineSpacingRule = format.LineSpacingRule
        };
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

    private static InlineRunSummary CloneInlineRun(InlineRunSummary run)
    {
        return new InlineRunSummary
        {
            Path = run.Path,
            Text = run.Text,
            Format = CloneRunFormat(run.Format)
        };
    }

    private static SemanticHints BuildHints(BodyBlockSummary block)
    {
        return new SemanticHints
        {
            IsLikelyHeading = block.HeadingLevel.HasValue,
            IsLikelyFigureCaption = LooksLikeFigureCaption(block.Text),
            IsLikelyTableCaption = LooksLikeTableCaption(block.Text),
            IsLikelyEquation = LooksLikeEquation(block.Text)
        };
    }

    private static bool LooksLikeFigureCaption(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.StartsWith("\u56FE", StringComparison.OrdinalIgnoreCase)
            || text.StartsWith("Figure", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeTableCaption(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.StartsWith("\u8868", StringComparison.OrdinalIgnoreCase)
            || text.StartsWith("Table", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeEquation(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.Contains("=", StringComparison.Ordinal)
            || text.Contains("\u516C\u5F0F", StringComparison.Ordinal);
    }
}
