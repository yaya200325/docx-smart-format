namespace DocxWeb;

internal static class BuiltinTemplateFactory
{
    public static TemplateAnalysis CreateUndergraduateThesisTemplate(
        string templateDocId,
        string universityName,
        string thesisTitle)
    {
        var oddHeaderText = $"{(string.IsNullOrWhiteSpace(universityName) ? "XXX大学" : universityName)}毕业论文";
        var evenHeaderText = string.IsNullOrWhiteSpace(thesisTitle) ? "论文标题" : thesisTitle;

        var styles = new TemplateStyleCatalog
        {
            Heading1 = CreateStyle("Heading1", "18", "宋体", "Times New Roman", true, "center", pageBreakBefore: true, before: "240", after: "120", line: "400", lineRule: "exact"),
            Heading2 = CreateStyle("Heading2", "16", "宋体", "Times New Roman", true, "left", pageBreakBefore: false, before: "200", after: "120", line: "400", lineRule: "exact"),
            Heading3 = CreateStyle("Heading3", "14", "宋体", "Times New Roman", true, "left", pageBreakBefore: false, before: "160", after: "80", line: "400", lineRule: "exact"),
            Body = CreateStyle("Normal", "12", "宋体", "Times New Roman", false, "both", pageBreakBefore: false, firstLineChars: "200", before: "0", after: "0", line: "400", lineRule: "exact"),
            Reference = CreateStyle("Reference", "12", "宋体", "Times New Roman", false, "both", pageBreakBefore: false, left: "0", right: "0", firstLineChars: "0", hangingChars: "200", before: "0", after: "0", line: "400", lineRule: "exact"),
            FigureCaption = CreateStyle("FigureCaption", "12", "宋体", "Times New Roman", false, "center", pageBreakBefore: false, before: "80", after: "80", line: "360", lineRule: "exact"),
            TableCaption = CreateStyle("TableCaption", "12", "宋体", "Times New Roman", false, "center", pageBreakBefore: false, before: "80", after: "80", line: "360", lineRule: "exact"),
            Equation = CreateStyle("Equation", "12", "宋体", "Times New Roman", false, "center", pageBreakBefore: false, before: "160", after: "160")
        };

        var headers = new List<HeaderFooterSummary>
        {
            CreateHeaderFooter("/builtin/header/odd", "header", "odd", "rIdBuiltinHeaderOdd", oddHeaderText),
            CreateHeaderFooter("/builtin/header/even", "header", "even", "rIdBuiltinHeaderEven", evenHeaderText)
        };

        var footers = new List<HeaderFooterSummary>
        {
            CreateHeaderFooter("/builtin/footer/default", "footer", "default", "rIdBuiltinFooterDefault", "第PAGE页", containsPageField: true),
            CreateHeaderFooter("/builtin/footer/even", "footer", "even", "rIdBuiltinFooterEven", "第PAGE页", containsPageField: true)
        };

        var sections = new List<SectionSummary>
        {
            new()
            {
                Path = "/section[1]",
                Type = "continuous",
                PageWidth = "11906",
                PageHeight = "16838",
                Orientation = "portrait",
                PageStart = 1,
                PageNumFmt = "decimal",
                HeaderRef = "odd:rIdBuiltinHeaderOdd;even:rIdBuiltinHeaderEven",
                FooterRef = "default:rIdBuiltinFooterDefault;even:rIdBuiltinFooterEven",
                HeaderSourceRefs = "odd:/builtin/header/odd;even:/builtin/header/even",
                FooterSourceRefs = "default:/builtin/footer/default;even:/builtin/footer/even",
                TitlePage = false,
                MarginTop = "1440",
                MarginBottom = "1440",
                MarginLeft = "1800",
                MarginRight = "1800"
            }
        };

        var document = new DocumentSummary
        {
            PageWidth = "11906",
            PageHeight = "16838",
            Orientation = "portrait",
            MarginTop = "1440",
            MarginBottom = "1440",
            MarginLeft = "1800",
            MarginRight = "1800",
            SectionCount = 1,
            HeaderCount = headers.Count,
            FooterCount = footers.Count
        };

        var rules = new TemplateRuleSet
        {
            DocumentRules = new TemplateDocumentRules
            {
                PageWidth = document.PageWidth,
                PageHeight = document.PageHeight,
                Orientation = document.Orientation,
                MarginTop = document.MarginTop,
                MarginBottom = document.MarginBottom,
                MarginLeft = document.MarginLeft,
                MarginRight = document.MarginRight,
                SectionTypes = ["continuous"],
                PageNumberFormats = ["decimal"]
            },
            HeaderFooterRules = new HeaderFooterRuleSet
            {
                HasHeader = true,
                HasFooter = true,
                HeaderTypes = ["odd", "even"],
                FooterTypes = ["default", "even"],
                ContainsPageNumber = true,
                ContainsTotalPages = false
            },
            StyleProfiles = styles,
            CaptionRules = new CaptionRuleSet
            {
                FigurePrefixes = ["图", "Figure"],
                TablePrefixes = ["表", "Table"],
                FigureCaptionStyleKey = "template.figureCaption",
                TableCaptionStyleKey = "template.tableCaption",
                FigureCaptionPosition = "below",
                TableCaptionPosition = "above"
            },
            OutlineRules = new OutlineRuleSet
            {
                HeadingStyleKeys = ["template.heading1", "template.heading2", "template.heading3"],
                MaxHeadingLevel = 3,
                HeadingRecognitionHints = ["styleName", "headingLevel", "numberedPrefix"]
            },
            NormalizedRequirements =
            [
                "正文小四，中文宋体，英文 Times New Roman，首行缩进 2 字符，固定值 20pt 行距。",
                "参考文献小四，中文宋体，英文 Times New Roman，两端对齐，悬挂缩进 2 字符，固定值 20pt 行距。",
                "页眉奇偶页不同，奇数页为学校毕业论文，偶数页为论文标题。",
                "页脚居中页码，小四字号。"
            ]
        };

        return new TemplateAnalysis
        {
            TemplateDocId = templateDocId,
            Document = document,
            Styles = styles,
            Headers = headers,
            Footers = footers,
            Sections = sections,
            Rules = rules
        };
    }

    private static TemplateStyleProfile CreateStyle(
        string styleName,
        string fontSizePt,
        string eastAsiaFont,
        string asciiFont,
        bool bold,
        string align,
        bool pageBreakBefore,
        string? left = null,
        string? right = null,
        string? firstLineChars = null,
        string? hangingChars = null,
        string? before = null,
        string? after = null,
        string? line = null,
        string? lineRule = null)
    {
        var halfPoint = double.TryParse(fontSizePt, out var pt)
            ? (pt * 2d).ToString("0.##")
            : fontSizePt;

        return new TemplateStyleProfile
        {
            StyleName = styleName,
            Paragraph = new ParagraphFormatProfile
            {
                Align = align,
                PageBreakBefore = pageBreakBefore,
                LeftIndent = left,
                RightIndent = right,
                FirstLineChars = firstLineChars,
                HangingChars = hangingChars,
                BeforeSpacing = before,
                AfterSpacing = after,
                LineSpacing = line,
                LineSpacingRule = lineRule
            },
            Run = new RunFormatProfile
            {
                FontSize = halfPoint,
                EastAsiaFont = eastAsiaFont,
                AsciiFont = asciiFont,
                Bold = bold
            }
        };
    }

    private static HeaderFooterSummary CreateHeaderFooter(
        string path,
        string kind,
        string type,
        string relationshipId,
        string text,
        bool containsPageField = false)
    {
        return new HeaderFooterSummary
        {
            Path = path,
            RelationshipId = relationshipId,
            Kind = kind,
            Type = type,
            ContainsPageField = containsPageField,
            Paragraphs =
            [
                new HeaderFooterParagraph
                {
                    Path = $"{path}/paragraph[1]",
                    Text = text,
                    Align = "center",
                    Run = new RunFormatProfile
                    {
                        EastAsiaFont = "宋体",
                        AsciiFont = "Times New Roman",
                        FontSize = "24"
                    }
                }
            ]
        };
    }
}
