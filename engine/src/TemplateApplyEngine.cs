namespace DocxWeb;

internal static class TemplateApplyEngine
{
    public static RenderSpec BuildRenderSpec(
        TemplateAnalysis template,
        SourceDocumentAnalysis source,
        FormatDecisionEnvelope decision)
    {
        var headerFooterPolicyMap = BuildHeaderFooterPolicyMap(decision.HeaderFooterDecisions);
        var styleResolver = BuildStyleResolver(template);
        var styleProfileMap = BuildStyleProfileMap(template);
        var decisionMap = decision.BlockDecisions
            .Where(x => !string.IsNullOrWhiteSpace(x.BlockId))
            .GroupBy(x => x.BlockId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
        var spec = new RenderSpec
        {
            Document = new RenderDocumentSpec
            {
                Title = source.Document.Title,
                Author = source.Document.Author,
                PageWidth = template.Document.PageWidth ?? source.Document.PageWidth,
                PageHeight = template.Document.PageHeight ?? source.Document.PageHeight,
                Orientation = template.Document.Orientation ?? source.Document.Orientation,
                MarginTop = template.Document.MarginTop ?? source.Document.MarginTop,
                MarginBottom = template.Document.MarginBottom ?? source.Document.MarginBottom,
                MarginLeft = template.Document.MarginLeft ?? source.Document.MarginLeft,
                MarginRight = template.Document.MarginRight ?? source.Document.MarginRight
            },
            Headers = BuildHeaderFooterSpecs(template.Headers, decision.HeaderFooterDecisions, decision.DocumentRules.UseTemplateHeaders, "header"),
            Footers = BuildHeaderFooterSpecs(template.Footers, decision.HeaderFooterDecisions, decision.DocumentRules.UseTemplateFooters, "footer")
        };

        var bodyBlocks = new List<RenderBodyBlock>();
        var referenceBlockIds = CollectReferenceBlockIds(source.Blocks, decisionMap);
        foreach (var block in source.Blocks)
        {
            decisionMap.TryGetValue(block.BlockId, out var rawDecisionItem);
            var decisionItem = ResolveBlockDecision(block, rawDecisionItem, referenceBlockIds);
            bodyBlocks.Add(ToRenderBodyBlock(block, decisionItem, styleResolver, styleProfileMap));
        }

        PrependTemplateFirstSection(bodyBlocks, template);
        ApplyCaptionDecisions(bodyBlocks, decision.CaptionDecisions, styleResolver, styleProfileMap);
        ApplyRunDecisions(bodyBlocks, decision.RunDecisions);
        ApplySectionDecisions(bodyBlocks, template, source, decision.SectionDecisions, headerFooterPolicyMap);
        spec.Body.AddRange(bodyBlocks);

        return spec;
    }

    private static BlockDecision? ResolveBlockDecision(
        SemanticBlock block,
        BlockDecision? explicitDecision,
        ISet<string> referenceBlockIds)
    {
        if (string.Equals(block.Kind, "paragraph", StringComparison.OrdinalIgnoreCase)
            && referenceBlockIds.Contains(block.BlockId)
            && !IsReferenceDecision(explicitDecision))
        {
            return new BlockDecision
            {
                BlockId = block.BlockId,
                SemanticRole = "reference",
                ApplyStyleKey = "template.reference"
            };
        }

        if (explicitDecision != null
            && (!string.IsNullOrWhiteSpace(explicitDecision.ApplyStyleKey)
                || !string.IsNullOrWhiteSpace(explicitDecision.SemanticRole)))
        {
            return explicitDecision;
        }

        if (!string.Equals(block.Kind, "paragraph", StringComparison.OrdinalIgnoreCase))
            return explicitDecision;

        var semanticRole = InferSemanticRole(block, referenceBlockIds);
        var applyStyleKey = semanticRole switch
        {
            "heading1" => "template.heading1",
            "heading2" => "template.heading2",
            "heading3" => "template.heading3",
            "figurecaption" => "template.figureCaption",
            "tablecaption" => "template.tableCaption",
            "equation" => "template.equation",
            "reference" => "template.reference",
            "body" => "template.body",
            _ => "template.body"
        };

        return new BlockDecision
        {
            BlockId = block.BlockId,
            SemanticRole = semanticRole,
            ApplyStyleKey = applyStyleKey
        };
    }

    private static string InferSemanticRole(SemanticBlock block, ISet<string> referenceBlockIds)
    {
        if (LooksLikePrimaryHeading(block.Text))
            return "heading1";

        if (block.HeadingLevel is >= 1 and <= 3)
            return $"heading{block.HeadingLevel.Value}";

        if (referenceBlockIds.Contains(block.BlockId))
            return "reference";

        if (block.Hints.IsLikelyFigureCaption)
            return "figureCaption";

        if (block.Hints.IsLikelyTableCaption)
            return "tableCaption";

        if (block.Hints.IsLikelyEquation)
            return "equation";

        return "body";
    }

    private static bool LooksLikePrimaryHeading(string? text)
    {
        var normalized = NormalizeText(text);
        if (string.IsNullOrWhiteSpace(normalized))
            return false;

        if (normalized is "引言" or "绪论")
            return true;

        if (System.Text.RegularExpressions.Regex.IsMatch(normalized, @"^第\s*[0-9一二三四五六七八九十百千万]+\s*章"))
            return true;

        if (System.Text.RegularExpressions.Regex.IsMatch(normalized, @"^chapter\s*1\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;

        return false;
    }

    private static HashSet<string> CollectReferenceBlockIds(
        IReadOnlyList<SemanticBlock> blocks,
        IReadOnlyDictionary<string, BlockDecision> decisionMap)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var inReferenceSection = false;

        foreach (var block in blocks)
        {
            if (!string.Equals(block.Kind, "paragraph", StringComparison.OrdinalIgnoreCase))
                continue;

            if (decisionMap.TryGetValue(block.BlockId, out var explicitDecision)
                && IsReferenceDecision(explicitDecision))
            {
                result.Add(block.BlockId);
                inReferenceSection = true;
                continue;
            }

            if (LooksLikeReferenceHeading(block))
            {
                inReferenceSection = true;
                continue;
            }

            if (!inReferenceSection)
                continue;

            if (IsSectionBoundary(block))
            {
                inReferenceSection = false;
                continue;
            }

            if (LooksLikeReferenceEntry(block.Text))
            {
                result.Add(block.BlockId);
                continue;
            }

            if (IsBlankParagraph(block))
                continue;

            inReferenceSection = false;
        }

        return result;
    }

    private static bool IsReferenceDecision(BlockDecision? decision)
    {
        if (decision == null)
            return false;

        return string.Equals(decision.SemanticRole, "reference", StringComparison.OrdinalIgnoreCase)
            || string.Equals(decision.ApplyStyleKey, "template.reference", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeReferenceHeading(SemanticBlock block)
    {
        var normalized = NormalizeText(block.Text);
        return normalized is "\u53c2\u8003\u6587\u732e" or "\u53c2\u8003\u8d44\u6599" or "references" or "bibliography";
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

    private static bool IsSectionBoundary(SemanticBlock block)
    {
        if (block.HeadingLevel is >= 1 and <= 3)
            return true;

        if (block.Hints.IsLikelyFigureCaption || block.Hints.IsLikelyTableCaption || block.Hints.IsLikelyEquation)
            return true;

        return false;
    }

    private static bool IsBlankParagraph(SemanticBlock block)
    {
        return string.IsNullOrWhiteSpace(block.Text);
    }

    private static string NormalizeText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        return text.Replace(" ", "", StringComparison.Ordinal)
            .Replace("\t", "", StringComparison.Ordinal)
            .Replace("\r", "", StringComparison.Ordinal)
            .Replace("\n", "", StringComparison.Ordinal)
            .Trim()
            .ToLowerInvariant();
    }

    private static void PrependTemplateFirstSection(List<RenderBodyBlock> blocks, TemplateAnalysis template)
    {
        var firstSection = template.Sections.FirstOrDefault();
        if (firstSection == null)
            return;

        blocks.Insert(0, new RenderBodyBlock
        {
            Kind = "section",
            Section = BuildTemplateSectionSpec(firstSection, template.Document)
        });
    }

    private static Dictionary<string, string?> BuildStyleResolver(TemplateAnalysis template)
    {
        return new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
        {
            ["template.heading1"] = template.Styles.Heading1?.StyleName,
            ["template.heading2"] = template.Styles.Heading2?.StyleName,
            ["template.heading3"] = template.Styles.Heading3?.StyleName,
            ["template.body"] = template.Styles.Body?.StyleName,
            ["template.reference"] = template.Styles.Reference?.StyleName,
            ["template.figurecaption"] = template.Styles.FigureCaption?.StyleName,
            ["template.tablecaption"] = template.Styles.TableCaption?.StyleName,
            ["template.equation"] = template.Styles.Equation?.StyleName
        };
    }

    private static Dictionary<string, TemplateStyleProfile?> BuildStyleProfileMap(TemplateAnalysis template)
    {
        return new Dictionary<string, TemplateStyleProfile?>(StringComparer.OrdinalIgnoreCase)
        {
            ["template.heading1"] = template.Styles.Heading1,
            ["template.heading2"] = template.Styles.Heading2,
            ["template.heading3"] = template.Styles.Heading3,
            ["template.body"] = template.Styles.Body,
            ["template.reference"] = template.Styles.Reference,
            ["template.figurecaption"] = template.Styles.FigureCaption,
            ["template.tablecaption"] = template.Styles.TableCaption,
            ["template.equation"] = template.Styles.Equation
        };
    }

    private static Dictionary<string, HeaderFooterDecision> BuildHeaderFooterPolicyMap(IEnumerable<HeaderFooterDecision> decisions)
    {
        return decisions
            .Where(x => !string.IsNullOrWhiteSpace(x.DecisionKey))
            .GroupBy(x => x.DecisionKey!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
    }

    private static List<RenderHeaderFooterSpec> BuildHeaderFooterSpecs(
        IEnumerable<HeaderFooterSummary> summaries,
        IEnumerable<HeaderFooterDecision> decisions,
        bool enabled,
        string kind)
    {
        if (!enabled)
            return [];

        var result = new List<RenderHeaderFooterSpec>();
        foreach (var summary in summaries)
        {
            var decision = decisions.FirstOrDefault(x =>
                string.Equals(x.Kind, kind, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.Type, summary.Type, StringComparison.OrdinalIgnoreCase));

            if (decision != null && string.Equals(decision.Action, "omit", StringComparison.OrdinalIgnoreCase))
                continue;

            result.Add(ToRenderHeaderFooter(summary));
        }

        return result;
    }

    private static void ApplyCaptionDecisions(
        List<RenderBodyBlock> blocks,
        IEnumerable<CaptionDecision> decisions,
        Dictionary<string, string?> styleResolver,
        Dictionary<string, TemplateStyleProfile?> styleProfileMap)
    {
        foreach (var captionDecision in decisions)
        {
            var captionIndex = blocks.FindIndex(x => string.Equals(x.SourceBlockId, captionDecision.BlockId, StringComparison.OrdinalIgnoreCase));
            if (captionIndex < 0)
                continue;

            var targetIndex = FindNeighborAssetIndex(blocks, captionIndex, captionDecision.CaptionKind);
            if (targetIndex < 0)
                continue;

            var captionBlock = blocks[captionIndex];
            if (captionBlock.Paragraph != null && !string.IsNullOrWhiteSpace(captionDecision.ApplyStyleKey))
            {
                captionBlock.Paragraph.Style = MapStyle(captionDecision.ApplyStyleKey, captionBlock.Paragraph.Style, captionBlock.Paragraph.HeadingLevel, styleResolver);
                ApplyTemplateStyle(captionBlock.Paragraph, captionDecision.ApplyStyleKey, styleProfileMap);
            }

            blocks.RemoveAt(captionIndex);
            var insertIndex = string.Equals(captionDecision.Position, "above", StringComparison.OrdinalIgnoreCase)
                ? Math.Min(targetIndex, blocks.Count)
                : Math.Min(targetIndex + 1, blocks.Count);
            blocks.Insert(insertIndex, captionBlock);
        }
    }

    private static void ApplyRunDecisions(
        List<RenderBodyBlock> blocks,
        IEnumerable<RunDecision> decisions)
    {
        foreach (var decision in decisions)
        {
            if (string.IsNullOrWhiteSpace(decision.BlockId))
                continue;

            var paragraph = blocks
                .FirstOrDefault(x => string.Equals(x.SourceBlockId, decision.BlockId, StringComparison.OrdinalIgnoreCase))
                ?.Paragraph;
            if (paragraph == null)
                continue;

            if (paragraph.Runs.Count == 0 && !string.IsNullOrWhiteSpace(paragraph.Text))
            {
                paragraph.Runs.Add(new InlineRunSummary
                {
                    Path = $"{decision.BlockId}/run[1]",
                    Text = paragraph.Text,
                    Format = CloneRunFormat(paragraph.Run)
                });
            }

            var matchedRuns = paragraph.Runs
                .Where(run => MatchRunDecision(run, decision))
                .ToList();
            if (decision.Occurrence.HasValue && decision.Occurrence.Value > 0)
                matchedRuns = matchedRuns.Skip(decision.Occurrence.Value - 1).Take(1).ToList();

            foreach (var run in matchedRuns)
            {
                if (decision.ReplaceText != null)
                    run.Text = decision.ReplaceText;
                run.Format = ApplyRunDecision(run.Format, decision);
                if (run.Segments.Count > 0)
                    ApplySegmentDecision(run.Segments, decision);
            }

            if (paragraph.Runs.Count > 0)
                paragraph.Run = CloneRunFormat(paragraph.Runs[0].Format);
        }
    }

    private static void ApplySectionDecisions(
        List<RenderBodyBlock> blocks,
        TemplateAnalysis template,
        SourceDocumentAnalysis source,
        IEnumerable<SectionDecision> decisions,
        Dictionary<string, HeaderFooterDecision> headerFooterPolicyMap)
    {
        var sectionDecisions = decisions.ToList();
        if (sectionDecisions.Count == 0)
        {
            ApplyTemplateSections(blocks, template, source);
            return;
        }

        foreach (var sectionDecision in sectionDecisions)
        {
            var insertIndex = string.IsNullOrWhiteSpace(sectionDecision.AfterBlockId)
                ? blocks.Count
                : blocks.FindIndex(x => string.Equals(x.SourceBlockId, sectionDecision.AfterBlockId, StringComparison.OrdinalIgnoreCase));

            if (insertIndex < 0)
                insertIndex = blocks.Count - 1;

            var headerPolicy = ResolveHeaderFooterPolicy(sectionDecision.HeaderDecisionKey, headerFooterPolicyMap);
            var footerPolicy = ResolveHeaderFooterPolicy(sectionDecision.FooterDecisionKey, headerFooterPolicyMap);

            var sectionBlock = new RenderBodyBlock
            {
                Kind = "section",
                Section = new RenderSectionSpec
                {
                    Type = sectionDecision.SectionType,
                    PageStart = sectionDecision.PageStart,
                    PageNumFmt = null,
                    Orientation = sectionDecision.Orientation,
                    Headers = BuildExplicitHeaderFooterRefs(sectionDecision.HeaderType, sectionDecision.HeaderSourcePath),
                    Footers = BuildExplicitHeaderFooterRefs(sectionDecision.FooterType, sectionDecision.FooterSourcePath),
                    HeaderType = sectionDecision.HeaderType ?? headerPolicy?.Type,
                    FooterType = sectionDecision.FooterType ?? footerPolicy?.Type,
                    HeaderSourcePath = sectionDecision.HeaderSourcePath ?? headerPolicy?.SourcePath,
                    FooterSourcePath = sectionDecision.FooterSourcePath ?? footerPolicy?.SourcePath
                }
            };

            blocks.Insert(Math.Min(insertIndex + 1, blocks.Count), sectionBlock);
        }
    }

    private static void ApplyTemplateSections(
        List<RenderBodyBlock> blocks,
        TemplateAnalysis template,
        SourceDocumentAnalysis source)
    {
        var templateSections = template.Sections;
        if (templateSections.Count == 0)
            return;

        var sourceSections = source.Sections.OrderBy(x => ExtractSectionOrder(x.Path)).ToList();
        for (var i = 1; i < templateSections.Count && i <= sourceSections.Count; i++)
        {
            var sourceSection = sourceSections[i - 1];
            var lastSourceBlock = source.Blocks.LastOrDefault(x => x.SectionIndex == i);

            var insertAfterBlockId = lastSourceBlock?.BlockId;
            var insertIndex = string.IsNullOrWhiteSpace(insertAfterBlockId)
                ? blocks.Count - 1
                : blocks.FindIndex(x => string.Equals(x.SourceBlockId, insertAfterBlockId, StringComparison.OrdinalIgnoreCase));

            var templateSection = templateSections[i];
            var sectionSpec = BuildTemplateSectionSpec(templateSection, template.Document);
            blocks.Insert(Math.Min(insertIndex + 1, blocks.Count), new RenderBodyBlock
            {
                Kind = "section",
                Section = sectionSpec
            });
        }

        if (blocks.Count == 0)
            blocks.Add(new RenderBodyBlock { Kind = "section", Section = BuildTemplateSectionSpec(templateSections[0], template.Document) });
    }

    private static RenderSectionSpec BuildTemplateSectionSpec(SectionSummary templateSection, DocumentSummary templateDocument)
    {
        var headers = ParseHeaderFooterRefs(templateSection.HeaderSourceRefs);
        var footers = ParseHeaderFooterRefs(templateSection.FooterSourceRefs);

        return new RenderSectionSpec
        {
            Type = templateSection.Type,
            PageWidth = templateSection.PageWidth ?? templateDocument.PageWidth,
            PageHeight = templateSection.PageHeight ?? templateDocument.PageHeight,
            Orientation = templateSection.Orientation ?? templateDocument.Orientation,
            MarginTop = templateSection.MarginTop ?? templateDocument.MarginTop,
            MarginBottom = templateSection.MarginBottom ?? templateDocument.MarginBottom,
            MarginLeft = templateSection.MarginLeft ?? templateDocument.MarginLeft,
            MarginRight = templateSection.MarginRight ?? templateDocument.MarginRight,
            PageStart = templateSection.PageStart,
            PageNumFmt = templateSection.PageNumFmt,
            TitlePage = templateSection.TitlePage,
            HeaderType = headers.FirstOrDefault()?.Type,
            FooterType = footers.FirstOrDefault()?.Type,
            HeaderSourcePath = headers.FirstOrDefault()?.SourcePath,
            FooterSourcePath = footers.FirstOrDefault()?.SourcePath,
            Headers = headers,
            Footers = footers
        };
    }

    private static List<RenderHeaderFooterReferenceSpec> ParseHeaderFooterRefs(string? refs)
    {
        if (string.IsNullOrWhiteSpace(refs))
            return [];

        return refs
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => x.Split(':', 2))
            .Where(x => x.Length == 2)
            .Select(x => new RenderHeaderFooterReferenceSpec
            {
                Type = x[0],
                SourcePath = x[1]
            })
            .ToList();
    }

    private static List<RenderHeaderFooterReferenceSpec> BuildExplicitHeaderFooterRefs(string? type, string? sourcePath)
    {
        if (string.IsNullOrWhiteSpace(type) && string.IsNullOrWhiteSpace(sourcePath))
            return [];

        return
        [
            new RenderHeaderFooterReferenceSpec
            {
                Type = string.IsNullOrWhiteSpace(type) ? "default" : type,
                SourcePath = sourcePath
            }
        ];
    }

    private static int ExtractSectionOrder(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return int.MaxValue;

        var start = path.IndexOf('[');
        var end = path.IndexOf(']');
        if (start >= 0 && end > start && int.TryParse(path[(start + 1)..end], out var value))
            return value;

        return int.MaxValue;
    }

    private static HeaderFooterDecision? ResolveHeaderFooterPolicy(
        string? decisionKey,
        Dictionary<string, HeaderFooterDecision> headerFooterPolicyMap)
    {
        if (string.IsNullOrWhiteSpace(decisionKey))
            return null;

        return headerFooterPolicyMap.TryGetValue(decisionKey, out var policy) ? policy : null;
    }

    private static int FindNeighborAssetIndex(List<RenderBodyBlock> blocks, int captionIndex, string captionKind)
    {
        var expectedKind = string.Equals(captionKind, "table", StringComparison.OrdinalIgnoreCase) ? "table" : "image";

        for (var i = captionIndex - 1; i >= 0; i--)
        {
            if (string.Equals(blocks[i].Kind, expectedKind, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        for (var i = captionIndex + 1; i < blocks.Count; i++)
        {
            if (string.Equals(blocks[i].Kind, expectedKind, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        return -1;
    }

    private static RenderHeaderFooterSpec ToRenderHeaderFooter(HeaderFooterSummary summary)
    {
        return new RenderHeaderFooterSpec
        {
            SourcePath = summary.Path,
            RelationshipId = summary.RelationshipId,
            Kind = summary.Kind,
            Type = summary.Type,
            Paragraphs = summary.Paragraphs.Select(x => new RenderParagraphSpec
            {
                Text = summary.ContainsPageField ? null : x.Text,
                PageNumberPrefix = GetPagePlaceholderPrefix(x.Text),
                PageNumberSuffix = GetPagePlaceholderSuffix(x.Text),
                Style = x.Style,
                Run = CloneRunFormat(x.Run),
                Align = x.Align,
                Paragraph = new ParagraphFormatProfile
                {
                    Align = x.Align
                },
                PageNumberField = summary.ContainsPageField && ContainsPagePlaceholder(x.Text),
                NumPagesField = summary.ContainsNumPagesField
            }).ToList()
        };
    }

    private static bool ContainsPagePlaceholder(string? text)
    {
        return !string.IsNullOrWhiteSpace(text)
            && text.Contains("PAGE", StringComparison.OrdinalIgnoreCase);
    }

    private static string? NormalizeHeaderFooterText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        return text.Replace("PAGE", "", StringComparison.OrdinalIgnoreCase)
            .Replace("  ", " ", StringComparison.Ordinal)
            .Trim();
    }

    private static string? GetPagePlaceholderPrefix(string? text)
    {
        if (!ContainsPagePlaceholder(text))
            return null;

        var index = text!.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase);
        return index <= 0 ? string.Empty : text[..index];
    }

    private static string? GetPagePlaceholderSuffix(string? text)
    {
        if (!ContainsPagePlaceholder(text))
            return null;

        var index = text!.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase);
        var suffixIndex = index + 4;
        return suffixIndex >= text.Length ? string.Empty : text[suffixIndex..];
    }

    private static RenderBodyBlock ToRenderBodyBlock(
        SemanticBlock block,
        BlockDecision? decision,
        Dictionary<string, string?> styleResolver,
        Dictionary<string, TemplateStyleProfile?> styleProfileMap)
    {
        var styleKey = decision?.ApplyStyleKey?.ToLowerInvariant();
        var role = decision?.SemanticRole?.ToLowerInvariant();

        return block.Kind.ToLowerInvariant() switch
        {
            "table" => new RenderBodyBlock
            {
                Kind = "table",
                SourceBlockId = block.BlockId,
                Table = block.Table
            },
            "image" => new RenderBodyBlock
            {
                Kind = "image",
                SourceBlockId = block.BlockId,
                Image = block.Image == null ? null : new RenderImageSpec
                {
                    AssetId = block.Image.AssetId,
                    Path = block.Image.Path,
                    ContentType = block.Image.ContentType,
                    WidthEmu = block.Image.WidthEmu,
                    HeightEmu = block.Image.HeightEmu,
                    AltText = block.Image.AltText
                }
            },
            "equation" => new RenderBodyBlock
            {
                Kind = "equation",
                SourceBlockId = block.BlockId,
                Equation = block.Equation == null ? null : new RenderEquationSpec
                {
                    Text = block.Equation.Text,
                    Xml = block.Equation.Xml,
                    DisplayMode = block.Equation.DisplayMode,
                    Paragraph = BuildEquationParagraphFormat(styleProfileMap),
                    Run = BuildEquationRunFormat(styleProfileMap)
                }
            },
            _ => new RenderBodyBlock
            {
                Kind = "paragraph",
                SourceBlockId = block.BlockId,
                Paragraph = BuildParagraphSpec(block, decision, styleKey, role, styleResolver, styleProfileMap)
            }
        };
    }

    private static RenderParagraphSpec BuildParagraphSpec(
        SemanticBlock block,
        BlockDecision? decision,
        string? styleKey,
        string? role,
        Dictionary<string, string?> styleResolver,
        Dictionary<string, TemplateStyleProfile?> styleProfileMap)
    {
        var effectiveStyleKey = NormalizeStyleKey(styleKey, role);
        var paragraphSpec = new RenderParagraphSpec
        {
            Text = NormalizeParagraphText(block.Text, effectiveStyleKey),
            Style = MapStyle(effectiveStyleKey, block.Style, block.HeadingLevel, styleResolver),
            HeadingLevel = MapHeadingLevel(role, block.HeadingLevel),
            Paragraph = CloneParagraphFormat(block.Paragraph),
            Run = CloneRunFormat(block.Run),
            Runs = NormalizeInlineRuns(block.Runs.Select(CloneInlineRun).ToList(), effectiveStyleKey),
            Align = block.Paragraph.Align,
            PageBreakBefore = block.Paragraph.PageBreakBefore
        };

        ApplyTemplateStyle(paragraphSpec, effectiveStyleKey, styleProfileMap);
        ApplyBlockDecision(paragraphSpec, decision);
        return paragraphSpec;
    }

    private static void ApplyBlockDecision(RenderParagraphSpec paragraphSpec, BlockDecision? decision)
    {
        if (decision == null)
            return;

        paragraphSpec.Align = decision.Align ?? paragraphSpec.Align;
        if (decision.PageBreakBefore.HasValue)
            paragraphSpec.PageBreakBefore = decision.PageBreakBefore.Value;

        paragraphSpec.Paragraph = new ParagraphFormatProfile
        {
            Align = decision.Align ?? paragraphSpec.Paragraph.Align,
            PageBreakBefore = decision.PageBreakBefore ?? paragraphSpec.Paragraph.PageBreakBefore,
            LeftIndent = decision.LeftIndent ?? paragraphSpec.Paragraph.LeftIndent,
            RightIndent = decision.RightIndent ?? paragraphSpec.Paragraph.RightIndent,
            FirstLineIndent = decision.FirstLineIndent ?? paragraphSpec.Paragraph.FirstLineIndent,
            FirstLineChars = decision.FirstLineChars ?? paragraphSpec.Paragraph.FirstLineChars,
            HangingIndent = decision.HangingIndent ?? paragraphSpec.Paragraph.HangingIndent,
            HangingChars = decision.HangingChars ?? paragraphSpec.Paragraph.HangingChars,
            BeforeSpacing = decision.BeforeSpacing ?? paragraphSpec.Paragraph.BeforeSpacing,
            AfterSpacing = decision.AfterSpacing ?? paragraphSpec.Paragraph.AfterSpacing,
            LineSpacing = decision.LineSpacing ?? paragraphSpec.Paragraph.LineSpacing,
            LineSpacingRule = decision.LineSpacingRule ?? paragraphSpec.Paragraph.LineSpacingRule
        };

        paragraphSpec.Run = ApplyBlockRunDecision(paragraphSpec.Run, decision);
        if (paragraphSpec.Runs.Count > 0)
        {
            paragraphSpec.Runs = paragraphSpec.Runs
                .Select(run => new InlineRunSummary
                {
                    Path = run.Path,
                    Text = run.Text,
                    Format = ApplyBlockRunDecision(run.Format, decision),
                    Segments = run.Segments.Select(segment => new InlineTextSegment
                    {
                        Path = segment.Path,
                        Text = segment.Text,
                        Format = ApplyBlockRunDecision(segment.Format, decision)
                    }).ToList()
                })
                .ToList();
        }
    }

    private static RunFormatProfile ApplyBlockRunDecision(RunFormatProfile format, BlockDecision decision)
    {
        return new RunFormatProfile
        {
            FontSize = decision.FontSize ?? format.FontSize,
            AsciiFont = decision.AsciiFont ?? format.AsciiFont,
            EastAsiaFont = decision.EastAsiaFont ?? format.EastAsiaFont,
            Bold = decision.Bold ?? format.Bold,
            Italic = decision.Italic ?? format.Italic,
            Underline = decision.Underline ?? format.Underline,
            FontColor = decision.FontColor ?? format.FontColor,
            Highlight = decision.Highlight ?? format.Highlight,
            VerticalAlign = decision.VerticalAlign ?? format.VerticalAlign
        };
    }

    private static string? NormalizeParagraphText(string? text, string? styleKey)
    {
        if (!string.Equals(styleKey, "template.reference", StringComparison.OrdinalIgnoreCase))
            return text;

        return NormalizeReferenceSpacing(text);
    }

    private static List<InlineRunSummary> NormalizeInlineRuns(List<InlineRunSummary> runs, string? styleKey)
    {
        if (!string.Equals(styleKey, "template.reference", StringComparison.OrdinalIgnoreCase))
            return runs;

        return runs
            .Select(run => new InlineRunSummary
            {
                Path = run.Path,
                Text = NormalizeReferenceSpacing(run.Text),
                Format = CloneRunFormat(run.Format),
                Segments = run.Segments.Select(segment => new InlineTextSegment
                {
                    Path = segment.Path,
                    Text = NormalizeReferenceSpacing(segment.Text),
                    Format = CloneRunFormat(segment.Format)
                }).ToList()
            })
            .ToList();
    }

    private static string? NormalizeReferenceSpacing(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        return System.Text.RegularExpressions.Regex.Replace(text, @"^\[(\d+)\]\s*", "[$1]\t");
    }

    private static void ApplyTemplateStyle(
        RenderParagraphSpec paragraphSpec,
        string? styleKey,
        Dictionary<string, TemplateStyleProfile?> styleProfileMap)
    {
        if (string.IsNullOrWhiteSpace(styleKey))
            return;

        if (!styleProfileMap.TryGetValue(styleKey, out var profile) || profile == null)
            return;

        paragraphSpec.Paragraph = MergeParagraphFormat(profile.Paragraph, paragraphSpec.Paragraph);
        paragraphSpec.Run = MergeRunFormat(profile.Run, paragraphSpec.Run);
        paragraphSpec.Align = profile.Paragraph.Align ?? paragraphSpec.Align;
        paragraphSpec.PageBreakBefore = profile.Paragraph.PageBreakBefore || paragraphSpec.PageBreakBefore;

        if (paragraphSpec.Runs.Count > 0)
        {
            paragraphSpec.Runs = paragraphSpec.Runs
                .Select(run => new InlineRunSummary
                {
                    Path = run.Path,
                    Text = run.Text,
                    Format = MergeRunFormat(profile.Run, run.Format),
                    Segments = run.Segments.Select(segment => new InlineTextSegment
                    {
                        Path = segment.Path,
                        Text = segment.Text,
                        Format = MergeRunFormat(profile.Run, segment.Format)
                    }).ToList()
                })
                .ToList();
        }
    }

    private static string? MapStyle(
        string? styleKey,
        string? fallbackStyle,
        int? fallbackHeadingLevel,
        Dictionary<string, string?> styleResolver)
    {
        if (!string.IsNullOrWhiteSpace(styleKey)
            && styleResolver.TryGetValue(styleKey, out var resolvedStyle)
            && !string.IsNullOrWhiteSpace(resolvedStyle))
        {
            return resolvedStyle;
        }

        return styleKey switch
        {
            "template.heading1" => "Heading1",
            "template.heading2" => "Heading2",
            "template.heading3" => "Heading3",
            "template.body" => "Normal",
            "template.reference" => "Reference",
            "template.figurecaption" => "Caption",
            "template.tablecaption" => "Caption",
            _ => string.IsNullOrWhiteSpace(fallbackStyle) && fallbackHeadingLevel.HasValue
                ? $"Heading{fallbackHeadingLevel.Value}"
                : fallbackStyle
        };
    }

    private static string? NormalizeStyleKey(string? styleKey, string? role)
    {
        if (!string.IsNullOrWhiteSpace(styleKey))
            return styleKey;

        return role switch
        {
            "heading1" => "template.heading1",
            "heading2" => "template.heading2",
            "heading3" => "template.heading3",
            "body" => "template.body",
            "reference" => "template.reference",
            "figurecaption" => "template.figureCaption",
            "tablecaption" => "template.tableCaption",
            "equation" => "template.equation",
            _ => styleKey
        };
    }

    private static int? MapHeadingLevel(string? role, int? fallback)
    {
        return role switch
        {
            "heading1" => 1,
            "heading2" => 2,
            "heading3" => 3,
            _ => fallback
        };
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
            Format = CloneRunFormat(run.Format),
            Segments = run.Segments.Select(segment => new InlineTextSegment
            {
                Path = segment.Path,
                Text = segment.Text,
                Format = CloneRunFormat(segment.Format)
            }).ToList()
        };
    }

    private static ParagraphFormatProfile MergeParagraphFormat(ParagraphFormatProfile templateFormat, ParagraphFormatProfile currentFormat)
    {
        return new ParagraphFormatProfile
        {
            Align = templateFormat.Align ?? currentFormat.Align,
            PageBreakBefore = templateFormat.PageBreakBefore || currentFormat.PageBreakBefore,
            LeftIndent = templateFormat.LeftIndent ?? currentFormat.LeftIndent,
            RightIndent = templateFormat.RightIndent ?? currentFormat.RightIndent,
            FirstLineIndent = templateFormat.FirstLineIndent ?? currentFormat.FirstLineIndent,
            FirstLineChars = templateFormat.FirstLineChars ?? currentFormat.FirstLineChars,
            HangingIndent = templateFormat.HangingIndent ?? currentFormat.HangingIndent,
            HangingChars = templateFormat.HangingChars ?? currentFormat.HangingChars,
            BeforeSpacing = templateFormat.BeforeSpacing ?? currentFormat.BeforeSpacing,
            AfterSpacing = templateFormat.AfterSpacing ?? currentFormat.AfterSpacing,
            LineSpacing = templateFormat.LineSpacing ?? currentFormat.LineSpacing,
            LineSpacingRule = templateFormat.LineSpacingRule ?? currentFormat.LineSpacingRule
        };
    }

    private static RunFormatProfile MergeRunFormat(RunFormatProfile templateFormat, RunFormatProfile currentFormat)
    {
        return new RunFormatProfile
        {
            FontSize = templateFormat.FontSize ?? currentFormat.FontSize,
            AsciiFont = templateFormat.AsciiFont ?? currentFormat.AsciiFont,
            EastAsiaFont = templateFormat.EastAsiaFont ?? currentFormat.EastAsiaFont,
            Bold = templateFormat.Bold || currentFormat.Bold,
            Italic = templateFormat.Italic || currentFormat.Italic,
            Underline = templateFormat.Underline ?? currentFormat.Underline,
            FontColor = templateFormat.FontColor ?? currentFormat.FontColor,
            Highlight = templateFormat.Highlight ?? currentFormat.Highlight,
            VerticalAlign = templateFormat.VerticalAlign ?? currentFormat.VerticalAlign
        };
    }

    private static ParagraphFormatProfile BuildEquationParagraphFormat(
        Dictionary<string, TemplateStyleProfile?> styleProfileMap)
    {
        if (styleProfileMap.TryGetValue("template.equation", out var profile) && profile != null)
            return CloneParagraphFormat(profile.Paragraph);

        return new ParagraphFormatProfile
        {
            Align = "center",
            BeforeSpacing = "160",
            AfterSpacing = "160"
        };
    }

    private static RunFormatProfile BuildEquationRunFormat(
        Dictionary<string, TemplateStyleProfile?> styleProfileMap)
    {
        if (styleProfileMap.TryGetValue("template.equation", out var profile) && profile != null)
            return CloneRunFormat(profile.Run);

        return new RunFormatProfile
        {
            EastAsiaFont = "宋体",
            AsciiFont = "Times New Roman",
            FontSize = "24"
        };
    }

    private static bool MatchRunDecision(InlineRunSummary run, RunDecision decision)
    {
        if (!string.IsNullOrWhiteSpace(decision.RunPath)
            && string.Equals(run.Path, decision.RunPath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(decision.MatchText)
            && string.Equals(run.Text, decision.MatchText, StringComparison.Ordinal))
        {
            return true;
        }

        return string.IsNullOrWhiteSpace(decision.RunPath) && string.IsNullOrWhiteSpace(decision.MatchText);
    }

    private static bool MatchSegmentDecision(InlineTextSegment segment, RunDecision decision)
    {
        if (!string.IsNullOrWhiteSpace(decision.SegmentPath)
            && string.Equals(segment.Path, decision.SegmentPath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(decision.SegmentText)
            && string.Equals(segment.Text, decision.SegmentText, StringComparison.Ordinal))
        {
            return true;
        }

        return string.IsNullOrWhiteSpace(decision.SegmentPath) && string.IsNullOrWhiteSpace(decision.SegmentText);
    }

    private static void ApplySegmentDecision(List<InlineTextSegment> segments, RunDecision decision)
    {
        var matchedSegments = segments
            .Where(segment => MatchSegmentDecision(segment, decision))
            .ToList();
        if (decision.Occurrence.HasValue && decision.Occurrence.Value > 0)
            matchedSegments = matchedSegments.Skip(decision.Occurrence.Value - 1).Take(1).ToList();

        foreach (var segment in matchedSegments)
        {
            if (decision.ReplaceSegmentText != null)
                segment.Text = decision.ReplaceSegmentText;
            segment.Format = ApplyRunDecision(segment.Format, decision);
        }
    }

    private static RunFormatProfile ApplyRunDecision(RunFormatProfile format, RunDecision decision)
    {
        return new RunFormatProfile
        {
            FontSize = decision.FontSize ?? format.FontSize,
            AsciiFont = decision.AsciiFont ?? format.AsciiFont,
            EastAsiaFont = decision.EastAsiaFont ?? format.EastAsiaFont,
            Bold = decision.Bold ?? format.Bold,
            Italic = decision.Italic ?? format.Italic,
            Underline = decision.Underline ?? format.Underline,
            FontColor = decision.FontColor ?? format.FontColor,
            Highlight = decision.Highlight ?? format.Highlight,
            VerticalAlign = decision.VerticalAlign ?? format.VerticalAlign
        };
    }
}
