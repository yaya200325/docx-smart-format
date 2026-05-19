using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxWeb;

internal static class DocxAnalyzer
{
    public static AnalysisReport Analyze(string inputPath, string? assetRoot = null)
    {
        using var document = WordprocessingDocument.Open(inputPath, false);
        var mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("文档缺少 MainDocumentPart。");
        var styleResolver = new WordEffectiveStyleResolver(mainPart);
        var mainDocument = mainPart.Document
            ?? throw new InvalidOperationException("文档缺少主文档节点。");
        var body = mainDocument.Body
            ?? throw new InvalidOperationException("文档缺少 Body。");

        var blocks = new List<BodyBlockSummary>();
        var sections = new List<SectionSummary>();
        var assets = new List<ExtractedAsset>();
        var assetMap = new Dictionary<string, ExtractedAsset>(StringComparer.OrdinalIgnoreCase);
        var currentSectionIndex = 1;
        var paragraphIndex = 0;
        var tableIndex = 0;
        var imageIndex = 0;

        foreach (var element in body.Elements())
        {
            switch (element)
            {
                case Paragraph paragraph:
                    paragraphIndex++;

                    var images = ExtractParagraphImages(
                        paragraph,
                        mainPart,
                        styleResolver,
                        inputPath,
                        assetRoot,
                        assetMap,
                        assets,
                        ref imageIndex,
                        currentSectionIndex);
                    var equations = ExtractParagraphEquations(
                        paragraph,
                        ref paragraphIndex,
                        currentSectionIndex);

                    var paragraphWithoutImages = BuildParagraphBlock(paragraph, paragraphIndex, currentSectionIndex, styleResolver);
                    if ((!string.IsNullOrWhiteSpace(paragraphWithoutImages.Text) || (!images.Any() && !equations.Any()))
                        && !IsEquationOnlyParagraph(paragraphWithoutImages, equations))
                        blocks.Add(paragraphWithoutImages);

                    blocks.AddRange(images);
                    blocks.AddRange(equations);

                    if (paragraphWithoutImages.ContainsSectionBreak && paragraph.ParagraphProperties?.SectionProperties != null)
                    {
                        sections.Add(BuildSectionSummary(
                            paragraph.ParagraphProperties.SectionProperties,
                            currentSectionIndex));
                        currentSectionIndex++;
                    }
                    break;
                case Table table:
                    tableIndex++;
                    blocks.Add(BuildTableBlock(table, tableIndex, currentSectionIndex));
                    break;
            }
        }

        var finalSectionProperties = body.GetFirstChild<SectionProperties>();
        if (finalSectionProperties != null)
        {
            sections.Add(BuildSectionSummary(finalSectionProperties, currentSectionIndex));
        }
        else if (sections.Count == 0)
        {
            sections.Add(new SectionSummary
            {
                Path = "/section[1]"
            });
        }

        var headers = ReadHeaderFooters(mainPart, sections, isHeader: true);
        var footers = ReadHeaderFooters(mainPart, sections, isHeader: false);
        var firstSection = sections.FirstOrDefault();

        return new AnalysisReport
        {
            SourceFile = Path.GetFullPath(inputPath),
            Document = new DocumentSummary
            {
                Title = document.PackageProperties.Title,
                Author = document.PackageProperties.Creator,
                PageWidth = firstSection?.PageWidth,
                PageHeight = firstSection?.PageHeight,
                Orientation = firstSection?.Orientation,
                MarginTop = firstSection?.MarginTop,
                MarginBottom = firstSection?.MarginBottom,
                MarginLeft = firstSection?.MarginLeft,
                MarginRight = firstSection?.MarginRight,
                SectionCount = sections.Count,
                HeaderCount = headers.Count,
                FooterCount = footers.Count
            },
            Sections = sections,
            Headers = headers,
            Footers = footers,
            Body = blocks,
            Assets = assets,
            Issues = []
        };
    }

    private static BodyBlockSummary BuildParagraphBlock(Paragraph paragraph, int paragraphIndex, int sectionIndex, WordEffectiveStyleResolver styleResolver)
    {
        var properties = paragraph.ParagraphProperties;
        var style = properties?.ParagraphStyleId?.Val?.Value;
        var headingLevel = DetectHeadingLevel(style, properties?.OutlineLevel?.Val?.Value);
        var hasPageBreak = properties?.PageBreakBefore != null
            || paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);
        var hasSectionBreak = properties?.SectionProperties != null;
        var text = GetParagraphTextWithoutImages(paragraph);
        var paragraphFormat = styleResolver.ResolveParagraphFormat(paragraph, hasPageBreak);
        var runs = styleResolver.ResolveInlineRuns(paragraph, paragraphIndex);
        var dominantRunFormat = DetermineDominantRunFormat(runs);

        return new BodyBlockSummary
        {
            Path = $"/body/paragraph[{paragraphIndex}]",
            Kind = "paragraph",
            Text = text,
            Style = style,
            HeadingLevel = headingLevel,
            SectionIndex = sectionIndex,
            PageBreakBefore = hasPageBreak,
            ContainsSectionBreak = hasSectionBreak,
            OutlineRole = headingLevel.HasValue ? $"heading{headingLevel.Value}" : "paragraph",
            Paragraph = paragraphFormat,
            Run = dominantRunFormat,
            Runs = runs,
            HasImage = paragraph.Descendants().Any(e => e.LocalName is "drawing" or "pict"),
            HasTable = false,
            HasEquation = paragraph.Descendants().Any(IsMathElement)
        };
    }

    private static bool IsEquationOnlyParagraph(BodyBlockSummary paragraphBlock, List<BodyBlockSummary> equations)
    {
        return equations.Count > 0 && string.IsNullOrWhiteSpace(paragraphBlock.Text);
    }

    private static List<BodyBlockSummary> ExtractParagraphImages(
        Paragraph paragraph,
        MainDocumentPart mainPart,
        WordEffectiveStyleResolver styleResolver,
        string inputPath,
        string? assetRoot,
        Dictionary<string, ExtractedAsset> assetMap,
        List<ExtractedAsset> assets,
        ref int imageIndex,
        int sectionIndex)
    {
        var result = new List<BodyBlockSummary>();
        foreach (var drawing in paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
        {
            var blip = drawing.Descendants<Drawing.Blip>().FirstOrDefault();
            var relId = blip?.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relId))
                continue;

            if (mainPart.GetPartById(relId) is not ImagePart imagePart)
                continue;

            var asset = ExportImageAsset(inputPath, assetRoot, relId, imagePart, assetMap, assets);
            imageIndex++;

            var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();
            var nonVisual = drawing.Descendants<Pic.NonVisualPictureProperties>().FirstOrDefault();

            result.Add(new BodyBlockSummary
            {
                Path = $"/body/image[{imageIndex}]",
                Kind = "image",
                SectionIndex = sectionIndex,
                OutlineRole = "image",
                Paragraph = styleResolver.ResolveParagraphFormat(paragraph, hasPageBreak: false),
                HasImage = true,
                HasTable = false,
                Image = new ImageAssetSpec
                {
                    AssetId = asset.AssetId,
                    Path = asset.Path,
                    ContentType = asset.ContentType,
                    WidthEmu = extent?.Cx?.Value,
                    HeightEmu = extent?.Cy?.Value,
                    AltText = nonVisual?.NonVisualDrawingProperties?.Description?.Value
                }
            });
        }

        return result;
    }

    private static List<BodyBlockSummary> ExtractParagraphEquations(
        Paragraph paragraph,
        ref int paragraphIndex,
        int sectionIndex)
    {
        var result = new List<BodyBlockSummary>();
        var equationIndex = 0;

        foreach (var officeMathParagraph in paragraph.Elements().Where(x => x.LocalName == "oMathPara"))
        {
            equationIndex++;
            result.Add(BuildEquationBlock(
                officeMathParagraph,
                $"/body/equation[{paragraphIndex}.{equationIndex}]",
                sectionIndex,
                "display"));
        }

        foreach (var officeMath in paragraph.Descendants().Where(x => x.LocalName == "oMath"
                     && x.Ancestors().Any(a => a.LocalName == "oMathPara") == false))
        {
            equationIndex++;
            result.Add(BuildEquationBlock(
                officeMath,
                $"/body/equation[{paragraphIndex}.{equationIndex}]",
                sectionIndex,
                "inline"));
        }

        return result;
    }

    private static BodyBlockSummary BuildEquationBlock(OpenXmlElement mathElement, string path, int sectionIndex, string displayMode)
    {
        var text = NormalizeText(mathElement.InnerText);
        return new BodyBlockSummary
        {
            Path = path,
            Kind = "equation",
            Text = string.IsNullOrWhiteSpace(text) ? null : text,
            SectionIndex = sectionIndex,
            OutlineRole = "equation",
            Paragraph = new ParagraphFormatProfile(),
            Run = new RunFormatProfile(),
            HasImage = false,
            HasTable = false,
            HasEquation = true,
            Equation = new EquationSpec
            {
                Text = string.IsNullOrWhiteSpace(text) ? null : text,
                Xml = mathElement.OuterXml,
                DisplayMode = displayMode
            }
        };
    }

    private static bool IsMathElement(OpenXmlElement element)
    {
        return element.LocalName is "oMath" or "oMathPara";
    }

    private static ExtractedAsset ExportImageAsset(
        string inputPath,
        string? assetRoot,
        string relId,
        ImagePart imagePart,
        Dictionary<string, ExtractedAsset> assetMap,
        List<ExtractedAsset> assets)
    {
        var assetKey = $"{inputPath}:{relId}";
        if (assetMap.TryGetValue(assetKey, out var existing))
            return existing;

        var extension = GetExtensionFromContentType(imagePart.ContentType) ?? Path.GetExtension(imagePart.Uri.ToString());
        if (string.IsNullOrWhiteSpace(extension))
            extension = ".bin";

        var assetId = CreateStableAssetId(assetKey);
        var outputPath = assetRoot == null
            ? imagePart.Uri.ToString()
            : Path.Combine(assetRoot, assetId + extension);

        if (assetRoot != null)
        {
            Directory.CreateDirectory(assetRoot);
            using var input = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var output = File.Create(outputPath);
            input.CopyTo(output);
        }

        var asset = new ExtractedAsset
        {
            AssetId = assetId,
            Kind = "image",
            Path = outputPath,
            ContentType = imagePart.ContentType,
            SourcePath = imagePart.Uri.ToString()
        };

        assetMap[assetKey] = asset;
        assets.Add(asset);
        return asset;
    }

    private static BodyBlockSummary BuildTableBlock(Table table, int tableIndex, int sectionIndex)
    {
        var rows = table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>()
                .Select(cell => NormalizeText(cell.InnerText))
                .ToList())
            .ToList();

        return new BodyBlockSummary
        {
            Path = $"/body/table[{tableIndex}]",
            Kind = "table",
            Text = string.Join(" | ", rows.SelectMany(r => r)),
            SectionIndex = sectionIndex,
            OutlineRole = "table",
            Paragraph = new ParagraphFormatProfile(),
            HasImage = table.Descendants().Any(e => e.LocalName is "drawing" or "pict"),
            HasTable = true,
            TableRows = rows
        };
    }

    private static SectionSummary BuildSectionSummary(SectionProperties sectionProperties, int index)
    {
        var pageSize = sectionProperties.GetFirstChild<PageSize>();
        var pageMargin = sectionProperties.GetFirstChild<PageMargin>();
        var pageNumber = sectionProperties.GetFirstChild<PageNumberType>();
        var sectionType = sectionProperties.GetFirstChild<SectionType>();

        return new SectionSummary
        {
            Path = $"/section[{index}]",
            Type = sectionType?.Val?.InnerText,
            PageWidth = pageSize?.Width?.Value.ToString(),
            PageHeight = pageSize?.Height?.Value.ToString(),
            Orientation = pageSize?.Orient?.InnerText,
            PageStart = pageNumber?.Start?.Value is { } start ? (int)start : null,
            PageNumFmt = pageNumber?.Format?.InnerText,
            HeaderRef = JoinReferences(sectionProperties.Elements<HeaderReference>().Select(ToReferenceString)),
            FooterRef = JoinReferences(sectionProperties.Elements<FooterReference>().Select(ToReferenceString)),
            HeaderSourceRefs = JoinReferences(sectionProperties.Elements<HeaderReference>().Select(ToSourceReferenceString)),
            FooterSourceRefs = JoinReferences(sectionProperties.Elements<FooterReference>().Select(ToSourceReferenceString)),
            TitlePage = sectionProperties.GetFirstChild<TitlePage>() != null,
            MarginTop = pageMargin?.Top?.Value.ToString(),
            MarginBottom = pageMargin?.Bottom?.Value.ToString(),
            MarginLeft = pageMargin?.Left?.Value.ToString(),
            MarginRight = pageMargin?.Right?.Value.ToString()
        };
    }

    private static List<HeaderFooterSummary> ReadHeaderFooters(
        MainDocumentPart mainPart,
        IEnumerable<SectionSummary> sections,
        bool isHeader)
    {
        var result = new List<HeaderFooterSummary>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var section in sections)
        {
            var refs = (isHeader ? section.HeaderRef : section.FooterRef)?
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                ?? [];

            foreach (var reference in refs)
            {
                var parts = reference.Split(':', 2);
                if (parts.Length != 2 || !seen.Add(reference))
                    continue;

                var type = parts[0];
                var relId = parts[1];
                var part = mainPart.GetPartById(relId);

                if (part is HeaderPart headerPart && isHeader && headerPart.Header != null)
                {
                    result.Add(BuildHeaderFooterSummary(headerPart.Header, "header", type, headerPart.Uri.ToString(), relId));
                }
                else if (part is FooterPart footerPart && !isHeader && footerPart.Footer != null)
                {
                    result.Add(BuildHeaderFooterSummary(footerPart.Footer, "footer", type, footerPart.Uri.ToString(), relId));
                }
            }
        }

        return result;
    }

    private static HeaderFooterSummary BuildHeaderFooterSummary(
        OpenXmlCompositeElement root,
        string kind,
        string type,
        string uri,
        string relationshipId)
    {
        var paragraphs = root.Descendants<Paragraph>()
            .Select((paragraph, index) => new HeaderFooterParagraph
            {
                Path = $"/{kind}/{type}/paragraph[{index + 1}]",
                Text = GetParagraphTextWithoutImages(paragraph),
                Style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
                Align = paragraph.ParagraphProperties?.Justification?.Val?.InnerText
            })
            .ToList();

        var xml = root.OuterXml;
        return new HeaderFooterSummary
        {
            Path = uri,
            RelationshipId = relationshipId,
            Kind = kind,
            Type = type,
            Paragraphs = paragraphs,
            ContainsPageField = xml.Contains("PAGE", StringComparison.OrdinalIgnoreCase),
            ContainsNumPagesField = xml.Contains("NUMPAGES", StringComparison.OrdinalIgnoreCase)
        };
    }

    private static string JoinReferences(IEnumerable<string> refs)
    {
        return string.Join(";", refs.Where(static item => !string.IsNullOrWhiteSpace(item)));
    }

    private static string ToReferenceString(HeaderReference headerReference)
    {
        return $"{headerReference.Type?.InnerText?.ToLowerInvariant() ?? "default"}:{headerReference.Id?.Value}";
    }

    private static string ToReferenceString(FooterReference footerReference)
    {
        return $"{footerReference.Type?.InnerText?.ToLowerInvariant() ?? "default"}:{footerReference.Id?.Value}";
    }

    private static string ToSourceReferenceString(HeaderReference headerReference)
    {
        return $"{headerReference.Type?.InnerText?.ToLowerInvariant() ?? "default"}:{headerReference.Id?.Value}";
    }

    private static string ToSourceReferenceString(FooterReference footerReference)
    {
        return $"{footerReference.Type?.InnerText?.ToLowerInvariant() ?? "default"}:{footerReference.Id?.Value}";
    }

    private static int? DetectHeadingLevel(string? style, int? outlineLevel)
    {
        if (!string.IsNullOrWhiteSpace(style))
        {
            var normalized = style.Replace(" ", "", StringComparison.OrdinalIgnoreCase).ToLowerInvariant();
            if (normalized.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(normalized["heading".Length..], out var parsed))
            {
                return parsed;
            }
        }

        if (outlineLevel.HasValue)
            return outlineLevel.Value + 1;

        return null;
    }

    private static RunFormatProfile DetermineDominantRunFormat(List<InlineRunSummary> runs)
    {
        if (runs.Count == 0)
            return new RunFormatProfile();

        var dominant = runs
            .GroupBy(run => BuildFormatKey(run.Format))
            .OrderByDescending(group => group.Count())
            .ThenByDescending(group => group.Sum(item => item.Text?.Length ?? 0))
            .First()
            .First()
            .Format;

        return CloneRunFormat(dominant);
    }

    private static string BuildFormatKey(RunFormatProfile format)
    {
        return string.Join("|", [
            format.FontSize ?? "",
            format.AsciiFont ?? "",
            format.EastAsiaFont ?? "",
            format.Bold ? "1" : "0",
            format.Italic ? "1" : "0",
            format.Underline ?? "",
            format.FontColor ?? "",
            format.Highlight ?? ""
        ]);
    }

    private static RunFormatProfile CloneRunFormat(RunFormatProfile format)
    {
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

    private static string? GetParagraphTextWithoutImages(Paragraph paragraph)
    {
        var texts = paragraph.Descendants<Text>()
            .Where(text => !text.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any())
            .Select(text => text.Text);
        var value = string.Concat(texts);
        value = NormalizeText(value);
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string NormalizeText(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        return value.Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Trim();
    }

    private static string CreateStableAssetId(string value)
    {
        var bytes = SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(value));
        return Convert.ToHexString(bytes[..8]).ToLowerInvariant();
    }

    private static string? GetExtensionFromContentType(string? contentType)
    {
        return contentType?.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tif",
            "image/x-emf" => ".emf",
            "image/x-wmf" => ".wmf",
            "image/svg+xml" => ".svg",
            _ => null
        };
    }
}
