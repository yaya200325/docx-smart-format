using System.Text.Json.Serialization;

namespace DocxWeb;

public sealed class AnalysisReport
{
    [JsonPropertyName("sourceFile")]
    public string SourceFile { get; set; } = "";

    [JsonPropertyName("document")]
    public DocumentSummary Document { get; set; } = new();

    [JsonPropertyName("sections")]
    public List<SectionSummary> Sections { get; set; } = [];

    [JsonPropertyName("headers")]
    public List<HeaderFooterSummary> Headers { get; set; } = [];

    [JsonPropertyName("footers")]
    public List<HeaderFooterSummary> Footers { get; set; } = [];

    [JsonPropertyName("body")]
    public List<BodyBlockSummary> Body { get; set; } = [];

    [JsonPropertyName("assets")]
    public List<ExtractedAsset> Assets { get; set; } = [];

    [JsonPropertyName("issues")]
    public List<AnalysisIssue> Issues { get; set; } = [];
}

public sealed class DocumentSummary
{
    [JsonPropertyName("title")]
    public string? Title { get; set; }

    [JsonPropertyName("author")]
    public string? Author { get; set; }

    [JsonPropertyName("pageWidth")]
    public string? PageWidth { get; set; }

    [JsonPropertyName("pageHeight")]
    public string? PageHeight { get; set; }

    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }

    [JsonPropertyName("marginTop")]
    public string? MarginTop { get; set; }

    [JsonPropertyName("marginBottom")]
    public string? MarginBottom { get; set; }

    [JsonPropertyName("marginLeft")]
    public string? MarginLeft { get; set; }

    [JsonPropertyName("marginRight")]
    public string? MarginRight { get; set; }

    [JsonPropertyName("sectionCount")]
    public int SectionCount { get; set; }

    [JsonPropertyName("headerCount")]
    public int HeaderCount { get; set; }

    [JsonPropertyName("footerCount")]
    public int FooterCount { get; set; }
}

public sealed class SectionSummary
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("pageWidth")]
    public string? PageWidth { get; set; }

    [JsonPropertyName("pageHeight")]
    public string? PageHeight { get; set; }

    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }

    [JsonPropertyName("pageStart")]
    public int? PageStart { get; set; }

    [JsonPropertyName("pageNumFmt")]
    public string? PageNumFmt { get; set; }

    [JsonPropertyName("headerRef")]
    public string? HeaderRef { get; set; }

    [JsonPropertyName("footerRef")]
    public string? FooterRef { get; set; }

    [JsonPropertyName("headerSourceRefs")]
    public string? HeaderSourceRefs { get; set; }

    [JsonPropertyName("footerSourceRefs")]
    public string? FooterSourceRefs { get; set; }

    [JsonPropertyName("titlePage")]
    public bool TitlePage { get; set; }

    [JsonPropertyName("marginTop")]
    public string? MarginTop { get; set; }

    [JsonPropertyName("marginBottom")]
    public string? MarginBottom { get; set; }

    [JsonPropertyName("marginLeft")]
    public string? MarginLeft { get; set; }

    [JsonPropertyName("marginRight")]
    public string? MarginRight { get; set; }
}

public sealed class HeaderFooterSummary
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("relationshipId")]
    public string? RelationshipId { get; set; }

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "default";

    [JsonPropertyName("paragraphs")]
    public List<HeaderFooterParagraph> Paragraphs { get; set; } = [];

    [JsonPropertyName("containsPageField")]
    public bool ContainsPageField { get; set; }

    [JsonPropertyName("containsNumPagesField")]
    public bool ContainsNumPagesField { get; set; }
}

public sealed class HeaderFooterParagraph
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("align")]
    public string? Align { get; set; }

    [JsonPropertyName("run")]
    public RunFormatProfile Run { get; set; } = new();
}

public sealed class BodyBlockSummary
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "";

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("headingLevel")]
    public int? HeadingLevel { get; set; }

    [JsonPropertyName("sectionIndex")]
    public int SectionIndex { get; set; }

    [JsonPropertyName("pageBreakBefore")]
    public bool PageBreakBefore { get; set; }

    [JsonPropertyName("containsSectionBreak")]
    public bool ContainsSectionBreak { get; set; }

    [JsonPropertyName("outlineRole")]
    public string? OutlineRole { get; set; }

    [JsonPropertyName("paragraph")]
    public ParagraphFormatProfile Paragraph { get; set; } = new();

    [JsonPropertyName("run")]
    public RunFormatProfile Run { get; set; } = new();

    [JsonPropertyName("runs")]
    public List<InlineRunSummary> Runs { get; set; } = [];

    [JsonPropertyName("hasImage")]
    public bool HasImage { get; set; }

    [JsonPropertyName("hasTable")]
    public bool HasTable { get; set; }

    [JsonPropertyName("hasEquation")]
    public bool HasEquation { get; set; }

    [JsonPropertyName("tableRows")]
    public List<List<string>> TableRows { get; set; } = [];

    [JsonPropertyName("image")]
    public ImageAssetSpec? Image { get; set; }

    [JsonPropertyName("equation")]
    public EquationSpec? Equation { get; set; }
}

public sealed class AnalysisIssue
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("description")]
    public string Description { get; set; } = "";

    [JsonPropertyName("path")]
    public string? Path { get; set; }
}

public sealed class RenderSpec
{
    [JsonPropertyName("document")]
    public RenderDocumentSpec Document { get; set; } = new();

    [JsonPropertyName("headers")]
    public List<RenderHeaderFooterSpec> Headers { get; set; } = [];

    [JsonPropertyName("footers")]
    public List<RenderHeaderFooterSpec> Footers { get; set; } = [];

    [JsonPropertyName("body")]
    public List<RenderBodyBlock> Body { get; set; } = [];
}

public sealed class RenderDocumentSpec
{
    [JsonPropertyName("title")]
    public string? Title { get; set; }

    [JsonPropertyName("author")]
    public string? Author { get; set; }

    [JsonPropertyName("locale")]
    public string? Locale { get; set; }

    [JsonPropertyName("pageWidth")]
    public string? PageWidth { get; set; }

    [JsonPropertyName("pageHeight")]
    public string? PageHeight { get; set; }

    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }

    [JsonPropertyName("marginTop")]
    public string? MarginTop { get; set; }

    [JsonPropertyName("marginBottom")]
    public string? MarginBottom { get; set; }

    [JsonPropertyName("marginLeft")]
    public string? MarginLeft { get; set; }

    [JsonPropertyName("marginRight")]
    public string? MarginRight { get; set; }
}

public sealed class RenderHeaderFooterSpec
{
    [JsonPropertyName("sourcePath")]
    public string? SourcePath { get; set; }

    [JsonPropertyName("relationshipId")]
    public string? RelationshipId { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "default";

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "";

    [JsonPropertyName("paragraphs")]
    public List<RenderParagraphSpec> Paragraphs { get; set; } = [];

    [JsonPropertyName("textBoxes")]
    public List<RenderTextBoxSpec> TextBoxes { get; set; } = [];
}

public sealed class RenderTextBoxSpec
{
    [JsonPropertyName("posXEmu")]
    public long? PosXEmu { get; set; }

    [JsonPropertyName("posYEmu")]
    public long? PosYEmu { get; set; }

    [JsonPropertyName("widthEmu")]
    public long? WidthEmu { get; set; }

    [JsonPropertyName("heightEmu")]
    public long? HeightEmu { get; set; }

    [JsonPropertyName("relativeFromH")]
    public string? RelativeFromH { get; set; }

    [JsonPropertyName("relativeFromV")]
    public string? RelativeFromV { get; set; }

    [JsonPropertyName("textDirection")]
    public string? TextDirection { get; set; }

    [JsonPropertyName("borderStyle")]
    public string? BorderStyle { get; set; }

    [JsonPropertyName("fillColor")]
    public string? FillColor { get; set; }

    [JsonPropertyName("anchor")]
    public string? Anchor { get; set; }

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("paragraphs")]
    public List<RenderParagraphSpec> Paragraphs { get; set; } = [];
}

public sealed class RenderBodyBlock
{
    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "paragraph";

    [JsonPropertyName("sourceBlockId")]
    public string? SourceBlockId { get; set; }

    [JsonPropertyName("paragraph")]
    public RenderParagraphSpec? Paragraph { get; set; }

    [JsonPropertyName("section")]
    public RenderSectionSpec? Section { get; set; }

    [JsonPropertyName("table")]
    public RenderTableSpec? Table { get; set; }

    [JsonPropertyName("image")]
    public RenderImageSpec? Image { get; set; }

    [JsonPropertyName("equation")]
    public RenderEquationSpec? Equation { get; set; }
}

public sealed class RenderParagraphSpec
{
    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("pageNumberPrefix")]
    public string? PageNumberPrefix { get; set; }

    [JsonPropertyName("pageNumberSuffix")]
    public string? PageNumberSuffix { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("headingLevel")]
    public int? HeadingLevel { get; set; }

    [JsonPropertyName("paragraph")]
    public ParagraphFormatProfile Paragraph { get; set; } = new();

    [JsonPropertyName("run")]
    public RunFormatProfile Run { get; set; } = new();

    [JsonPropertyName("runs")]
    public List<InlineRunSummary> Runs { get; set; } = [];

    [JsonPropertyName("align")]
    public string? Align { get; set; }

    [JsonPropertyName("pageBreakBefore")]
    public bool PageBreakBefore { get; set; }

    [JsonPropertyName("pageNumberField")]
    public bool PageNumberField { get; set; }

    [JsonPropertyName("numPagesField")]
    public bool NumPagesField { get; set; }
}

public sealed class RenderSectionSpec
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("pageWidth")]
    public string? PageWidth { get; set; }

    [JsonPropertyName("pageHeight")]
    public string? PageHeight { get; set; }

    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }

    [JsonPropertyName("marginTop")]
    public string? MarginTop { get; set; }

    [JsonPropertyName("marginBottom")]
    public string? MarginBottom { get; set; }

    [JsonPropertyName("marginLeft")]
    public string? MarginLeft { get; set; }

    [JsonPropertyName("marginRight")]
    public string? MarginRight { get; set; }

    [JsonPropertyName("pageStart")]
    public int? PageStart { get; set; }

    [JsonPropertyName("pageNumFmt")]
    public string? PageNumFmt { get; set; }

    [JsonPropertyName("titlePage")]
    public bool TitlePage { get; set; }

    [JsonPropertyName("headerType")]
    public string? HeaderType { get; set; }

    [JsonPropertyName("footerType")]
    public string? FooterType { get; set; }

    [JsonPropertyName("headerSourcePath")]
    public string? HeaderSourcePath { get; set; }

    [JsonPropertyName("footerSourcePath")]
    public string? FooterSourcePath { get; set; }

    [JsonPropertyName("headers")]
    public List<RenderHeaderFooterReferenceSpec> Headers { get; set; } = [];

    [JsonPropertyName("footers")]
    public List<RenderHeaderFooterReferenceSpec> Footers { get; set; } = [];
}

public sealed class RenderHeaderFooterReferenceSpec
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "default";

    [JsonPropertyName("sourcePath")]
    public string? SourcePath { get; set; }
}

public sealed class RenderTableSpec
{
    [JsonPropertyName("rows")]
    public List<List<string>> Rows { get; set; } = [];
}

public sealed class EquationSpec
{
    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("xml")]
    public string Xml { get; set; } = "";

    [JsonPropertyName("displayMode")]
    public string DisplayMode { get; set; } = "inline";
}

public sealed class RenderEquationSpec
{
    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("xml")]
    public string Xml { get; set; } = "";

    [JsonPropertyName("displayMode")]
    public string DisplayMode { get; set; } = "inline";

    [JsonPropertyName("paragraph")]
    public ParagraphFormatProfile Paragraph { get; set; } = new();

    [JsonPropertyName("run")]
    public RunFormatProfile Run { get; set; } = new();
}

public sealed class TemplateAnalysis
{
    [JsonPropertyName("templateDocId")]
    public string TemplateDocId { get; set; } = "";

    [JsonPropertyName("document")]
    public DocumentSummary Document { get; set; } = new();

    [JsonPropertyName("styles")]
    public TemplateStyleCatalog Styles { get; set; } = new();

    [JsonPropertyName("headers")]
    public List<HeaderFooterSummary> Headers { get; set; } = [];

    [JsonPropertyName("footers")]
    public List<HeaderFooterSummary> Footers { get; set; } = [];

    [JsonPropertyName("sections")]
    public List<SectionSummary> Sections { get; set; } = [];

    [JsonPropertyName("rules")]
    public TemplateRuleSet Rules { get; set; } = new();
}

public sealed class TemplateOverview
{
    [JsonPropertyName("styles")]
    public List<TemplateStyleOverview> Styles { get; set; } = [];

    [JsonPropertyName("heading1")]
    public TemplateStyleOverview? Heading1 { get; set; }

    [JsonPropertyName("heading2")]
    public TemplateStyleOverview? Heading2 { get; set; }

    [JsonPropertyName("heading3")]
    public TemplateStyleOverview? Heading3 { get; set; }

    [JsonPropertyName("body")]
    public TemplateStyleOverview? Body { get; set; }

    [JsonPropertyName("figureCaption")]
    public TemplateStyleOverview? FigureCaption { get; set; }

    [JsonPropertyName("tableCaption")]
    public TemplateStyleOverview? TableCaption { get; set; }

    [JsonPropertyName("equation")]
    public TemplateStyleOverview? Equation { get; set; }

    [JsonPropertyName("reference")]
    public TemplateStyleOverview? Reference { get; set; }

    [JsonPropertyName("sections")]
    public List<TemplateSectionOverview> Sections { get; set; } = [];

    [JsonPropertyName("headerFooter")]
    public TemplateHeaderFooterOverview HeaderFooter { get; set; } = new();
}

public sealed class TemplateHeaderFooterOverview
{
    [JsonPropertyName("hasHeader")]
    public bool HasHeader { get; set; }

    [JsonPropertyName("hasFooter")]
    public bool HasFooter { get; set; }

    [JsonPropertyName("hasOddEvenHeader")]
    public bool HasOddEvenHeader { get; set; }

    [JsonPropertyName("hasOddEvenFooter")]
    public bool HasOddEvenFooter { get; set; }

    [JsonPropertyName("hasPageNumber")]
    public bool HasPageNumber { get; set; }

    [JsonPropertyName("pageNumberScheme")]
    public string? PageNumberScheme { get; set; }

    [JsonPropertyName("hasTitlePage")]
    public bool HasTitlePage { get; set; }
}

public sealed class TemplateStyleOverview
{
    [JsonPropertyName("key")]
    public string Key { get; set; } = "";

    [JsonPropertyName("styleName")]
    public string? StyleName { get; set; }

    [JsonPropertyName("asciiFont")]
    public string? AsciiFont { get; set; }

    [JsonPropertyName("eastAsiaFont")]
    public string? EastAsiaFont { get; set; }

    [JsonPropertyName("fontFamily")]
    public string? FontFamily { get; set; }

    [JsonPropertyName("fontSize")]
    public string? FontSize { get; set; }

    [JsonPropertyName("bold")]
    public bool Bold { get; set; }

    [JsonPropertyName("italic")]
    public bool Italic { get; set; }

    [JsonPropertyName("pageBreakBefore")]
    public bool PageBreakBefore { get; set; }

    [JsonPropertyName("align")]
    public string? Align { get; set; }

    [JsonPropertyName("leftIndent")]
    public string? LeftIndent { get; set; }

    [JsonPropertyName("rightIndent")]
    public string? RightIndent { get; set; }

    [JsonPropertyName("firstLineIndent")]
    public string? FirstLineIndent { get; set; }

    [JsonPropertyName("firstLineChars")]
    public string? FirstLineChars { get; set; }

    [JsonPropertyName("hangingIndent")]
    public string? HangingIndent { get; set; }

    [JsonPropertyName("hangingChars")]
    public string? HangingChars { get; set; }

    [JsonPropertyName("beforeSpacing")]
    public string? BeforeSpacing { get; set; }

    [JsonPropertyName("afterSpacing")]
    public string? AfterSpacing { get; set; }

    [JsonPropertyName("lineSpacing")]
    public string? LineSpacing { get; set; }

    [JsonPropertyName("underline")]
    public string? Underline { get; set; }

    [JsonPropertyName("fontColor")]
    public string? FontColor { get; set; }

    [JsonPropertyName("verticalAlign")]
    public string? VerticalAlign { get; set; }
}

public sealed class TemplateSectionOverview
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("titlePage")]
    public bool TitlePage { get; set; }

    [JsonPropertyName("pageStart")]
    public int? PageStart { get; set; }

    [JsonPropertyName("pageNumFmt")]
    public string? PageNumFmt { get; set; }

    [JsonPropertyName("headerType")]
    public string? HeaderType { get; set; }

    [JsonPropertyName("footerType")]
    public string? FooterType { get; set; }

    [JsonPropertyName("headerText")]
    public string? HeaderText { get; set; }

    [JsonPropertyName("evenHeaderText")]
    public string? EvenHeaderText { get; set; }

    [JsonPropertyName("footerText")]
    public string? FooterText { get; set; }

    [JsonPropertyName("headerSourcePath")]
    public string? HeaderSourcePath { get; set; }

    [JsonPropertyName("footerSourcePath")]
    public string? FooterSourcePath { get; set; }

    [JsonPropertyName("hasEvenHeader")]
    public bool HasEvenHeader { get; set; }

    [JsonPropertyName("hasEvenFooter")]
    public bool HasEvenFooter { get; set; }
}

public sealed class TemplateStyleCatalog
{
    [JsonPropertyName("heading1")]
    public TemplateStyleProfile? Heading1 { get; set; }

    [JsonPropertyName("heading2")]
    public TemplateStyleProfile? Heading2 { get; set; }

    [JsonPropertyName("heading3")]
    public TemplateStyleProfile? Heading3 { get; set; }

    [JsonPropertyName("body")]
    public TemplateStyleProfile? Body { get; set; }

    [JsonPropertyName("figureCaption")]
    public TemplateStyleProfile? FigureCaption { get; set; }

    [JsonPropertyName("tableCaption")]
    public TemplateStyleProfile? TableCaption { get; set; }

    [JsonPropertyName("equation")]
    public TemplateStyleProfile? Equation { get; set; }

    [JsonPropertyName("reference")]
    public TemplateStyleProfile? Reference { get; set; }
}

public sealed class TemplateRuleSet
{
    [JsonPropertyName("documentRules")]
    public TemplateDocumentRules DocumentRules { get; set; } = new();

    [JsonPropertyName("headerFooterRules")]
    public HeaderFooterRuleSet HeaderFooterRules { get; set; } = new();

    [JsonPropertyName("styleProfiles")]
    public TemplateStyleCatalog StyleProfiles { get; set; } = new();

    [JsonPropertyName("captionRules")]
    public CaptionRuleSet CaptionRules { get; set; } = new();

    [JsonPropertyName("outlineRules")]
    public OutlineRuleSet OutlineRules { get; set; } = new();

    [JsonPropertyName("normalizedRequirements")]
    public List<string> NormalizedRequirements { get; set; } = [];
}

public sealed class TemplateDocumentRules
{
    [JsonPropertyName("pageWidth")]
    public string? PageWidth { get; set; }

    [JsonPropertyName("pageHeight")]
    public string? PageHeight { get; set; }

    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }

    [JsonPropertyName("marginTop")]
    public string? MarginTop { get; set; }

    [JsonPropertyName("marginBottom")]
    public string? MarginBottom { get; set; }

    [JsonPropertyName("marginLeft")]
    public string? MarginLeft { get; set; }

    [JsonPropertyName("marginRight")]
    public string? MarginRight { get; set; }

    [JsonPropertyName("sectionTypes")]
    public List<string> SectionTypes { get; set; } = [];

    [JsonPropertyName("pageNumberFormats")]
    public List<string> PageNumberFormats { get; set; } = [];
}

public sealed class HeaderFooterRuleSet
{
    [JsonPropertyName("hasHeader")]
    public bool HasHeader { get; set; }

    [JsonPropertyName("hasFooter")]
    public bool HasFooter { get; set; }

    [JsonPropertyName("headerTypes")]
    public List<string> HeaderTypes { get; set; } = [];

    [JsonPropertyName("footerTypes")]
    public List<string> FooterTypes { get; set; } = [];

    [JsonPropertyName("containsPageNumber")]
    public bool ContainsPageNumber { get; set; }

    [JsonPropertyName("containsTotalPages")]
    public bool ContainsTotalPages { get; set; }
}

public sealed class CaptionRuleSet
{
    [JsonPropertyName("figurePrefixes")]
    public List<string> FigurePrefixes { get; set; } = [];

    [JsonPropertyName("tablePrefixes")]
    public List<string> TablePrefixes { get; set; } = [];

    [JsonPropertyName("figureCaptionStyleKey")]
    public string? FigureCaptionStyleKey { get; set; }

    [JsonPropertyName("tableCaptionStyleKey")]
    public string? TableCaptionStyleKey { get; set; }

    [JsonPropertyName("figureCaptionPosition")]
    public string? FigureCaptionPosition { get; set; }

    [JsonPropertyName("tableCaptionPosition")]
    public string? TableCaptionPosition { get; set; }
}

public sealed class OutlineRuleSet
{
    [JsonPropertyName("headingStyleKeys")]
    public List<string> HeadingStyleKeys { get; set; } = [];

    [JsonPropertyName("maxHeadingLevel")]
    public int MaxHeadingLevel { get; set; }

    [JsonPropertyName("headingRecognitionHints")]
    public List<string> HeadingRecognitionHints { get; set; } = [];
}

public sealed class TemplateStyleProfile
{
    [JsonPropertyName("styleName")]
    public string? StyleName { get; set; }

    [JsonPropertyName("paragraph")]
    public ParagraphFormatProfile Paragraph { get; set; } = new();

    [JsonPropertyName("run")]
    public RunFormatProfile Run { get; set; } = new();
}

public sealed class ParagraphFormatProfile
{
    [JsonPropertyName("align")]
    public string? Align { get; set; }

    [JsonPropertyName("pageBreakBefore")]
    public bool PageBreakBefore { get; set; }

    [JsonPropertyName("leftIndent")]
    public string? LeftIndent { get; set; }

    [JsonPropertyName("rightIndent")]
    public string? RightIndent { get; set; }

    [JsonPropertyName("firstLineIndent")]
    public string? FirstLineIndent { get; set; }

    [JsonPropertyName("firstLineChars")]
    public string? FirstLineChars { get; set; }

    [JsonPropertyName("hangingIndent")]
    public string? HangingIndent { get; set; }

    [JsonPropertyName("hangingChars")]
    public string? HangingChars { get; set; }

    [JsonPropertyName("beforeSpacing")]
    public string? BeforeSpacing { get; set; }

    [JsonPropertyName("afterSpacing")]
    public string? AfterSpacing { get; set; }

    [JsonPropertyName("lineSpacing")]
    public string? LineSpacing { get; set; }

    [JsonPropertyName("lineSpacingRule")]
    public string? LineSpacingRule { get; set; }

    [JsonPropertyName("tabStops")]
    public List<TabStopSpec> TabStops { get; set; } = [];
}

public sealed class TabStopSpec
{
    [JsonPropertyName("position")]
    public string? Position { get; set; }

    [JsonPropertyName("alignment")]
    public string? Alignment { get; set; }
}

public sealed class RunFormatProfile
{
    [JsonPropertyName("fontSize")]
    public string? FontSize { get; set; }

    [JsonPropertyName("asciiFont")]
    public string? AsciiFont { get; set; }

    [JsonPropertyName("eastAsiaFont")]
    public string? EastAsiaFont { get; set; }

    [JsonPropertyName("bold")]
    public bool Bold { get; set; }

    [JsonPropertyName("italic")]
    public bool Italic { get; set; }

    [JsonPropertyName("underline")]
    public string? Underline { get; set; }

    [JsonPropertyName("fontColor")]
    public string? FontColor { get; set; }

    [JsonPropertyName("highlight")]
    public string? Highlight { get; set; }

    [JsonPropertyName("verticalAlign")]
    public string? VerticalAlign { get; set; }
}

public sealed class InlineRunSummary
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("format")]
    public RunFormatProfile Format { get; set; } = new();

    [JsonPropertyName("segments")]
    public List<InlineTextSegment> Segments { get; set; } = [];

    [JsonPropertyName("kind")]
    public string? Kind { get; set; }

    [JsonPropertyName("bookmarkName")]
    public string? BookmarkName { get; set; }

    [JsonPropertyName("bookmarkId")]
    public string? BookmarkId { get; set; }

    [JsonPropertyName("fieldInstruction")]
    public string? FieldInstruction { get; set; }

    [JsonPropertyName("fieldDisplayText")]
    public string? FieldDisplayText { get; set; }
}

public sealed class InlineTextSegment
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("format")]
    public RunFormatProfile Format { get; set; } = new();
}

public sealed class SourceDocumentAnalysis
{
    [JsonPropertyName("sourceDocId")]
    public string SourceDocId { get; set; } = "";

    [JsonPropertyName("document")]
    public DocumentSummary Document { get; set; } = new();

    [JsonPropertyName("blocks")]
    public List<SemanticBlock> Blocks { get; set; } = [];

    [JsonPropertyName("sections")]
    public List<SectionSummary> Sections { get; set; } = [];

    [JsonPropertyName("headers")]
    public List<HeaderFooterSummary> Headers { get; set; } = [];

    [JsonPropertyName("footers")]
    public List<HeaderFooterSummary> Footers { get; set; } = [];

    [JsonPropertyName("assets")]
    public List<ExtractedAsset> Assets { get; set; } = [];
}

public sealed class SemanticBlock
{
    [JsonPropertyName("blockId")]
    public string BlockId { get; set; } = "";

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "";

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("sectionIndex")]
    public int SectionIndex { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("headingLevel")]
    public int? HeadingLevel { get; set; }

    [JsonPropertyName("paragraph")]
    public ParagraphFormatProfile Paragraph { get; set; } = new();

    [JsonPropertyName("run")]
    public RunFormatProfile Run { get; set; } = new();

    [JsonPropertyName("runs")]
    public List<InlineRunSummary> Runs { get; set; } = [];

    [JsonPropertyName("image")]
    public ImageAssetSpec? Image { get; set; }

    [JsonPropertyName("table")]
    public RenderTableSpec? Table { get; set; }

    [JsonPropertyName("equation")]
    public EquationSpec? Equation { get; set; }

    [JsonPropertyName("semanticHints")]
    public SemanticHints Hints { get; set; } = new();
}

public sealed class SemanticHints
{
    [JsonPropertyName("isLikelyHeading")]
    public bool IsLikelyHeading { get; set; }

    [JsonPropertyName("isLikelyFigureCaption")]
    public bool IsLikelyFigureCaption { get; set; }

    [JsonPropertyName("isLikelyTableCaption")]
    public bool IsLikelyTableCaption { get; set; }

    [JsonPropertyName("isLikelyEquation")]
    public bool IsLikelyEquation { get; set; }
}

public sealed class FormatDecisionEnvelope
{
    [JsonPropertyName("documentRules")]
    public DecisionDocumentRules DocumentRules { get; set; } = new();

    [JsonPropertyName("blockDecisions")]
    public List<BlockDecision> BlockDecisions { get; set; } = [];

    [JsonPropertyName("sectionDecisions")]
    public List<SectionDecision> SectionDecisions { get; set; } = [];

    [JsonPropertyName("headerFooterDecisions")]
    public List<HeaderFooterDecision> HeaderFooterDecisions { get; set; } = [];

    [JsonPropertyName("captionDecisions")]
    public List<CaptionDecision> CaptionDecisions { get; set; } = [];

    [JsonPropertyName("runDecisions")]
    public List<RunDecision> RunDecisions { get; set; } = [];
}

public sealed class DecisionDocumentRules
{
    [JsonPropertyName("headerStyleKey")]
    public string? HeaderStyleKey { get; set; }

    [JsonPropertyName("footerStyleKey")]
    public string? FooterStyleKey { get; set; }

    [JsonPropertyName("pageNumberScheme")]
    public string? PageNumberScheme { get; set; }

    [JsonPropertyName("useTemplateHeaders")]
    public bool UseTemplateHeaders { get; set; } = true;

    [JsonPropertyName("useTemplateFooters")]
    public bool UseTemplateFooters { get; set; } = true;
}

public sealed class BlockDecision
{
    [JsonPropertyName("blockId")]
    public string BlockId { get; set; } = "";

    [JsonPropertyName("semanticRole")]
    public string SemanticRole { get; set; } = "";

    [JsonPropertyName("applyStyleKey")]
    public string ApplyStyleKey { get; set; } = "";

    [JsonPropertyName("align")]
    public string? Align { get; set; }

    [JsonPropertyName("pageBreakBefore")]
    public bool? PageBreakBefore { get; set; }

    [JsonPropertyName("leftIndent")]
    public string? LeftIndent { get; set; }

    [JsonPropertyName("rightIndent")]
    public string? RightIndent { get; set; }

    [JsonPropertyName("firstLineIndent")]
    public string? FirstLineIndent { get; set; }

    [JsonPropertyName("firstLineChars")]
    public string? FirstLineChars { get; set; }

    [JsonPropertyName("hangingIndent")]
    public string? HangingIndent { get; set; }

    [JsonPropertyName("hangingChars")]
    public string? HangingChars { get; set; }

    [JsonPropertyName("beforeSpacing")]
    public string? BeforeSpacing { get; set; }

    [JsonPropertyName("afterSpacing")]
    public string? AfterSpacing { get; set; }

    [JsonPropertyName("lineSpacing")]
    public string? LineSpacing { get; set; }

    [JsonPropertyName("lineSpacingRule")]
    public string? LineSpacingRule { get; set; }

    [JsonPropertyName("fontSize")]
    public string? FontSize { get; set; }

    [JsonPropertyName("asciiFont")]
    public string? AsciiFont { get; set; }

    [JsonPropertyName("eastAsiaFont")]
    public string? EastAsiaFont { get; set; }

    [JsonPropertyName("bold")]
    public bool? Bold { get; set; }

    [JsonPropertyName("italic")]
    public bool? Italic { get; set; }

    [JsonPropertyName("underline")]
    public string? Underline { get; set; }

    [JsonPropertyName("fontColor")]
    public string? FontColor { get; set; }

    [JsonPropertyName("highlight")]
    public string? Highlight { get; set; }

    [JsonPropertyName("verticalAlign")]
    public string? VerticalAlign { get; set; }
}

public sealed class SectionDecision
{
    [JsonPropertyName("afterBlockId")]
    public string? AfterBlockId { get; set; }

    [JsonPropertyName("sectionType")]
    public string? SectionType { get; set; }

    [JsonPropertyName("applySectionKey")]
    public string? ApplySectionKey { get; set; }

    [JsonPropertyName("pageStart")]
    public int? PageStart { get; set; }

    [JsonPropertyName("orientation")]
    public string? Orientation { get; set; }

    [JsonPropertyName("headerType")]
    public string? HeaderType { get; set; }

    [JsonPropertyName("footerType")]
    public string? FooterType { get; set; }

    [JsonPropertyName("headerSourcePath")]
    public string? HeaderSourcePath { get; set; }

    [JsonPropertyName("footerSourcePath")]
    public string? FooterSourcePath { get; set; }

    [JsonPropertyName("headerDecisionKey")]
    public string? HeaderDecisionKey { get; set; }

    [JsonPropertyName("footerDecisionKey")]
    public string? FooterDecisionKey { get; set; }
}

public sealed class HeaderFooterDecision
{
    [JsonPropertyName("decisionKey")]
    public string? DecisionKey { get; set; }

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "";

    [JsonPropertyName("type")]
    public string Type { get; set; } = "default";

    [JsonPropertyName("action")]
    public string Action { get; set; } = "useTemplate";

    [JsonPropertyName("applyStyleKey")]
    public string? ApplyStyleKey { get; set; }

    [JsonPropertyName("targetSectionKey")]
    public string? TargetSectionKey { get; set; }

    [JsonPropertyName("sourcePath")]
    public string? SourcePath { get; set; }
}

public sealed class CaptionDecision
{
    [JsonPropertyName("blockId")]
    public string BlockId { get; set; } = "";

    [JsonPropertyName("captionKind")]
    public string CaptionKind { get; set; } = "";

    [JsonPropertyName("position")]
    public string Position { get; set; } = "";

    [JsonPropertyName("applyStyleKey")]
    public string? ApplyStyleKey { get; set; }
}

public sealed class RunDecision
{
    [JsonPropertyName("blockId")]
    public string BlockId { get; set; } = "";

    [JsonPropertyName("runPath")]
    public string? RunPath { get; set; }

    [JsonPropertyName("matchText")]
    public string? MatchText { get; set; }

    [JsonPropertyName("occurrence")]
    public int? Occurrence { get; set; }

    [JsonPropertyName("segmentPath")]
    public string? SegmentPath { get; set; }

    [JsonPropertyName("segmentText")]
    public string? SegmentText { get; set; }

    [JsonPropertyName("replaceText")]
    public string? ReplaceText { get; set; }

    [JsonPropertyName("replaceSegmentText")]
    public string? ReplaceSegmentText { get; set; }

    [JsonPropertyName("fontSize")]
    public string? FontSize { get; set; }

    [JsonPropertyName("asciiFont")]
    public string? AsciiFont { get; set; }

    [JsonPropertyName("eastAsiaFont")]
    public string? EastAsiaFont { get; set; }

    [JsonPropertyName("bold")]
    public bool? Bold { get; set; }

    [JsonPropertyName("italic")]
    public bool? Italic { get; set; }

    [JsonPropertyName("underline")]
    public string? Underline { get; set; }

    [JsonPropertyName("fontColor")]
    public string? FontColor { get; set; }

    [JsonPropertyName("highlight")]
    public string? Highlight { get; set; }

    [JsonPropertyName("verticalAlign")]
    public string? VerticalAlign { get; set; }
}

public sealed class ExtractedAsset
{
    [JsonPropertyName("assetId")]
    public string AssetId { get; set; } = "";

    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "image";

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("contentType")]
    public string? ContentType { get; set; }

    [JsonPropertyName("sourcePath")]
    public string? SourcePath { get; set; }
}

public sealed class ImageAssetSpec
{
    [JsonPropertyName("assetId")]
    public string AssetId { get; set; } = "";

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("contentType")]
    public string? ContentType { get; set; }

    [JsonPropertyName("widthEmu")]
    public long? WidthEmu { get; set; }

    [JsonPropertyName("heightEmu")]
    public long? HeightEmu { get; set; }

    [JsonPropertyName("altText")]
    public string? AltText { get; set; }
}

public sealed class RenderImageSpec
{
    [JsonPropertyName("path")]
    public string? Path { get; set; }

    [JsonPropertyName("assetId")]
    public string? AssetId { get; set; }

    [JsonPropertyName("contentType")]
    public string? ContentType { get; set; }

    [JsonPropertyName("widthEmu")]
    public long? WidthEmu { get; set; }

    [JsonPropertyName("heightEmu")]
    public long? HeightEmu { get; set; }

    [JsonPropertyName("altText")]
    public string? AltText { get; set; }
}
