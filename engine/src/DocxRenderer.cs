using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxWeb;

internal static class DocxRenderer
{
    public static void Render(RenderSpec spec, string outputPath, string? templatePath = null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        using var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        if (!string.IsNullOrWhiteSpace(templatePath) && File.Exists(templatePath))
            CopyTemplateStyles(templatePath, mainPart);

        EnsureDocumentSettings(mainPart, spec);

        ApplyPackageProperties(document, spec.Document);

        var headerRefs = CreateHeaderParts(mainPart, spec.Headers);
        var footerRefs = CreateFooterParts(mainPart, spec.Footers);

        var body = mainPart.Document.Body!;
        var segments = SplitSegments(spec);
        if (segments.Count == 0)
        {
            segments.Add(new BodySegment
            {
                Section = ToDefaultSection(spec.Document)
            });
        }

        for (var index = 0; index < segments.Count; index++)
        {
            var segment = segments[index];
            foreach (var block in segment.Blocks)
            {
                switch (block.Kind.ToLowerInvariant())
                {
                    case "paragraph":
                    case "heading":
                    case "reference":
                        body.Append(BuildParagraph(block.Paragraph));
                        break;
                    case "table":
                        body.Append(BuildTable(block.Table));
                        break;
                    case "image":
                        var imageParagraph = BuildImageParagraph(mainPart, block.Image);
                        if (imageParagraph != null)
                            body.Append(imageParagraph);
                        break;
                    case "equation":
                        var equationParagraph = BuildEquationParagraph(block.Equation);
                        if (equationParagraph != null)
                            body.Append(equationParagraph);
                        break;
                }
            }

            if (index < segments.Count - 1)
            {
                AttachSectionBreak(body, BuildSectionProperties(
                    segment.Section ?? ToDefaultSection(spec.Document),
                    headerRefs,
                    footerRefs));
            }
        }

        body.Append(BuildSectionProperties(
            segments[^1].Section ?? ToDefaultSection(spec.Document),
            headerRefs,
            footerRefs));

        mainPart.Document.Save();
    }

    private static void ApplyPackageProperties(WordprocessingDocument document, RenderDocumentSpec spec)
    {
        if (!string.IsNullOrWhiteSpace(spec.Title))
            document.PackageProperties.Title = spec.Title;

        if (!string.IsNullOrWhiteSpace(spec.Author))
            document.PackageProperties.Creator = spec.Author;
    }

    private static Dictionary<string, string> CreateHeaderParts(
        MainDocumentPart mainPart,
        IEnumerable<RenderHeaderFooterSpec> headers)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var header in headers)
        {
            var type = NormalizeHeaderFooterType(header.Type);
            var part = mainPart.AddNewPart<HeaderPart>();
            part.Header = new Header();
            foreach (var paragraph in header.Paragraphs)
            {
                paragraph.Paragraph.FirstLineChars = "0";
                paragraph.Paragraph.FirstLineIndent = "0";
                part.Header.Append(BuildParagraph(paragraph));
            }

            AppendTextBoxParagraphs(part.Header, header.TextBoxes, ensureHostParagraph: header.Paragraphs.Count == 0);

            var partId = mainPart.GetIdOfPart(part);
            if (!result.ContainsKey($"type:{type}"))
                result[$"type:{type}"] = partId;
            if (!string.IsNullOrWhiteSpace(header.SourcePath))
                result[$"path:{header.SourcePath}"] = partId;
            if (!string.IsNullOrWhiteSpace(header.RelationshipId))
                result[$"rel:{header.RelationshipId}"] = partId;
        }

        return result;
    }

    private static Dictionary<string, string> CreateFooterParts(
        MainDocumentPart mainPart,
        IEnumerable<RenderHeaderFooterSpec> footers)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var footer in footers)
        {
            var type = NormalizeHeaderFooterType(footer.Type);
            var part = mainPart.AddNewPart<FooterPart>();
            part.Footer = new Footer();
            foreach (var paragraph in footer.Paragraphs)
            {
                paragraph.Paragraph.FirstLineChars = "0";
                paragraph.Paragraph.FirstLineIndent = "0";
                part.Footer.Append(BuildParagraph(paragraph));
            }

            AppendTextBoxParagraphs(part.Footer, footer.TextBoxes, ensureHostParagraph: footer.Paragraphs.Count == 0);

            var partId = mainPart.GetIdOfPart(part);
            if (!result.ContainsKey($"type:{type}"))
                result[$"type:{type}"] = partId;
            if (!string.IsNullOrWhiteSpace(footer.SourcePath))
                result[$"path:{footer.SourcePath}"] = partId;
            if (!string.IsNullOrWhiteSpace(footer.RelationshipId))
                result[$"rel:{footer.RelationshipId}"] = partId;
        }

        return result;
    }

    private static void AppendTextBoxParagraphs(
        OpenXmlElement container,
        IList<RenderTextBoxSpec> textBoxes,
        bool ensureHostParagraph)
    {
        if (textBoxes.Count == 0)
        {
            if (ensureHostParagraph && !container.Elements<Paragraph>().Any())
                container.Append(new Paragraph(new Run()));
            return;
        }

        var hostParagraph = container.Elements<Paragraph>().LastOrDefault();
        if (hostParagraph == null)
        {
            hostParagraph = new Paragraph();
            var properties = new ParagraphProperties();
            properties.Append(new SpacingBetweenLines
            {
                Before = "0",
                After = "0",
                Line = "240",
                LineRule = LineSpacingRuleValues.Auto
            });
            hostParagraph.Append(properties);
            container.Append(hostParagraph);
        }

        foreach (var textBox in textBoxes)
        {
            var run = BuildTextBoxRun(textBox);
            if (run != null)
                hostParagraph.Append(run);
        }
    }

    private static Run? BuildTextBoxRun(RenderTextBoxSpec textBox)
    {
        var xml = BuildTextBoxAlternateContentXml(textBox);
        if (string.IsNullOrEmpty(xml))
            return null;

        var run = new Run();
        try
        {
            run.InnerXml = xml;
        }
        catch
        {
            return null;
        }

        return run;
    }

    private static string BuildTextBoxAlternateContentXml(RenderTextBoxSpec textBox)
    {
        var posX = textBox.PosXEmu ?? 457200L;
        var posY = textBox.PosYEmu ?? 1828800L;
        var width = textBox.WidthEmu ?? 457200L;
        var height = textBox.HeightEmu ?? 7772400L;
        var textDirection = string.IsNullOrWhiteSpace(textBox.TextDirection) ? "vert270" : textBox.TextDirection!;
        var relativeFromH = string.IsNullOrWhiteSpace(textBox.RelativeFromH) ? "page" : textBox.RelativeFromH!;
        var relativeFromV = string.IsNullOrWhiteSpace(textBox.RelativeFromV) ? "page" : textBox.RelativeFromV!;
        var anchor = string.IsNullOrWhiteSpace(textBox.Anchor) ? "ctr" : textBox.Anchor!;
        var name = string.IsNullOrWhiteSpace(textBox.Name) ? "VerticalTextBox" : textBox.Name!;

        var isFilled = !string.IsNullOrWhiteSpace(textBox.FillColor)
            && !string.Equals(textBox.FillColor, "none", StringComparison.OrdinalIgnoreCase);
        var fill = isFilled
            ? $"<a:solidFill><a:srgbClr val=\"{textBox.FillColor}\"/></a:solidFill>"
            : "<a:noFill/>";

        var isStroked = !string.IsNullOrWhiteSpace(textBox.BorderStyle)
            && !string.Equals(textBox.BorderStyle, "none", StringComparison.OrdinalIgnoreCase);
        var border = isStroked
            ? "<a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>"
            : "<a:ln><a:noFill/></a:ln>";

        var paragraphXml = string.Concat(textBox.Paragraphs.Select(p => BuildParagraph(p).OuterXml));
        if (string.IsNullOrEmpty(paragraphXml))
            paragraphXml = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>";

        var culture = System.Globalization.CultureInfo.InvariantCulture;
        var vmlLeft = (posX / 12700.0).ToString("0.##", culture);
        var vmlTop = (posY / 12700.0).ToString("0.##", culture);
        var vmlWidth = (width / 12700.0).ToString("0.##", culture);
        var vmlHeight = (height / 12700.0).ToString("0.##", culture);

        var vmlLayoutFlow = textDirection switch
        {
            "vert270" => "vertical;mso-layout-flow-alt:bottom-to-top",
            "vert" => "vertical",
            _ => "vertical"
        };

        var fillVml = isFilled ? $"filled=\"t\" fillcolor=\"#{textBox.FillColor}\"" : "filled=\"f\"";
        var strokedVml = isStroked ? "stroked=\"t\"" : "stroked=\"f\"";

        return $"<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<mc:Choice Requires=\"wps\">"
            + "<w:drawing>"
            + "<wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
            + "<wp:simplePos x=\"0\" y=\"0\"/>"
            + $"<wp:positionH relativeFrom=\"{relativeFromH}\"><wp:posOffset>{posX}</wp:posOffset></wp:positionH>"
            + $"<wp:positionV relativeFrom=\"{relativeFromV}\"><wp:posOffset>{posY}</wp:posOffset></wp:positionV>"
            + $"<wp:extent cx=\"{width}\" cy=\"{height}\"/>"
            + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>"
            + "<wp:wrapNone/>"
            + $"<wp:docPr id=\"1\" name=\"{name}\"/>"
            + "<wp:cNvGraphicFramePr/>"
            + "<a:graphic>"
            + "<a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<wps:wsp>"
            + "<wps:cNvSpPr txBox=\"1\"/>"
            + "<wps:spPr>"
            + $"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm>"
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            + fill
            + border
            + "</wps:spPr>"
            + "<wps:txbx>"
            + $"<w:txbxContent>{paragraphXml}</w:txbxContent>"
            + "</wps:txbx>"
            + $"<wps:bodyPr rot=\"0\" spcFirstLastPara=\"0\" vertOverflow=\"visible\" horzOverflow=\"visible\" vert=\"{textDirection}\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" numCol=\"1\" spcCol=\"0\" rtlCol=\"0\" fromWordArt=\"0\" anchor=\"{anchor}\" anchorCtr=\"0\" forceAA=\"0\" compatLnSpc=\"1\"/>"
            + "</wps:wsp>"
            + "</a:graphicData>"
            + "</a:graphic>"
            + "</wp:anchor>"
            + "</w:drawing>"
            + "</mc:Choice>"
            + "<mc:Fallback>"
            + "<w:pict xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\">"
            + $"<v:rect id=\"_x0000_s1026\" style=\"position:absolute;margin-left:{vmlLeft}pt;margin-top:{vmlTop}pt;width:{vmlWidth}pt;height:{vmlHeight}pt;z-index:251659264;mso-position-horizontal-relative:{relativeFromH};mso-position-vertical-relative:{relativeFromV}\" {strokedVml} {fillVml}>"
            + $"<v:textbox style=\"layout-flow:{vmlLayoutFlow}\" inset=\"7.2pt,3.6pt,7.2pt,3.6pt\">"
            + $"<w:txbxContent>{paragraphXml}</w:txbxContent>"
            + "</v:textbox>"
            + "</v:rect>"
            + "</w:pict>"
            + "</mc:Fallback>"
            + "</mc:AlternateContent>";
    }

    private static List<BodySegment> SplitSegments(RenderSpec spec)
    {
        var segments = new List<BodySegment>
        {
            new()
            {
                Section = ToDefaultSection(spec.Document)
            }
        };

        var current = segments[0];

        foreach (var block in spec.Body)
        {
            if (string.Equals(block.Kind, "section", StringComparison.OrdinalIgnoreCase))
            {
                if (current.Blocks.Count == 0 && segments.Count == 1)
                {
                    current.Section = MergeSection(current.Section, block.Section);
                    continue;
                }

                current = new BodySegment
                {
                    Section = MergeSection(ToDefaultSection(spec.Document), block.Section)
                };
                segments.Add(current);
                continue;
            }

            current.Blocks.Add(block);
        }

        return segments;
    }

    private static RenderSectionSpec ToDefaultSection(RenderDocumentSpec spec)
    {
        return new RenderSectionSpec
        {
            PageWidth = spec.PageWidth,
            PageHeight = spec.PageHeight,
            Orientation = spec.Orientation,
            MarginTop = spec.MarginTop,
            MarginBottom = spec.MarginBottom,
            MarginLeft = spec.MarginLeft,
            MarginRight = spec.MarginRight,
            PageNumFmt = null,
            Headers = [],
            Footers = []
        };
    }

    private static RenderSectionSpec MergeSection(RenderSectionSpec? fallback, RenderSectionSpec? overrideSpec)
    {
        fallback ??= new RenderSectionSpec();
        if (overrideSpec == null)
            return fallback;

        return new RenderSectionSpec
        {
            Type = overrideSpec.Type ?? fallback.Type,
            PageWidth = overrideSpec.PageWidth ?? fallback.PageWidth,
            PageHeight = overrideSpec.PageHeight ?? fallback.PageHeight,
            Orientation = overrideSpec.Orientation ?? fallback.Orientation,
            MarginTop = overrideSpec.MarginTop ?? fallback.MarginTop,
            MarginBottom = overrideSpec.MarginBottom ?? fallback.MarginBottom,
            MarginLeft = overrideSpec.MarginLeft ?? fallback.MarginLeft,
            MarginRight = overrideSpec.MarginRight ?? fallback.MarginRight,
            PageStart = overrideSpec.PageStart ?? fallback.PageStart,
            PageNumFmt = overrideSpec.PageNumFmt ?? fallback.PageNumFmt,
            TitlePage = overrideSpec.TitlePage || fallback.TitlePage,
            HeaderType = overrideSpec.HeaderType ?? fallback.HeaderType,
            FooterType = overrideSpec.FooterType ?? fallback.FooterType,
            HeaderSourcePath = overrideSpec.HeaderSourcePath ?? fallback.HeaderSourcePath,
            FooterSourcePath = overrideSpec.FooterSourcePath ?? fallback.FooterSourcePath,
            Headers = overrideSpec.Headers.Count > 0 ? [.. overrideSpec.Headers] : [.. fallback.Headers],
            Footers = overrideSpec.Footers.Count > 0 ? [.. overrideSpec.Footers] : [.. fallback.Footers]
        };
    }

    private static Paragraph BuildParagraph(RenderParagraphSpec? spec)
    {
        spec ??= new RenderParagraphSpec();

        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        var headingLevel = spec.HeadingLevel;
        var styleId = spec.Style;
        var paragraphFormat = spec.Paragraph ?? new ParagraphFormatProfile();
        var align = spec.Align ?? paragraphFormat.Align;
        var defaultRunFormat = GetDefaultRunFormat(headingLevel, spec.Run);

        if (string.IsNullOrWhiteSpace(styleId) && headingLevel.HasValue)
            styleId = $"Heading{headingLevel.Value}";

        if (!string.IsNullOrWhiteSpace(styleId))
            properties.Append(new ParagraphStyleId { Val = styleId });

        if (headingLevel.HasValue)
            properties.Append(new OutlineLevel { Val = headingLevel.Value - 1 });

        var justification = ParseAlignment(align);
        if (justification.HasValue)
            properties.Append(new Justification { Val = justification.Value });

        AppendIndentation(properties, paragraphFormat);
        AppendSpacing(properties, paragraphFormat);
        AppendTabStops(properties, paragraphFormat);

        if (spec.PageBreakBefore || paragraphFormat.PageBreakBefore)
            properties.Append(new PageBreakBefore());

        if (properties.ChildElements.Count > 0)
            paragraph.Append(properties);

        AppendContentRuns(paragraph, spec, headingLevel);

        if (spec.PageNumberField)
        {
            if (!string.IsNullOrWhiteSpace(spec.PageNumberPrefix))
                paragraph.Append(BuildTextRun(spec.PageNumberPrefix, defaultRunFormat, headingLevel));

            paragraph.Append(BuildField(" PAGE ", defaultRunFormat, headingLevel));

            if (!string.IsNullOrWhiteSpace(spec.PageNumberSuffix))
                paragraph.Append(BuildTextRun(spec.PageNumberSuffix, defaultRunFormat, headingLevel));
        }

        if (spec.NumPagesField)
        {
            if (!string.IsNullOrEmpty(spec.Text) || spec.PageNumberField)
                paragraph.Append(BuildTextRun(" / ", defaultRunFormat, headingLevel));

            paragraph.Append(BuildField(" NUMPAGES ", defaultRunFormat, headingLevel));
        }

        if (!paragraph.Elements<Run>().Any() && !paragraph.Elements<SimpleField>().Any())
            paragraph.Append(new Run());

        return paragraph;
    }

    private static Paragraph? BuildImageParagraph(MainDocumentPart mainPart, RenderImageSpec? spec)
    {
        if (spec == null)
            return null;

        var imagePath = spec.Path;
        if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            return null;

        var imagePartType = GetImagePartType(imagePath, spec.ContentType);
        var imagePart = mainPart.AddImagePart(imagePartType);
        using (var stream = File.OpenRead(imagePath))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);
        var width = spec.WidthEmu.GetValueOrDefault(4_000_000L);
        var height = spec.HeightEmu.GetValueOrDefault(3_000_000L);
        var name = Path.GetFileName(imagePath);
        var altText = spec.AltText ?? name;

        var element =
            new Drawing(
                new Dw.Inline(
                    new Dw.Extent { Cx = width, Cy = height },
                    new Dw.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new Dw.DocProperties
                    {
                        Id = (UInt32Value)1U,
                        Name = name,
                        Description = altText
                    },
                    new Dw.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new Pic.Picture(
                                new Pic.NonVisualPictureProperties(
                                    new Pic.NonVisualDrawingProperties
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = name,
                                        Description = altText
                                    },
                                    new Pic.NonVisualPictureDrawingProperties()),
                                new Pic.BlipFill(
                                    new A.Blip { Embed = relationshipId },
                                    new A.Stretch(new A.FillRectangle())),
                                new Pic.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = width, Cy = height }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle
                                    })))
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });

        return new Paragraph(new Run(element));
    }

    private static Paragraph? BuildEquationParagraph(RenderEquationSpec? spec)
    {
        if (spec == null)
            return null;

        if (string.IsNullOrWhiteSpace(spec.Xml) && !string.IsNullOrWhiteSpace(spec.Text))
        {
            var paragraph = new Paragraph();
            ApplyEquationParagraphProperties(paragraph, spec);

            AppendFormattedText(paragraph, spec.Text, GetDefaultEquationRunFormat(spec.Run), null);
            return paragraph;
        }

        try
        {
            var paragraph = new Paragraph();
            ApplyEquationParagraphProperties(paragraph, spec);

            var sanitizedXml = SanitizeEquationXml(spec.Xml);
            if (sanitizedXml.Contains("oMathPara", StringComparison.OrdinalIgnoreCase))
            {
                paragraph.Append(new DocumentFormat.OpenXml.Math.Paragraph(sanitizedXml));
            }
            else
            {
                paragraph.Append(new DocumentFormat.OpenXml.Math.OfficeMath(sanitizedXml));
            }
            return paragraph;
        }
        catch
        {
            var paragraph = new Paragraph();
            ApplyEquationParagraphProperties(paragraph, spec);

            var label = string.Equals(spec.DisplayMode, "display", StringComparison.OrdinalIgnoreCase)
                ? "[Equation]"
                : "[Inline Equation]";
            var text = string.IsNullOrWhiteSpace(spec.Text) ? label : spec.Text;
            AppendFormattedText(paragraph, text, GetDefaultEquationRunFormat(spec.Run), null);
            return paragraph;
        }
    }

    private static Table BuildTable(RenderTableSpec? spec)
    {
        spec ??= new RenderTableSpec();
        var rows = spec.Rows.ToList();

        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableJustification { Val = TableRowAlignmentValues.Center },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 12U },
                new BottomBorder { Val = BorderValues.Single, Size = 12U },
                new LeftBorder { Val = BorderValues.None, Size = 0U },
                new RightBorder { Val = BorderValues.None, Size = 0U },
                new InsideHorizontalBorder { Val = BorderValues.None, Size = 0U },
                new InsideVerticalBorder { Val = BorderValues.None, Size = 0U })));

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var isHeaderRow = rowIndex == 0 && rows.Count > 1;
            var tableRow = new TableRow();

            foreach (var cellText in rows[rowIndex])
            {
                var cellParagraph = new Paragraph();
                var cellFormat = isHeaderRow ? new RunFormatProfile { Bold = true } : null;
                AppendFormattedText(cellParagraph, cellText ?? string.Empty, cellFormat, null);
                if (!cellParagraph.Elements<Run>().Any())
                    cellParagraph.Append(new Run());

                var cellProperties = new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Auto });

                if (isHeaderRow)
                    cellProperties.Append(new TableCellBorders(
                        new BottomBorder { Val = BorderValues.Single, Size = 6U }));

                var cell = new TableCell(cellProperties, cellParagraph);
                tableRow.Append(cell);
            }

            table.Append(tableRow);
        }

        if (rows.Count == 0)
        {
            table.Append(new TableRow(
                new TableCell(
                    new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }),
                    new Paragraph(new Run(new Text(string.Empty))))));
        }

        return table;
    }

    private static void AppendContentRuns(Paragraph paragraph, RenderParagraphSpec spec, int? headingLevel)
    {
        if (spec.Runs.Count > 0)
        {
            foreach (var run in spec.Runs)
            {
                var kind = run.Kind?.ToLowerInvariant();
                switch (kind)
                {
                    case "bookmarkstart":
                        if (!string.IsNullOrWhiteSpace(run.BookmarkId) && !string.IsNullOrWhiteSpace(run.BookmarkName))
                        {
                            paragraph.Append(new BookmarkStart
                            {
                                Id = run.BookmarkId,
                                Name = run.BookmarkName
                            });
                        }
                        continue;
                    case "bookmarkend":
                        if (!string.IsNullOrWhiteSpace(run.BookmarkId))
                            paragraph.Append(new BookmarkEnd { Id = run.BookmarkId });
                        continue;
                    case "tab":
                    {
                        var tabRun = new Run();
                        var tabProps = BuildRunProperties(run.Format, headingLevel);
                        if (tabProps.ChildElements.Count > 0)
                            tabRun.Append(tabProps);
                        tabRun.Append(new TabChar());
                        paragraph.Append(tabRun);
                        continue;
                    }
                    case "reffield":
                    {
                        var field = new SimpleField { Instruction = run.FieldInstruction ?? string.Empty };
                        var fieldRun = new Run();
                        var fieldRunProps = BuildRunProperties(run.Format, headingLevel);
                        if (fieldRunProps.ChildElements.Count > 0)
                            fieldRun.Append(fieldRunProps);
                        AppendRunTextContent(fieldRun, run.FieldDisplayText ?? string.Empty);
                        field.Append(fieldRun);
                        paragraph.Append(field);
                        continue;
                    }
                }

                if (run.Segments.Count > 0)
                {
                    foreach (var segment in run.Segments)
                    {
                        if (string.IsNullOrEmpty(segment.Text))
                            continue;

                        AppendFormattedText(paragraph, segment.Text, segment.Format, headingLevel);
                    }
                    continue;
                }

                if (!string.IsNullOrEmpty(run.Text))
                    AppendFormattedText(paragraph, run.Text, run.Format, headingLevel);
            }

            return;
        }

        if (!string.IsNullOrEmpty(spec.Text))
            AppendFormattedText(paragraph, spec.Text, spec.Run, headingLevel);
    }

    private static Run BuildTextRun(string text, RunFormatProfile? format, int? headingLevel)
    {
        var run = new Run();
        var properties = BuildRunProperties(format, headingLevel);

        if (properties.ChildElements.Count > 0)
            run.Append(properties);

        AppendRunTextContent(run, text);
        return run;
    }

    private static void AppendFormattedText(Paragraph paragraph, string text, RunFormatProfile? format, int? headingLevel)
    {
        foreach (var segment in SplitAutoFormatSegments(text, format))
            paragraph.Append(BuildTextRun(segment.Text, segment.Format, headingLevel));
    }

    private static List<AutoFormatSegment> SplitAutoFormatSegments(string text, RunFormatProfile? baseFormat)
    {
        baseFormat ??= new RunFormatProfile();
        if (string.IsNullOrEmpty(text))
            return [new AutoFormatSegment(text, CloneRunFormat(baseFormat))];

        if (!string.IsNullOrWhiteSpace(baseFormat.VerticalAlign))
            return [new AutoFormatSegment(text, CloneRunFormat(baseFormat))];

        var result = new List<AutoFormatSegment>();
        var buffer = "";
        var index = 0;

        while (index < text.Length)
        {
            if (TryReadScriptMarker(text, ref index, out var markerText, out var markerAlign))
            {
                FlushBuffer(result, ref buffer, baseFormat);
                result.Add(new AutoFormatSegment(markerText, CloneRunFormat(baseFormat, markerAlign)));
                continue;
            }

            if (TryReadToken(text, ref index, out var token))
            {
                if (TrySplitChemistryToken(token, baseFormat, out var chemicalSegments)
                    || TrySplitUnitExponentToken(token, baseFormat, out chemicalSegments))
                {
                    FlushBuffer(result, ref buffer, baseFormat);
                    result.AddRange(chemicalSegments);
                }
                else
                {
                    buffer += token;
                }

                continue;
            }

            buffer += text[index];
            index++;
        }

        FlushBuffer(result, ref buffer, baseFormat);
        return result.Count == 0
            ? [new AutoFormatSegment(text, CloneRunFormat(baseFormat))]
            : result;
    }

    private static bool TryReadScriptMarker(string text, ref int index, out string markerText, out string markerAlign)
    {
        markerText = "";
        markerAlign = "";

        if (index >= text.Length)
            return false;

        var marker = text[index];
        if (marker != '^' && marker != '_')
            return false;

        var start = index + 1;
        if (start >= text.Length)
            return false;

        if (text[start] == '{')
        {
            var endBrace = text.IndexOf('}', start + 1);
            if (endBrace <= start + 1)
                return false;

            markerText = text[(start + 1)..endBrace];
            markerAlign = marker == '^' ? "superscript" : "subscript";
            index = endBrace + 1;
            return markerText.Length > 0;
        }

        var cursor = start;
        while (cursor < text.Length && IsScriptChar(text[cursor]))
            cursor++;

        if (cursor == start)
            return false;

        markerText = text[start..cursor];
        markerAlign = marker == '^' ? "superscript" : "subscript";
        index = cursor;
        return true;
    }

    private static bool TryReadToken(string text, ref int index, out string token)
    {
        token = "";
        if (index >= text.Length || !IsTokenChar(text[index]))
            return false;

        var start = index;
        while (index < text.Length && IsTokenChar(text[index]))
            index++;

        token = text[start..index];
        return token.Length > 0;
    }

    private static bool TrySplitChemistryToken(string token, RunFormatProfile baseFormat, out List<AutoFormatSegment> segments)
    {
        segments = [];
        if (string.IsNullOrWhiteSpace(token))
            return false;

        var hasUpper = token.Any(char.IsUpper);
        var hasDigit = token.Any(char.IsDigit);
        if (!hasUpper || (!hasDigit && !token.Contains('+') && !token.Contains('-')))
            return false;

        var chargeMatch = Regex.Match(token, @"^(?<base>.*?)(?<charge>\d*[+-]+)$");
        var core = chargeMatch.Success ? chargeMatch.Groups["base"].Value : token;
        var charge = chargeMatch.Success ? chargeMatch.Groups["charge"].Value : null;

        var plain = "";
        for (var i = 0; i < core.Length; i++)
        {
            var ch = core[i];
            if (char.IsDigit(ch) && i > 0 && IsChemicalSubscriptAnchor(core[i - 1]))
            {
                FlushBuffer(segments, ref plain, baseFormat);
                var start = i;
                while (i < core.Length && char.IsDigit(core[i]))
                    i++;

                segments.Add(new AutoFormatSegment(core[start..i], CloneRunFormat(baseFormat, "subscript")));
                i--;
                continue;
            }

            plain += ch;
        }

        FlushBuffer(segments, ref plain, baseFormat);

        if (!string.IsNullOrEmpty(charge))
            segments.Add(new AutoFormatSegment(charge, CloneRunFormat(baseFormat, "superscript")));

        return segments.Any(x => !string.IsNullOrWhiteSpace(x.Format.VerticalAlign));
    }

    private static bool TrySplitUnitExponentToken(string token, RunFormatProfile baseFormat, out List<AutoFormatSegment> segments)
    {
        segments = [];
        var match = Regex.Match(token, @"^(?<base>(?:mm|cm|dm|m|km|nm|um|μm|µm|L|mL|μL|µL|g|kg|mg|s|min|h|Hz|kHz|MHz|GHz|Pa|kPa|MPa|V|mV|A|mA|W|kW|J|kJ|N|mol|℃|°C|K))(?<exp>[23])$");
        if (!match.Success)
            return false;

        segments.Add(new AutoFormatSegment(match.Groups["base"].Value, CloneRunFormat(baseFormat)));
        segments.Add(new AutoFormatSegment(match.Groups["exp"].Value, CloneRunFormat(baseFormat, "superscript")));
        return true;
    }

    private static void FlushBuffer(List<AutoFormatSegment> segments, ref string buffer, RunFormatProfile baseFormat)
    {
        if (buffer.Length == 0)
            return;

        segments.Add(new AutoFormatSegment(buffer, CloneRunFormat(baseFormat)));
        buffer = "";
    }

    private static bool IsChemicalSubscriptAnchor(char ch)
    {
        return char.IsLetter(ch) || ch == ')' || ch == ']' || ch == '}' ;
    }

    private static bool IsTokenChar(char ch)
    {
        return char.IsLetterOrDigit(ch)
            || ch == '('
            || ch == ')'
            || ch == '['
            || ch == ']'
            || ch == '{'
            || ch == '}'
            || ch == '+'
            || ch == '-'
            || ch == '·'
            || ch == '°'
            || ch == 'μ'
            || ch == 'µ';
    }

    private static bool IsScriptChar(char ch)
    {
        return char.IsLetterOrDigit(ch) || ch == '+' || ch == '-';
    }

    private static RunFormatProfile CloneRunFormat(RunFormatProfile source, string? verticalAlign = null)
    {
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
            VerticalAlign = verticalAlign ?? source.VerticalAlign
        };
    }

    private sealed record AutoFormatSegment(string Text, RunFormatProfile Format);

    private static void AppendRunTextContent(Run run, string text)
    {
        var parts = text.Split('\t');
        for (var i = 0; i < parts.Length; i++)
        {
            if (parts[i].Length > 0)
                run.Append(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });

            if (i < parts.Length - 1)
                run.Append(new TabChar());
        }
    }

    private static RunProperties BuildRunProperties(RunFormatProfile? format, int? headingLevel)
    {
        format = GetDefaultRunFormat(headingLevel, format);
        var properties = new RunProperties();
        var asciiFont = NormalizeAsciiFont(format.AsciiFont);
        var eastAsiaFont = NormalizeEastAsiaFont(format.EastAsiaFont);

        if (!string.IsNullOrWhiteSpace(asciiFont) || !string.IsNullOrWhiteSpace(eastAsiaFont))
        {
            properties.Append(new RunFonts
            {
                Ascii = asciiFont,
                HighAnsi = asciiFont,
                EastAsia = eastAsiaFont,
                ComplexScript = eastAsiaFont ?? asciiFont
            });
        }

        if (format.Bold || headingLevel.HasValue)
            properties.Append(new Bold());

        if (format.Italic)
            properties.Append(new Italic());

        if (!string.IsNullOrWhiteSpace(format.Underline)
            && TryParseUnderline(format.Underline, out var underline))
        {
            properties.Append(new Underline { Val = underline });
        }

        if (!string.IsNullOrWhiteSpace(format.FontColor))
            properties.Append(new Color { Val = format.FontColor });

        if (!string.IsNullOrWhiteSpace(format.Highlight)
            && TryParseHighlight(format.Highlight, out var highlight))
        {
            properties.Append(new Highlight { Val = highlight });
        }

        if (TryParseVerticalAlign(format.VerticalAlign, out var verticalAlign))
            properties.Append(new VerticalTextAlignment { Val = verticalAlign });

        var fontSize = format.FontSize ?? GetFallbackFontSize(headingLevel);
        if (!string.IsNullOrWhiteSpace(fontSize))
        {
            properties.Append(new FontSize { Val = fontSize });
            properties.Append(new FontSizeComplexScript { Val = fontSize });
        }

        return properties;
    }

    private static string GetFallbackFontSize(int? headingLevel)
    {
        return headingLevel switch
        {
            1 => "36",
            2 => "32",
            3 => "28",
            _ => "24"
        };
    }

    private static void AppendIndentation(ParagraphProperties properties, ParagraphFormatProfile format)
    {
        var indentation = new Indentation();
        var hasIndentation = false;

        if (TryParseStringValue(format.LeftIndent, out var leftIndent))
        {
            indentation.Left = leftIndent;
            hasIndentation = true;
        }

        if (TryParseStringValue(format.RightIndent, out var rightIndent))
        {
            indentation.Right = rightIndent;
            hasIndentation = true;
        }

        if (TryParseStringValue(format.FirstLineIndent, out var firstLineIndent))
        {
            indentation.FirstLine = firstLineIndent;
            hasIndentation = true;
        }

        if (TryParseInt32String(format.FirstLineChars, out var firstLineChars))
        {
            indentation.FirstLineChars = firstLineChars;
            hasIndentation = true;
        }

        if (TryParseStringValue(format.HangingIndent, out var hangingIndent))
        {
            indentation.Hanging = hangingIndent;
            hasIndentation = true;
        }

        if (TryParseInt32String(format.HangingChars, out var hangingChars))
        {
            indentation.HangingChars = hangingChars;
            hasIndentation = true;
        }

        if (hasIndentation)
            properties.Append(indentation);
    }

    private static void AppendTabStops(ParagraphProperties properties, ParagraphFormatProfile format)
    {
        if (format.TabStops == null || format.TabStops.Count == 0)
            return;

        var tabs = new Tabs();
        foreach (var stop in format.TabStops)
        {
            if (!TryParseInt32String(stop.Position, out var position))
                continue;
            var tabStop = new TabStop
            {
                Val = ParseTabStopAlignment(stop.Alignment),
                Position = position
            };
            tabs.Append(tabStop);
        }

        if (tabs.HasChildren)
            properties.Append(tabs);
    }

    private static TabStopValues ParseTabStopAlignment(string? alignment)
    {
        return (alignment ?? "left").Trim().ToLowerInvariant() switch
        {
            "center" => TabStopValues.Center,
            "right" => TabStopValues.Right,
            "decimal" => TabStopValues.Decimal,
            "bar" => TabStopValues.Bar,
            "clear" => TabStopValues.Clear,
            _ => TabStopValues.Left
        };
    }

    private static void AppendSpacing(ParagraphProperties properties, ParagraphFormatProfile format)
    {
        var spacing = new SpacingBetweenLines();
        var hasSpacing = false;

        if (TryParseStringValue(format.BeforeSpacing, out var beforeSpacing))
        {
            spacing.Before = beforeSpacing;
            hasSpacing = true;
        }

        if (TryParseStringValue(format.AfterSpacing, out var afterSpacing))
        {
            spacing.After = afterSpacing;
            hasSpacing = true;
        }

        if (TryParseStringValue(format.LineSpacing, out var lineSpacing))
        {
            spacing.Line = lineSpacing;
            hasSpacing = true;
        }

        if (TryParseLineRule(format.LineSpacingRule, out var lineRule))
        {
            spacing.LineRule = lineRule;
            hasSpacing = true;
        }

        if (hasSpacing)
            properties.Append(spacing);
    }

    private static SimpleField BuildField(string instruction, RunFormatProfile? format, int? headingLevel)
    {
        var field = new SimpleField
        {
            Instruction = instruction
        };

        var run = new Run();
        var properties = BuildRunProperties(format, headingLevel);
        if (properties.ChildElements.Count > 0)
            run.Append(properties);

        run.Append(new Text(string.Empty));
        field.Append(run);
        return field;
    }

    private static RunFormatProfile GetDefaultRunFormat(int? headingLevel, RunFormatProfile? format)
    {
        format ??= new RunFormatProfile();
        return new RunFormatProfile
        {
            FontSize = format.FontSize ?? GetFallbackFontSize(headingLevel),
            AsciiFont = string.IsNullOrWhiteSpace(format.AsciiFont) ? "Times New Roman" : format.AsciiFont,
            EastAsiaFont = string.IsNullOrWhiteSpace(format.EastAsiaFont) ? "宋体" : format.EastAsiaFont,
            Bold = format.Bold,
            Italic = format.Italic,
            Underline = format.Underline,
            FontColor = NormalizeFontColor(format.FontColor),
            Highlight = format.Highlight,
            VerticalAlign = format.VerticalAlign
        };
    }

    private static string? NormalizeAsciiFont(string? font)
    {
        if (string.IsNullOrWhiteSpace(font))
            return "Times New Roman";

        return font;
    }

    private static string? NormalizeEastAsiaFont(string? font)
    {
        if (string.IsNullOrWhiteSpace(font))
            return "宋体";

        return font.Contains("瀹", StringComparison.Ordinal)
            || font.Contains("嬩", StringComparison.Ordinal)
            ? "宋体"
            : font;
    }

    private static RunFormatProfile GetDefaultEquationRunFormat(RunFormatProfile? format)
    {
        return GetDefaultRunFormat(null, format);
    }

    private static string? NormalizeFontColor(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
            return null;

        return color.Equals("auto", StringComparison.OrdinalIgnoreCase)
            || color.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase)
            || color.Equals("white", StringComparison.OrdinalIgnoreCase)
            ? "000000"
            : color;
    }

    private static void ApplyEquationParagraphProperties(Paragraph paragraph, RenderEquationSpec spec)
    {
        var properties = new ParagraphProperties();
        var paragraphFormat = spec.Paragraph ?? new ParagraphFormatProfile();
        var align = string.Equals(spec.DisplayMode, "display", StringComparison.OrdinalIgnoreCase)
            ? "center"
            : paragraphFormat.Align;

        var justification = ParseAlignment(align);
        if (justification.HasValue)
            properties.Append(new Justification { Val = justification.Value });

        AppendIndentation(properties, paragraphFormat);

        // OMML 公式高度由内容决定，不能用 exact 行高，否则高公式（分式、积分等）被截断
        var isOmml = !string.IsNullOrWhiteSpace(spec.Xml);
        var spacingFormat = isOmml
            ? new ParagraphFormatProfile
              {
                  BeforeSpacing = paragraphFormat.BeforeSpacing,
                  AfterSpacing = paragraphFormat.AfterSpacing
              }
            : paragraphFormat;
        AppendSpacing(properties, spacingFormat);

        if (properties.ChildElements.Count > 0)
            paragraph.Append(properties);
    }

    private static string SanitizeEquationXml(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
            return xml;

        try
        {
            var element = XElement.Parse(xml, LoadOptions.PreserveWhitespace);
            foreach (var runProperties in element.Descendants().Where(x => x.Name.LocalName == "rPr"))
            {
                runProperties.Elements()
                    .Where(x => x.Name.LocalName is "color" or "highlight" or "shd")
                    .Remove();
            }

            return element.ToString(SaveOptions.DisableFormatting);
        }
        catch
        {
            return xml;
        }
    }

    private static void AttachSectionBreak(Body body, SectionProperties sectionProperties)
    {
        Paragraph hostParagraph;
        if (body.LastChild is Paragraph lastParagraph)
        {
            hostParagraph = lastParagraph;
        }
        else
        {
            hostParagraph = new Paragraph(new Run());
            body.Append(hostParagraph);
        }

        var properties = hostParagraph.GetFirstChild<ParagraphProperties>();
        if (properties == null)
        {
            properties = new ParagraphProperties();
            hostParagraph.PrependChild(properties);
        }

        properties.RemoveAllChildren<SectionProperties>();
        properties.Append(sectionProperties);
    }

    private static SectionProperties BuildSectionProperties(
        RenderSectionSpec section,
        Dictionary<string, string> headerRefs,
        Dictionary<string, string> footerRefs)
    {
        var properties = new SectionProperties();

        foreach (var pair in SelectHeaderFooterRefs(headerRefs, section.Headers, section.HeaderType, section.HeaderSourcePath))
        {
            properties.Append(new HeaderReference
            {
                Type = ParseHeaderFooterType(pair.Key),
                Id = pair.Value
            });
        }

        foreach (var pair in SelectHeaderFooterRefs(footerRefs, section.Footers, section.FooterType, section.FooterSourcePath))
        {
            properties.Append(new FooterReference
            {
                Type = ParseHeaderFooterType(pair.Key),
                Id = pair.Value
            });
        }

        var pageSize = new PageSize();
        var hasPageSize = false;

        if (TryParseUInt(section.PageWidth, out var width))
        {
            pageSize.Width = width;
            hasPageSize = true;
        }

        if (TryParseUInt(section.PageHeight, out var height))
        {
            pageSize.Height = height;
            hasPageSize = true;
        }

        if (TryParseOrientation(section.Orientation, out var orientation))
        {
            pageSize.Orient = orientation;
            hasPageSize = true;
        }

        if (hasPageSize)
            properties.Append(pageSize);

        var pageMargin = new PageMargin();
        var hasMargin = false;

        if (TryParseInt(section.MarginTop, out var top))
        {
            pageMargin.Top = top;
            hasMargin = true;
        }

        if (TryParseInt(section.MarginBottom, out var bottom))
        {
            pageMargin.Bottom = bottom;
            hasMargin = true;
        }

        if (TryParseUInt(section.MarginLeft, out var left))
        {
            pageMargin.Left = left;
            hasMargin = true;
        }

        if (TryParseUInt(section.MarginRight, out var right))
        {
            pageMargin.Right = right;
            hasMargin = true;
        }

        if (hasMargin)
            properties.Append(pageMargin);

        if (section.PageStart.HasValue)
        {
            var pageNumberType = new PageNumberType
            {
                Start = section.PageStart.Value
            };

            if (TryParsePageNumberFormat(section.PageNumFmt, out var pageNumberFormat))
                pageNumberType.Format = pageNumberFormat;

            properties.Append(pageNumberType);
        }
        else if (TryParsePageNumberFormat(section.PageNumFmt, out var pageNumberFormat))
        {
            properties.Append(new PageNumberType
            {
                Format = pageNumberFormat
            });
        }

        if (section.TitlePage)
            properties.Append(new TitlePage());

        if (TryParseSectionType(section.Type, out var sectionType))
        {
            properties.Append(new SectionType
            {
                Val = sectionType
            });
        }

        return properties;
    }

    private static HeaderFooterValues ParseHeaderFooterType(string? value)
    {
        return value?.Trim().ToLowerInvariant() switch
        {
            "first" => HeaderFooterValues.First,
            "even" => HeaderFooterValues.Even,
            _ => HeaderFooterValues.Default
        };
    }

    private static bool TryParsePageNumberFormat(string? value, out NumberFormatValues format)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "decimal":
                format = NumberFormatValues.Decimal;
                return true;
            case "upperroman":
                format = NumberFormatValues.UpperRoman;
                return true;
            case "lowerroman":
                format = NumberFormatValues.LowerRoman;
                return true;
            default:
                format = default;
                return false;
        }
    }

    private static string NormalizeHeaderFooterType(string? value)
    {
        return value?.Trim().ToLowerInvariant() switch
        {
            "first" => "first",
            "even" => "even",
            _ => "default"
        };
    }

    private static IEnumerable<KeyValuePair<string, string>> SelectHeaderFooterRefs(
        Dictionary<string, string> refs,
        IReadOnlyList<RenderHeaderFooterReferenceSpec> explicitRefs,
        string? preferredType,
        string? preferredSourcePath)
    {
        if (explicitRefs.Count > 0)
        {
            var explicitResult = new List<KeyValuePair<string, string>>();
            foreach (var item in explicitRefs)
            {
                var type = NormalizeHeaderFooterType(item.Type);
                if (!string.IsNullOrWhiteSpace(item.SourcePath)
                    && refs.TryGetValue($"path:{item.SourcePath}", out var partByPath))
                {
                    explicitResult.Add(new KeyValuePair<string, string>(type, partByPath));
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.SourcePath)
                    && refs.TryGetValue($"rel:{item.SourcePath}", out var partByRel))
                {
                    explicitResult.Add(new KeyValuePair<string, string>(type, partByRel));
                    continue;
                }

                if (refs.TryGetValue($"type:{type}", out var partByType))
                    explicitResult.Add(new KeyValuePair<string, string>(type, partByType));
            }

            if (explicitResult.Count > 0)
                return explicitResult;
        }

        if (!string.IsNullOrWhiteSpace(preferredSourcePath)
            && refs.TryGetValue($"path:{preferredSourcePath}", out var preferredPath))
        {
            return [new KeyValuePair<string, string>(NormalizeHeaderFooterType(preferredType), preferredPath)];
        }

        if (!string.IsNullOrWhiteSpace(preferredSourcePath)
            && refs.TryGetValue($"rel:{preferredSourcePath}", out var preferredRel))
        {
            return [new KeyValuePair<string, string>(NormalizeHeaderFooterType(preferredType), preferredRel)];
        }

        var normalizedPreferred = NormalizeHeaderFooterType(preferredType);
        if (refs.TryGetValue($"type:{normalizedPreferred}", out var preferred))
            return [new KeyValuePair<string, string>(normalizedPreferred, preferred)];

        if (refs.TryGetValue("type:default", out var fallback))
            return [new KeyValuePair<string, string>("default", fallback)];

        return refs
            .Where(x => x.Key.StartsWith("type:", StringComparison.OrdinalIgnoreCase))
            .Select(x => new KeyValuePair<string, string>(x.Key["type:".Length..], x.Value));
    }

    private static void CopyTemplateStyles(string templatePath, MainDocumentPart targetMainPart)
    {
        using var templateDocument = WordprocessingDocument.Open(templatePath, false);
        var templateMainPart = templateDocument.MainDocumentPart;
        if (templateMainPart?.StyleDefinitionsPart?.Styles == null)
            return;

        var stylePart = targetMainPart.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles = (Styles)templateMainPart.StyleDefinitionsPart.Styles.CloneNode(true);

        if (templateMainPart.StylesWithEffectsPart?.Styles != null)
        {
            var effectsPart = targetMainPart.AddNewPart<StylesWithEffectsPart>();
            effectsPart.Styles = (DocumentFormat.OpenXml.Wordprocessing.Styles)templateMainPart.StylesWithEffectsPart.Styles.CloneNode(true);
        }
    }

    private static void EnsureDocumentSettings(MainDocumentPart mainPart, RenderSpec spec)
    {
        var needsOddEvenHeaders = spec.Headers.Any(x =>
            string.Equals(x.Type, "odd", StringComparison.OrdinalIgnoreCase)
            || string.Equals(x.Type, "even", StringComparison.OrdinalIgnoreCase))
            || spec.Footers.Any(x =>
                string.Equals(x.Type, "odd", StringComparison.OrdinalIgnoreCase)
                || string.Equals(x.Type, "even", StringComparison.OrdinalIgnoreCase))
            || spec.Body
                .Where(x => x.Section != null)
                .SelectMany(x => x.Section!.Headers.Concat(x.Section!.Footers))
                .Any(x =>
                    string.Equals(x.Type, "odd", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(x.Type, "even", StringComparison.OrdinalIgnoreCase));

        if (!needsOddEvenHeaders)
            return;

        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();

        if (!settingsPart.Settings.Elements<EvenAndOddHeaders>().Any())
            settingsPart.Settings.Append(new EvenAndOddHeaders());

        settingsPart.Settings.Save();
    }

    private static JustificationValues? ParseAlignment(string? value)
    {
        return value?.Trim().ToLowerInvariant() switch
        {
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "both" => JustificationValues.Both,
            "justify" => JustificationValues.Both,
            "distribute" => JustificationValues.Distribute,
            "left" => JustificationValues.Left,
            _ => null
        };
    }

    private static bool TryParseOrientation(string? value, out PageOrientationValues orientation)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "landscape":
                orientation = PageOrientationValues.Landscape;
                return true;
            case "portrait":
                orientation = PageOrientationValues.Portrait;
                return true;
            default:
                orientation = default;
                return false;
        }
    }

    private static bool TryParseSectionType(string? value, out SectionMarkValues sectionType)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "continuous":
                sectionType = SectionMarkValues.Continuous;
                return true;
            case "evenpage":
                sectionType = SectionMarkValues.EvenPage;
                return true;
            case "oddpage":
                sectionType = SectionMarkValues.OddPage;
                return true;
            case "nextcolumn":
                sectionType = SectionMarkValues.NextColumn;
                return true;
            case "nextpage":
                sectionType = SectionMarkValues.NextPage;
                return true;
            default:
                sectionType = default;
                return false;
        }
    }

    private static bool TryParseUnderline(string? value, out UnderlineValues underline)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "single":
                underline = UnderlineValues.Single;
                return true;
            case "double":
                underline = UnderlineValues.Double;
                return true;
            case "none":
                underline = UnderlineValues.None;
                return true;
            default:
                underline = default;
                return false;
        }
    }

    private static bool TryParseLineRule(string? value, out LineSpacingRuleValues lineRule)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "auto":
                lineRule = LineSpacingRuleValues.Auto;
                return true;
            case "exact":
                lineRule = LineSpacingRuleValues.Exact;
                return true;
            case "atleast":
            case "at-least":
                lineRule = LineSpacingRuleValues.AtLeast;
                return true;
            default:
                lineRule = default;
                return false;
        }
    }

    private static bool TryParseHighlight(string? value, out HighlightColorValues highlight)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "yellow":
                highlight = HighlightColorValues.Yellow;
                return true;
            case "green":
                highlight = HighlightColorValues.Green;
                return true;
            case "cyan":
                highlight = HighlightColorValues.Cyan;
                return true;
            case "magenta":
                highlight = HighlightColorValues.Magenta;
                return true;
            case "blue":
                highlight = HighlightColorValues.Blue;
                return true;
            case "red":
                highlight = HighlightColorValues.Red;
                return true;
            case "darkyellow":
                highlight = HighlightColorValues.DarkYellow;
                return true;
            case "darkblue":
                highlight = HighlightColorValues.DarkBlue;
                return true;
            case "darkcyan":
                highlight = HighlightColorValues.DarkCyan;
                return true;
            case "darkgreen":
                highlight = HighlightColorValues.DarkGreen;
                return true;
            case "darkmagenta":
                highlight = HighlightColorValues.DarkMagenta;
                return true;
            case "darkred":
                highlight = HighlightColorValues.DarkRed;
                return true;
            case "darkgray":
                highlight = HighlightColorValues.DarkGray;
                return true;
            case "lightgray":
                highlight = HighlightColorValues.LightGray;
                return true;
            case "black":
                highlight = HighlightColorValues.Black;
                return true;
            case "white":
                highlight = HighlightColorValues.White;
                return true;
            default:
                highlight = default;
                return false;
        }
    }

    private static bool TryParseVerticalAlign(string? value, out VerticalPositionValues verticalAlign)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "superscript":
            case "super":
            case "上标":
                verticalAlign = VerticalPositionValues.Superscript;
                return true;
            case "subscript":
            case "sub":
            case "下标":
                verticalAlign = VerticalPositionValues.Subscript;
                return true;
            case "baseline":
            case "normal":
            case "基线":
                verticalAlign = VerticalPositionValues.Baseline;
                return true;
            default:
                verticalAlign = default;
                return false;
        }
    }

    private static bool TryParseStringValue(string? value, out string result)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            result = value;
            return true;
        }

        result = string.Empty;
        return false;
    }

    private static PartTypeInfo GetImagePartType(string path, string? contentType)
    {
        return (contentType?.ToLowerInvariant(), Path.GetExtension(path).ToLowerInvariant()) switch
        {
            ("image/png", _) or (_, ".png") => ImagePartType.Png,
            ("image/jpeg", _) or ("image/jpg", _) or (_, ".jpg") or (_, ".jpeg") => ImagePartType.Jpeg,
            ("image/gif", _) or (_, ".gif") => ImagePartType.Gif,
            ("image/bmp", _) or (_, ".bmp") => ImagePartType.Bmp,
            ("image/tiff", _) or (_, ".tif") or (_, ".tiff") => ImagePartType.Tiff,
            ("image/x-emf", _) or (_, ".emf") => ImagePartType.Emf,
            ("image/x-wmf", _) or (_, ".wmf") => ImagePartType.Wmf,
            _ => ImagePartType.Png
        };
    }

    private static bool TryParseUInt(string? value, out UInt32Value result)
    {
        if (uint.TryParse(value, out var parsed))
        {
            result = parsed;
            return true;
        }

        result = default!;
        return false;
    }

    private static bool TryParseInt(string? value, out Int32Value result)
    {
        if (int.TryParse(value, out var parsed))
        {
            result = parsed;
            return true;
        }

        result = default!;
        return false;
    }

    private static bool TryParseInt32String(string? value, out Int32Value result)
    {
        if (int.TryParse(value, out var parsed))
        {
            result = parsed;
            return true;
        }

        result = default!;
        return false;
    }

    private sealed class BodySegment
    {
        public RenderSectionSpec? Section { get; set; }

        public List<RenderBodyBlock> Blocks { get; } = [];
    }
}
