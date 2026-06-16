using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Converters;

public sealed class DocxToTypstConverter : IDisposable
{
    private const double TwipsPerInch = 1440.0;
    private const double EmusPerInch = 914400.0;

    private readonly WordprocessingDocument document;
    private readonly string tempDirectory;
    private readonly string assetsDirectory;
    private readonly Dictionary<string, string> extractedImages = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Style> stylesById = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, bool> orderedNumberingByNumId = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> themeColors = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> diagnostics = [];
    private string? majorThemeFont;
    private string? minorThemeFont;

    public DocxToTypstConverter(WordprocessingDocument document)
    {
        this.document = document ?? throw new ArgumentNullException(nameof(document));
        tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeEditor", "docx-to-typst-" + Guid.NewGuid().ToString("N"));
        assetsDirectory = Path.Combine(tempDirectory, "assets");
        Directory.CreateDirectory(assetsDirectory);
        LoadTheme();
        LoadStyles();
        LoadNumbering();
    }

    public TypstDocument Convert()
    {
        Body? body = document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return new TypstDocument { TempDirectory = tempDirectory, AssetsDirectory = assetsDirectory };
        }

        MainDocumentPart? mainPart = document.MainDocumentPart;
        List<SectionProperties> sections = GetSectionPropertiesInDocumentOrder(body);
        TypstPageSetup pageSetup = ExtractPageSetup(sections.FirstOrDefault());
        List<TypstBlock> blocks = [];
        List<OpenXmlElement> elements = body.Elements().ToList();
        for (int i = 0; i < elements.Count; i++)
        {
            OpenXmlElement element = elements[i];
            switch (element)
            {
                case Wp.Paragraph paragraph when IsListParagraph(paragraph):
                    TypstListBlock list = ConvertList(elements, i, out int lastListIndex);
                    blocks.Add(list);
                    i = lastListIndex;
                    break;
                case Wp.Paragraph paragraph:
                    AddParagraphBlocks(blocks, paragraph, mainPart, pageSetup);
                    AddSectionBoundary(blocks, paragraph, sections, pageSetup);
                    break;
                case Wp.Table table:
                    blocks.Add(ConvertTable(table));
                    break;
            }
        }

        TypstHeaderFooterSet headerFooter = ConvertHeaderFooterSet(sections.FirstOrDefault(), pageSetup);
        return new TypstDocument
        {
            PageSetup = pageSetup,
            DefaultFontFamily = ExtractDefaultFontFamily(),
            DefaultFontSizePt = ExtractDefaultFontSize(),
            HeaderBlocks = headerFooter.DefaultHeaderBlocks,
            FooterBlocks = headerFooter.DefaultFooterBlocks,
            HeaderFooter = headerFooter,
            Blocks = blocks,
            Diagnostics = diagnostics.ToList(),
            TempDirectory = tempDirectory,
            AssetsDirectory = assetsDirectory
        };
    }

    public string GenerateTypstSource(TypstDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        TypstPageSetup page = document.PageSetup;
        List<string> pageOptions = BuildPageOptions(page);
        TypstHeaderFooterSet headerFooter = document.HeaderFooter;
        string? pageAnchored = AddHeaderFooterPageOptions(pageOptions, headerFooter, page);

        List<string> lines =
        [
            // Belt-and-braces: strip any embedded newlines from the joined page options
            // so the entire `#set page(...)` call always lives on a single physical line.
            $"#set page({string.Join(", ", pageOptions).Replace("\n", " ").Replace("\r", " ")})",
            $"#set text(font: {TypstFontValue(document.DefaultFontFamily)}, size: {FormatPt(document.DefaultFontSizePt)})",
            string.Empty
        ];

        if (pageAnchored is not null)
        {
            lines.Add(pageAnchored);
        }

        bool defaultHeaderFooterApplied = !headerFooter.ApplyDefaultAfterFirstPageBreak
            || (headerFooter.DefaultHeaderBlocks.Count == 0 && headerFooter.DefaultFooterBlocks.Count == 0);
        foreach (TypstBlock block in document.Blocks)
        {
            lines.Add(RenderBlock(block));
            if (block is TypstPageSettingsBlock)
            {
                defaultHeaderFooterApplied = true;
            }
            else if (!defaultHeaderFooterApplied && block is TypstPageBreakBlock)
            {
                List<string> defaultPageOptions = [];
                string? defaultAnchored = AddHeaderFooterPageOptions(defaultPageOptions, headerFooter with
                {
                    FirstHeaderBlocks = [],
                    FirstFooterBlocks = [],
                    HasTitlePage = false,
                    ApplyDefaultAfterFirstPageBreak = false
                }, page);
                if (defaultPageOptions.Count > 0)
                {
                    lines.Add($"#set page({string.Join(", ", defaultPageOptions)})");
                }

                if (defaultAnchored is not null)
                {
                    lines.Add(defaultAnchored);
                }

                defaultHeaderFooterApplied = true;
            }

            lines.Add(string.Empty);
        }

        return string.Join(Environment.NewLine, lines);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private TypstListBlock ConvertList(IReadOnlyList<OpenXmlElement> elements, int startIndex, out int lastListIndex)
    {
        Wp.Paragraph firstParagraph = (Wp.Paragraph)elements[startIndex];
        bool ordered = IsOrderedList(firstParagraph);
        List<TypstListItem> items = [];
        lastListIndex = startIndex;

        for (int i = startIndex; i < elements.Count; i++)
        {
            OpenXmlElement element = elements[i];
            if (element is not Wp.Paragraph paragraph || !IsListParagraph(paragraph))
            {
                break;
            }

            items.Add(new TypstListItem { Inlines = ConvertParagraphInlines(paragraph).Where(i => i.Kind != TypstInlineKind.LineBreak).ToList() });
            lastListIndex = i;
        }

        return new TypstListBlock { Ordered = ordered, Items = items };
    }

    private void AddParagraphBlocks(List<TypstBlock> blocks, Wp.Paragraph paragraph, OpenXmlPart? owningPart, TypstPageSetup pageSetup)
    {
        blocks.AddRange(ExtractShapes(paragraph, pageSetup));
        List<TypstInline> inlines = ConvertParagraphInlines(paragraph);
        List<TypstImageBlock> images = ExtractImages(paragraph, owningPart);
        TypstParagraphBlock block = CreateParagraphBlock(paragraph, inlines);

        if (inlines.Count > 0)
        {
            blocks.Add(block);
            blocks.AddRange(images);
        }
        else if (images.Count > 0)
        {
            blocks.Add(block with { ImageBlocks = images });
        }
        else if (ShouldPreserveEmptyParagraph(paragraph, block))
        {
            blocks.Add(block);
        }

        if (HasPageBreak(paragraph))
        {
            blocks.Add(new TypstPageBreakBlock());
        }
    }

    private static bool ShouldPreserveEmptyParagraph(Wp.Paragraph paragraph, TypstParagraphBlock block)
    {
        ParagraphProperties? properties = paragraph.ParagraphProperties;
        if (properties is null)
        {
            return false;
        }

        string? styleId = properties.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrEmpty(styleId) && !styleId.Equals("Normal", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return block.SpaceBeforePt is > 0 || block.SpaceAfterPt is > 0;
    }

    private List<TypstInline> ConvertParagraphInlines(Wp.Paragraph paragraph)
    {
        List<TypstInline> inlines = [];
        List<OpenXmlElement?> paragraphStyleRunProperties = GetParagraphStyleRunProperties(paragraph);
        FieldParseState fieldState = new();

        foreach (OpenXmlElement child in paragraph.ChildElements)
        {
            ConvertInlineElement(child, inlines, paragraphStyleRunProperties, fieldState);
        }

        return inlines;
    }

    private void ConvertInlineElement(OpenXmlElement element, List<TypstInline> inlines, List<OpenXmlElement?> paragraphStyleRunProperties, FieldParseState fieldState)
    {
        if (IsInTextBoxContent(element))
        {
            return;
        }

        if (element is SimpleField simpleField)
        {
            RunFormatting formatting = new();
            Run? firstRun = simpleField.Descendants<Run>().FirstOrDefault();
            if (firstRun is not null)
            {
                List<OpenXmlElement?> styleRunProperties = GetRunStyleProperties(firstRun);
                formatting = ComputeRunFormatting(firstRun.RunProperties, styleRunProperties, paragraphStyleRunProperties);
            }

            TypstInline? fieldInline = CreateFieldInline(simpleField.Instruction?.Value ?? GetXmlAttribute(simpleField, "instr") ?? string.Empty, formatting);
            if (fieldInline is not null)
            {
                inlines.Add(fieldInline);
                return;
            }
        }

        if (element is Run run)
        {
            ConvertRunInlines(run, inlines, paragraphStyleRunProperties, fieldState);
            return;
        }

        foreach (OpenXmlElement child in element.ChildElements)
        {
            ConvertInlineElement(child, inlines, paragraphStyleRunProperties, fieldState);
        }
    }

    private void ConvertRunInlines(Run run, List<TypstInline> inlines, List<OpenXmlElement?> paragraphStyleRunProperties, FieldParseState fieldState)
    {
        List<OpenXmlElement?> styleRunProperties = GetRunStyleProperties(run);
        RunFormatting runFormatting = ComputeRunFormatting(run.RunProperties, styleRunProperties, paragraphStyleRunProperties);
        if (fieldState.InField)
        {
            MergeFieldFormatting(fieldState.FieldFormatting, runFormatting);
        }

        foreach (OpenXmlElement child in run.ChildElements)
        {
            switch (child)
            {
                case FieldChar fieldChar:
                    HandleFieldChar(fieldChar, fieldState, inlines);
                    break;
                case FieldCode fieldCode:
                    fieldState.Instruction.Append(fieldCode.Text);
                    break;
                case Text text when !string.IsNullOrEmpty(text.Text) && !fieldState.SuppressResult:
                    inlines.Add(CreateTextInline(text.Text, run.RunProperties, styleRunProperties, paragraphStyleRunProperties));
                    break;
                case Break br when br.Type?.Value != BreakValues.Page && !fieldState.SuppressResult:
                    inlines.Add(new TypstInline { Kind = TypstInlineKind.LineBreak });
                    break;
                case TabChar when !fieldState.SuppressResult:
                    inlines.Add(new TypstInline { Kind = TypstInlineKind.Tab });
                    break;
                case SymbolChar symbol when !fieldState.SuppressResult:
                    inlines.Add(CreateSymbolInline(symbol, run.RunProperties, styleRunProperties, paragraphStyleRunProperties));
                    break;
            }
        }
    }

    private TypstParagraphBlock CreateParagraphBlock(Wp.Paragraph paragraph, List<TypstInline> inlines)
    {
        List<OpenXmlElement?> properties = GetParagraphFormattingProperties(paragraph);
        properties.Add(paragraph.ParagraphProperties);

        string? alignment = null;
        double? before = null;
        double? after = null;
        double? leading = null;
        foreach (OpenXmlElement? property in properties)
        {
            Justification? justification = GetChild<Justification>(property);
            if (justification?.Val?.Value is not null)
            {
                alignment = MapAlignment(justification.Val.Value);
            }

            SpacingBetweenLines? spacing = GetChild<SpacingBetweenLines>(property);
            if (spacing is not null)
            {
                // Defensive read: when w:afterAutospacing="1" is set on a <w:spacing> element,
                // the OpenXml SDK's typed Before/After properties can misread the value and
                // return the `after` value for `before` (collapsing both to the same number).
                // We detect this suspicious equality (typed Before == typed After) and fall
                // back to the raw XML attribute for `before`, which the SDK does not munge.
                string? beforeRaw = spacing.Before?.Value;
                string? afterRaw = spacing.After?.Value;
                if (beforeRaw is null || (afterRaw is not null && beforeRaw == afterRaw))
                {
                    beforeRaw = GetSpacingAttributeValue(spacing, "before") ?? beforeRaw;
                }

                double? beforePt = TwipsToPoints(beforeRaw);
                double? afterPt = TwipsToPoints(afterRaw);
                if (beforePt is not null)
                {
                    before = beforePt;
                }

                if (afterPt is not null)
                {
                    after = afterPt;
                }

                leading = ResolveAutoLineSpacingLeading(spacing, inlines) ?? leading;
            }
        }

        return new TypstParagraphBlock
        {
            Inlines = inlines,
            Alignment = alignment,
            SpaceBeforePt = before,
            SpaceAfterPt = after,
            LeadingPt = leading,
            BottomBorder = ExtractParagraphBottomBorder(paragraph)
        };
    }

    private TypstInline CreateTextInline(string text, OpenXmlElement? direct, IEnumerable<OpenXmlElement?> runStyle, IEnumerable<OpenXmlElement?> paragraphStyle)
    {
        RunFormatting formatting = ComputeRunFormatting(direct, runStyle, paragraphStyle);
        string normalizedText = NormalizeSymbolText(text, formatting.FontFamily);
        TypstInline? checkbox = CreateCheckboxInline(normalizedText);
        if (checkbox is not null)
        {
            return checkbox;
        }

        return new TypstInline
        {
            Kind = TypstInlineKind.Text,
            Text = normalizedText,
            Bold = formatting.Bold,
            Italic = formatting.Italic,
            Underline = formatting.Underline,
            Color = formatting.Color,
            FontSizePt = formatting.FontSizePt,
            FontFamily = formatting.FontFamily
        };
    }

    private RunFormatting ComputeRunFormatting(OpenXmlElement? direct, IEnumerable<OpenXmlElement?> runStyle, IEnumerable<OpenXmlElement?> paragraphStyle)
    {
        RunFormatting formatting = new();
        ApplyRunFormatting(formatting, document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle);
        foreach (OpenXmlElement? property in paragraphStyle)
        {
            ApplyRunFormatting(formatting, property);
        }

        foreach (OpenXmlElement? property in runStyle)
        {
            ApplyRunFormatting(formatting, property);
        }

        ApplyRunFormatting(formatting, direct);
        return formatting;
    }

    private TypstInline CreateSymbolInline(SymbolChar symbol, OpenXmlElement? direct, IEnumerable<OpenXmlElement?> runStyle, IEnumerable<OpenXmlElement?> paragraphStyle)
    {
        string font = symbol.Font?.Value ?? GetXmlAttribute(symbol, "font") ?? string.Empty;
        string character = symbol.Char?.Value ?? GetXmlAttribute(symbol, "char") ?? string.Empty;
        string text = MapSymbolCharacter(font, character) ?? string.Empty;
        if (string.IsNullOrEmpty(text) && int.TryParse(character, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int codePoint))
        {
            text = char.ConvertFromUtf32(codePoint);
        }

        return CreateTextInline(text, direct, runStyle, paragraphStyle);
    }

    private void HandleFieldChar(FieldChar fieldChar, FieldParseState fieldState, List<TypstInline> inlines)
    {
        FieldCharValues? type = fieldChar.FieldCharType?.Value;
        if (type == FieldCharValues.Begin)
        {
            fieldState.Instruction.Clear();
            fieldState.InField = true;
            fieldState.SuppressResult = false;
            fieldState.FieldFormatting = new RunFormatting();
            return;
        }

        if (type == FieldCharValues.Separate)
        {
            TypstInline? fieldInline = CreateFieldInline(fieldState.Instruction.ToString(), fieldState.FieldFormatting);
            if (fieldInline is not null)
            {
                inlines.Add(fieldInline);
            }

            fieldState.SuppressResult = true;
            return;
        }

        if (type == FieldCharValues.End)
        {
            fieldState.Instruction.Clear();
            fieldState.InField = false;
            fieldState.SuppressResult = false;
        }
    }

    private TypstInline? CreateFieldInline(string instruction, RunFormatting? formatting = null)
    {
        string normalized = Regex.Replace(instruction, "\\s+", " ").Trim();
        string? rawTypst = null;
        if (Regex.IsMatch(normalized, "^PAGE(?:\\s|$)", RegexOptions.IgnoreCase))
        {
            rawTypst = "#context counter(page).display()";
        }
        else if (Regex.IsMatch(normalized, "^NUMPAGES(?:\\s|$)", RegexOptions.IgnoreCase))
        {
            rawTypst = "#context counter(page).final().at(0)";
        }

        if (rawTypst is null)
        {
            return null;
        }

        return new TypstInline
        {
            Kind = TypstInlineKind.RawTypst,
            RawTypst = rawTypst,
            Bold = formatting?.Bold ?? false,
            Italic = formatting?.Italic ?? false,
            Underline = formatting?.Underline ?? false,
            Color = formatting?.Color,
            FontSizePt = formatting?.FontSizePt,
            FontFamily = formatting?.FontFamily
        };
    }

    private TypstTableBlock ConvertTable(Wp.Table table)
    {
        List<TypstTableRow> rows = [];
        foreach (Wp.TableRow row in table.Elements<Wp.TableRow>())
        {
            TableFirstRowFormatting? firstRowFormatting = GetTableFirstRowFormatting(table, row, rows.Count == 0);
            List<TypstTableCell> cells = [];
            foreach (Wp.TableCell cell in row.Elements<Wp.TableCell>())
            {
                List<TypstParagraphBlock> paragraphs = cell.Elements<Wp.Paragraph>()
                    .Select(p => CreateParagraphBlock(p, ConvertParagraphInlines(p)))
                    .Where(p => p.Inlines.Count > 0)
                    .ToList();
                TableStyleProperties? conditional = GetTableStyleProperties(table, row, firstRow: rows.Count == 0);
                OpenXmlElement? conditionalRunProperties = conditional?.GetFirstChild<RunProperties>();
                if (conditionalRunProperties is not null)
                {
                    paragraphs = paragraphs.Select(p => p with { Inlines = p.Inlines.Select(i => ApplyConditionalRunFormatting(i, conditionalRunProperties)).ToList() }).ToList();
                }

                if (firstRowFormatting is not null)
                {
                    paragraphs = paragraphs.Select(p => p with { Inlines = p.Inlines.Select(i => ApplyTableFirstRowFormatting(i, firstRowFormatting)).ToList() }).ToList();
                }

                TableCellProperties? conditionalCellProperties = conditional?.GetFirstChild<TableCellProperties>();
                cells.Add(new TypstTableCell
                {
                    Paragraphs = paragraphs,
                    ShadingColor = ResolveShadingColor(cell.TableCellProperties?.Shading) ?? ResolveShadingColor(conditionalCellProperties?.Shading) ?? firstRowFormatting?.ShadingColor
                });
            }

            rows.Add(new TypstTableRow { Cells = cells });
        }

        return new TypstTableBlock { Rows = rows, HasBorders = table.Descendants<TableBorders>().Any() || GetTableStyle(table)?.Descendants<TableBorders>().Any() == true };
    }

    private void AddSectionBoundary(List<TypstBlock> blocks, Wp.Paragraph paragraph, IReadOnlyList<SectionProperties> sections, TypstPageSetup currentPageSetup)
    {
        SectionProperties? sectionProperties = paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>();
        if (sectionProperties is null)
        {
            return;
        }

        int sectionIndex = sections.ToList().FindIndex(section => ReferenceEquals(section, sectionProperties));
        if (sectionIndex < 0 || sectionIndex + 1 >= sections.Count)
        {
            return;
        }

        SectionProperties nextSection = sections[sectionIndex + 1];
        TypstPageSetup nextPageSetup = ExtractPageSetup(nextSection, currentPageSetup);
        blocks.Add(new TypstPageSettingsBlock
        {
            PageSetup = nextPageSetup,
            HeaderFooter = ConvertHeaderFooterSet(nextSection, nextPageSetup),
            PageBreakBefore = !HasPageBreak(paragraph)
        });
    }

    private TypstHeaderFooterSet ConvertHeaderFooterSet(SectionProperties? sectionProperties, TypstPageSetup pageSetup)
    {
        bool hasTitlePage = sectionProperties?.GetFirstChild<TitlePage>() is not null;
        HeaderFooterValues defaultVariant = HasHeaderFooterReference(sectionProperties, HeaderFooterValues.Default)
            || !HasHeaderFooterReference(sectionProperties, HeaderFooterValues.Even)
            ? HeaderFooterValues.Default
            : HeaderFooterValues.Even;

        List<TypstBlock> defaultHeaderBlocks = ConvertHeaderFooterBlocks(sectionProperties, isHeader: true, defaultVariant, pageSetup);
        List<TypstBlock> defaultFooterBlocks = ConvertHeaderFooterBlocks(sectionProperties, isHeader: false, defaultVariant, pageSetup);

        return new TypstHeaderFooterSet
        {
            FirstHeaderBlocks = hasTitlePage ? ConvertHeaderFooterBlocks(sectionProperties, isHeader: true, HeaderFooterValues.First, pageSetup) : [],
            FirstFooterBlocks = hasTitlePage ? ConvertHeaderFooterBlocks(sectionProperties, isHeader: false, HeaderFooterValues.First, pageSetup) : [],
            DefaultHeaderBlocks = defaultHeaderBlocks,
            DefaultFooterBlocks = defaultFooterBlocks,
            HasTitlePage = hasTitlePage,
            ApplyDefaultAfterFirstPageBreak = hasTitlePage && (defaultHeaderBlocks.Count > 0 || defaultFooterBlocks.Count > 0)
        };
    }

    private static List<SectionProperties> GetSectionPropertiesInDocumentOrder(Body body)
    {
        List<SectionProperties> sections = body.Descendants<SectionProperties>().ToList();
        SectionProperties? bodySection = body.Elements<SectionProperties>().LastOrDefault();
        if (bodySection is not null && !sections.Contains(bodySection))
        {
            sections.Add(bodySection);
        }

        return sections;
    }

    private static bool HasHeaderFooterReference(SectionProperties? sectionProperties, HeaderFooterValues variant)
        => sectionProperties?.Elements<HeaderReference>().Any(reference => reference.Type?.Value == variant) == true
            || sectionProperties?.Elements<FooterReference>().Any(reference => reference.Type?.Value == variant) == true;

    private List<TypstBlock> ConvertHeaderFooterBlocks(SectionProperties? sectionProperties, bool isHeader, HeaderFooterValues variant, TypstPageSetup pageSetup)
    {
        MainDocumentPart? mainPart = document.MainDocumentPart;
        if (mainPart is null)
        {
            return [];
        }

        OpenXmlPart? part = isHeader
            ? ResolveHeaderPart(mainPart, sectionProperties, variant)
            : ResolveFooterPart(mainPart, sectionProperties, variant);
        OpenXmlElement? root = part switch
        {
            HeaderPart headerPart => headerPart.Header,
            FooterPart footerPart => footerPart.Footer,
            _ => null
        };

        if (root is null || part is null)
        {
            return [];
        }

        List<TypstBlock> blocks = [];
        foreach (OpenXmlElement element in root.ChildElements)
        {
            AddHeaderFooterElementBlocks(blocks, element, part, pageSetup);
        }

        return blocks;
    }

    private void AddHeaderFooterElementBlocks(List<TypstBlock> blocks, OpenXmlElement element, OpenXmlPart part, TypstPageSetup pageSetup)
    {
        switch (element)
        {
            case Wp.Paragraph paragraph when !IsInTextBoxContent(paragraph):
                AddParagraphBlocks(blocks, paragraph, part, pageSetup);
                break;
            case Wp.Table table:
                blocks.Add(ConvertTable(table));
                break;
            default:
                foreach (OpenXmlElement child in element.ChildElements)
                {
                    AddHeaderFooterElementBlocks(blocks, child, part, pageSetup);
                }

                break;
        }
    }

    private static HeaderPart? ResolveHeaderPart(MainDocumentPart mainPart, SectionProperties? sectionProperties, HeaderFooterValues variant)
    {
        string? relationshipId = sectionProperties?.Elements<HeaderReference>()
            .Where(reference => reference.Type?.Value == variant)
            .Select(reference => reference.Id?.Value)
            .FirstOrDefault(id => !string.IsNullOrWhiteSpace(id));

        return GetRelatedPart<HeaderPart>(mainPart, relationshipId);
    }

    private static FooterPart? ResolveFooterPart(MainDocumentPart mainPart, SectionProperties? sectionProperties, HeaderFooterValues variant)
    {
        string? relationshipId = sectionProperties?.Elements<FooterReference>()
            .Where(reference => reference.Type?.Value == variant)
            .Select(reference => reference.Id?.Value)
            .FirstOrDefault(id => !string.IsNullOrWhiteSpace(id));

        return GetRelatedPart<FooterPart>(mainPart, relationshipId);
    }

    private static TPart? GetRelatedPart<TPart>(OpenXmlPart ownerPart, string? relationshipId)
        where TPart : OpenXmlPart
    {
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            return null;
        }

        try
        {
            return ownerPart.GetPartById(relationshipId) as TPart;
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    private List<TypstImageBlock> ExtractImages(OpenXmlElement scope, OpenXmlPart? owningPart)
    {
        if (owningPart is null)
        {
            return [];
        }

        List<TypstImageBlock> images = [];
        foreach (A.Blip blip in scope.Descendants<A.Blip>())
        {
            string? relationshipId = blip.Embed?.Value ?? blip.Link?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            ExtractedImage? extracted = ExtractImage(owningPart, relationshipId);
            if (extracted is null)
            {
                continue;
            }

            DW.Anchor? anchor = blip.Ancestors<DW.Anchor>().FirstOrDefault();
            DW.Extent? extent = blip.Ancestors<DW.Inline>().FirstOrDefault()?.Extent ?? anchor?.Extent;
            (double? xPt, double? yPt) = ExtractAnchorOffset(anchor);
            Wp.Paragraph? parentParagraph = blip.Ancestors<Wp.Paragraph>().FirstOrDefault();
            images.Add(new TypstImageBlock
            {
                Path = extracted.TypstPath,
                WidthInches = extent?.Cx is null ? null : extent.Cx.Value / EmusPerInch,
                HeightInches = extent?.Cy is null ? null : extent.Cy.Value / EmusPerInch,
                XPt = xPt,
                YPt = yPt,
                IsUnsupportedFormat = extracted.IsUnsupportedFormat,
                TopBorder = ExtractParagraphTopBorder(parentParagraph)
            });
        }

        return images;
    }

    private TypstBorderInfo? ExtractParagraphBottomBorder(Wp.Paragraph paragraph)
    {
        List<OpenXmlElement?> properties = GetParagraphFormattingProperties(paragraph);
        properties.Add(paragraph.ParagraphProperties);

        for (int i = properties.Count - 1; i >= 0; i--)
        {
            if (properties[i]?.GetFirstChild<ParagraphBorders>()?.BottomBorder is BottomBorder bottomBorder)
            {
                return ExtractBorderInfo(bottomBorder);
            }
        }

        return null;
    }

    private static TypstBorderInfo? ExtractBorderInfo(BorderType border)
    {
        BorderValues? borderValue = border.Val?.Value;
        if (borderValue == BorderValues.Nil || borderValue == BorderValues.None)
        {
            return null;
        }

        double sizeEighthPoints = 0;
        if (border.Size?.Value is uint size)
        {
            sizeEighthPoints = size;
        }

        string color = border.Color?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(color))
        {
            color = "000000";
        }

        return new TypstBorderInfo
        {
            Color = color,
            SizeEighthPoints = sizeEighthPoints
        };
    }

    private static TypstBorderInfo? ExtractParagraphTopBorder(Wp.Paragraph? paragraph)
    {
        TopBorder? topBorder = paragraph?.ParagraphProperties?.ParagraphBorders?.TopBorder;
        if (topBorder is null)
        {
            return null;
        }

        // Ignore borders that are explicitly disabled.
        BorderValues? borderValue = topBorder.Val?.Value;
        if (borderValue == BorderValues.Nil || borderValue == BorderValues.None)
        {
            return null;
        }

        double sizeEighthPoints = 0;
        if (topBorder.Size?.Value is uint size)
        {
            sizeEighthPoints = size;
        }

        string color = topBorder.Color?.Value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(color))
        {
            color = "000000";
        }

        return new TypstBorderInfo
        {
            Color = color,
            SizeEighthPoints = sizeEighthPoints
        };
    }

    private static (double? XPt, double? YPt) ExtractAnchorOffset(DW.Anchor? anchor)
    {
        if (anchor is null)
        {
            return (null, null);
        }

        List<double> offsets = Regex.Matches(anchor.OuterXml, "<wp:posOffset>(-?\\d+)</wp:posOffset>")
            .Select(match => double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) / 12700.0)
            .ToList();
        return offsets.Count >= 2 ? (offsets[0], offsets[1]) : (null, null);
    }

    private List<TypstShapeBlock> ExtractShapes(OpenXmlElement scope, TypstPageSetup pageSetup)
    {
        List<TypstShapeBlock> shapes = [];
        Dictionary<string, int> shapeTextIndexes = new(StringComparer.OrdinalIgnoreCase);
        foreach (OpenXmlElement element in scope.Descendants().Where(IsShapeElement))
        {
            if (element.Ancestors().Any(IsShapeElement))
            {
                continue;
            }

            ShapeGeometry geometry = ValidateShapeGeometry(ExtractShapeGeometry(element), pageSetup);
            List<TypstParagraphBlock> paragraphs = element.Descendants<Wp.Paragraph>()
                .Where(IsInTextBoxContent)
                .Select(p => CreateParagraphBlock(p, ConvertTextBoxParagraphInlines(p)))
                .Where(p => p.Inlines.Count > 0)
                .ToList();

            if (!geometry.IsValid && paragraphs.Count == 0)
            {
                continue;
            }

            if (paragraphs.Count == 0 && geometry.FillColor is null && (geometry.WidthPt is null || geometry.HeightPt is null))
            {
                continue;
            }

            TypstShapeBlock shape = new()
            {
                Paragraphs = paragraphs,
                XPt = geometry.XPt,
                YPt = geometry.YPt,
                WidthPt = geometry.WidthPt,
                HeightPt = geometry.HeightPt,
                FillColor = geometry.FillColor,
                StrokeColor = geometry.StrokeColor
            };

            string normalizedText = NormalizeShapeText(paragraphs);
            if (!string.IsNullOrEmpty(normalizedText) && shapeTextIndexes.TryGetValue(normalizedText, out int existingIndex))
            {
                TypstShapeBlock existing = shapes[existingIndex];
                if (IsPositioned(shape) && !IsPositioned(existing))
                {
                    shapes[existingIndex] = shape;
                }

                continue;
            }

            if (!string.IsNullOrEmpty(normalizedText))
            {
                shapeTextIndexes[normalizedText] = shapes.Count;
            }

            shapes.Add(shape);
        }

        return shapes;
    }

    private static bool IsPositioned(TypstShapeBlock shape) => shape.XPt is not null || shape.YPt is not null;

    private static string NormalizeShapeText(IEnumerable<TypstParagraphBlock> paragraphs)
    {
        string text = string.Concat(paragraphs.SelectMany(p => p.Inlines).Where(i => i.Kind == TypstInlineKind.Text).Select(i => i.Text));
        return Regex.Replace(text, "\\s+", " ").Trim();
    }

    private List<TypstInline> ConvertTextBoxParagraphInlines(Wp.Paragraph paragraph)
    {
        List<TypstInline> inlines = [];
        List<OpenXmlElement?> paragraphStyleRunProperties = GetParagraphStyleRunProperties(paragraph);
        foreach (Run run in paragraph.Descendants<Run>())
        {
            List<OpenXmlElement?> styleRunProperties = GetRunStyleProperties(run);
            foreach (OpenXmlElement child in run.ChildElements)
            {
                switch (child)
                {
                    case Text text when !string.IsNullOrEmpty(text.Text):
                        inlines.Add(CreateTextInline(text.Text, run.RunProperties, styleRunProperties, paragraphStyleRunProperties));
                        break;
                    case Break br when br.Type?.Value != BreakValues.Page:
                        inlines.Add(new TypstInline { Kind = TypstInlineKind.LineBreak });
                        break;
                    case TabChar:
                        inlines.Add(new TypstInline { Kind = TypstInlineKind.Tab });
                        break;
                }
            }
        }

        return inlines;
    }

    private static bool IsShapeElement(OpenXmlElement element)
    {
        string localName = element.LocalName;
        if (localName == "shape" || localName == "rect")
        {
            return true;
        }

        return localName == "wsp" && element.Descendants().Any(e => e.LocalName == "txbxContent" || e.LocalName == "txbx");
    }

    private static bool IsInTextBoxContent(OpenXmlElement element) => element.Ancestors().Any(e => e.LocalName == "txbxContent");

    private static ShapeGeometry ExtractShapeGeometry(OpenXmlElement element)
    {
        string style = GetXmlAttribute(element, "style") ?? string.Empty;
        return new ShapeGeometry(
            ReadStylePointValue(style, "margin-left") ?? ReadStylePointValue(style, "left"),
            ReadStylePointValue(style, "margin-top") ?? ReadStylePointValue(style, "top"),
            ReadStylePointValue(style, "width"),
            ReadStylePointValue(style, "height"),
            NormalizeColor(GetXmlAttribute(element, "fillcolor") ?? Regex.Match(element.OuterXml, "<v:fill[^>]*(?:color|fillcolor)=\"([^\"]+)\"").Groups[1].Value),
            NormalizeColor(GetXmlAttribute(element, "strokecolor") ?? Regex.Match(element.OuterXml, "<v:stroke[^>]*color=\"([^\"]+)\"").Groups[1].Value));
    }

    private static ShapeGeometry ValidateShapeGeometry(ShapeGeometry geometry, TypstPageSetup pageSetup)
    {
        double pageWidthPt = pageSetup.WidthInches * 72.0;
        double pageHeightPt = pageSetup.HeightInches * 72.0;
        double maxXPt = pageWidthPt * 2.0;
        double maxYPt = pageHeightPt * 2.0;

        bool impossible = IsImpossible(geometry.XPt, maxXPt)
            || IsImpossible(geometry.YPt, maxYPt)
            || IsImpossible(geometry.WidthPt, maxXPt)
            || IsImpossible(geometry.HeightPt, maxYPt);

        return impossible
            ? geometry with { XPt = null, YPt = null, WidthPt = null, HeightPt = null, IsValid = false }
            : geometry with { IsValid = true };
    }

    private static bool IsImpossible(double? value, double maxValue) => value is not null && Math.Abs(value.Value) > maxValue;

    private static double? ReadStylePointValue(string style, string property)
    {
        Match match = Regex.Match(style, $"(?:^|;)\\s*{Regex.Escape(property)}\\s*:\\s*(-?[0-9.]+)\\s*(pt|in)?", RegexOptions.IgnoreCase);
        if (!match.Success || !double.TryParse(match.Groups[1].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double value))
        {
            return null;
        }

        string unit = match.Groups[2].Value;
        if (unit.Equals("in", StringComparison.OrdinalIgnoreCase))
        {
            return value * 72.0;
        }

        return string.IsNullOrEmpty(unit) && Math.Abs(value) > 2000.0 ? value / 100.0 : value;
    }

    private ExtractedImage? ExtractImage(OpenXmlPart owningPart, string relationshipId)
    {
        string cacheKey = $"{owningPart.Uri}|{relationshipId}";
        if (extractedImages.TryGetValue(cacheKey, out string? existing))
        {
            return new ExtractedImage(existing, IsUnsupportedImagePath(existing));
        }

        OpenXmlPart? part;
        try
        {
            part = owningPart.GetPartById(relationshipId);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }

        if (part is not ImagePart imagePart)
        {
            return null;
        }

        string extension = ContentTypeToExtension(imagePart.ContentType);
        bool unsupported = IsUnsupportedImageExtension(extension);
        string fileName = $"image-{extractedImages.Count + 1}{extension}";
        string path = Path.Combine(assetsDirectory, fileName);
        using Stream input = imagePart.GetStream(FileMode.Open, FileAccess.Read);
        using FileStream output = File.Create(path);
        input.CopyTo(output);
        string typstPath = Path.Combine("assets", fileName).Replace('\\', '/');
        extractedImages[cacheKey] = typstPath;
        if (unsupported)
        {
            diagnostics.Add($"Skipped unsupported image format '{extension}' from {owningPart.Uri}; rendered a placeholder.");
        }

        return new ExtractedImage(typstPath, unsupported);
    }

    private TypstPageSetup ExtractPageSetup(SectionProperties? sectionProperties, TypstPageSetup? fallback = null)
    {
        PageSize? size = sectionProperties?.GetFirstChild<PageSize>();
        PageMargin? margin = sectionProperties?.GetFirstChild<PageMargin>();
        fallback ??= new TypstPageSetup();

        return new TypstPageSetup
        {
            WidthInches = TwipsToInches(size?.Width?.Value, fallback.WidthInches),
            HeightInches = TwipsToInches(size?.Height?.Value, fallback.HeightInches),
            Margins = new TypstMargins
            {
                TopInches = TwipsToInches(margin?.Top?.Value, fallback.Margins.TopInches),
                RightInches = TwipsToInches(margin?.Right?.Value, fallback.Margins.RightInches),
                BottomInches = TwipsToInches(margin?.Bottom?.Value, fallback.Margins.BottomInches),
                LeftInches = TwipsToInches(margin?.Left?.Value, fallback.Margins.LeftInches)
            },
            HeaderDistanceInches = margin?.Header?.Value is uint headerTwips ? headerTwips / TwipsPerInch : fallback.HeaderDistanceInches,
            FooterDistanceInches = margin?.Footer?.Value is uint footerTwips ? footerTwips / TwipsPerInch : fallback.FooterDistanceInches
        };
    }

    private string RenderBlock(TypstBlock block) => block switch
    {
        TypstParagraphBlock paragraph => RenderParagraph(paragraph),
        TypstPageBreakBlock => "#pagebreak()",
        TypstPageSettingsBlock settings => RenderPageSettings(settings),
        TypstListBlock list => RenderList(list),
        TypstTableBlock table => RenderTable(table),
        TypstImageBlock image => RenderImage(image),
        TypstShapeBlock shape => RenderShape(shape),
        _ => string.Empty
    };

    private string RenderParagraph(TypstParagraphBlock paragraph)
    {
        bool hasImages = paragraph.ImageBlocks.Count > 0;
        string content = hasImages
            ? string.Join(Environment.NewLine, paragraph.ImageBlocks.Select(RenderImage))
            : RenderInlines(paragraph.Inlines);

        if (paragraph.BottomBorder is not null)
        {
            content += $"\n#line(length: 100%, stroke: {FormatPt(paragraph.BottomBorder.SizeEighthPoints / 8.0)} + rgb(\"#{paragraph.BottomBorder.Color}\"))";
        }

        if (paragraph.Alignment == "justify" && !hasImages)
        {
            List<string> parOptions = ["justify: true"];
            if (paragraph.LeadingPt is > 0)
            {
                parOptions.Add($"leading: {FormatPt(paragraph.LeadingPt.Value)}");
            }

            content = $"#par({string.Join(", ", parOptions)})[{content}]";
        }
        else if (paragraph.Alignment is not null && !(hasImages && paragraph.Alignment == "justify"))
        {
            if (paragraph.LeadingPt is > 0 && !hasImages)
            {
                content = $"#align({paragraph.Alignment})[#par(leading: {FormatPt(paragraph.LeadingPt.Value)})[{content}]]";
            }
            else
            {
                content = $"#align({paragraph.Alignment})[{content}]";
            }
        }
        else if (paragraph.LeadingPt is > 0 && !hasImages)
        {
            content = $"#par(leading: {FormatPt(paragraph.LeadingPt.Value)})[{content}]";
        }

        List<string> spacingOptions = [];
        if (paragraph.SpaceBeforePt is > 0)
        {
            spacingOptions.Add($"above: {FormatPt(paragraph.SpaceBeforePt.Value)}");
        }

        if (paragraph.SpaceAfterPt is > 0)
        {
            spacingOptions.Add($"below: {FormatPt(paragraph.SpaceAfterPt.Value)}");
        }

        if (spacingOptions.Count == 0)
        {
            return content;
        }

        // Without an explicit width, a spacing block shrinks to its content, which makes
        // inner #align(center)/#par(justify) ineffective. Force full width when alignment
        // is applied so the paragraph occupies the full body column.
        if (paragraph.Alignment is not null)
        {
            spacingOptions.Insert(0, "width: 100%");
        }

        return $"#block({string.Join(", ", spacingOptions)})[{content}]";
    }

    private string RenderList(TypstListBlock list)
    {
        string function = list.Ordered ? "enum" : "list";
        return $"#{function}" + string.Concat(list.Items.Select(item => $"[{RenderInlines(item.Inlines)}]"));
    }

    private string RenderTable(TypstTableBlock table)
    {
        int columns = table.Rows.Select(r => r.Cells.Count).DefaultIfEmpty(1).Max();
        List<string> cells = [];
        foreach (TypstTableRow row in table.Rows)
        {
            cells.AddRange(row.Cells.Select(RenderTableCell));
            for (int i = row.Cells.Count; i < columns; i++)
            {
                cells.Add("[]");
            }
        }

        string stroke = table.HasBorders ? ", stroke: 0.5pt" : ", stroke: none";
        return $"#table(columns: {columns}, inset: 4pt{stroke}, {string.Join(", ", cells)})";
    }

    private string RenderTableCell(TypstTableCell cell)
    {
        string content = cell.Paragraphs.Count == 0
            ? string.Empty
            : string.Join("\n", cell.Paragraphs.Select(p => RenderInlines(p.Inlines)));

        return cell.ShadingColor is null
            ? $"[{content}]"
            : $"table.cell(fill: rgb(\"#{cell.ShadingColor}\"))[{content}]";
    }

    private string RenderImage(TypstImageBlock image)
    {
        string rendered = image.IsUnsupportedFormat
            ? RenderUnsupportedImagePlaceholder(image)
            : RenderSupportedImage(image);

        if (image.XPt is not null || image.YPt is not null)
        {
            return $"#place(dx: {FormatPt(image.XPt ?? 0)}, dy: {FormatPt(image.YPt ?? 0)})[{rendered}]";
        }

        if (image.TopBorder is not null)
        {
            rendered = $"#line(length: 100%, stroke: {FormatPt(image.TopBorder.SizeEighthPoints / 8.0)} + rgb(\"#{image.TopBorder.Color}\")){rendered}";
        }

        return rendered;
    }

    private string RenderSupportedImage(TypstImageBlock image)
    {
        List<string> arguments = [$"{TypstString(image.Path)}"];
        if (image.WidthInches is > 0)
        {
            arguments.Add($"width: {FormatIn(image.WidthInches.Value)}");
        }

        if (image.HeightInches is > 0)
        {
            arguments.Add($"height: {FormatIn(image.HeightInches.Value)}");
        }

        return $"#image({string.Join(", ", arguments)})";
    }

    private string RenderUnsupportedImagePlaceholder(TypstImageBlock image)
    {
        List<string> arguments = ["stroke: 0.5pt + rgb(\"#888888\")", "inset: 4pt"];
        if (image.WidthInches is > 0)
        {
            arguments.Add($"width: {FormatIn(image.WidthInches.Value)}");
        }

        if (image.HeightInches is > 0)
        {
            arguments.Add($"height: {FormatIn(image.HeightInches.Value)}");
        }

        return $"#rect({string.Join(", ", arguments)})[#align(center + horizon)[#text(size: 8pt, fill: rgb(\"#666666\"))[Exclusive Windows image format]]]";
    }

    private string RenderBlocksAsContent(IEnumerable<TypstBlock> blocks)
    {
        // Join the rendered blocks into a single string suitable for embedding inside
        // a Typst content block (e.g. `header: [<content>]` or `footer: [<content>]`).
        // For a single block, return its rendered text unchanged so the outer `[...]`
        // added by AddHeaderFooterPageOptions produces a clean `header: [block]` form.
        // For multiple blocks, concatenate them with an explicit `#parbreak()` so they
        // render as separate paragraphs. Using `#parbreak()` instead of `\n` keeps the
        // surrounding `#set page(...)` directive on a single line (avoiding parser
        // issues with a floating closing delimiter) and guarantees separation even when
        // a block renders as plain text, which a single newline would collapse into one
        // paragraph.
        List<string> rendered = blocks.Select(RenderBlock).Where(b => !string.IsNullOrWhiteSpace(b)).ToList();
        if (rendered.Count == 0)
        {
            return string.Empty;
        }

        if (rendered.Count == 1)
        {
            return rendered[0];
        }

        return string.Join("#parbreak()", rendered);
    }

    private string RenderPageSettings(TypstPageSettingsBlock settings)
    {
        List<string> pageOptions = BuildPageOptions(settings.PageSetup);
        string? anchored = AddHeaderFooterPageOptions(pageOptions, settings.HeaderFooter, settings.PageSetup);
        string setPage = $"#set page({string.Join(", ", pageOptions)})";
        string result = settings.PageBreakBefore ? $"#pagebreak()\n{setPage}" : setPage;
        if (anchored is not null)
        {
            result += $"\n{anchored}";
        }

        return result;
    }

    private static List<string> BuildPageOptions(TypstPageSetup page)
    {
        TypstMargins margins = page.Margins;
        double topMargin = margins.TopInches;
        double bottomMargin = margins.BottomInches;

        double? headerDistance = page.HeaderDistanceInches;
        if (headerDistance is not null && headerDistance.Value > topMargin)
        {
            topMargin = headerDistance.Value;
        }

        double? footerDistance = page.FooterDistanceInches;
        if (footerDistance is not null && footerDistance.Value > bottomMargin)
        {
            bottomMargin = footerDistance.Value;
        }

        List<string> options =
        [
            $"width: {FormatIn(page.WidthInches)}",
            $"height: {FormatIn(page.HeightInches)}",
            $"margin: (left: {FormatIn(margins.LeftInches)}, right: {FormatIn(margins.RightInches)}, top: {FormatIn(topMargin)}, bottom: {FormatIn(bottomMargin)})"
        ];

        if (headerDistance is not null)
        {
            double headerAscent = topMargin - headerDistance.Value;
            if (headerAscent >= 0)
            {
                options.Add($"header-ascent: {FormatIn(headerAscent)}");
            }
        }

        if (footerDistance is not null)
        {
            double footerDescent = bottomMargin - footerDistance.Value;
            if (footerDescent >= 0)
            {
                options.Add($"footer-descent: {FormatIn(footerDescent)}");
            }
        }

        return options;
    }

    private string? AddHeaderFooterPageOptions(List<string> pageOptions, TypstHeaderFooterSet headerFooter, TypstPageSetup pageSetup)
    {
        (string? header, string? headerAnchored) = BuildHeaderFooterContent(headerFooter.FirstHeaderBlocks, headerFooter.DefaultHeaderBlocks, headerFooter.HasTitlePage, isHeader: true, pageSetup);
        (string? footer, string? footerAnchored) = BuildHeaderFooterContent(headerFooter.FirstFooterBlocks, headerFooter.DefaultFooterBlocks, headerFooter.HasTitlePage, isHeader: false, pageSetup);
        if (header is not null)
        {
            pageOptions.Add($"header: [{header}]");
        }

        if (footer is not null)
        {
            pageOptions.Add($"footer: [{footer}]");
        }

        // Decorative images (e.g. a bottom chevron anchored to the page bottom) are
        // emitted as separate `#place(bottom + center, ...)` lines AFTER the
        // `#set page(...)` call, so they are not constrained by the small footer area.
        return CombineAnchored(headerAnchored, footerAnchored);
    }

    private static string? CombineAnchored(string? headerAnchored, string? footerAnchored)
    {
        List<string> parts = [];
        if (headerAnchored is not null)
        {
            parts.Add(headerAnchored);
        }

        if (footerAnchored is not null)
        {
            parts.Add(footerAnchored);
        }

        return parts.Count == 0 ? null : string.Join("\n", parts);
    }

    private (string? Content, string? Decorative) BuildHeaderFooterContent(List<TypstBlock> firstBlocks, List<TypstBlock> defaultBlocks, bool hasTitlePage, bool isHeader, TypstPageSetup pageSetup)
    {
        if (!hasTitlePage)
        {
            (List<TypstBlock> nonDecorative, List<TypstImageBlock> decorativeImages) = SplitDecorativeBlocks(defaultBlocks);
            string renderedContent = RenderBlocksAsContent(nonDecorative);
            string? nonTitleContent = string.IsNullOrWhiteSpace(renderedContent) ? null : renderedContent;
            string? nonTitleDecorative = RenderAnchoredImages(decorativeImages);
            return (WrapHeaderFooterContent(nonTitleContent, isHeader, pageSetup), nonTitleDecorative);
        }

        (List<TypstBlock> firstNonDecorative, List<TypstImageBlock> firstDecorativeImages) = SplitDecorativeBlocks(firstBlocks);
        (List<TypstBlock> defaultNonDecorative, List<TypstImageBlock> defaultDecorativeImages) = SplitDecorativeBlocks(defaultBlocks);
        string first = RenderBlocksAsContent(firstNonDecorative);
        string defaultContent = RenderBlocksAsContent(defaultNonDecorative);
        string? firstDecorative = RenderAnchoredImages(firstDecorativeImages);
        string? defaultDecorative = RenderAnchoredImages(defaultDecorativeImages);
        string? resolvedDecorative = firstDecorative ?? defaultDecorative;
        if (string.IsNullOrWhiteSpace(first))
        {
            return (null, resolvedDecorative);
        }

        string? resolvedContent;
        if (string.IsNullOrWhiteSpace(defaultContent))
        {
            resolvedContent = first;
        }
        else
        {
            resolvedContent = $"#context if counter(page).get().first() == 1 [{first}] else [{defaultContent}]";
        }

        return (WrapHeaderFooterContent(resolvedContent, isHeader, pageSetup), resolvedDecorative);
    }

    private string? WrapHeaderFooterContent(string? content, bool isHeader, TypstPageSetup pageSetup)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return content;
        }

        if (isHeader && pageSetup.HeaderDistanceInches is double headerDistance)
        {
            return $"#place(top, dy: {FormatIn(headerDistance)})[#block(width: 100%)[{content}]]";
        }

        if (!isHeader && pageSetup.FooterDistanceInches is double footerDistance)
        {
            return $"#place(bottom, dy: -{FormatIn(footerDistance)})[#block(width: 100%)[{content}]]";
        }

        return content;
    }

    // Splits header/footer blocks into (content, decorative). Decorative blocks are
    // anchored images that should be placed at the page bottom (e.g. a chevron in a
    // small footer area). Inline images remain in the content flow so they keep their
    // original vertical order relative to surrounding paragraphs. A single-block
    // header/footer is never split, so a header that contains only an image still
    // renders the image inside `header: [...]`.
    private static (List<TypstBlock> Content, List<TypstImageBlock> Decorative) SplitDecorativeBlocks(List<TypstBlock> blocks)
    {
        if (blocks.Count <= 1)
        {
            return (blocks, []);
        }

        List<TypstBlock> content = [];
        List<TypstImageBlock> decorative = [];
        foreach (TypstBlock block in blocks)
        {
            if (block is TypstParagraphBlock { Inlines.Count: 0 } paragraph && paragraph.ImageBlocks.Count > 0)
            {
                List<TypstImageBlock> inlineImages = [];
                foreach (TypstImageBlock image in paragraph.ImageBlocks)
                {
                    if (image.XPt is not null || image.YPt is not null)
                    {
                        decorative.Add(image);
                    }
                    else
                    {
                        inlineImages.Add(image);
                    }
                }

                if (inlineImages.Count > 0)
                {
                    content.Add(paragraph with { ImageBlocks = inlineImages });
                }
            }
            else if (block is TypstImageBlock image && (image.XPt is not null || image.YPt is not null))
            {
                decorative.Add(image);
            }
            else
            {
                content.Add(block);
            }
        }

        return (content, decorative);
    }

    private string? RenderAnchoredImages(List<TypstImageBlock> images)
    {
        if (images.Count == 0)
        {
            return null;
        }

        return string.Concat(images.Select(image =>
        {
            string line = image.TopBorder is not null
                ? $"#line(length: 100%, stroke: {FormatPt(image.TopBorder.SizeEighthPoints / 8.0)} + rgb(\"#{image.TopBorder.Color}\")) "
                : string.Empty;
            return $"#place(bottom + center, [{line}{RenderImage(image)}])";
        }));
    }

    private string RenderShape(TypstShapeBlock shape)
    {
        List<string> arguments = [];
        if (shape.WidthPt is > 0)
        {
            arguments.Add($"width: {FormatPt(shape.WidthPt.Value)}");
        }

        if (shape.HeightPt is > 0)
        {
            arguments.Add($"height: {FormatPt(shape.HeightPt.Value)}");
        }

        if (shape.FillColor is not null)
        {
            arguments.Add($"fill: rgb(\"#{shape.FillColor}\")");
        }

        arguments.Add(shape.StrokeColor is null ? "stroke: none" : $"stroke: rgb(\"#{shape.StrokeColor}\")");
        if (shape.Paragraphs.Count > 0)
        {
            arguments.Add("inset: 4pt");
        }

        string content = shape.Paragraphs.Count == 0
            ? string.Empty
            : $"[{string.Join("\n", shape.Paragraphs.Select(RenderParagraph))}]";
        string rectangle = $"#rect({string.Join(", ", arguments)}){content}";

        return shape.XPt is not null || shape.YPt is not null
            ? $"#place(dx: {FormatPt(shape.XPt ?? 0)}, dy: {FormatPt(shape.YPt ?? 0)})[{rectangle}]"
            : rectangle;
    }

    private string RenderInlines(IEnumerable<TypstInline> inlines) => string.Concat(inlines.Select(RenderInline));

    private string RenderInline(TypstInline inline)
    {
        if (inline.Kind == TypstInlineKind.LineBreak)
        {
            return "#linebreak()";
        }

        if (inline.Kind == TypstInlineKind.Tab)
        {
            return "#h(1em)";
        }

        if (inline.Kind == TypstInlineKind.RawTypst)
        {
            string rawContent = inline.RawTypst;
            if (inline.Bold)
            {
                rawContent = $"#strong[{rawContent}]";
            }

            if (inline.Italic)
            {
                rawContent = $"#emph[{rawContent}]";
            }

            if (inline.Underline)
            {
                rawContent = $"#underline[{rawContent}]";
            }

            List<string> rawTextOptions = [];
            if (inline.Color is not null)
            {
                rawTextOptions.Add($"fill: rgb(\"#{inline.Color}\")");
            }

            if (inline.FontSizePt is > 0)
            {
                rawTextOptions.Add($"size: {FormatPt(inline.FontSizePt.Value)}");
            }

            if (!string.IsNullOrWhiteSpace(inline.FontFamily))
            {
                rawTextOptions.Add($"font: {TypstFontValue(inline.FontFamily)}");
            }

            return rawTextOptions.Count == 0 ? rawContent : $"#text({string.Join(", ", rawTextOptions)})[{rawContent}]";
        }

        string content = EscapeTypstContent(inline.Text);
        if (inline.Bold)
        {
            content = $"#strong[{content}]";
        }

        if (inline.Italic)
        {
            content = $"#emph[{content}]";
        }

        if (inline.Underline)
        {
            content = $"#underline[{content}]";
        }

        List<string> textOptions = [];
        if (inline.Color is not null)
        {
            textOptions.Add($"fill: rgb(\"#{inline.Color}\")");
        }

        if (inline.FontSizePt is > 0)
        {
            textOptions.Add($"size: {FormatPt(inline.FontSizePt.Value)}");
        }

        if (!string.IsNullOrWhiteSpace(inline.FontFamily))
        {
            textOptions.Add($"font: {TypstFontValue(inline.FontFamily)}");
        }

        return textOptions.Count == 0 ? content : $"#text({string.Join(", ", textOptions)})[{content}]";
    }

    private static T? GetChild<T>(OpenXmlElement? element)
        where T : OpenXmlElement
    {
        return element?.GetFirstChild<T>();
    }

    private string? GetFontFamily(OpenXmlElement? element)
    {
        RunFonts? fonts = GetChild<RunFonts>(element);
        string? direct = fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value ?? fonts?.ComplexScript?.Value;
        if (!string.IsNullOrWhiteSpace(direct))
        {
            return direct;
        }

        string? theme = fonts?.AsciiTheme?.Value.ToString() ?? fonts?.HighAnsiTheme?.Value.ToString() ?? fonts?.ComplexScriptTheme?.Value.ToString();
        return ResolveThemeFont(theme);
    }

    private void LoadTheme()
    {
        A.Theme? theme = document.MainDocumentPart?.ThemePart?.Theme;
        A.ThemeElements? elements = theme?.ThemeElements;
        A.ColorScheme? colors = elements?.ColorScheme;
        AddThemeColor("dark1", colors?.Dark1Color);
        AddThemeColor("light1", colors?.Light1Color);
        AddThemeColor("dark2", colors?.Dark2Color);
        AddThemeColor("light2", colors?.Light2Color);
        AddThemeColor("accent1", colors?.Accent1Color);
        AddThemeColor("accent2", colors?.Accent2Color);
        AddThemeColor("accent3", colors?.Accent3Color);
        AddThemeColor("accent4", colors?.Accent4Color);
        AddThemeColor("accent5", colors?.Accent5Color);
        AddThemeColor("accent6", colors?.Accent6Color);
        AddThemeColor("hyperlink", colors?.Hyperlink);
        AddThemeColor("followedHyperlink", colors?.FollowedHyperlinkColor);

        majorThemeFont = elements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
        minorThemeFont = elements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
    }

    private void AddThemeColor(string key, OpenXmlCompositeElement? color)
    {
        string? value = color?.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
        value ??= color?.Descendants<A.SystemColor>().FirstOrDefault()?.LastColor?.Value;
        value = NormalizeColor(value);
        if (value is not null)
        {
            themeColors[key] = value;
        }
    }

    private void LoadStyles()
    {
        Styles? styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles is null)
        {
            return;
        }

        foreach (Style style in styles.Elements<Style>())
        {
            string? id = style.StyleId?.Value ?? GetXmlAttribute(style, "styleId");
            if (!string.IsNullOrWhiteSpace(id))
            {
                stylesById[id] = style;
            }
        }
    }

    private void LoadNumbering()
    {
        Numbering? numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
        {
            return;
        }

        Dictionary<string, bool> abstractOrdered = [];
        foreach (AbstractNum abstractNum in numbering.Elements<AbstractNum>())
        {
            string? abstractId = abstractNum.AbstractNumberId?.Value.ToString(CultureInfo.InvariantCulture);
            NumberFormatValues? format = abstractNum.Descendants<NumberingFormat>().FirstOrDefault()?.Val?.Value;
            if (abstractId is not null)
            {
                abstractOrdered[abstractId] = format != NumberFormatValues.Bullet;
            }
        }

        foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>())
        {
            string? numId = instance.NumberID?.Value.ToString(CultureInfo.InvariantCulture);
            string? abstractId = instance.AbstractNumId?.Val?.Value.ToString(CultureInfo.InvariantCulture);
            if (numId is not null && abstractId is not null && abstractOrdered.TryGetValue(abstractId, out bool ordered))
            {
                orderedNumberingByNumId[numId] = ordered;
            }
        }
    }

    private List<OpenXmlElement?> GetRunStyleProperties(Run run)
    {
        string? styleId = run.RunProperties?.RunStyle?.Val?.Value;
        return GetStyleChain(styleId).Select(s => (OpenXmlElement?)s.StyleRunProperties).ToList();
    }

    private List<OpenXmlElement?> GetParagraphStyleRunProperties(Wp.Paragraph paragraph)
    {
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return GetStyleChain(styleId).Select(s => (OpenXmlElement?)s.StyleRunProperties).ToList();
    }

    private List<OpenXmlElement?> GetParagraphStyleParagraphProperties(Wp.Paragraph paragraph)
    {
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return GetStyleChain(styleId).Select(s => (OpenXmlElement?)s.StyleParagraphProperties).ToList();
    }

    private List<OpenXmlElement?> GetParagraphFormattingProperties(Wp.Paragraph paragraph)
    {
        List<OpenXmlElement?> properties =
        [
            document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle
        ];
        properties.AddRange(GetParagraphStyleParagraphProperties(paragraph));
        return properties;
    }

    private List<Style> GetStyleChain(string? styleId)
    {
        List<Style> chain = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        while (!string.IsNullOrWhiteSpace(styleId) && seen.Add(styleId) && stylesById.TryGetValue(styleId, out Style? style))
        {
            chain.Insert(0, style);
            styleId = style.BasedOn?.Val?.Value;
        }

        return chain;
    }

    private Style? GetTableStyle(Wp.Table table)
    {
        TableStyle? tableStyle = table.TableProperties?.GetFirstChild<TableStyle>() ?? table.Descendants<TableStyle>().FirstOrDefault();
        string? styleId = tableStyle?.Val?.Value ?? GetXmlAttribute(tableStyle, "val");
        return styleId is not null && stylesById.TryGetValue(styleId, out Style? style) ? style : null;
    }

    private TableStyleProperties? GetTableStyleProperties(Wp.Table table, Wp.TableRow row, bool firstRow)
    {
        if (!firstRow && row.TableRowProperties?.GetFirstChild<TableHeader>() is null)
        {
            return null;
        }

        return GetTableStyle(table)?.Descendants<TableStyleProperties>()
            .FirstOrDefault(p => p.Type?.Value == TableStyleOverrideValues.FirstRow || p.Type?.Value == TableStyleOverrideValues.Band1Horizontal);
    }

    private TypstInline ApplyConditionalRunFormatting(TypstInline inline, OpenXmlElement properties)
    {
        if (inline.Kind != TypstInlineKind.Text)
        {
            return inline;
        }

        RunFormatting formatting = new()
        {
            Bold = inline.Bold,
            Italic = inline.Italic,
            Underline = inline.Underline,
            Color = inline.Color,
            FontSizePt = inline.FontSizePt,
            FontFamily = inline.FontFamily
        };
        ApplyRunFormatting(formatting, properties);
        return inline with
        {
            Bold = formatting.Bold,
            Italic = formatting.Italic,
            Underline = formatting.Underline,
            Color = formatting.Color,
            FontSizePt = formatting.FontSizePt,
            FontFamily = formatting.FontFamily
        };
    }

    private TypstInline ApplyTableFirstRowFormatting(TypstInline inline, TableFirstRowFormatting formatting)
    {
        if (inline.Kind != TypstInlineKind.Text)
        {
            return inline;
        }

        return inline with
        {
            Bold = formatting.Bold || inline.Bold,
            Color = formatting.TextColor ?? inline.Color
        };
    }

    private TableFirstRowFormatting? GetTableFirstRowFormatting(Wp.Table table, Wp.TableRow row, bool firstRow)
    {
        if (!firstRow && row.TableRowProperties?.GetFirstChild<TableHeader>() is null)
        {
            return null;
        }

        Style? style = GetTableStyle(table);
        if (style is null)
        {
            return null;
        }

        string xml = style.OuterXml;
        Match match = Regex.Match(xml, "<w:tblStylePr[^>]*w:type=\"(?:firstRow|band1Horz)\"[\\s\\S]*?</w:tblStylePr>");
        if (!match.Success)
        {
            return null;
        }

        string block = match.Value;
        string? shading = NormalizeColor(Regex.Match(block, "w:fill=\"([0-9A-Fa-f]{6})\"").Groups[1].Value);
        string? textColor = NormalizeColor(Regex.Match(block, "<w:color[^>]*w:val=\"([0-9A-Fa-f]{6})\"").Groups[1].Value);
        bool bold = Regex.IsMatch(block, "<w:b(?:\\s|/|>)");
        return shading is null && textColor is null && !bold ? null : new TableFirstRowFormatting(shading, bold, textColor);
    }

    private void ApplyRunFormatting(RunFormatting formatting, OpenXmlElement? properties)
    {
        if (properties is null)
        {
            return;
        }

        Bold? bold = GetChild<Bold>(properties);
        if (bold is not null)
        {
            formatting.Bold = IsOn(bold);
        }

        Italic? italic = GetChild<Italic>(properties);
        if (italic is not null)
        {
            formatting.Italic = IsOn(italic);
        }

        Underline? underline = GetChild<Underline>(properties);
        if (underline is not null)
        {
            formatting.Underline = IsUnderlineOn(underline);
        }

        Color? color = GetChild<Color>(properties);
        formatting.Color = ResolveColor(color) ?? formatting.Color;
        formatting.FontSizePt = HalfPointsToPoints(GetChild<FontSize>(properties)?.Val?.Value) ?? formatting.FontSizePt;
        formatting.FontFamily = GetFontFamily(properties) ?? formatting.FontFamily;
    }

    private string ExtractDefaultFontFamily()
    {
        RunPropertiesBaseStyle? defaults = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        return GetFontFamily(defaults)
            ?? GetFontFamily(stylesById.TryGetValue("Normal", out Style? normal) ? normal.StyleRunProperties : null)
            ?? "Liberation Serif";
    }

    private double ExtractDefaultFontSize()
    {
        string? size = document.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.FontSize?.Val?.Value;
        return HalfPointsToPoints(size) ?? 11;
    }

    private bool IsListParagraph(Wp.Paragraph paragraph) => paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val is not null;

    private bool IsOrderedList(Wp.Paragraph paragraph)
    {
        string? numId = paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value.ToString(CultureInfo.InvariantCulture);
        return numId is null || !orderedNumberingByNumId.TryGetValue(numId, out bool ordered) || ordered;
    }

    private static bool HasPageBreak(Wp.Paragraph paragraph) => paragraph.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);

    private static bool IsOn(OnOffType? value) => value is not null && (value.Val is null || value.Val.Value);

    private static bool IsUnderlineOn(Underline? underline) => underline?.Val is not null && underline.Val.Value != UnderlineValues.None;

    private static double TwipsToInches(long? value, double fallback) => value.HasValue ? value.Value / TwipsPerInch : fallback;

    private static double? TwipsToPoints(string? value) => double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double twips) ? twips / 20.0 : null;

    private double? ResolveAutoLineSpacingLeading(SpacingBetweenLines? spacing, IReadOnlyList<TypstInline> inlines)
    {
        if (spacing?.Line?.Value is null)
        {
            return null;
        }

        LineSpacingRuleValues? rule = spacing.LineRule?.Value;
        if (rule is not null && rule != LineSpacingRuleValues.Auto)
        {
            return null;
        }

        double? lineHeightPt = TwipsToPoints(spacing.Line.Value);
        if (lineHeightPt is null)
        {
            return null;
        }

        double fontSizePt = inlines.Select(i => i.FontSizePt).FirstOrDefault(s => s is > 0) ?? ExtractDefaultFontSize();
        double leadingPt = lineHeightPt.Value - fontSizePt;
        return leadingPt > 0 ? Math.Min(leadingPt, 6.0) : null;
    }

    private string? ResolveColor(Color? color) => NormalizeColor(color?.Val?.Value) ?? ResolveThemeColor(color?.ThemeColor?.Value.ToString()) ?? ResolveThemeColor(GetXmlAttribute(color, "themeColor"));

    private string? ResolveShadingColor(Shading? shading) => NormalizeColor(shading?.Fill?.Value) ?? ResolveThemeColor(shading?.ThemeFill?.Value.ToString()) ?? ResolveThemeColor(shading?.ThemeColor?.Value.ToString()) ?? ResolveThemeColor(GetXmlAttribute(shading, "themeFill")) ?? ResolveThemeColor(GetXmlAttribute(shading, "themeColor"));

    private string? ResolveThemeColor(string? themeColor)
    {
        if (string.IsNullOrWhiteSpace(themeColor))
        {
            return null;
        }

        return themeColors.TryGetValue(themeColor, out string? color) ? color : null;
    }

    private string? ResolveThemeFont(string? themeFont)
    {
        if (string.IsNullOrWhiteSpace(themeFont))
        {
            return null;
        }

        return themeFont.Contains("major", StringComparison.OrdinalIgnoreCase) ? majorThemeFont : minorThemeFont;
    }

    private static string? GetXmlAttribute(OpenXmlElement? element, string localName)
    {
        if (element is null)
        {
            return null;
        }

        Match match = Regex.Match(element.OuterXml, $"\\b(?:\\w+:)?{Regex.Escape(localName)}=\"([^\"]+)\"");
        return match.Success ? match.Groups[1].Value : null;
    }

    // Defensive fallback for <w:spacing> attributes. The OpenXml SDK's typed
    // `SpacingBetweenLines.Before`/`After` properties can misread values when
    // `w:afterAutospacing="1"` is present on the same element (returning the
    // `after` value for `before`). Reading the raw XML attribute avoids that
    // SDK bug and gives the true stored value.
    private static string? GetSpacingAttributeValue(OpenXmlElement? element, string localName)
    {
        if (element is null)
        {
            return null;
        }

        Match match = Regex.Match(element.OuterXml, $"\\bw:{Regex.Escape(localName)}=\"([^\"]+)\"");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string? MapAlignment(JustificationValues value)
    {
        if (value == JustificationValues.Center)
        {
            return "center";
        }

        if (value == JustificationValues.Right)
        {
            return "right";
        }

        return value == JustificationValues.Both ? "justify" : null;
    }

    private static double? HalfPointsToPoints(string? value) => double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double halfPoints) ? halfPoints / 2.0 : null;

    private sealed class RunFormatting
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string? Color { get; set; }
        public double? FontSizePt { get; set; }
        public string? FontFamily { get; set; }
    }

    private static void MergeFieldFormatting(RunFormatting target, RunFormatting source)
    {
        if (source.Bold)
        {
            target.Bold = true;
        }

        if (source.Italic)
        {
            target.Italic = true;
        }

        if (source.Underline)
        {
            target.Underline = true;
        }

        if (source.Color is not null)
        {
            target.Color = source.Color;
        }

        if (source.FontSizePt is not null)
        {
            target.FontSizePt = source.FontSizePt;
        }

        if (!string.IsNullOrWhiteSpace(source.FontFamily))
        {
            target.FontFamily = source.FontFamily;
        }
    }

    private sealed class FieldParseState
    {
        public bool InField { get; set; }
        public bool SuppressResult { get; set; }
        public StringBuilder Instruction { get; } = new();
        public RunFormatting FieldFormatting { get; set; } = new();
    }

    private sealed record TableFirstRowFormatting(string? ShadingColor, bool Bold, string? TextColor);

    private sealed record ShapeGeometry(double? XPt, double? YPt, double? WidthPt, double? HeightPt, string? FillColor, string? StrokeColor)
    {
        public bool IsValid { get; init; } = true;
    }

    private sealed record ExtractedImage(string TypstPath, bool IsUnsupportedFormat);

    private static string? NormalizeColor(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        string color = value.Trim().TrimStart('#');
        return color.Length == 6 && color.All(Uri.IsHexDigit) ? color.ToUpperInvariant() : null;
    }

    private static string NormalizeSymbolText(string text, string? fontFamily)
    {
        if (string.IsNullOrEmpty(text) || !IsSymbolFont(fontFamily))
        {
            return text;
        }

        return string.Concat(text.Select(ch => MapSymbolCharacter(fontFamily!, ((int)ch).ToString("X4", CultureInfo.InvariantCulture)) ?? ch.ToString()));
    }

    private static bool IsSymbolFont(string? fontFamily)
        => !string.IsNullOrWhiteSpace(fontFamily)
            && (fontFamily.Contains("Wingdings", StringComparison.OrdinalIgnoreCase)
                || fontFamily.Contains("Symbol", StringComparison.OrdinalIgnoreCase)
                || fontFamily.Contains("Segoe UI Symbol", StringComparison.OrdinalIgnoreCase));

    private static string? MapSymbolCharacter(string fontFamily, string hex)
    {
        string normalized = hex.Trim().TrimStart('x').ToUpperInvariant();
        if (normalized.StartsWith("F0", StringComparison.Ordinal) && normalized.Length >= 4)
        {
            normalized = normalized[^2..];
        }
        else if (normalized.Length == 4 && normalized.StartsWith("00", StringComparison.Ordinal))
        {
            normalized = normalized[2..];
        }

        if (fontFamily.Contains("Wingdings", StringComparison.OrdinalIgnoreCase))
        {
            return normalized switch
            {
                "A7" or "6C" => "•",
                "9F" or "71" => "▪",
                "A8" or "6E" => "◦",
                "A3" or "6F" or "6D" => "☐",
                "FE" or "FC" or "52" => "☑",
                "FD" or "FB" => "☒",
                _ => null
            };
        }

        if (fontFamily.Contains("Symbol", StringComparison.OrdinalIgnoreCase))
        {
            return normalized switch
            {
                "B7" or "95" => "•",
                "A7" => "§",
                "D8" => "×",
                "E0" => "◊",
                _ => null
            };
        }

        return normalized switch
        {
            "2610" => "☐",
            "2611" => "☑",
            "2612" => "☒",
            _ => null
        };
    }

    private static TypstInline? CreateCheckboxInline(string text)
    {
        string normalized = text.Trim();
        return normalized switch
        {
            "☐" => new TypstInline { Kind = TypstInlineKind.RawTypst, RawTypst = "#box(width: 8pt, height: 8pt, stroke: 0.7pt)" },
            "☑" => new TypstInline { Kind = TypstInlineKind.RawTypst, RawTypst = "#box(width: 8pt, height: 8pt, stroke: 0.7pt)[#place(dx: 1pt, dy: -2pt)[#text(size: 9pt)[✓]]]" },
            "☒" => new TypstInline { Kind = TypstInlineKind.RawTypst, RawTypst = "#box(width: 8pt, height: 8pt, stroke: 0.7pt)[#place(dx: 1pt, dy: -2pt)[#text(size: 9pt)[×]]]" },
            _ => null
        };
    }

    private static string ContentTypeToExtension(string contentType) => contentType.ToLowerInvariant() switch
    {
        "image/png" => ".png",
        "image/jpeg" or "image/jpg" => ".jpg",
        "image/gif" => ".gif",
        "image/bmp" => ".bmp",
        "image/tiff" => ".tiff",
        "image/x-emf" or "image/emf" => ".emf",
        "image/x-wmf" or "image/wmf" => ".wmf",
        "image/svg+xml" => ".svg",
        _ => ".bin"
    };

    private static bool IsUnsupportedImagePath(string path)
        => IsUnsupportedImageExtension(Path.GetExtension(path));

    private static bool IsUnsupportedImageExtension(string extension)
        => extension.Equals(".emf", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".wmf", StringComparison.OrdinalIgnoreCase);

    private static string EscapeTypstContent(string value)
    {
        return value
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("#", "\\#", StringComparison.Ordinal)
            .Replace("$", "\\$", StringComparison.Ordinal)
            .Replace("%", "\\%", StringComparison.Ordinal)
            .Replace("&", "\\&", StringComparison.Ordinal)
            .Replace("@", "\\@", StringComparison.Ordinal)
            .Replace("*", "\\*", StringComparison.Ordinal)
            .Replace("_", "\\_", StringComparison.Ordinal)
            .Replace("^", "\\^", StringComparison.Ordinal)
            .Replace("~", "\\~", StringComparison.Ordinal)
            .Replace("`", "\\`", StringComparison.Ordinal)
            .Replace("<", "\\<", StringComparison.Ordinal)
            .Replace(">", "\\>", StringComparison.Ordinal)
            .Replace("[", "\\[", StringComparison.Ordinal)
            .Replace("]", "\\]", StringComparison.Ordinal);
    }

    private static string TypstString(string value) => "\"" + value.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";

    private static string TypstFontValue(string value)
    {
        string[] fallback = GetFontFallback(value);
        return fallback.Length == 0
            ? TypstString(value)
            : TypstString(fallback[0]);
    }

    private static string[] GetFontFallback(string value)
    {
        if (value.Equals("Corbel", StringComparison.OrdinalIgnoreCase)
            || value.Equals("Calibri", StringComparison.OrdinalIgnoreCase)
            || value.Equals("Segoe UI", StringComparison.OrdinalIgnoreCase)
            || value.Equals("Arial", StringComparison.OrdinalIgnoreCase))
        {
            return ["Liberation Sans", "Noto Sans", "Aptos"];
        }

        return [];
    }

    private static string FormatIn(double value) => value.ToString("0.###", CultureInfo.InvariantCulture) + "in";

    private static string FormatPt(double value) => value.ToString("0.###", CultureInfo.InvariantCulture) + "pt";
}
