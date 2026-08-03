using System.Globalization;
using System.Text.RegularExpressions;
using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Schema;

/// <summary>
/// Validates the semantic invariants of an already materialized DOCX generation model.
/// JSON shape, value-kind, unknown-property, and typo diagnostics remain the parser's concern.
/// </summary>
internal static class DocxGenerationModelValidator
{
    private static readonly Regex HexColorPattern = new(
        @"^#[0-9a-fA-F]{6}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static DocxGenerationValidationResult Validate(DocxGenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        ValidationSession session = new(document);
        return session.Validate();
    }

    private sealed class ValidationSession(DocxGenerationDocument document)
    {
        private readonly List<DocxGenerationIssue> _errors = [];
        private readonly List<DocxGenerationIssue> _warnings = [];
        private IReadOnlyDictionary<string, string> _palette = DesignThemeCatalog.Editorial.Palette;
        private IReadOnlyDictionary<string, TypographyToken> _typography = new Dictionary<string, TypographyToken>();
        private FontTokens _fonts = new();
        private PageDefaults _pageDefaults = new();
        private string? _themeName;

        public DocxGenerationValidationResult Validate()
        {
            ValidateRoot();
            return new DocxGenerationValidationResult(
                _errors.Count == 0 ? document : null,
                [.. _errors],
                [.. _warnings]);
        }

        private void ValidateRoot()
        {
            if (document.Version is null)
            {
                Error("$.version", $"'version' is required (\"{DocxGenerationDocument.SupportedVersion}\").");
            }
            else if (!string.Equals(document.Version, DocxGenerationDocument.SupportedVersion, StringComparison.Ordinal))
            {
                Error("$.version", $"unsupported version '{document.Version}'; expected \"{DocxGenerationDocument.SupportedVersion}\".");
            }

            if (document.TemplatePath is not null && string.IsNullOrWhiteSpace(document.TemplatePath))
            {
                Error("$.template", "must not be empty or whitespace (omit it to generate from a blank document).");
            }

            if (document.Design is not null)
            {
                ValidateDesign(document.Design);
            }

            IReadOnlyList<Section>? sections = document.Sections;
            if (sections is null || sections.Count == 0)
            {
                Error("$.sections", "at least one section is required.");
                return;
            }

            for (var i = 0; i < sections.Count; i++)
            {
                Section? section = sections[i];
                string path = $"$.sections[{i}]";
                if (section is null)
                {
                    Error(path, "each section must be a section object and must not be null.");
                    continue;
                }

                ValidateSection(section, path, i == 0);
            }
        }

        private void ValidateDesign(DesignTokens design)
        {
            string? themeName = design.Theme;
            if (themeName is not null && !DesignThemeCatalog.Themes.ContainsKey(themeName))
            {
                Error("$.design.theme", $"unknown theme '{themeName}'. Known themes: {string.Join(", ", DesignThemeCatalog.Themes.Keys)}.", Suggest(themeName, DesignThemeCatalog.Themes.Keys));
            }
            _themeName = themeName;

            IReadOnlyDictionary<string, string>? palette = design.Palette;
            // A null/omitted palette is valid: the active theme's palette applies as the base.
            _palette = DesignThemeCatalog.Resolve(_themeName).EffectivePalette(palette);
            if (palette is not null)
            {
                foreach (var (name, color) in palette)
                {
                    string path = $"$.design.palette.{name}";
                    if (color is null || !HexColorPattern.IsMatch(color))
                    {
                        Error(path, "palette colors must be #RRGGBB hex literals.");
                    }
                }
            }

            FontTokens? fonts = design.Fonts;
            if (fonts is not null)
            {
                _fonts = fonts;
            }

            IReadOnlyDictionary<string, TypographyToken>? typography = design.Typography;
            if (typography is null)
            {
                Error("$.design.typography", "must not be null.");
            }
            else
            {
                _typography = typography;
                foreach (var (name, token) in typography)
                {
                    string path = $"$.design.typography.{name}";
                    if (token is null)
                    {
                        Error(path, "typography token must not be null.");
                        continue;
                    }

                    CheckFontReference(token.FontFamily, $"{path}.font");
                    CheckOptionalPositive(token.SizePt, $"{path}.size");
                    CheckColor(token.Color, $"{path}.color");
                }
            }

            IReadOnlyDictionary<string, double>? spacing = design.Spacing;
            if (spacing is null)
            {
                Error("$.design.spacing", "must not be null.");
            }
            else
            {
                foreach (var (name, value) in spacing)
                {
                    CheckNonNegative(value, $"$.design.spacing.{name}");
                }
            }

            ShapeDefaults? shapes = design.Shapes;
            if (shapes is not null)
            {
                CheckNonNegative(shapes.CornerRadiusPt, "$.design.shapes.cornerRadius");
                CheckColor(shapes.DefaultFill, "$.design.shapes.defaultFill");
                CheckColor(shapes.DefaultStrokeColor, "$.design.shapes.defaultStroke");
                CheckPositive(shapes.DefaultStrokeWidthPt, "$.design.shapes.defaultStrokeWidth");
            }

            PageDefaults? page = design.Page;
            if (page is null)
            {
                Error("$.design.page", "must not be null.");
            }
            else
            {
                _pageDefaults = page;
                CheckNullableEnum(page.PageSize, "$.design.page.size", "page size");
                CheckEnum(page.Orientation, "$.design.page.orientation", "orientation");
                if (page.Margins is not null)
                {
                    ValidateMargins(page.Margins, "$.design.page.margins");
                }
                CheckFontReference(page.DefaultFontFamily, "$.design.page.defaultFont");
                CheckColor(page.DefaultTextColor, "$.design.page.defaultTextColor");
            }

            LayoutDefaults? layout = design.Layout;
            if (layout is not null)
            {
                CheckNullableEnum(layout.Density, "$.design.layout.density", "density");
                CheckOptionalNonNegative(layout.MinBodySizePt, "$.design.layout.minBodySizePt");
                CheckOptionalPositive(layout.MaxTableWidthPt, "$.design.layout.maxTableWidthPt");
            }
        }

        private void ValidateSection(Section section, string path, bool isFirst)
        {
            if (section.PageSetup is not null)
            {
                ValidatePageSetup(section.PageSetup, $"{path}.pageSetup");
                if (isFirst && section.PageSetup.BreakType is not null)
                {
                    Warn($"{path}.pageSetup.breakType", "the first section has no preceding section; 'breakType' is ignored.");
                }
            }

            ValidateFlowList(section.Header, $"{path}.header");
            ValidateFlowList(section.Footer, $"{path}.footer");
            ValidateFlowList(section.Blocks, $"{path}.blocks", "a section must contain at least one flow block.");
            ValidatePositionedList(section.Positioned, $"{path}.positioned");
            ValidateSectionGeometry(section, path);
        }

        private void ValidatePageSetup(PageSetup setup, string path)
        {
            if (setup.PageSize is not null)
            {
                ValidatePageSize(setup.PageSize, $"{path}.size");
            }
            CheckNullableEnum(setup.Orientation, $"{path}.orientation", "orientation");
            if (setup.Margins is not null)
            {
                ValidateMargins(setup.Margins, $"{path}.margins");
            }
            if (setup.Columns is not null)
            {
                ValidateColumns(setup.Columns, $"{path}.columns");
            }
            CheckNullableEnum(setup.BreakType, $"{path}.breakType", "section break type");
        }

        private void ValidatePageSize(PageSize size, string path)
        {
            if (size.Name is not null)
            {
                CheckNullableEnum(size.Name, path, "page size");
                if (size.WidthPt is not null || size.HeightPt is not null)
                {
                    Error(path, "a named page size cannot also declare custom width or height.");
                }
                return;
            }

            if (size.WidthPt is null || size.HeightPt is null)
            {
                Error(path, "a custom page size requires both positive 'width' and 'height'.");
                return;
            }
            CheckPositive(size.WidthPt.Value, $"{path}.width");
            CheckPositive(size.HeightPt.Value, $"{path}.height");
        }

        private void ValidateMargins(Margins margins, string path)
        {
            CheckNonNegative(margins.TopPt, $"{path}.top");
            CheckNonNegative(margins.RightPt, $"{path}.right");
            CheckNonNegative(margins.BottomPt, $"{path}.bottom");
            CheckNonNegative(margins.LeftPt, $"{path}.left");
        }

        private void ValidateColumns(PageColumns columns, string path)
        {
            if (columns.Count < 1)
            {
                Error($"{path}.count", $"must be between 1 and {int.MaxValue} (got {columns.Count}).");
            }
            CheckNonNegative(columns.SpacingPt, $"{path}.spacing");
            if (columns.Count == 1 && (columns.SpacingPt > 0 || columns.SeparatorLine))
            {
                Warn(path, "'spacing'/'separator' have no effect with a single column (count: 1).");
            }
        }

        private void ValidateFlowList(
            IReadOnlyList<FlowBlock>? blocks,
            string path,
            string? emptyMessage = null)
        {
            if (blocks is null)
            {
                Error(path, "must not be null.");
                return;
            }
            if (emptyMessage is not null && blocks.Count == 0)
            {
                Error(path, emptyMessage);
            }
            for (var i = 0; i < blocks.Count; i++)
            {
                FlowBlock? block = blocks[i];
                string blockPath = $"{path}[{i}]";
                if (block is null)
                {
                    Error(blockPath, "flow block must not be null.");
                    continue;
                }
                ValidateFlowBlock(block, blockPath);
            }
        }

        private void ValidateFlowBlock(FlowBlock block, string path)
        {
            switch (block)
            {
                case ParagraphBlock paragraph:
                    ValidateRequiredText(paragraph.Content, path);
                    break;
                case HeadingBlock heading:
                    if (heading.Level is < 1 or > 6)
                    {
                        Error($"{path}.level", $"must be between 1 and 6 (got {heading.Level}).");
                    }
                    ValidateRequiredText(heading.Content, path);
                    break;
                case ListBlock list:
                    ValidateList(list, path);
                    break;
                case TableBlock table:
                    ValidateTable(table, path);
                    break;
                case ImageElement image:
                    ValidateFlowImage(image, path);
                    break;
                case CalloutBlock callout:
                    CheckEnum(callout.Tone, $"{path}.tone", "callout tone");
                    ValidateRequiredText(callout.Content, path);
                    break;
                case PageBreakBlock:
                    break;
                case FlowContainerBlock group:
                    ValidateFlowList(group.Blocks, $"{path}.blocks", "a group must contain at least one flow block.");
                    break;
                case CoverBlock cover:
                    ValidateOptionalText(cover.Eyebrow, $"{path}.eyebrow");
                    ValidateRequiredText(cover.Title, $"{path}.title");
                    ValidateOptionalText(cover.Subtitle, $"{path}.subtitle");
                    ValidateOptionalText(cover.Metadata, $"{path}.metadata");
                    if (cover.Kpis is not null)
                    {
                        ValidateKpiItems(cover.Kpis, $"{path}.kpis");
                    }
                    break;
                case KpiRowBlock kpiRow:
                    ValidateKpiItems(kpiRow.Items, $"{path}.items");
                    break;
                case SemanticSectionBlock section:
                    ValidateRequiredText(section.Title, $"{path}.title");
                    ValidateOptionalText(section.Intro, $"{path}.intro");
                    ValidateFlowList(section.Blocks, $"{path}.blocks", "a semantic section must contain at least one flow block.");
                    break;
                case ComparisonTableBlock comparison:
                    ValidateComparisonTable(comparison, path);
                    break;
                case RoadmapBlock roadmap:
                    ValidateRoadmap(roadmap, path);
                    break;
                default:
                    Error(path, $"unsupported flow block type '{block.GetType().Name}'.");
                    break;
            }
        }

        private void ValidateList(ListBlock list, string path)
        {
            CheckEnum(list.Kind, $"{path}.kind", "list kind");
            if (list.StartIndex is <= 0)
            {
                Error($"{path}.start", $"must be ≥ 1 (got {list.StartIndex.Value}).");
            }
            if (list.StartIndex is not null && list.Kind == ListKind.Bullet)
            {
                Warn($"{path}.start", "'start' is ignored for bullet lists; use kind 'ordered' for numbered lists.");
            }

            IReadOnlyList<TextModel>? items = list.Items;
            if (items is null || items.Count == 0)
            {
                Error($"{path}.items", "a list must contain at least one item.");
                return;
            }
            for (var i = 0; i < items.Count; i++)
            {
                ValidateRequiredText(items[i], $"{path}.items[{i}]");
            }
        }

        private void ValidateTable(TableBlock table, string path)
        {
            CheckNullableEnum(table.Alignment, $"{path}.alignment", "alignment");
            IReadOnlyList<TableRow>? rows = table.Rows;
            if (rows is null || rows.Count == 0)
            {
                Error($"{path}.rows", "a table must contain at least one row.");
                return;
            }

            int? columnCount = null;
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                TableRow? row = rows[rowIndex];
                string rowPath = $"{path}.rows[{rowIndex}]";
                if (row is null)
                {
                    Error(rowPath, "table row must not be null.");
                    continue;
                }
                IReadOnlyList<TableCell>? cells = row.Cells;
                if (cells is null || cells.Count == 0)
                {
                    Error($"{rowPath}.cells", "a row must contain at least one cell.");
                    continue;
                }

                if (columnCount is null)
                {
                    columnCount = cells.Count;
                }
                else if (cells.Count != columnCount.Value)
                {
                    Error(rowPath, $"row has {cells.Count} cells but the table has {columnCount.Value} columns; all rows must have the same number of cells.");
                }

                for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                {
                    TableCell? cell = cells[cellIndex];
                    string cellPath = $"{rowPath}.cells[{cellIndex}]";
                    if (cell is null)
                    {
                        Error(cellPath, "table cell must not be null.");
                        continue;
                    }
                    ValidateText(cell.Content, cellPath, allowEmpty: true);
                    CheckColor(cell.Fill, $"{cellPath}.fill");
                    CheckNullableEnum(cell.Alignment, $"{cellPath}.alignment", "alignment");
                }
            }

            IReadOnlyList<double>? widths = table.ColumnWidthsPt;
            if (widths is not null)
            {
                for (var i = 0; i < widths.Count; i++)
                {
                    CheckPositive(widths[i], $"{path}.widths[{i}]");
                }
                if (columnCount is not null && widths.Count != columnCount.Value)
                {
                    Error($"{path}.widths", $"got {widths.Count} column widths but the table has {columnCount.Value} columns.");
                }
            }
        }

        private void ValidateFlowImage(ImageElement image, string path)
        {
            CheckRequiredString(image.Source, $"{path}.src", "image source");
            CheckEnum(image.Fit, $"{path}.fit", "image fit mode");
            ValidateCrop(image.Crop, $"{path}.crop");
            CheckOptionalPositive(image.WidthPt, $"{path}.width");
            CheckOptionalPositive(image.HeightPt, $"{path}.height");
        }

        private void ValidateRequiredText(TextModel? content, string path)
        {
            if (content is null)
            {
                Error(path, "text content is required.");
                return;
            }
            ValidateText(content, path, allowEmpty: false);
        }

        private void ValidateOptionalText(TextModel? content, string path)
        {
            if (content is not null)
            {
                ValidateText(content, path, allowEmpty: false);
            }
        }

        private void ValidateKpiItems(IReadOnlyList<KpiItem>? items, string path)
        {
            if (items is null)
            {
                Error(path, "must not be null.");
                return;
            }
            for (var i = 0; i < items.Count; i++)
            {
                KpiItem? item = items[i];
                string itemPath = $"{path}[{i}]";
                if (item is null)
                {
                    Error(itemPath, "KPI item must not be null.");
                    continue;
                }
                ValidateRequiredText(item.Value, $"{itemPath}.value");
                ValidateRequiredText(item.Label, $"{itemPath}.label");
                CheckEnum(item.Tone, $"{itemPath}.tone", "report tone");
            }
        }

        private void ValidateComparisonTable(ComparisonTableBlock table, string path)
        {
            IReadOnlyList<TextModel>? columns = table.Columns;
            if (columns is null || columns.Count == 0)
            {
                Error($"{path}.columns", "a comparison table must have at least one column.");
                return;
            }
            for (var i = 0; i < columns.Count; i++)
            {
                ValidateRequiredText(columns[i], $"{path}.columns[{i}]");
            }

            IReadOnlyList<ComparisonTableRow>? rows = table.Rows;
            if (rows is null || rows.Count == 0)
            {
                Error($"{path}.rows", "a comparison table must contain at least one row.");
                return;
            }
            int columnCount = columns.Count;
            for (var r = 0; r < rows.Count; r++)
            {
                ComparisonTableRow? row = rows[r];
                string rowPath = $"{path}.rows[{r}]";
                if (row is null)
                {
                    Error(rowPath, "comparison table row must not be null.");
                    continue;
                }
                IReadOnlyList<TextModel>? cells = row.Cells;
                if (cells is null || cells.Count == 0)
                {
                    Error($"{rowPath}.cells", "a row must contain at least one cell.");
                    continue;
                }
                if (cells.Count != columnCount)
                {
                    Error(rowPath, $"row has {cells.Count} cells but the comparison table has {columnCount} columns; all rows must have the same number of cells.");
                }
                for (var c = 0; c < cells.Count; c++)
                {
                    ValidateRequiredText(cells[c], $"{rowPath}.cells[{c}]");
                }
            }
        }

        private void ValidateRoadmap(RoadmapBlock roadmap, string path)
        {
            IReadOnlyList<RoadmapPhase>? phases = roadmap.Phases;
            if (phases is null || phases.Count == 0)
            {
                Error($"{path}.phases", "a roadmap must contain at least one phase.");
                return;
            }
            for (var i = 0; i < phases.Count; i++)
            {
                RoadmapPhase? phase = phases[i];
                string phasePath = $"{path}.phases[{i}]";
                if (phase is null)
                {
                    Error(phasePath, "roadmap phase must not be null.");
                    continue;
                }
                ValidateRequiredText(phase.Window, $"{phasePath}.window");
                ValidateRequiredText(phase.Action, $"{phasePath}.action");
                ValidateOptionalText(phase.Evidence, $"{phasePath}.evidence");
                CheckEnum(phase.Tone, $"{phasePath}.tone", "report tone");
            }
        }

        private void ValidateText(TextModel? content, string path, bool allowEmpty)
        {
            if (content is null)
            {
                if (!allowEmpty)
                {
                    Error(path, "text content is required.");
                }
                return;
            }

            bool hasText = content.Text is not null;
            bool hasRuns = content.Runs is not null;
            if (hasText == hasRuns)
            {
                if (!(allowEmpty && !hasText))
                {
                    Error(path, "text content needs exactly one of 'text' (string) or 'runs' (array).");
                }
            }

            if (content.Runs is { } runs)
            {
                if (runs.Count == 0)
                {
                    Error($"{path}.runs", "'runs' must contain at least one run.");
                }
                for (var i = 0; i < runs.Count; i++)
                {
                    Run? run = runs[i];
                    string runPath = $"{path}.runs[{i}]";
                    if (run is null)
                    {
                        Error(runPath, "run must not be null.");
                        continue;
                    }
                    CheckRequiredString(run.Text, $"{runPath}.text", "run text");
                    CheckFontReference(run.FontFamily, $"{runPath}.font");
                    CheckOptionalPositive(run.FontSizePt, $"{runPath}.size");
                    CheckColor(run.Color, $"{runPath}.color");
                }
            }

            if (content.Token is not null && !_typography.ContainsKey(content.Token))
            {
                Error($"{path}.token", $"unknown typography token '{content.Token}'.", Suggest(content.Token, _typography.Keys));
            }
            CheckNullableEnum(content.Role, $"{path}.role", "text role");
            CheckNullableEnum(content.Alignment, $"{path}.alignment", "alignment");
            if (content.Spacing is not null)
            {
                CheckOptionalNonNegative(content.Spacing.BeforePt, $"{path}.spacing.before");
                CheckOptionalNonNegative(content.Spacing.AfterPt, $"{path}.spacing.after");
                CheckOptionalPositive(content.Spacing.LineMultiple, $"{path}.spacing.line");
            }
        }

        private void ValidatePositionedList(IReadOnlyList<PositionedElement>? elements, string path)
        {
            if (elements is null)
            {
                Error(path, "must not be null.");
                return;
            }
            for (var i = 0; i < elements.Count; i++)
            {
                PositionedElement? element = elements[i];
                string elementPath = $"{path}[{i}]";
                if (element is null)
                {
                    Error(elementPath, "positioned element must not be null.");
                    continue;
                }
                ValidatePositionedElement(element, elementPath);
            }
        }

        private void ValidatePositionedElement(PositionedElement element, string path)
        {
            PositionSpec? position = element.Position;
            if (position is null)
            {
                Error(path, "positioned element position is required.");
                return;
            }

            ValidatePosition(position, path);
            switch (element)
            {
                case PositionedTextBox textBox:
                    ValidateBox(position, path);
                    ValidateRequiredText(textBox.Content, path);
                    ValidateShape(textBox.Fill, textBox.Stroke, textBox.CornerRadiusPt, path);
                    break;
                case PositionedImage image:
                    ValidateBox(position, path);
                    CheckRequiredString(image.Source, $"{path}.src", "image source");
                    CheckEnum(image.Fit, $"{path}.fit", "image fit mode");
                    ValidateCrop(image.Crop, $"{path}.crop");
                    break;
                case PositionedRectangle rectangle:
                    ValidateBox(position, path);
                    ValidateShape(rectangle.Fill, rectangle.Stroke, rectangle.CornerRadiusPt, path);
                    break;
                case PositionedLine line:
                    CheckOptionalRequiredPositive(position.WidthPt, $"{path}.width");
                    if (position.HeightPt is not null)
                    {
                        Error($"{path}.height", "a line has no height; use 'width' as its length and 'orientation' for its axis.");
                    }
                    CheckEnum(line.Orientation, $"{path}.orientation", "line orientation");
                    ValidateStroke(line.Stroke, $"{path}.stroke");
                    break;
                case PositionedCallout callout:
                    ValidateBox(position, path);
                    CheckEnum(callout.Tone, $"{path}.tone", "callout tone");
                    ValidateRequiredText(callout.Content, path);
                    ValidateShape(callout.Fill, callout.Stroke, callout.CornerRadiusPt, path);
                    break;
                default:
                    Error(path, $"unsupported positioned element type '{element.GetType().Name}'.");
                    break;
            }
        }

        private void ValidatePosition(PositionSpec position, string path)
        {
            CheckNonNegative(position.X, $"{path}.x");
            CheckNonNegative(position.Y, $"{path}.y");
            CheckFinite(position.Rotation, $"{path}.rotation");
            CheckEnum(position.Anchor, $"{path}.anchor", "anchor reference");
            CheckEnum(position.Wrap, $"{path}.wrap", "wrap mode");
            if (position.WrapDistances is { } distances)
            {
                CheckNonNegative(distances.TopPt, $"{path}.wrapDistances.top");
                CheckNonNegative(distances.LeftPt, $"{path}.wrapDistances.left");
                CheckNonNegative(distances.BottomPt, $"{path}.wrapDistances.bottom");
                CheckNonNegative(distances.RightPt, $"{path}.wrapDistances.right");
                if (position.Wrap is WrapMode.None or WrapMode.BehindText or WrapMode.InFrontOfText)
                {
                    Warn($"{path}.wrapDistances", "'wrapDistances' only applies to square/tight/through wraps (or topAndBottom).");
                }
            }
        }

        private void ValidateBox(PositionSpec position, string path)
        {
            CheckOptionalRequiredPositive(position.WidthPt, $"{path}.width");
            CheckOptionalRequiredPositive(position.HeightPt, $"{path}.height");
        }

        private void ValidateShape(string? fill, StrokeSpec? stroke, double? cornerRadius, string path)
        {
            CheckColor(fill, $"{path}.fill");
            ValidateStroke(stroke, $"{path}.stroke");
            CheckOptionalNonNegative(cornerRadius, $"{path}.cornerRadius");
        }

        private void ValidateStroke(StrokeSpec? stroke, string path)
        {
            if (stroke is null)
            {
                return;
            }
            if (stroke.Color is null)
            {
                Error($"{path}.color", "stroke color is required and must not be null.");
            }
            else if (stroke.Color.Length == 0)
            {
                Error($"{path}.color", "must not be empty.");
            }
            else
            {
                CheckColor(stroke.Color, $"{path}.color");
            }
            CheckPositive(stroke.WidthPt, $"{path}.width");
        }

        private void ValidateCrop(ImageCrop? crop, string path)
        {
            if (crop is null)
            {
                return;
            }
            bool leftValid = CheckFraction(crop.Left, $"{path}.left");
            bool topValid = CheckFraction(crop.Top, $"{path}.top");
            bool rightValid = CheckFraction(crop.Right, $"{path}.right");
            bool bottomValid = CheckFraction(crop.Bottom, $"{path}.bottom");
            if (leftValid && rightValid && crop.Left + crop.Right >= 1)
            {
                Error(path, "left + right crop ≥ 1 removes the whole image horizontally.");
            }
            if (topValid && bottomValid && crop.Top + crop.Bottom >= 1)
            {
                Error(path, "top + bottom crop ≥ 1 removes the whole image vertically.");
            }
        }

        private void ValidateSectionGeometry(Section section, string path)
        {
            if (!TryResolvePage(section.PageSetup, out double pageWidth, out double pageHeight, out Margins margins))
            {
                return;
            }
            string geometryPath = section.PageSetup is null ? path : $"{path}.pageSetup";
            if (margins.TopPt + margins.BottomPt >= pageHeight)
            {
                Error(geometryPath, $"margins exceed the page height: top+bottom = {FormatNumber(margins.TopPt + margins.BottomPt)}pt ≥ {FormatNumber(pageHeight)}pt.");
            }
            if (margins.LeftPt + margins.RightPt >= pageWidth)
            {
                Error(geometryPath, $"margins exceed the page width: left+right = {FormatNumber(margins.LeftPt + margins.RightPt)}pt ≥ {FormatNumber(pageWidth)}pt.");
            }

            if (section.PageSetup?.Columns is { Count: > 0 } columns &&
                IsFiniteNonNegative(columns.SpacingPt))
            {
                double textWidth = pageWidth - margins.LeftPt - margins.RightPt;
                if (textWidth > 0)
                {
                    double gutters = columns.SpacingPt * (columns.Count - 1);
                    if (double.IsFinite(gutters) && gutters >= textWidth)
                    {
                        Error(geometryPath, $"columns do not fit the text area: {columns.Count} columns × {FormatNumber(columns.SpacingPt)}pt gutters need {FormatNumber(gutters)}pt ≥ {FormatNumber(textWidth)}pt.");
                    }
                }
            }

            IReadOnlyList<PositionedElement>? positioned = section.Positioned;
            if (positioned is null)
            {
                return;
            }
            for (var i = 0; i < positioned.Count; i++)
            {
                PositionSpec? position = positioned[i]?.Position;
                if (position is null || !IsFiniteNonNegative(position.X) || !IsFiniteNonNegative(position.Y))
                {
                    continue;
                }
                string positionPath = $"{path}.positioned[{i}]";
                double rightEdge = position.X + (position.WidthPt ?? 0);
                if (double.IsFinite(rightEdge) && rightEdge > pageWidth)
                {
                    Warn(positionPath, $"extends past the right page edge ({FormatNumber(rightEdge)}pt > {FormatNumber(pageWidth)}pt).");
                }
                double bottomEdge = position.Y + (position.HeightPt ?? 0);
                if (double.IsFinite(bottomEdge) && bottomEdge > pageHeight)
                {
                    Warn(positionPath, $"extends past the bottom page edge ({FormatNumber(bottomEdge)}pt > {FormatNumber(pageHeight)}pt).");
                }
            }
        }

        private bool TryResolvePage(PageSetup? setup, out double width, out double height, out Margins margins)
        {
            width = 0;
            height = 0;
            margins = setup?.Margins ?? _pageDefaults.Margins ?? Margins.Defaults;
            if (!ValidMargins(margins))
            {
                return false;
            }

            PageOrientation orientation = setup?.Orientation ?? _pageDefaults.Orientation;
            if (!Enum.IsDefined(orientation))
            {
                return false;
            }
            PageSize? size = setup?.PageSize;
            if (size is null)
            {
                if (_pageDefaults.PageSize is { } defaultName)
                {
                    if (!Enum.IsDefined(defaultName) || !PageSize.Catalog.TryGetValue(defaultName, out size))
                    {
                        return false;
                    }
                }
                else
                {
                    size = PageSize.Default;
                }
            }
            if (!TryPageDimensions(size, out width, out height))
            {
                return false;
            }
            if (orientation == PageOrientation.Landscape)
            {
                (width, height) = (height, width);
            }
            return true;
        }

        private static bool TryPageDimensions(PageSize size, out double width, out double height)
        {
            width = 0;
            height = 0;
            if (size.Name is { } name)
            {
                if (size.WidthPt is not null || size.HeightPt is not null ||
                    !Enum.IsDefined(name) || !PageSize.Catalog.TryGetValue(name, out PageSize? catalogSize))
                {
                    return false;
                }
                width = catalogSize.WidthPt!.Value;
                height = catalogSize.HeightPt!.Value;
                return true;
            }
            if (size.WidthPt is not { } customWidth || size.HeightPt is not { } customHeight ||
                !IsFinitePositive(customWidth) || !IsFinitePositive(customHeight))
            {
                return false;
            }
            width = customWidth;
            height = customHeight;
            return true;
        }

        private void CheckColor(string? value, string path)
        {
            if (value is null)
            {
                return;
            }
            if (_palette.ContainsKey(value))
            {
                return;
            }
            if (HexColorPattern.IsMatch(value))
            {
                Warn(path, $"raw hex color '{value}' is off-palette: prefer a design palette token.");
                return;
            }
            if (value.StartsWith("#", StringComparison.Ordinal))
            {
                Error(path, $"'{value}' is not a valid hex color: expected #RRGGBB.");
                return;
            }
            Error(path, $"unknown color '{value}': not a palette token and not #RRGGBB hex.", Suggest(value, _palette.Keys));
        }

        private void CheckFontReference(string? value, string path)
        {
            var theme = DesignThemeCatalog.Resolve(_themeName);
            if (string.Equals(value, "display", StringComparison.OrdinalIgnoreCase) &&
                string.IsNullOrWhiteSpace(_fonts.Display) && string.IsNullOrWhiteSpace(theme.DisplayFontFamily))
            {
                Warn(path, "font slot 'display' is not defined in 'design.fonts' nor the active theme; emitters fall back to the document default.");
            }
            else if (string.Equals(value, "body", StringComparison.OrdinalIgnoreCase) &&
                     string.IsNullOrWhiteSpace(_fonts.Body) && string.IsNullOrWhiteSpace(theme.BodyFontFamily))
            {
                Warn(path, "font slot 'body' is not defined in 'design.fonts' nor the active theme; emitters fall back to the document default.");
            }
        }

        private void CheckRequiredString(string? value, string path, string subject)
        {
            if (value is null)
            {
                Error(path, $"{subject} is required and must not be null.");
            }
            else if (value.Length == 0)
            {
                Error(path, "must not be empty.");
            }
        }

        private void CheckOptionalRequiredPositive(double? value, string path)
        {
            if (value is null)
            {
                Error(path, "is required and must be > 0.");
            }
            else
            {
                CheckPositive(value.Value, path);
            }
        }

        private void CheckOptionalPositive(double? value, string path)
        {
            if (value is not null)
            {
                CheckPositive(value.Value, path);
            }
        }

        private void CheckOptionalNonNegative(double? value, string path)
        {
            if (value is not null)
            {
                CheckNonNegative(value.Value, path);
            }
        }

        private void CheckPositive(double value, string path)
        {
            if (!CheckFinite(value, path))
            {
                return;
            }
            if (value <= 0)
            {
                Error(path, $"must be > 0 (got {FormatNumber(value)}).");
            }
        }

        private void CheckNonNegative(double value, string path)
        {
            if (!CheckFinite(value, path))
            {
                return;
            }
            if (value < 0)
            {
                Error(path, $"must be ≥ 0 (got {FormatNumber(value)}).");
            }
        }

        private bool CheckFraction(double value, string path)
        {
            if (!CheckFinite(value, path))
            {
                return false;
            }
            if (value < 0)
            {
                Error(path, $"must be ≥ 0 (got {FormatNumber(value)}).");
                return false;
            }
            if (value > 1)
            {
                Error(path, $"must be ≤ 1 (got {FormatNumber(value)}).");
                return false;
            }
            return true;
        }

        private bool CheckFinite(double value, string path)
        {
            if (double.IsFinite(value))
            {
                return true;
            }
            Error(path, "must be a finite number.");
            return false;
        }

        private void CheckEnum<T>(T value, string path, string what) where T : struct, Enum
        {
            if (!Enum.IsDefined(value))
            {
                Error(path, $"'{value}' is not a valid {what}.");
            }
        }

        private void CheckNullableEnum<T>(T? value, string path, string what) where T : struct, Enum
        {
            if (value is not null)
            {
                CheckEnum(value.Value, path, what);
            }
        }

        private void Error(string path, string message, string? suggestion = null) =>
            _errors.Add(new DocxGenerationIssue(path, message, suggestion, DocxGenerationIssueSeverity.Error));

        private void Warn(string path, string message) =>
            _warnings.Add(new DocxGenerationIssue(path, message, null, DocxGenerationIssueSeverity.Warning));

        private static bool ValidMargins(Margins margins) =>
            IsFiniteNonNegative(margins.TopPt) &&
            IsFiniteNonNegative(margins.RightPt) &&
            IsFiniteNonNegative(margins.BottomPt) &&
            IsFiniteNonNegative(margins.LeftPt);

        private static bool IsFiniteNonNegative(double value) => double.IsFinite(value) && value >= 0;

        private static bool IsFinitePositive(double value) => double.IsFinite(value) && value > 0;

        private static string FormatNumber(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

        private static string? Suggest(string value, IEnumerable<string> candidates)
        {
            string? best = null;
            var bestDistance = int.MaxValue;
            foreach (string candidate in candidates)
            {
                int distance = Levenshtein(value, candidate);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    best = candidate;
                }
            }
            return bestDistance <= 2 ? $"Did you mean '{best}'?" : null;
        }

        private static int Levenshtein(string a, string b)
        {
            a = a.ToLowerInvariant();
            b = b.ToLowerInvariant();
            int[] previous = new int[b.Length + 1];
            int[] current = new int[b.Length + 1];
            for (var j = 0; j <= b.Length; j++)
            {
                previous[j] = j;
            }
            for (var i = 1; i <= a.Length; i++)
            {
                current[0] = i;
                for (var j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
                }
                (previous, current) = (current, previous);
            }
            return previous[b.Length];
        }
    }
}
