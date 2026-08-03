using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Instructions;

/// <summary>
/// Executes an <see cref="XlsxInstructionSet"/> (or its preflighted <see cref="XlsxPlan"/>)
/// against a <see cref="WorkbookBuilder"/>, producing the workbook.
///
/// The executor is plan-first: every instruction set is compiled through
/// <see cref="XlsxPlanner"/>, which validates the whole set and resolves every variable
/// (deterministically, with cycle detection), cell type, address, range and style
/// reference before anything is mutated. The executor then preflights the remaining
/// runtime conversions (typed values that must parse, style aspects the style builder can
/// represent, legacy numeric style ids) and only then applies the plan to the builder — so
/// a rejected set can never leave a partially-written workbook behind. Sets built
/// programmatically in code (bypassing the parser) are held to exactly the same rules as
/// parsed JSON.
/// </summary>
public static class XlsxInstructionExecutor
{
    private static readonly DateOnly ExcelMinDate = new(1900, 1, 1);
    private static readonly DateTime ExcelMinDateTime = new(1900, 1, 1);

    /// <summary>
    /// Compiles and applies all instructions from the set to the given workbook builder.
    /// Throws <see cref="XlsxException"/> with the first validation/planning error when the
    /// set cannot be executed; the builder is left untouched in that case.
    /// </summary>
    public static void Execute(XlsxInstructionSet instructions, IWorkbookBuilder builder)
    {
        ArgumentNullException.ThrowIfNull(instructions);
        ArgumentNullException.ThrowIfNull(builder);

        var plan = XlsxPlanner.Plan(instructions);
        if (!plan.IsValid)
        {
            var first = plan.Validation.Errors.First();
            throw new XlsxException(
                string.IsNullOrEmpty(first.Path) ? first.Message : $"{first.Message} (at {first.Path})");
        }

        Execute(plan.Plan!, builder);
    }

    /// <summary>
    /// Applies an already-preflighted plan. Internal entry point used by
    /// <see cref="XlsxGenerator"/> so a set is only planned once.
    ///
    /// The executor is strict about atomicity. Everything that can reject the set runs
    /// before the first mutation: style aspects (in <see cref="StyleContext.Create"/>),
    /// builder-dependent conflicts (worksheets/styles/tables already present), typed values
    /// that must parse, and legacy numeric style ids (range-checked after the stylesheet is
    /// populated, so the check is exact). Only stylesheet entries are produced by the
    /// prepare step — no metadata, sheet, or cell is written until every check has passed.
    /// </summary>
    internal static void Execute(XlsxPlan plan, IWorkbookBuilder builder)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(builder);

        var styles = StyleContext.Create(plan, builder);
        PreflightBuilderConflicts(plan, styles);
        PreflightCellConversions(plan);
        PrepareStyles(plan, builder, styles);
        PreflightRawStyleIds(plan, styles);

        ApplyPlan(plan, builder, styles);
    }

    // ─── Preflight (no cell/sheet/metadata mutation) ───────────────

    private static void PreflightBuilderConflicts(XlsxPlan plan, StyleContext styles)
    {
        var existingSheets = new HashSet<string>(
            styles.Builder.GetWorksheetNames(), StringComparer.OrdinalIgnoreCase);
        foreach (var ws in plan.Worksheets)
        {
            if (existingSheets.Contains(ws.Name))
            {
                throw new XlsxException(
                    $"Worksheet '{ws.Name}' already exists in the target workbook, so the " +
                    "instruction set cannot be applied without overwriting it. Remove or " +
                    "rename the existing worksheet, or apply the set to an empty workbook.");
            }
        }

        var existingStyles = new HashSet<string>(
            styles.Builder.GetDefinedStyleNames(), StringComparer.OrdinalIgnoreCase);
        foreach (var named in plan.Styles)
        {
            if (existingStyles.Contains(named.Name))
            {
                throw new XlsxException(
                    $"A named style '{named.Name}' is already defined in the target workbook, " +
                    "so the instruction set cannot redefine it. Use a different style name or " +
                    "apply the set to a workbook that does not define it.");
            }
        }

        if (styles.Workbook is { } workbook)
        {
            foreach (var ws in plan.Worksheets)
            {
                foreach (var table in ws.Tables)
                {
                    if (workbook.IsTableNameRegistered(table.Name))
                    {
                        throw new XlsxException(
                            $"A table named '{table.Name}' already exists in the target " +
                            "workbook. Table names must be unique workbook-wide; use a " +
                            "different table name or apply the set to an empty workbook.");
                    }
                }
            }
        }
    }

    private static void PreflightCellConversions(XlsxPlan plan)
    {
        foreach (var ws in plan.Worksheets)
        {
            foreach (var header in ws.Headers)
            {
                ValidateCellConversion(header);
            }

            foreach (var row in ws.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    ValidateCellConversion(cell);
                }
            }

            foreach (var cell in ws.Cells)
            {
                ValidateCellConversion(cell);
            }
        }
    }

    /// <summary>
    /// Populates the workbook stylesheet with everything the plan needs BEFORE any cell is
    /// written: the named style definitions, the cached legacy header-bold style (when a
    /// header row is written without an explicit style), and every number-format override
    /// resolution. The apply phase then reuses the cached indices, so a rejected set can
    /// never leave partial cell content behind — the only mutations made here are the
    /// stylesheet's own bookkeeping, and only after every read-only check has passed.
    /// </summary>
    private static void PrepareStyles(XlsxPlan plan, IWorkbookBuilder builder, StyleContext styles)
    {
        foreach (var spec in styles.Specs)
        {
            builder.DefineStyle(spec);
        }

        if (styles.Workbook is { } workbook)
        {
            if (plan.Worksheets.Any(ws => ws.Headers.Count > 0 && ws.HeaderStyle is null))
            {
                workbook.EnsureHeaderStyleIndex();
            }

            foreach (var ws in plan.Worksheets)
            {
                foreach (var cell in EnumerateCells(ws))
                {
                    if (cell.NumberFormat is null)
                    {
                        continue;
                    }

                    var styleName = cell.StyleRef is not null && !IsNumericId(cell.StyleRef)
                        ? cell.StyleRef
                        : null;
                    workbook.ResolveStyleIndex(styleName, cell.NumberFormat);
                }
            }
        }
    }

    /// <summary>
    /// Range-checks every legacy numeric style id against the stylesheet, which the prepare
    /// step has fully populated — so the check is exact and runs BEFORE any cell is
    /// written. An out-of-range id therefore fails atomically instead of after earlier
    /// cells have landed. A numeric id on a non-<see cref="WorkbookBuilder"/> is also
    /// rejected here, before any mutation.
    /// </summary>
    private static void PreflightRawStyleIds(XlsxPlan plan, StyleContext styles)
    {
        var workbook = styles.Workbook;
        if (workbook is null)
        {
            foreach (var ws in plan.Worksheets)
            {
                foreach (var cell in EnumerateCells(ws))
                {
                    if (IsNumericId(cell.StyleRef))
                    {
                        throw new XlsxException(
                            $"Cell {cell.Reference} references legacy numeric style id '{cell.StyleRef}', " +
                            "which requires a WorkbookBuilder; custom IWorkbookBuilder " +
                            "implementations cannot apply raw cellXf indices.");
                    }
                }
            }

            return;
        }

        foreach (var ws in plan.Worksheets)
        {
            foreach (var cell in EnumerateCells(ws))
            {
                // The number-format-override path resolves through the default cell format
                // and cannot fail; only the direct id path is range-checked.
                if (!IsNumericId(cell.StyleRef) || cell.NumberFormat is not null)
                {
                    continue;
                }

                var rawId = uint.Parse(cell.StyleRef!, NumberStyles.None, CultureInfo.InvariantCulture);
                workbook.EnsureStyleIndexExists(rawId, cell.Reference);
            }
        }
    }

    private static IEnumerable<XlsxPlanCell> EnumerateCells(XlsxPlanWorksheet ws)
    {
        foreach (var header in ws.Headers)
        {
            yield return header;
        }

        foreach (var row in ws.Rows)
        {
            foreach (var cell in row.Cells)
            {
                yield return cell;
            }
        }

        foreach (var cell in ws.Cells)
        {
            yield return cell;
        }
    }

    /// <summary>
    /// True when any cell in the plan carries a formula. Used to force full recalculation
    /// on load, because formula cells are written without cached values and must not show
    /// stale or empty results.
    /// </summary>
    private static bool HasAnyFormula(XlsxPlan plan)
    {
        foreach (var ws in plan.Worksheets)
        {
            foreach (var cell in EnumerateCells(ws))
            {
                if (cell.Formula is not null)
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Verifies every typed value converts cleanly and every formula is non-empty BEFORE
    /// anything is written, so a bad value mid-plan can never leave earlier cells behind.
    /// The builder methods re-validate the same inputs at write time; this mirrors them.
    /// </summary>
    private static void ValidateCellConversion(XlsxPlanCell cell)
    {
        if (cell.Formula is { Length: > 0 } formula)
        {
            var formulaText = formula.StartsWith('=') ? formula[1..] : formula;
            if (string.IsNullOrWhiteSpace(formulaText))
            {
                throw new XlsxException(
                    $"Cell {cell.Reference} has an empty formula. " +
                    "A formula must contain at least one token.");
            }

            return;
        }

        var value = cell.Value;
        switch (cell.Type)
        {
            case XlsxCellType.Number:
                if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                    || !double.IsFinite(number))
                {
                    throw new XlsxException(
                        $"Cell {cell.Reference} is typed as 'number' but its value '{value}' is not " +
                        "a finite number.");
                }

                break;

            case XlsxCellType.Boolean:
                if (!bool.TryParse(value, out _))
                {
                    throw new XlsxException(
                        $"Cell {cell.Reference} is typed as 'boolean' but its value '{value}' is not " +
                        "'true' or 'false'.");
                }

                break;

            case XlsxCellType.Date:
                if (!DateOnly.TryParseExact(
                        value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
                    || date < ExcelMinDate)
                {
                    throw new XlsxException(
                        $"Cell {cell.Reference} is typed as 'date' but its value '{value}' is not an " +
                        "ISO-8601 date (yyyy-MM-dd) on or after 1900-01-01.");
                }

                break;

            case XlsxCellType.DateTime:
                if (!DateTime.TryParseExact(
                        value,
                        new[] { "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.fff" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out var dateTime)
                    || dateTime < ExcelMinDateTime)
                {
                    throw new XlsxException(
                        $"Cell {cell.Reference} is typed as 'datetime' but its value '{value}' is not an " +
                        "ISO-8601 datetime (yyyy-MM-ddTHH:mm:ss) on or after 1900-01-01.");
                }

                break;
        }
    }

    // ─── Apply (mutates the builder) ───────────────────────────────

    private static void ApplyPlan(XlsxPlan plan, IWorkbookBuilder builder, StyleContext styles)
    {
        if (plan.Metadata is { } metadata)
        {
            builder.SetCoreProperties(new WorkbookCoreProperties
            {
                Title = metadata.Title,
                Subject = metadata.Subject,
                Creator = metadata.Author,
                Keywords = metadata.Keywords,
                Description = metadata.Comments,
                Category = metadata.Category
            });
        }

        if (HasAnyFormula(plan))
        {
            // Formula cells are written without cached values; full recalculation on load
            // is what makes them display real results instead of stale or empty ones.
            builder.SetCalculationProperties(new WorkbookCalculationProperties { FullCalcOnLoad = true });
        }

        foreach (var ws in plan.Worksheets)
        {
            ApplyWorksheet(ws, builder, styles);
        }
    }

    private static void ApplyWorksheet(XlsxPlanWorksheet ws, IWorkbookBuilder builder, StyleContext styles)
    {
        var sheet = builder.AddWorksheet(ws.Name);

        foreach (var column in ws.Columns)
        {
            if (column.Width is { } width)
            {
                sheet.SetColumnWidth(column.ColumnLetter, width);
            }
        }

        if (ws.Headers.Count > 0)
        {
            if (ws.HeaderStyle is null)
            {
                // No explicit header style: preserve the legacy bold header behaviour.
                sheet.AddHeaderRow(ws.Headers.Select(h => h.Value!).ToList(), ws.StartRow);
            }
            else
            {
                foreach (var header in ws.Headers)
                {
                    WriteCell(sheet, header, styles);
                }
            }
        }

        foreach (var row in ws.Rows)
        {
            foreach (var cell in row.Cells)
            {
                WriteCell(sheet, cell, styles);
            }
        }

        foreach (var cell in ws.Cells)
        {
            WriteCell(sheet, cell, styles);
        }

        foreach (var (rowIndex, height) in ws.RowHeights)
        {
            sheet.SetRowHeight(rowIndex, height);
        }

        foreach (var merge in ws.Merges)
        {
            sheet.MergeCells(merge);
        }

        if (ws.FreezePanes is { } freeze)
        {
            var frozenRows = freeze.Row - 1;
            var frozenColumns = freeze.Column - 1;
            if (frozenRows > 0 || frozenColumns > 0)
            {
                sheet.FreezePanes(frozenRows, frozenColumns);
            }
        }

        if (ws.AutoFilterRange is { } autoFilter)
        {
            sheet.SetAutoFilter(autoFilter);
        }

        foreach (var table in ws.Tables)
        {
            var bounds = table.Range.Split(':');
            sheet.AddTable(bounds[0], bounds[1], table.Name);
        }
    }

    private static void WriteCell(IWorksheetBuilder sheet, XlsxPlanCell cell, StyleContext styles)
    {
        if (cell.Formula is not null)
        {
            sheet.AddFormula(cell.Reference, cell.Formula, styles.StyleNameFor(cell), cell.NumberFormat);
        }
        else
        {
            WriteValueCell(sheet, cell, styles);
        }

        styles.ApplyLegacyStyle(sheet, cell);
    }

    private static void WriteValueCell(IWorksheetBuilder sheet, XlsxPlanCell cell, StyleContext styles)
    {
        var styleName = styles.StyleNameFor(cell);
        switch (cell.Type)
        {
            case XlsxCellType.String:
                sheet.AddCellString(cell.Reference, cell.Value ?? string.Empty, styleName);
                break;

            case XlsxCellType.Number:
                sheet.AddCellNumber(
                    cell.Reference,
                    double.Parse(cell.Value!, NumberStyles.Float, CultureInfo.InvariantCulture),
                    cell.NumberFormat,
                    styleName);
                break;

            case XlsxCellType.Boolean:
                sheet.AddCellBoolean(cell.Reference, bool.Parse(cell.Value!), styleName);
                break;

            case XlsxCellType.Date:
                sheet.AddCellDate(cell.Reference, cell.Value!, cell.NumberFormat, styleName);
                break;

            case XlsxCellType.DateTime:
                sheet.AddCellDateTime(cell.Reference, cell.Value!, cell.NumberFormat, styleName);
                break;

            default:
                throw new XlsxException(
                    $"Cell {cell.Reference} has no concrete type to write; expected a value or " +
                    "formula cell.");
        }
    }

    // ─── Styles ────────────────────────────────────────────────────

    private static bool IsNumericId(string? styleRef) =>
        styleRef is not null
        && uint.TryParse(styleRef, NumberStyles.None, CultureInfo.InvariantCulture, out _);

    /// <summary>
    /// Resolves named style references and number formats for the plan, and carries the
    /// legacy numeric-style-id path (which needs the concrete <see cref="WorkbookBuilder"/>
    /// and <see cref="WorksheetBuilder"/> to apply raw cellXf indices).
    /// </summary>
    private sealed class StyleContext
    {
        private readonly WorkbookBuilder? _workbook;

        private StyleContext(IReadOnlyList<CellStyleSpec> specs, IWorkbookBuilder builder, WorkbookBuilder? workbook)
        {
            Specs = specs;
            Builder = builder;
            _workbook = workbook;
        }

        public IReadOnlyList<CellStyleSpec> Specs { get; }

        /// <summary>The target builder (interface), used for read-only conflict preflights.</summary>
        public IWorkbookBuilder Builder { get; }

        /// <summary>The concrete workbook when the target is a <see cref="WorkbookBuilder"/>.</summary>
        public WorkbookBuilder? Workbook => _workbook;

        public static StyleContext Create(XlsxPlan plan, IWorkbookBuilder builder)
        {
            var specs = new List<CellStyleSpec>(plan.Styles.Count);
            foreach (var named in plan.Styles)
            {
                specs.Add(ConvertStyle(named));
            }

            return new StyleContext(specs, builder, builder as WorkbookBuilder);
        }

        /// <summary>The named style to pass to the typed builder APIs, or null for legacy numeric ids.</summary>
        public string? StyleNameFor(XlsxPlanCell cell) =>
            cell.StyleRef is not null && !IsNumericId(cell.StyleRef) ? cell.StyleRef : null;

        /// <summary>
        /// Applies a legacy numeric style id (e.g. "0") to an already-written cell. With no
        /// number-format override the id IS the cellXf index (v1 semantics, range-checked
        /// against the stylesheet before the cell is mutated); with one, the override
        /// resolves over the workbook's default cell format.
        /// </summary>
        public void ApplyLegacyStyle(IWorksheetBuilder sheet, XlsxPlanCell cell)
        {
            if (!IsNumericId(cell.StyleRef))
            {
                return;
            }

            var workbook = _workbook
                ?? throw new XlsxException(
                    $"Cell {cell.Reference} references legacy numeric style id '{cell.StyleRef}', which " +
                    "requires a WorkbookBuilder; custom IWorkbookBuilder implementations cannot apply " +
                    "raw cellXf indices.");
            var rawId = uint.Parse(cell.StyleRef!, NumberStyles.None, CultureInfo.InvariantCulture);

            uint index;
            if (cell.NumberFormat is not null)
            {
                index = workbook.ResolveStyleIndex(null, cell.NumberFormat);
            }
            else
            {
                workbook.EnsureStyleIndexExists(rawId, cell.Reference);
                index = rawId;
            }

            ApplyRawStyleIndex(sheet, cell.Reference, index);
        }
    }

    /// <summary>
    /// Applies a raw cellXf style index to a cell. Requires the concrete
    /// <see cref="WorksheetBuilder"/> (the only implementation of <see cref="IWorksheetBuilder"/>);
    /// custom implementations cannot represent raw indices.
    /// </summary>
    private static void ApplyRawStyleIndex(IWorksheetBuilder sheet, string reference, uint styleIndex)
    {
        if (sheet is WorksheetBuilder concrete)
        {
            concrete.ApplyStyleIndex(reference, styleIndex);
        }
        else
        {
            throw new XlsxException(
                $"Cell {reference} needs a raw cellXf style index ({styleIndex}), which custom " +
                "IWorksheetBuilder implementations cannot apply; use WorksheetBuilder.");
        }
    }

    private static CellStyleSpec ConvertStyle(NamedStyle named)
    {
        CellFontSpec? font = null;
        if (named.Font is { } f)
        {
            font = new CellFontSpec
            {
                Bold = f.Bold ?? false,
                Italic = f.Italic ?? false,
                ColorArgb = f.Color,
                Size = f.Size
            };
        }

        CellFillSpec? fill = null;
        if (named.Fill is { } fl)
        {
            var pattern = fl.Pattern?.Trim().ToLowerInvariant();
            if (pattern is not ("none" or null or ""))
            {
                fill = new CellFillSpec
                {
                    Pattern = MapFillPattern(pattern, named.Name),
                    SolidColorArgb = fl.Color
                };
            }
        }

        CellBorderSpec? border = null;
        if (named.Border is { } b)
        {
            var left = ConvertBorderEdge(b.Left);
            var right = ConvertBorderEdge(b.Right);
            var top = ConvertBorderEdge(b.Top);
            var bottom = ConvertBorderEdge(b.Bottom);
            if (left is not null || right is not null || top is not null || bottom is not null)
            {
                border = new CellBorderSpec { Left = left, Right = right, Top = top, Bottom = bottom };
            }
        }

        CellAlignmentSpec? alignment = null;
        if (named.Alignment is { } a)
        {
            var horizontal = ParseHorizontal(a.Horizontal);
            var vertical = ParseVertical(a.Vertical);
            if (horizontal is not null || vertical is not null || a.WrapText == true)
            {
                alignment = new CellAlignmentSpec
                {
                    Horizontal = horizontal,
                    Vertical = vertical,
                    WrapText = a.WrapText ?? false
                };
            }
        }

        var spec = new CellStyleSpec
        {
            Name = named.Name,
            Font = font,
            Fill = fill,
            Border = border,
            Alignment = alignment,
            NumberFormat = named.NumberFormat
        };
        spec.Validate();
        return spec;
    }

    /// <summary>
    /// Maps a vocabulary fill pattern to its SpreadsheetML pattern type. Only real Excel
    /// pattern types are representable; anything else (including made-up "diagonal" names)
    /// throws here, and the semantic validator rejects the same names with a path-qualified
    /// diagnostic — validation and execution can never disagree.
    /// </summary>
    private static PatternValues MapFillPattern(string pattern, string styleName)
    {
        switch (pattern)
        {
            case "solid":
                return PatternValues.Solid;
            case "mediumgray":
                return PatternValues.MediumGray;
            case "darkgray":
                return PatternValues.DarkGray;
            case "lightgray":
                return PatternValues.LightGray;
            case "gray125":
                return PatternValues.Gray125;
            case "gray0625":
                return PatternValues.Gray0625;
            case "darkhorizontal":
                return PatternValues.DarkHorizontal;
            case "darkvertical":
                return PatternValues.DarkVertical;
            case "darkdown":
                return PatternValues.DarkDown;
            case "darkup":
                return PatternValues.DarkUp;
            case "darkgrid":
                return PatternValues.DarkGrid;
            case "darktrellis":
                return PatternValues.DarkTrellis;
            case "lighthorizontal":
                return PatternValues.LightHorizontal;
            case "lightvertical":
                return PatternValues.LightVertical;
            case "lightdown":
                return PatternValues.LightDown;
            case "lightup":
                return PatternValues.LightUp;
            case "lightgrid":
                return PatternValues.LightGrid;
            case "lighttrellis":
                return PatternValues.LightTrellis;
            default:
                throw new XlsxException(
                    $"Style '{styleName}' uses fill pattern '{pattern}', which is not a valid " +
                    "Excel pattern type. Use one of: none, solid, gray125, gray0625, darkgray, " +
                    "mediumgray, lightgray, darkHorizontal, darkVertical, darkDown, darkUp, " +
                    "darkGrid, darkTrellis, lightHorizontal, lightVertical, lightDown, lightUp, " +
                    "lightGrid, lightTrellis.");
        }
    }

    private static CellEdgeSpec? ConvertBorderEdge(BorderEdgeInstruction? edge)
    {
        if (edge?.Style is not { } rawStyle)
        {
            return null;
        }

        var style = rawStyle.Trim().ToLowerInvariant();
        if (style is "" or "none")
        {
            return null;
        }

        return new CellEdgeSpec
        {
            Style = MapBorderStyle(style),
            ColorArgb = edge.Color
        };
    }

    /// <summary>
    /// Maps a vocabulary border style to its SpreadsheetML border style. Every real Excel
    /// border style is representable; unknown names throw here and are rejected by the
    /// semantic validator too, so validation and execution agree.
    /// </summary>
    private static BorderStyleValues MapBorderStyle(string style)
    {
        switch (style)
        {
            case "thin":
                return BorderStyleValues.Thin;
            case "medium":
                return BorderStyleValues.Medium;
            case "thick":
                return BorderStyleValues.Thick;
            case "double":
                return BorderStyleValues.Double;
            case "dashed":
                return BorderStyleValues.Dashed;
            case "dotted":
                return BorderStyleValues.Dotted;
            case "dashdot":
                return BorderStyleValues.DashDot;
            case "dashdotdot":
                return BorderStyleValues.DashDotDot;
            case "hair":
                return BorderStyleValues.Hair;
            case "mediumdashed":
                return BorderStyleValues.MediumDashed;
            case "mediumdashdot":
                return BorderStyleValues.MediumDashDot;
            case "mediumdashdotdot":
                return BorderStyleValues.MediumDashDotDot;
            case "slantdashdot":
                return BorderStyleValues.SlantDashDot;
            default:
                throw new XlsxException(
                    $"Unknown border style '{style}'. Known styles: none, thin, medium, thick, " +
                    "double, dashed, dotted, dashDot, dashDotDot, hair, mediumDashed, " +
                    "mediumDashDot, mediumDashDotDot, slantDashDot.");
        }
    }

    private static HorizontalAlignmentValues? ParseHorizontal(string? value)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case null:
                return null;
            case "general":
                return HorizontalAlignmentValues.General;
            case "left":
                return HorizontalAlignmentValues.Left;
            case "center":
                return HorizontalAlignmentValues.Center;
            case "right":
                return HorizontalAlignmentValues.Right;
            case "fill":
                return HorizontalAlignmentValues.Fill;
            case "justify":
                return HorizontalAlignmentValues.Justify;
            case "centercontinuous":
                return HorizontalAlignmentValues.CenterContinuous;
            default:
                throw new XlsxException(
                    $"Unknown horizontal alignment '{value}'. Known values: general, left, center, " +
                    "right, fill, justify, centerContinuous.");
        }
    }

    private static VerticalAlignmentValues? ParseVertical(string? value)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case null:
                return null;
            case "top":
                return VerticalAlignmentValues.Top;
            case "center":
                return VerticalAlignmentValues.Center;
            case "bottom":
                return VerticalAlignmentValues.Bottom;
            case "justify":
                return VerticalAlignmentValues.Justify;
            case "distributed":
                return VerticalAlignmentValues.Distributed;
            default:
                throw new XlsxException(
                    $"Unknown vertical alignment '{value}'. Known values: top, center, bottom, " +
                    "justify, distributed.");
        }
    }
}
