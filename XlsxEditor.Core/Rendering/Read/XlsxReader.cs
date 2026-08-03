using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Rendering.Models;

namespace XlsxEditor.Core.Rendering.Read;

public sealed class XlsxReader
{
    private static readonly HashSet<uint> DateFormatIds = new()
    {
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47,
        50, 51, 52, 53, 54, 55, 56, 57, 58
    };

    private static readonly DateTime Epoch1900 = new(1899, 12, 30);
    private static readonly DateTime Epoch1904 = new(1904, 1, 1);

    public XlsxReadResult Read(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        using var doc = SpreadsheetDocument.Open(path, false);
        return ReadCore(doc);
    }

    public XlsxReadResult Read(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms, false);
        return ReadCore(doc);
    }

    public XlsxReadResult Read(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);
        return ReadCore(doc);
    }

    private static XlsxReadResult ReadCore(SpreadsheetDocument doc)
    {
        var issues = new List<XlsxReadIssue>();

        if (doc.WorkbookPart is null)
        {
            issues.Add(new XlsxReadIssue
            {
                Severity = XlsxReadIssueSeverity.Warning,
                Message = "Workbook has no workbook part."
            });
            return new XlsxReadResult { Issues = issues };
        }

        var workbookPart = doc.WorkbookPart;
        var workbook = workbookPart.Workbook;

        string? title = doc.PackageProperties.Title;
        bool date1904 = false;
        if (workbook?.WorkbookProperties is { } props)
        {
            date1904 = props.Date1904?.Value ?? false;
        }

        var sharedStrings = ReadSharedStrings(workbookPart, issues);
        var styleCache = ReadStyles(workbookPart, issues);

        var definedNames = ReadDefinedNames(workbook, issues);

        var sheets = workbook?.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
        var renderSheets = new List<XlsxRenderSheet>();

        foreach (var sheet in sheets)
        {
            if (sheet.Name?.Value is not { } sheetName)
            {
                issues.Add(new XlsxReadIssue
                {
                    Severity = XlsxReadIssueSeverity.Warning,
                    Message = "Sheet has no name, skipping."
                });
                continue;
            }

            if (sheet.Id?.Value is not { } relationshipId)
            {
                issues.Add(new XlsxReadIssue
                {
                    Severity = XlsxReadIssueSeverity.Warning,
                    Message = $"Sheet '{sheetName}' has no relationship id, skipping.",
                    Location = sheetName
                });
                continue;
            }

            var sheetIssues = new List<XlsxReadIssue>();
            XlsxRenderSheet renderSheet;

            try
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                renderSheet = ReadSheet(
                    worksheetPart,
                    sheetName,
                    sharedStrings,
                    styleCache,
                    definedNames,
                    date1904,
                    sheetIssues);
            }
            catch (Exception ex)
            {
                sheetIssues.Add(new XlsxReadIssue
                {
                    Severity = XlsxReadIssueSeverity.Warning,
                    Message = $"Failed to read sheet '{sheetName}': {ex.Message}",
                    Location = sheetName
                });
                continue;
            }

            issues.AddRange(sheetIssues);
            renderSheets.Add(renderSheet);
        }

        return new XlsxReadResult
        {
            Workbook = new XlsxRenderWorkbook
            {
                Title = title,
                Sheets = renderSheets
            },
            Issues = issues
        };
    }

    private static Dictionary<int, string> ReadSharedStrings(
        WorkbookPart workbookPart,
        List<XlsxReadIssue> issues)
    {
        var result = new Dictionary<int, string>();

        if (workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() is not { } sstPart)
        {
            return result;
        }

        try
        {
            var sst = sstPart.SharedStringTable;

            // Accessing SharedStringTable loads the XML; if the file is corrupt, it may throw.
            // We wrap in a defensive block below.
            if (sst is null) return result;

            int index = 0;
            foreach (var si in sst.Elements<SharedStringItem>())
            {
                var text = GetSharedStringText(si);
                result[index++] = text;
            }
        }
        catch (Exception ex)
        {
            issues.Add(new XlsxReadIssue
            {
                Severity = XlsxReadIssueSeverity.Warning,
                Message = $"Failed to read shared strings table: {ex.Message}"
            });
        }

        return result;
    }

    private static string GetSharedStringText(SharedStringItem si)
    {
        var t = si.Elements<Text>().FirstOrDefault();
        if (t is not null)
        {
            return t.InnerText;
        }

        var runs = si.Elements<Run>();
        var parts = new List<string>();
        foreach (var run in runs)
        {
            var runText = run.Elements<Text>().FirstOrDefault();
            if (runText is not null)
            {
                parts.Add(runText.InnerText);
            }
        }

        return string.Concat(parts);
    }

    private static XlsxStyleCache ReadStyles(
        WorkbookPart workbookPart,
        List<XlsxReadIssue> issues)
    {
        var cache = new XlsxStyleCache();

        if (workbookPart.WorkbookStylesPart?.Stylesheet is not { } stylesheet)
        {
            return cache;
        }

        try
        {
            ReadNumberFormats(stylesheet, cache, issues);
            ReadFonts(stylesheet, cache);
            ReadFills(stylesheet, cache);
            ReadBorders(stylesheet, cache);
            ReadCellFormats(stylesheet, cache);
        }
        catch (Exception ex)
        {
            issues.Add(new XlsxReadIssue
            {
                Severity = XlsxReadIssueSeverity.Warning,
                Message = $"Failed to read styles: {ex.Message}"
            });
        }

        return cache;
    }

    private static void ReadNumberFormats(
        Stylesheet stylesheet,
        XlsxStyleCache cache,
        List<XlsxReadIssue> issues)
    {
        if (stylesheet.NumberingFormats is not { } numFmts) return;

        foreach (var nf in numFmts.Elements<NumberingFormat>())
        {
            if (nf.NumberFormatId?.Value is { } id && nf.FormatCode?.Value is { } code)
            {
                cache.SetCustomNumberFormat(id, code);
            }
        }
    }

    private static void ReadFonts(Stylesheet stylesheet, XlsxStyleCache cache)
    {
        if (stylesheet.Fonts is not { } fonts) return;

        foreach (var font in fonts.Elements<Font>())
        {
            var name = font.FontName?.Val?.Value;
            var size = font.FontSize?.Val?.Value;
            var bold = font.Bold != null;
            var italic = font.Italic != null;
            string? colorArgb = null;

            var color = font.Color;
            if (color is not null)
            {
                if (color.Rgb?.Value is { } rgb)
                {
                    colorArgb = rgb;
                }
                else if (color.Theme is not null)
                {
                    colorArgb = null;
                }
            }

            cache.AddFont(new CachedFont(name, size, bold, italic, colorArgb));
        }
    }

    private static void ReadFills(Stylesheet stylesheet, XlsxStyleCache cache)
    {
        if (stylesheet.Fills is not { } fills) return;

        foreach (var fill in fills.Elements<Fill>())
        {
            string? fillArgb = null;

            var patternFill = fill.PatternFill;
            if (patternFill?.ForegroundColor is { } fg)
            {
                if (fg.Rgb?.Value is { } rgb)
                {
                    fillArgb = rgb;
                }
            }

            cache.AddFill(fillArgb);
        }
    }

    private static void ReadBorders(Stylesheet stylesheet, XlsxStyleCache cache)
    {
        if (stylesheet.Borders is not { } borders) return;

        foreach (var border in borders.Elements<Border>())
        {
            var cached = new CachedBorder(
                ReadEdge(border.LeftBorder),
                ReadEdge(border.RightBorder),
                ReadEdge(border.TopBorder),
                ReadEdge(border.BottomBorder));
            cache.AddBorder(cached);
        }
    }

    private static CachedBorderEdge ReadEdge(BorderPropertiesType? edge)
    {
        if (edge is null) return new CachedBorderEdge(null, null);

        var style = edge.Style?.ToString();
        string? colorArgb = null;

        if (edge.Color is { } c)
        {
            if (c.Rgb?.Value is { } rgb)
            {
                colorArgb = rgb;
            }
        }

        return new CachedBorderEdge(style, colorArgb);
    }

    private static void ReadCellFormats(Stylesheet stylesheet, XlsxStyleCache cache)
    {
        if (stylesheet.CellStyleFormats is { } styleFormats)
        {
            foreach (var xf in styleFormats.Elements<CellFormat>())
            {
                cache.AddCellStyleXf(new CachedCellFormat(
                    xf.NumberFormatId?.Value ?? 0U,
                    xf.FontId?.Value ?? 0U,
                    xf.FillId?.Value ?? 0U,
                    xf.BorderId?.Value ?? 0U,
                    xf.ApplyNumberFormat?.Value ?? true,
                    xf.ApplyFont?.Value ?? true,
                    xf.ApplyFill?.Value ?? true,
                    xf.ApplyBorder?.Value ?? true,
                    xf.ApplyAlignment?.Value ?? true,
                    ReadAlignment(xf),
                    xf.FormatId?.Value));
            }
        }

        if (stylesheet.CellFormats is { } cellFormats)
        {
            foreach (var xf in cellFormats.Elements<CellFormat>())
            {
                cache.AddCellXf(new CachedCellFormat(
                    xf.NumberFormatId?.Value ?? 0U,
                    xf.FontId?.Value ?? 0U,
                    xf.FillId?.Value ?? 0U,
                    xf.BorderId?.Value ?? 0U,
                    xf.ApplyNumberFormat?.Value ?? true,
                    xf.ApplyFont?.Value ?? true,
                    xf.ApplyFill?.Value ?? true,
                    xf.ApplyBorder?.Value ?? true,
                    xf.ApplyAlignment?.Value ?? true,
                    ReadAlignment(xf),
                    xf.FormatId?.Value));
            }
        }
    }

    private static CachedAlignment ReadAlignment(CellFormat xf)
    {
        var al = xf.Alignment;
        if (al is null) return new CachedAlignment(null, null, false);

        return new CachedAlignment(
            al.Horizontal?.ToString(),
            al.Vertical?.ToString(),
            al.WrapText?.Value ?? false);
    }

    private static Dictionary<string, string> ReadDefinedNames(
        Workbook? workbook,
        List<XlsxReadIssue> issues)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (workbook?.DefinedNames is not { } definedNames) return result;

        foreach (var dn in definedNames.Elements<DefinedName>())
        {
            if (dn.Name?.Value is { } name && dn.InnerText is { } value)
            {
                result[name] = value;
            }
        }

        return result;
    }

    private static XlsxRenderSheet ReadSheet(
        WorksheetPart worksheetPart,
        string sheetName,
        Dictionary<int, string> sharedStrings,
        XlsxStyleCache styleCache,
        Dictionary<string, string> definedNames,
        bool date1904,
        List<XlsxReadIssue> issues)
    {
        var worksheet = worksheetPart.Worksheet;
        if (worksheet is null)
        {
            issues.Add(new XlsxReadIssue
            {
                Severity = XlsxReadIssueSeverity.Warning,
                Message = $"Worksheet part for '{sheetName}' has no worksheet element.",
                Location = sheetName
            });
            return new XlsxRenderSheet { Name = sheetName };
        }

        var (pageWidthPt, pageHeightPt, landscape) = ReadPageSetup(worksheet, issues, sheetName);

        var columns = ReadColumns(worksheet, issues, sheetName);

        var (rows, allCells) = ReadRows(worksheet, sharedStrings, styleCache, date1904, issues, sheetName);

        var merges = ReadMerges(worksheet, issues, sheetName);

        var freezePanes = ReadFreezePanes(worksheet, issues, sheetName);

        var autoFilterRange = ReadAutoFilter(worksheet, issues, sheetName);

        var tables = ReadTables(worksheetPart, issues, sheetName);

        var printArea = ReadPrintArea(definedNames, sheetName, issues);

        return new XlsxRenderSheet
        {
            Name = sheetName,
            PageWidthPt = pageWidthPt,
            PageHeightPt = pageHeightPt,
            Landscape = landscape,
            Columns = columns,
            Rows = rows,
            Merges = merges,
            FreezePanes = freezePanes,
            Tables = tables,
            PrintArea = printArea,
            AutoFilterRange = autoFilterRange
        };
    }

    private static (double widthPt, double heightPt, bool landscape) ReadPageSetup(
        Worksheet worksheet,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var pageSetup = worksheet.GetFirstChild<PageSetup>();
        if (pageSetup is null) return (0, 0, false);

        bool landscape = pageSetup.Orientation?.Value == OrientationValues.Landscape;

        double paperSize = pageSetup.PaperSize?.Value ?? 0U;
        var (paperW, paperH) = PaperSizeToPt(paperSize);
        double widthPt = paperW;
        double heightPt = paperH;

        if (pageSetup.FitToWidth?.Value is > 0 || pageSetup.FitToHeight?.Value is > 0)
        {
            // fitToPage is metadata for the emitter; page dimensions still apply
        }

        return (widthPt, heightPt, landscape);
    }

    private static (double w, double h) PaperSizeToPt(double paperSize)
    {
        return paperSize switch
        {
            1 => (612, 792),       // Letter
            5 => (612, 1008),      // Legal
            9 => (595, 842),       // A4
            13 => (612, 1224),     // Ledger
            _ => (0, 0)
        };
    }

    private static double InchesToPoints(double inches) => inches * 72.0;

    private static IReadOnlyList<XlsxRenderColumn> ReadColumns(
        Worksheet worksheet,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var result = new List<XlsxRenderColumn>();
        var cols = worksheet.GetFirstChild<Columns>();
        if (cols is null) return result;

        foreach (var col in cols.Elements<Column>())
        {
            var startIdx = col.Min?.Value is { } minVal ? (int)minVal : 1;
            var endIdx = col.Max?.Value is { } maxVal ? (int)maxVal : startIdx;
            var widthChars = col.Width?.Value ?? 0.0;
            var widthPt = widthChars * 7.0;
            var hidden = col.Hidden?.Value ?? false;

            for (int i = startIdx; i <= endIdx; i++)
            {
                result.Add(new XlsxRenderColumn
                {
                    Index = i,
                    WidthPt = widthPt,
                    Hidden = hidden
                });
            }
        }

        return result;
    }

    private static (IReadOnlyList<XlsxRenderRow> rows, List<XlsxRenderCell> allCells) ReadRows(
        Worksheet worksheet,
        Dictionary<int, string> sharedStrings,
        XlsxStyleCache styleCache,
        bool date1904,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            return (Array.Empty<XlsxRenderRow>(), new List<XlsxRenderCell>());
        }

        var renderRows = new List<XlsxRenderRow>();
        var allCells = new List<XlsxRenderCell>();

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx == 0) continue;

            var heightPt = row.Height?.Value ?? 0;
            var hidden = row.Hidden?.Value ?? false;

            var cells = new List<XlsxRenderCell>();

            foreach (var cell in row.Elements<Cell>())
            {
                var renderCell = ReadCell(
                    cell,
                    sharedStrings,
                    styleCache,
                    date1904,
                    issues,
                    sheetName);
                cells.Add(renderCell);
                allCells.Add(renderCell);
            }

            renderRows.Add(new XlsxRenderRow
            {
                Index = rowIdx,
                HeightPt = heightPt,
                Hidden = hidden,
                Cells = cells
            });
        }

        return (renderRows, allCells);
    }

    private static XlsxRenderCell ReadCell(
        Cell cell,
        Dictionary<int, string> sharedStrings,
        XlsxStyleCache styleCache,
        bool date1904,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var (row, col) = ParseCellReference(cell.CellReference?.Value);

        var styleIndex = (int)(cell.StyleIndex?.Value ?? 0);
        var resolvedStyle = styleCache.ResolveEffectiveStyle(styleIndex);
        var numberFormatCode = resolvedStyle.NumberFormatCode;

        var formula = cell.CellFormula?.Text;
        if (formula is not null && formula.Length > 0)
        {
            formula = "=" + formula;
        }

        var value = ReadCellValue(cell, sharedStrings, date1904, numberFormatCode);

        return new XlsxRenderCell
        {
            Row = row,
            Column = col,
            Value = value,
            Formula = formula,
            Style = resolvedStyle.Style,
            NumberFormatCode = numberFormatCode
        };
    }

    private static XlsxRenderValue ReadCellValue(
        Cell cell,
        Dictionary<int, string> sharedStrings,
        bool date1904,
        string? numberFormatCode)
    {
        var dataType = cell.DataType?.Value;
        var rawValue = cell.CellValue?.Text;

        if (dataType is null && rawValue is null)
        {
            var inlineStr = cell.InlineString;
            if (inlineStr is not null)
            {
                var text = inlineStr.InnerText;
                return string.IsNullOrEmpty(text)
                    ? XlsxRenderValue.Empty()
                    : XlsxRenderValue.Text(text);
            }
            return XlsxRenderValue.Empty();
        }

        if (dataType == CellValues.SharedString)
            return ReadSharedStringValue(rawValue, sharedStrings);

        if (dataType == CellValues.Boolean)
            return rawValue == "1"
                ? XlsxRenderValue.Boolean(true)
                : XlsxRenderValue.Boolean(false);

        if (dataType == CellValues.Error)
            return XlsxRenderValue.Error(rawValue ?? "#UNKNOWN");

        if (dataType == CellValues.InlineString)
            return ReadInlineStringValue(cell);

        if (dataType == CellValues.String)
            return !string.IsNullOrEmpty(rawValue)
                ? XlsxRenderValue.Text(rawValue)
                : XlsxRenderValue.Empty();

        return ReadNumericValue(rawValue, date1904, numberFormatCode);
    }

    private static XlsxRenderValue ReadSharedStringValue(
        string? rawValue,
        Dictionary<int, string> sharedStrings)
    {
        if (rawValue is null) return XlsxRenderValue.Empty();

        if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
            && sharedStrings.TryGetValue(idx, out var text))
        {
            return string.IsNullOrEmpty(text) ? XlsxRenderValue.Empty() : XlsxRenderValue.Text(text);
        }

        return XlsxRenderValue.Empty();
    }

    private static XlsxRenderValue ReadInlineStringValue(Cell cell)
    {
        var inlineStr = cell.InlineString;
        if (inlineStr is null) return XlsxRenderValue.Empty();

        var text = inlineStr.InnerText;
        return string.IsNullOrEmpty(text) ? XlsxRenderValue.Empty() : XlsxRenderValue.Text(text);
    }

    private static XlsxRenderValue ReadNumericValue(
        string? rawValue,
        bool date1904,
        string? numberFormatCode)
    {
        if (rawValue is null) return XlsxRenderValue.Empty();

        if (!double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return XlsxRenderValue.Empty();
        }

        if (IsDateFormatCode(numberFormatCode))
        {
            var dateTime = SerialToDateTime(number, date1904);
            return XlsxRenderValue.DateTime(dateTime, number);
        }

        return XlsxRenderValue.Number(number);
    }

    internal static bool IsDateFormatCode(string? formatCode)
    {
        if (formatCode is null) return false;

        if (uint.TryParse(formatCode, NumberStyles.Integer, CultureInfo.InvariantCulture, out var builtInId))
        {
            return DateFormatIds.Contains(builtInId);
        }

        var upper = formatCode.ToUpperInvariant();
        return upper.Contains('Y') || upper.Contains('D') || upper.Contains('H') || upper.Contains('S');
    }

    internal static DateTime SerialToDateTime(double serial, bool use1904)
    {
        if (use1904)
        {
            var days = (int)Math.Floor(serial);
            var fraction = serial - days;
            var datePart = Epoch1904.AddDays(days);
            var ticks1904 = (long)Math.Round(fraction * TimeSpan.TicksPerDay, MidpointRounding.ToEven);
            return datePart.AddTicks(ticks1904);
        }

        var d = (int)Math.Floor(serial);
        var f = serial - d;
        var dp = Epoch1900.AddDays(d);
        var tks = (long)Math.Round(f * TimeSpan.TicksPerDay, MidpointRounding.ToEven);
        return dp.AddTicks(tks);
    }

    private static IReadOnlyList<XlsxRenderMerge> ReadMerges(
        Worksheet worksheet,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var merges = new List<XlsxRenderMerge>();
        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells is null) return merges;

        foreach (var mc in mergeCells.Elements<MergeCell>())
        {
            if (mc.Reference?.Value is not { } refStr) continue;

            try
            {
                var parts = refStr.Split(':');
                if (parts.Length != 2) continue;

                var (firstRow, firstCol) = ParseCellReference(parts[0]);
                var (lastRow, lastCol) = ParseCellReference(parts[1]);

                merges.Add(new XlsxRenderMerge
                {
                    FirstRow = firstRow,
                    LastRow = lastRow,
                    FirstCol = firstCol,
                    LastCol = lastCol
                });
            }
            catch
            {
                issues.Add(new XlsxReadIssue
                {
                    Severity = XlsxReadIssueSeverity.Warning,
                    Message = $"Unparseable merge range '{refStr}' in sheet '{sheetName}'.",
                    Location = sheetName
                });
            }
        }

        return merges;
    }

    private static XlsxRenderFreezePanes? ReadFreezePanes(
        Worksheet worksheet,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var sheetViews = worksheet.GetFirstChild<SheetViews>();
        if (sheetViews is null) return null;

        var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
        if (sheetView?.Pane is not { } pane) return null;

        if (pane.State?.Value != PaneStateValues.Frozen) return null;

        return new XlsxRenderFreezePanes
        {
            FrozenRows = (int)(pane.VerticalSplit?.Value ?? 0),
            FrozenCols = (int)(pane.HorizontalSplit?.Value ?? 0)
        };
    }

    private static string? ReadAutoFilter(
        Worksheet worksheet,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var autoFilter = worksheet.GetFirstChild<AutoFilter>();
        return autoFilter?.Reference?.Value;
    }

    private static IReadOnlyList<XlsxRenderTable> ReadTables(
        WorksheetPart worksheetPart,
        List<XlsxReadIssue> issues,
        string sheetName)
    {
        var tables = new List<XlsxRenderTable>();

        var tableParts = worksheetPart.TableDefinitionParts;
        if (tableParts is null) return tables;

        foreach (var tablePart in tableParts)
        {
            try
            {
                var table = tablePart.Table;
                if (table is null) continue;

                var name = table.Name?.Value
                    ?? table.DisplayName?.Value
                    ?? "Table";

                if (table.Reference?.Value is not { } refStr) continue;

                var parts = refStr.Split(':');
                if (parts.Length != 2) continue;

                var (firstRow, firstCol) = ParseCellReference(parts[0]);
                var (lastRow, lastCol) = ParseCellReference(parts[1]);

                tables.Add(new XlsxRenderTable
                {
                    Name = name,
                    FirstRow = firstRow,
                    LastRow = lastRow,
                    FirstCol = firstCol,
                    LastCol = lastCol
                });
            }
            catch (Exception ex)
            {
                issues.Add(new XlsxReadIssue
                {
                    Severity = XlsxReadIssueSeverity.Warning,
                    Message = $"Failed to read table in sheet '{sheetName}': {ex.Message}",
                    Location = sheetName
                });
            }
        }

        return tables;
    }

    private static string? ReadPrintArea(
        Dictionary<string, string> definedNames,
        string sheetName,
        List<XlsxReadIssue> issues)
    {
        var candidates = new[]
        {
            $"'{sheetName}'!Print_Area",
            $"{sheetName}!Print_Area",
            "Print_Area"
        };

        foreach (var candidate in candidates)
        {
            if (definedNames.TryGetValue(candidate, out var value))
            {
                return ExtractPrintAreaRange(value, sheetName);
            }
        }

        return null;
    }

    private static string? ExtractPrintAreaRange(string definedNameValue, string sheetName)
    {
        var value = definedNameValue.Trim();

        var bangIdx = value.LastIndexOf('!');
        if (bangIdx >= 0)
        {
            value = value[(bangIdx + 1)..];
        }

        value = value.Replace("$", "");

        if (string.IsNullOrWhiteSpace(value)) return null;

        return value;
    }

    internal static (int row, int col) ParseCellReference(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) return (0, 0);

        int i = 0;
        while (i < reference.Length && char.IsLetter(reference[i])) i++;

        if (i == 0 || i == reference.Length) return (0, 0);

        var colStr = reference[..i];
        var rowStr = reference[i..];

        if (!int.TryParse(rowStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row)) return (0, 0);

        int col = 0;
        foreach (var ch in colStr)
        {
            col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }

        return (row, col);
    }
}

internal sealed class XlsxStyleCache
{
    private readonly List<CachedFont> _fonts = new();
    private readonly List<CachedFill> _fills = new();
    private readonly List<CachedBorder> _borders = new();
    private readonly List<CachedCellFormat> _cellStyleXfs = new();
    private readonly List<CachedCellFormat> _cellXfs = new();
    private readonly Dictionary<uint, string> _customNumberFormats = new();

    private static readonly Dictionary<uint, string> BuiltInNumberFormats = new()
    {
        [0] = "General",
        [1] = "0",
        [2] = "0.00",
        [3] = "#,##0",
        [4] = "#,##0.00",
        [9] = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "m/d/yyyy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yyyy h:mm",
        [37] = "#,##0 ;(#,##0)",
        [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)",
        [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mm:ss.0",
        [48] = "##0.0E+0",
        [49] = "@"
    };

    public void SetCustomNumberFormat(uint id, string code)
    {
        _customNumberFormats[id] = code;
    }

    public string GetNumberFormatCode(uint id)
    {
        if (_customNumberFormats.TryGetValue(id, out var code)) return code;
        if (BuiltInNumberFormats.TryGetValue(id, out var builtIn)) return builtIn;
        return string.Empty;
    }

    public void AddFont(CachedFont font) => _fonts.Add(font);
    public void AddFill(string? fillArgb) => _fills.Add(new CachedFill(fillArgb));
    public void AddBorder(CachedBorder border) => _borders.Add(border);
    public void AddCellStyleXf(CachedCellFormat xf) => _cellStyleXfs.Add(xf);
    public void AddCellXf(CachedCellFormat xf) => _cellXfs.Add(xf);

    public ResolvedStyle ResolveEffectiveStyle(int cellXfIndex)
    {
        if (cellXfIndex < 0 || cellXfIndex >= _cellXfs.Count)
        {
            return new ResolvedStyle(null, null);
        }

        var xf = _cellXfs[cellXfIndex];
        return ResolveFromXf(xf);
    }

    private ResolvedStyle ResolveFromXf(CachedCellFormat xf, int depth = 0)
    {
        if (depth > 10) return new ResolvedStyle(null, null);

        var numFmtId = xf.ApplyNumberFormat ? xf.NumFmtId : ResolveAncestorNumFmtId(xf, depth + 1);
        var fontId = xf.ApplyFont ? xf.FontId : ResolveAncestorFontId(xf, depth + 1);
        var fillId = xf.ApplyFill ? xf.FillId : ResolveAncestorFillId(xf, depth + 1);
        var borderId = xf.ApplyBorder ? xf.BorderId : ResolveAncestorBorderId(xf, depth + 1);
        var alignment = xf.ApplyAlignment ? xf.Alignment : ResolveAncestorAlignment(xf, depth + 1);

        var formatCode = GetNumberFormatCode(numFmtId);

        var font = fontId < _fonts.Count ? _fonts[(int)fontId] : null;
        var fill = fillId < _fills.Count ? _fills[(int)fillId] : null;
        var border = borderId < _borders.Count ? _borders[(int)borderId] : null;

        var style = BuildRenderStyle(font, fill, border, alignment, formatCode);

        return new ResolvedStyle(style, formatCode);
    }

    private uint ResolveAncestorNumFmtId(CachedCellFormat xf, int depth)
    {
        if (xf.XfId is { } xfId && xfId < _cellStyleXfs.Count)
        {
            return _cellStyleXfs[(int)xfId].NumFmtId;
        }
        return 0;
    }

    private uint ResolveAncestorFontId(CachedCellFormat xf, int depth)
    {
        if (xf.XfId is { } xfId && xfId < _cellStyleXfs.Count)
        {
            return _cellStyleXfs[(int)xfId].FontId;
        }
        return 0;
    }

    private uint ResolveAncestorFillId(CachedCellFormat xf, int depth)
    {
        if (xf.XfId is { } xfId && xfId < _cellStyleXfs.Count)
        {
            return _cellStyleXfs[(int)xfId].FillId;
        }
        return 0;
    }

    private uint ResolveAncestorBorderId(CachedCellFormat xf, int depth)
    {
        if (xf.XfId is { } xfId && xfId < _cellStyleXfs.Count)
        {
            return _cellStyleXfs[(int)xfId].BorderId;
        }
        return 0;
    }

    private CachedAlignment ResolveAncestorAlignment(CachedCellFormat xf, int depth)
    {
        if (xf.XfId is { } xfId && xfId < _cellStyleXfs.Count)
        {
            return _cellStyleXfs[(int)xfId].Alignment;
        }
        return new CachedAlignment(null, null, false);
    }

    private static XlsxRenderStyle? BuildRenderStyle(
        CachedFont? font,
        CachedFill? fill,
        CachedBorder? border,
        CachedAlignment alignment,
        string? numberFormatCode)
    {
        var hasAnyProperty =
            font?.Name is not null
            || font?.Size is not null
            || font?.Bold == true
            || font?.Italic == true
            || font?.ColorArgb is not null
            || fill?.ColorArgb is not null
            || border?.Left.Style is not null || border?.Right.Style is not null
            || border?.Top.Style is not null || border?.Bottom.Style is not null
            || alignment.Horizontal is not null || alignment.Vertical is not null
            || alignment.WrapText
            || numberFormatCode is not null && numberFormatCode.Length > 0 && numberFormatCode != "General";

        if (!hasAnyProperty) return null;

        return new XlsxRenderStyle
        {
            FontName = font?.Name,
            FontSizePt = font?.Size,
            Bold = font?.Bold ?? false,
            Italic = font?.Italic ?? false,
            FontColorArgb = font?.ColorArgb,
            FillArgb = fill?.ColorArgb,
            BorderLeft = BuildBorderEdge(border?.Left),
            BorderRight = BuildBorderEdge(border?.Right),
            BorderTop = BuildBorderEdge(border?.Top),
            BorderBottom = BuildBorderEdge(border?.Bottom),
            HorizontalAlignment = alignment.Horizontal,
            VerticalAlignment = alignment.Vertical,
            WrapText = alignment.WrapText,
            NumberFormatCode = numberFormatCode
        };
    }

    private static XlsxRenderBorderEdge? BuildBorderEdge(CachedBorderEdge? edge)
    {
        if (edge is null || edge.Style is null) return null;

        return new XlsxRenderBorderEdge
        {
            Style = edge.Style,
            ColorArgb = edge.ColorArgb
        };
    }
}

internal sealed record CachedFont(
    string? Name,
    double? Size,
    bool Bold,
    bool Italic,
    string? ColorArgb);

internal sealed record CachedFill(
    string? ColorArgb);

internal sealed record CachedBorder(
    CachedBorderEdge Left,
    CachedBorderEdge Right,
    CachedBorderEdge Top,
    CachedBorderEdge Bottom);

internal sealed record CachedBorderEdge(
    string? Style,
    string? ColorArgb);

internal sealed record CachedCellFormat(
    uint NumFmtId,
    uint FontId,
    uint FillId,
    uint BorderId,
    bool ApplyNumberFormat,
    bool ApplyFont,
    bool ApplyFill,
    bool ApplyBorder,
    bool ApplyAlignment,
    CachedAlignment Alignment,
    uint? XfId);

internal sealed record CachedAlignment(
    string? Horizontal,
    string? Vertical,
    bool WrapText);

internal sealed record ResolvedStyle(
    XlsxRenderStyle? Style,
    string? NumberFormatCode);
