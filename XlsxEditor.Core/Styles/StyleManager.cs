using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Styles;

/// <summary>
/// Append-only, deduplicated cell-style machinery for a workbook's stylesheet.
///
/// Every named style and every style/number-format combination requested by a cell
/// write resolves to a cellXf index in the workbook stylesheet. Resolution is
/// structural: two styles with identical font/fill/alignment/number-format content
/// share one index, and existing entries loaded from an opened workbook are matched
/// (and therefore reused) rather than recreated. New font/fill/numFmt/cellXf entries
/// are always <em>appended</em>; existing style definitions are never mutated.
///
/// The stylesheet's own collections (numFmts, fonts, fills, borders, cellStyleXfs,
/// cellXfs) are kept in ECMA-376 schema order and their Count attributes stay in sync
/// with the elements they contain.
/// </summary>
internal sealed class StyleManager
{
    private readonly WorkbookPart _workbookPart;

    // name -> (resolved cellXf index, the spec that produced it). Ordinal comparison:
    // style names are case-insensitive in Excel, so "Bold" and "bold" are one name.
    private readonly Dictionary<string, (uint Index, CellStyleSpec Spec)> _namedStyles =
        new(StringComparer.OrdinalIgnoreCase);

    // canonical content key -> cellXf index, populated from the existing stylesheet on
    // first use and extended as new styles are appended. This is what makes identical
    // styles share a single cellXf across define calls and reopened workbooks.
    private readonly Dictionary<string, uint> _cellXfKeys = new(StringComparer.Ordinal);

    // canonical content key -> font index / fill index (dedup within and across appends).
    private readonly Dictionary<string, uint> _fontKeys = new(StringComparer.Ordinal);
    private readonly Dictionary<string, uint> _fillKeys = new(StringComparer.Ordinal);

    // custom number-format code -> numFmt id (ids >= 164), plus the reverse map needed
    // to build content keys from existing cellXfs.
    private readonly Dictionary<string, uint> _numberFormatIds = new(StringComparer.Ordinal);
    private readonly Dictionary<uint, string> _numberFormatCodeById = new();

    private bool _scanned;

    public StyleManager(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
    }

    /// <summary>
    /// Registers a named style and returns its cellXf index. The name must be unique
    /// (case-insensitive); redefining an existing name throws before any stylesheet
    /// mutation. The underlying OOXML entries are deduplicated: an identical style —
    /// whether defined earlier in this session or already present in a reopened
    /// workbook's stylesheet — is reused instead of appended.
    /// </summary>
    public uint DefineStyle(CellStyleSpec style)
    {
        ArgumentNullException.ThrowIfNull(style);
        style.Validate();

        if (_namedStyles.ContainsKey(style.Name))
        {
            throw new XlsxException(
                $"A named style '{style.Name}' is already defined in this workbook. " +
                "Style names must be unique (comparison is case-insensitive, as in Excel); " +
                "redefine the style under a different name or reuse the existing one.");
        }

        EnsureScanned();
        var index = ResolveCellXf(style);
        _namedStyles[style.Name] = (index, style);
        return index;
    }

    /// <summary>Returns the cellXf index of a previously defined named style.</summary>
    public uint GetStyleIndex(string styleName)
    {
        if (_namedStyles.TryGetValue(styleName, out var existing))
        {
            return existing.Index;
        }

        throw new XlsxException(
            $"No named style '{styleName}' is defined in this workbook. " +
            "Define it with DefineStyle before referencing it from a cell write.");
    }

    /// <summary>Names of all styles defined through this manager, in definition order.</summary>
    public IReadOnlyList<string> GetDefinedStyleNames() => _namedStyles.Keys.ToList();

    /// <summary>
    /// Resolves a style name plus an optional number-format override to a cellXf index.
    /// When both are null/empty, returns 0 (the stylesheet's default cell format). The
    /// number-format override wins over a named style's own number format. This is the
    /// signature cell writes use to attach styling.
    /// </summary>
    public uint ResolveStyleIndex(string? styleName, string? numberFormat)
    {
        var hasName = !string.IsNullOrWhiteSpace(styleName);
        var hasFormat = !string.IsNullOrWhiteSpace(numberFormat);
        if (!hasName && !hasFormat)
        {
            return 0;
        }

        EnsureScanned();

        CellStyleSpec spec;
        if (!hasName)
        {
            spec = new CellStyleSpec { NumberFormat = numberFormat };
        }
        else
        {
            if (!_namedStyles.TryGetValue(styleName!, out var named))
            {
                throw new XlsxException(
                    $"No named style '{styleName}' is defined in this workbook. " +
                    "Define it with DefineStyle before referencing it from a cell write.");
            }

            spec = named.Spec with { NumberFormat = hasFormat ? numberFormat : named.Spec.NumberFormat };
        }

        return ResolveCellXf(spec);
    }

    // ─── CellXf resolution ─────────────────────────────────────────

    /// <summary>
    /// Resolves a fully-specified style to a cellXf index, appending the (deduplicated)
    /// font/fill/numFmt/cellXf entries needed and reusing any structurally identical
    /// existing entry. All component resolution happens after the cellXf content key is
    /// computed, so a dedup hit never appends orphaned components.
    /// </summary>
    private uint ResolveCellXf(CellStyleSpec spec)
    {
        var fontKey = BuildFontKey(spec.Font);
        var fillKey = BuildFillKey(spec.Fill);
        var numFmtCode = spec.NumberFormat;
        var key = BuildCellXfKey(fontKey, fillKey, numFmtCode, spec.Alignment);

        if (_cellXfKeys.TryGetValue(key, out var existing))
        {
            return existing;
        }

        var numFmtId = ResolveNumberFormat(numFmtCode);
        var fontId = ResolveFont(fontKey, spec.Font);
        var fillId = ResolveFill(fillKey, spec.Fill);

        var cellFormat = new CellFormat
        {
            NumberFormatId = numFmtId,
            FontId = fontId,
            FillId = fillId,
            BorderId = 0
        };

        if (numFmtCode is not null)
        {
            cellFormat.ApplyNumberFormat = true;
        }

        if (fontKey is not null)
        {
            cellFormat.ApplyFont = true;
        }

        if (fillKey is not null)
        {
            cellFormat.ApplyFill = true;
        }

        if (spec.Alignment is { IsDefault: false } alignment)
        {
            cellFormat.Alignment = BuildAlignment(alignment);
            cellFormat.ApplyAlignment = true;
        }

        var cellFormats = EnsureCellFormats();
        var index = cellFormats.Count?.Value ?? (uint)cellFormats.Elements<CellFormat>().Count();
        cellFormats.Append(cellFormat);
        cellFormats.Count = index + 1;
        _cellXfKeys[key] = index;
        return index;
    }

    // ─── Component resolution ──────────────────────────────────────

    private uint ResolveFont(string? fontKey, CellFontSpec? font)
    {
        if (fontKey is null)
        {
            return 0; // the default font, never appended
        }

        if (_fontKeys.TryGetValue(fontKey, out var existing))
        {
            return existing;
        }

        // The typed property setters insert each child at its correct position in the
        // CT_Font sequence, so the produced <font> is schema-ordered no matter which
        // aspects are set and in what order the caller supplied them.
        var newFont = new Font();
        if (font!.Bold)
        {
            newFont.Bold = new Bold();
        }

        if (font.Italic)
        {
            newFont.Italic = new Italic();
        }

        if (font.ColorArgb is { } color)
        {
            newFont.Color = new Color { Rgb = CellStyleSpec.NormalizeColorArgb(color) };
        }

        if (font.Size is { } size)
        {
            newFont.FontSize = new FontSize { Val = size };
        }

        var fonts = EnsureFonts();
        var id = fonts.Count?.Value ?? (uint)fonts.Elements<Font>().Count();
        fonts.Append(newFont);
        fonts.Count = id + 1;
        _fontKeys[fontKey] = id;
        return id;
    }

    private uint ResolveFill(string? fillKey, CellFillSpec? fill)
    {
        if (fillKey is null)
        {
            return 0; // fill 0 is the schema-required "no fill"
        }

        if (_fillKeys.TryGetValue(fillKey, out var existing))
        {
            return existing;
        }

        var newFill = new Fill(new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new ForegroundColor { Rgb = CellStyleSpec.NormalizeColorArgb(fill!.SolidColorArgb!) }
        });

        var fills = EnsureFills();
        var id = fills.Count?.Value ?? (uint)fills.Elements<Fill>().Count();
        fills.Append(newFill);
        fills.Count = id + 1;
        _fillKeys[fillKey] = id;
        return id;
    }

    private uint ResolveNumberFormat(string? code)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return 0;
        }

        if (BuiltInNumberFormatIds.TryGetValue(code, out var builtIn))
        {
            return builtIn;
        }

        if (_numberFormatIds.TryGetValue(code, out var existing))
        {
            return existing;
        }

        var numFmts = EnsureNumberingFormats();
        var id = NextCustomNumberFormatId(numFmts);
        numFmts.Append(new NumberingFormat { NumberFormatId = id, FormatCode = code });
        numFmts.Count = (uint)numFmts.Elements<NumberingFormat>().Count();
        _numberFormatIds[code] = id;
        _numberFormatCodeById[id] = code;
        return id;
    }

    // ─── Existing-stylesheet scan (open-preservation + dedup) ──────

    /// <summary>
    /// Scans the existing stylesheet once and indexes its fonts, fills, custom number
    /// formats and cellXfs by content. Read-only: existing style definitions are never
    /// modified. The index lets a style defined after an open reuse an identical entry
    /// already in the workbook instead of appending a duplicate.
    /// </summary>
    private void EnsureScanned()
    {
        if (_scanned)
        {
            return;
        }

        _scanned = true;

        var stylesheet = _workbookPart.WorkbookStylesPart?.Stylesheet;
        if (stylesheet is null)
        {
            return;
        }

        foreach (var nf in stylesheet.NumberingFormats?.Elements<NumberingFormat>() ?? Enumerable.Empty<NumberingFormat>())
        {
            if (nf.FormatCode?.Value is { } code && nf.NumberFormatId?.Value is { } id)
            {
                _numberFormatIds.TryAdd(code, id);
                _numberFormatCodeById.TryAdd(id, code);
            }
        }

        var fonts = stylesheet.Fonts?.Elements<Font>().ToList() ?? new List<Font>();
        for (var i = 0; i < fonts.Count; i++)
        {
            if (BuildFontKeyFromFont(fonts[i]) is { } fontKey)
            {
                _fontKeys.TryAdd(fontKey, (uint)i);
            }
        }

        var fills = stylesheet.Fills?.Elements<Fill>().ToList() ?? new List<Fill>();
        for (var i = 0; i < fills.Count; i++)
        {
            if (BuildFillKeyFromFill(fills[i]) is { } fillKey)
            {
                _fillKeys.TryAdd(fillKey, (uint)i);
            }
        }

        var cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToList() ?? new List<CellFormat>();
        for (var i = 0; i < cellFormats.Count; i++)
        {
            if (BuildCellXfKeyFromCellFormat(cellFormats[i], fonts, fills) is { } key)
            {
                _cellXfKeys.TryAdd(key, (uint)i);
            }
        }
    }

    // ─── Content keys ──────────────────────────────────────────────

    private static string? BuildFontKey(CellFontSpec? font)
    {
        if (font is null || font.IsDefault)
        {
            return null;
        }

        var color = font.ColorArgb is null ? "-" : CellStyleSpec.NormalizeColorArgb(font.ColorArgb);
        var size = font.Size is { } sz ? sz.ToString("R", CultureInfo.InvariantCulture) : "-";
        return $"b:{(font.Bold ? 1 : 0)};i:{(font.Italic ? 1 : 0)};c:{color};s:{size}";
    }

    private static string? BuildFillKey(CellFillSpec? fill)
    {
        if (fill is null || fill.IsDefault)
        {
            return null;
        }

        return $"solid:{CellStyleSpec.NormalizeColorArgb(fill.SolidColorArgb!)}";
    }

    private static string BuildCellXfKey(
        string? fontKey, string? fillKey, string? numFmtCode, CellAlignmentSpec? alignment)
    {
        var horizontal = alignment?.Horizontal is null ? "-" : alignment.Horizontal.Value.ToString();
        var vertical = alignment?.Vertical is null ? "-" : alignment.Vertical.Value.ToString();
        var wrap = alignment?.WrapText == true ? 1 : 0;
        return $"font:{fontKey ?? "-"}|fill:{fillKey ?? "-"}|num:{numFmtCode ?? "-"}|" +
               $"align:{horizontal};{vertical};{wrap}";
    }

    private static string? BuildFontKeyFromFont(Font font)
    {
        var bold = font.Elements<Bold>().Any();
        var italic = font.Elements<Italic>().Any();
        var color = font.Elements<Color>().FirstOrDefault()?.Rgb?.Value;
        var size = font.Elements<FontSize>().FirstOrDefault()?.Val?.Value;

        if (!bold && !italic && color is null && size is null)
        {
            return null;
        }

        var normalizedColor = color is null ? "-" : color.ToUpperInvariant();
        var normalizedSize = size is { } sz ? sz.ToString("R", CultureInfo.InvariantCulture) : "-";
        return $"b:{(bold ? 1 : 0)};i:{(italic ? 1 : 0)};c:{normalizedColor};s:{normalizedSize}";
    }

    private static string? BuildFillKeyFromFill(Fill fill)
    {
        var patternFill = fill.PatternFill;
        if (patternFill?.PatternType?.Value != PatternValues.Solid)
        {
            return null;
        }

        var color = patternFill.ForegroundColor?.Rgb?.Value;
        return color is null ? null : $"solid:{color.ToUpperInvariant()}";
    }

    private string? BuildCellXfKeyFromCellFormat(CellFormat xf, List<Font> fonts, List<Fill> fills)
    {
        string? fontKey = null;
        if (xf.FontId?.Value is { } fontId && fontId < fonts.Count)
        {
            fontKey = BuildFontKeyFromFont(fonts[(int)fontId]);
        }

        string? fillKey = null;
        if (xf.FillId?.Value is { } fillId && fillId < fills.Count)
        {
            fillKey = BuildFillKeyFromFill(fills[(int)fillId]);
        }

        string? numFmtCode = null;
        if (xf.NumberFormatId?.Value is { } numFmtId)
        {
            // numFmtId 0 is "General", the default; it must key as "-" so a plain cellXf
            // dedups against a style spec that sets no number format.
            if (numFmtId == 0)
            {
                numFmtCode = null;
            }
            else if (numFmtId >= 164)
            {
                _numberFormatCodeById.TryGetValue(numFmtId, out numFmtCode);
            }
            else
            {
                BuiltInNumberFormatCodes.TryGetValue(numFmtId, out numFmtCode);
            }
        }

        var alignment = xf.Alignment;
        var horizontal = alignment?.Horizontal is null ? "-" : alignment.Horizontal.Value.ToString();
        var vertical = alignment?.Vertical is null ? "-" : alignment.Vertical.Value.ToString();
        var wrap = alignment?.WrapText?.Value == true ? 1 : 0;

        return $"font:{fontKey ?? "-"}|fill:{fillKey ?? "-"}|num:{numFmtCode ?? "-"}|" +
               $"align:{horizontal};{vertical};{wrap}";
    }

    private static Alignment BuildAlignment(CellAlignmentSpec alignment)
    {
        var result = new Alignment();
        if (alignment.Horizontal is { } horizontal)
        {
            result.Horizontal = horizontal;
        }

        if (alignment.Vertical is { } vertical)
        {
            result.Vertical = vertical;
        }

        if (alignment.WrapText)
        {
            result.WrapText = true;
        }

        return result;
    }

    // ─── Stylesheet plumbing ───────────────────────────────────────

    /// <summary>
    /// Returns the workbook's stylesheet, creating the standard Excel baseline (default
    /// font, the two required fills, one empty border, one default cellStyleXf and one
    /// default cellXf) when the workbook has none. On a reopened workbook the existing
    /// stylesheet is returned untouched.
    /// </summary>
    private Stylesheet EnsureStylesheet()
    {
        var part = _workbookPart.WorkbookStylesPart ?? _workbookPart.AddNewPart<WorkbookStylesPart>();
        if (part.Stylesheet is not null)
        {
            return part.Stylesheet;
        }

        var stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            ) { Count = 2 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(new CellFormat()) { Count = 1 });

        part.Stylesheet = stylesheet;
        return stylesheet;
    }

    private NumberingFormats EnsureNumberingFormats()
    {
        var stylesheet = EnsureStylesheet();
        if (stylesheet.NumberingFormats is { } existing)
        {
            return existing;
        }

        var numFmts = new NumberingFormats();
        InsertStylesheetElementAtSchemaPosition(stylesheet, numFmts);
        return numFmts;
    }

    private Fonts EnsureFonts()
    {
        var stylesheet = EnsureStylesheet();
        if (stylesheet.Fonts is { } existing)
        {
            return existing;
        }

        var fonts = new Fonts();
        InsertStylesheetElementAtSchemaPosition(stylesheet, fonts);
        return fonts;
    }

    private Fills EnsureFills()
    {
        var stylesheet = EnsureStylesheet();
        if (stylesheet.Fills is { } existing)
        {
            return existing;
        }

        var fills = new Fills();
        InsertStylesheetElementAtSchemaPosition(stylesheet, fills);
        return fills;
    }

    private CellFormats EnsureCellFormats()
    {
        var stylesheet = EnsureStylesheet();
        if (stylesheet.CellFormats is { } existing)
        {
            return existing;
        }

        var cellFormats = new CellFormats();
        InsertStylesheetElementAtSchemaPosition(stylesheet, cellFormats);
        return cellFormats;
    }

    /// <summary>
    /// Inserts a new stylesheet child at its ECMA-376 position in the CT_Stylesheet
    /// sequence (numFmts, fonts, fills, borders, cellStyleXfs, cellXfs, …). Existing
    /// children are already in schema order, so inserting before the first child that
    /// must follow is correct for both fresh and reopened stylesheets.
    /// </summary>
    private static void InsertStylesheetElementAtSchemaPosition(Stylesheet stylesheet, OpenXmlElement element)
    {
        var order = StylesheetChildOrder.GetValueOrDefault(element.LocalName, int.MaxValue);
        foreach (var child in stylesheet.ChildElements)
        {
            if (StylesheetChildOrder.GetValueOrDefault(child.LocalName, int.MaxValue) > order)
            {
                stylesheet.InsertBefore(element, child);
                return;
            }
        }

        stylesheet.Append(element);
    }

    private static uint NextCustomNumberFormatId(NumberingFormats numFmts)
    {
        uint max = 163;
        foreach (var nf in numFmts.Elements<NumberingFormat>())
        {
            if (nf.NumberFormatId?.Value is { } id && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    /// <summary>
    /// Ordinal positions of the stylesheet children in the CT_Stylesheet sequence
    /// (ECMA-376 §18.8.39). Elements not listed are treated as following every known
    /// child — the safe side for an element we never append.
    /// </summary>
    private static readonly Dictionary<string, int> StylesheetChildOrder = new()
    {
        ["numFmts"] = 1,
        ["fonts"] = 2,
        ["fills"] = 3,
        ["borders"] = 4,
        ["cellStyleXfs"] = 5,
        ["cellXfs"] = 6,
        ["cellStyles"] = 7,
        ["dxfs"] = 8,
        ["tableStyles"] = 9,
        ["colors"] = 10,
        ["extLst"] = 11
    };

    /// <summary>
    /// Standard Excel number-format codes that map to a built-in numFmt id (ECMA-376
    /// §18.8.30, built-in formats 1-49). Codes not listed here become custom numFmts
    /// with an id &gt;= 164.
    /// </summary>
    private static readonly Dictionary<string, uint> BuiltInNumberFormatIds = new(StringComparer.Ordinal)
    {
        ["General"] = 0,
        ["0"] = 1,
        ["0.00"] = 2,
        ["#,##0"] = 3,
        ["#,##0.00"] = 4,
        ["$#,##0_);($#,##0)"] = 5,
        ["$#,##0_);[Red]($#,##0)"] = 6,
        ["$#,##0.00_);($#,##0.00)"] = 7,
        ["$#,##0.00_);[Red]($#,##0.00)"] = 8,
        ["0%"] = 9,
        ["0.00%"] = 10,
        ["0.00E+00"] = 11,
        ["# ?/?"] = 12,
        ["# ??/??"] = 13,
        ["mm-dd-yy"] = 14,
        ["d-mmm-yy"] = 15,
        ["d-mmm"] = 16,
        ["mmm-yy"] = 17,
        ["h:mm AM/PM"] = 18,
        ["h:mm:ss AM/PM"] = 19,
        ["h:mm"] = 20,
        ["h:mm:ss"] = 21,
        ["m/d/yy h:mm"] = 22,
        ["mm:ss"] = 45,
        ["[h]:mm:ss"] = 46,
        ["mmss.0"] = 47,
        ["##0.0E+0"] = 48,
        ["@"] = 49
    };

    private static readonly Dictionary<uint, string> BuiltInNumberFormatCodes =
        BuiltInNumberFormatIds.ToDictionary(pair => pair.Value, pair => pair.Key);
}
