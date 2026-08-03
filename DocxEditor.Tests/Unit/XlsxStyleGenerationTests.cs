using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Focused tests for the JSON→XLSX style vocabulary: border styles and colors, fill
/// patterns, the named-color contract, structural dedup (identical styles/borders share
/// one index), and reopen round-trips. Every generated package is validated with
/// OpenXmlValidator, so accepted styles are always representable — validation and
/// generation can never disagree.
/// </summary>
public class XlsxStyleGenerationTests
{
    private static SpreadsheetDocument OpenBytes(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);
        return SpreadsheetDocument.Open(stream, false);
    }

    private static void AssertNoValidationErrors(byte[] bytes)
    {
        using var document = OpenBytes(bytes);
        OpenXmlAssert.NoValidationErrors(document);
    }

    private static string Errors(XlsxGenerateResult result) =>
        string.Join("; ", result.Validation.Errors.Select(e => e.Message));

    [Theory]
    [InlineData("thin")]
    [InlineData("medium")]
    [InlineData("thick")]
    [InlineData("double")]
    [InlineData("dashed")]
    [InlineData("dotted")]
    [InlineData("dashDot")]
    [InlineData("dashDotDot")]
    [InlineData("hair")]
    [InlineData("mediumDashed")]
    [InlineData("mediumDashDot")]
    [InlineData("mediumDashDotDot")]
    [InlineData("slantDashDot")]
    public void Generate_ShouldAcceptEverySupportedBorderStyle(string style)
    {
        var result = XlsxGenerator.Generate($$"""
        {
            "version": "1.0",
            "styles": [{ "name": "boxed", "border": { "top": { "style": "{{style}}" } } }],
            "worksheets": [{ "name": "S", "cells": [{ "address": "A1", "value": "x", "style": "boxed" }] }]
        }
        """);

        Assert.True(result.IsValid, Errors(result));
        Assert.NotNull(result.Bytes);
        AssertNoValidationErrors(result.Bytes!);
    }

    [Theory]
    [InlineData("solid")]
    [InlineData("gray125")]
    [InlineData("gray0625")]
    [InlineData("darkgray")]
    [InlineData("mediumgray")]
    [InlineData("lightgray")]
    [InlineData("darkhorizontal")]
    [InlineData("darkvertical")]
    [InlineData("darkdown")]
    [InlineData("darkup")]
    [InlineData("darkgrid")]
    [InlineData("darktrellis")]
    [InlineData("lighthorizontal")]
    [InlineData("lightvertical")]
    [InlineData("lightdown")]
    [InlineData("lightup")]
    [InlineData("lightgrid")]
    [InlineData("lighttrellis")]
    public void Generate_ShouldAcceptEverySupportedFillPattern(string pattern)
    {
        var result = XlsxGenerator.Generate($$"""
        {
            "version": "1.0",
            "styles": [{ "name": "filled", "fill": { "color": "FF0000", "pattern": "{{pattern}}" } }],
            "worksheets": [{ "name": "S", "cells": [{ "address": "A1", "value": "x", "style": "filled" }] }]
        }
        """);

        Assert.True(result.IsValid, Errors(result));
        Assert.NotNull(result.Bytes);
        AssertNoValidationErrors(result.Bytes!);
    }

    [Fact]
    public void Generate_ShouldApplyBorderStylesAndColors()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "styles": [{
                "name": "boxed",
                "border": {
                    "top": { "style": "thin", "color": "000000" },
                    "bottom": { "style": "medium" },
                    "left": { "style": "dashed" },
                    "right": { "style": "dotted" }
                }
            }],
            "worksheets": [{ "name": "S", "cells": [{ "address": "A1", "value": "x", "style": "boxed" }] }]
        }
        """);

        Assert.True(result.IsValid, Errors(result));
        AssertNoValidationErrors(result.Bytes!);

        using var document = OpenBytes(result.Bytes!);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(2u, stylesheet.Borders!.Count!.Value);

        var border = stylesheet.Borders.Elements<Border>().Skip(1).Single();
        Assert.Equal(BorderStyleValues.Thin, border.TopBorder!.Style!.Value);
        Assert.Equal("FF000000", border.TopBorder.Color!.Rgb!.Value);
        Assert.Equal(BorderStyleValues.Medium, border.BottomBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Dashed, border.LeftBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Dotted, border.RightBorder!.Style!.Value);

        var cell = document.WorkbookPart!.WorksheetParts.First().Worksheet!
            .Descendants<Cell>().Single(c => c.CellReference == "A1");
        var cellXf = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal(1u, cellXf.BorderId!.Value);
        Assert.True(cellXf.ApplyBorder!.Value);
    }

    [Fact]
    public void Generate_ShouldApplyNonSolidFillPatterns()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "styles": [
                { "name": "striped", "fill": { "color": "FF0000", "pattern": "gray125" } },
                { "name": "hatched", "fill": { "color": "FF00FF", "pattern": "darkdown" } }
            ],
            "worksheets": [{ "name": "S", "cells": [
                { "address": "A1", "value": "x", "style": "striped" },
                { "address": "A2", "value": "y", "style": "hatched" }
            ] }]
        }
        """);

        Assert.True(result.IsValid, Errors(result));
        AssertNoValidationErrors(result.Bytes!);

        using var document = OpenBytes(result.Bytes!);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        // fills 0 (none) and 1 (baseline gray125) plus the two explicit fills.
        Assert.Equal(4u, stylesheet.Fills!.Count!.Value);

        var striped = stylesheet.Fills.Elements<Fill>().ElementAt(2);
        Assert.Equal(PatternValues.Gray125, striped.PatternFill!.PatternType!.Value);
        Assert.Equal("FFFF0000", striped.PatternFill.ForegroundColor!.Rgb!.Value);

        var hatched = stylesheet.Fills.Elements<Fill>().ElementAt(3);
        Assert.Equal(PatternValues.DarkDown, hatched.PatternFill!.PatternType!.Value);
        Assert.Equal("FFFF00FF", hatched.PatternFill.ForegroundColor!.Rgb!.Value);
    }

    [Fact]
    public void Generate_ShouldNormalizeNamedColors_ToHex()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "styles": [{
                "name": "named",
                "font": { "color": "white" },
                "fill": { "color": "red", "pattern": "solid" },
                "border": { "top": { "style": "thin", "color": "darkgray" } }
            }],
            "worksheets": [{ "name": "S", "cells": [{ "address": "A1", "value": "x", "style": "named" }] }]
        }
        """);

        Assert.True(result.IsValid, Errors(result));
        AssertNoValidationErrors(result.Bytes!);

        using var document = OpenBytes(result.Bytes!);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal("FFFFFFFF", stylesheet.Fonts!.Elements<Font>().Skip(1).Single().Color!.Rgb!.Value);
        Assert.Equal("FFFF0000", stylesheet.Fills!.Elements<Fill>().Skip(2).Single().PatternFill!.ForegroundColor!.Rgb!.Value);
        Assert.Equal("FFA9A9A9", stylesheet.Borders!.Elements<Border>().Skip(1).Single().TopBorder!.Color!.Rgb!.Value);
    }

    [Fact]
    public void Generate_ShouldDeduplicateStructurallyIdenticalBorderStyles()
    {
        var result = XlsxGenerator.Generate("""
        {
            "version": "1.0",
            "styles": [
                { "name": "b1", "border": { "top": { "style": "thin", "color": "000000" } } },
                { "name": "b2", "border": { "top": { "style": "thin", "color": "000000" } } }
            ],
            "worksheets": [{ "name": "S", "cells": [
                { "address": "A1", "value": "x", "style": "b1" },
                { "address": "A2", "value": "y", "style": "b2" }
            ] }]
        }
        """);

        Assert.True(result.IsValid, Errors(result));
        AssertNoValidationErrors(result.Bytes!);

        using var document = OpenBytes(result.Bytes!);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        // Default border + one shared bordered style: identical content must not duplicate.
        Assert.Equal(2u, stylesheet.Borders!.Count!.Value);
        Assert.Equal(2u, stylesheet.CellFormats!.Count!.Value);

        var cells = document.WorkbookPart!.WorksheetParts.First().Worksheet!
            .Descendants<Cell>().OrderBy(c => c.CellReference!.Value).ToList();
        Assert.Equal(cells[0].StyleIndex, cells[1].StyleIndex);
        Assert.NotEqual(0u, cells[0].StyleIndex!.Value);
    }

    [Fact]
    public void OpenReopen_ShouldReuseExistingBorderAndFillStyles_ByContent()
    {
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "boxed",
                Border = new CellBorderSpec { Top = new CellEdgeSpec { Style = BorderStyleValues.Thin, ColorArgb = "000000" } },
                Fill = new CellFillSpec { Pattern = PatternValues.Solid, SolidColorArgb = "FF0000" }
            });
            bytes = builder.SaveToBytes();
        }

        using var reopened = WorkbookBuilder.Open(bytes);
        var reusedIndex = reopened.DefineStyle(new CellStyleSpec
        {
            Name = "again",
            Border = new CellBorderSpec { Top = new CellEdgeSpec { Style = BorderStyleValues.Thin, ColorArgb = "000000" } },
            Fill = new CellFillSpec { Pattern = PatternValues.Solid, SolidColorArgb = "FF0000" }
        });
        // Reuses the single existing bordered cellXf (index 1); nothing is appended.
        Assert.Equal(1u, reusedIndex);

        var saved = reopened.SaveToBytes();
        using var document = OpenBytes(saved);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(2u, stylesheet.Borders!.Count!.Value);
        Assert.Equal(3u, stylesheet.Fills!.Count!.Value);
        Assert.Equal(2u, stylesheet.CellFormats!.Count!.Value);
    }
}
