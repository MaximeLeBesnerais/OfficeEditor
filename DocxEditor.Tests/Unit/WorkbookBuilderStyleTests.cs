using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

public class WorkbookBuilderStyleTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_xlsx_styles_{Guid.NewGuid()}.xlsx");

    // ─── Basic definition & stylesheet growth ─────────────────────

    [Fact]
    public void DefineStyle_ReturnsIndexAndAppendsDedupComponents()
    {
        // Act: a bold+italic+color+size style on a fresh workbook
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            var index = builder.DefineStyle(new CellStyleSpec
            {
                Name = "Title",
                Font = new CellFontSpec { Bold = true, Italic = true, ColorArgb = "FF0000", Size = 14 }
            });
            Assert.Equal(1U, index);
            builder.Save();
        }

        // Assert: baseline (font, 2 fills, border, cellStyleXf, cellXf) + one new font + one new cellXf
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.NotNull(stylesheet);
        Assert.Equal(2U, stylesheet.Fonts!.Count?.Value);
        Assert.Equal(2U, stylesheet.Fills!.Count?.Value);
        Assert.Equal(1U, stylesheet.Borders!.Count?.Value);
        Assert.Equal(1U, stylesheet.CellStyleFormats!.Count?.Value);
        Assert.Equal(2U, stylesheet.CellFormats!.Count?.Value);

        var font = stylesheet.Fonts.Elements<Font>().Last();
        Assert.NotNull(font.GetFirstChild<Bold>());
        Assert.NotNull(font.GetFirstChild<Italic>());
        Assert.Equal("FFFF0000", font.GetFirstChild<Color>()?.Rgb?.Value);
        Assert.Equal(14.0, font.GetFirstChild<FontSize>()?.Val?.Value);

        var cellFormat = stylesheet.CellFormats.Elements<CellFormat>().Last();
        Assert.Equal(1U, cellFormat.FontId?.Value);
        Assert.Equal(0U, cellFormat.FillId?.Value);
        Assert.Equal(0U, cellFormat.BorderId?.Value);
        Assert.True(cellFormat.ApplyFont?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void DefineStyle_DuplicateName_ThrowsBeforeAnyMutation()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.DefineStyle(new CellStyleSpec { Name = "Bold", Font = new CellFontSpec { Bold = true } });

        // Same name, even with identical spec, is a duplicate definition.
        var ex = Assert.Throws<XlsxException>(() =>
            builder.DefineStyle(new CellStyleSpec { Name = "bold", Font = new CellFontSpec { Bold = true } }));
        Assert.Contains("already defined", ex.Message);
        Assert.Contains("bold", ex.Message);

        // Nothing was appended for the rejected duplicate: still one extra font + one cellXf.
        Assert.Single(builder.GetDefinedStyleNames());
        Assert.Equal(1U, builder.GetStyleIndex("Bold"));
    }

    [Fact]
    public void DefineStyle_TwoNamesIdenticalSpec_ShareOneCellXf()
    {
        // Act: two differently-named styles with identical content
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            var a = builder.DefineStyle(new CellStyleSpec
            {
                Name = "RedBold",
                Font = new CellFontSpec { Bold = true, ColorArgb = "FF0000" }
            });
            var b = builder.DefineStyle(new CellStyleSpec
            {
                Name = "RedBoldTitle",
                Font = new CellFontSpec { Bold = true, ColorArgb = "FF0000" }
            });
            Assert.Equal(a, b);
            builder.Save();
        }

        // Assert: one shared font + one shared cellXf, no duplicates
        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(2U, stylesheet.Fonts!.Count?.Value);
        Assert.Equal(2U, stylesheet.CellFormats!.Count?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void DefineStyle_Color6Digit_NormalizesToArgb8()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "Six",
                Fill = new CellFillSpec { SolidColorArgb = "FFC000" }
            });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var fill = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fills!.Elements<Fill>().Last();
        Assert.Equal("FFFFC000", fill.PatternFill?.ForegroundColor?.Rgb?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void DefineStyle_FullStyle_EmitsSchemaOrderedOoxml()
    {
        // Act: every supported aspect at once
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "Money",
                Font = new CellFontSpec { Bold = true, ColorArgb = "00FF00", Size = 11.5 },
                Fill = new CellFillSpec { SolidColorArgb = "FFFFC000" },
                Alignment = new CellAlignmentSpec { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true },
                NumberFormat = "0.00%"
            });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;

        // Font child order is schema order: b, sz, color
        var font = stylesheet.Fonts!.Elements<Font>().Last();
        var fontChildren = font.ChildElements.Select(c => c.LocalName).ToList();
        Assert.Equal(new[] { "b", "sz", "color" }, fontChildren);

        // Solid fill with fgColor
        var fill = stylesheet.Fills!.Elements<Fill>().Last();
        Assert.Equal(PatternValues.Solid, fill.PatternFill?.PatternType?.Value);
        Assert.Equal("FFFFC000", fill.PatternFill?.ForegroundColor?.Rgb?.Value);

        // Alignment
        var cellFormat = stylesheet.CellFormats!.Elements<CellFormat>().Last();
        Assert.Equal(HorizontalAlignmentValues.Center, cellFormat.Alignment?.Horizontal?.Value);
        Assert.Equal(VerticalAlignmentValues.Center, cellFormat.Alignment?.Vertical?.Value);
        Assert.True(cellFormat.Alignment?.WrapText?.Value);
        Assert.True(cellFormat.ApplyAlignment?.Value);

        // "0.00%" maps to built-in numFmt id 10; no custom numFmts created
        Assert.Equal(10U, cellFormat.NumberFormatId?.Value);
        Assert.True(cellFormat.ApplyNumberFormat?.Value);
        Assert.Null(stylesheet.NumberingFormats);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void DefineStyle_CustomNumberFormat_RegistersNumFmtId164()
    {
        // Act: a format with no built-in id must become a custom numFmt
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "IsoDate",
                NumberFormat = "yyyy-mm-dd"
            });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var numFmt = Assert.Single(stylesheet.NumberingFormats!.Elements<NumberingFormat>());
        Assert.Equal(164U, numFmt.NumberFormatId?.Value);
        Assert.Equal("yyyy-mm-dd", numFmt.FormatCode?.Value);

        var cellFormat = stylesheet.CellFormats!.Elements<CellFormat>().Last();
        Assert.Equal(164U, cellFormat.NumberFormatId?.Value);
        Assert.True(cellFormat.ApplyNumberFormat?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void DefineStyle_SameCustomFormat_TwoStyles_ReuseOneNumFmt()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec { Name = "A", NumberFormat = "0.000" });
            builder.DefineStyle(new CellStyleSpec { Name = "B", NumberFormat = "0.000" });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var numFmts = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats;
        Assert.NotNull(numFmts);
        Assert.Single(numFmts.Elements<NumberingFormat>());
        Assert.Equal(164U, numFmts.Elements<NumberingFormat>().Single().NumberFormatId?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void DefineStyle_SubsequentCustomFormats_IncrementNumFmtId()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec { Name = "A", NumberFormat = "0.000" });
            builder.DefineStyle(new CellStyleSpec { Name = "B", NumberFormat = "0.0000" });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var ids = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
            .Elements<NumberingFormat>().Select(n => n.NumberFormatId!.Value).ToList();
        Assert.Equal(new[] { 164U, 165U }, ids);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ─── Stylesheet schema order ───────────────────────────────────

    [Fact]
    public void DefineStyle_WithNumberFormat_KeepsStylesheetChildrenInSchemaOrder()
    {
        using (var builder = WorkbookBuilder.Create(_testFilePath))
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "All",
                Font = new CellFontSpec { Bold = true },
                Fill = new CellFillSpec { SolidColorArgb = "FFFF0000" },
                NumberFormat = "0.000",
                Alignment = new CellAlignmentSpec { Horizontal = HorizontalAlignmentValues.Right }
            });
            builder.Save();
        }

        using var doc = SpreadsheetDocument.Open(_testFilePath, false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var childNames = stylesheet.ChildElements.Select(c => c.LocalName).ToList();
        // numFmts must come before fonts, which come before fills, before borders, etc.
        var order = childNames.ToDictionary(n => n, n => childNames.IndexOf(n));
        Assert.True(order["numFmts"] < order["fonts"], $"got {string.Join(",", childNames)}");
        Assert.True(order["fonts"] < order["fills"], $"got {string.Join(",", childNames)}");
        Assert.True(order["fills"] < order["borders"], $"got {string.Join(",", childNames)}");
        Assert.True(order["borders"] < order["cellStyleXfs"], $"got {string.Join(",", childNames)}");
        Assert.True(order["cellStyleXfs"] < order["cellXfs"], $"got {string.Join(",", childNames)}");
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ─── Open-existing preservation & dedup ────────────────────────

    [Fact]
    public void ReopenWorkbook_DefineIdenticalStyle_ReusesExistingEntries_WithoutMutation()
    {
        // Arrange: a workbook with one bold style
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec { Name = "Bold", Font = new CellFontSpec { Bold = true } });
            bytes = builder.SaveToBytes();
        }

        var originalCellXfXml = LastCellXfXmlOf(bytes);

        // Act: reopen and define a structurally identical style under another name
        using (var builder = WorkbookBuilder.Open(bytes))
        {
            var index = builder.DefineStyle(new CellStyleSpec { Name = "BoldTitle", Font = new CellFontSpec { Bold = true } });
            Assert.Equal(1U, index); // reuses the existing bold cellXf, does not append
            bytes = builder.SaveToBytes();
        }

        // Assert: no new font/cellXf appeared and the existing definition is untouched
        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(2U, stylesheet.Fonts!.Count?.Value);
        Assert.Equal(2U, stylesheet.CellFormats!.Count?.Value);
        Assert.Equal(originalCellXfXml, stylesheet.CellFormats.Elements<CellFormat>().Last().OuterXml);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    [Fact]
    public void ReopenWorkbook_DefineNewStyle_AppendsAfterExistingWithoutMutating()
    {
        // Arrange: a workbook authored with an external-style stylesheet (default + 8 cellXfs)
        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new Fonts(new Font()) { Count = 1 },
                    new Fills(
                        new Fill(new PatternFill { PatternType = PatternValues.None }),
                        new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                    ) { Count = 2 },
                    new Borders(new Border()) { Count = 1 },
                    new CellStyleFormats(new CellFormat()) { Count = 1 },
                    new CellFormats(
                        Enumerable.Range(0, 8).Select(_ =>
                            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 })
                    ) { Count = 8 });
                workbookPart.Workbook = new Workbook(new Sheets(new Sheet
                {
                    Name = "Sheet1",
                    SheetId = 1u,
                    Id = workbookPart.GetIdOfPart(worksheetPart)
                }));
                doc.Save();
            }
            bytes = ms.ToArray();
        }

        // Act: reopen, define a brand-new style
        using (var builder = WorkbookBuilder.Open(bytes))
        {
            var index = builder.DefineStyle(new CellStyleSpec
            {
                Name = "Fresh",
                Font = new CellFontSpec { Bold = true },
                Fill = new CellFillSpec { SolidColorArgb = "FFFFC000" }
            });
            Assert.Equal(8U, index); // appended after the 8 existing cellXfs
            bytes = builder.SaveToBytes();
        }

        // Assert: existing entries preserved, new entries appended at the end
        using var reopened = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var stylesheet = reopened.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.Equal(9U, stylesheet.CellFormats!.Count?.Value);
        Assert.Equal(2U, stylesheet.Fonts!.Count?.Value);
        Assert.Equal(3U, stylesheet.Fills!.Count?.Value);
        // The 8th original cellXf (index 7) is untouched: still the default font id 0
        Assert.Equal(0U, stylesheet.CellFormats.Elements<CellFormat>().ElementAt(7).FontId?.Value);
        var fresh = stylesheet.CellFormats.Elements<CellFormat>().Last();
        Assert.Equal(1U, fresh.FontId?.Value);
        Assert.Equal(2U, fresh.FillId?.Value);
        OpenXmlAssert.NoValidationErrors(reopened);
    }

    [Fact]
    public void ReopenWorkbook_DefineMatchingExistingComplexStyle_ReusesCellXf()
    {
        // A reopened workbook that already contains the exact style we define (solid fill
        // + centered alignment) must reuse it rather than append a duplicate.
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec
            {
                Name = "Existing",
                Fill = new CellFillSpec { SolidColorArgb = "FFFFC000" },
                Alignment = new CellAlignmentSpec { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            });
            bytes = builder.SaveToBytes();
        }

        using (var builder = WorkbookBuilder.Open(bytes))
        {
            var index = builder.DefineStyle(new CellStyleSpec
            {
                Name = "ExistingAgain",
                Fill = new CellFillSpec { SolidColorArgb = "FFFFC000" },
                Alignment = new CellAlignmentSpec { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            });
            Assert.Equal(1U, index);
            bytes = builder.SaveToBytes();
        }

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var stylesheet = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        // The first session appended one solid fill and one cellXf (indices 2 and 1);
        // reopening and re-defining the same spec must NOT append any more.
        Assert.Equal(3U, stylesheet.Fills!.Count?.Value);
        Assert.Equal(2U, stylesheet.CellFormats!.Count?.Value);
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ─── Resolve/read API ──────────────────────────────────────────

    [Fact]
    public void GetStyleIndex_UnknownName_ThrowsXlsxException()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var ex = Assert.Throws<XlsxException>(() => builder.GetStyleIndex("Missing"));
        Assert.Contains("Missing", ex.Message);
        Assert.Contains("DefineStyle", ex.Message);
    }

    [Fact]
    public void GetStyleIndex_CaseInsensitive_ReturnsIndex()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.DefineStyle(new CellStyleSpec { Name = "Bold", Font = new CellFontSpec { Bold = true } });
        Assert.Equal(builder.GetStyleIndex("Bold"), builder.GetStyleIndex("bold"));
    }

    [Fact]
    public void GetDefinedStyleNames_ReturnsDefinitionOrder()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        builder.DefineStyle(new CellStyleSpec { Name = "First", NumberFormat = "0.00" });
        builder.DefineStyle(new CellStyleSpec { Name = "Second", Font = new CellFontSpec { Italic = true } });
        Assert.Equal(new[] { "First", "Second" }, builder.GetDefinedStyleNames());
    }

    [Fact]
    public void StyleIndexResolvesAfterReopen_ThroughCellWrite()
    {
        // Named styles are a builder-session concept; re-defining the same spec after a
        // reopen is cheap (structural dedup reuses the existing cellXf) and keeps the
        // stylesheet free of duplicate entries.
        byte[] bytes;
        using (var builder = WorkbookBuilder.Create())
        {
            builder.AddWorksheet("Sheet1");
            builder.DefineStyle(new CellStyleSpec { Name = "Pct", NumberFormat = "0.00%" });
            bytes = builder.SaveToBytes();
        }

        using var builder2 = WorkbookBuilder.Open(bytes);
        builder2.DefineStyle(new CellStyleSpec { Name = "Pct", NumberFormat = "0.00%" });
        var sheet = builder2.AddWorksheet("Sheet2");
        sheet.AddCellNumber("A1", 0.123, styleName: "Pct");
        bytes = builder2.SaveToBytes();

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var cell = doc.WorkbookPart!.WorksheetParts
            .SelectMany(wp => wp.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>())
            .SelectMany(r => r.Elements<Cell>())
            .Single(c => c.CellReference!.Value == "A1");
        Assert.Equal(1U, cell.StyleIndex?.Value);
        var cellFormat = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!
            .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal(10U, cellFormat.NumberFormatId?.Value); // built-in "0.00%"
        OpenXmlAssert.NoValidationErrors(doc);
    }

    // ─── Validation ────────────────────────────────────────────────

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void DefineStyle_InvalidName_Throws(string? name)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var ex = Assert.Throws<XlsxException>(() =>
            builder.DefineStyle(new CellStyleSpec { Name = name!, Font = new CellFontSpec { Bold = true } }));
        Assert.Contains("name", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(builder.GetDefinedStyleNames());
    }

    [Theory]
    [InlineData("FF000")]
    [InlineData("FF00001")]
    [InlineData("ZZZZZZ")]
    [InlineData("Ff000")]   // invalid length, though hex digits are case-insensitive
    public void DefineStyle_InvalidColor_Throws(string color)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var ex = Assert.Throws<XlsxException>(() =>
            builder.DefineStyle(new CellStyleSpec { Name = "Bad", Font = new CellFontSpec { ColorArgb = color } }));
        Assert.Contains("color", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(builder.GetDefinedStyleNames());
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(409.6)]
    public void DefineStyle_InvalidFontSize_Throws(double size)
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var ex = Assert.Throws<XlsxException>(() =>
            builder.DefineStyle(new CellStyleSpec { Name = "Bad", Font = new CellFontSpec { Size = size } }));
        Assert.Contains("size", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(builder.GetDefinedStyleNames());
    }

    [Fact]
    public void DefineStyle_WhitespaceNumberFormat_Throws()
    {
        using var builder = WorkbookBuilder.Create(_testFilePath);
        var ex = Assert.Throws<XlsxException>(() =>
            builder.DefineStyle(new CellStyleSpec { Name = "Bad", NumberFormat = "  " }));
        Assert.Contains("number format", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(builder.GetDefinedStyleNames());
    }

    private static string LastCellXfXmlOf(byte[] bytes)
    {
        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        return doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().Last().OuterXml;
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
    }
}
