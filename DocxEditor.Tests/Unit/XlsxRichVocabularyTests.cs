using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Covers the additive rich JSON vocabulary: workbook metadata, named style definitions,
/// worksheet layout (startRow, headerStyle, typed columns, row heights, merges, freeze
/// panes, autofilter, tables) and typed discrete cells.
/// </summary>
public class XlsxRichVocabularyTests
{
    [Fact]
    public void Parse_ShouldDeserializeWorkbookMetadata()
    {
        var json = """
        {
            "version": "1.0",
            "metadata": {
                "title": "Q1 Report",
                "subject": "Quarterly",
                "author": "Acme",
                "category": "Finance",
                "keywords": "budget,revenue",
                "comments": "Draft"
            },
            "worksheets": [{"name": "S", "rows": [["1"]]}]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        Assert.NotNull(set.Metadata);
        Assert.Equal("Q1 Report", set.Metadata!.Title);
        Assert.Equal("Acme", set.Metadata.Author);
        Assert.Equal("Finance", set.Metadata.Category);
    }

    [Fact]
    public void Parse_ShouldDeserializeNamedStyles_WithFullDefinition()
    {
        var json = """
        {
            "version": "1.0",
            "styles": [
                {
                    "name": "Header",
                    "font": { "bold": true, "italic": false, "color": "FFFFFF", "size": 12 },
                    "fill": { "color": "4472C4", "pattern": "solid" },
                    "border": {
                        "top": { "style": "thin", "color": "000000" },
                        "bottom": { "style": "medium" }
                    },
                    "alignment": { "horizontal": "center", "vertical": "center", "wrapText": true },
                    "numberFormat": "0.00"
                }
            ],
            "worksheets": [{"name": "S", "rows": [["1"]]}]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        var style = Assert.Single(set.Styles!);
        Assert.Equal("Header", style.Name);
        Assert.True(style.Font!.Bold);
        Assert.False(style.Font.Italic);
        Assert.Equal("FFFFFF", style.Font.Color);
        Assert.Equal(12d, style.Font.Size);
        Assert.Equal("solid", style.Fill!.Pattern);
        Assert.Equal("thin", style.Border!.Top!.Style);
        Assert.Equal("medium", style.Border.Bottom!.Style);
        Assert.Equal("center", style.Alignment!.Horizontal);
        Assert.True(style.Alignment.WrapText);
        Assert.Equal("0.00", style.NumberFormat);
    }

    [Fact]
    public void Parse_ShouldRejectDuplicateStyleNames()
    {
        var json = """
        {
            "version": "1.0",
            "styles": [
                {"name": "Money", "numberFormat": "0.00"},
                {"name": "money", "font": {"bold": true}}
            ],
            "worksheets": [{"name": "S", "rows": [["1"]]}]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("Duplicate style name", ex.Message);
    }

    [Fact]
    public void Parse_ShouldDeserializeWorksheetLayoutVocabulary()
    {
        var json = """
        {
            "version": "1.0",
            "styles": [
                {"name": "header", "font": {"bold": true}},
                {"name": "num", "numberFormat": "#,##0.00"}
            ],
            "worksheets": [
                {
                    "name": "Report",
                    "startRow": 3,
                    "headerStyle": "header",
                    "columns": [
                        {"name": "Region", "width": 18, "type": "string"},
                        {"name": "Amount", "width": 12, "style": "num", "type": "number"}
                    ],
                    "headers": ["Region", "Amount"],
                    "rows": [["North", "100.5"]],
                    "rowHeights": [{"row": 3, "height": 24}, {"row": 4, "height": 20}],
                    "merges": ["A1:C1"],
                    "freezePanes": {"row": 4, "column": 1},
                    "autoFilter": "A20:D30",
                    "tables": [{"name": "ReportTable", "range": "A3:B4"}],
                    "cells": [
                        {"address": "D5", "formula": "=SUM(B4:B5)", "type": "number", "numberFormat": "#,##0"}
                    ]
                }
            ]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        var ws = Assert.Single(set.Worksheets);
        Assert.Equal(3, ws.StartRow);
        Assert.Equal("header", ws.HeaderStyle);
        Assert.Equal(2, ws.Columns!.Count);
        Assert.Equal("Region", ws.Columns[0].Name);
        Assert.Equal(18d, ws.Columns[0].Width);
        Assert.Equal("string", ws.Columns[0].Type);
        Assert.Equal("num", ws.Columns[1].Style);
        Assert.Equal(2, ws.RowHeights!.Count);
        Assert.Equal(24d, ws.RowHeights[0].Height);
        Assert.Equal("A1:C1", ws.Merges![0]);
        Assert.Equal(4, ws.FreezePanes!.Row);
        Assert.Equal("A20:D30", ws.AutoFilter);
        Assert.Equal("ReportTable", ws.Tables![0].Name);
        Assert.Equal("A3:B4", ws.Tables[0].Range);
        Assert.Equal("number", ws.Cells![0].Type);
        Assert.Equal("#,##0", ws.Cells[0].NumberFormat);
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("string")]
    [InlineData("number")]
    [InlineData("boolean")]
    [InlineData("date")]
    [InlineData("datetime")]
    [InlineData("NUMBER")]
    public void Parse_ShouldAcceptEveryDiscreteCellType(string type)
    {
        var json = $$"""
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "x", "type": "{{type}}"}]
            }]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        Assert.Equal(type, set.Worksheets[0].Cells![0].Type);
    }

    [Fact]
    public void Parse_ShouldRejectUnknownCellType()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "x", "type": "currency"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("unknown 'type'", ex.Message);
        Assert.Contains("currency", ex.Message);
    }

    [Fact]
    public void Parse_ShouldAcceptExplicitEmptyStringValue()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "", "type": "string"}]
            }]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        var cell = set.Worksheets[0].Cells![0];
        Assert.Equal(string.Empty, cell.Value);
    }

    [Fact]
    public void Parse_ShouldRejectCellReferencingUndefinedNamedStyle()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "x", "style": "missing"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("unknown style 'missing'", ex.Message);
    }

    [Fact]
    public void Parse_ShouldAcceptLegacyNumericStyleReference()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "cells": [{"address": "A1", "value": "x", "style": "0"}]
            }]
        }
        """;

        var set = XlsxInstructionParser.Parse(json);
        Assert.Equal("0", set.Worksheets[0].Cells![0].Style);
    }

    [Fact]
    public void Parse_ShouldRejectInvalidFillPattern()
    {
        var json = """
        {
            "version": "1.0",
            "styles": [{"name": "S", "fill": {"color": "FF0000", "pattern": "stripes"}}],
            "worksheets": [{"name": "W", "rows": [["1"]]}]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("fill pattern", ex.Message);
    }

    [Fact]
    public void Parse_ShouldRejectMergeOverlap()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["a"],
                "merges": ["A1:C3", "B2:D4"]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("overlaps", ex.Message);
    }

    [Fact]
    public void Parse_ShouldRejectTableAndMergeOverlap()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["a", "b", "c"],
                "rows": [["1", "2", "3"]],
                "tables": [{"name": "SalesTable", "range": "A1:C2"}],
                "merges": ["B1:D1"]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("overlap", ex.Message);
    }

    [Fact]
    public void Parse_ShouldRejectOverlappingTables()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["a", "b", "c", "d"],
                "rows": [["1", "2", "3", "4"]],
                "tables": [
                    {"name": "TableOne", "range": "A1:C2"},
                    {"name": "TableTwo", "range": "B2:D3"}
                ]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("overlaps table", ex.Message);
    }

    [Fact]
    public void Parse_ShouldRejectTableNameThatLooksLikeACellReference()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "headers": ["a"],
                "rows": [["1"]],
                "tables": [{"name": "A1", "range": "A1:B2"}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("cell reference", ex.Message);
    }

    [Fact]
    public void Parse_ShouldRejectColumnWidthBeyondExcelLimit()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "columns": [{"width": 300}],
                "rows": [["1"]]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("255", ex.Message);
    }

    [Fact]
    public void Parse_ShouldRejectRowHeightBeyondExcelLimit()
    {
        var json = """
        {
            "version": "1.0",
            "worksheets": [{
                "name": "S",
                "rows": [["1"]],
                "rowHeights": [{"row": 1, "height": 500}]
            }]
        }
        """;

        var ex = Assert.Throws<XlsxException>(() => XlsxInstructionParser.Parse(json));
        Assert.Contains("409.5", ex.Message);
    }
}
