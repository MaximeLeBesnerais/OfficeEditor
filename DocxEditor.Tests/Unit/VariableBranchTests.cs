using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Variables;
using XlsxEditor.Core.Variables;
using W = DocumentFormat.OpenXml.Wordprocessing;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace DocxEditor.Tests.Unit;

public class VariableBranchTests : IDisposable
{
    private readonly List<string> _files = new();

    [Fact]
    public void DocxDetector_ShouldScanBodyHeadersFootersAndDeduplicatePerLocation()
    {
        var path = CreateDocx(
            new W.Paragraph(new W.Run(new W.Text("Hello {{ name |Guest}} and {{name|Again}} {{empty|}} {{missing"))),
            headerText: "Header {{title}} {{title}}",
            footerText: "Footer {{name|FooterDefault}}");

        using var document = WordprocessingDocument.Open(path, false);
        var variables = new DocxVariableDetector().Scan(document);

        Assert.Equal(4, variables.Count);
        Assert.Contains(variables, v => v.Name == "name" && v.Location == "body" && v.DefaultValue == "Guest");
        Assert.Contains(variables, v => v.Name == "empty" && v.Location == "body" && v.DefaultValue == "");
        Assert.Contains(variables, v => v.Name == "title" && v.Location.StartsWith("header:"));
        Assert.Contains(variables, v => v.Name == "name" && v.Location.StartsWith("footer:") && v.DefaultValue == "FooterDefault");
        Assert.DoesNotContain(variables, v => v.FullMatch == "{{missing");
    }

    [Fact]
    public void DocxDetector_WithEmptyDocumentPart_ShouldReturnEmptyList()
    {
        var path = NewTempFile("docx");
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
        }

        using var reopened = WordprocessingDocument.Open(path, false);

        Assert.Empty(new DocxVariableDetector().Scan(reopened));
    }

    [Fact]
    public void DocxReplacer_ShouldReplaceSingleRunSplitRunsDefaultsAndPreserveMissingVariables()
    {
        var split = new W.Paragraph(
            new W.Run(new W.Text("Split {{fi")),
            new W.Run(new W.Text("rst}} and {{second|Default}} and {{empty|}} and {{missing}}")));
        var path = CreateDocx(new[]
            {
            new W.Paragraph(new W.Run(new W.Text("Hi {{ name }}; repeated {{name}}"))),
            split
            },
            headerText: "Header {{header|Fallback}}",
            footerText: "Footer {{footer}}");

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["name"] = "Ada",
                ["first"] = "One",
                ["footer"] = "Done"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var allText = reopened.MainDocumentPart!.Document!.InnerText
            + string.Concat(reopened.MainDocumentPart.HeaderParts.Select(h => h.Header!.InnerText))
            + string.Concat(reopened.MainDocumentPart.FooterParts.Select(f => f.Footer!.InnerText));

        Assert.Contains("Hi Ada; repeated Ada", allText);
        Assert.Contains("Split One and Default and  and {{missing}}", allText);
        Assert.Contains("Header Fallback", allText);
        Assert.Contains("Footer Done", allText);

        var textNodes = reopened.MainDocumentPart.Document!.Body!.Descendants<W.Text>().ToList();
        Assert.Contains(textNodes, t => t.Text.Contains("Split One") && t.Space?.Value == SpaceProcessingModeValues.Preserve);
    }

    [Fact]
    public void DocxReplacer_WithNoVariablesOrSingleTextRun_ShouldLeaveTextUnchanged()
    {
        var path = CreateDocx(new W.Paragraph(new W.Run(new W.Text("Plain text"))));

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>());
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("Plain text", reopened.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DocxReplacer_WithMultiRunMixedFormatting_ShouldPreserveRunProperties()
    {
        var para = new W.Paragraph(
            new W.Run(
                new W.RunProperties(new W.Bold()),
                new W.Text("Bold {{user")),
            new W.Run(
                new W.RunProperties(new W.Italic()),
                new W.Text("name}} suffix")),
            new W.Run(
                new W.RunProperties(new W.Bold(), new W.Italic()),
                new W.Text(" tail")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["username"] = "Ada"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("Bold Ada suffix tail", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().First().Elements<W.Run>().ToList();
        Assert.Equal(3, runs.Count);

        var rp0 = runs[0].GetFirstChild<W.RunProperties>();
        var rp1 = runs[1].GetFirstChild<W.RunProperties>();
        var rp2 = runs[2].GetFirstChild<W.RunProperties>();

        Assert.NotNull(rp0);
        Assert.NotEmpty(rp0!.Elements<W.Bold>());
        Assert.Empty(rp0.Elements<W.Italic>());

        Assert.NotNull(rp1);
        Assert.Empty(rp1!.Elements<W.Bold>());
        Assert.NotEmpty(rp1.Elements<W.Italic>());

        Assert.NotNull(rp2);
        Assert.NotEmpty(rp2!.Elements<W.Bold>());
        Assert.NotEmpty(rp2.Elements<W.Italic>());
    }

    [Fact]
    public void DocxReplacer_WithMultiRunVariableSplitAcrossManyRuns_ShouldPreserveAll()
    {
        var para = new W.Paragraph(
            new W.Run(new W.Text("Prefix {{my")),
            new W.Run(new W.Text("_var}}")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["my_var"] = "replaced"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        var innerText = body.InnerText;
        Assert.Equal("Prefix replaced", innerText);
    }

    [Fact]
    public void DocxReplacer_WithVariableSpanningMoreThanTwoRuns_ShouldReplaceAcrossAllRuns()
    {
        var para = new W.Paragraph(
            new W.Run(new W.Text("Start {{va")),
            new W.Run(new W.Text("ri")),
            new W.Run(new W.Text("able}} end")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["variable"] = "VALUE"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("Start VALUE end", reopened.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DocxReplacer_WithAdjacentVariablesSplitAcrossRuns_ShouldReplaceEach()
    {
        var para = new W.Paragraph(
            new W.Run(new W.Text("{{fi")),
            new W.Run(new W.Text("rst}}{{sec")),
            new W.Run(new W.Text("ond}}")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["first"] = "1",
                ["second"] = "2"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("12", reopened.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DocxReplacer_WithVariableInsideHyperlinkRun_ShouldReplace()
    {
        var para = new W.Paragraph(
            new W.Hyperlink(new W.Run(new W.Text("Visit {{site}}"))));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["site"] = "example.com"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("Visit example.com", reopened.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DocxReplacer_SingleRunReplacement_ShouldPreserveWhitespace()
    {
        var path = CreateDocx(new W.Paragraph(new W.Run(new W.Text("Hello {{name}}!"))));

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["name"] = " Ada "
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var text = reopened.MainDocumentPart!.Document!.Body!.Descendants<W.Text>().Single();
        Assert.Equal("Hello  Ada !", text.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, text.Space?.Value);
    }

    [Fact]
    public void DocxReplacer_WithNullDataValue_ShouldTreatAsMissingInBothPaths()
    {
        // Null values reach the dictionary through JSON deserialization of merge data;
        // semantics: null == missing, so defaults apply and bare placeholders are preserved.
        var data = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(
            """{"single": null, "split": null}""");
        Assert.NotNull(data);

        var para = new W.Paragraph(
            new W.Run(new W.Text("A {{single|Fallback}} B {{single}} C {{sp")),
            new W.Run(new W.Text("lit|SplitDefault}} D {{split}}")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, data);
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("A Fallback B {{single}} C SplitDefault D {{split}}",
            reopened.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void DocxReplacer_WithSplitPlaceholderAcrossMixedFormatRuns_ShouldPreserveExactRunTextAndProperties()
    {
        var para = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("Bold {{fi")),
            new W.Run(new W.RunProperties(new W.Italic()), new W.Text("rst}} and {{la")),
            new W.Run(new W.RunProperties(new W.Bold(), new W.Italic()), new W.Text("st}} done")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["first"] = "1st",
                ["last"] = "2nd"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("Bold 1st and 2nd done", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().Single().Elements<W.Run>().ToList();
        Assert.Equal(3, runs.Count);

        AssertRun(runs[0], "Bold 1st", bold: true, italic: false);
        AssertRun(runs[1], " and 2nd", bold: false, italic: true);
        AssertRun(runs[2], " done", bold: true, italic: true);

        Assert.Equal(SpaceProcessingModeValues.Preserve, runs[1].GetFirstChild<W.Text>()!.Space?.Value);
        Assert.Equal(SpaceProcessingModeValues.Preserve, runs[2].GetFirstChild<W.Text>()!.Space?.Value);
    }

    [Fact]
    public void DocxReplacer_WithPrefixAndSuffixInDifferentlyFormattedRuns_ShouldKeepEachRunFormatting()
    {
        var para = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.Italic()), new W.Text("Italic {{na")),
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("me}} bold")),
            new W.Run(new W.Text(" plain suffix")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["name"] = "Ada"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("Italic Ada bold plain suffix", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().Single().Elements<W.Run>().ToList();
        Assert.Equal(3, runs.Count);

        AssertRun(runs[0], "Italic Ada", bold: false, italic: true);
        AssertRun(runs[1], " bold", bold: true, italic: false);
        AssertRun(runs[2], " plain suffix", bold: false, italic: false);
        Assert.Equal(SpaceProcessingModeValues.Preserve, runs[1].GetFirstChild<W.Text>()!.Space?.Value);
        Assert.Null(runs[2].GetFirstChild<W.Text>()!.Space?.Value);
    }

    [Fact]
    public void DocxReplacer_WithMultiplePlaceholdersSpanningRuns_ShouldReplaceAllInOrder()
    {
        var para = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("{{a}} {{b")),
            new W.Run(new W.RunProperties(new W.Italic()), new W.Text("}} and {{c")),
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("}} end")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["a"] = "X",
                ["b"] = "Y",
                ["c"] = "Z"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("X Y and Z end", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().Single().Elements<W.Run>().ToList();
        Assert.Equal(3, runs.Count);

        AssertRun(runs[0], "X Y", bold: true, italic: false);
        AssertRun(runs[1], " and Z", bold: false, italic: true);
        AssertRun(runs[2], " end", bold: true, italic: false);

        Assert.Equal(SpaceProcessingModeValues.Preserve, runs[1].GetFirstChild<W.Text>()!.Space?.Value);
        Assert.Equal(SpaceProcessingModeValues.Preserve, runs[2].GetFirstChild<W.Text>()!.Space?.Value);
    }

    [Fact]
    public void DocxReplacer_WithMissingVariableSpanningRuns_ShouldPreservePlaceholderAndFormatting()
    {
        var para = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("Bold {{miss")),
            new W.Run(new W.RunProperties(new W.Italic()), new W.Text("ing}} tail")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>());
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("Bold {{missing}} tail", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().Single().Elements<W.Run>().ToList();
        Assert.Equal(2, runs.Count);

        AssertRun(runs[0], "Bold {{miss", bold: true, italic: false);
        AssertRun(runs[1], "ing}} tail", bold: false, italic: true);
    }

    [Fact]
    public void DocxReplacer_WithSingleRunPlaceholderAndOtherFormattedRuns_ShouldNotDisturbNeighbors()
    {
        var para = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("Hello {{name}}")),
            new W.Run(new W.RunProperties(new W.Italic()), new W.Text(" world")),
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("!")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["name"] = "Ada"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("Hello Ada world!", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().Single().Elements<W.Run>().ToList();
        Assert.Equal(3, runs.Count);

        AssertRun(runs[0], "Hello Ada", bold: true, italic: false);
        AssertRun(runs[1], " world", bold: false, italic: true);
        AssertRun(runs[2], "!", bold: true, italic: false);
    }

    [Fact]
    public void DocxReplacer_WithReplacementContainingPlaceholderPattern_ShouldNotResubstitute()
    {
        var para = new W.Paragraph(
            new W.Run(new W.Text("Hello {{a}}")),
            new W.Run(new W.Text(" world")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["a"] = "X{{y}}Z",
                ["y"] = "evil"
            });
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Equal("Hello X{{y}}Z world", body.InnerText);
        Assert.DoesNotContain("evil", body.InnerText);

        var runs = body.Descendants<W.Paragraph>().Single().Elements<W.Run>().ToList();
        Assert.Equal(2, runs.Count);
        AssertRun(runs[0], "Hello X{{y}}Z", bold: false, italic: false);
        AssertRun(runs[1], " world", bold: false, italic: false);
    }

    [Fact]
    public void DocxReplacer_SaveAndReopen_ShouldProduceValidatorCleanDocument()
    {
        var para = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("Bold {{fi")),
            new W.Run(new W.RunProperties(new W.Italic()), new W.Text("rst}} and {{la")),
            new W.Run(new W.RunProperties(new W.Bold(), new W.Italic()), new W.Text("st}} done")),
            new W.Run(new W.Text(" {{miss")),
            new W.Run(new W.Text("ing}} tail")));
        var path = CreateDocx(new[] { para });

        using (var document = WordprocessingDocument.Open(path, true))
        {
            new DocxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["first"] = "1st",
                ["last"] = "2nd"
            });
        }

        OpenXmlAssert.NoDocxValidationErrors(path);

        using var reopened = WordprocessingDocument.Open(path, false);
        Assert.Equal("Bold 1st and 2nd done {{missing}} tail",
            reopened.MainDocumentPart!.Document!.Body!.InnerText);
    }

    [Fact]
    public void XlsxDetector_ShouldScanSharedAndPlainCellsWithDefaultsAndDuplicateLocations()
    {
        var path = CreateXlsx(includeSharedStringPart: true,
            ("A1", CellValues.SharedString, "0"),
            ("B1", CellValues.String, "{{name}} {{name|Other}}"),
            ("C1", null, "{{empty|}} {{broken"),
            (null, CellValues.String, "{{noRef|Default}}"));

        using var document = SpreadsheetDocument.Open(path, false);
        var variables = new XlsxVariableDetector().Scan(document);

        Assert.Equal(4, variables.Count);
        Assert.Contains(variables, v => v.Name == "shared" && v.DefaultValue == "SharedDefault" && v.Location == "sheet:Sheet1:cell:A1");
        Assert.Single(variables, v => v.Name == "name" && v.Location == "sheet:Sheet1:cell:B1");
        Assert.Contains(variables, v => v.Name == "empty" && v.DefaultValue == "" && v.Location == "sheet:Sheet1:cell:C1");
        Assert.Contains(variables, v => v.Name == "noRef" && v.Location == "sheet:Sheet1:cell:Unknown");
    }

    [Fact]
    public void XlsxDetector_WithMissingWorkbookOrSheets_ShouldReturnEmptyList()
    {
        var noWorkbookPath = NewTempFile("xlsx");
        using (SpreadsheetDocument.Create(noWorkbookPath, SpreadsheetDocumentType.Workbook))
        {
        }

        using (var document = SpreadsheetDocument.Open(noWorkbookPath, false))
        {
            Assert.Empty(new XlsxVariableDetector().Scan(document));
        }

        var noSheetsPath = NewTempFile("xlsx");
        using (var document = SpreadsheetDocument.Create(noSheetsPath, SpreadsheetDocumentType.Workbook))
        {
            document.AddWorkbookPart().Workbook = new Workbook();
        }

        using (var document = SpreadsheetDocument.Open(noSheetsPath, false))
        {
            Assert.Empty(new XlsxVariableDetector().Scan(document));
        }
    }

    [Fact]
    public void XlsxReplacer_ShouldReplaceSharedStringsPlainCellsDefaultsAndMissingValues()
    {
        var path = CreateXlsx(includeSharedStringPart: true,
            ("A1", CellValues.SharedString, "0"),
            ("B1", CellValues.String, "{{provided}} {{defaulted|Fallback}} {{empty|}} {{missing}}"),
            ("C1", null, "plain"));

        using (var document = SpreadsheetDocument.Open(path, true))
        {
            new XlsxVariableReplacer().Replace(document, new Dictionary<string, string>
            {
                ["shared"] = "SharedValue",
                ["provided"] = "Given"
            });
        }

        using var reopened = SpreadsheetDocument.Open(path, false);
        var workbookPart = reopened.WorkbookPart!;
        var cells = workbookPart.WorksheetParts.First().Worksheet!.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);

        Assert.Equal(CellValues.SharedString, cells["A1"].DataType!.Value);
        Assert.Equal("SharedValue", ReadCell(cells["A1"], workbookPart));
        Assert.Equal(CellValues.String, cells["B1"].DataType!.Value);
        Assert.Equal("Given Fallback {{empty|}} {{missing}}", cells["B1"].CellValue!.Text);
        Assert.Equal("plain", cells["C1"].CellValue!.Text);
    }

    [Fact]
    public void XlsxReplacer_WithExistingSharedStringResult_ShouldReuseExistingIndex()
    {
        var path = CreateXlsx(includeSharedStringPart: true,
            ("A1", CellValues.SharedString, "0"),
            ("B1", CellValues.SharedString, "1"));

        using (var document = SpreadsheetDocument.Open(path, true))
        {
            new XlsxVariableReplacer().Replace(document, new Dictionary<string, string> { ["shared"] = "Existing" });
        }

        using var reopened = SpreadsheetDocument.Open(path, false);
        var a1 = reopened.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().First(c => c.CellReference == "A1");
        Assert.Equal("1", a1.CellValue!.Text);
    }

    [Fact]
    public void XlsxReplacer_NullDocument_ThrowsArgumentNullException()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => new XlsxVariableReplacer().Replace(null!, new Dictionary<string, string>()));
        Assert.Equal("document", ex.ParamName);
    }

    [Fact]
    public void XlsxReplacer_NullData_ThrowsArgumentNullException()
    {
        using var document = SpreadsheetDocument.Create(NewTempFile("xlsx"), SpreadsheetDocumentType.Workbook);
        var ex = Assert.Throws<ArgumentNullException>(
            () => new XlsxVariableReplacer().Replace(document, null!));
        Assert.Equal("data", ex.ParamName);
    }

    [Fact]
    public void XlsxReplacer_NullValue_ThrowsArgumentNullException_AndLeavesCellUnchanged()
    {
        var path = CreateXlsx(false, ("A1", CellValues.String, "Hello {{name}}"));

        using (var document = SpreadsheetDocument.Open(path, true))
        {
            var ex = Assert.Throws<ArgumentNullException>(() =>
                new XlsxVariableReplacer().Replace(document, new Dictionary<string, string> { ["name"] = null! }));
            Assert.Contains("name", ex.Message);
        }

        // Atomic: the rejected merge left the placeholder untouched.
        using var reopened = SpreadsheetDocument.Open(path, false);
        var cell = reopened.WorkbookPart!.WorksheetParts.First().Worksheet!
            .GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().First();
        Assert.Equal("Hello {{name}}", cell.CellValue!.Text);
    }

    public void Dispose()
    {
        foreach (var file in _files)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }

    private string CreateDocx(W.Paragraph bodyParagraph, string? headerText = null, string? footerText = null)
        => CreateDocx(new[] { bodyParagraph }, headerText, footerText);

    private string CreateDocx(IEnumerable<W.Paragraph> bodyParagraphs, string? headerText = null, string? footerText = null)
    {
        var path = NewTempFile("docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body(bodyParagraphs.Cast<OpenXmlElement>()));

        if (headerText != null)
        {
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text(headerText))));
            headerPart.Header.Save();
        }

        if (footerText != null)
        {
            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text(footerText))));
            footerPart.Footer.Save();
        }

        mainPart.Document.Save();
        return path;
    }

    private string CreateXlsx(bool includeSharedStringPart, params (string? Reference, CellValues? DataType, string Value)[] cells)
    {
        var path = NewTempFile("xlsx");
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());

        if (includeSharedStringPart)
        {
            var shared = workbookPart.AddNewPart<SharedStringTablePart>();
            shared.SharedStringTable = new SharedStringTable(
                new SharedStringItem(new S.Text("{{shared|SharedDefault}}")),
                new SharedStringItem(new S.Text("Existing")));
            shared.SharedStringTable.Save();
        }

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var row = new Row { RowIndex = 1 };
        foreach (var spec in cells)
        {
            var cell = new Cell { CellValue = new CellValue(spec.Value) };
            if (spec.Reference != null)
            {
                cell.CellReference = spec.Reference;
            }
            if (spec.DataType.HasValue)
            {
                cell.DataType = spec.DataType.Value;
            }
            row.Append(cell);
        }

        worksheetPart.Worksheet = new Worksheet(new SheetData(row));
        worksheetPart.Worksheet.Save();
        var relationshipId = workbookPart.GetIdOfPart(worksheetPart);
        workbookPart.Workbook.Sheets!.Append(new Sheet { Id = relationshipId, SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        return path;
    }

    private string NewTempFile(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"branch_vars_{Guid.NewGuid():N}.{extension}");
        _files.Add(path);
        return path;
    }

    private static string ReadCell(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var index = int.Parse(cell.CellValue!.Text);
            return workbookPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(index).InnerText;
        }

        return cell.CellValue?.Text ?? string.Empty;
    }

    private static void AssertRun(W.Run run, string expectedText, bool bold, bool italic)
    {
        Assert.Equal(expectedText, run.Elements<W.Text>().SingleOrDefault()?.Text ?? string.Empty);
        var rp = run.GetFirstChild<W.RunProperties>();
        Assert.Equal(bold, rp?.Elements<W.Bold>().Any() ?? false);
        Assert.Equal(italic, rp?.Elements<W.Italic>().Any() ?? false);
    }
}
