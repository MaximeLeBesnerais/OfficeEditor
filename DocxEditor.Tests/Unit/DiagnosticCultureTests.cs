using System.Globalization;
using DocxEditor.Core.Generation.Emit.Ooxml.Positioned;
using DocxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Fixtures;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;
using XlsxEditor.Core.Instructions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Diagnostics are read by humans and parsed by tools (CLI output, API responses, MCP
/// payloads, RMSE gates), so their numbers must not follow the ambient locale. Before
/// this was enforced, the same messages rendered as "1 048 576" / "4,1%" on a French
/// machine and the suite failed in six places while CI — running in an English
/// environment — stayed green.
///
/// The culture used here is a clone of the invariant culture with a space group
/// separator and a comma decimal separator, so the test is deterministic and does not
/// depend on ICU locale data being installed.
/// </summary>
public sealed class DiagnosticCultureTests
{
    [Fact]
    public void XlsxPlannerDiagnostics_ShouldFormatNumbersInvariantly()
    {
        UnderNonEnglishCulture(() =>
        {
            var tooLongText = new XlsxInstructionSet
            {
                Version = "1.0",
                Worksheets =
                [
                    new WorksheetInstruction
                    {
                        Name = "S",
                        Cells = [new CellInstruction { Address = "A1", Value = new string('x', 32_768) }]
                    }
                ]
            };

            var textResult = XlsxPlanner.Plan(tooLongText);
            var textError = Assert.Single(textResult.Validation.Errors, d => d.Code == XlsxDiagnosticCode.TextTooLong);
            Assert.Contains("32,767", textError.Message);
            Assert.DoesNotContain("32\u00A0767", textError.Message);

            var tooLongFormula = new XlsxInstructionSet
            {
                Version = "1.0",
                Worksheets =
                [
                    new WorksheetInstruction
                    {
                        Name = "S",
                        Cells = [new CellInstruction { Address = "A1", Formula = "=" + new string('f', 9_000) }]
                    }
                ]
            };

            var formulaResult = XlsxPlanner.Plan(tooLongFormula);
            var formulaError = Assert.Single(
                formulaResult.Validation.Errors,
                d => d.Code == XlsxDiagnosticCode.FormulaTooLong);
            Assert.Contains("8,192", formulaError.Message);
            Assert.DoesNotContain("8\u00A0192", formulaError.Message);
        });
    }

    [Fact]
    public void XlsxBuilderExceptionMessages_ShouldFormatNumbersInvariantly()
    {
        UnderNonEnglishCulture(() =>
        {
            var ex = Assert.Throws<XlsxException>(
                () => WorksheetBuilder.GetCellReference(0, 1_048_577));
            Assert.Contains("1,048,576", ex.Message);
            Assert.DoesNotContain("1\u00A0048\u00A0576", ex.Message);
        });
    }

    [Fact]
    public void ParityFailureText_ShouldFormatPercentagesInvariantly()
    {
        UnderNonEnglishCulture(() =>
        {
            var thresholds = new Dictionary<string, double>(StringComparer.Ordinal) { ["rect-radii"] = 0.04 };
            // Two pages, so only the breaching page fires (a single page would also breach
            // the per-document average and yield two failure lines).
            var document = new ParityDocumentSample(
                "rect-radii",
                false,
                [new ParityPageSample(1, 0.01), new ParityPageSample(2, 0.041)]);

            var failures = ParityEvaluator.Evaluate([document], thresholds);

            var failure = Assert.Single(failures);
            Assert.Contains("4.1%", failure);
            Assert.DoesNotContain("4,1%", failure);
        });
    }

    [Fact]
    public void ParityThresholdException_ShouldFormatNumbersInvariantly()
    {
        UnderNonEnglishCulture(() =>
        {
            var thresholds = new Dictionary<string, double>(StringComparer.Ordinal) { ["rect-radii"] = 1.5 };
            var document = new ParityDocumentSample("rect-radii", false, [new ParityPageSample(1, 0.01)]);

            var ex = Assert.Throws<ArgumentException>(() => ParityEvaluator.Evaluate([document], thresholds));
            Assert.Contains("1.5", ex.Message);
            Assert.DoesNotContain("1,5", ex.Message);
        });
    }

    [Fact]
    public void DocxPositionedEmitterMessages_ShouldFormatMeasurementsInvariantly()
    {
        UnderNonEnglishCulture(() =>
        {
            var container = new DocumentFormat.OpenXml.Wordprocessing.Body();
            var line = new PositionedLine
            {
                Position = new PositionSpec { X = 0, Y = 0, WidthPt = -0.5 }
            };

            var ex = Assert.Throws<PositionedElementEmitException>(
                () => new PositionedElementEmitter().Emit(null, container, line));

            Assert.Contains("-0.5", ex.Message);
            Assert.DoesNotContain("-0,5", ex.Message);
        });
    }

    /// <summary>
    /// Runs <paramref name="action"/> with a deliberately non-English <see cref="CultureInfo.CurrentCulture"/>
    /// and restores the previous culture afterwards. The two sanity assertions make the test
    /// self-checking: if the fake culture ever stopped differing from the invariant one, the
    /// assertions below would fail loudly instead of passing vacuously.
    /// </summary>
    private static void UnderNonEnglishCulture(Action action)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        var fake = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        fake.NumberFormat.NumberGroupSeparator = "\u00A0";
        fake.NumberFormat.NumberDecimalSeparator = ",";
        fake.NumberFormat.PercentDecimalSeparator = ",";

        try
        {
            CultureInfo.CurrentCulture = fake;

            Assert.Equal("32\u00A0767", 32_767.ToString("N0", CultureInfo.CurrentCulture));
            Assert.Equal("4,1", (0.041 * 100).ToString("0.##", CultureInfo.CurrentCulture));

            action();
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}
