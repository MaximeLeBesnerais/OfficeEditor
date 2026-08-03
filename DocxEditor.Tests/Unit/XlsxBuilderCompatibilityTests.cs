using OfficeEditor.Core.Models;
using XlsxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Proves the fluent builder interfaces stay source-compatible with external
/// implementations. The members added by the rich builder/planner work
/// (typed cells, styles, freeze panes, autofilter, core/calculation properties) are
/// default-implemented on the interface, so a class written against the original surface
/// still compiles without them: read-only queries return inert defaults and mutating
/// operations fail loudly with <see cref="NotSupportedException"/> instead of breaking
/// compilation.
/// </summary>
public class XlsxBuilderCompatibilityTests
{
    [Fact]
    public void MinimalWorksheetImplementation_ShouldUseDefaultInterfaceMembers()
    {
        IWorksheetBuilder sheet = new MinimalWorksheetBuilder();

        // Read-only queries added later fall back to inert defaults.
        Assert.Null(sheet.GetFreezePanes());
        Assert.Null(sheet.GetAutoFilterRange());

        // Mutating operations added later fail loudly rather than silently no-op.
        Assert.Throws<NotSupportedException>(() => sheet.AddCellString("A1", "x"));
        Assert.Throws<NotSupportedException>(() => sheet.AddCellNumber("A1", 1));
        Assert.Throws<NotSupportedException>(() => sheet.AddCellBoolean("A1", true));
        Assert.Throws<NotSupportedException>(() => sheet.AddCellDate("A1", "2026-01-01"));
        Assert.Throws<NotSupportedException>(() => sheet.AddCellDateTime("A1", "2026-01-01T00:00:00"));
        Assert.Throws<NotSupportedException>(() => sheet.AddFormula("A1", "=1+1"));
        Assert.Throws<NotSupportedException>(() => sheet.FreezePanes(1, 0));
        Assert.Throws<NotSupportedException>(() => sheet.SetAutoFilter("A1:B2"));
        Assert.Throws<NotSupportedException>(() => sheet.RemoveAutoFilter());

        // The original surface still works.
        Assert.Same(sheet, sheet.AddCell("A1", "x"));
        Assert.Empty(sheet.GetMergeRanges());
    }

    [Fact]
    public void MinimalWorkbookImplementation_ShouldUseDefaultInterfaceMembers()
    {
        IWorkbookBuilder book = new MinimalWorkbookBuilder();

        Assert.Empty(book.GetDefinedStyleNames());
        Assert.NotNull(book.GetCoreProperties());
        Assert.NotNull(book.GetCalculationProperties());

        Assert.Throws<NotSupportedException>(() => book.DefineStyle(new CellStyleSpec { Name = "s" }));
        Assert.Throws<NotSupportedException>(() => book.GetStyleIndex("s"));
        Assert.Throws<NotSupportedException>(() => book.SetCoreProperties(new WorkbookCoreProperties()));
        Assert.Throws<NotSupportedException>(() => book.SetCalculationProperties(new WorkbookCalculationProperties()));

        Assert.IsType<MinimalWorksheetBuilder>(book.AddWorksheet("S"));
        Assert.Same(book, book.RemoveWorksheet("S"));
    }

    [Fact]
    public void ConcreteBuilders_ShouldStillOverrideNewMembers()
    {
        // WorkbookBuilder/WorksheetBuilder implement every member explicitly, so the
        // defaults are never invoked for the shipped implementations.
        using var builder = WorkbookBuilder.Create();
        IWorkbookBuilder asInterface = builder;
        Assert.Empty(asInterface.GetDefinedStyleNames());
        asInterface.DefineStyle(new CellStyleSpec { Name = "money", NumberFormat = "0.00" });
        Assert.Equal(new List<string> { "money" }, asInterface.GetDefinedStyleNames());
    }

    /// <summary>
    /// A minimal external <see cref="IWorkbookBuilder"/> written against the original
    /// surface only — deliberately does NOT implement the members added later, proving
    /// the default interface members keep it source-compatible.
    /// </summary>
    private sealed class MinimalWorkbookBuilder : IWorkbookBuilder
    {
        public static IWorkbookBuilder Create() => new MinimalWorkbookBuilder();

        public static IWorkbookBuilder Open(Stream stream) => new MinimalWorkbookBuilder();

        public static IWorkbookBuilder Open(byte[] bytes) => new MinimalWorkbookBuilder();

        public IWorksheetBuilder AddWorksheet(string name) => new MinimalWorksheetBuilder();

        public IWorkbookBuilder RemoveWorksheet(string name) => this;

        public IWorksheetBuilder GetWorksheet(string name) => new MinimalWorksheetBuilder();

        public List<string> GetWorksheetNames() => new();

        public List<VariableInfo> DetectVariables() => new();

        public IWorkbookBuilder MergeVariables(Dictionary<string, string> data) => this;

        public void Save(string? path = null)
        {
        }

        public void Save(Stream stream)
        {
        }

        public byte[] SaveToBytes() => Array.Empty<byte>();

        public void Dispose()
        {
        }
    }

    /// <summary>
    /// A minimal external <see cref="IWorksheetBuilder"/> implementing only the original
    /// members; the typed-cell/layout members added later are inherited from the
    /// interface defaults.
    /// </summary>
    private sealed class MinimalWorksheetBuilder : IWorksheetBuilder
    {
        public IWorksheetBuilder AddCell(string cellReference, string value) => this;

        public IWorksheetBuilder AddCell(string cellReference, string formula, bool isFormula) => this;

        public IWorksheetBuilder AddCell(string cellReference, string value, string styleId) => this;

        public IWorksheetBuilder AddHeaderRow(List<string> values, int rowIndex = 1) => this;

        public IWorksheetBuilder AddDataRow(List<string> values, int rowIndex) => this;

        public IWorksheetBuilder AddFormulaRow(List<string> formulas, int rowIndex) => this;

        public IWorksheetBuilder AddTable(string startCell, string endCell, string tableName) => this;

        public IWorksheetBuilder AddChart(ChartType type, string dataRange) => this;

        public IWorksheetBuilder SetColumnWidth(string column, double width) => this;

        public IWorksheetBuilder SetRowHeight(int rowIndex, double height) => this;

        public IWorksheetBuilder MergeCells(string range) => this;

        public IWorksheetBuilder UnmergeCells(string range) => this;

        public string? GetCellValue(string cellReference) => null;

        public string? GetCellFormula(string cellReference) => null;

        public bool CellExists(string cellReference) => false;

        public CellInfo? GetCellInfo(string cellReference) => null;

        public List<CellInfo> GetRange(string start, string end) => new();

        public List<RowInfo> GetRows() => new();

        public RowInfo? GetRow(int rowIndex) => null;

        public (int firstRow, int lastRow, int firstCol, int lastCol) GetDimensions() => (0, 0, 0, 0);

        public List<string> GetMergeRanges() => new();

        public double? GetColumnWidth(string column) => null;

        public double? GetRowHeight(int rowIndex) => null;

        public IWorksheetBuilder DeleteCell(string cellReference) => this;

        public IWorksheetBuilder DeleteRow(int rowIndex) => this;

        public IWorksheetBuilder ClearRange(string start, string end) => this;
    }
}
