using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeEditor.Core.Models;
using XlsxEditor.Core.Exceptions;

namespace XlsxEditor.Core.Builders;

public interface IWorkbookBuilder : IDisposable
{
    IWorksheetBuilder AddWorksheet(string name);
    IWorkbookBuilder RemoveWorksheet(string name);
    IWorksheetBuilder GetWorksheet(string name);
    List<string> GetWorksheetNames();
    
    // Variables
    List<VariableInfo> DetectVariables();
    IWorkbookBuilder MergeVariables(Dictionary<string, string> data);

    void Save(string? path = null);
    void Save(Stream stream);
    byte[] SaveToBytes();

    static abstract IWorkbookBuilder Create();
    static abstract IWorkbookBuilder Open(Stream stream);
    static abstract IWorkbookBuilder Open(byte[] bytes);
}

public interface IWorksheetBuilder
{
    // Write
    IWorksheetBuilder AddCell(string cellReference, string value);
    IWorksheetBuilder AddCell(string cellReference, string formula, bool isFormula);
    IWorksheetBuilder AddCell(string cellReference, string value, string styleId);
    IWorksheetBuilder AddHeaderRow(List<string> values, int rowIndex = 1);
    IWorksheetBuilder AddDataRow(List<string> values, int rowIndex);
    IWorksheetBuilder AddFormulaRow(List<string> formulas, int rowIndex);
    IWorksheetBuilder AddTable(string startCell, string endCell, string tableName);
    IWorksheetBuilder AddChart(ChartType type, string dataRange);

    // Read
    string? GetCellValue(string cellReference);
    string? GetCellFormula(string cellReference);
    bool CellExists(string cellReference);
    CellInfo? GetCellInfo(string cellReference);
    List<CellInfo> GetRange(string start, string end);
    List<RowInfo> GetRows();
    RowInfo? GetRow(int rowIndex);
    (int firstRow, int lastRow, int firstCol, int lastCol) GetDimensions();

    // Edit
    IWorksheetBuilder DeleteCell(string cellReference);
    IWorksheetBuilder DeleteRow(int rowIndex);
    IWorksheetBuilder ClearRange(string start, string end);
}

public sealed record CellInfo
{
    public string Reference { get; init; } = string.Empty;
    public string? Value { get; init; }
    public string? Formula { get; init; }
    public CellValues? DataType { get; init; }
}

public sealed record RowInfo
{
    public int RowIndex { get; init; }
    public List<CellInfo> Cells { get; init; } = new();
}

public enum ChartType
{
    Bar,
    Line,
    Pie,
    Column
}

public class WorkbookBuilder : IWorkbookBuilder
{
    private readonly SpreadsheetDocument _document;
    private readonly string? _path;
    private readonly MemoryStream? _documentStream;
    private readonly bool _isNewDocument;
    // Excel worksheet names are case-insensitive; an Ordinal dictionary would allow
    // "Sales" and "SALES" to coexist and produce a corrupt workbook.
    private readonly Dictionary<string, WorksheetBuilder> _worksheets = new(StringComparer.OrdinalIgnoreCase);
    private WorkbookPart _workbookPart;
    private SharedStringTablePart? _sharedStringPart;
    private uint _nextSheetId = 1;
    private uint _nextTableId = 1;

    // Excel table display names must be unique workbook-wide (case-insensitive).
    private readonly HashSet<string> _tableNames = new(StringComparer.OrdinalIgnoreCase);

    private WorkbookBuilder(SpreadsheetDocument document, string? path, bool isNew, MemoryStream? documentStream = null)
    {
        _document = document;
        _path = path;
        _documentStream = documentStream;
        _isNewDocument = isNew;
        _workbookPart = document.WorkbookPart!;

        if (isNew)
        {
            InitializeNewWorkbook();
        }
        else
        {
            LoadExistingWorksheets();
        }
    }

    public static IWorkbookBuilder Create(string path)
    {
        var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        return new WorkbookBuilder(document, path, true);
    }

    public static IWorkbookBuilder Open(string path)
    {
        var document = SpreadsheetDocument.Open(path, true);
        return new WorkbookBuilder(document, path, false);
    }

    public static IWorkbookBuilder Create()
    {
        var memoryStream = new MemoryStream();
        var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
        document.AddWorkbookPart();
        return new WorkbookBuilder(document, null, true, memoryStream);
    }

    /// <summary>
    /// Opens an existing XLSX workbook from a stream.
    /// The stream content is copied to an internal buffer; the caller retains ownership of the original stream.
    /// </summary>
    public static IWorkbookBuilder Open(Stream stream)
    {
        var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        memoryStream.Position = 0;
        var document = SpreadsheetDocument.Open(memoryStream, true);
        return new WorkbookBuilder(document, null, false, memoryStream);
    }

    /// <summary>
    /// Opens an existing XLSX workbook from a byte array.
    /// The bytes are copied to an internal writable buffer so the caller cannot mutate the document's backing store.
    /// </summary>
    public static IWorkbookBuilder Open(byte[] bytes)
    {
        var memoryStream = new MemoryStream(bytes.Length);
        memoryStream.Write(bytes, 0, bytes.Length);
        memoryStream.Position = 0;
        var document = SpreadsheetDocument.Open(memoryStream, true);
        return new WorkbookBuilder(document, null, false, memoryStream);
    }

    /// <summary>
    /// Regex matching characters illegal in Excel worksheet names: : \ / ? * [ ]
    /// </summary>
    private static readonly Regex InvalidSheetNameChars = new(@"[:\\\/\?\*\[\]]", RegexOptions.Compiled);

    public IWorksheetBuilder AddWorksheet(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new XlsxException("Worksheet name cannot be empty or whitespace.");
        }

        if (name.Length > 31)
        {
            throw new XlsxException(
                $"Worksheet name '{name}' is {name.Length} characters; Excel limits names to 31 characters.");
        }

        if (InvalidSheetNameChars.IsMatch(name))
        {
            throw new XlsxException(
                $"Worksheet name '{name}' contains characters that are illegal in Excel " +
                $"(: \\ / ? * [ ]). Remove them and try again.");
        }

        if (_worksheets.ContainsKey(name))
        {
            throw new XlsxException(
                $"A worksheet named '{name}' already exists in this workbook. " +
                "Worksheet names must be unique (comparison is case-insensitive, as in Excel).");
        }

        var worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        var worksheet = new Worksheet(
            new SheetData()
        );
        worksheetPart.Worksheet = worksheet;

        // Add sheet to workbook
        var workbook = _workbookPart.Workbook!;
        var sheets = workbook.Sheets ?? new Sheets();
        if (workbook.Sheets == null)
        {
            workbook.Append(sheets);
        }

        var sheet = new Sheet
        {
            Name = name,
            SheetId = _nextSheetId++,
            Id = _workbookPart.GetIdOfPart(worksheetPart)
        };
        sheets.Append(sheet);

        var worksheetBuilder = new WorksheetBuilder(worksheetPart, worksheet, this);
        _worksheets[name] = worksheetBuilder;

        return worksheetBuilder;
    }

    public IWorkbookBuilder RemoveWorksheet(string name)
    {
        if (!_worksheets.ContainsKey(name))
        {
            throw new ArgumentException($"Worksheet '{name}' not found.");
        }

        var worksheetBuilder = _worksheets[name];
        var worksheetPart = worksheetBuilder.WorksheetPart;

        // Remove from workbook sheets
        var workbook = _workbookPart.Workbook!;
        var sheets = workbook.Sheets;
        if (sheets != null)
        {
            var sheet = sheets.Elements<Sheet>().FirstOrDefault(
                s => string.Equals(s.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (sheet != null)
            {
                sheet.Remove();
            }
        }

        // Remove the worksheet part
        _workbookPart.DeletePart(worksheetPart);
        _worksheets.Remove(name);

        return this;
    }

    public IWorksheetBuilder GetWorksheet(string name)
    {
        if (!_worksheets.TryGetValue(name, out var worksheet))
        {
            throw new ArgumentException($"Worksheet '{name}' not found.");
        }
        return worksheet;
    }

    public List<string> GetWorksheetNames()
    {
        return _worksheets.Keys.ToList();
    }

    public List<VariableInfo> DetectVariables()
    {
        var detector = new Variables.XlsxVariableDetector();
        return detector.Scan(_document);
    }

    public IWorkbookBuilder MergeVariables(Dictionary<string, string> data)
    {
        var replacer = new Variables.XlsxVariableReplacer();
        replacer.Replace(_document, data);
        return this;
    }

    public void Save(string? path = null)
    {
        if (string.IsNullOrEmpty(path))
        {
            if (string.IsNullOrEmpty(_path))
            {
                throw new InvalidOperationException("This workbook was created in memory. Use Save(Stream) or SaveToBytes() to persist it.");
            }

            _document.Save();
            return;
        }

        _document.Save();

        if (!string.IsNullOrEmpty(_path) && Path.GetFullPath(path) == Path.GetFullPath(_path))
        {
            return;
        }

        using var clone = _document.Clone(path, true);
        clone.Save();
    }

    /// <summary>
    /// Writes the current workbook content to the provided stream and leaves it open.
    /// </summary>
    public void Save(Stream stream)
    {
        _document.Save();

        if (_documentStream != null)
        {
            _documentStream.Position = 0;
            _documentStream.CopyTo(stream);
            _documentStream.Position = 0;
        }
        else
        {
            using var fileStream = File.OpenRead(_path!);
            fileStream.CopyTo(stream);
        }
    }

    /// <summary>
    /// Returns the current workbook content as a byte array.
    /// </summary>
    public byte[] SaveToBytes()
    {
        _document.Save();

        if (_documentStream != null)
        {
            _documentStream.Position = 0;
            return _documentStream.ToArray();
        }

        return File.ReadAllBytes(_path!);
    }

    public void Dispose()
    {
        _document.Dispose();
        _documentStream?.Dispose();
    }

    internal string GetSharedString(string text)
    {
        var sharedStringPart = _sharedStringPart;
        if (sharedStringPart == null)
        {
            sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _sharedStringPart = sharedStringPart;
        }

        var sharedStringTable = sharedStringPart.SharedStringTable ??= new SharedStringTable();
        
        // Check if string already exists
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return item.InnerText;
            }
        }

        // Add new shared string
        var newItem = new SharedStringItem(new Text(text));
        sharedStringTable.Append(newItem);
        return text;
    }

    internal int GetSharedStringIndex(string text)
    {
        var sharedStringPart = _sharedStringPart;
        if (sharedStringPart == null)
        {
            sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _sharedStringPart = sharedStringPart;
        }

        var sharedStringTable = sharedStringPart.SharedStringTable ??= new SharedStringTable();
        int index = 0;
        
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return index;
            }
            index++;
        }

        // Add new shared string
        var newItem = new SharedStringItem(new Text(text));
        sharedStringTable.Append(newItem);
        return index;
    }

    internal string? GetSharedStringByIndex(int index)
    {
        var sharedStringPart = _workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        if (sharedStringPart?.SharedStringTable is not SharedStringTable table)
        {
            return null;
        }

        var items = table.Elements<SharedStringItem>().ToList();
        if (index < 0 || index >= items.Count)
        {
            return null;
        }

        return items[index].InnerText;
    }

    /// <summary>
    /// Ensures the workbook has a stylesheet whose cell formats (cellXfs) include
    /// <paramref name="styleIndex"/>; otherwise writing s= on the cell would corrupt
    /// the file (Excel repair prompt).
    /// </summary>
    internal void EnsureStyleIndexExists(uint styleIndex, string cellReference)
    {
        var cellFormats = _workbookPart.WorkbookStylesPart?.Stylesheet?.CellFormats;
        if (cellFormats == null)
        {
            throw new XlsxException(
                $"Cell {cellReference} references styleId {styleIndex}, but this workbook has no " +
                "stylesheet (no cell formats are defined). Create a style first " +
                "(e.g. via AddHeaderRow) or use the Phase 3 style builder.");
        }

        var count = cellFormats.Count?.Value ?? (uint)cellFormats.Elements<CellFormat>().Count();
        if (styleIndex >= count)
        {
            throw new XlsxException(
                $"Cell {cellReference} references styleId {styleIndex}, but the stylesheet only " +
                $"defines {count} cell format(s) (valid ids: 0–{count - 1}). " +
                "Create the style first or use the Phase 3 style builder.");
        }
    }

    internal uint EnsureHeaderStyleIndex()
    {
        var stylesPart = _workbookPart.WorkbookStylesPart ?? _workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= new Stylesheet();
        var stylesheet = stylesPart.Stylesheet;

        stylesheet.Fonts ??= new Fonts(new Font()) { Count = 1 };
        stylesheet.Fills ??= new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
        ) { Count = 2 };
        stylesheet.Borders ??= new Borders(new Border()) { Count = 1 };
        stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat()) { Count = 1 };
        stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1 };

        var fontId = stylesheet.Fonts.Count?.Value ?? (uint)stylesheet.Fonts.Elements<Font>().Count();
        stylesheet.Fonts.Append(new Font(new Bold()));
        stylesheet.Fonts.Count = fontId + 1;

        var styleIndex = stylesheet.CellFormats.Count?.Value ?? (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
        stylesheet.CellFormats.Append(new CellFormat { FontId = fontId, FillId = 0, BorderId = 0, ApplyFont = true });
        stylesheet.CellFormats.Count = styleIndex + 1;

        stylesPart.Stylesheet.Save();
        return styleIndex;
    }

    internal uint NextTableId()
    {
        return _nextTableId++;
    }

    /// <summary>
    /// Registers a table display name, throwing if it is already used in this workbook.
    /// </summary>
    internal void RegisterTableName(string tableName)
    {
        if (!_tableNames.Add(tableName))
        {
            throw new XlsxException(
                $"A table named '{tableName}' already exists in this workbook. " +
                "Table names must be unique workbook-wide (case-insensitive).");
        }
    }

    private void InitializeNewWorkbook()
    {
        _workbookPart.Workbook = new Workbook();
        _workbookPart.Workbook.Sheets = new Sheets();
    }

    private void LoadExistingWorksheets()
    {
        var workbook = _workbookPart.Workbook!;
        var sheets = workbook.Sheets;
        if (sheets == null) return;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value;
            if (name == null) continue;

            var worksheetPart = (WorksheetPart)_workbookPart.GetPartById(sheet.Id!);
            var worksheetBuilder = new WorksheetBuilder(worksheetPart, worksheetPart.Worksheet!, this);
            _worksheets[name] = worksheetBuilder;

            if (sheet.SheetId?.Value >= _nextSheetId)
            {
                _nextSheetId = sheet.SheetId.Value + 1;
            }
        }

        // Load shared string part if exists
        _sharedStringPart = _workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        // Scan for max existing table ID to avoid collisions, and register existing
        // table names so new tables cannot reuse them.
        uint maxTableId = 0;
        foreach (var wsPart in _workbookPart.WorksheetParts)
        {
            foreach (var tdPart in wsPart.TableDefinitionParts)
            {
                if (tdPart.Table?.Id?.Value > maxTableId)
                {
                    maxTableId = tdPart.Table.Id.Value;
                }

                if (tdPart.Table?.DisplayName?.Value is { Length: > 0 } displayName)
                {
                    _tableNames.Add(displayName);
                }
            }
        }
        _nextTableId = maxTableId + 1;
    }
}
