using System.Text.RegularExpressions;
using System.Xml;
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

        // A package that opens as a SpreadsheetDocument but has no workbook main part
        // (e.g. an empty stream, a wrong-format package, or a missing /xl/workbook.xml)
        // must fail here at the Open boundary instead of surfacing as a null-deref later.
        _workbookPart = GetRequiredWorkbookPart(document);

        if (isNew)
        {
            InitializeNewWorkbook();
        }
        else
        {
            // Accessing Workbook loads/parses the part, so a corrupt or wrong-typed root
            // (XmlException / InvalidDataException) is raised here and normalized by the
            // Open boundary, and a relationship targeting a missing part surfaces as the
            // SDK's InvalidOperationException (converted at this structural site). A part
            // that parses but has no root element yields null.
            var workbook = GetWorkbookRoot(_workbookPart);
            if (workbook is null)
            {
                throw new XlsxException(
                    "This file is not a valid XLSX workbook: the workbook part has no " +
                    "workbook root element.");
            }

            LoadExistingWorksheets();
        }
    }

    /// <summary>
    /// Returns the workbook part's root element. Loading it parses the part and, in the
    /// process, the SDK resolves every relationship declared on the workbook part — a
    /// relationship targeting a part that does not exist in the package surfaces as the
    /// SDK's <see cref="InvalidOperationException"/>. That structural failure is converted to
    /// <see cref="XlsxException"/> here, at the exact call site that triggers it (preserving
    /// the original as the inner exception), instead of being classified as malformed input at
    /// the Open boundary — which would also swallow genuine programmer errors of the broad
    /// <see cref="InvalidOperationException"/> family. An <see cref="ObjectDisposedException"/>
    /// (an <see cref="InvalidOperationException"/> subclass) always propagates unchanged.
    /// </summary>
    private static Workbook? GetWorkbookRoot(WorkbookPart workbookPart)
    {
        try
        {
            return workbookPart.Workbook;
        }
        catch (InvalidOperationException ex) when (ex is not ObjectDisposedException)
        {
            throw new XlsxException(
                "This file is not a valid XLSX workbook: the workbook part references a " +
                "relationship target that does not exist in the package, so its content " +
                "cannot be loaded.",
                ex);
        }
    }

    /// <summary>
    /// Returns the required workbook main part, or a domain exception when it cannot be
    /// produced. The SDK returns null when the package has no officeDocument relationship,
    /// and throws <see cref="InvalidOperationException"/> when that relationship targets a
    /// part that does not exist in the package. Both structural failures surface as
    /// <see cref="XlsxException"/> (the original SDK failure preserved as the inner
    /// exception). Converting at this structural site — instead of classifying the broad
    /// <see cref="InvalidOperationException"/> family at the Open boundary — keeps a genuine
    /// programmer bug such as an <see cref="ObjectDisposedException"/> (an
    /// <see cref="InvalidOperationException"/> subclass) from being mislabeled as malformed
    /// input; it propagates unchanged.
    /// </summary>
    private static WorkbookPart GetRequiredWorkbookPart(SpreadsheetDocument document)
    {
        try
        {
            return document.WorkbookPart
                ?? throw new XlsxException(
                    "This file is not a valid XLSX workbook: it has no workbook main part. " +
                    "Expected a package whose workbook part resolves to /xl/workbook.xml.");
        }
        catch (InvalidOperationException ex) when (ex is not ObjectDisposedException)
        {
            throw new XlsxException(
                "This file is not a valid XLSX workbook: the package's workbook relationship " +
                "targets a part that does not exist. Expected a package whose workbook part " +
                "resolves to /xl/workbook.xml.",
                ex);
        }
    }

    /// <summary>
    /// Creates a new XLSX workbook at the given path and returns a builder over it.
    /// The path must be a non-empty, non-whitespace file path, mirroring
    /// <see cref="Open(string)"/>; null, empty, or whitespace paths are rejected up front as
    /// argument exceptions. If creating or initializing the package fails, the opened document
    /// is disposed best-effort so no file handle leaks and no partial file is left behind, and
    /// the original failure propagates unchanged — a dispose error during cleanup can never mask
    /// the primary create/initialization failure. Ordinary path/permission and IO errors are
    /// never normalized into the domain exception. On success the returned builder owns the
    /// document; callers must dispose it.
    /// </summary>
    public static IWorkbookBuilder Create(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty or whitespace.", nameof(path));
        }

        SpreadsheetDocument? document = null;
        try
        {
            document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            return new WorkbookBuilder(document, path, true);
        }
        catch
        {
            // A failed create must never leak the partially opened package handle; the
            // caller-visible failure is preserved and rethrown unchanged. Cleanup is
            // best-effort so a dispose error cannot replace the primary failure.
            DisposeFailedOpen(document);
            throw;
        }
    }

    public static IWorkbookBuilder Open(string path)
    {
        ArgumentNullException.ThrowIfNull(path);

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path cannot be empty or whitespace.", nameof(path));
        }

        SpreadsheetDocument? document = null;
        try
        {
            document = SpreadsheetDocument.Open(path, true);
            return new WorkbookBuilder(document, path, false);
        }
        catch (XlsxException)
        {
            // Our own structural validation rejected the package after the SDK opened it:
            // dispose the half-opened document, then let the typed exception propagate
            // unchanged so callers still see exactly the failure we raised.
            DisposeFailedOpen(document);
            throw;
        }
        catch (Exception ex) when (IsInvalidPackageFailure(ex))
        {
            DisposeFailedOpen(document);
            throw new XlsxException(
                $"Cannot open the XLSX workbook at '{path}': the file is not a valid or " +
                "readable OpenXML spreadsheet package.", ex);
        }
        catch
        {
            // Anything else — programmer errors, FileNotFoundException, permission or path
            // errors — propagates unchanged (after cleanup) instead of being normalized.
            DisposeFailedOpen(document);
            throw;
        }
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
        ArgumentNullException.ThrowIfNull(stream);

        SpreadsheetDocument? document = null;
        MemoryStream? memoryStream = null;
        try
        {
            memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            memoryStream.Position = 0;
            document = SpreadsheetDocument.Open(memoryStream, true);
            return new WorkbookBuilder(document, null, false, memoryStream);
        }
        catch (XlsxException)
        {
            DisposeFailedOpen(document, memoryStream);
            throw;
        }
        catch (Exception ex) when (IsInvalidPackageFailure(ex))
        {
            DisposeFailedOpen(document, memoryStream);
            throw new XlsxException(
                "Cannot open the XLSX workbook: the input stream is not a valid or " +
                "readable OpenXML spreadsheet package.", ex);
        }
        catch
        {
            DisposeFailedOpen(document, memoryStream);
            throw;
        }
    }

    /// <summary>
    /// Opens an existing XLSX workbook from a byte array.
    /// The bytes are copied to an internal writable buffer so the caller cannot mutate the document's backing store.
    /// </summary>
    public static IWorkbookBuilder Open(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        SpreadsheetDocument? document = null;
        MemoryStream? memoryStream = null;
        try
        {
            memoryStream = new MemoryStream(bytes.Length);
            memoryStream.Write(bytes, 0, bytes.Length);
            memoryStream.Position = 0;
            document = SpreadsheetDocument.Open(memoryStream, true);
            return new WorkbookBuilder(document, null, false, memoryStream);
        }
        catch (XlsxException)
        {
            DisposeFailedOpen(document, memoryStream);
            throw;
        }
        catch (Exception ex) when (IsInvalidPackageFailure(ex))
        {
            DisposeFailedOpen(document, memoryStream);
            throw new XlsxException(
                "Cannot open the XLSX workbook: the byte array is not a valid or " +
                "readable OpenXML spreadsheet package.", ex);
        }
        catch
        {
            DisposeFailedOpen(document, memoryStream);
            throw;
        }
    }

    /// <summary>
    /// The well-known exception families the OpenXML SDK raises when a package is not a
    /// valid XLSX workbook (corrupt zip, wrong root element, malformed XML, unresolvable
    /// part content type). These are parse failures of malformed input at the public Open
    /// boundary and are normalized into <see cref="XlsxException"/>.
    /// Deliberately excludes the broad programmer exception types — <see cref="InvalidOperationException"/>
    /// (and its <see cref="ObjectDisposedException"/> subclass) and <see cref="ArgumentOutOfRangeException"/> —
    /// which the SDK also uses for missing/broken relationships: those are converted to the domain
    /// exception at the exact structural sites that call into the SDK (see
    /// <see cref="GetRequiredWorkbookPart"/> and <see cref="ResolveWorksheetPart"/>), so a genuine
    /// programmer bug can never be mislabeled as malformed input. Filesystem errors are also
    /// excluded; those propagate untouched.
    /// </summary>
    private static bool IsInvalidPackageFailure(Exception ex) =>
        ex is FileFormatException
            or InvalidDataException
            or XmlException
            or OpenXmlPackageException;

    /// <summary>
    /// Best-effort cleanup after a failed open. Only the OpenXML document and the
    /// internally-owned buffer are disposed; the caller's source stream is never touched.
    /// Cleanup exceptions are deliberately ignored so they cannot mask the primary failure.
    /// </summary>
    private static void DisposeFailedOpen(SpreadsheetDocument? document, MemoryStream? ownedStream = null)
    {
        try
        {
            document?.Dispose();
        }
        catch
        {
            // The primary open failure is rethrown by the caller; a dispose error here must not replace it.
        }

        ownedStream?.Dispose();
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

        // Excel forbids sheet names that begin or end with an apostrophe (a leading
        // apostrophe in particular is Excel's escape character for the R1C1-style
        // and would corrupt the workbook).
        if (name[0] == '\'' || name[^1] == '\'')
        {
            throw new XlsxException(
                $"Worksheet name '{name}' begins or ends with an apostrophe ('), " +
                "which Excel does not allow. Remove the leading or trailing apostrophe.");
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
        ArgumentNullException.ThrowIfNull(data);

        // A null replacement value would be written as an empty string — never the
        // caller's intent. Fail loudly before any cell is mutated.
        foreach (var pair in data)
        {
            if (pair.Value is null)
            {
                throw new ArgumentNullException(
                    $"data['{pair.Key}']",
                    $"Value for variable '{pair.Key}' must not be null; use an empty string to clear a value.");
            }
        }

        var replacer = new Variables.XlsxVariableReplacer();
        replacer.Replace(_document, data);
        return this;
    }

    public void Save(string? path = null)
    {
        // Reject empty/whitespace destinations before flushing so an invalid path fails fast
        // without touching the source document or materializing a junk file; matches the
        // Create/Open path contract.
        if (path is not null && string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be empty or whitespace.", nameof(path));
        }

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

    // Text → index cache so writes are O(1) instead of rescanning the table per cell.
    private Dictionary<string, int>? _sharedStringIndices;

    private Dictionary<string, int> GetSharedStringIndices()
    {
        if (_sharedStringIndices == null)
        {
            _sharedStringIndices = new Dictionary<string, int>(StringComparer.Ordinal);
            var table = _sharedStringPart?.SharedStringTable;
            if (table != null)
            {
                var index = 0;
                foreach (var item in table.Elements<SharedStringItem>())
                {
                    _sharedStringIndices.TryAdd(item.InnerText, index);
                    index++;
                }
            }
        }

        return _sharedStringIndices;
    }

    internal int GetSharedStringIndex(string text)
    {
        var indices = GetSharedStringIndices();
        if (indices.TryGetValue(text, out var existingIndex))
        {
            return existingIndex;
        }

        var sharedStringPart = _sharedStringPart;
        if (sharedStringPart == null)
        {
            sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _sharedStringPart = sharedStringPart;
        }

        var sharedStringTable = sharedStringPart.SharedStringTable ??= new SharedStringTable();

        // xml:space="preserve" so leading/trailing whitespace survives the
        // save → reload round-trip.
        var index = sharedStringTable.Elements<SharedStringItem>().Count();
        var newItem = new SharedStringItem(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        sharedStringTable.Append(newItem);
        indices[text] = index;
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

    // Cached so repeated AddHeaderRow calls reuse one bold style instead of
    // appending a new font + cell format to the stylesheet every time.
    private uint? _headerStyleIndex;

    internal uint EnsureHeaderStyleIndex()
    {
        if (_headerStyleIndex is { } cached)
        {
            return cached;
        }

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
        _headerStyleIndex = styleIndex;
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

            var relationshipId = sheet.Id?.Value
                ?? throw new XlsxException(
                    $"Sheet '{name}' is missing its relationship id (r:id), so its worksheet " +
                    "part cannot be located.");

            var worksheetPart = ResolveWorksheetPart(relationshipId, name);

            // Load the worksheet root now: a corrupt or wrong-typed root surfaces here
            // (normalized by the Open boundary) instead of on first worksheet use.
            var worksheet = worksheetPart.Worksheet
                ?? throw new XlsxException(
                    $"The worksheet part for sheet '{name}' (relationship '{relationshipId}') " +
                    "has no worksheet root element.");

            var worksheetBuilder = new WorksheetBuilder(worksheetPart, worksheet, this);
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

    /// <summary>
    /// Resolves a sheet's relationship id to a <see cref="WorksheetPart"/>, rejecting parts
    /// of any other type. The SDK's failure modes for a broken sheet relationship — a missing
    /// target part (<see cref="InvalidOperationException"/>) or an undeclared relationship id
    /// (<see cref="ArgumentOutOfRangeException"/>) — are converted to <see cref="XlsxException"/>
    /// here, at the exact call site that triggers them, with the original failure preserved as
    /// the inner exception. Converting at the structural site (rather than at the Open boundary)
    /// keeps the boundary from misclassifying genuine programmer errors of the same broad types;
    /// an <see cref="ObjectDisposedException"/> (an <see cref="InvalidOperationException"/> subclass)
    /// always propagates unchanged.
    /// </summary>
    private WorksheetPart ResolveWorksheetPart(string relationshipId, string sheetName)
    {
        OpenXmlPart part;
        try
        {
            part = _workbookPart.GetPartById(relationshipId);
        }
        catch (InvalidOperationException ex) when (ex is not ObjectDisposedException)
        {
            throw new XlsxException(
                $"Sheet '{sheetName}' references relationship '{relationshipId}', which targets " +
                "a part that does not exist in this workbook, so the worksheet cannot be loaded.",
                ex);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            throw new XlsxException(
                $"Sheet '{sheetName}' references relationship '{relationshipId}', which is not " +
                "declared in this workbook's relationships, so the worksheet cannot be loaded.",
                ex);
        }

        if (part is not WorksheetPart worksheetPart)
        {
            throw new XlsxException(
                $"Sheet '{sheetName}' references relationship '{relationshipId}', which " +
                $"resolves to a {part.GetType().Name} part instead of a worksheet part.");
        }

        return worksheetPart;
    }
}
