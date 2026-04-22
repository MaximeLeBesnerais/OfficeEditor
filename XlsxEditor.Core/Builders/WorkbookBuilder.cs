using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeEditor.Core.Models;

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
}

public interface IWorksheetBuilder
{
    IWorksheetBuilder AddCell(string cellReference, string value);
    IWorksheetBuilder AddCell(string cellReference, string formula, bool isFormula);
    IWorksheetBuilder AddCell(string cellReference, string value, string styleId);
    IWorksheetBuilder AddHeaderRow(List<string> values, int rowIndex = 1);
    IWorksheetBuilder AddDataRow(List<string> values, int rowIndex);
    IWorksheetBuilder AddFormulaRow(List<string> formulas, int rowIndex);
    IWorksheetBuilder AddTable(string startCell, string endCell, string tableName);
    IWorksheetBuilder AddChart(ChartType type, string dataRange);
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
    private readonly bool _isNewDocument;
    private readonly Dictionary<string, WorksheetBuilder> _worksheets = new();
    private WorkbookPart _workbookPart;
    private SharedStringTablePart? _sharedStringPart;
    private uint _nextSheetId = 1;

    private WorkbookBuilder(SpreadsheetDocument document, bool isNew)
    {
        _document = document;
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
        return new WorkbookBuilder(document, true);
    }

    public static IWorkbookBuilder Open(string path)
    {
        var document = SpreadsheetDocument.Open(path, true);
        return new WorkbookBuilder(document, false);
    }

    public IWorksheetBuilder AddWorksheet(string name)
    {
        var worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        var worksheet = new Worksheet(
            new SheetData()
        );
        worksheetPart.Worksheet = worksheet;

        // Add sheet to workbook
        var sheets = _workbookPart.Workbook.Sheets ?? new Sheets();
        if (_workbookPart.Workbook.Sheets == null)
        {
            _workbookPart.Workbook.Append(sheets);
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
        var sheets = _workbookPart.Workbook.Sheets;
        if (sheets != null)
        {
            var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name?.Value == name);
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
        _document.Save();
    }

    public void Dispose()
    {
        _document.Dispose();
    }

    internal string GetSharedString(string text)
    {
        if (_sharedStringPart == null)
        {
            _sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _sharedStringPart.SharedStringTable = new SharedStringTable();
        }

        var sharedStringTable = _sharedStringPart.SharedStringTable;
        
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
        if (_sharedStringPart == null)
        {
            _sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _sharedStringPart.SharedStringTable = new SharedStringTable();
        }

        var sharedStringTable = _sharedStringPart.SharedStringTable;
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

    private void InitializeNewWorkbook()
    {
        _workbookPart.Workbook = new Workbook();
        _workbookPart.Workbook.Sheets = new Sheets();
    }

    private void LoadExistingWorksheets()
    {
        var sheets = _workbookPart.Workbook.Sheets;
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
    }
}
