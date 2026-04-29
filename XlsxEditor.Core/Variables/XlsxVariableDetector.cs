using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeEditor.Core.Models;

namespace XlsxEditor.Core.Variables;

public class XlsxVariableDetector
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public List<VariableInfo> Scan(SpreadsheetDocument document)
    {
        var variables = new List<VariableInfo>();
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) return variables;

        var workbook = workbookPart.Workbook;
        if (workbook == null) return variables;

        var sheets = workbook.Sheets;
        if (sheets == null) return variables;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var sheetName = sheet.Name?.Value ?? "Unknown";
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null) continue;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellText = GetCellText(cell, workbookPart);
                    var cellReference = cell.CellReference?.Value ?? "Unknown";
                    var location = $"sheet:{sheetName}:cell:{cellReference}";
                    
                    var matches = VariablePattern.Matches(cellText);
                    foreach (Match match in matches)
                    {
                        var variableName = match.Groups[1].Value.Trim();
                        var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;
                        
                        if (!variables.Any(v => v.Name == variableName && v.Location == location))
                        {
                            variables.Add(new VariableInfo
                            {
                                Name = variableName,
                                FullMatch = match.Value,
                                Location = location,
                                DefaultValue = defaultValue
                            });
                        }
                    }
                }
            }
        }

        return variables;
    }

    private string GetCellText(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType?.Value == CellValues.SharedString && cell.CellValue?.Text != null)
        {
            if (int.TryParse(cell.CellValue.Text, out var sharedStringIndex))
            {
                var sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringPart?.SharedStringTable != null)
                {
                    var items = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ToList();
                    if (sharedStringIndex >= 0 && sharedStringIndex < items.Count)
                    {
                        return items[sharedStringIndex].InnerText;
                    }
                }
            }
        }
        
        return cell.CellValue?.Text ?? string.Empty;
    }
}
