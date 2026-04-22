using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxEditor.Core.Variables;

public class XlsxVariableReplacer
{
    private static readonly Regex VariablePattern = new(
        @"\{\{([^}|]+)(?:\|([^}]*))?\}\}",
        RegexOptions.Compiled);

    public void Replace(SpreadsheetDocument document, Dictionary<string, string> data)
    {
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) return;

        var sheets = workbookPart.Workbook.Sheets;
        if (sheets == null) return;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    ReplaceInCell(cell, data, workbookPart);
                }
            }
        }
    }

    private void ReplaceInCell(Cell cell, Dictionary<string, string> data, WorkbookPart workbookPart)
    {
        var cellText = GetCellText(cell, workbookPart);
        
        if (!cellText.Contains("{{")) return;

        var newText = VariablePattern.Replace(cellText, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var defaultValue = match.Groups[2].Success ? match.Groups[2].Value : null;

            if (data.TryGetValue(variableName, out var value))
            {
                return value;
            }
            
            if (!string.IsNullOrEmpty(defaultValue))
            {
                return defaultValue;
            }
            
            return match.Value;
        });

        // Update cell value
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            // Update shared string
            var sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sharedStringPart?.SharedStringTable != null)
            {
                var newIndex = GetOrAddSharedString(sharedStringPart.SharedStringTable, newText);
                cell.CellValue = new CellValue(newIndex.ToString());
            }
        }
        else
        {
            cell.CellValue = new CellValue(newText);
            cell.DataType = CellValues.String;
        }
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

    private int GetOrAddSharedString(SharedStringTable sharedStringTable, string text)
    {
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
}
