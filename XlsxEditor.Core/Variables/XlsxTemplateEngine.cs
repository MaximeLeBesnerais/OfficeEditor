using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxEditor.Core.Variables;

public class XlsxTemplateEngine
{
    private static readonly Regex IfPattern = new(
        @"\{\{#if\s+([^}]+)\}\}(.*?)\{\{/if\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex IfNotPattern = new(
        @"\{\{#ifnot\s+([^}]+)\}\}(.*?)\{\{/ifnot\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex EachPattern = new(
        @"\{\{#each\s+([^}]+)\}\}(.*?)\{\{/each\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex ComparisonPattern = new(
        @"^(\w+)\s*([\u003e\u003c=!]+)\s*(.+)$",
        RegexOptions.Compiled);

    public void Process(SpreadsheetDocument document, Dictionary<string, object> data)
    {
        // Reject null arguments up front so behavior is not marker-dependent: a null data
        // dictionary was previously only touched when a cell actually contained a '{{…}}'
        // marker (NRE/ANE deep in evaluation), while a document without markers "worked".
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(data);

        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) return;

        var workbook = workbookPart.Workbook;
        if (workbook == null) return;

        var sheets = workbook.Sheets;
        if (sheets == null) return;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null) continue;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>().ToList())
            {
                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    ProcessCell(cell, data, workbookPart);
                }
            }
        }
    }

    private void ProcessCell(Cell cell, Dictionary<string, object> data, WorkbookPart workbookPart)
    {
        var cellText = GetCellText(cell, workbookPart);
        
        if (!cellText.Contains("{{")) return;

        var newText = ProcessConditionals(cellText, data);
        newText = ProcessLoops(newText, data);

        if (newText != cellText)
        {
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringPart?.SharedStringTable != null)
                {
                    var newIndex = GetOrAddSharedString(sharedStringPart.SharedStringTable, newText);
                    cell.CellValue = new CellValue(newIndex.ToString());
                }
            }
            else if (cell.DataType?.Value == CellValues.InlineString && cell.InlineString is not null)
            {
                // Inline-string cells keep their text in <is><t>, not <v>. Writing a
                // t="str" literal here would leave the stale <is> alongside the new <v>
                // and change the cell's type; rewrite the inline string in place instead.
                cell.InlineString = new InlineString(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });
                cell.CellValue = null;
            }
            else
            {
                cell.CellValue = new CellValue(newText);
                cell.DataType = CellValues.String;
            }
        }
    }

    private string GetCellText(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType?.Value == CellValues.InlineString && cell.InlineString is not null)
        {
            return cell.InlineString.InnerText;
        }

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

        // Append new shared string. xml:space="preserve" so leading/trailing whitespace
        // in a processed value survives (Excel trims <t> text without it), matching the
        // contract the WorkbookBuilder applies to its own shared-string writes.
        var newItem = new SharedStringItem(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        sharedStringTable.Append(newItem);
        return index;
    }

    private string ProcessConditionals(string text, Dictionary<string, object> data)
    {
        text = IfPattern.Replace(text, match =>
        {
            var condition = match.Groups[1].Value.Trim();
            var content = match.Groups[2].Value;
            
            return EvaluateCondition(condition, data) ? content : string.Empty;
        });

        text = IfNotPattern.Replace(text, match =>
        {
            var condition = match.Groups[1].Value.Trim();
            var content = match.Groups[2].Value;
            
            return !EvaluateCondition(condition, data) ? content : string.Empty;
        });

        return text;
    }

    private string ProcessLoops(string text, Dictionary<string, object> data)
    {
        return EachPattern.Replace(text, match =>
        {
            var arrayName = match.Groups[1].Value.Trim();
            var template = match.Groups[2].Value;
            
            if (data.TryGetValue(arrayName, out var arrayValue) && arrayValue is List<Dictionary<string, object>> array)
            {
                var results = new List<string>();
                foreach (var item in array)
                {
                    var itemText = template;
                    foreach (var kvp in item)
                    {
                        itemText = itemText.Replace($"{{{kvp.Key}}}", kvp.Value?.ToString() ?? "");
                    }
                    results.Add(itemText);
                }
                return string.Join("\n", results);
            }
            
            return string.Empty;
        });
    }

    private bool EvaluateCondition(string condition, Dictionary<string, object> data)
    {
        var comparisonMatch = ComparisonPattern.Match(condition);
        if (comparisonMatch.Success)
        {
            var left = comparisonMatch.Groups[1].Value.Trim();
            var op = comparisonMatch.Groups[2].Value.Trim();
            var right = comparisonMatch.Groups[3].Value.Trim().Trim('\'', '"');

            if (data.TryGetValue(left, out var leftValue))
            {
                var leftStr = leftValue?.ToString() ?? "";
                
                return op switch
                {
                    "==" => leftStr == right,
                    "!=" => leftStr != right,
                    ">" => CompareValues(leftStr, right) > 0,
                    "<" => CompareValues(leftStr, right) < 0,
                    ">=" => CompareValues(leftStr, right) >= 0,
                    "<=" => CompareValues(leftStr, right) <= 0,
                    _ => false
                };
            }
            
            return false;
        }

        if (data.TryGetValue(condition, out var value))
        {
            if (value is bool boolValue)
                return boolValue;
            
            var strValue = value?.ToString() ?? "";
            return !string.IsNullOrEmpty(strValue) && 
                   !strValue.Equals("false", StringComparison.OrdinalIgnoreCase) &&
                   !strValue.Equals("0", StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    private int CompareValues(string left, string right)
    {
        if (double.TryParse(left, out var leftNum) && double.TryParse(right, out var rightNum))
        {
            return leftNum.CompareTo(rightNum);
        }
        
        return string.Compare(left, right, StringComparison.Ordinal);
    }
}
