using XlsxEditor.Core.Builders;

namespace XlsxExamples;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== XLSX Examples ===\n");
        
        var outputDir = Path.Combine("..", "output", "xlsx");
        Directory.CreateDirectory(outputDir);
        
        // Example 1: Basic Workbook
        BasicWorkbook(outputDir);
        
        // Example 2: Formulas and Calculations
        FormulasAndCalculations(outputDir);
        
        // Example 3: Multiple Worksheets
        MultipleWorksheets(outputDir);
        
        // Example 4: Variable Detection
        VariableDetection(outputDir);
        
        // Example 5: Template Processing
        TemplateProcessing(outputDir);
        
        Console.WriteLine($"\nAll XLSX examples completed! Check {outputDir}");
    }
    
    static void BasicWorkbook(string outputDir)
    {
        Console.WriteLine("1. Creating basic workbook...");
        
        var path = Path.Combine(outputDir, "01-basic.xlsx");
        using var builder = WorkbookBuilder.Create(path);
        
        // Create sales worksheet
        var sales = builder.AddWorksheet("Sales");
        
        // Add header row
        sales.AddHeaderRow(new List<string> { "Product", "Q1", "Q2", "Q3", "Q4" });
        
        // Add data rows
        sales.AddDataRow(new List<string> { "Widget", "100", "150", "200", "250" }, 2);
        sales.AddDataRow(new List<string> { "Gadget", "80", "120", "160", "200" }, 3);
        sales.AddDataRow(new List<string> { "Tool", "60", "90", "120", "150" }, 4);
        
        builder.Save();
        Console.WriteLine($"   Created: {path}");
    }
    
    static void FormulasAndCalculations(string outputDir)
    {
        Console.WriteLine("2. Creating workbook with formulas...");
        
        var path = Path.Combine(outputDir, "02-formulas.xlsx");
        using var builder = WorkbookBuilder.Create(path);
        
        var sheet = builder.AddWorksheet("Budget");
        
        // Headers
        sheet.AddHeaderRow(new List<string> { "Category", "Q1", "Q2", "Q3", "Q4", "Total" });
        
        // Data rows with formulas
        sheet.AddDataRow(new List<string> { "Revenue", "10000", "12000", "11000", "15000" }, 2);
        sheet.AddDataRow(new List<string> { "Expenses", "6000", "7000", "6500", "8000" }, 3);
        sheet.AddDataRow(new List<string> { "Profit", "4000", "5000", "4500", "7000" }, 4);
        
        // Add totals row with formulas
        sheet.AddFormulaRow(new List<string> 
        { 
            "TOTAL", 
            "=SUM(B2:B4)", 
            "=SUM(C2:C4)", 
            "=SUM(D2:D4)", 
            "=SUM(E2:E4)",
            "=SUM(F2:F4)"
        }, 5);
        
        // Add average row
        sheet.AddFormulaRow(new List<string> 
        { 
            "AVERAGE", 
            "=AVERAGE(B2:B4)", 
            "=AVERAGE(C2:C4)", 
            "=AVERAGE(D2:D4)", 
            "=AVERAGE(E2:E4)",
            "=AVERAGE(F2:F4)"
        }, 6);
        
        builder.Save();
        Console.WriteLine($"   Created: {path}");
    }
    
    static void MultipleWorksheets(string outputDir)
    {
        Console.WriteLine("3. Creating workbook with multiple worksheets...");
        
        var path = Path.Combine(outputDir, "03-multi-sheet.xlsx");
        using var builder = WorkbookBuilder.Create(path);
        
        // Sales data sheet
        var sales = builder.AddWorksheet("Sales Data");
        sales.AddHeaderRow(new List<string> { "Month", "Revenue", "Expenses", "Profit" });
        sales.AddDataRow(new List<string> { "Jan", "10000", "6000", "4000" }, 2);
        sales.AddDataRow(new List<string> { "Feb", "12000", "7000", "5000" }, 3);
        sales.AddDataRow(new List<string> { "Mar", "11000", "6500", "4500" }, 4);
        
        // Summary sheet with formulas referencing Sales Data
        var summary = builder.AddWorksheet("Summary");
        summary.AddHeaderRow(new List<string> { "Metric", "Value" });
        summary.AddDataRow(new List<string> { "Total Revenue", "" }, 2);
        summary.AddCell("B2", "=SUM('Sales Data'!B2:B4)", true);
        summary.AddDataRow(new List<string> { "Total Expenses", "" }, 3);
        summary.AddCell("B3", "=SUM('Sales Data'!C2:C4)", true);
        summary.AddDataRow(new List<string> { "Total Profit", "" }, 4);
        summary.AddCell("B4", "=SUM('Sales Data'!D2:D4)", true);
        summary.AddDataRow(new List<string> { "Average Profit", "" }, 5);
        summary.AddCell("B5", "=AVERAGE('Sales Data'!D2:D4)", true);
        
        builder.Save();
        Console.WriteLine($"   Created: {path}");
    }
    
    static void VariableDetection(string outputDir)
    {
        Console.WriteLine("4. Detecting variables in spreadsheet...");
        
        // Create template
        var templatePath = Path.Combine(outputDir, "04-template.xlsx");
        using (var builder = WorkbookBuilder.Create(templatePath))
        {
            var sheet = builder.AddWorksheet("Report");
            sheet.AddHeaderRow(new List<string> { "Field", "Value" });
            sheet.AddDataRow(new List<string> { "Company", "{{companyName}}" }, 2);
            sheet.AddDataRow(new List<string> { "Report Date", "{{reportDate}}" }, 3);
            sheet.AddDataRow(new List<string> { "Prepared By", "{{preparedBy}}" }, 4);
            sheet.AddDataRow(new List<string> { "Total Sales", "{{totalSales}}" }, 5);
            builder.Save();
        }
        
        // Detect variables
        using var detector = WorkbookBuilder.Open(templatePath);
        var variables = detector.DetectVariables();
        
        Console.WriteLine($"   Template: {templatePath}");
        Console.WriteLine($"   Found {variables.Count} variables:");
        foreach (var v in variables)
        {
            Console.WriteLine($"     - {v.Name}");
        }
    }
    
    static void TemplateProcessing(string outputDir)
    {
        Console.WriteLine("5. Processing spreadsheet template...");
        
        // Create template
        var templatePath = Path.Combine(outputDir, "05-template.xlsx");
        using (var builder = WorkbookBuilder.Create(templatePath))
        {
            var sheet = builder.AddWorksheet("Invoice");
            sheet.AddHeaderRow(new List<string> { "Description", "Quantity", "Price", "Total" });
            sheet.AddDataRow(new List<string> { "{{item1}}", "{{qty1}}", "{{price1}}", "=B2*C2" }, 2);
            sheet.AddDataRow(new List<string> { "{{item2}}", "{{qty2}}", "{{price2}}", "=B3*C3" }, 3);
            sheet.AddDataRow(new List<string> { "{{item3}}", "{{qty3}}", "{{price3}}", "=B4*C4" }, 4);
            sheet.AddDataRow(new List<string> { "TOTAL", "", "", "=SUM(D2:D4)" }, 5);
            builder.Save();
        }
        
        // Process template
        var outputPath = Path.Combine(outputDir, "05-processed.xlsx");
        using (var builder = WorkbookBuilder.Open(templatePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["item1"] = "Consulting Services",
                ["qty1"] = "10",
                ["price1"] = "150",
                ["item2"] = "Software License",
                ["qty2"] = "1",
                ["price2"] = "500",
                ["item3"] = "Training",
                ["qty3"] = "5",
                ["price3"] = "100"
            });
            builder.SaveAs(outputPath);
        }
        
        Console.WriteLine($"   Template: {templatePath}");
        Console.WriteLine($"   Processed: {outputPath}");
    }
}
