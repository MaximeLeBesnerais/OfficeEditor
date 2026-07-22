using DocxEditor.Core.Builders;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;
using OfficeEditor.Core.Models;

namespace OfficeEditor.Examples;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("╔════════════════════════════════════════════════════════╗");
        Console.WriteLine("║       OfficeEditor Examples - Complete Suite           ║");
        Console.WriteLine("╚════════════════════════════════════════════════════════╝");
        Console.WriteLine();

        Console.WriteLine("This program demonstrates all OfficeEditor features:");
        Console.WriteLine("  • DOCX - Word document creation and editing");
        Console.WriteLine("  • PPTX - PowerPoint presentations with Typst export");
        Console.WriteLine("  • XLSX - Excel workbooks with formulas");
        Console.WriteLine();

        // Create output directory
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        Console.WriteLine($"Output directory: {outputDir}");
        Console.WriteLine();

        // Run all examples
        try
        {
            Console.WriteLine("Starting DOCX examples...");
            Console.WriteLine();
            RunDocxExamples(outputDir);
            
            Console.WriteLine("\nStarting PPTX examples...");
            Console.WriteLine();
            RunPptxExamples(outputDir);
            
            Console.WriteLine("\nStarting XLSX examples...");
            Console.WriteLine();
            RunXlsxExamples(outputDir);
            
            Console.WriteLine("\n╔════════════════════════════════════════════════════════╗");
            Console.WriteLine("║              All Examples Completed!                   ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════╝");
            Console.WriteLine();
            Console.WriteLine($"Check the output folder: {outputDir}");
            Console.WriteLine();
            Console.WriteLine("Generated files:");
            Console.WriteLine("  DOCX - Word documents with various content");
            Console.WriteLine("  PPTX - PowerPoint presentations");
            Console.WriteLine("  XLSX - Excel workbooks with data and formulas");
            Console.WriteLine("  PDF  - Exported from PPTX via Typst");
            Console.WriteLine("  PNG  - Slide thumbnails from PPTX");
            Console.WriteLine("  TYP  - Typst source code from PPTX");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nError: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }

        Console.WriteLine("\nDone!");
    }

    static void RunDocxExamples(string baseOutputDir)
    {
        Console.WriteLine("=== DOCX Examples ===\n");
        
        var outputDir = Path.Combine(baseOutputDir, "docx");
        Directory.CreateDirectory(outputDir);
        
        // Example 1: Basic Document Creation
        BasicDocumentCreation(outputDir);
        
        // Example 2: Markdown Conversion
        MarkdownConversion(outputDir);
        
        // Example 3: Variable Detection
        VariableDetection(outputDir);
        
        // Example 4: Variable Replacement (Mail Merge)
        VariableReplacement(outputDir);
        
        Console.WriteLine($"\nDOCX examples completed! Check {outputDir}");
    }

    static void BasicDocumentCreation(string outputDir)
    {
        Console.WriteLine("1. Creating basic document...");
        
        var path = Path.Combine(outputDir, "01-basic.docx");
        using var builder = DocumentBuilder.Create(path);
        
        builder.AddParagraph("Welcome to OfficeEditor", "Heading1");
        builder.AddParagraph("This document was created using the fluent C# API.");
        builder.AddParagraph("You can easily add formatted text and paragraphs.");
        
        builder.AddParagraph("Features:", "Heading2");
        builder.AddParagraph("• Fluent API");
        builder.AddParagraph("• Markdown support");
        builder.AddParagraph("• JSON instructions");
        builder.AddParagraph("• Variable replacement");
        
        builder.AddParagraph("Getting Started:", "Heading2");
        builder.AddParagraph("1. Install the NuGet package");
        builder.AddParagraph("2. Create a DocumentBuilder");
        builder.AddParagraph("3. Add content");
        builder.AddParagraph("4. Save");
        
        builder.Save();
        Console.WriteLine($"   Created: {path}");
    }

    static void MarkdownConversion(string outputDir)
    {
        Console.WriteLine("2. Converting Markdown to DOCX...");
        
        var markdown = @"# Project Proposal

## Executive Summary

This is a **bold** proposal for a new *innovative* project.

## Key Benefits

1. Increased productivity
2. Cost savings
3. Better collaboration

## Implementation

> ""The best way to predict the future is to create it."" - Peter Drucker

### Phase 1: Planning
- Requirements gathering
- Resource allocation
- Timeline definition

### Phase 2: Development
- Core features
- Testing
- Documentation

---

*Generated by OfficeEditor*";
        
        var path = Path.Combine(outputDir, "02-markdown.docx");
        using var builder = DocumentBuilder.Create(path);
        builder.AddMarkdown(markdown);
        builder.Save();
        
        Console.WriteLine($"   Created: {path}");
    }

    static void VariableDetection(string outputDir)
    {
        Console.WriteLine("3. Detecting variables in template...");
        
        // Create template with variables
        var templatePath = Path.Combine(outputDir, "03-detection-template.docx");
        using (var builder = DocumentBuilder.Create(templatePath))
        {
            builder.AddParagraph("Invoice for {{customerName}}");
            builder.AddParagraph("Order #: {{orderNumber}}");
            builder.AddParagraph("Total: ${{totalAmount}}");
            builder.AddParagraph("Due date: {{dueDate}}");
            builder.Save();
        }
        
        // Detect variables
        using var detector = DocumentBuilder.Open(templatePath);
        var variables = detector.DetectVariables();
        
        Console.WriteLine($"   Template: {templatePath}");
        Console.WriteLine($"   Found {variables.Count} variables:");
        foreach (var v in variables)
        {
            Console.WriteLine($"     - {v.Name} (at {v.Location})");
        }
    }

    static void VariableReplacement(string outputDir)
    {
        Console.WriteLine("4. Variable replacement (mail merge)...");
        
        // Create template
        var templatePath = Path.Combine(outputDir, "04-merge-template.docx");
        using (var builder = DocumentBuilder.Create(templatePath))
        {
            builder.AddParagraph("Hello {{name}}!");
            builder.AddParagraph("");
            builder.AddParagraph("Your order details:");
            builder.AddParagraph("Order #: {{orderNumber}}");
            builder.AddParagraph("Total: ${{total}}");
            builder.AddParagraph("Status: {{status}}");
            builder.Save();
        }
        
        // Process template with variable replacement
        var outputPath = Path.Combine(outputDir, "04-merge-result.docx");
        using (var builder = DocumentBuilder.Open(templatePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["name"] = "Alice",
                ["orderNumber"] = "ORD-2024-001",
                ["total"] = "299.99",
                ["status"] = "Shipped"
            });
            builder.Save(outputPath);
        }
        
        Console.WriteLine($"   Template: {templatePath}");
        Console.WriteLine($"   Result: {outputPath}");
    }

    static void RunPptxExamples(string baseOutputDir)
    {
        Console.WriteLine("=== PPTX Examples ===\n");
        
        var outputDir = Path.Combine(baseOutputDir, "pptx");
        Directory.CreateDirectory(outputDir);
        
        // Example 1: Basic Presentation
        BasicPresentation(outputDir);
        
        // Example 2: Advanced Content
        AdvancedContent(outputDir);
        
        // Example 3: Variable Detection & Mail Merge
        VariableDetectionAndMerge(outputDir);
        
        // Example 4: Export to PDF via Typst
        ExportToPdf(outputDir);
        
        // Example 5: Export to Typst Source
        ExportTypstSource(outputDir);
        
        Console.WriteLine($"\nPPTX examples completed! Check {outputDir}");
    }

    static void BasicPresentation(string outputDir)
    {
        Console.WriteLine("1. Creating basic presentation...");
        
        var path = Path.Combine(outputDir, "01-basic.pptx");
        using var builder = PresentationBuilder.Create(path);
        
        // Title slide
        builder.AddSlide();
        builder.CurrentSlide
            .AddTitle("OfficeEditor Presentation")
            .AddSubtitle("Created with Fluent C# API");
        
        // Content slide 1
        builder.AddSlide();
        builder.CurrentSlide
            .AddTitle("Key Features")
            .AddBulletList(new[] {
                "Easy slide creation",
                "Multiple content types",
                "Variable detection",
                "Typst export"
            });
        
        // Content slide 2
        builder.AddSlide();
        builder.CurrentSlide
            .AddTitle("Getting Started")
            .AddText("1. Install the NuGet package")
            .AddText("2. Create a PresentationBuilder")
            .AddText("3. Add slides and content")
            .AddText("4. Save or export");
        
        builder.Save();
        Console.WriteLine($"   Created: {path}");
    }

    static void AdvancedContent(string outputDir)
    {
        Console.WriteLine("2. Creating presentation with advanced content...");
        
        var path = Path.Combine(outputDir, "02-advanced.pptx");
        using var builder = PresentationBuilder.Create(path);
        
        // Title slide
        builder.AddSlide();
        builder.CurrentSlide
            .AddTitle("Q4 Sales Report")
            .AddSubtitle("2024 Performance Review");
        
        // Table slide
        builder.AddSlide();
        builder.CurrentSlide
            .AddTitle("Sales by Region")
            .AddTable(new List<List<string>>
            {
                new() { "Region", "Q1", "Q2", "Q3", "Q4", "Total" },
                new() { "North", "$100K", "$120K", "$110K", "$140K", "$470K" },
                new() { "South", "$80K", "$95K", "$105K", "$130K", "$410K" },
                new() { "East", "$90K", "$100K", "$120K", "$150K", "$460K" },
                new() { "West", "$110K", "$125K", "$135K", "$160K", "$530K" }
            });
        
        // Numbered list slide
        builder.AddSlide();
        builder.CurrentSlide
            .AddTitle("Top Priorities")
            .AddNumberedList(new[] {
                "Expand to new markets",
                "Improve customer retention",
                "Launch new product line",
                "Optimize operations"
            });
        
        builder.Save();
        Console.WriteLine($"   Created: {path}");
    }

    static void VariableDetectionAndMerge(string outputDir)
    {
        Console.WriteLine("3. Variable detection and mail merge...");
        
        // Create template
        var templatePath = Path.Combine(outputDir, "03-template.pptx");
        using (var builder = PresentationBuilder.Create(templatePath))
        {
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Welcome {{companyName}}")
                .AddSubtitle("Presented by {{presenterName}}");
            
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Project: {{projectName}}")
                .AddText("Budget: ${{budget}}")
                .AddText("Timeline: {{timeline}}")
                .AddText("Status: {{status}}");
            
            builder.Save();
        }
        
        // Detect variables
        List<VariableInfo> variables;
        using (var detector = PresentationBuilder.Open(templatePath))
        {
            variables = detector.DetectVariables();
        }
        
        Console.WriteLine($"   Template: {templatePath}");
        Console.WriteLine($"   Found {variables.Count} variables:");
        foreach (var v in variables)
        {
            Console.WriteLine($"     - {v.Name}");
        }
        
        // Mail merge
        var outputPath = Path.Combine(outputDir, "03-merged.pptx");
        using (var builder = PresentationBuilder.Open(templatePath))
        {
            builder.MergeVariables(new Dictionary<string, string>
            {
                ["companyName"] = "TechCorp Inc.",
                ["presenterName"] = "Jane Smith",
                ["projectName"] = "Cloud Migration",
                ["budget"] = "500,000",
                ["timeline"] = "6 months",
                ["status"] = "On Track"
            });
            builder.Save(outputPath);
        }
        
        Console.WriteLine($"   Merged: {outputPath}");
    }

    static void ExportToPdf(string outputDir)
    {
        Console.WriteLine("4. Exporting to PDF via Typst...");
        
        // Create a presentation
        var pptxPath = Path.Combine(outputDir, "04-export.pptx");
        using (var builder = PresentationBuilder.Create(pptxPath))
        {
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("PDF Export Demo")
                .AddSubtitle("Using Typst Integration");
            
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Features")
                .AddBulletList(new[] {
                    "High-fidelity PDF output",
                    "Preserves formatting",
                    "Embeds fonts and images",
                    "Fast compilation"
                });
            
            builder.Save();
        }
        
        // Export to PDF
        var pdfPath = Path.Combine(outputDir, "04-export.pdf");
        using (var builder = PresentationBuilder.Open(pptxPath))
        {
            try
            {
                var pdfBytes = builder.ExportToPdf();
                File.WriteAllBytes(pdfPath, pdfBytes);
                Console.WriteLine($"   PDF: {pdfPath} ({pdfBytes.Length} bytes)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   Note: PDF export uses TypstBridge first, with Typst CLI fallback.");
                Console.WriteLine($"   Error: {ex.Message}");
            }
        }
        
        Console.WriteLine($"   PPTX: {pptxPath}");
    }

    static void ExportTypstSource(string outputDir)
    {
        Console.WriteLine("5. Exporting to Typst source code...");
        
        // Create a presentation
        var pptxPath = Path.Combine(outputDir, "05-typst-source.pptx");
        using (var builder = PresentationBuilder.Create(pptxPath))
        {
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Typst Source Export")
                .AddSubtitle("View the generated code");
            
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("How it works")
                .AddText("1. Extract PPTX content")
                .AddText("2. Convert to Typst elements")
                .AddText("3. Generate Typst source")
                .AddText("4. Compile with TypstBridge");
            
            builder.Save();
        }
        
        // Export Typst source
        var typPath = Path.Combine(outputDir, "05-source.typ");
        using (var builder = PresentationBuilder.Open(pptxPath))
        {
            var typstSource = builder.ExportToTypst();
            File.WriteAllText(typPath, typstSource);
        }
        
        Console.WriteLine($"   PPTX: {pptxPath}");
        Console.WriteLine($"   Typst: {typPath}");
        Console.WriteLine($"   (Open {typPath} to see the generated Typst code)");
    }

    static void RunXlsxExamples(string baseOutputDir)
    {
        Console.WriteLine("=== XLSX Examples ===\n");
        
        var outputDir = Path.Combine(baseOutputDir, "xlsx");
        Directory.CreateDirectory(outputDir);
        
        // Example 1: Basic Workbook
        BasicWorkbook(outputDir);
        
        // Example 2: Formulas and Calculations
        FormulasAndCalculations(outputDir);
        
        // Example 3: Multiple Worksheets
        MultipleWorksheets(outputDir);
        
        // Example 4: Variable Detection
        VariableDetectionXlsx(outputDir);
        
        // Example 5: Variable Replacement
        VariableReplacementXlsx(outputDir);
        
        Console.WriteLine($"\nXLSX examples completed! Check {outputDir}");
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

    static void VariableDetectionXlsx(string outputDir)
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

    static void VariableReplacementXlsx(string outputDir)
    {
        Console.WriteLine("5. Variable replacement in spreadsheet...");
        
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
            builder.Save(outputPath);
        }
        
        Console.WriteLine($"   Template: {templatePath}");
        Console.WriteLine($"   Processed: {outputPath}");
    }
}
