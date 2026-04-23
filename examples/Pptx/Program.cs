using PptxEditor.Core.Builders;
using System.Text.Json;

namespace PptxExamples;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== PPTX Examples ===\n");
        
        var outputDir = Path.Combine("..", "output", "pptx");
        Directory.CreateDirectory(outputDir);
        
        // Example 1: Basic Presentation
        BasicPresentation(outputDir);
        
        // Example 2: Advanced Content
        AdvancedContent(outputDir);
        
        // Example 3: Variable Detection & Mail Merge
        VariableDetectionAndMerge(outputDir);
        
        // Example 4: Export to PDF via Typst
        ExportToPdf(outputDir);
        
        // Example 5: Export Slide Thumbnails
        ExportThumbnails(outputDir);
        
        // Example 6: Export to Typst Source
        ExportTypstSource(outputDir);
        
        Console.WriteLine($"\nAll PPTX examples completed! Check {outputDir}");
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
        using var detector = PresentationBuilder.Open(templatePath);
        var variables = detector.DetectVariables();
        
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
            builder.SaveAs(outputPath);
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
            var pdfBytes = builder.ExportToPdf();
            File.WriteAllBytes(pdfPath, pdfBytes);
        }
        
        Console.WriteLine($"   PPTX: {pptxPath}");
        Console.WriteLine($"   PDF: {pdfPath}");
    }
    
    static void ExportThumbnails(string outputDir)
    {
        Console.WriteLine("5. Exporting slide thumbnails...");
        
        // Create a presentation
        var pptxPath = Path.Combine(outputDir, "05-thumbnails.pptx");
        using (var builder = PresentationBuilder.Create(pptxPath))
        {
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Slide 1")
                .AddText("First slide content");
            
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Slide 2")
                .AddText("Second slide content");
            
            builder.AddSlide();
            builder.CurrentSlide
                .AddTitle("Slide 3")
                .AddText("Third slide content");
            
            builder.Save();
        }
        
        // Export thumbnails
        using (var builder = PresentationBuilder.Open(pptxPath))
        {
            var thumbnails = builder.ExportThumbnails(new ThumbnailOptions 
            { 
                Ppi = 150 
            });
            
            for (int i = 0; i < thumbnails.Length; i++)
            {
                var pngPath = Path.Combine(outputDir, $"05-slide-{i + 1}.png");
                File.WriteAllBytes(pngPath, thumbnails[i]);
                Console.WriteLine($"   Thumbnail {i + 1}: {pngPath}");
            }
        }
        
        Console.WriteLine($"   Generated {thumbnails.Length} thumbnails at 150 PPI");
    }
    
    static void ExportTypstSource(string outputDir)
    {
        Console.WriteLine("6. Exporting to Typst source code...");
        
        // Create a presentation
        var pptxPath = Path.Combine(outputDir, "06-typst-source.pptx");
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
                .AddText("4. Compile with typstsharp");
            
            builder.Save();
        }
        
        // Export Typst source
        var typPath = Path.Combine(outputDir, "06-source.typ");
        using (var builder = PresentationBuilder.Open(pptxPath))
        {
            var typstSource = builder.ExportToTypst();
            File.WriteAllText(typPath, typstSource);
        }
        
        Console.WriteLine($"   PPTX: {pptxPath}");
        Console.WriteLine($"   Typst: {typPath}");
        Console.WriteLine($"   (Open {typPath} to see the generated Typst code)");
    }
}
