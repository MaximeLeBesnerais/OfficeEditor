using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeEditor.Core.Models;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Builders;

public interface IPresentationBuilder : IDisposable
{
    IPresentationBuilder AddSlide(string? layoutName = null);
    IPresentationBuilder RemoveSlide(int index);
    IPresentationBuilder ReorderSlide(int fromIndex, int toIndex);
    
    // Content
    ISlideBuilder CurrentSlide { get; }
    ISlideBuilder GetSlide(int index);
    int SlideCount { get; }
    
    // Variables
    List<VariableInfo> DetectVariables();
    IPresentationBuilder MergeVariables(Dictionary<string, string> data);
    IPresentationBuilder ProcessTemplate(Dictionary<string, object> data);
    
    // Anatomizer
    List<SlideAnatomy> Analyze();
    IPresentationBuilder ReplaceElement(uint elementId, string newText);
    IPresentationBuilder ReplaceTable(uint elementId, List<List<string>> newData);
    IPresentationBuilder ReplaceImage(uint elementId, string newImagePath);
    
    // Export to Typst/PDF/Thumbnails
    string ExportToTypst();
    byte[] ExportToPdf(PdfOptions? options = null);
    byte[][] ExportThumbnails(ThumbnailOptions? options = null);
    
    void Save(string? path = null);
}

public sealed record PdfOptions
{
    public bool SingleFile { get; init; } = true;
}

public sealed record ThumbnailOptions
{
    public float Ppi { get; init; } = 150;
    public string Format { get; init; } = "png";
}

public interface ISlideBuilder
{
    ISlideBuilder AddTitle(string text);
    ISlideBuilder AddSubtitle(string text);
    ISlideBuilder AddText(string text);
    ISlideBuilder AddBulletList(IEnumerable<string> items);
    ISlideBuilder AddNumberedList(IEnumerable<string> items);
    ISlideBuilder AddImage(string imagePath);
    ISlideBuilder AddTable(List<List<string>> rows);
    ISlideBuilder AddChart(ChartType type, Dictionary<string, int> data);
}

public enum ChartType
{
    Bar,
    Line,
    Pie,
    Column
}

public class PresentationBuilder : IPresentationBuilder
{
    private readonly PresentationDocument _document;
    private readonly bool _isNewDocument;
    private readonly List<SlideBuilder> _slides = new();
    private int _currentSlideIndex = -1;

    private PresentationBuilder(PresentationDocument document, bool isNew)
    {
        _document = document;
        _isNewDocument = isNew;
        
        if (isNew)
        {
            InitializeNewPresentation();
        }
        else
        {
            LoadExistingSlides();
        }
    }

    public static IPresentationBuilder Create(string path)
    {
        var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        return new PresentationBuilder(document, true);
    }

    public static IPresentationBuilder Open(string path)
    {
        var document = PresentationDocument.Open(path, true);
        return new PresentationBuilder(document, false);
    }

    public ISlideBuilder CurrentSlide => _currentSlideIndex >= 0 && _currentSlideIndex < _slides.Count 
        ? _slides[_currentSlideIndex] 
        : throw new InvalidOperationException("No slide selected.");

    public int SlideCount => _slides.Count;

    public IPresentationBuilder AddSlide(string? layoutName = null)
    {
        var presentation = _document.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList!;
        
        // Get slide master and layout
        var slideMasterPart = _document.PresentationPart!.SlideMasterParts.First();
        SlideLayoutPart slideLayoutPart;
        
        if (!string.IsNullOrEmpty(layoutName))
        {
            slideLayoutPart = slideMasterPart.SlideLayoutParts
                .FirstOrDefault(slp => slp.SlideLayout.CommonSlideData?.Name?.Value == layoutName)
                ?? slideMasterPart.SlideLayoutParts.First();
        }
        else
        {
            // Default to first layout (usually "Title Slide" or "Title and Content")
            slideLayoutPart = slideMasterPart.SlideLayoutParts.First();
        }

        // Create new slide
        var slidePart = _document.PresentationPart.AddNewPart<SlidePart>();
        var slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 0, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(
                        new Drawing.TransformGroup(
                            new Drawing.Offset { X = 0, Y = 0 },
                            new Drawing.Extents { Cx = 0, Cy = 0 },
                            new Drawing.ChildOffset { X = 0, Y = 0 },
                            new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                        )
                    )
                )
            )
        );
        slidePart.Slide = slide;
        slidePart.AddPart(slideLayoutPart);

        // Add slide to presentation
        var slideIds = slideIdList.ChildElements
            .OfType<SlideId>()
            .Select(sid => sid.Id?.Value ?? 0)
            .ToList();
        var maxSlideId = slideIds.Count > 0 ? slideIds.Max() : 256;
        
        var newSlideId = new SlideId
        {
            Id = maxSlideId + 1,
            RelationshipId = _document.PresentationPart.GetIdOfPart(slidePart)
        };
        slideIdList.Append(newSlideId);

        // Create slide builder
        var slideBuilder = new SlideBuilder(slidePart, slide);
        _slides.Add(slideBuilder);
        _currentSlideIndex = _slides.Count - 1;

        return this;
    }

    public IPresentationBuilder RemoveSlide(int index)
    {
        if (index < 0 || index >= _slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var slideBuilder = _slides[index];
        var slidePart = slideBuilder.SlidePart;
        var presentation = _document.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList!;

        // Find and remove the slide ID
        var slideId = slideIdList.ChildElements
            .OfType<SlideId>()
            .FirstOrDefault(sid => sid.RelationshipId?.Value == _document.PresentationPart.GetIdOfPart(slidePart));
        
        if (slideId != null)
        {
            slideId.Remove();
        }

        // Remove the slide part
        _document.PresentationPart.DeletePart(slidePart);
        _slides.RemoveAt(index);

        if (_currentSlideIndex >= _slides.Count)
        {
            _currentSlideIndex = _slides.Count - 1;
        }

        return this;
    }

    public IPresentationBuilder ReorderSlide(int fromIndex, int toIndex)
    {
        if (fromIndex < 0 || fromIndex >= _slides.Count || toIndex < 0 || toIndex >= _slides.Count)
        {
            throw new ArgumentOutOfRangeException();
        }

        var slideBuilder = _slides[fromIndex];
        _slides.RemoveAt(fromIndex);
        _slides.Insert(toIndex, slideBuilder);

        // Reorder in presentation
        var presentation = _document.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList!;
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
        
        var slideId = slideIds[fromIndex];
        slideId.Remove();
        
        // Insert at the correct position
        var targetIndex = fromIndex < toIndex ? toIndex - 1 : toIndex;
        if (targetIndex >= slideIdList.ChildElements.Count)
        {
            slideIdList.Append(slideId);
        }
        else
        {
            var targetSlideId = slideIdList.ChildElements.OfType<SlideId>().ElementAt(targetIndex);
            slideIdList.InsertAfter(targetSlideId, slideId);
        }

        if (_currentSlideIndex == fromIndex)
        {
            _currentSlideIndex = toIndex;
        }

        return this;
    }

    public ISlideBuilder GetSlide(int index)
    {
        if (index < 0 || index >= _slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }
        
        _currentSlideIndex = index;
        return _slides[index];
    }

    public List<VariableInfo> DetectVariables()
    {
        var detector = new Variables.PptxVariableDetector();
        return detector.Scan(_document);
    }

    public IPresentationBuilder MergeVariables(Dictionary<string, string> data)
    {
        var replacer = new Variables.PptxVariableReplacer();
        replacer.Replace(_document, data);
        return this;
    }

    public IPresentationBuilder ProcessTemplate(Dictionary<string, object> data)
    {
        var engine = new Variables.PptxTemplateEngine();
        engine.Process(_document, data);
        return this;
    }

    public List<SlideAnatomy> Analyze()
    {
        var anatomizer = new Services.PptxAnatomizer();
        return anatomizer.Analyze(_document);
    }

    public IPresentationBuilder ReplaceElement(uint elementId, string newText)
    {
        if (_currentSlideIndex >= 0 && _currentSlideIndex < _slides.Count)
        {
            var slidePart = _slides[_currentSlideIndex].SlidePart;
            var replacer = new Services.PptxElementReplacer();
            replacer.ReplaceText(slidePart, elementId, newText);
        }
        return this;
    }

    public IPresentationBuilder ReplaceTable(uint elementId, List<List<string>> newData)
    {
        if (_currentSlideIndex >= 0 && _currentSlideIndex < _slides.Count)
        {
            var slidePart = _slides[_currentSlideIndex].SlidePart;
            var replacer = new Services.PptxElementReplacer();
            replacer.ReplaceTableData(slidePart, elementId, newData);
        }
        return this;
    }

    public IPresentationBuilder ReplaceImage(uint elementId, string newImagePath)
    {
        if (_currentSlideIndex >= 0 && _currentSlideIndex < _slides.Count)
        {
            var slidePart = _slides[_currentSlideIndex].SlidePart;
            var replacer = new Services.PptxElementReplacer();
            replacer.ReplaceImage(slidePart, elementId, newImagePath);
        }
        return this;
    }

    public string ExportToTypst()
    {
        using var converter = new Converters.PptxToTypstConverter(_document);
        var presentation = converter.Convert();
        return converter.GenerateTypstSource(presentation);
    }

    public byte[] ExportToPdf(PdfOptions? options = null)
    {
        options ??= new PdfOptions();
        
        using var converter = new Converters.PptxToTypstConverter(_document);
        var presentation = converter.Convert();
        var typstSource = converter.GenerateTypstSource(presentation);
        
        // Write Typst source to temp file
        var typstFile = Path.Combine(presentation.TempDirectory, "presentation.typ");
        File.WriteAllText(typstFile, typstSource);
        
        using var compiler = new OfficeEditor.Core.Services.TypstCompilerService();
        var compileOptions = new OfficeEditor.Core.Services.CompileOptions
        {
            Format = OfficeEditor.Core.Services.OutputFormat.Pdf,
            FontDirectory = presentation.FontFiles.Count > 0 ? Path.Combine(presentation.TempDirectory, "fonts") : null
        };
        
        var result = compiler.Compile(typstSource, compileOptions);
        
        if (result.Pages.Length == 0)
        {
            throw new InvalidOperationException("PDF compilation produced no output.");
        }
        
        return result.Pages[0];
    }

    public byte[][] ExportThumbnails(ThumbnailOptions? options = null)
    {
        options ??= new ThumbnailOptions();
        
        using var converter = new Converters.PptxToTypstConverter(_document);
        var presentation = converter.Convert();
        var typstSource = converter.GenerateTypstSource(presentation);
        
        using var compiler = new OfficeEditor.Core.Services.TypstCompilerService();
        var compileOptions = new OfficeEditor.Core.Services.CompileOptions
        {
            Format = OfficeEditor.Core.Services.OutputFormat.Png,
            Ppi = options.Ppi,
            FontDirectory = presentation.FontFiles.Count > 0 ? Path.Combine(presentation.TempDirectory, "fonts") : null
        };
        
        var result = compiler.Compile(typstSource, compileOptions);
        return result.Pages;
    }

    public void Save(string? path = null)
    {
        if (path != null)
        {
            // Clone to new path
            var newDoc = _document.Clone(path);
            newDoc.Dispose();
        }
        else
        {
            _document.Save();
        }
    }

    public void Dispose()
    {
        _document.Dispose();
    }

    private void InitializeNewPresentation()
    {
        var presentationPart = _document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

        // Create slide master
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideMaster = new SlideMaster(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 0, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(
                        new Drawing.TransformGroup(
                            new Drawing.Offset { X = 0, Y = 0 },
                            new Drawing.Extents { Cx = 0, Cy = 0 },
                            new Drawing.ChildOffset { X = 0, Y = 0 },
                            new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                        )
                    ),
                    new P.Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "Title Placeholder" },
                            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Title })
                        ),
                        new ShapeProperties(),
                        new TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph()
                        )
                    ),
                    new P.Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 2, Name = "Content Placeholder" },
                            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 })
                        ),
                        new ShapeProperties(),
                        new TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph()
                        )
                    )
                )
            ),
            new ColorMap
            {
                Background1 = Drawing.ColorSchemeIndexValues.Light1,
                Text1 = Drawing.ColorSchemeIndexValues.Dark1,
                Background2 = Drawing.ColorSchemeIndexValues.Light2,
                Text2 = Drawing.ColorSchemeIndexValues.Dark2,
                Accent1 = Drawing.ColorSchemeIndexValues.Accent1,
                Accent2 = Drawing.ColorSchemeIndexValues.Accent2,
                Accent3 = Drawing.ColorSchemeIndexValues.Accent3,
                Accent4 = Drawing.ColorSchemeIndexValues.Accent4,
                Accent5 = Drawing.ColorSchemeIndexValues.Accent5,
                Accent6 = Drawing.ColorSchemeIndexValues.Accent6,
                Hyperlink = Drawing.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = Drawing.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList()
        );
        slideMasterPart.SlideMaster = slideMaster;
        presentationPart.Presentation.SlideMasterIdList.Append(new SlideMasterId
        {
            Id = 2147483648,
            RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
        });

        // Create default slide layout
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        var slideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 0, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(
                        new Drawing.TransformGroup(
                            new Drawing.Offset { X = 0, Y = 0 },
                            new Drawing.Extents { Cx = 0, Cy = 0 },
                            new Drawing.ChildOffset { X = 0, Y = 0 },
                            new Drawing.ChildExtents { Cx = 0, Cy = 0 }
                        )
                    )
                )
            )
        );
        slideLayoutPart.SlideLayout = slideLayout;
        slideLayoutPart.AddPart(slideMasterPart);

        // Link layout to master
        var layoutId = new SlideLayoutId
        {
            Id = 2147483649,
            RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
        };
        slideMaster.SlideLayoutIdList!.Append(layoutId);

        // Initialize slide ID list
        presentationPart.Presentation.SlideIdList = new SlideIdList();
        presentationPart.Presentation.SlideSize = new SlideSize
        {
            Cx = 9144000,
            Cy = 6858000,
            Type = SlideSizeValues.Screen4x3
        };
        presentationPart.Presentation.NotesSize = new NotesSize
        {
            Cx = 6858000,
            Cy = 9144000
        };
        
        // Add theme
        var themePart = slideMasterPart.AddNewPart<ThemePart>();
        themePart.Theme = new Drawing.Theme(
            new Drawing.ThemeElements(
                new Drawing.ColorScheme(
                    new Drawing.Dark1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
                    new Drawing.Light1Color(new Drawing.SystemColor { Val = Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new Drawing.Dark2Color(new Drawing.RgbColorModelHex { Val = "1F497D" }),
                    new Drawing.Light2Color(new Drawing.RgbColorModelHex { Val = "EEECE1" }),
                    new Drawing.Accent1Color(new Drawing.RgbColorModelHex { Val = "4F81BD" }),
                    new Drawing.Accent2Color(new Drawing.RgbColorModelHex { Val = "C0504D" }),
                    new Drawing.Accent3Color(new Drawing.RgbColorModelHex { Val = "9BBB59" }),
                    new Drawing.Accent4Color(new Drawing.RgbColorModelHex { Val = "8064A2" }),
                    new Drawing.Accent5Color(new Drawing.RgbColorModelHex { Val = "4BACC6" }),
                    new Drawing.Accent6Color(new Drawing.RgbColorModelHex { Val = "F79646" }),
                    new Drawing.Hyperlink(new Drawing.RgbColorModelHex { Val = "0000FF" }),
                    new Drawing.FollowedHyperlinkColor(new Drawing.RgbColorModelHex { Val = "800080" })
                ) { Name = "Office" },
                new Drawing.FontScheme(
                    new Drawing.MajorFont(
                        new Drawing.LatinFont { Typeface = "Calibri" },
                        new Drawing.EastAsianFont { Typeface = "" },
                        new Drawing.ComplexScriptFont { Typeface = "" }
                    ),
                    new Drawing.MinorFont(
                        new Drawing.LatinFont { Typeface = "Calibri" },
                        new Drawing.EastAsianFont { Typeface = "" },
                        new Drawing.ComplexScriptFont { Typeface = "" }
                    )
                ) { Name = "Office" },
                new Drawing.FormatScheme(
                    new Drawing.FillStyleList(
                        new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor }),
                        new Drawing.GradientFill(),
                        new Drawing.NoFill()
                    ),
                    new Drawing.LineStyleList(
                        new Drawing.Outline(
                            new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor }),
                            new Drawing.PresetDash { Val = Drawing.PresetLineDashValues.Solid }
                        ) { Width = 6350, CapType = Drawing.LineCapValues.Flat, CompoundLineType = Drawing.CompoundLineValues.Single },
                        new Drawing.Outline(
                            new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor }),
                            new Drawing.PresetDash { Val = Drawing.PresetLineDashValues.Solid }
                        ) { Width = 12700, CapType = Drawing.LineCapValues.Flat, CompoundLineType = Drawing.CompoundLineValues.Single },
                        new Drawing.Outline(
                            new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor }),
                            new Drawing.PresetDash { Val = Drawing.PresetLineDashValues.Solid }
                        ) { Width = 19050, CapType = Drawing.LineCapValues.Flat, CompoundLineType = Drawing.CompoundLineValues.Single }
                    ),
                    new Drawing.EffectStyleList(
                        new Drawing.EffectStyle(new Drawing.EffectList()),
                        new Drawing.EffectStyle(new Drawing.EffectList()),
                        new Drawing.EffectStyle(new Drawing.EffectList())
                    ),
                    new Drawing.BackgroundFillStyleList(
                        new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor }),
                        new Drawing.SolidFill(new Drawing.SchemeColor { Val = Drawing.SchemeColorValues.PhColor }),
                        new Drawing.GradientFill()
                    )
                ) { Name = "Office" }
            ),
            new Drawing.ObjectDefaults(),
            new Drawing.ExtraColorSchemeList()
        ) { Name = "Office Theme" };

        // Add package properties
        _document.PackageProperties.Creator = "OfficeEditor";
        _document.PackageProperties.Created = DateTime.Now;
        _document.PackageProperties.Modified = DateTime.Now;
    }

    private void LoadExistingSlides()
    {
        var presentation = _document.PresentationPart!.Presentation;
        var slideIdList = presentation.SlideIdList;
        
        if (slideIdList == null) return;

        foreach (var slideId in slideIdList.ChildElements.OfType<SlideId>())
        {
            var slidePart = (SlidePart)_document.PresentationPart.GetPartById(slideId.RelationshipId!);
            var slideBuilder = new SlideBuilder(slidePart, slidePart.Slide!);
            _slides.Add(slideBuilder);
        }

        if (_slides.Count > 0)
        {
            _currentSlideIndex = 0;
        }
    }
}
