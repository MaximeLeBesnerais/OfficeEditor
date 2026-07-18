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
    IPresentationBuilder DuplicateSlide(int index, int? position = null);
    
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
    byte[] ExportThumbnail(int slideIndex, ThumbnailOptions? options = null);

    void Save(string? path = null);
    void Save(Stream stream);
    byte[] SaveToBytes();

    static abstract IPresentationBuilder Create();
    static abstract IPresentationBuilder Open(Stream stream);
    static abstract IPresentationBuilder Open(byte[] bytes);
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
    private readonly string? _filePath;
    private readonly MemoryStream? _documentStream;

    private PresentationBuilder(PresentationDocument document, bool isNew, string? filePath = null, MemoryStream? documentStream = null)
    {
        _document = document;
        _isNewDocument = isNew;
        _filePath = filePath;
        _documentStream = documentStream;

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
        return new PresentationBuilder(document, true, path);
    }

    public static IPresentationBuilder Open(string path)
    {
        var document = PresentationDocument.Open(path, true);
        return new PresentationBuilder(document, false, path);
    }

    public static IPresentationBuilder Create()
    {
        var memoryStream = new MemoryStream();
        var document = PresentationDocument.Create(memoryStream, PresentationDocumentType.Presentation);
        return new PresentationBuilder(document, true, null, memoryStream);
    }

    public static IPresentationBuilder Open(Stream stream)
    {
        var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        memoryStream.Position = 0;
        var document = PresentationDocument.Open(memoryStream, true);
        return new PresentationBuilder(document, false, null, memoryStream);
    }

    public static IPresentationBuilder Open(byte[] bytes)
    {
        var memoryStream = new MemoryStream(bytes.Length);
        memoryStream.Write(bytes, 0, bytes.Length);
        memoryStream.Position = 0;
        var document = PresentationDocument.Open(memoryStream, true);
        return new PresentationBuilder(document, false, null, memoryStream);
    }

    public ISlideBuilder CurrentSlide => _currentSlideIndex >= 0 && _currentSlideIndex < _slides.Count
        ? _slides[_currentSlideIndex]
        : throw new InvalidOperationException("No slide selected.");

    public int SlideCount => _slides.Count;

    public IPresentationBuilder AddSlide(string? layoutName = null)
    {
        var presentation = _document.PresentationPart!.Presentation!;
        var slideIdList = presentation.SlideIdList!;

        // Get slide master and layout
        var slideMasterPart = _document.PresentationPart!.SlideMasterParts.First();
        SlideLayoutPart slideLayoutPart;

        if (!string.IsNullOrEmpty(layoutName))
        {
            slideLayoutPart = slideMasterPart.SlideLayoutParts
                .FirstOrDefault(slp => slp.SlideLayout?.CommonSlideData?.Name?.Value == layoutName)
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
        var newSlideId = new SlideId
        {
            Id = GetNextSlideId(slideIdList),
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
        var presentation = _document.PresentationPart!.Presentation!;
        var slideIdList = presentation.SlideIdList!;

        // Find and remove the slide ID
        var slideId = slideIdList.ChildElements
            .OfType<SlideId>()
            .FirstOrDefault(sid => sid.RelationshipId?.Value == _document.PresentationPart.GetIdOfPart(slidePart));
        
        if (slideId != null)
        {
            slideId.Remove();
        }

        // Prune parts referenced only by this slide (slide-only images, notes, ...) so they
        // do not stay orphaned in the package. Shared parts (layouts, images used by other
        // slides) survive. Master/layout/theme parts are never pruned: they are always
        // referenced from the master/presentation graph.
        var referencedElsewhere = CollectPartsReferencedOutside(slidePart);
        PruneUnreferencedParts(slidePart, referencedElsewhere, new HashSet<OpenXmlPartContainer>());

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
        var presentation = _document.PresentationPart!.Presentation!;
        var slideIdList = presentation.SlideIdList!;
        var slideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
        
        var slideId = slideIds[fromIndex];
        slideId.Remove();

        // After removal the list has one fewer entry; insert BEFORE the element currently
        // at toIndex so the slide lands exactly at toIndex (both for earlier and later moves).
        if (toIndex >= slideIdList.ChildElements.Count)
        {
            slideIdList.Append(slideId);
        }
        else
        {
            var targetSlideId = slideIdList.ChildElements.OfType<SlideId>().ElementAt(toIndex);
            slideIdList.InsertBefore(slideId, targetSlideId);
        }

        if (_currentSlideIndex == fromIndex)
        {
            _currentSlideIndex = toIndex;
        }

        return this;
    }

    /// <summary>
    /// Creates a verbatim copy of the slide at <paramref name="index"/> and inserts it at
    /// <paramref name="position"/> (defaults to right after the source slide).
    /// The slide XML is copied byte-for-byte (preserving all styling); image and layout
    /// parts are shared with the source slide, not copied. Every relationship of the
    /// source slide part is re-wired with an IDENTICAL relationship id so the copied
    /// XML's r:embed / r:id references stay valid (mismatched rIds = corrupt deck).
    /// Notes are per-slide state: the duplicate starts without a notes part.
    /// </summary>
    public IPresentationBuilder DuplicateSlide(int index, int? position = null)
    {
        if (index < 0 || index >= _slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var insertPosition = position ?? index + 1;
        if (insertPosition < 0 || insertPosition > _slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(position));
        }

        var presentationPart = _document.PresentationPart!;
        var slideIdList = presentationPart.Presentation!.SlideIdList!;
        var sourceSlidePart = _slides[index].SlidePart;

        // Verbatim copy of the slide XML (style-preservation rule).
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        using (var sourceStream = sourceSlidePart.GetStream())
        {
            newSlidePart.FeedData(sourceStream);
        }

        foreach (var pair in sourceSlidePart.Parts)
        {
            if (pair.OpenXmlPart is NotesSlidePart)
            {
                continue; // notes are per-slide state; the duplicate starts without them
            }

            newSlidePart.AddPart(pair.OpenXmlPart, pair.RelationshipId);
        }

        foreach (var hyperlink in sourceSlidePart.HyperlinkRelationships)
        {
            newSlidePart.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal, hyperlink.Id);
        }

        foreach (var external in sourceSlidePart.ExternalRelationships)
        {
            newSlidePart.AddExternalRelationship(external.RelationshipType, external.Uri, external.Id);
        }

        var newSlideId = new SlideId
        {
            Id = GetNextSlideId(slideIdList),
            RelationshipId = presentationPart.GetIdOfPart(newSlidePart)
        };

        var existingSlideIds = slideIdList.ChildElements.OfType<SlideId>().ToList();
        if (insertPosition >= existingSlideIds.Count)
        {
            slideIdList.Append(newSlideId);
        }
        else
        {
            slideIdList.InsertBefore(newSlideId, existingSlideIds[insertPosition]);
        }

        // Keep the _slides ↔ slideIdList invariant that ReorderSlide/RemoveSlide rely on.
        _slides.Insert(insertPosition, new SlideBuilder(newSlidePart, newSlidePart.Slide!));
        _currentSlideIndex = insertPosition;

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
            FontDirectory = presentation.FontFiles.Count > 0 ? Path.Combine(presentation.TempDirectory, "fonts") : null,
            WorkingDirectory = presentation.TempDirectory
        };
        
        var result = compiler.Compile(typstSource, compileOptions);
        
        if (result.Pages.Length == 0)
        {
            throw new InvalidOperationException($"PDF compilation produced no output. Error: {result.ErrorMessage}");
        }
        
        return result.Pages[0];
    }

    public byte[][] ExportThumbnails(ThumbnailOptions? options = null)
    {
        options ??= new ThumbnailOptions();

        var outputFormat = options.Format?.Trim().ToLowerInvariant() switch
        {
            null or "" or "png" => OfficeEditor.Core.Services.OutputFormat.Png,
            "svg" => OfficeEditor.Core.Services.OutputFormat.Svg,
            var unsupported => throw new ArgumentException(
                $"Unsupported thumbnail format '{unsupported}'. Supported formats: \"png\" (default), \"svg\".",
                nameof(options))
        };

        using var converter = new Converters.PptxToTypstConverter(_document);
        var presentation = converter.Convert();
        var typstSource = converter.GenerateTypstSource(presentation);

        using var compiler = new OfficeEditor.Core.Services.TypstCompilerService();
        var compileOptions = new OfficeEditor.Core.Services.CompileOptions
        {
            Format = outputFormat,
            Ppi = options.Ppi,
            FontDirectory = presentation.FontFiles.Count > 0 ? Path.Combine(presentation.TempDirectory, "fonts") : null,
            WorkingDirectory = presentation.TempDirectory
        };

        var result = compiler.Compile(typstSource, compileOptions);
        return result.Pages;
    }

    /// <summary>
    /// Renders a single slide to an image and returns the encoded bytes.
    /// </summary>
    /// <remarks>
    /// v0 implementation: renders the whole deck via <see cref="ExportThumbnails"/> and
    /// returns the page at <paramref name="slideIndex"/>. This full-render fallback assumes
    /// a 1:1 page-to-slide mapping, so the rendered page count is validated against
    /// <see cref="SlideCount"/> before slicing and an <see cref="InvalidOperationException"/>
    /// is thrown on mismatch (rather than returning the wrong page). The internals can be
    /// swapped to true single-slide compilation later without breaking callers.
    /// </remarks>
    public byte[] ExportThumbnail(int slideIndex, ThumbnailOptions? options = null)
    {
        if (slideIndex < 0 || slideIndex >= _slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(slideIndex));
        }

        var pages = ExportThumbnails(options);

        if (pages.Length != _slides.Count)
        {
            throw new InvalidOperationException(
                $"Rendered page count ({pages.Length}) does not match the slide count ({_slides.Count}); " +
                $"cannot safely slice page {slideIndex}. The full-render fallback requires a 1:1 page-to-slide mapping.");
        }

        return pages[slideIndex];
    }

    public void Save(string? path = null)
    {
        if (path != null)
        {
            // Clone to new path
            var newDoc = _document.Clone(path);
            newDoc.Dispose();
        }
        else if (_filePath != null)
        {
            _document.Save();
        }
        else
        {
            throw new InvalidOperationException("Cannot save stream-backed document without a path. Use Save(Stream) or SaveToBytes().");
        }
    }

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
            using var fileStream = File.OpenRead(_filePath!);
            fileStream.CopyTo(stream);
        }
    }

    public byte[] SaveToBytes()
    {
        _document.Save();

        if (_documentStream != null)
        {
            return _documentStream.ToArray();
        }

        return File.ReadAllBytes(_filePath!);
    }

    public void Dispose()
    {
        _document.Dispose();
        _documentStream?.Dispose();
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

    /// <summary>
    /// Collects every part reachable from the package root WITHOUT traversing
    /// <paramref name="excludedPart"/>. After <paramref name="excludedPart"/> is deleted,
    /// any part it referenced that is not in this set would be orphaned.
    /// The excluded part itself is still added to the set so back-references
    /// (e.g. notes slide → slide) are never treated as orphans.
    /// </summary>
    private HashSet<OpenXmlPart> CollectPartsReferencedOutside(OpenXmlPart excludedPart)
    {
        var referenced = new HashSet<OpenXmlPart>();
        var visitedContainers = new HashSet<OpenXmlPartContainer>();
        var stack = new Stack<OpenXmlPartContainer>();
        stack.Push(_document);

        while (stack.Count > 0)
        {
            var container = stack.Pop();
            if (!visitedContainers.Add(container))
            {
                continue;
            }

            foreach (var pair in container.Parts)
            {
                referenced.Add(pair.OpenXmlPart);
                if (!ReferenceEquals(pair.OpenXmlPart, excludedPart))
                {
                    stack.Push(pair.OpenXmlPart);
                }
            }
        }

        return referenced;
    }

    /// <summary>
    /// Recursively deletes every part under <paramref name="container"/> that is not
    /// referenced anywhere outside the subtree being removed. Children are pruned before
    /// their parent so exclusively-owned sub-parts do not leak.
    /// </summary>
    private static void PruneUnreferencedParts(
        OpenXmlPartContainer container,
        HashSet<OpenXmlPart> referencedElsewhere,
        HashSet<OpenXmlPartContainer> visited)
    {
        if (!visited.Add(container))
        {
            return;
        }

        foreach (var pair in container.Parts.ToList())
        {
            if (referencedElsewhere.Contains(pair.OpenXmlPart))
            {
                continue;
            }

            PruneUnreferencedParts(pair.OpenXmlPart, referencedElsewhere, visited);
            container.DeletePart(pair.OpenXmlPart);
        }
    }

    private static uint GetNextSlideId(SlideIdList slideIdList)
    {
        var slideIds = slideIdList.ChildElements
            .OfType<SlideId>()
            .Select(sid => sid.Id?.Value ?? 0)
            .ToList();
        var maxSlideId = slideIds.Count > 0 ? slideIds.Max() : 256;
        return maxSlideId + 1;
    }

    private void LoadExistingSlides()
    {
        var presentation = _document.PresentationPart!.Presentation!;
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
