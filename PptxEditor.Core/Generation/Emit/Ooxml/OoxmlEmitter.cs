using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;
using Drawing = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// OOXML emitter (plan.md §2): turns the absolute draw tree (<see cref="LayoutResult"/>)
/// into a native .pptx. Every element is emitted at its resolved absolute geometry —
/// no layout happens here (rule 1), and no preview tricks enter the PPTX (rule 4).
/// Units: points → EMU (×12700); spcPct-family values (gradient stops, alpha, adj,
/// srcRect) are 1/1000ths of a percent (AGENTS.pptx.md rule 2). Autofit is never emitted:
/// the shrink pass (P3) resolves <see cref="ResolvedText.FontScale"/> and run sizes are
/// scaled literally, so no renderer-dependent normAutofit is needed. Clip overflow is a
/// no-op in OOXML — an spTree does not clip its children (PowerPoint behavior).
/// </summary>
public sealed class OoxmlEmitter
{
    private const double EmuPerPoint = 12700.0;
    private const int Pct1000Scale = 100000; // spcPct family: 100000 = 100% (rule 2)

    private readonly List<string> _warnings = new();
    private uint _nextShapeId;

    /// <summary>Emits the whole resolved document as a .pptx package.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="layout"/> is null.</exception>
    /// <exception cref="ArgumentException">The document has no slides.</exception>
    public OoxmlEmissionResult Emit(LayoutResult layout)
    {
        ArgumentNullException.ThrowIfNull(layout);
        if (layout.Slides.Count == 0)
        {
            throw new ArgumentException("A deck must contain at least one slide.", nameof(layout));
        }

        _warnings.Clear();
        var stream = new MemoryStream();
        using (var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new Presentation
            {
                SlideMasterIdList = new SlideMasterIdList(),
                SlideIdList = new SlideIdList()
            };

            var slideMasterPart = AddSlideMaster(presentationPart);
            var slideLayoutPart = AddSlideLayout(slideMasterPart);
            AddTheme(slideMasterPart);

            var first = layout.Slides[0];
            presentationPart.Presentation.SlideSize = new P.SlideSize
            {
                Cx = (int)PtToEmu(first.WidthPt),
                Cy = (int)PtToEmu(first.HeightPt)
            };
            presentationPart.Presentation.NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 };

            uint slideId = 256;
            foreach (var slide in layout.Slides)
            {
                EmitSlide(presentationPart, slideLayoutPart, slide, slideId++);
            }

            document.PackageProperties.Creator = "OfficeEditor";
            document.PackageProperties.Created = DateTime.Now;
            document.PackageProperties.Modified = DateTime.Now;
            document.Save();
        }

        return new OoxmlEmissionResult { Bytes = stream.ToArray(), Warnings = [.. _warnings] };
    }

    #region Package scaffolding (minimal master/layout/theme, modeled on PresentationBuilder's proven init)

    private static SlideMasterPart AddSlideMaster(PresentationPart presentationPart)
    {
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(CreateEmptyShapeTree()),
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
            new SlideLayoutIdList());

        presentationPart.Presentation!.SlideMasterIdList!.Append(new SlideMasterId
        {
            Id = 2147483648,
            RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
        });
        return slideMasterPart;
    }

    private static SlideLayoutPart AddSlideLayout(SlideMasterPart slideMasterPart)
    {
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = new P.SlideLayout(
            new CommonSlideData(CreateEmptyShapeTree()) { Name = "Blank" })
        {
            Type = SlideLayoutValues.Blank
        };
        slideLayoutPart.AddPart(slideMasterPart);

        slideMasterPart.SlideMaster!.SlideLayoutIdList!.Append(new SlideLayoutId
        {
            Id = 2147483649,
            RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
        });
        return slideLayoutPart;
    }

    private static void AddTheme(SlideMasterPart slideMasterPart)
    {
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
    }

    private static ShapeTree CreateEmptyShapeTree() =>
        new(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1, Name = "" },
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
        );

    #endregion

    #region Slide + element dispatch

    private void EmitSlide(PresentationPart presentationPart, SlideLayoutPart slideLayoutPart, ResolvedSlide slide, uint slideId)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        var shapeTree = CreateEmptyShapeTree();
        slidePart.Slide = new Slide(new CommonSlideData(shapeTree));
        slidePart.AddPart(slideLayoutPart);

        _nextShapeId = 2;
        EmitContainer(slidePart, shapeTree, slide.Root);

        presentationPart.Presentation!.SlideIdList!.Append(new SlideId
        {
            Id = slideId,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        });
    }

    private void EmitElement(SlidePart slidePart, ShapeTree shapeTree, ResolvedElement element)
    {
        switch (element)
        {
            case ResolvedContainer container:
                EmitContainer(slidePart, shapeTree, container);
                break;
            case ResolvedText text:
                EmitText(shapeTree, text);
                break;
            case ResolvedRect rect:
                EmitRect(shapeTree, rect);
                break;
            case ResolvedEllipse ellipse:
                EmitEllipse(shapeTree, ellipse);
                break;
            case ResolvedLine line:
                EmitLine(shapeTree, line);
                break;
            case ResolvedImage image:
                EmitImage(slidePart, shapeTree, image);
                break;
            case ResolvedGroup group:
                // Children are already placed in absolute coordinates; a group is a
                // paint-order construct, so it flattens into the spTree in document order.
                foreach (var child in group.Children)
                {
                    EmitElement(slidePart, shapeTree, child);
                }
                break;
            default:
                throw new NotSupportedException($"Resolved element '{element.GetType().Name}' is not supported by the OOXML emitter.");
        }
    }

    private void EmitContainer(SlidePart slidePart, ShapeTree shapeTree, ResolvedContainer container)
    {
        // The container's own surface paints first (below its children = document order).
        if (container.Fill is not null || container.Stroke is not null || container.Shadow is not null)
        {
            var geometry = RadiusGeometryMapper.Map(container.Radius, container.Width, container.Height);
            var shape = new P.Shape(
                CreateNonVisualShapeProperties("Container"),
                CreateShapeProperties(container, geometry, container.Fill, container.Stroke, container.Shadow));
            shapeTree.Append(shape);
        }

        foreach (var child in container.Children)
        {
            EmitElement(slidePart, shapeTree, child);
        }
    }

    private void EmitRect(ShapeTree shapeTree, ResolvedRect rect)
    {
        var geometry = RadiusGeometryMapper.Map(rect.Radius, rect.Width, rect.Height);
        var shape = new P.Shape(
            CreateNonVisualShapeProperties("Rectangle"),
            CreateShapeProperties(rect, geometry, rect.Fill, rect.Stroke, rect.Shadow));
        shapeTree.Append(shape);
    }

    private void EmitEllipse(ShapeTree shapeTree, ResolvedEllipse ellipse)
    {
        var shape = new P.Shape(
            CreateNonVisualShapeProperties("Ellipse"),
            CreateShapeProperties(ellipse, null, ellipse.Fill, ellipse.Stroke, ellipse.Shadow, Drawing.ShapeTypeValues.Ellipse));
        shapeTree.Append(shape);
    }

    private void EmitLine(ShapeTree shapeTree, ResolvedLine line)
    {
        // Straight lines are centered in their box with a zero-size cross axis, the same
        // way PowerPoint persists a perfectly horizontal/vertical line (ext cy/cx = 0).
        var (x, y, cx, cy) = line.Orientation == LineOrientation.Horizontal
            ? (line.X, line.Y + line.Height / 2.0, line.Width, 0.0)
            : (line.X + line.Width / 2.0, line.Y, 0.0, line.Height);

        var stroke = line.Stroke ?? new StrokeSpec { Color = "#000000", WidthPt = 1 };

        var shapeProperties = new P.ShapeProperties(
            new Drawing.Transform2D(
                new Drawing.Offset { X = PtToEmu(x), Y = PtToEmu(y) },
                new Drawing.Extents { Cx = PtToEmu(cx), Cy = PtToEmu(cy) }),
            new Drawing.PresetGeometry(new Drawing.AdjustValueList())
            {
                Preset = line.IsConnector ? Drawing.ShapeTypeValues.StraightConnector1 : Drawing.ShapeTypeValues.Line
            });
        AppendOutline(shapeProperties, stroke);

        if (line.IsConnector)
        {
            shapeTree.Append(new P.ConnectionShape(
                new P.NonVisualConnectionShapeProperties(
                    new P.NonVisualDrawingProperties { Id = NextShapeId(), Name = $"Connector {_nextShapeId - 1}" },
                    new P.NonVisualConnectorShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                shapeProperties));
        }
        else
        {
            shapeTree.Append(new P.Shape(
                CreateNonVisualShapeProperties("Line"),
                shapeProperties));
        }
    }

    private void EmitText(ShapeTree shapeTree, ResolvedText text)
    {
        var bodyProperties = new Drawing.BodyProperties
        {
            LeftInset = (int)PtToEmu(text.Insets.Left),
            TopInset = (int)PtToEmu(text.Insets.Top),
            RightInset = (int)PtToEmu(text.Insets.Right),
            BottomInset = (int)PtToEmu(text.Insets.Bottom),
            Anchor = text.Anchor switch
            {
                TextAnchor.Top => Drawing.TextAnchoringTypeValues.Top,
                TextAnchor.Middle => Drawing.TextAnchoringTypeValues.Center,
                TextAnchor.Bottom => Drawing.TextAnchoringTypeValues.Bottom,
                _ => throw new NotSupportedException($"Text anchor '{text.Anchor}' is not supported.")
            }
        };

        var paragraph = new Drawing.Paragraph
        {
            ParagraphProperties = new Drawing.ParagraphProperties
            {
                Alignment = text.TextAlign switch
                {
                    Model.TextAlign.Left => Drawing.TextAlignmentTypeValues.Left,
                    Model.TextAlign.Center => Drawing.TextAlignmentTypeValues.Center,
                    Model.TextAlign.Right => Drawing.TextAlignmentTypeValues.Right,
                    _ => throw new NotSupportedException($"Text align '{text.TextAlign}' is not supported.")
                }
            }
        };
        foreach (var run in text.Runs)
        {
            AppendRun(paragraph, run, text.FontScale);
        }

        // FontScale is applied literally to run sizes (P3 already resolved the shrink);
        // no normAutofit is emitted, so nothing is renderer-dependent (rule 4).
        var shapeProperties = new P.ShapeProperties(
            CreateTransform(text),
            new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle },
            new Drawing.NoFill());
        AppendOutline(shapeProperties, null);
        AppendShadow(shapeProperties, text.Shadow);

        shapeTree.Append(new P.Shape(
            CreateNonVisualShapeProperties("Text"),
            shapeProperties,
            new P.TextBody(bodyProperties, paragraph)));
    }

    private void AppendRun(Drawing.Paragraph paragraph, ResolvedTextRun run, double fontScale)
    {
        var runProperties = new Drawing.RunProperties
        {
            FontSize = (int)Math.Round(run.FontSizePt * fontScale * 100.0, MidpointRounding.AwayFromZero)
        };
        if (run.Bold)
        {
            runProperties.Bold = true;
        }
        if (run.Italic)
        {
            runProperties.Italic = true;
        }
        if (run.ColorHex is not null)
        {
            runProperties.Append(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = Hex(run.ColorHex) }));
        }
        if (run.FontFamily is not null)
        {
            runProperties.Append(new Drawing.LatinFont { Typeface = run.FontFamily });
        }

        // Line breaks inside a run become a:br elements carrying the same run properties.
        var segments = run.Text.Split('\n');
        for (var i = 0; i < segments.Length; i++)
        {
            if (i > 0)
            {
                paragraph.Append(new Drawing.Break((Drawing.RunProperties)runProperties.CloneNode(true)));
            }
            var text = new Drawing.Text(segments[i]);
            if (segments[i] != segments[i].Trim())
            {
                text.SetAttribute(new OpenXmlAttribute("xml:space", "http://www.w3.org/XML/1998/namespace", "preserve"));
            }
            paragraph.Append(new Drawing.Run((Drawing.RunProperties)runProperties.CloneNode(true), text));
        }
    }

    private void EmitImage(SlidePart slidePart, ShapeTree shapeTree, ResolvedImage image)
    {
        var (bytes, extension) = ResolveImageSource(image.Source);
        var imagePart = slidePart.AddImagePart(GetImagePartType(extension));
        imagePart.FeedData(new MemoryStream(bytes));

        var x = PtToEmu(image.X);
        var y = PtToEmu(image.Y);
        var cx = PtToEmu(image.Width);
        var cy = PtToEmu(image.Height);

        SourceRect? sourceRect = null;
        // Crop without an explicit rect is Fill (F7 semantics).
        var fit = image.Fit == ImageFitMode.Crop && image.Crop is null ? ImageFitMode.Fill : image.Fit;
        switch (fit)
        {
            case ImageFitMode.Crop:
                sourceRect = image.Crop!.Value;
                break;
            case ImageFitMode.Fill:
            case ImageFitMode.Contain:
                var dimensions = ImageHeaderSniffer.TryGetPixelDimensions(bytes);
                if (dimensions is not { } dims)
                {
                    _warnings.Add(
                        $"Image dimensions could not be determined (unsupported or unrecognized format); " +
                        $"fit '{fit.ToString().ToLowerInvariant()}' fell back to 'stretch'.");
                    break;
                }
                if (fit == ImageFitMode.Fill)
                {
                    sourceRect = ImageFitGeometry.ComputeFillCrop(dims.Width, dims.Height, cx, cy);
                }
                else
                {
                    (x, y, cx, cy) = ImageFitGeometry.ComputeContainFrame(dims.Width, dims.Height, x, y, cx, cy);
                }
                break;
            default:
                throw new NotSupportedException($"Image fit mode '{image.Fit}' is not supported by the OOXML emitter.");
        }

        var blipFill = new P.BlipFill(new Drawing.Blip { Embed = slidePart.GetIdOfPart(imagePart) });
        if (sourceRect is { } rect)
        {
            // CT_BlipFillProperties sequence: blip?, srcRect?, (tile|stretch)?.
            blipFill.Append(new Drawing.SourceRectangle
            {
                Left = rect.Left,
                Top = rect.Top,
                Right = rect.Right,
                Bottom = rect.Bottom
            });
        }
        blipFill.Append(new Drawing.Stretch(new Drawing.FillRectangle()));

        var id = NextShapeId();
        var nonVisual = new P.NonVisualDrawingProperties { Id = id, Name = $"Image {id}" };
        if (image.Alt is not null)
        {
            nonVisual.Description = image.Alt;
        }

        shapeTree.Append(new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisual,
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            blipFill,
            new P.ShapeProperties(
                new Drawing.Transform2D(
                    new Drawing.Offset { X = x, Y = y },
                    new Drawing.Extents { Cx = cx, Cy = cy }),
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle })));
    }

    #endregion

    #region Shared shape plumbing

    private P.NonVisualShapeProperties CreateNonVisualShapeProperties(string kind)
    {
        var id = NextShapeId();
        return new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = $"{kind} {id}" },
            new P.NonVisualShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties());
    }

    private uint NextShapeId() => _nextShapeId++;

    private static Drawing.Transform2D CreateTransform(ResolvedElement element) =>
        new(
            new Drawing.Offset { X = PtToEmu(element.X), Y = PtToEmu(element.Y) },
            new Drawing.Extents { Cx = PtToEmu(element.Width), Cy = PtToEmu(element.Height) });

    private P.ShapeProperties CreateShapeProperties(
        ResolvedElement element,
        RadiusGeometryMapper.RadiusGeometry? geometry,
        FillSpec? fill,
        StrokeSpec? stroke,
        ShadowSpec? shadow,
        Drawing.ShapeTypeValues? presetOverride = null)
    {
        // CT_ShapeProperties sequence: xfrm, geometry, fill, ln, effectLst.
        var shapeProperties = new P.ShapeProperties(CreateTransform(element));
        shapeProperties.Append(new Drawing.PresetGeometry(BuildAdjustValueList(geometry))
        {
            Preset = presetOverride ?? geometry?.Preset ?? Drawing.ShapeTypeValues.Rectangle
        });
        AppendFill(shapeProperties, fill);
        AppendOutline(shapeProperties, stroke);
        AppendShadow(shapeProperties, shadow);
        return shapeProperties;
    }

    private static Drawing.AdjustValueList BuildAdjustValueList(RadiusGeometryMapper.RadiusGeometry? geometry)
    {
        var adjustValues = new Drawing.AdjustValueList();
        if (geometry is not null)
        {
            adjustValues.Append(new Drawing.ShapeGuide { Name = "adj1", Formula = $"val {geometry.Adj1}" });
            if (geometry.Preset != Drawing.ShapeTypeValues.RoundRectangle)
            {
                adjustValues.Append(new Drawing.ShapeGuide { Name = "adj2", Formula = $"val {geometry.Adj2}" });
            }
        }
        return adjustValues;
    }

    private static void AppendFill(P.ShapeProperties shapeProperties, FillSpec? fill)
    {
        switch (fill)
        {
            case null:
                shapeProperties.Append(new Drawing.NoFill());
                break;
            case SolidFill solid:
                shapeProperties.Append(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = Hex(solid.Color) }));
                break;
            case LinearGradientFill gradient:
                shapeProperties.Append(BuildGradientFill(gradient));
                break;
            default:
                throw new NotSupportedException($"Fill type '{fill.GetType().Name}' is not supported by the OOXML emitter.");
        }
    }

    private static Drawing.GradientFill BuildGradientFill(LinearGradientFill gradient)
    {
        var stopList = new Drawing.GradientStopList();
        foreach (var stop in gradient.Stops.OrderBy(s => s.Offset))
        {
            var color = new Drawing.RgbColorModelHex { Val = Hex(stop.Color) };
            if (stop.Alpha is { } alpha)
            {
                color.Append(new Drawing.Alpha { Val = FractionToPct1000(alpha) });
            }
            stopList.Append(new Drawing.GradientStop(color) { Position = FractionToPct1000(stop.Offset) });
        }
        return new Drawing.GradientFill(
            stopList,
            new Drawing.LinearGradientFill { Angle = NormalizeAngle(gradient.Angle), Scaled = true });
    }

    private static void AppendOutline(P.ShapeProperties shapeProperties, StrokeSpec? stroke)
    {
        if (stroke is null)
        {
            // Explicit noFill keeps the outline theme-independent (no p:style is emitted).
            shapeProperties.Append(new Drawing.Outline(new Drawing.NoFill()));
            return;
        }
        shapeProperties.Append(new Drawing.Outline(
            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = Hex(stroke.Color) }))
        {
            Width = (int)PtToEmu(stroke.WidthPt)
        });
    }

    private static void AppendShadow(P.ShapeProperties shapeProperties, ShadowSpec? shadow)
    {
        if (shadow is null)
        {
            return;
        }
        var color = new Drawing.RgbColorModelHex { Val = Hex(shadow.Color) };
        if (shadow.Alpha is { } alpha)
        {
            color.Append(new Drawing.Alpha { Val = FractionToPct1000(alpha) });
        }
        shapeProperties.Append(new Drawing.EffectList(
            new Drawing.OuterShadow(color)
            {
                BlurRadius = PtToEmu(shadow.Blur),
                Distance = PtToEmu(Math.Sqrt(shadow.Dx * shadow.Dx + shadow.Dy * shadow.Dy)),
                Direction = NormalizeAngle(Math.Atan2(shadow.Dy, shadow.Dx) * 180.0 / Math.PI)
            }));
    }

    #endregion

    #region Image sources + units

    private static (byte[] Bytes, string Extension) ResolveImageSource(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            throw new ArgumentException("Image source is required.", nameof(source));
        }

        if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            // data:<mime>;base64,<payload>
            var comma = source.IndexOf(',');
            if (comma < 0 || !source[..comma].EndsWith(";base64", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    "Image data URIs must be of the form 'data:<mime>;base64,<payload>'.", nameof(source));
            }
            var mime = source["data:".Length..(comma - ";base64".Length)];
            return (Convert.FromBase64String(source[(comma + 1)..]), MimeToExtension(mime));
        }

        if (source.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || source.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                "Remote image URLs are not supported in v1; pass a local file path or a data URI.", nameof(source));
        }

        if (!File.Exists(source))
        {
            throw new FileNotFoundException($"Image not found: {source}", source);
        }
        return (File.ReadAllBytes(source), Path.GetExtension(source).ToLowerInvariant());
    }

    private static string MimeToExtension(string mime) => mime.Trim().ToLowerInvariant() switch
    {
        "image/png" => ".png",
        "image/jpeg" => ".jpg",
        "image/gif" => ".gif",
        "image/bmp" => ".bmp",
        "image/tiff" => ".tiff",
        "image/svg+xml" => ".svg",
        var other => throw new ArgumentException($"Unsupported image media type '{other}'.", nameof(mime))
    };

    // Extension → part-type switch mirrored from SlideBuilder.AddImage / PptxElementReplacer
    // (those files are owned by other workstreams and must not be edited).
    private static PartTypeInfo GetImagePartType(string extension) => extension.ToLowerInvariant() switch
    {
        ".png" => ImagePartType.Png,
        ".gif" => ImagePartType.Gif,
        ".bmp" => ImagePartType.Bmp,
        ".tiff" or ".tif" => ImagePartType.Tiff,
        ".svg" => ImagePartType.Svg,
        _ => ImagePartType.Jpeg
    };

    private static long PtToEmu(double points) =>
        (long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero);

    private static int FractionToPct1000(double fraction) =>
        (int)Math.Round(fraction * Pct1000Scale, MidpointRounding.AwayFromZero);

    /// <summary>OOXML angles are 60000ths of a degree, normalized to [0, 21600000).</summary>
    private static int NormalizeAngle(double degrees)
    {
        var angle = (int)Math.Round(degrees * 60000.0, MidpointRounding.AwayFromZero) % 21600000;
        return angle < 0 ? angle + 21600000 : angle;
    }

    private static string Hex(string resolvedColor)
    {
        var hex = resolvedColor.StartsWith('#') ? resolvedColor[1..] : resolvedColor;
        return hex.ToUpperInvariant();
    }

    #endregion
}
