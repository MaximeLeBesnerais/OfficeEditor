using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Model;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Positioned-tier OOXML emitter: turns the generation model's <see cref="PositionedElement"/>
/// hierarchy into WordprocessingDrawing floating anchors (<c>wp:anchor</c>) appended to a
/// document/header/footer container. This is the write-side counterpart of the
/// <c>DocxToTypstConverter</c> read side (the semantics reference in the generation README):
/// every primitive lands as a real floating object that Word, LibreOffice and the converter
/// can read back.
///
/// Supported primitives: text boxes and callouts (<c>wps:wsp</c> + <c>w:txbxContent</c>),
/// rectangles, straight lines, and floating pictures (<c>p:pic</c>) whose relationship id is
/// supplied through the <see cref="IPositionedImageResolver"/> seam — this emitter never loads
/// asset bytes. Wrap modes square/tight/through/topAndBottom map to their <c>wp:wrap*</c>
/// elements; none/behindText/inFrontOfText map to <c>wp:wrapNone</c> with the anchor
/// <c>behindDoc</c>/z-order semantics. Unsupported or approximated combinations always
/// surface as <see cref="PositionedElementEmitWarning"/> — never silent omission.
///
/// Determinism: elements are emitted in ascending z-order (stable within equal z), later
/// elements paint on top, and every drawing id is unique within the owning part.
///
/// SDK caveat: the OpenXML SDK ships no typed classes for the <c>wps</c> vocabulary and its
/// Office2010 schema for it is incorrect (it models <c>wps:cNvSpPr</c> as containing only
/// <c>a:spLocks</c>, wrongly requires <c>bodyPr</c> on textless shapes, and validates
/// <c>w:txbxContent</c> paragraphs against a reduced model), and its <c>pic</c> schema omits
/// <c>pic:nvPr</c>. Consequently <c>OpenXmlValidator</c> reports spurious errors on output
/// this emitter produces — the XML itself matches real Word output (verified against the
/// <c>DocxToTypstConverter</c> read side and LibreOffice).
/// </summary>
public sealed class PositionedElementEmitter
{
    private const double EmuPerPoint = 12700.0;

    /// <summary>Word's default anchor <c>relativeHeight</c> (0x0F000000).</summary>
    private const uint RelativeHeightBase = 251658240;

    /// <summary>Coordinate units of a <c>wp:wrapPolygon</c> (21600 per 100% of the box).</summary>
    private const long WrapPolygonScale = 21600;

    private readonly PositionedElementEmitOptions _options;
    private readonly DesignTokenResolver _design;
    private readonly TextBoxContentFactory _textBoxes;

    /// <summary>
    /// Creates the emitter. <see cref="PositionedElementEmitOptions"/> is optional; omit it
    /// for built-in fallbacks (no design tokens, no image resolver).
    /// </summary>
    public PositionedElementEmitter(PositionedElementEmitOptions? options = null)
    {
        _options = options ?? new PositionedElementEmitOptions();
        _design = new DesignTokenResolver(_options.Design);
        _textBoxes = new TextBoxContentFactory(_design);
    }

    /// <summary>
    /// Emits <paramref name="elements"/> as floating anchors appended to
    /// <paramref name="container"/> (a <see cref="Body"/>, header, footer or text-box
    /// content). <paramref name="owningPart"/> hosts image relationships and existing
    /// drawing ids; it is optional except when emitting positioned images through the
    /// resolver seam. Elements are z-ordered deterministically before emission.
    /// </summary>
    public PositionedElementEmitResult Emit(OpenXmlPart? owningPart, OpenXmlCompositeElement container, IEnumerable<PositionedElement> elements)
    {
        ArgumentNullException.ThrowIfNull(container);
        ArgumentNullException.ThrowIfNull(elements);

        EmitContext context = new(owningPart, container, _options, _design, _textBoxes);
        List<PositionedElement> ordered = [.. elements.OrderBy(e => e.Position.ZOrder)];
        for (int i = 0; i < ordered.Count; i++)
        {
            context.EmitElement(i, ordered[i]);
        }

        return context.Result;
    }

    /// <summary>Emits a single positioned element as a floating anchor.</summary>
    public PositionedElementEmitResult Emit(OpenXmlPart? owningPart, OpenXmlCompositeElement container, PositionedElement element)
    {
        ArgumentNullException.ThrowIfNull(element);
        return Emit(owningPart, container, new[] { element });
    }

    /// <summary>
    /// Per-call emit state: the owning part + target container, the monotonically unique
    /// drawing ids and z-heights, and the collected warnings. Created fresh per
    /// <see cref="Emit"/> call so a shared emitter is deterministic and reuse-safe.
    /// </summary>
    private sealed class EmitContext
    {
        private readonly List<PositionedElementEmitWarning> _warnings = [];
        private readonly OpenXmlPart? _owningPart;
        private readonly OpenXmlCompositeElement _container;
        private readonly PositionedElementEmitOptions _options;
        private readonly DesignTokenResolver _design;
        private readonly TextBoxContentFactory _textBoxes;
        private uint _nextDrawingId;
        private uint _emitted;
        private int _skipped;

        public EmitContext(
            OpenXmlPart? owningPart,
            OpenXmlCompositeElement container,
            PositionedElementEmitOptions options,
            DesignTokenResolver design,
            TextBoxContentFactory textBoxes)
        {
            _owningPart = owningPart;
            _container = container;
            _options = options;
            _design = design;
            _textBoxes = textBoxes;
            _nextDrawingId = ComputeNextDrawingId();
        }

        public OpenXmlPart? OwningPart => _owningPart;

        public PositionedElementEmitOptions Options => _options;

        public PositionedElementEmitResult Result => new()
        {
            EmittedCount = (int)_emitted,
            SkippedCount = _skipped,
            Warnings = [.. _warnings]
        };

        public uint NextDrawingId() => _nextDrawingId++;

        /// <summary>Monotonic anchor z-height: higher values paint on top, matching the emit (z) order.</summary>
        public uint NextRelativeHeight() => RelativeHeightBase + _emitted;

        public void RecordEmitted() => _emitted++;

        public void RecordSkipped() => _skipped++;

        public void Warn(int index, string elementType, string message)
            => _warnings.Add(new PositionedElementEmitWarning(index, elementType, message));

        public void EmitElement(int index, PositionedElement element)
        {
            switch (element)
            {
                case PositionedTextBox textBox:
                    EmitTextBox(this, index, textBox);
                    break;
                case PositionedRectangle rectangle:
                    EmitRectangle(this, index, rectangle);
                    break;
                case PositionedLine line:
                    EmitLine(this, index, line);
                    break;
                case PositionedImage image:
                    EmitImage(this, index, image);
                    break;
                case PositionedCallout callout:
                    EmitCallout(this, index, callout);
                    break;
                default:
                    throw new PositionedElementEmitException(
                        $"Positioned element '{element.GetType().Name}' is not supported by the positioned OOXML emitter.");
            }
        }

        private uint ComputeNextDrawingId()
        {
            if (_options.DrawingIdBase is { } explicitBase)
            {
                return explicitBase;
            }

            long maxId = 0;
            OpenXmlElement? scope = _owningPart?.RootElement ?? _container;
            if (scope is not null)
            {
                // Match by local name + id attribute rather than the typed class: the SDK
                // re-parses elements inside a wps:wsp (which has no typed class) as unknown
                // elements, so Descendants<NonVisualDrawingProperties> would miss their ids.
                foreach (OpenXmlElement element in scope.Descendants())
                {
                    if (element.LocalName != "cNvPr")
                    {
                        continue;
                    }

                    OpenXmlAttribute idAttribute = element.GetAttribute("id", string.Empty);
                    if (uint.TryParse(idAttribute.Value, System.Globalization.CultureInfo.InvariantCulture, out uint id))
                    {
                        maxId = Math.Max(maxId, id);
                    }
                }
            }

            return (uint)(maxId + 1);
        }

        private void EmitTextBox(EmitContext context, int index, PositionedTextBox element)
        {
            PositionSpec position = element.Position;
            ValidateBox(position, "textBox");

            string? fillHex = _design.ResolveBoxFillHex(element.Fill);
            (string? strokeHex, double strokeWidthPt) = ResolveStroke(_design.ResolveBoxStroke(element.Stroke));
            double cornerRadiusPt = element.CornerRadiusPt ?? _design.DefaultCornerRadiusPt;
            W.TextBoxContent content = _textBoxes.Build(element.Content);

            uint id = NextDrawingId();
            OpenXmlUnknownElement shape = WordprocessingShapeBuilder.BuildTextBox(
                id, $"TextBox {id}", position.Alt, BoxWidthPt(position), BoxHeightPt(position),
                position.Rotation, cornerRadiusPt, fillHex, strokeHex, strokeWidthPt, content);

            AppendAnchor(context, index, "textBox", "TextBox", position, BoxWidthPt(position), BoxHeightPt(position), shape);
        }

        private void EmitRectangle(EmitContext context, int index, PositionedRectangle element)
        {
            PositionSpec position = element.Position;
            ValidateBox(position, "rect");

            string? fillHex = _design.ResolveBoxFillHex(element.Fill);
            (string? strokeHex, double strokeWidthPt) = ResolveStroke(_design.ResolveBoxStroke(element.Stroke));
            double cornerRadiusPt = element.CornerRadiusPt ?? _design.DefaultCornerRadiusPt;

            uint id = NextDrawingId();
            OpenXmlUnknownElement shape = WordprocessingShapeBuilder.BuildRectangle(
                id, $"Rectangle {id}", position.Alt, BoxWidthPt(position), BoxHeightPt(position),
                position.Rotation, cornerRadiusPt, fillHex, strokeHex, strokeWidthPt);

            AppendAnchor(context, index, "rect", "Rectangle", position, BoxWidthPt(position), BoxHeightPt(position), shape);
        }

        private void EmitLine(EmitContext context, int index, PositionedLine element)
        {
            PositionSpec position = element.Position;
            if (position.WidthPt is not { } widthOrNull || widthOrNull <= 0)
            {
                throw new PositionedElementEmitException($"A positioned line requires a positive 'width' (its length); got {position.WidthPt}.");
            }

            double lengthPt = widthOrNull;
            bool vertical = element.Orientation == LineOrientation.Vertical;
            StrokeSpec stroke = _design.ResolveLineStroke(element.Stroke);
            (string? strokeHex, double strokeWidthPt) = ResolveStroke(stroke);

            uint id = NextDrawingId();
            OpenXmlUnknownElement shape = WordprocessingShapeBuilder.BuildLine(
                id, $"Line {id}", position.Alt, lengthPt, vertical, strokeHex!, strokeWidthPt);

            // The anchor extents mirror the shape's zero-cross-axis box: (length, 0) for a
            // horizontal line, (0, length) for a vertical one.
            double cxPt = vertical ? 0 : lengthPt;
            double cyPt = vertical ? lengthPt : 0;
            AppendAnchor(context, index, "line", "Line", position, cxPt, cyPt, shape);
        }

        private void EmitCallout(EmitContext context, int index, PositionedCallout element)
        {
            PositionSpec position = element.Position;
            ValidateBox(position, "callout");

            CalloutAppearance appearance = CalloutAppearance.For(element.Tone);
            string fillHex = element.Fill is null ? appearance.FillHex : _design.ResolveHex(element.Fill, appearance.FillHex);
            StrokeSpec? stroke = element.Stroke;
            string strokeHex = _design.ResolveHex(stroke?.Color ?? appearance.StrokeHex, appearance.StrokeHex);
            double strokeWidthPt = stroke?.WidthPt ?? 1;
            double cornerRadiusPt = element.CornerRadiusPt ?? _design.DefaultCornerRadiusPt;
            W.TextBoxContent content = _textBoxes.Build(element.Content);

            uint id = NextDrawingId();
            OpenXmlUnknownElement shape = WordprocessingShapeBuilder.BuildTextBox(
                id, $"Callout {id}", position.Alt, BoxWidthPt(position), BoxHeightPt(position),
                position.Rotation, cornerRadiusPt, fillHex, strokeHex, strokeWidthPt, content);

            AppendAnchor(context, index, "callout", "Callout", position, BoxWidthPt(position), BoxHeightPt(position), shape);
        }

        private void EmitImage(EmitContext context, int index, PositionedImage element)
        {
            PositionSpec position = element.Position;
            ValidateBox(position, "image");

            if (_options.ImageResolver is not { } resolver)
            {
                context.Warn(index, "image",
                    "skipped: no image-part resolver is configured (PositionedElementEmitOptions.ImageResolver); " +
                    "the emitter never loads image assets itself.");
                RecordSkipped();
                return;
            }

            if (_owningPart is null)
            {
                context.Warn(index, "image",
                    "skipped: no owning part is available to host the image relationship.");
                RecordSkipped();
                return;
            }

            string? embedId = resolver.ResolveEmbedId(_owningPart, element);
            if (string.IsNullOrWhiteSpace(embedId))
            {
                context.Warn(index, "image",
                    $"skipped: the image-part manager returned no relationship id for source '{element.Source}'.");
                RecordSkipped();
                return;
            }

            double widthPt = BoxWidthPt(position);
            double heightPt = BoxHeightPt(position);
            (int Left, int Top, int Right, int Bottom)? srcRect = null;
            switch (element.Fit)
            {
                case ImageFitMode.Stretch:
                    break;
                case ImageFitMode.Crop when element.Crop is { } crop:
                    srcRect = (Percent(crop.Left), Percent(crop.Top), Percent(crop.Right), Percent(crop.Bottom));
                    break;
                case ImageFitMode.Crop:
                    break;
                case ImageFitMode.Fill:
                case ImageFitMode.Contain:
                    if (resolver.TryGetNaturalPixelSize(element.Source, out int naturalWidth, out int naturalHeight))
                    {
                        if (element.Fit == ImageFitMode.Fill)
                        {
                            srcRect = ComputeCoverCrop(naturalWidth, naturalHeight, widthPt, heightPt);
                        }
                        else
                        {
                            (double x, double y, widthPt, heightPt) = ComputeContainBox(position.X, position.Y, widthPt, heightPt, naturalWidth, naturalHeight);
                            position = position with { X = x, Y = y };
                        }
                    }
                    else
                    {
                        context.Warn(index, "image",
                            $"fit '{element.Fit.ToString().ToLowerInvariant()}' requires natural image dimensions, " +
                            "which the image-part manager does not provide; approximated as 'stretch'.");
                    }

                    break;
                default:
                    throw new PositionedElementEmitException($"Image fit mode '{element.Fit}' is not supported by the positioned OOXML emitter.");
            }

            uint id = NextDrawingId();
            A.Pictures.Picture picture = WordprocessingShapeBuilder.BuildPicture(
                id, $"Picture {id}", position.Alt, widthPt, heightPt, embedId, srcRect);

            AppendAnchor(context, index, "image", "Picture", position, widthPt, heightPt, picture);
        }

        private void AppendAnchor(
            EmitContext context,
            int index,
            string elementType,
            string shapeKind,
            PositionSpec position,
            double cxPt,
            double cyPt,
            OpenXmlElement payload)
        {
            AppendAnchorApproximationWarnings(context, index, elementType, position.Anchor);
            WrapDistances distances = position.WrapDistances ?? new WrapDistances();

            uint id = NextDrawingId();
            DW.DocProperties docProperties = new() { Id = id, Name = $"{shapeKind} {id}" };
            if (!string.IsNullOrWhiteSpace(position.Alt))
            {
                docProperties.Description = position.Alt;
            }

            DW.Anchor anchor = new()
            {
                SimplePos = false,
                RelativeHeight = NextRelativeHeight(),
                BehindDoc = position.Wrap == WrapMode.BehindText,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
                DistanceFromTop = Emu(distances.TopPt),
                DistanceFromBottom = Emu(distances.BottomPt),
                DistanceFromLeft = Emu(distances.LeftPt),
                DistanceFromRight = Emu(distances.RightPt)
            };
            anchor.Append(new DW.SimplePosition { X = 0, Y = 0 });
            anchor.Append(BuildHorizontalPosition(position));
            anchor.Append(BuildVerticalPosition(position));
            anchor.Append(new DW.Extent { Cx = PtToEmu(cxPt), Cy = PtToEmu(cyPt) });
            anchor.Append(new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 });
            anchor.Append(BuildWrap(context, index, elementType, position));
            anchor.Append(docProperties);
            anchor.Append(new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }));
            anchor.Append(new A.Graphic(new A.GraphicData(payload)
            {
                Uri = payload is A.Pictures.Picture
                    ? WordprocessingShapeBuilder.PictureNamespace
                    : WordprocessingShapeBuilder.WpsNamespace
            }));

            _container.Append(BuildAnchorParagraph(anchor));
            RecordEmitted();
        }

        /// <summary>
        /// A floating anchor lives in its own <c>w:p</c>. The run is sized to the smallest
        /// valid half-point so the anchor paragraph contributes a negligible line to the flow.
        /// </summary>
        private static W.Paragraph BuildAnchorParagraph(DW.Anchor anchor)
        {
            W.Run run = new();
            run.Append(new W.RunProperties(new W.FontSize { Val = "1" }));
            run.Append(new W.Drawing(anchor));
            return new W.Paragraph(run);
        }

        private static DW.HorizontalPosition BuildHorizontalPosition(PositionSpec position)
        {
            DW.HorizontalRelativePositionValues relativeFrom = position.Anchor switch
            {
                AnchorReference.Page => DW.HorizontalRelativePositionValues.Page,
                AnchorReference.Margin => DW.HorizontalRelativePositionValues.Margin,
                AnchorReference.Column => DW.HorizontalRelativePositionValues.Column,
                AnchorReference.Paragraph => DW.HorizontalRelativePositionValues.Character,
                AnchorReference.Character => DW.HorizontalRelativePositionValues.Character,
                _ => throw new PositionedElementEmitException($"Anchor reference '{position.Anchor}' is not supported.")
            };

            DW.HorizontalPosition horizontal = new() { RelativeFrom = relativeFrom };
            horizontal.Append(new DW.PositionOffset { Text = PtToEmu(position.X).ToString(System.Globalization.CultureInfo.InvariantCulture) });
            return horizontal;
        }

        private static DW.VerticalPosition BuildVerticalPosition(PositionSpec position)
        {
            DW.VerticalRelativePositionValues relativeFrom = position.Anchor switch
            {
                AnchorReference.Page => DW.VerticalRelativePositionValues.Page,
                AnchorReference.Margin => DW.VerticalRelativePositionValues.Margin,
                AnchorReference.Column => DW.VerticalRelativePositionValues.Margin,
                AnchorReference.Paragraph => DW.VerticalRelativePositionValues.Paragraph,
                AnchorReference.Character => DW.VerticalRelativePositionValues.Line,
                _ => throw new PositionedElementEmitException($"Anchor reference '{position.Anchor}' is not supported.")
            };

            DW.VerticalPosition vertical = new() { RelativeFrom = relativeFrom };
            vertical.Append(new DW.PositionOffset { Text = PtToEmu(position.Y).ToString(System.Globalization.CultureInfo.InvariantCulture) });
            return vertical;
        }

        private OpenXmlElement BuildWrap(EmitContext context, int index, string elementType, PositionSpec position)
        {
            WrapDistances distances = position.WrapDistances ?? new WrapDistances();
            switch (position.Wrap)
            {
                case WrapMode.None:
                case WrapMode.BehindText:
                case WrapMode.InFrontOfText:
                    // The anchor's distT/distB/distL/distR carry the wrap distances for these
                    // modes (there is no wp:wrap* element with attributes for them).
                    return new DW.WrapNone();

                case WrapMode.Square:
                    return new DW.WrapSquare
                    {
                        DistanceFromTop = Emu(distances.TopPt),
                        DistanceFromBottom = Emu(distances.BottomPt),
                        DistanceFromLeft = Emu(distances.LeftPt),
                        DistanceFromRight = Emu(distances.RightPt),
                        WrapText = DW.WrapTextValues.Largest
                    };

                case WrapMode.Tight:
                case WrapMode.Through:
                    if (distances.TopPt > 0 || distances.BottomPt > 0)
                    {
                        context.Warn(index, elementType,
                            $"wrap '{position.Wrap.ToString().ToLowerInvariant()}' only supports left/right wrap distances; " +
                            "top/bottom distances were dropped.");
                    }

                    DW.WrapPolygon polygon = BuildRectangleWrapPolygon();
                    if (position.Wrap == WrapMode.Tight)
                    {
                        return new DW.WrapTight
                        {
                            DistanceFromLeft = Emu(distances.LeftPt),
                            DistanceFromRight = Emu(distances.RightPt),
                            WrapText = DW.WrapTextValues.Largest,
                            WrapPolygon = polygon
                        };
                    }

                    return new DW.WrapThrough
                    {
                        DistanceFromLeft = Emu(distances.LeftPt),
                        DistanceFromRight = Emu(distances.RightPt),
                        WrapText = DW.WrapTextValues.Largest,
                        WrapPolygon = polygon
                    };

                case WrapMode.TopAndBottom:
                    if (distances.LeftPt > 0 || distances.RightPt > 0)
                    {
                        context.Warn(index, elementType,
                            "wrap 'topAndBottom' only supports top/bottom wrap distances; left/right distances were dropped.");
                    }

                    return new DW.WrapTopBottom
                    {
                        DistanceFromTop = Emu(distances.TopPt),
                        DistanceFromBottom = Emu(distances.BottomPt)
                    };

                default:
                    throw new PositionedElementEmitException($"Wrap mode '{position.Wrap}' is not supported by the positioned OOXML emitter.");
            }
        }

        /// <summary>
        /// Tight/through wrap needs a contour polygon. The vocabulary only carries box
        /// extents (no custom contour), so the contour is approximated as the bounding
        /// rectangle in wrap units — the same approximation Word makes for a rectangle.
        /// </summary>
        private static DW.WrapPolygon BuildRectangleWrapPolygon()
        {
            DW.WrapPolygon polygon = new()
            {
                Edited = false,
                StartPoint = new DW.StartPoint { X = 0, Y = 0 }
            };
            polygon.Append(new DW.LineTo { X = WrapPolygonScale, Y = 0 });
            polygon.Append(new DW.LineTo { X = WrapPolygonScale, Y = WrapPolygonScale });
            polygon.Append(new DW.LineTo { X = 0, Y = WrapPolygonScale });
            polygon.Append(new DW.LineTo { X = 0, Y = 0 });
            return polygon;
        }

        private void AppendAnchorApproximationWarnings(EmitContext context, int index, string elementType, AnchorReference anchor)
        {
            switch (anchor)
            {
                case AnchorReference.Column:
                    context.Warn(index, elementType,
                        "anchor 'column': WordprocessingDrawing has no vertical 'column' reference " +
                        "(vertical relative-from values are page/margin/paragraph/line); the vertical axis is approximated as relative to the margin.");
                    break;
                case AnchorReference.Paragraph:
                    context.Warn(index, elementType,
                        "anchor 'paragraph': WordprocessingDrawing has no horizontal 'paragraph' reference; " +
                        "the horizontal axis is approximated as relative to the character (vertical stays paragraph).");
                    break;
                case AnchorReference.Character:
                    context.Warn(index, elementType,
                        "anchor 'character': WordprocessingDrawing has no horizontal 'character' row reference; " +
                        "approximated as horizontal relative to the character and vertical relative to the line.");
                    break;
            }
        }

        private (string? Hex, double WidthPt) ResolveStroke(StrokeSpec? stroke)
        {
            if (stroke is null)
            {
                return (null, 1);
            }

            return (_design.ResolveHex(stroke.Color, "000000"), stroke.WidthPt);
        }

        private static void ValidateBox(PositionSpec position, string elementType)
        {
            if (position.WidthPt is not { } widthOrNull || widthOrNull <= 0)
            {
                throw new PositionedElementEmitException($"A positioned {elementType} requires a positive 'width'; got {position.WidthPt}.");
            }

            if (position.HeightPt is not { } heightOrNull || heightOrNull <= 0)
            {
                throw new PositionedElementEmitException($"A positioned {elementType} requires a positive 'height'; got {position.HeightPt}.");
            }
        }

        private static double BoxWidthPt(PositionSpec position) => position.WidthPt!.Value;

        private static double BoxHeightPt(PositionSpec position) => position.HeightPt!.Value;

        /// <summary>
        /// Cover-crop (<see cref="ImageFitMode.Fill"/>) srcRect: the source is center-cropped
        /// so its visible region matches the box aspect. Fractions are returned in
        /// 1/1000ths of a percent (100000 = 100%), the DrawingML srcRect convention. Returns
        /// null when the aspects already match (no crop needed).
        /// </summary>
        private static (int Left, int Top, int Right, int Bottom)? ComputeCoverCrop(int imageWidth, int imageHeight, double boxWidthPt, double boxHeightPt)
        {
            double imageAspect = (double)imageWidth / imageHeight;
            double boxAspect = boxWidthPt / boxHeightPt;
            if (Math.Abs(imageAspect - boxAspect) < 0.0001)
            {
                return null;
            }

            if (imageAspect > boxAspect)
            {
                double visibleFraction = boxAspect / imageAspect;
                int crop = Percent((1.0 - visibleFraction) / 2.0);
                return crop == 0 ? null : (crop, 0, crop, 0);
            }

            double verticalVisibleFraction = imageAspect / boxAspect;
            int verticalCrop = Percent((1.0 - verticalVisibleFraction) / 2.0);
            return verticalCrop == 0 ? null : (0, verticalCrop, 0, verticalCrop);
        }

        /// <summary>
        /// Contain (<see cref="ImageFitMode.Contain"/>) box: shrinks the box to the image
        /// aspect and re-centers it by shifting the anchor position by half the delta, so the
        /// picture is letterboxed inside the declared box.
        /// </summary>
        private static (double X, double Y, double Width, double Height) ComputeContainBox(
            double xPt, double yPt, double widthPt, double heightPt, int imageWidth, int imageHeight)
        {
            double imageAspect = (double)imageWidth / imageHeight;
            double boxAspect = widthPt / heightPt;
            if (Math.Abs(imageAspect - boxAspect) < 0.0001)
            {
                return (xPt, yPt, widthPt, heightPt);
            }

            if (imageAspect > boxAspect)
            {
                double newHeight = widthPt / imageAspect;
                return (xPt, yPt + (heightPt - newHeight) / 2.0, widthPt, newHeight);
            }

            double newWidth = heightPt * imageAspect;
            return (xPt + (widthPt - newWidth) / 2.0, yPt, newWidth, heightPt);
        }

        private static int Percent(double fraction) =>
            (int)Math.Round(fraction * 100000.0, MidpointRounding.AwayFromZero);

        private static uint Emu(double points) =>
            (uint)Math.Clamp((long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero), 0L, uint.MaxValue);

        private static long PtToEmu(double points) =>
            (long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero);
    }

    /// <summary>Tone-driven default fill/stroke for callouts (hex, resolved).</summary>
    private sealed record CalloutAppearance(string FillHex, string StrokeHex)
    {
        public static CalloutAppearance For(CalloutTone tone) => tone switch
        {
            CalloutTone.Note => new("EAF1FB", "2E74B5"),
            CalloutTone.Tip => new("E6F4EA", "2E7D32"),
            CalloutTone.Warning => new("FEF7E0", "B26B00"),
            CalloutTone.Error => new("FDECEA", "C62828"),
            _ => new("EAF1FB", "2E74B5")
        };
    }
}
