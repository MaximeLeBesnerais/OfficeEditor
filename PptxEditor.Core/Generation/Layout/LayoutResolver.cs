using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Layout;

/// <summary>
/// Pure layout engine (layout is resolved exactly once, in C#; size/overflow
/// semantics). Maps a container tree plus design tokens to an absolute draw tree where every
/// element has x/y/w/h in points and resolved fills/fonts. No OOXML or Typst imports.
/// <para>
/// Semantics summary:
/// grow shares the space remaining on the layout axis after fixed children and gaps;
/// aspect resolves against the dimension the parent constrains first (row → width,
/// column → height, grid → cell width, free canvas → whichever of w/h is fixed);
/// stretch fills the cross axis unless the child constrains it; undetermined cross size
/// fills the cross axis; paint order = document order.
/// </para>
/// </summary>
public sealed class LayoutResolver
{
    /// <summary>Smallest font scale accepted before a shrink warning is raised.</summary>
    public const double MinFontScale = 0.5;

    private const double Eps = 0.001;
    private const int Precision = 3;

    private readonly ITextMeasurer? _textMeasurer;
    private readonly List<string> _warnings = [];
    private DesignTokens _design = null!;

    /// <summary>Creates a resolver. <paramref name="textMeasurer"/> is the P3 seam; null = no text fit pass.</summary>
    public LayoutResolver(ITextMeasurer? textMeasurer = null)
    {
        _textMeasurer = textMeasurer;
    }

    /// <summary>Resolves every slide of <paramref name="document"/> to an absolute draw tree.</summary>
    public LayoutResult Resolve(GenerationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        _design = document.Design;
        _warnings.Clear();

        var slides = new List<ResolvedSlide>(document.Slides.Count);
        for (var i = 0; i < document.Slides.Count; i++)
        {
            slides.Add(ResolveSlide(document.Slides[i], document.SlideSize, $"slides[{i}]"));
        }
        return new LayoutResult { Slides = slides, Warnings = [.. _warnings] };
    }

    private ResolvedSlide ResolveSlide(ContainerElement slide, SlideSize size, string path)
    {
        if (slide.Size is not null)
        {
            throw new LayoutException(path, "the root slide container takes its size from the slide; 'size' is only valid on children.");
        }
        if (slide.At is not null)
        {
            throw new LayoutException(path, "the root slide container fills the slide; 'at' is only valid on children of layout-less parents.");
        }

        var root = (ResolvedContainer)ResolveElement(slide, new Rect(0, 0, size.WidthPt, size.HeightPt), path);
        return new ResolvedSlide
        {
            WidthPt = size.WidthPt,
            HeightPt = size.HeightPt,
            Root = root,
            Notes = slide.Notes,
            Id = slide.Id,
            Elements = Flatten(root)
        };
    }

    private ResolvedElement ResolveElement(GenElement element, Rect rect, string path)
    {
        rect = rect.Rounded();
        return element switch
        {
            ContainerElement c => ResolveContainer(c, rect, path),
            TextElement t => ResolveText(t, rect, path),
            RectElement r => new ResolvedRect
            {
                X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = r.Id,
                Fill = ResolveFill(r.Fill, path),
                Stroke = ResolveStroke(r.Stroke, path),
                Radius = r.Radius,
                Shadow = ResolveShadow(r.Shadow, path)
            },
            EllipseElement e => new ResolvedEllipse
            {
                X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = e.Id,
                Fill = ResolveFill(e.Fill, path),
                Stroke = ResolveStroke(e.Stroke, path),
                Shadow = ResolveShadow(e.Shadow, path)
            },
            LineElement l => new ResolvedLine
            {
                X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = l.Id,
                IsConnector = l.IsConnector,
                Orientation = l.Orientation,
                Stroke = ResolveStroke(l.Stroke, path)
            },
            ImageElement im => new ResolvedImage
            {
                X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = im.Id,
                Source = im.Source,
                Fit = im.Fit,
                Crop = im.Crop,
                Alt = im.Alt
            },
            GroupElement g => ResolveGroup(g, rect, path),
            ComponentElement c => throw new LayoutException(path,
                $"component '{c.Name}' must be expanded into primitives by the component layer before layout."),
            _ => throw new LayoutException(path, $"unsupported element type '{element.GetType().Name}'.")
        };
    }

    private ResolvedContainer ResolveContainer(ContainerElement container, Rect rect, string path)
    {
        var content = ApplyPadding(rect, container.Padding);
        IReadOnlyList<ResolvedElement> children;
        if (container.Layout is null)
        {
            children = ResolveFreeChildren(container.Children, content, path, container.Overflow);
        }
        else
        {
            children = container.Layout.Mode switch
            {
                LayoutMode.Row => ResolveFlow(container, content, path, horizontal: true),
                LayoutMode.Column => ResolveFlow(container, content, path, horizontal: false),
                LayoutMode.Grid => ResolveGrid(container, content, path),
                _ => throw new LayoutException(path, $"unknown layout mode '{container.Layout.Mode}'.")
            };
        }

        return new ResolvedContainer
        {
            X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = container.Id,
            Fill = ResolveFill(container.Fill, path),
            Stroke = ResolveStroke(container.Stroke, path),
            Radius = container.Radius,
            Shadow = ResolveShadow(container.Shadow, path),
            Overflow = container.Overflow,
            Children = children
        };
    }

    private ResolvedGroup ResolveGroup(GroupElement group, Rect rect, string path)
    {
        // A group is a paint-order construct, not a container: no padding, no overflow policy.
        var children = new List<ResolvedElement>(group.Children.Count);
        for (var i = 0; i < group.Children.Count; i++)
        {
            children.Add(ResolveFreeChild(group.Children[i], rect, $"{path}.children[{i}]", OverflowPolicy.Clip));
        }
        return new ResolvedGroup { X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = group.Id, Children = children };
    }

    private IReadOnlyList<ResolvedElement> ResolveFreeChildren(
        IReadOnlyList<GenElement> elements, Rect content, string path, OverflowPolicy overflow)
    {
        var resolved = new List<ResolvedElement>(elements.Count);
        for (var i = 0; i < elements.Count; i++)
        {
            resolved.Add(ResolveFreeChild(elements[i], content, $"{path}.children[{i}]", overflow));
        }
        return resolved;
    }

    private ResolvedElement ResolveFreeChild(GenElement child, Rect content, string path, OverflowPolicy overflow)
    {
        if (child.At is not { } at)
        {
            throw new LayoutException(path, "elements in a layout-less parent require 'at' ({\"x\":…,\"y\":…}).");
        }
        var size = child.Size;
        if (size?.Grow is not null)
        {
            throw new LayoutException(path, "'grow' shares space on a layout axis: it requires a layout container parent.");
        }

        var (w, h) = ResolveFreeSize(size, path);
        var rect = new Rect(content.X + at.X, content.Y + at.Y, w, h);

        var overX = rect.Right - content.Right;
        var overY = rect.Bottom - content.Bottom;
        if ((overX > Eps || overY > Eps) && overflow == OverflowPolicy.Error)
        {
            throw new LayoutException(path,
                $"element at ({at.X}, {at.Y}) sized {Round(w)}×{Round(h)} pt overflows its layout-less parent " +
                $"(content box {Round(content.W)}×{Round(content.H)} pt) by {Round(Math.Max(overX, 0))} pt horizontally and " +
                $"{Round(Math.Max(overY, 0))} pt vertically; set \"overflow\": \"clip\" on the parent to allow it.");
        }
        return ResolveElement(child, rect, path);
    }

    private static (double W, double H) ResolveFreeSize(SizeSpec? size, string path)
    {
        var w = size?.Width;
        var h = size?.Height;
        if (size?.Aspect is { } ar)
        {
            if (w is not null && h is not null)
            {
                throw new LayoutException(path,
                    "size is over-constrained: 'aspect' with both 'w' and 'h' fixed. Fix one dimension and let aspect derive the other.");
            }
            if (w is not null)
            {
                h = w.Value / ar.Value;
            }
            else if (h is not null)
            {
                w = h.Value * ar.Value;
            }
        }
        if (w is null || h is null)
        {
            throw new LayoutException(path,
                "elements in a layout-less parent require a fixed 'size': \"w\" and \"h\", or one dimension plus \"aspect\".");
        }
        return (w.Value, h.Value);
    }

    private IReadOnlyList<ResolvedElement> ResolveFlow(ContainerElement container, Rect content, string path, bool horizontal)
    {
        var layout = container.Layout!;
        var children = container.Children;
        var n = children.Count;
        if (n == 0)
        {
            return [];
        }

        var contentMain = horizontal ? content.W : content.H;
        var contentCross = horizontal ? content.H : content.W;
        var gap = layout.Gap;

        // 1) Main-axis size per child: fixed → grow → aspect-derived → error.
        var mains = new double[n];
        var grows = new double[n];
        double fixedTotal = 0, growTotal = 0;
        var hasGrow = false;
        for (var i = 0; i < n; i++)
        {
            var childPath = $"{path}.children[{i}]";
            var size = children[i].Size;
            var fixedMain = horizontal ? size?.Width : size?.Height;
            var fixedCross = horizontal ? size?.Height : size?.Width;

            if (size?.Aspect is not null && fixedMain is not null && fixedCross is not null)
            {
                throw new LayoutException(childPath,
                    "size is over-constrained: 'aspect' with both main and cross dimensions fixed. Fix one dimension and let aspect derive the other.");
            }
            if (size?.Grow is { } grow)
            {
                if (fixedMain is not null)
                {
                    throw new LayoutException(childPath,
                        "'grow' and a fixed main-axis size are mutually exclusive: grow shares the space remaining after fixed children.");
                }
                grows[i] = grow;
                growTotal += grow;
                hasGrow = true;
            }
            else if (fixedMain is { } fm)
            {
                mains[i] = fm;
                fixedTotal += fm;
            }
            else if (size?.Aspect is { } ar)
            {
                // Aspect resolves against the dimension the parent constrains first: in a
                // row/column that is the cross axis only when it is fixed; an unstretched
                // cross falls back to the full content cross (stretch semantics).
                var crossBasis = fixedCross ?? contentCross;
                mains[i] = horizontal ? crossBasis * ar.Value : crossBasis / ar.Value;
                fixedTotal += mains[i];
            }
            else
            {
                throw new LayoutException(childPath,
                    "undetermined main-axis size: give the child a fixed main dimension, 'grow', or 'aspect'.");
            }
        }

        // 2) Grow distribution over the remaining main space.
        var gapsTotal = gap * (n - 1);
        var remaining = contentMain - gapsTotal - fixedTotal;
        if (hasGrow)
        {
            if (remaining < -Eps)
            {
                CheckMainOverflow(container, path, -remaining);
                remaining = 0;
            }
            var distributable = Math.Max(remaining, 0);
            for (var i = 0; i < n; i++)
            {
                if (children[i].Size?.Grow is not null)
                {
                    mains[i] = growTotal > 0 ? distributable * grows[i] / growTotal : 0;
                }
            }
            remaining = 0;
        }
        else if (remaining < -Eps)
        {
            CheckMainOverflow(container, path, -remaining);
        }

        // 3) Justify: leading offset + extra space added to each declared gap.
        var free = Math.Max(remaining, 0);
        double lead, extra;
        switch (layout.Justify)
        {
            case Justify.Center: lead = free / 2; extra = 0; break;
            case Justify.End: lead = free; extra = 0; break;
            case Justify.SpaceBetween when n >= 2: lead = 0; extra = free / (n - 1); break;
            case Justify.SpaceEvenly: lead = free / (n + 1); extra = free / (n + 1); break;
            default: lead = 0; extra = 0; break; // Start (and SpaceBetween with a single child)
        }

        // 4) Place children on both axes.
        var resolved = new List<ResolvedElement>(n);
        var cursor = (horizontal ? content.X : content.Y) + lead;
        for (var i = 0; i < n; i++)
        {
            var childPath = $"{path}.children[{i}]";
            var size = children[i].Size;
            var fixedCross = horizontal ? size?.Height : size?.Width;
            var main = mains[i];

            double cross;
            if (fixedCross is { } fc)
            {
                cross = fc;
            }
            else if (size?.Aspect is { } ar)
            {
                cross = horizontal ? main / ar.Value : main * ar.Value;
            }
            else
            {
                cross = contentCross;
            }

            if (cross > contentCross + Eps && container.Overflow == OverflowPolicy.Error)
            {
                throw new LayoutException(childPath,
                    $"child cross-axis size {Round(cross)} pt overflows the container's content cross axis " +
                    $"({Round(contentCross)} pt); set \"overflow\": \"clip\" on the container to allow it.");
            }

            var align = size?.AlignSelf ?? layout.Align;
            var alignOffset = align switch
            {
                AlignItems.Center => (contentCross - cross) / 2,
                AlignItems.End => contentCross - cross,
                _ => 0 // Start, Stretch (cross already fills)
            };

            var rect = horizontal
                ? new Rect(cursor, content.Y + alignOffset, main, cross)
                : new Rect(content.X + alignOffset, cursor, cross, main);
            resolved.Add(ResolveElement(children[i], rect, childPath));
            cursor += main + gap + extra;
        }
        return resolved;
    }

    private IReadOnlyList<ResolvedElement> ResolveGrid(ContainerElement container, Rect content, string path)
    {
        var layout = container.Layout!;
        var children = container.Children;
        var n = children.Count;
        if (n == 0)
        {
            return [];
        }
        if (layout.Columns is not { } cols || cols < 1)
        {
            throw new LayoutException(path, "grid layout requires 'cols' ≥ 1 (rows derive from child count).");
        }

        var rows = (n + cols - 1) / cols;
        var colGap = layout.ColumnGap ?? layout.Gap;
        var rowGap = layout.RowGap ?? layout.Gap;

        var cellW = (content.W - colGap * (cols - 1)) / cols;
        if (cellW <= 0)
        {
            if (container.Overflow == OverflowPolicy.Error)
            {
                throw new LayoutException(path,
                    $"grid columns do not fit: {cols} columns plus {Round(colGap)} pt gaps exceed the content width " +
                    $"({Round(content.W)} pt).");
            }
            cellW = 0;
        }

        // Cell width is the dimension a grid constrains first: aspect derives height from it.
        var widths = new double[n];
        var heights = new double?[n];
        for (var i = 0; i < n; i++)
        {
            var childPath = $"{path}.children[{i}]";
            var size = children[i].Size;
            if (size?.Grow is not null)
            {
                throw new LayoutException(childPath,
                    "'grow' is not supported in grid layout: size rows via fixed 'h' or 'aspect'.");
            }
            if (size?.Aspect is not null && size.Width is not null && size.Height is not null)
            {
                throw new LayoutException(childPath,
                    "size is over-constrained: 'aspect' with both 'w' and 'h' fixed. Fix one dimension and let aspect derive the other.");
            }
            widths[i] = size?.Width ?? cellW;
            heights[i] = size?.Height ?? (size?.Aspect is { } ar ? widths[i] / ar.Value : null);
        }

        // Row height = the tallest determined child; rows with no determined child share
        // the remaining height equally. Leftover stays at the bottom (v1).
        var rowHeights = new double[rows];
        var determined = new bool[rows];
        double determinedTotal = 0;
        var undeterminedRows = 0;
        for (var i = 0; i < n; i++)
        {
            var r = i / cols;
            if (heights[i] is { } h)
            {
                if (!determined[r] || h > rowHeights[r])
                {
                    rowHeights[r] = Math.Max(rowHeights[r], h);
                }
                determined[r] = true;
            }
        }
        for (var r = 0; r < rows; r++)
        {
            if (determined[r])
            {
                determinedTotal += rowHeights[r];
            }
            else
            {
                undeterminedRows++;
            }
        }

        var remainingH = content.H - rowGap * (rows - 1) - determinedTotal;
        if (undeterminedRows > 0)
        {
            var share = remainingH / undeterminedRows;
            if (share < -Eps)
            {
                CheckMainOverflow(container, path, -share * undeterminedRows);
                share = 0;
            }
            for (var r = 0; r < rows; r++)
            {
                if (!determined[r])
                {
                    rowHeights[r] = Math.Max(share, 0);
                }
            }
        }
        else if (remainingH < -Eps)
        {
            CheckMainOverflow(container, path, -remainingH);
        }

        var resolved = new List<ResolvedElement>(n);
        for (var i = 0; i < n; i++)
        {
            var childPath = $"{path}.children[{i}]";
            var size = children[i].Size;
            var r = i / cols;
            var c = i % cols;

            var cellX = content.X + c * (cellW + colGap);
            var cellY = content.Y;
            for (var rr = 0; rr < r; rr++)
            {
                cellY += rowHeights[rr] + rowGap;
            }

            var w = widths[i];
            var h = heights[i] ?? rowHeights[r];

            if (w > cellW + Eps && container.Overflow == OverflowPolicy.Error)
            {
                throw new LayoutException(childPath,
                    $"child width {Round(w)} pt overflows its grid cell ({Round(cellW)} pt); " +
                    "set \"overflow\": \"clip\" on the container to allow it.");
            }

            var xOffset = layout.Justify switch
            {
                Justify.Center => (cellW - w) / 2,
                Justify.End => cellW - w,
                _ => 0 // Start; space-* is meaningless for one child per cell
            };
            var align = size?.AlignSelf ?? layout.Align;
            var yOffset = align switch
            {
                AlignItems.Center => (rowHeights[r] - h) / 2,
                AlignItems.End => rowHeights[r] - h,
                _ => 0 // Start, Stretch (height already fills the row)
            };

            resolved.Add(ResolveElement(children[i], new Rect(cellX + xOffset, cellY + yOffset, w, h), childPath));
        }
        return resolved;
    }

    private ResolvedText ResolveText(TextElement text, Rect rect, string path)
    {
        var defaultFamily = ResolveFontFamily(text.Font) ?? _design.Fonts.Body;
        var defaultSize = text.FontSize ?? _design.Metrics.BodySizePt;
        var defaultColor = ResolveColor(text.Color, path);

        var runs = new List<ResolvedTextRun>();
        if (text.Runs is { } modelRuns)
        {
            foreach (var run in modelRuns)
            {
                runs.Add(new ResolvedTextRun
                {
                    Text = run.Text,
                    FontFamily = ResolveFontFamily(run.Font) ?? defaultFamily,
                    FontSizePt = run.FontSize ?? defaultSize,
                    ColorHex = ResolveColor(run.Color, path) ?? defaultColor,
                    Bold = run.Bold || text.Bold,
                    Italic = run.Italic || text.Italic
                });
            }
        }
        else
        {
            runs.Add(new ResolvedTextRun
            {
                Text = text.Value ?? string.Empty,
                FontFamily = defaultFamily,
                FontSizePt = defaultSize,
                ColorHex = defaultColor,
                Bold = text.Bold,
                Italic = text.Italic
            });
        }

        var insets = text.Insets ?? new EdgeInsets(0, 0, 0, 0);
        var scale = 1.0;
        if (_textMeasurer is not null)
        {
            var request = new TextMeasureRequest
            {
                Runs = runs,
                BoxWidthPt = Math.Max(rect.W - insets.Left - insets.Right, 0),
                BoxHeightPt = Math.Max(rect.H - insets.Top - insets.Bottom, 0),
                MinScale = MinFontScale
            };
            var fit = Math.Clamp(_textMeasurer.FitScale(request), 0.01, 1.0);
            if (text.Overflow == OverflowPolicy.Clip)
            {
                // Clip passes through unscaled, but never silently.
                if (fit < 1 - Eps)
                {
                    _warnings.Add(
                        $"{path}: text does not fit its box ({Round(rect.W)}×{Round(rect.H)} pt) and its overflow policy is \"clip\"; content is clipped at the box edge.");
                }
            }
            else
            {
                scale = fit;
                if (scale < 1 - Eps && text.Overflow == OverflowPolicy.Error)
                {
                    throw new LayoutException(path,
                        $"text does not fit its box ({Round(rect.W)}×{Round(rect.H)} pt) and its overflow policy is \"error\".");
                }
                if (scale < MinFontScale - Eps && text.Overflow == OverflowPolicy.Shrink)
                {
                    _warnings.Add(
                        $"{path}: text shrunk to fontScale {Round(scale)} below MinScale {MinFontScale}; content may still overflow.");
                }
            }
        }

        return new ResolvedText
        {
            X = rect.X, Y = rect.Y, Width = rect.W, Height = rect.H, Id = text.Id,
            Runs = runs,
            TextAlign = text.TextAlign,
            Anchor = text.Anchor,
            Insets = insets,
            Overflow = text.Overflow,
            FontScale = scale,
            Shadow = ResolveShadow(text.Shadow, path)
        };
    }

    private void CheckMainOverflow(ContainerElement container, string path, double overflowBy)
    {
        if (container.Overflow == OverflowPolicy.Error)
        {
            throw new LayoutException(path,
                $"children overflow the container by {Round(overflowBy)} pt on the layout axis " +
                "(overflow \"error\" is the default for layout containers; " +
                "reduce sizes/gaps or set \"overflow\": \"clip\").");
        }
    }

    private string? ResolveFontFamily(string? font)
        => font switch
        {
            null => null,
            "display" => _design.Fonts.Display,
            "body" => _design.Fonts.Body,
            _ => font
        };

    private string? ResolveColor(string? color, string path)
    {
        if (color is null)
        {
            return null;
        }
        if (color.StartsWith('#'))
        {
            return color;
        }
        if (_design.Palette.TryGetValue(color, out var hex))
        {
            return hex;
        }
        throw new LayoutException(path, $"unknown palette token '{color}' (declared tokens: {string.Join(", ", _design.Palette.Keys)}).");
    }

    private FillSpec? ResolveFill(FillSpec? fill, string path)
        => fill switch
        {
            null => null,
            SolidFill solid => new SolidFill(ResolveColor(solid.Color, path)!),
            LinearGradientFill gradient => new LinearGradientFill
            {
                Angle = gradient.Angle,
                Stops = [.. gradient.Stops.Select(s => new GradientStop
                {
                    Color = ResolveColor(s.Color, path)!,
                    Offset = s.Offset,
                    Alpha = s.Alpha
                })]
            },
            _ => throw new LayoutException(path, $"unsupported fill type '{fill.GetType().Name}'.")
        };

    private StrokeSpec? ResolveStroke(StrokeSpec? stroke, string path)
        => stroke is null ? null : new StrokeSpec { Color = ResolveColor(stroke.Color, path)!, WidthPt = stroke.WidthPt };

    private ShadowSpec? ResolveShadow(ShadowSpec? shadow, string path)
        => shadow is null
            ? null
            : new ShadowSpec
            {
                Color = ResolveColor(shadow.Color, path)!,
                Dx = shadow.Dx,
                Dy = shadow.Dy,
                Blur = shadow.Blur,
                Alpha = shadow.Alpha
            };

    private static Rect ApplyPadding(Rect rect, EdgeInsets? padding)
    {
        if (padding is not { } pad)
        {
            return rect;
        }
        return new Rect(
            rect.X + pad.Left,
            rect.Y + pad.Top,
            Math.Max(rect.W - pad.Left - pad.Right, 0),
            Math.Max(rect.H - pad.Top - pad.Bottom, 0));
    }

    /// <summary>
    /// Builds the slide's flat paint-order element list for introspection. Depth-first so
    /// groups are flattened between their own entry and the next sibling — group children
    /// are already placed in absolute coordinates by the resolver, so no geometry is
    /// recomputed here (pure exposure of the single layout pass).
    /// </summary>
    private static IReadOnlyList<ResolvedElementInfo> Flatten(ResolvedElement root)
    {
        var list = new List<ResolvedElementInfo>();
        FlattenInto(root, list);
        return list;
    }

    private static void FlattenInto(ResolvedElement element, List<ResolvedElementInfo> list)
    {
        list.Add(new ResolvedElementInfo
        {
            Id = element.Id,
            Type = element switch
            {
                ResolvedContainer => ResolvedElementType.Container,
                ResolvedText => ResolvedElementType.Text,
                ResolvedRect => ResolvedElementType.Rect,
                ResolvedEllipse => ResolvedElementType.Ellipse,
                ResolvedLine => ResolvedElementType.Line,
                ResolvedImage => ResolvedElementType.Image,
                ResolvedGroup => ResolvedElementType.Group,
                _ => throw new NotSupportedException($"Unknown resolved element '{element.GetType().Name}'.")
            },
            X = element.X,
            Y = element.Y,
            Width = element.Width,
            Height = element.Height
        });
        switch (element)
        {
            case ResolvedContainer container:
                foreach (var child in container.Children)
                {
                    FlattenInto(child, list);
                }
                break;
            case ResolvedGroup group:
                foreach (var child in group.Children)
                {
                    FlattenInto(child, list);
                }
                break;
        }
    }

    private static double Round(double value) => Math.Round(value, Precision);

    private readonly record struct Rect(double X, double Y, double W, double H)
    {
        public double Right => X + W;
        public double Bottom => Y + H;

        public Rect Rounded() => new(Round(X), Round(Y), Round(W), Round(H));
    }
}
