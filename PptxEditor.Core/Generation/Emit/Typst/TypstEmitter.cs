using System.Globalization;
using System.Text;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Emit.Typst;

/// <summary>
/// Typst emitter (Typst is a dumb renderer): maps the absolute draw
/// tree produced by <see cref="LayoutResolver"/> to Typst source using ONLY
/// <c>#place(top + left, dx, dy)</c> plus absolute boxes. Typst performs no layout of its
/// own — every coordinate comes from the resolved tree.
/// <para>
/// Primitive mapping: text → <c>#block</c> + <c>#align(anchor + align)</c>
/// with one <c>#text</c> span per run; rect → <c>#rect(radius: (top-left: …))</c> with
/// per-corner radii in points (the same resolved values P4 maps to round1Rect/round2SameRect
/// adj — Typst takes pt directly, so no shared mapping table is needed); line → <c>#line</c>
/// centered in its box; ellipse → <c>#ellipse</c>; image → <c>#image(fit:)</c> (fill=cover,
/// crop=clipped scaled box or caller srcRect, contain=letterboxed); group → children in
/// paint order; linear gradient → <c>gradient.linear((color, offset)…, angle:)</c> — the
/// angle passes through verbatim (OOXML a:lin ang and Typst both measure clockwise from
/// left→right; the P7 parity fixture settles any residual divergence). Shadows are faked
/// as an offset copy of the shape/text in the shadow color (blur is Tier-3
/// and ignored in the preview). Container <c>overflow: clip</c> maps to
/// <c>#block(clip: true)</c> with children placed relative to the container.
/// </para>
/// </summary>
public sealed class TypstEmitter
{
    /// <summary>Stroke used for lines with no explicit stroke (emitter default, mirrored in tests).</summary>
    public const string DefaultLineColorHex = "#000000";

    /// <summary>Default stroke width for lines with no explicit stroke, in points.</summary>
    public const double DefaultLineWidthPt = 1;

    private const int Precision = 3;

    /// <summary>Emits a whole resolved document as one Typst source (one page per slide).</summary>
    public string Emit(LayoutResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        var sb = new StringBuilder();
        for (var i = 0; i < result.Slides.Count; i++)
        {
            EmitSlideBody(sb, result.Slides[i], $"slides[{i}]", isFirst: i == 0);
        }
        return sb.ToString();
    }

    /// <summary>Emits a single resolved slide as a standalone Typst source.</summary>
    public string EmitSlide(ResolvedSlide slide)
    {
        ArgumentNullException.ThrowIfNull(slide);
        var sb = new StringBuilder();
        EmitSlideBody(sb, slide, "slide", isFirst: true);
        return sb.ToString();
    }

    private void EmitSlideBody(StringBuilder sb, ResolvedSlide slide, string path, bool isFirst)
    {
        // Page setup per slide: Typst persists page properties across pagebreaks.
        sb.AppendLine($"#set page(width: {Pt(slide.WidthPt)}, height: {Pt(slide.HeightPt)}, margin: 0pt)");
        // No explicit background means white (the root container paints its own fill on top).
        sb.AppendLine("#set page(fill: rgb(\"#FFFFFF\"))");
        if (!isFirst)
        {
            sb.AppendLine("#pagebreak()");
        }
        EmitElement(sb, slide.Root, path + ".root", originX: 0, originY: 0);
    }

    private void EmitElement(StringBuilder sb, ResolvedElement element, string path, double originX, double originY)
    {
        switch (element)
        {
            case ResolvedContainer container:
                EmitContainer(sb, container, path, originX, originY);
                break;
            case ResolvedText text:
                EmitText(sb, text, path, originX, originY);
                break;
            case ResolvedRect rect:
                EmitShadowedShape(sb, rect, rect.Fill, rect.Stroke, rect.Radius, rect.Shadow, originX, originY, ShapeKind.Rect);
                break;
            case ResolvedEllipse ellipse:
                EmitShadowedShape(sb, ellipse, ellipse.Fill, ellipse.Stroke, radius: null, ellipse.Shadow, originX, originY, ShapeKind.Ellipse);
                break;
            case ResolvedLine line:
                EmitLine(sb, line, originX, originY);
                break;
            case ResolvedImage image:
                EmitImage(sb, image, path, originX, originY);
                break;
            case ResolvedGroup group:
                EmitGroup(sb, group, path, originX, originY);
                break;
            default:
                throw new TypstEmitException(path, $"unsupported resolved element type '{element.GetType().Name}'.");
        }
    }

    private void EmitContainer(StringBuilder sb, ResolvedContainer container, string path, double originX, double originY)
    {
        // 1) Faked shadow: offset copy of the container surface in the shadow color.
        if (container.Shadow is { } shadow)
        {
            EmitShapeLine(sb, container, shadow.Dx, shadow.Dy, FillFromShadow(shadow), stroke: null, container.Radius, originX, originY, ShapeKind.Rect);
        }

        // 2) Visible background (radius alone with no fill/stroke paints nothing).
        if (container.Fill is not null || container.Stroke is not null)
        {
            EmitShapeLine(sb, container, dx: 0, dy: 0, container.Fill, container.Stroke, container.Radius, originX, originY, ShapeKind.Rect);
        }

        // 3) Children in paint order. overflow: clip wraps them in a clipped block whose
        //    origin is the container, so children are placed relative to it.
        if (container.Overflow == OverflowPolicy.Clip && container.Children.Count > 0)
        {
            sb.AppendLine($"#place(top + left, dx: {Pt(container.X - originX)}, dy: {Pt(container.Y - originY)})[#block(width: {Pt(container.Width)}, height: {Pt(container.Height)}, clip: true)[");
            for (var i = 0; i < container.Children.Count; i++)
            {
                EmitElement(sb, container.Children[i], ChildPath(path, i), container.X, container.Y);
            }
            sb.AppendLine("]]");
            return;
        }

        for (var i = 0; i < container.Children.Count; i++)
        {
            EmitElement(sb, container.Children[i], ChildPath(path, i), originX, originY);
        }
    }

    private void EmitGroup(StringBuilder sb, ResolvedGroup group, string path, double originX, double originY)
    {
        // A group is paint order only: children already carry absolute coordinates.
        for (var i = 0; i < group.Children.Count; i++)
        {
            EmitElement(sb, group.Children[i], ChildPath(path, i), originX, originY);
        }
    }

    private enum ShapeKind { Rect, Ellipse }

    private void EmitShadowedShape(
        StringBuilder sb, ResolvedElement element, FillSpec? fill, StrokeSpec? stroke, CornerRadii? radius,
        ShadowSpec? shadow, double originX, double originY, ShapeKind kind)
    {
        if (shadow is not null)
        {
            EmitShapeLine(sb, element, shadow.Dx, shadow.Dy, FillFromShadow(shadow), stroke: null, radius, originX, originY, kind);
        }
        EmitShapeLine(sb, element, 0, 0, fill, stroke, radius, originX, originY, kind);
    }

    private void EmitShapeLine(
        StringBuilder sb, ResolvedElement element, double dx, double dy, FillSpec? fill, StrokeSpec? stroke,
        CornerRadii? radius, double originX, double originY, ShapeKind kind)
    {
        var name = kind == ShapeKind.Rect ? "rect" : "ellipse";
        var parameters = new StringBuilder($"width: {Pt(element.Width)}, height: {Pt(element.Height)}");
        if (fill is not null)
        {
            parameters.Append($", fill: {FormatFill(fill)}");
        }
        if (stroke is not null)
        {
            parameters.Append($", stroke: {FormatStroke(stroke)}");
        }
        if (kind == ShapeKind.Rect && radius is { } r && (r.TopLeft != 0 || r.TopRight != 0 || r.BottomRight != 0 || r.BottomLeft != 0))
        {
            // Typst corner order matches CornerRadii: top-left, top-right, bottom-right, bottom-left.
            parameters.Append(
                $", radius: (top-left: {Pt(r.TopLeft)}, top-right: {Pt(r.TopRight)}, bottom-right: {Pt(r.BottomRight)}, bottom-left: {Pt(r.BottomLeft)})");
        }
        sb.AppendLine($"#place(top + left, dx: {Pt(element.X + dx - originX)}, dy: {Pt(element.Y + dy - originY)})[#{name}({parameters})]");
    }

    private void EmitText(StringBuilder sb, ResolvedText text, string path, double originX, double originY)
    {
        if (text.Runs.Count == 0)
        {
            throw new TypstEmitException(path, "text element has no runs; the resolver always produces at least one.");
        }

        // Insets shrink the text box (visually identical to OOXML bodyPr insets).
        var x = text.X + text.Insets.Left;
        var y = text.Y + text.Insets.Top;
        var w = Math.Max(text.Width - text.Insets.Left - text.Insets.Right, 0);
        var h = Math.Max(text.Height - text.Insets.Top - text.Insets.Bottom, 0);

        if (text.Shadow is { } shadow)
        {
            // Faked text shadow: offset copy of the same runs, all in the shadow color.
            var shadowColor = WithAlpha(shadow.Color, shadow.Alpha);
            AppendTextLine(sb, x + shadow.Dx - originX, y + shadow.Dy - originY, w, h, text, colorOverride: shadowColor);
        }
        AppendTextLine(sb, x - originX, y - originY, w, h, text, colorOverride: null);
    }

    private static void AppendTextLine(
        StringBuilder sb, double dx, double dy, double w, double h, ResolvedText text, string? colorOverride)
    {
        var anchor = text.Anchor switch
        {
            TextAnchor.Middle => "horizon",
            TextAnchor.Bottom => "bottom",
            _ => "top"
        };
        var align = text.TextAlign switch
        {
            TextAlign.Center => "center",
            TextAlign.Right => "right",
            _ => "left"
        };

        var runs = new StringBuilder();
        foreach (var run in text.Runs)
        {
            var parameters = new StringBuilder();
            if (run.FontFamily is not null)
            {
                parameters.Append($", font: \"{EscapeStringLiteral(run.FontFamily)}\"");
            }
            parameters.Append($", size: {Pt(run.FontSizePt * text.FontScale)}");
            var color = colorOverride ?? run.ColorHex;
            if (color is not null)
            {
                parameters.Append($", fill: rgb(\"{color}\")");
            }
            if (run.Bold)
            {
                parameters.Append(", weight: \"bold\"");
            }
            if (run.Italic)
            {
                parameters.Append(", style: \"italic\"");
            }
            runs.Append($"#text({parameters.ToString().TrimStart(',', ' ')})[{EscapeMarkup(run.Text)}]");
        }

        sb.AppendLine($"#place(top + left, dx: {Pt(dx)}, dy: {Pt(dy)})[#block(width: {Pt(w)}, height: {Pt(h)})[#align({anchor} + {align})[{runs}]]]");
    }

    private void EmitLine(StringBuilder sb, ResolvedLine line, double originX, double originY)
    {
        // Straight only in v1: the line spans its box on the orientation axis,
        // centered on the cross axis. Connector vs line is an OOXML distinction only.
        var (sx, sy, ex, ey) = line.Orientation == LineOrientation.Vertical
            ? (line.Width / 2, 0.0, line.Width / 2, line.Height)
            : (0.0, line.Height / 2, line.Width, line.Height / 2);
        var stroke = line.Stroke ?? new StrokeSpec { Color = DefaultLineColorHex, WidthPt = DefaultLineWidthPt };
        sb.AppendLine(
            $"#place(top + left, dx: {Pt(line.X - originX)}, dy: {Pt(line.Y - originY)})" +
            $"[#line(start: ({Pt(sx)}, {Pt(sy)}), end: ({Pt(ex)}, {Pt(ey)}), stroke: {FormatStroke(stroke)})]");
    }

    private void EmitImage(StringBuilder sb, ResolvedImage image, string path, double originX, double originY)
    {
        var source = image.Source;
        if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            throw new TypstEmitException(path,
                "base64 image sources are not supported by the Typst emitter (Typst cannot decode them in-source); " +
                "materialize the image to a file and pass its path.");
        }
        if (source.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            source.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            throw new TypstEmitException(path,
                "remote image URLs are not supported by the Typst emitter (Typst reads local files only); " +
                "download the image and pass its local path.");
        }

        var src = $"\"{EscapeStringLiteral(source)}\"";
        var place = $"#place(top + left, dx: {Pt(image.X - originX)}, dy: {Pt(image.Y - originY)})";
        var w = Pt(image.Width);
        var h = Pt(image.Height);

        switch (image.Fit)
        {
            case ImageFitMode.Stretch:
                sb.AppendLine($"{place}[#image({src}, width: {w}, height: {h}, fit: \"stretch\")]");
                break;
            case ImageFitMode.Fill:
            case ImageFitMode.Crop when image.Crop is null:
                // Fill = cover: center-crop at the frame's aspect (F7 semantics; crop with no
                // srcRect is treated as fill, see ImageFitMode.Crop).
                sb.AppendLine($"{place}[#image({src}, width: {w}, height: {h}, fit: \"cover\")]");
                break;
            case ImageFitMode.Contain:
                // The frame never moves in the draw tree; contain letterboxes inside the box.
                sb.AppendLine($"{place}[#block(width: {w}, height: {h})[#align(center + horizon)[#image({src}, width: {w}, height: {h}, fit: \"contain\")]]]");
                break;
            case ImageFitMode.Crop:
                EmitCroppedImage(sb, image, src, place, path);
                break;
            default:
                throw new TypstEmitException(path, $"unknown image fit mode '{image.Fit}'.");
        }
    }

    private void EmitCroppedImage(StringBuilder sb, ResolvedImage image, string src, string place, string path)
    {
        // Caller srcRect (1/1000ths of a percent cropped from each edge, spcPct
        // scale family): PowerPoint stretches the remaining source region into the frame.
        // Emulated with a clipped block containing the image scaled so the visible region
        // covers the box, offset by the cropped left/top fractions.
        var crop = image.Crop!.Value;
        const double full = 100000.0;
        var visibleW = (full - crop.Left - crop.Right) / full;
        var visibleH = (full - crop.Top - crop.Bottom) / full;
        if (visibleW <= 0 || visibleH <= 0)
        {
            throw new TypstEmitException(path,
                $"crop srcRect ({crop.Left}, {crop.Top}, {crop.Right}, {crop.Bottom}) leaves no visible source region; " +
                "left+right and top+bottom must each be below 100000 (100%).");
        }

        var scaledW = image.Width / visibleW;
        var scaledH = image.Height / visibleH;
        var dx = -crop.Left / full * scaledW;
        var dy = -crop.Top / full * scaledH;

        sb.AppendLine(
            $"{place}[#block(width: {Pt(image.Width)}, height: {Pt(image.Height)}, clip: true)" +
            $"[#place(top + left, dx: {Pt(dx)}, dy: {Pt(dy)})[#image({src}, width: {Pt(scaledW)}, height: {Pt(scaledH)}, fit: \"stretch\")]]]");
    }

    private static SolidFill FillFromShadow(ShadowSpec shadow) => new(WithAlpha(shadow.Color, shadow.Alpha));

    private static string FormatFill(FillSpec fill)
        => fill switch
        {
            SolidFill solid => $"rgb(\"{solid.Color}\")",
            // Angle passes through verbatim: OOXML a:lin ang and Typst both measure degrees
            // clockwise from the left→right axis (parity settled by P7 fixtures).
            LinearGradientFill gradient =>
                $"gradient.linear({string.Join(", ", gradient.Stops.Select(s => $"(rgb(\"{WithAlpha(s.Color, s.Alpha)}\"), {Fmt(s.Offset * 100)}%)"))}, angle: {Fmt(gradient.Angle)}deg)",
            _ => throw new TypstEmitException("fill", $"unsupported fill type '{fill.GetType().Name}'.")
        };

    private static string FormatStroke(StrokeSpec stroke) => $"{Pt(stroke.WidthPt)} + rgb(\"{stroke.Color}\")";

    /// <summary>Applies an alpha multiplier to a #RRGGBB color, yielding #RRGGBBAA (Typst accepts 8-digit hex).</summary>
    private static string WithAlpha(string colorHex, double? alpha)
    {
        if (alpha is not { } a || a >= 1)
        {
            return colorHex;
        }
        var channel = (int)Math.Round(Math.Clamp(a, 0, 1) * 255, MidpointRounding.AwayFromZero);
        return colorHex + channel.ToString("X2", CultureInfo.InvariantCulture);
    }

    private static string Pt(double value) => Fmt(value) + "pt";

    private static string Fmt(double value)
    {
        var rounded = Math.Round(value, Precision);
        if (rounded == 0)
        {
            rounded = 0; // normalize negative zero
        }
        return rounded.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string ChildPath(string path, int index) => $"{path}.children[{index}]";

    /// <summary>Escapes a Typst string literal (font families, image paths).</summary>
    private static string EscapeStringLiteral(string value)
        => value.Replace("\\", "\\\\").Replace("\"", "\\\"");

    /// <summary>Escapes markup-significant characters in text content (same set as the PPTX→Typst converter).</summary>
    private static string EscapeMarkup(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }
        return text
            .Replace("\\", "\\\\")
            .Replace("[", "\\[")
            .Replace("]", "\\]")
            .Replace("*", "\\*")
            .Replace("_", "\\_")
            .Replace("#", "\\#")
            .Replace("`", "\\`")
            .Replace("$", "\\$")
            .Replace("%", "\\%")
            .Replace("&", "\\&")
            .Replace("@", "\\@")
            .Replace("^", "\\^")
            .Replace("~", "\\~")
            .Replace("<", "\\<")
            .Replace(">", "\\>")
            .Replace("/", "\\/")
            .Replace("{", "\\{")
            .Replace("}", "\\}");
    }
}
