using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>image_card</c> (plan.md §4): an image growing into the card's space above an
/// optional caption block (title + muted subtitle) on a token-styled card surface.
/// Expands to a column container; the caption block height is computed from the line
/// estimates so the image takes the rest.
/// <para>Overflow contract: the image is clipped by its fit mode (default crop — the
/// fixed-canvas policy, plan.md §1); caption texts shrink; structural overflow errors.</para>
/// </summary>
public static class ImageCardComponent
{
    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "src", "title", "subtitle", "fit", "alt", "imageGrow" };

    private static readonly IReadOnlyDictionary<string, ImageFitMode> FitModes =
        new Dictionary<string, ImageFitMode>(StringComparer.OrdinalIgnoreCase)
        {
            ["fill"] = ImageFitMode.Fill, ["crop"] = ImageFitMode.Crop, ["contain"] = ImageFitMode.Contain
        };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    public static ContainerElement Expand(ComponentElement element, DesignTokens design, string path)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "image_card", element.Content, ContentProps, design.Palette);
        var src = reader.String("src", required: true);
        var title = reader.String("title");
        var subtitle = reader.String("subtitle");
        var fit = reader.Enum("fit", FitModes, "image fit mode") ?? ImageFitMode.Crop;
        var alt = reader.String("alt");
        var imageGrow = reader.Number("imageGrow", min: 0.001) ?? 1;
        reader.ThrowIfInvalid();

        var metrics = design.Metrics;
        var children = new List<GenElement>
        {
            new ImageElement
            {
                Source = src!,
                Fit = fit,
                Alt = alt,
                Size = new SizeSpec { Grow = imageGrow }
            }
        };

        if (title is not null || subtitle is not null)
        {
            var pad = metrics.GutterPt;
            var captionHeight = 2 * pad
                + (title is not null ? ComponentStyle.LineHeight(metrics.BodySizePt + 2) : 0)
                + (subtitle is not null ? ComponentStyle.LineHeight(metrics.BodySizePt) : 0)
                + (title is not null && subtitle is not null ? 4 : 0);

            var captionChildren = new List<GenElement>();
            if (title is not null)
            {
                captionChildren.Add(ComponentStyle.Line(title, "display", metrics.BodySizePt + 2, "ink", bold: true));
            }
            if (subtitle is not null)
            {
                captionChildren.Add(ComponentStyle.Line(subtitle, "body", metrics.BodySizePt, "muted"));
            }

            children.Add(new ContainerElement
            {
                Size = new SizeSpec { Height = Math.Round(captionHeight, 3) },
                Padding = EdgeInsets.All(pad),
                Layout = new LayoutSpec { Mode = LayoutMode.Column, Gap = 4 },
                Children = captionChildren
            });
        }

        ComponentStyle.RequireTokens(design, path, "image_card", "paper", title is not null ? "ink" : null, subtitle is not null ? "muted" : null);

        return ComponentStyle.ApplySurface(new ContainerElement
        {
            Size = element.Size,
            At = element.At,
            Layout = new LayoutSpec { Mode = LayoutMode.Column },
            Children = children
        }, design);
    }
}
