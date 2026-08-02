using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Model;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Images;

/// <summary>
/// Positioned (floating) picture construction: creates the
/// <see cref="PositionedPictureContent"/> seam from a registered part + geometry, and turns
/// it into a <c>wp:anchor</c> drawing for a <see cref="PositionSpec"/>. The anchor carries
/// wrap mode/distances, anchor reference, z-order (relativeHeight) and rotation, and wraps
/// the same <c>pic:pic</c> subtree as the inline emitter.
/// </summary>
public static class PositionedPictureFactory
{
    private const uint BaseRelativeHeight = 251658240; // Word's common drawing-layer base

    /// <summary>
    /// Creates the positioned-picture seam: content (relationship + geometry) and the
    /// deterministic docPr identity, decoupled from the anchor geometry. The positioned
    /// emitter calls this once per placement and later passes it to
    /// <see cref="BuildAnchored"/> with the resolved <see cref="PositionSpec"/>.
    /// </summary>
    public static PositionedPictureContent CreatePlacement(
        RegisteredImagePart part, ResolvedImageGeometry geometry, uint docPrId, string? altText)
    {
        ArgumentNullException.ThrowIfNull(part);
        ArgumentNullException.ThrowIfNull(geometry);
        return new PositionedPictureContent
        {
            ImagePart = part,
            Geometry = geometry,
            DocPrId = docPrId,
            AltText = altText
        };
    }

    /// <summary>
    /// Builds the anchored picture drawing (<c>wp:anchor</c>) for a placement and its
    /// position spec, ready to be appended inside a <c>w:drawing</c>.
    /// </summary>
    public static Drawing BuildAnchored(PositionedPictureContent placement, PositionSpec position)
    {
        ArgumentNullException.ThrowIfNull(placement);
        ArgumentNullException.ThrowIfNull(position);

        var geometry = placement.Geometry;
        var wrapDistances = position.WrapDistances;

        var anchor = new Wp.Anchor
        {
            DistanceFromTop = ToEmu(wrapDistances?.TopPt ?? 0),
            DistanceFromBottom = ToEmu(wrapDistances?.BottomPt ?? 0),
            DistanceFromLeft = ToEmu(wrapDistances?.LeftPt ?? 0),
            DistanceFromRight = ToEmu(wrapDistances?.RightPt ?? 0),
            SimplePos = false,
            RelativeHeight = BaseRelativeHeight + (uint)Math.Max(0, position.ZOrder),
            BehindDoc = position.Wrap == WrapMode.BehindText,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = true
        };

        anchor.Append(new Wp.SimplePosition { X = 0, Y = 0 });
        anchor.Append(BuildHorizontalPosition(position));
        anchor.Append(BuildVerticalPosition(position));
        anchor.Append(new Wp.Extent { Cx = geometry.ExtentsCxEmu, Cy = geometry.ExtentsCyEmu });
        anchor.Append(new Wp.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 });
        anchor.Append(BuildWrap(position));
        anchor.Append(new Wp.DocProperties
        {
            Id = placement.DocPrId,
            Name = $"Picture {placement.DocPrId}",
            Description = placement.AltText
        });
        anchor.Append(new Wp.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }));
        anchor.Append(PictureGraphicBuilder.BuildGraphic(
            placement.ImagePart,
            geometry,
            placement.DocPrId,
            placement.AltText,
            ToSixtiethsOfDegree(position.Rotation)));

        return new Drawing(anchor);
    }

    private static Wp.HorizontalPosition BuildHorizontalPosition(PositionSpec position)
    {
        var relativeFrom = position.Anchor switch
        {
            AnchorReference.Page => Wp.HorizontalRelativePositionValues.Page,
            AnchorReference.Column => Wp.HorizontalRelativePositionValues.Column,
            _ => Wp.HorizontalRelativePositionValues.Margin
        };
        var horizontal = new Wp.HorizontalPosition { RelativeFrom = relativeFrom };
        horizontal.Append(new Wp.PositionOffset(ToEmu(position.X).ToString(CultureInfo.InvariantCulture)));
        return horizontal;
    }

    private static Wp.VerticalPosition BuildVerticalPosition(PositionSpec position)
    {
        var relativeFrom = position.Anchor switch
        {
            AnchorReference.Page => Wp.VerticalRelativePositionValues.Page,
            AnchorReference.Paragraph => Wp.VerticalRelativePositionValues.Paragraph,
            AnchorReference.Character => Wp.VerticalRelativePositionValues.Line,
            _ => Wp.VerticalRelativePositionValues.Margin
        };
        var vertical = new Wp.VerticalPosition { RelativeFrom = relativeFrom };
        vertical.Append(new Wp.PositionOffset(ToEmu(position.Y).ToString(CultureInfo.InvariantCulture)));
        return vertical;
    }

    private static OpenXmlElement BuildWrap(PositionSpec position)
    {
        var distances = position.WrapDistances;
        var top = distances?.TopPt ?? 0;
        var bottom = distances?.BottomPt ?? 0;
        var left = distances?.LeftPt ?? 0;
        var right = distances?.RightPt ?? 0;

        return position.Wrap switch
        {
            WrapMode.None or WrapMode.BehindText or WrapMode.InFrontOfText => new Wp.WrapNone(),
            WrapMode.Square => new Wp.WrapSquare
            {
                WrapText = Wp.WrapTextValues.BothSides,
                DistanceFromTop = ToEmu(top),
                DistanceFromBottom = ToEmu(bottom),
                DistanceFromLeft = ToEmu(left),
                DistanceFromRight = ToEmu(right)
            },
            WrapMode.Tight => new Wp.WrapTight
            {
                WrapText = Wp.WrapTextValues.BothSides,
                DistanceFromLeft = ToEmu(left),
                DistanceFromRight = ToEmu(right)
            },
            WrapMode.Through => new Wp.WrapThrough
            {
                WrapText = Wp.WrapTextValues.BothSides,
                DistanceFromLeft = ToEmu(left),
                DistanceFromRight = ToEmu(right)
            },
            WrapMode.TopAndBottom => new Wp.WrapTopBottom
            {
                DistanceFromTop = ToEmu(top),
                DistanceFromBottom = ToEmu(bottom)
            },
            _ => throw new NotSupportedException($"Wrap mode '{position.Wrap}' is not supported.")
        };
    }

    private static uint ToEmu(double points) => (uint)Math.Max(0, DrawingUnits.ToEmu(points));

    private static long ToSixtiethsOfDegree(double degrees) =>
        (long)Math.Round(degrees * 60000, MidpointRounding.AwayFromZero);
}
