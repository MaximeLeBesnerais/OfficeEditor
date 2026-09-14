using System.Globalization;
using DocumentFormat.OpenXml;
using DocxEditor.Core.Models;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxEditor.Core.Converters;

public sealed partial class DocxToTypstConverter
{
    private static OpenXmlElement? Child(OpenXmlElement? element, string name) =>
        element?.Elements().FirstOrDefault(e => e.LocalName == name);

    private static bool IsDrawingGroup(OpenXmlElement element) =>
        element.LocalName is "wgp" or "grpSp"
        && element.NamespaceUri is "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            or "http://schemas.openxmlformats.org/drawingml/2006/main";

    private ShapeGeometry ResolveDrawingGeometry(OpenXmlElement element, ShapeGeometry geometry, TypstPageSetup page)
    {
        var groups = element.Ancestors().Where(IsDrawingGroup).ToList();
        double x = geometry.XPt ?? 0;
        double y = geometry.YPt ?? 0;
        double width = geometry.WidthPt ?? 0;
        double height = geometry.HeightPt ?? 0;
        double outerX = 0;
        double outerY = 0;

        // Walk from the innermost group outwards. A group's child coordinate space is
        // independent of its own extent; apply chOff/chExt before its parent transform.
        foreach (var group in groups)
        {
            var transform = Child(Child(group, "grpSpPr"), "xfrm");
            var off = Child(transform, "off");
            var ext = Child(transform, "ext");
            var childOff = Child(transform, "chOff");
            var childExt = Child(transform, "chExt");
            outerX = EmuToPoints(GetXmlAttribute(off, "x")) ?? 0;
            outerY = EmuToPoints(GetXmlAttribute(off, "y")) ?? 0;
            var groupWidth = EmuToPoints(GetXmlAttribute(ext, "cx"));
            var groupHeight = EmuToPoints(GetXmlAttribute(ext, "cy"));
            var childWidth = EmuToPoints(GetXmlAttribute(childExt, "cx"));
            var childHeight = EmuToPoints(GetXmlAttribute(childExt, "cy"));
            double sx = childWidth is > 0 && groupWidth is > 0 ? groupWidth.Value / childWidth.Value : 1;
            double sy = childHeight is > 0 && groupHeight is > 0 ? groupHeight.Value / childHeight.Value : 1;
            x = outerX + (x - (EmuToPoints(GetXmlAttribute(childOff, "x")) ?? 0)) * sx;
            y = outerY + (y - (EmuToPoints(GetXmlAttribute(childOff, "y")) ?? 0)) * sy;
            width *= sx;
            height *= sy;

            if (GetXmlAttribute(transform, "rot") is { } rotation && rotation != "0"
                || GetXmlAttribute(transform, "flipH") is "1" or "true"
                || GetXmlAttribute(transform, "flipV") is "1" or "true")
            {
                const string warning = "DrawingML group rotation/reflection is not supported; translation and scale were preserved.";
                if (!diagnostics.Contains(warning)) diagnostics.Add(warning);
            }
        }

        var anchor = element.Ancestors<DW.Anchor>().FirstOrDefault();
        if (anchor is null)
            return geometry with { XPt = x, YPt = y, WidthPt = width, HeightPt = height };

        bool canvas = element.Ancestors().Any(e => e.LocalName == "wpc");
        if (groups.Count > 0 && !canvas)
        {
            // The outer group's position in its host is supplied by wp:anchor. Its
            // a:xfrm/off is not a second page translation (same rule as a lone shape).
            x -= outerX;
            y -= outerY;
        }
        else if (!canvas)
        {
            x = 0;
            y = 0;
        }

        var position = ExtractAnchorPosition(anchor);
        var anchorWidth = anchor.Extent?.Cx?.Value / EmusPerPoint ?? width;
        var anchorHeight = anchor.Extent?.Cy?.Value / EmusPerPoint ?? height;
        var placement = ResolveDrawingPlacement(position, page, anchorWidth, anchorHeight);
        return geometry with
        {
            XPt = (position.XPt ?? 0) + x,
            YPt = (position.YPt ?? 0) + y,
            WidthPt = geometry.WidthPt is null ? anchorWidth : width,
            HeightPt = geometry.HeightPt is null ? anchorHeight : height,
            Placement = placement with { XPt = placement.XPt + x, YPt = placement.YPt + y }
        };
    }

    private TypstDrawingPlacement ResolveDrawingPlacement(AnchorPosition position, TypstPageSetup page, double width, double height)
    {
        double pageWidth = page.WidthInches * 72;
        double pageHeight = page.HeightInches * 72;
        double left = page.Margins.LeftInches * 72;
        double top = page.Margins.TopInches * 72;
        double right = page.Margins.RightInches * 72;
        double bottom = page.Margins.BottomInches * 72;
        var horizontal = position.HorizontalRelativeFrom?.ToLowerInvariant();
        var vertical = position.VerticalRelativeFrom?.ToLowerInvariant();

        // Typst's top-level place uses the text area. Express the OOXML reference
        // rectangle in that coordinate system, then align the drawing within it.
        var (xOrigin, xSpan) = horizontal switch
        {
            "page" => (-left, pageWidth),
            "leftmargin" => (-left, left),
            "rightmargin" => (pageWidth - left - right, right),
            _ => (0d, pageWidth - left - right)
        };
        var (yOrigin, ySpan) = vertical switch
        {
            "page" => (-top, pageHeight),
            "topmargin" => (-top, top),
            "bottommargin" => (pageHeight - top - bottom, bottom),
            _ => (0d, pageHeight - top - bottom)
        };
        bool followsParagraph = vertical is "paragraph" or "line";
        if (horizontal is "character" or "insidemargin" or "outsidemargin"
            || vertical is "insidemargin" or "outsidemargin")
        {
            const string warning = "Character and inside/outside margin anchors are approximated using the text area.";
            if (!diagnostics.Contains(warning)) diagnostics.Add(warning);
        }

        double dx = position.HorizontalAlignment switch
        {
            "center" => (xSpan - width) / 2,
            "right" => xSpan - width,
            _ => position.XPt ?? 0
        };
        double dy = position.VerticalAlignment switch
        {
            "horizon" => (ySpan - height) / 2,
            "bottom" => ySpan - height,
            _ => position.YPt ?? 0
        };
        return new TypstDrawingPlacement(xOrigin + dx, yOrigin + dy, followsParagraph ? null : "top");
    }

    private static TypstTextInsets? ExtractTextInsets(OpenXmlElement element)
    {
        if (element.LocalName == "wsp")
        {
            var body = Child(element, "bodyPr");
            return new TypstTextInsets(
                EmuToPoints(GetXmlAttribute(body, "lIns")) ?? 7.2,
                EmuToPoints(GetXmlAttribute(body, "tIns")) ?? 3.6,
                EmuToPoints(GetXmlAttribute(body, "rIns")) ?? 7.2,
                EmuToPoints(GetXmlAttribute(body, "bIns")) ?? 3.6);
        }

        // Keep the legacy fallback for VML without explicit insets; honor authored
        // values (including zero) when the textbox supplies them.
        var inset = GetXmlAttribute(Child(element, "textbox"), "inset");
        if (inset is null) return null;
        var sides = inset.Split(',');
        if (sides.Length != 4) return null;
        double ReadSide(int index, double fallback)
        {
            var raw = sides[index].Trim();
            var multiplier = raw.EndsWith("in", StringComparison.OrdinalIgnoreCase) ? 72d : 1d;
            if (raw.EndsWith("pt", StringComparison.OrdinalIgnoreCase) || multiplier == 72) raw = raw[..^2];
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) && double.IsFinite(value)
                ? value * multiplier : fallback;
        }
        return new TypstTextInsets(ReadSide(0, 7.2), ReadSide(1, 3.6), ReadSide(2, 7.2), ReadSide(3, 3.6));
    }

    private static string FormatTextInsets(TypstTextInsets? insets) => insets is null
        ? "inset: 4pt"
        : $"inset: (left: {FormatPt(insets.LeftPt)}, right: {FormatPt(insets.RightPt)}, top: {FormatPt(insets.TopPt)}, bottom: {FormatPt(insets.BottomPt)})";

    private static string PlaceDrawing(string content, TypstDrawingPlacement placement) =>
        $"#place({placement.Alignment}, dx: {FormatPt(placement.XPt)}, dy: {FormatPt(placement.YPt)})[{content}]";
}
