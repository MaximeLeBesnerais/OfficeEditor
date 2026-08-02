using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Positioned;

/// <summary>
/// Builds the DrawingML payloads that live inside a <c>wp:anchor</c> graphic data:
/// <c>wps:wsp</c> wordprocessing shapes (text boxes, rectangles, lines) and the <c>p:pic</c>
/// floating-picture element. The OpenXML SDK ships no typed classes for the
/// <c>wps</c> (Office 2010 wordprocessingShape) namespace, so those containers are built
/// as <see cref="OpenXmlUnknownElement"/> nodes with typed DrawingML children
/// (<c>a:xfrm</c>, <c>a:prstGeom</c>, <c>a:solidFill</c>, …), matching the exact element
/// order Word produces. All color/geometry inputs are pre-resolved by the emitter; this
/// class is a pure translator to OOXML.
/// </summary>
internal static class WordprocessingShapeBuilder
{
    /// <summary>XML namespace of the <c>wps</c> (Office 2010 wordprocessing shape) vocabulary.</summary>
    public const string WpsNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    /// <summary>XML namespace of the <c>pic</c> (DrawingML picture) vocabulary.</summary>
    public const string PictureNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    /// <summary>Default text insets for a text box: 0.1in left/right, 0.05in top/bottom (EMU), as Word writes.</summary>
    private const int DefaultLeftRightInsetEmu = 91440;

    private const int DefaultTopBottomInsetEmu = 45720;

    private const double EmuPerPoint = 12700.0;

    private const int AngleScale = 60000; // OOXML angles are 60000ths of a degree

    private const int FullAngle = 21600000; // 360° in OOXML angle units

    /// <summary>Builds a <c>wps:wsp</c> text box with WML text content.</summary>
    public static OpenXmlUnknownElement BuildTextBox(
        uint id,
        string name,
        string? alt,
        double widthPt,
        double heightPt,
        double rotation,
        double cornerRadiusPt,
        string? fillHex,
        string? strokeHex,
        double strokeWidthPt,
        TextBoxContent content)
    {
        OpenXmlUnknownElement wsp = BuildBoxShape(
            id, name, alt, widthPt, heightPt, rotation, cornerRadiusPt, fillHex, strokeHex, strokeWidthPt,
            txBox: true);
        wsp.Append(Wps("txbx", content));
        wsp.Append(BuildTextBodyProperties());
        return wsp;
    }

    /// <summary>Builds a <c>wps:wsp</c> rectangle (no text content).</summary>
    public static OpenXmlUnknownElement BuildRectangle(
        uint id,
        string name,
        string? alt,
        double widthPt,
        double heightPt,
        double rotation,
        double cornerRadiusPt,
        string? fillHex,
        string? strokeHex,
        double strokeWidthPt)
    {
        return BuildBoxShape(
            id, name, alt, widthPt, heightPt, rotation, cornerRadiusPt, fillHex, strokeHex, strokeWidthPt,
            txBox: false);
    }

    /// <summary>
    /// Builds a <c>wps:wsp</c> straight line. A horizontal line is a <c>straightConnector1</c>
    /// spanning its length with a zero-height box (the way Word persists a horizontal line);
    /// a vertical line is a <c>line</c> preset in a zero-width, full-length box (no rotation,
    /// so it stays exactly at the anchor's X offset — rotation would pivot around the box
    /// center and shift the line). Round-trip note: the converter maps the <c>line</c> preset
    /// to a rectangle-kind shape, so vertical lines read back as a thin rect; horizontal
    /// <c>straightConnector1</c> lines round-trip as line-kind.
    /// </summary>
    public static OpenXmlUnknownElement BuildLine(
        uint id,
        string name,
        string? alt,
        double lengthPt,
        bool vertical,
        string strokeHex,
        double strokeWidthPt)
    {
        double cxPt = vertical ? 0.0 : lengthPt;
        double cyPt = vertical ? lengthPt : 0.0;
        OpenXmlUnknownElement spPr = Wps("spPr");
        spPr.Append(BuildTransform(cxPt, cyPt, rotation: 0.0));
        spPr.Append(BuildPresetGeometry(
            vertical ? A.ShapeTypeValues.Line : A.ShapeTypeValues.StraightConnector1,
            cornerAdj: null));
        spPr.Append(new A.NoFill());
        AppendStroke(spPr, strokeHex, strokeWidthPt);

        OpenXmlUnknownElement nonVisual = BuildNonVisualProps(id, name, alt, txBox: false);
        return Wps("wsp", nonVisual, spPr);
    }

    /// <summary>
    /// Builds the <c>p:pic</c> floating picture shell around a supplied image relationship
    /// id (<paramref name="embedId"/>). <paramref name="srcRect"/> carries an optional
    /// <c>a:srcRect</c> crop (values are 1/1000ths of a percent, 100000 = 100%).
    /// </summary>
    public static Pic.Picture BuildPicture(
        uint id,
        string name,
        string? alt,
        double widthPt,
        double heightPt,
        string embedId,
        (int Left, int Top, int Right, int Bottom)? srcRect)
    {
        // In the DOCX pic context the non-visual properties live in the picture namespace
        // (pic:cNvPr), unlike wps shapes which reuse DrawingML-main a:cNvPr.
        Pic.NonVisualDrawingProperties nonVisual = new() { Id = id, Name = name };
        if (!string.IsNullOrWhiteSpace(alt))
        {
            nonVisual.Description = alt;
        }

        Pic.NonVisualPictureProperties nonVisualPictureProperties =
            new(nonVisual, new Pic.NonVisualPictureDrawingProperties());
        // Word writes the (schema-optional here) application non-visual properties node;
        // the SDK ships no typed class for pic:nvPr, so it is appended as an empty node.
        nonVisualPictureProperties.Append(new OpenXmlUnknownElement("pic", "nvPr", PictureNamespace));

        Pic.BlipFill blipFill = new(new A.Blip { Embed = embedId });
        if (srcRect is { } crop)
        {
            blipFill.Append(new A.SourceRectangle
            {
                Left = crop.Left,
                Top = crop.Top,
                Right = crop.Right,
                Bottom = crop.Bottom
            });
        }

        blipFill.Append(new A.Stretch(new A.FillRectangle()));

        Pic.ShapeProperties shapeProperties = new(
            new A.Transform2D(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = PtToEmu(widthPt), Cy = PtToEmu(heightPt) }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
            new A.NoFill(),
            new A.Outline(new A.NoFill()));

        return new Pic.Picture(nonVisualPictureProperties, blipFill, shapeProperties);
    }

    private static OpenXmlUnknownElement BuildBoxShape(
        uint id,
        string name,
        string? alt,
        double widthPt,
        double heightPt,
        double rotation,
        double cornerRadiusPt,
        string? fillHex,
        string? strokeHex,
        double strokeWidthPt,
        bool txBox)
    {
        (A.ShapeTypeValues preset, double? cornerAdj) = cornerRadiusPt > 0
            ? (A.ShapeTypeValues.RoundRectangle, CornerRadiusToAdjustment(cornerRadiusPt, widthPt, heightPt))
            : (A.ShapeTypeValues.Rectangle, (double?)null);

        OpenXmlUnknownElement spPr = Wps("spPr");
        spPr.Append(BuildTransform(widthPt, heightPt, rotation));
        spPr.Append(BuildPresetGeometry(preset, cornerAdj));
        AppendFill(spPr, fillHex);
        AppendStroke(spPr, strokeHex, strokeWidthPt);

        OpenXmlUnknownElement nonVisual = BuildNonVisualProps(id, name, alt, txBox);
        return Wps("wsp", nonVisual, spPr);
    }

    private static OpenXmlUnknownElement BuildNonVisualProps(uint id, string name, string? alt, bool txBox)
    {
        A.NonVisualDrawingProperties nonVisualDrawingProperties = new() { Id = id, Name = name };
        if (!string.IsNullOrWhiteSpace(alt))
        {
            nonVisualDrawingProperties.Description = alt;
        }

        A.NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties = new();
        if (txBox)
        {
            nonVisualShapeDrawingProperties.TextBox = true;
        }

        return Wps("cNvSpPr", nonVisualDrawingProperties, nonVisualShapeDrawingProperties);
    }

    /// <summary>
    /// The <c>wps:bodyPr</c> text-body properties: Word's default insets, top vertical
    /// anchor (the vocabulary models no vertical alignment — documented default), square
    /// text wrap and no autofit for deterministic sizing.
    /// </summary>
    private static OpenXmlUnknownElement BuildTextBodyProperties()
    {
        OpenXmlUnknownElement bodyProperties = Wps("bodyPr");
        SetAttribute(bodyProperties, "wrap", "square");
        SetAttribute(bodyProperties, "lIns", Invariant(DefaultLeftRightInsetEmu));
        SetAttribute(bodyProperties, "tIns", Invariant(DefaultTopBottomInsetEmu));
        SetAttribute(bodyProperties, "rIns", Invariant(DefaultLeftRightInsetEmu));
        SetAttribute(bodyProperties, "bIns", Invariant(DefaultTopBottomInsetEmu));
        SetAttribute(bodyProperties, "anchor", "t");
        bodyProperties.Append(new A.NoAutoFit());
        return bodyProperties;
    }

    private static A.Transform2D BuildTransform(double widthPt, double heightPt, double rotation)
    {
        A.Transform2D transform = new();
        if (rotation != 0.0)
        {
            transform.Rotation = NormalizeAngle(rotation);
        }

        transform.Append(new A.Offset { X = 0, Y = 0 });
        transform.Append(new A.Extents { Cx = PtToEmu(widthPt), Cy = PtToEmu(heightPt) });
        return transform;
    }

    private static A.PresetGeometry BuildPresetGeometry(A.ShapeTypeValues preset, double? cornerAdj)
    {
        A.AdjustValueList adjustValueList = new();
        if (cornerAdj is { } adjustment)
        {
            adjustValueList.Append(new A.ShapeGuide { Name = "adj", Formula = $"val {Invariant((int)adjustment)}" });
        }

        return new A.PresetGeometry(adjustValueList) { Preset = preset };
    }

    private static void AppendFill(OpenXmlCompositeElement shapeProperties, string? fillHex)
    {
        if (fillHex is null)
        {
            shapeProperties.Append(new A.NoFill());
            return;
        }

        shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = fillHex }));
    }

    private static void AppendStroke(OpenXmlCompositeElement shapeProperties, string? strokeHex, double strokeWidthPt)
    {
        if (strokeHex is null)
        {
            shapeProperties.Append(new A.Outline(new A.NoFill()));
            return;
        }

        shapeProperties.Append(new A.Outline(
            new A.SolidFill(new A.RgbColorModelHex { Val = strokeHex }))
        {
            Width = PtToEmu(strokeWidthPt)
        });
    }

    /// <summary>
    /// Converts a corner radius in points to the <c>roundRect</c> <c>adj</c> adjustment
    /// value: a percentage of the smaller box dimension in 1/1000ths of a percent
    /// (100000 = 100%), clamped to 50%.
    /// </summary>
    private static double CornerRadiusToAdjustment(double cornerRadiusPt, double widthPt, double heightPt)
    {
        double smaller = Math.Min(widthPt, heightPt);
        if (smaller <= 0.0)
        {
            return 0.0;
        }

        double fraction = cornerRadiusPt / smaller;
        return Math.Clamp(fraction, 0.0, 0.5) * 100000.0;
    }

    private static OpenXmlUnknownElement Wps(string localName, params OpenXmlElement[] children)
    {
        OpenXmlUnknownElement element = new("wps", localName, WpsNamespace);
        if (children.Length > 0)
        {
            element.Append(children);
        }

        return element;
    }

    private static void SetAttribute(OpenXmlElement element, string localName, string value)
        => element.SetAttribute(new OpenXmlAttribute(localName, string.Empty, value));

    private static int NormalizeAngle(double degrees)
    {
        int angle = (int)Math.Round(degrees * AngleScale, MidpointRounding.AwayFromZero) % FullAngle;
        return angle < 0 ? angle + FullAngle : angle;
    }

    private static int PtToEmu(double points) =>
        (int)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero);

    private static string Invariant(int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
