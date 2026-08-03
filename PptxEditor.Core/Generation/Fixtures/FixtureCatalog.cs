using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace PptxEditor.Core.Generation.Fixtures;

/// <summary>
/// The parity fixture set: one generated deck per Tier-1
/// primitive plus linear gradient, each with its per-primitive RMSE threshold.
/// The Typst render is the spec of record — OOXML is matched TO the preview.
/// Thresholds are mirrored in tools/visual-diff/baselines/gen/thresholds.json; a drift
/// guard test keeps the two in sync.
/// </summary>
public static class FixtureCatalog
{
    /// <summary>Vocabulary version emitted by every fixture.</summary>
    public const string DocumentVersion = "2.0";

    /// <summary>All fixtures, in stable catalog order.</summary>
    public static IReadOnlyList<ParityFixture> All { get; } =
    [
        new ParityFixture
        {
            Name = "rect-radii",
            Description = "Rectangle per-corner radius matrix: square, uniform radii, round1Rect (one corner distinct), round2SameRect (top pair), round2DiagRect (diagonals) — OOXML adj values ↔ Typst rect radius corners.",
            ThresholdRmse = 0.04,
            BuildDocument = _ => Document(
            [
                Rect(40, 50, 260, 180, new SolidFill("primary")),
                Rect(350, 50, 260, 180, new SolidFill("accent"), radius: CornerRadii.All(12)),
                Rect(660, 50, 260, 180, new SolidFill("muted"), radius: CornerRadii.All(40)),
                Rect(40, 290, 260, 180, new SolidFill("primary"), stroke: InkStroke(2), radius: new CornerRadii(24, 24, 0, 0)),
                Rect(350, 290, 260, 180, new SolidFill("accent"), stroke: InkStroke(2), radius: new CornerRadii(12, 32, 12, 12)),
                Rect(660, 290, 260, 180, new SolidFill("muted"), stroke: InkStroke(2), radius: new CornerRadii(24, 0, 24, 0))
            ])
        },
        new ParityFixture
        {
            Name = "gradient",
            Description = "Linear gradient matrix: axis angles 0/45/90/180, two- and three-stop, per-stop alpha (a:lin ang + gsLst ↔ gradient.linear).",
            ThresholdRmse = 0.05,
            BuildDocument = _ => Document(
            [
                Rect(40, 40, 420, 220, Gradient(0, [Stop("primary", 0), Stop("accent", 1)])),
                Rect(500, 40, 420, 220, Gradient(45, [Stop("primary", 0), Stop("accent", 0.5), Stop("ink", 1)])),
                Rect(40, 280, 420, 220, Gradient(90, [Stop("muted", 0, 0.6), Stop("primary", 1)])),
                Rect(500, 280, 420, 220, Gradient(180, [Stop("accent", 0, 0.9), Stop("primary", 1)]))
            ])
        },
        new ParityFixture
        {
            Name = "text-anchors",
            Description = "Text anchor × align 3×3 matrix on tinted backing rects, with a mixed-style runs cell (bold/italic/color/size) at center.",
            ThresholdRmse = 0.08,
            BuildDocument = _ => Document(TextAnchorCells())
        },
        new ParityFixture
        {
            Name = "line-connector",
            Description = "Straight lines and connectors, horizontal and vertical, varying stroke widths/colors, plus the default-stroke line (1pt black).",
            ThresholdRmse = 0.04,
            BuildDocument = _ => Document(
            [
                new LineElement { At = Pt(80, 100), Size = WH(800, 4), Stroke = new StrokeSpec { Color = "accent", WidthPt = 2 } },
                new LineElement { At = Pt(80, 140), Size = WH(800, 4), Stroke = new StrokeSpec { Color = "ink", WidthPt = 1 } },
                new LineElement { At = Pt(80, 180), Size = WH(800, 4), IsConnector = true, Stroke = new StrokeSpec { Color = "primary", WidthPt = 4 } },
                new LineElement { At = Pt(120, 220), Size = WH(4, 260), Orientation = LineOrientation.Vertical, Stroke = new StrokeSpec { Color = "accent", WidthPt = 2 } },
                new LineElement { At = Pt(200, 220), Size = WH(4, 260), Orientation = LineOrientation.Vertical, IsConnector = true, Stroke = new StrokeSpec { Color = "primary", WidthPt = 4 } },
                new LineElement { At = Pt(80, 500), Size = WH(800, 4) }
            ])
        },
        new ParityFixture
        {
            Name = "ellipse",
            Description = "Ellipse and circle with fill-only, stroke-only and fill+stroke treatments.",
            ThresholdRmse = 0.04,
            BuildDocument = _ => Document(
            [
                new EllipseElement { At = Pt(60, 80), Size = WH(300, 160), Fill = new SolidFill("primary") },
                new EllipseElement { At = Pt(420, 80), Size = WH(160, 160), Fill = new SolidFill("accent") },
                new EllipseElement { At = Pt(640, 80), Size = WH(260, 160), Stroke = InkStroke(3) },
                new EllipseElement { At = Pt(60, 300), Size = WH(400, 180), Fill = new SolidFill("muted"), Stroke = InkStroke(1) }
            ])
        },
        new ParityFixture
        {
            Name = "image-fit",
            Description = "Image fit modes (F7): fill on a matching-aspect frame (control), fill cropping into 1:1, contain letterboxing into 2:3, explicit crop srcRect (10% left/right).",
            ThresholdRmse = 0.06,
            BuildDocument = imageSource => Document(
            [
                new ImageElement { At = Pt(40, 60), Size = WH(300, 200), Source = imageSource, Fit = ImageFitMode.Fill, Alt = "checkerboard control" },
                new ImageElement { At = Pt(380, 60), Size = WH(300, 300), Source = imageSource, Fit = ImageFitMode.Fill },
                new ImageElement { At = Pt(720, 60), Size = WH(200, 300), Source = imageSource, Fit = ImageFitMode.Contain },
                new ImageElement { At = Pt(40, 400), Size = WH(400, 120), Source = imageSource, Fit = ImageFitMode.Crop, Crop = new SourceRect(10000, 0, 10000, 0) }
            ])
        },
        new ParityFixture
        {
            Name = "group",
            Description = "Nested groups with children placed in group coordinates; overlapping shapes demonstrate paint order = document order.",
            ThresholdRmse = 0.05,
            BuildDocument = _ => Document(
            [
                new GroupElement
                {
                    At = Pt(60, 60),
                    Size = WH(380, 420),
                    Children =
                    [
                        Rect(0, 0, 380, 420, new SolidFill("wash"), stroke: new StrokeSpec { Color = "muted", WidthPt = 1 }),
                        Rect(20, 20, 160, 160, new SolidFill("primary")),
                        new EllipseElement { At = Pt(120, 60), Size = WH(160, 160), Fill = new SolidFill("accent") },
                        Text(20, 220, 340, 120, "Group A", fontSize: 20)
                    ]
                },
                new GroupElement
                {
                    At = Pt(500, 60),
                    Size = WH(380, 420),
                    Children =
                    [
                        Rect(0, 0, 380, 420, new SolidFill("wash"), stroke: new StrokeSpec { Color = "muted", WidthPt = 1 }),
                        new GroupElement
                        {
                            At = Pt(40, 40),
                            Size = WH(300, 200),
                            Children =
                            [
                                Rect(0, 0, 300, 200, new SolidFill("primary")),
                                Text(16, 60, 268, 80, "Nested group", color: "paper", fontSize: 18)
                            ]
                        },
                        Rect(20, 260, 200, 140, new SolidFill("accent")),
                        Rect(100, 300, 200, 140, new SolidFill("primary"))
                    ]
                }
            ])
        },
        new ParityFixture
        {
            Name = "shadow",
            Description = "Drop shadows on rect, rounded container and text (native a:effectLst/outerShdw ↔ faked offset copy in the Typst preview).",
            ThresholdRmse = 0.12,
            BuildDocument = _ => Document(
            [
                Rect(80, 80, 300, 180, new SolidFill("paper"), stroke: new StrokeSpec { Color = "muted", WidthPt = 1 },
                    shadow: new ShadowSpec { Color = "ink", Dx = 6, Dy = 6, Blur = 8, Alpha = 0.35 }),
                new ContainerElement
                {
                    At = Pt(460, 80),
                    Size = WH(300, 180),
                    Fill = new SolidFill("wash"),
                    Radius = CornerRadii.All(12),
                    Shadow = new ShadowSpec { Color = "primary", Dx = 0, Dy = 8, Blur = 16, Alpha = 0.3 },
                    Children = [Text(24, 60, 252, 60, "Card shadow", fontSize: 18)]
                },
                Text(80, 320, 400, 120, "Soft shadow text", bold: true, fontSize: 28,
                    shadow: new ShadowSpec { Color = "ink", Dx = 2, Dy = 2, Blur = 3, Alpha = 0.25 }),
                Rect(560, 320, 240, 140, new SolidFill("accent"),
                    shadow: new ShadowSpec { Color = "ink", Dx = -6, Dy = 6, Blur = 10, Alpha = 0.4 })
            ])
        }
    ];

    /// <summary>Looks a fixture up by name (ordinal); null when unknown.</summary>
    public static ParityFixture? Find(string name) =>
        All.FirstOrDefault(f => string.Equals(f.Name, name, StringComparison.Ordinal));

    #region Document helpers

    private static DesignTokens Tokens { get; } = new()
    {
        Palette = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["primary"] = "#0B3D91",
            ["accent"] = "#FF6B00",
            ["ink"] = "#1A1A1A",
            ["paper"] = "#FFFFFF",
            ["muted"] = "#8A94A6",
            ["wash"] = "#EEF1F6"
        },
        Fonts = new FontTokens { Display = "Arial", Body = "Arial" }
    };

    private static GenerationDocument Document(IReadOnlyList<GenElement> children) => new()
    {
        Version = DocumentVersion,
        Design = Tokens,
        Slides =
        [
            new ContainerElement
            {
                Fill = new SolidFill("paper"),
                Children = children
            }
        ]
    };

    private static IReadOnlyList<GenElement> TextAnchorCells()
    {
        TextAnchor[] anchors = [TextAnchor.Top, TextAnchor.Middle, TextAnchor.Bottom];
        TextAlign[] aligns = [TextAlign.Left, TextAlign.Center, TextAlign.Right];
        var cells = new List<GenElement>();
        for (int row = 0; row < anchors.Length; row++)
        {
            for (int col = 0; col < aligns.Length; col++)
            {
                double x = 30 + col * 310;
                double y = 30 + row * 170;
                cells.Add(Rect(x, y, 280, 120, new SolidFill("wash")));
                if (anchors[row] == TextAnchor.Middle && aligns[col] == TextAlign.Center)
                {
                    cells.Add(new TextElement
                    {
                        At = Pt(x, y),
                        Size = WH(280, 120),
                        Anchor = anchors[row],
                        TextAlign = aligns[col],
                        Insets = EdgeInsets.All(8),
                        Runs =
                        [
                            new TextRun { Text = "Bold ", Bold = true, Color = "primary", FontSize = 20 },
                            new TextRun { Text = "italic ", Italic = true, Color = "accent", FontSize = 16 },
                            new TextRun { Text = "plain", Color = "ink", FontSize = 12 }
                        ]
                    });
                }
                else
                {
                    cells.Add(Text(x, y, 280, 120, "Aa Qq 123", fontSize: 18, anchor: anchors[row], align: aligns[col]));
                }
            }
        }
        return cells;
    }

    private static PointSpec Pt(double x, double y) => new(x, y);

    private static SizeSpec WH(double width, double height) => new() { Width = width, Height = height };

    private static StrokeSpec InkStroke(double width) => new() { Color = "ink", WidthPt = width };

    private static LinearGradientFill Gradient(double angle, IReadOnlyList<GradientStop> stops) =>
        new() { Angle = angle, Stops = stops };

    private static GradientStop Stop(string color, double offset, double? alpha = null) =>
        new() { Color = color, Offset = offset, Alpha = alpha };

    private static RectElement Rect(
        double x, double y, double w, double h,
        FillSpec fill, StrokeSpec? stroke = null, CornerRadii? radius = null, ShadowSpec? shadow = null) =>
        new() { At = Pt(x, y), Size = WH(w, h), Fill = fill, Stroke = stroke, Radius = radius, Shadow = shadow };

    private static TextElement Text(
        double x, double y, double w, double h, string value,
        string color = "ink", double fontSize = 14, bool bold = false,
        TextAnchor anchor = TextAnchor.Top, TextAlign align = TextAlign.Left, ShadowSpec? shadow = null) =>
        new()
        {
            At = Pt(x, y),
            Size = WH(w, h),
            Value = value,
            Font = "body",
            FontSize = fontSize,
            Color = color,
            Bold = bold,
            Anchor = anchor,
            TextAlign = align,
            Insets = EdgeInsets.All(8),
            Shadow = shadow
        };

    #endregion
}
