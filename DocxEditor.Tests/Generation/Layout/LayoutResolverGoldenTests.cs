using System.Text.Json;
using System.Text.Json.Nodes;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation.Layout;

/// <summary>
/// Golden-file tests: a generation document in → the absolute draw tree out,
/// compared against committed JSON snapshots. Set OE_UPDATE_SNAPSHOTS=1 to regenerate the
/// expected files (review the diff before committing).
/// </summary>
public sealed class LayoutResolverGoldenTests
{
    private const string UpdateSnapshotsEnvVar = "OE_UPDATE_SNAPSHOTS";

    [Theory]
    [InlineData("row-grow")]
    [InlineData("justify")]
    [InlineData("align")]
    [InlineData("padding")]
    [InlineData("aspect")]
    [InlineData("grid")]
    [InlineData("nested")]
    [InlineData("absolute")]
    [InlineData("tokens")]
    public void Resolve_GoldenFixture_MatchesExpectedDrawTree(string name)
    {
        var inputPath = Path.Combine(ResolveGoldenDirectory(), name + ".input.json");
        Assert.True(File.Exists(inputPath), $"Golden input not found: {inputPath}");

        var document = new GenerationDocumentParser().Parse(File.ReadAllText(inputPath));
        var result = new LayoutResolver().Resolve(document);
        var actual = DrawTreeSerializer.Serialize(result);

        var expectedPath = Path.Combine(ResolveGoldenDirectory(), name + ".expected.json");
        if (Environment.GetEnvironmentVariable(UpdateSnapshotsEnvVar) == "1")
        {
            File.WriteAllText(expectedPath, actual);
        }

        Assert.True(File.Exists(expectedPath),
            $"Golden snapshot not found: {expectedPath}. Run the test with {UpdateSnapshotsEnvVar}=1 to create it.");

        var expected = File.ReadAllText(expectedPath).Replace("\r\n", "\n");
        Assert.Equal(expected, actual);
    }

    private static string ResolveGoldenDirectory()
    {
        var baseDir = AppContext.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "Generation", "Layout", "Golden"));
    }
}

/// <summary>
/// Deterministic JSON shape for draw trees: camelCase keys in insertion order, a "kind"
/// discriminator per element, nulls omitted, trailing newline. Geometry is already rounded
/// to 3 decimals by the resolver.
/// </summary>
internal static class DrawTreeSerializer
{
    public static string Serialize(LayoutResult result)
    {
        var root = new JsonObject
        {
            ["slides"] = new JsonArray(result.Slides.Select(SerializeSlide).ToArray<JsonNode?>()),
            ["warnings"] = new JsonArray(result.Warnings.Select(w => (JsonNode?)JsonValue.Create(w)).ToArray())
        };
        return root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) + "\n";
    }

    private static JsonObject SerializeSlide(ResolvedSlide slide)
        => new()
        {
            ["widthPt"] = slide.WidthPt,
            ["heightPt"] = slide.HeightPt,
            ["root"] = SerializeElement(slide.Root)
        };

    private static JsonObject SerializeElement(ResolvedElement element)
    {
        var node = new JsonObject
        {
            ["kind"] = element switch
            {
                ResolvedContainer => "container",
                ResolvedText => "text",
                ResolvedRect => "rect",
                ResolvedEllipse => "ellipse",
                ResolvedLine => "line",
                ResolvedImage => "image",
                ResolvedGroup => "group",
                _ => throw new NotSupportedException($"Unknown resolved element '{element.GetType().Name}'.")
            },
            ["x"] = element.X,
            ["y"] = element.Y,
            ["w"] = element.Width,
            ["h"] = element.Height
        };

        switch (element)
        {
            case ResolvedContainer c:
                AddIfNotNull(node, "fill", SerializeFill(c.Fill));
                AddIfNotNull(node, "stroke", SerializeStroke(c.Stroke));
                AddIfNotNull(node, "radius", SerializeRadius(c.Radius));
                AddIfNotNull(node, "shadow", SerializeShadow(c.Shadow));
                node["overflow"] = SerializeOverflow(c.Overflow);
                node["children"] = new JsonArray(c.Children.Select(SerializeElement).ToArray<JsonNode?>());
                break;
            case ResolvedText t:
                node["runs"] = new JsonArray(t.Runs.Select(SerializeRun).ToArray<JsonNode?>());
                node["textAlign"] = t.TextAlign.ToString().ToLowerInvariant();
                node["anchor"] = t.Anchor.ToString().ToLowerInvariant();
                node["insets"] = SerializeInsets(t.Insets);
                node["overflow"] = SerializeOverflow(t.Overflow);
                node["fontScale"] = t.FontScale;
                AddIfNotNull(node, "shadow", SerializeShadow(t.Shadow));
                break;
            case ResolvedRect r:
                AddIfNotNull(node, "fill", SerializeFill(r.Fill));
                AddIfNotNull(node, "stroke", SerializeStroke(r.Stroke));
                AddIfNotNull(node, "radius", SerializeRadius(r.Radius));
                AddIfNotNull(node, "shadow", SerializeShadow(r.Shadow));
                break;
            case ResolvedEllipse e:
                AddIfNotNull(node, "fill", SerializeFill(e.Fill));
                AddIfNotNull(node, "stroke", SerializeStroke(e.Stroke));
                AddIfNotNull(node, "shadow", SerializeShadow(e.Shadow));
                break;
            case ResolvedLine l:
                node["connector"] = l.IsConnector;
                node["orientation"] = l.Orientation.ToString().ToLowerInvariant();
                AddIfNotNull(node, "stroke", SerializeStroke(l.Stroke));
                break;
            case ResolvedImage im:
                node["src"] = im.Source;
                node["fit"] = im.Fit.ToString().ToLowerInvariant();
                if (im.Crop is { } crop)
                {
                    node["crop"] = new JsonArray(crop.Left, crop.Top, crop.Right, crop.Bottom);
                }
                if (im.Alt is not null)
                {
                    node["alt"] = im.Alt;
                }
                break;
            case ResolvedGroup g:
                node["children"] = new JsonArray(g.Children.Select(SerializeElement).ToArray<JsonNode?>());
                break;
        }
        return node;
    }

    private static JsonObject SerializeRun(ResolvedTextRun run)
    {
        var node = new JsonObject { ["text"] = run.Text };
        if (run.FontFamily is not null)
        {
            node["font"] = run.FontFamily;
        }
        node["size"] = run.FontSizePt;
        if (run.ColorHex is not null)
        {
            node["color"] = run.ColorHex;
        }
        node["bold"] = run.Bold;
        node["italic"] = run.Italic;
        return node;
    }

    private static JsonObject? SerializeFill(FillSpec? fill)
        => fill switch
        {
            null => null,
            SolidFill solid => new JsonObject { ["type"] = "solid", ["color"] = solid.Color },
            LinearGradientFill gradient => new JsonObject
            {
                ["type"] = "gradient",
                ["angle"] = gradient.Angle,
                ["stops"] = new JsonArray(gradient.Stops.Select(s =>
                {
                    var stop = new JsonObject { ["color"] = s.Color, ["offset"] = s.Offset };
                    if (s.Alpha is { } alpha)
                    {
                        stop["alpha"] = alpha;
                    }
                    return (JsonNode?)stop;
                }).ToArray())
            },
            _ => throw new NotSupportedException($"Unknown fill '{fill.GetType().Name}'.")
        };

    private static JsonObject? SerializeStroke(StrokeSpec? stroke)
        => stroke is null ? null : new JsonObject { ["color"] = stroke.Color, ["width"] = stroke.WidthPt };

    private static JsonObject? SerializeShadow(ShadowSpec? shadow)
    {
        if (shadow is null)
        {
            return null;
        }
        var node = new JsonObject
        {
            ["color"] = shadow.Color,
            ["dx"] = shadow.Dx,
            ["dy"] = shadow.Dy,
            ["blur"] = shadow.Blur
        };
        if (shadow.Alpha is { } alpha)
        {
            node["alpha"] = alpha;
        }
        return node;
    }

    private static JsonArray? SerializeRadius(CornerRadii? radius)
        => radius is { } r ? new JsonArray(r.TopLeft, r.TopRight, r.BottomRight, r.BottomLeft) : null;

    private static JsonArray SerializeInsets(EdgeInsets insets)
        => new(insets.Top, insets.Right, insets.Bottom, insets.Left);

    private static string SerializeOverflow(OverflowPolicy overflow)
        => overflow.ToString().ToLowerInvariant();

    private static void AddIfNotNull(JsonObject node, string key, JsonNode? value)
    {
        if (value is not null)
        {
            node[key] = value;
        }
    }
}
