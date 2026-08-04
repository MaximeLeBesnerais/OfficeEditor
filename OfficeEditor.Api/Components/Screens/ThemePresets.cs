namespace OfficeEditor.Api.Components.Screens;

using System.Text.Json;
using System.Text.Json.Nodes;

public sealed class ThemePreset
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public required ThemeDesign Design { get; init; }
}

public sealed class ThemeDesign
{
    public required PaletteColors Palette { get; set; }
    public FontsConfig Fonts { get; set; } = new("Avenir Next", "Helvetica Neue");
    public ShapeConfig Shape { get; set; } = new(0, "outline");
    public MetricsConfig Metrics { get; set; } = new(48, 18, 42, 16);

    public JsonObject ToJsonObject()
    {
        return new JsonObject
        {
            ["palette"] = new JsonObject
            {
                ["primary"] = Palette.Primary,
                ["accent"] = Palette.Accent,
                ["ink"] = Palette.Ink,
                ["paper"] = Palette.Paper,
                ["muted"] = Palette.Muted,
                ["mist"] = Palette.Mist,
                ["glow"] = Palette.Glow,
            },
            ["fonts"] = new JsonObject
            {
                ["display"] = Fonts.Display,
                ["body"] = Fonts.Body,
            },
            ["shape"] = new JsonObject
            {
                ["cornerRadius"] = Shape.CornerRadius,
                ["cardStyle"] = Shape.CardStyle,
            },
            ["metrics"] = new JsonObject
            {
                ["marginPt"] = Metrics.MarginPt,
                ["gutterPt"] = Metrics.GutterPt,
                ["titleSizePt"] = Metrics.TitleSizePt,
                ["bodySizePt"] = Metrics.BodySizePt,
            },
        };
    }
}

public sealed class PaletteColors
{
    public string Primary { get; set; }
    public string Accent { get; set; }
    public string Ink { get; set; }
    public string Paper { get; set; }
    public string Muted { get; set; }
    public string Mist { get; set; }
    public string Glow { get; set; }

    public PaletteColors(string primary, string accent, string ink, string paper, string muted, string mist, string glow)
    {
        Primary = primary; Accent = accent; Ink = ink; Paper = paper; Muted = muted; Mist = mist; Glow = glow;
    }
}

public sealed record FontsConfig(string Display, string Body);
public sealed record ShapeConfig(int CornerRadius, string CardStyle);
public sealed record MetricsConfig(int MarginPt, int GutterPt, int TitleSizePt, int BodySizePt);

public static class ThemePresets
{
    public const string DefaultThemeId = "northwind";

    public static readonly IReadOnlyList<ThemePreset> All =
    [
        new ThemePreset
        {
            Id = "northwind",
            Name = "Northwind Teal",
            Design = new ThemeDesign
            {
                Palette = new PaletteColors("#0E7C7B", "#D9A441", "#1E2A32", "#FFFFFF", "#7A8288", "#EDF2F2", "#35B5B3"),
            },
        },
        new ThemePreset
        {
            Id = "corporate-blue",
            Name = "Corporate Blue",
            Design = new ThemeDesign
            {
                Palette = new PaletteColors("#1C5C9E", "#B07E28", "#17263E", "#FFFFFF", "#6E7B8A", "#EDF1F7", "#6FA8DC"),
            },
        },
    ];
}
