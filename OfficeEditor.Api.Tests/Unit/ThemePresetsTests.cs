using System.Text.Json.Nodes;
using OfficeEditor.Api.Components.Screens;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Theme presets backing the Generate screen. This is the one piece of the Blazor
/// surface that is plain C# (no Razor markup, no JS interop), so it is exercised
/// directly here instead of pulling in a bUnit/rendering harness.
/// </summary>
public sealed class ThemePresetsTests
{
    [Fact]
    public void DefaultThemeId_IsNorthwind()
    {
        Assert.Equal("northwind", ThemePresets.DefaultThemeId);
    }

    [Fact]
    public void All_ReturnsExactlyTheThreePresetsInCatalogOrder()
    {
        var presets = ThemePresets.All;

        Assert.Equal(
            ["northwind", "corporate-blue", "heritage"],
            presets.Select(p => p.Id));
        Assert.Equal(
            ["Northwind Teal", "Corporate Blue", " Heritage"],
            presets.Select(p => p.Name));
        Assert.All(presets, p =>
        {
            Assert.False(string.IsNullOrWhiteSpace(p.Id));
            Assert.False(string.IsNullOrWhiteSpace(p.Name));
            Assert.NotNull(p.Design.Palette);
        });
    }

    [Fact]
    public void All_PresetIdsAreUnique()
    {
        Assert.Equal(
            ThemePresets.All.Count,
            ThemePresets.All.Select(p => p.Id).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void All_NorthwindIsFirstAndIsTheDefault()
    {
        Assert.Equal(ThemePresets.DefaultThemeId, ThemePresets.All[0].Id);
    }

    [Fact]
    public void ToJsonObject_DefaultDesign_EmitsExpectedShape()
    {
        var design = new ThemeDesign
        {
            Palette = new PaletteColors("#0E7C7B", "#D9A441", "#1E2A32", "#FFFFFF", "#7A8288", "#EDF2F2", "#35B5B3"),
        };

        var json = design.ToJsonObject();

        var palette = json["palette"]!.AsObject();
        Assert.Equal("#0E7C7B", palette["primary"]!.GetValue<string>());
        Assert.Equal("#D9A441", palette["accent"]!.GetValue<string>());
        Assert.Equal("#1E2A32", palette["ink"]!.GetValue<string>());
        Assert.Equal("#FFFFFF", palette["paper"]!.GetValue<string>());
        Assert.Equal("#7A8288", palette["muted"]!.GetValue<string>());
        Assert.Equal("#EDF2F2", palette["mist"]!.GetValue<string>());
        Assert.Equal("#35B5B3", palette["glow"]!.GetValue<string>());

        var fonts = json["fonts"]!.AsObject();
        Assert.Equal("Avenir Next", fonts["display"]!.GetValue<string>());
        Assert.Equal("Helvetica Neue", fonts["body"]!.GetValue<string>());

        var shape = json["shape"]!.AsObject();
        Assert.Equal(0, shape["cornerRadius"]!.GetValue<int>());
        Assert.Equal("outline", shape["cardStyle"]!.GetValue<string>());

        var metrics = json["metrics"]!.AsObject();
        Assert.Equal(48, metrics["marginPt"]!.GetValue<int>());
        Assert.Equal(18, metrics["gutterPt"]!.GetValue<int>());
        Assert.Equal(42, metrics["titleSizePt"]!.GetValue<int>());
        Assert.Equal(16, metrics["bodySizePt"]!.GetValue<int>());
    }

    [Fact]
    public void ToJsonObject_OverriddenDesign_EmitsOverrides()
    {
        var design = new ThemeDesign
        {
            Palette = new PaletteColors("#111111", "#222222", "#333333", "#444444", "#555555", "#666666", "#777777"),
            Fonts = new("Inter", "Roboto"),
            Shape = new(12, "filled"),
            Metrics = new(60, 24, 50, 20),
        };

        var json = design.ToJsonObject();

        Assert.Equal("Inter", json["fonts"]!["display"]!.GetValue<string>());
        Assert.Equal("Roboto", json["fonts"]!["body"]!.GetValue<string>());
        Assert.Equal(12, json["shape"]!["cornerRadius"]!.GetValue<int>());
        Assert.Equal("filled", json["shape"]!["cardStyle"]!.GetValue<string>());
        Assert.Equal(60, json["metrics"]!["marginPt"]!.GetValue<int>());
        Assert.Equal(24, json["metrics"]!["gutterPt"]!.GetValue<int>());
        Assert.Equal(50, json["metrics"]!["titleSizePt"]!.GetValue<int>());
        Assert.Equal(20, json["metrics"]!["bodySizePt"]!.GetValue<int>());
    }

    [Fact]
    public void NorthwindPreset_ToJsonObject_MatchesThePresetPalette()
    {
        var northwind = ThemePresets.All.Single(p => p.Id == ThemePresets.DefaultThemeId);

        var json = northwind.Design.ToJsonObject();

        Assert.Equal("#0E7C7B", json["palette"]!["primary"]!.GetValue<string>());
        Assert.Equal("#35B5B3", json["palette"]!["glow"]!.GetValue<string>());
    }
}
