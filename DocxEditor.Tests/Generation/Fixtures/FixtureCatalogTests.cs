using PptxEditor.Core.Generation.Fixtures;

namespace DocxEditor.Tests.Generation.Fixtures;

/// <summary>
/// Fixture catalog acceptance (plan.md §2 rule 2): every Tier-1 primitive + linear
/// gradient has exactly one parity fixture with a per-primitive RMSE threshold, and the
/// committed thresholds file (the CI-facing record) never drifts from the catalog (the
/// single source of truth).
/// </summary>
public sealed class FixtureCatalogTests
{
    private static readonly string[] ExpectedFixtures =
    [
        "rect-radii", "gradient", "text-anchors", "line-connector",
        "ellipse", "image-fit", "group", "shadow"
    ];

    [Fact]
    public void Catalog_CoversEveryTier1Primitive_PlusLinearGradient()
    {
        Assert.Equal(ExpectedFixtures, FixtureCatalog.All.Select(f => f.Name).ToArray());
    }

    [Fact]
    public void Catalog_Names_AreUniqueSlugs()
    {
        foreach (var fixture in FixtureCatalog.All)
        {
            Assert.Matches("^[a-z0-9]+(-[a-z0-9]+)*$", fixture.Name);
        }
        Assert.Equal(FixtureCatalog.All.Count, FixtureCatalog.All.Select(f => f.Name).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void Catalog_Thresholds_AreValidNormalizedRmse()
    {
        foreach (var fixture in FixtureCatalog.All)
        {
            Assert.True(fixture.ThresholdRmse is > 0 and <= 1,
                $"fixture '{fixture.Name}' threshold {fixture.ThresholdRmse} must be a normalized RMSE in (0, 1].");
        }
    }

    [Fact]
    public void Catalog_Documents_UseSupportedVersion()
    {
        foreach (var fixture in FixtureCatalog.All)
        {
            Assert.Equal("2.0", fixture.BuildDocument("unused.png").Version);
        }
    }

    [Fact]
    public void Catalog_Find_ResolvesByName()
    {
        Assert.NotNull(FixtureCatalog.Find("rect-radii"));
        Assert.Null(FixtureCatalog.Find("does-not-exist"));
    }

    [Fact]
    public void CommittedThresholdsFile_MatchesCatalog()
    {
        var path = Path.Combine(ResolveRepositoryRoot(), "tools", "visual-diff", "baselines", "gen", "thresholds.json");
        Assert.True(File.Exists(path), $"Committed thresholds file not found: {path}");

        if (Environment.GetEnvironmentVariable("OE_UPDATE_SNAPSHOTS") == "1")
        {
            File.WriteAllText(path, ParityThresholds.SerializeCatalog() + Environment.NewLine);
        }

        var committed = ParityThresholds.Load(path);
        var expected = FixtureCatalog.All.ToDictionary(f => f.Name, f => f.ThresholdRmse, StringComparer.Ordinal);
        Assert.Equal(expected, committed);
    }

    internal static string ResolveRepositoryRoot() =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
}
