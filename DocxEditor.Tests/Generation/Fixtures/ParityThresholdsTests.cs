using PptxEditor.Core.Generation.Fixtures;

namespace DocxEditor.Tests.Generation.Fixtures;

/// <summary>
/// Thresholds-file loading: a missing file is an operational error with instructions
/// (never a silent pass), malformed files fail loudly, and the serialized catalog form
/// round-trips through the loader.
/// </summary>
public sealed class ParityThresholdsTests : IDisposable
{
    private readonly string _testDir = Path.Combine(Path.GetTempPath(), $"parity-threshold-tests-{Guid.NewGuid():N}");

    public void Dispose()
    {
        try { Directory.Delete(_testDir, recursive: true); } catch { /* best effort */ }
    }

    private string Write(string name, string content)
    {
        Directory.CreateDirectory(_testDir);
        var path = Path.Combine(_testDir, name);
        File.WriteAllText(path, content);
        return path;
    }

    [Fact]
    public void Load_MissingFile_ThrowsWithInstructions()
    {
        var ex = Assert.Throws<FileNotFoundException>(() =>
            ParityThresholds.Load(Path.Combine(_testDir, "nope.json")));
        Assert.Contains("tools/visual-diff/baselines/gen/thresholds.json", ex.Message);
    }

    [Fact]
    public void Load_MalformedJson_Throws()
    {
        var path = Write("bad.json", "{ not json");
        Assert.Throws<InvalidOperationException>(() => ParityThresholds.Load(path));
    }

    [Fact]
    public void Load_WrongVersion_Throws()
    {
        var path = Write("v0.json", """{ "Version": 99, "Thresholds": { "a": 0.1 } }""");
        var ex = Assert.Throws<InvalidOperationException>(() => ParityThresholds.Load(path));
        Assert.Contains("version", ex.Message);
    }

    [Fact]
    public void Load_MissingThresholdsObject_Throws()
    {
        var path = Write("empty.json", """{ "Version": 1 }""");
        Assert.Throws<InvalidOperationException>(() => ParityThresholds.Load(path));
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-0.5)]
    [InlineData(1.5)]
    public void Load_OutOfRangeThreshold_Throws(double value)
    {
        var path = Write("range.json", $$"""{ "Version": 1, "Thresholds": { "a": {{value}} } }""");
        Assert.Throws<InvalidOperationException>(() => ParityThresholds.Load(path));
    }

    [Fact]
    public void Load_ValidFile_ReturnsThresholds()
    {
        var path = Write("ok.json", """{ "Version": 1, "Thresholds": { "rect-radii": 0.04, "shadow": 0.12 } }""");
        var thresholds = ParityThresholds.Load(path);
        Assert.Equal(0.04, thresholds["rect-radii"]);
        Assert.Equal(0.12, thresholds["shadow"]);
    }

    [Fact]
    public void SerializeCatalog_RoundTripsThroughLoad()
    {
        var path = Write("catalog.json", ParityThresholds.SerializeCatalog());
        var thresholds = ParityThresholds.Load(path);
        var expected = FixtureCatalog.All.ToDictionary(f => f.Name, f => f.ThresholdRmse, StringComparer.Ordinal);
        Assert.Equal(expected, thresholds);
    }

    [Fact]
    public void CommittedFile_Loads_CoveringEveryFixture()
    {
        var path = Path.Combine(FixtureCatalogTests.ResolveRepositoryRoot(),
            "tools", "visual-diff", "baselines", "gen", "thresholds.json");
        var thresholds = ParityThresholds.Load(path);
        foreach (var fixture in FixtureCatalog.All)
        {
            Assert.True(thresholds.ContainsKey(fixture.Name), $"thresholds.json is missing fixture '{fixture.Name}'.");
        }
    }
}
