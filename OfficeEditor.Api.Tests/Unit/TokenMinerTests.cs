using System.Text.Json;
using PptxEditor.Core.Generation.Design;
using PptxEditor.Core.Generation.Model;

namespace OfficeEditor.Api.Tests.Unit;

public sealed class TokenMinerTests
{
    private const string UpdateSnapshotsEnvVar = "OE_UPDATE_SNAPSHOTS";

    private static string AetherLinkDeckPath =>
        Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..", "examples", "REF", "PPTX",
            "AetherLink-Glass-Shareholder-Overview.pptx"));

    private static string TokenPath =>
        Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..", "PptxEditor.Core", "Generation", "Design",
            "aetherlink.tokens.json"));

    [Fact]
    public void Mine_FromAetherLinkDeck_ProducesValidTokens()
    {
        Assert.True(File.Exists(AetherLinkDeckPath),
            $"AetherLink deck not found at {AetherLinkDeckPath}");

        var tokens = TokenMiner.Mine(AetherLinkDeckPath);

        Assert.NotNull(tokens);
        Assert.NotEmpty(tokens.Palette);
        Assert.True(tokens.Palette.ContainsKey("primary"),
            "Palette must have a 'primary' key");
        Assert.True(tokens.Palette.ContainsKey("ink"),
            "Palette must have an 'ink' key");
        Assert.True(tokens.Palette.ContainsKey("paper"),
            "Palette must have a 'paper' key");
        Assert.NotNull(tokens.Fonts.Display);
        Assert.NotNull(tokens.Fonts.Body);
        Assert.True(tokens.Metrics.TitleSizePt > 0);
        Assert.True(tokens.Metrics.BodySizePt > 0);
    }

    [Fact]
    public void Mine_FromAetherLinkDeck_MatchesCommittedTokens()
    {
        Assert.True(File.Exists(AetherLinkDeckPath));

        var tokens = TokenMiner.Mine(AetherLinkDeckPath);
        var actual = Serialize(tokens);

        if (Environment.GetEnvironmentVariable(UpdateSnapshotsEnvVar) == "1")
        {
            Directory.CreateDirectory(Path.GetDirectoryName(TokenPath)!);
            File.WriteAllText(TokenPath, actual);
        }

        Assert.True(File.Exists(TokenPath),
            $"Token file not found at {TokenPath}. Re-mine with {UpdateSnapshotsEnvVar}=1.");

        var expected = File.ReadAllText(TokenPath).Replace("\r\n", "\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Mine_Deserialized_RoundTrips()
    {
        Assert.True(File.Exists(AetherLinkDeckPath));

        var tokens = TokenMiner.Mine(AetherLinkDeckPath);
        var json = Serialize(tokens);

        var deserialized = JsonSerializer.Deserialize<DesignTokens>(json, SerializerOptions);
        Assert.NotNull(deserialized);

        var reSerialized = Serialize(deserialized!);
        Assert.Equal(json, reSerialized);
    }

    [Fact]
    public void Mine_NonExistentFile_Throws()
    {
        Assert.Throws<FileNotFoundException>(
            () => TokenMiner.Mine("/nonexistent/path.pptx"));
    }

    private static string Serialize(DesignTokens tokens)
        => JsonSerializer.Serialize(tokens, SerializerOptions) + "\n";

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
}
