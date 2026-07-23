using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Tests for <see cref="PresentationBuilder.CombineFontDirectories"/> — the rule that wires
/// <see cref="ThumbnailOptions.FontDirectory"/> into the Typst compile: embedded PPTX fonts
/// always come FIRST in the combined path (never overridden), the provided directory extends
/// the search path, and null/blank inputs collapse to null (compiler default behavior).
/// </summary>
public sealed class ThumbnailFontDirectoryTests
{
    [Fact]
    public void Combine_EmbeddedAndProvided_ReturnsEmbeddedFirstThenProvided()
    {
        var combined = PresentationBuilder.CombineFontDirectories("/tmp/deck-fonts", "/System/Library/Fonts");

        Assert.Equal(
            "/tmp/deck-fonts" + Path.PathSeparator + "/System/Library/Fonts",
            combined);
    }

    [Fact]
    public void Combine_ProvidedOnly_ReturnsProvided()
    {
        var combined = PresentationBuilder.CombineFontDirectories(null, "/System/Library/Fonts:/Library/Fonts");

        Assert.Equal("/System/Library/Fonts:/Library/Fonts", combined);
    }

    [Fact]
    public void Combine_EmbeddedOnly_ReturnsEmbedded()
    {
        var combined = PresentationBuilder.CombineFontDirectories("/tmp/deck-fonts", null);

        Assert.Equal("/tmp/deck-fonts", combined);
    }

    [Fact]
    public void Combine_Neither_ReturnsNull()
    {
        Assert.Null(PresentationBuilder.CombineFontDirectories(null, null));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Combine_BlankInputs_AreDropped(string? blank)
    {
        // Blank treated as absent: never emits a leading/trailing separator.
        Assert.Null(PresentationBuilder.CombineFontDirectories(blank, blank));
        Assert.Equal("/fonts", PresentationBuilder.CombineFontDirectories(blank, "/fonts"));
        Assert.Equal("/fonts", PresentationBuilder.CombineFontDirectories("/fonts", blank));
    }

    [Fact]
    public void ThumbnailOptions_FontDirectory_DefaultsToNull()
    {
        Assert.Null(new ThumbnailOptions().FontDirectory);
    }
}
