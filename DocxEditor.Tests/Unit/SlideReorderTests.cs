using PptxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Exact-order regression tests for <see cref="PresentationBuilder.ReorderSlide"/>.
/// The suspected off-by-one: moving a slide EARLIER (fromIndex &gt; toIndex) lands it
/// at toIndex + 1 because the code inserts AFTER the element at the target index.
/// </summary>
public sealed class SlideReorderTests : IDisposable
{
    private readonly string _tempDirectory;

    public SlideReorderTests()
    {
        _tempDirectory = SlideOpsTestHelpers.CreateTempDirectory(nameof(SlideReorderTests));
    }

    public void Dispose()
    {
        SlideOpsTestHelpers.BestEffortDelete(_tempDirectory);
    }

    [Theory]
    [InlineData(new[] { "A", "B", "C" }, 2, 0, new[] { "C", "A", "B" })] // move earlier: suspected off-by-one
    [InlineData(new[] { "A", "B", "C" }, 0, 2, new[] { "B", "C", "A" })] // move later to end
    [InlineData(new[] { "A", "B", "C", "D" }, 1, 3, new[] { "A", "C", "D", "B" })] // move later to end (4 slides)
    public void ReorderSlide_ExactOrderAfterReopen(string[] titles, int fromIndex, int toIndex, string[] expectedOrder)
    {
        var path = SlideOpsTestHelpers.CreateDeckWithTitledSlides(_tempDirectory, titles);

        using (var builder = PresentationBuilder.Open(path))
        {
            builder.ReorderSlide(fromIndex, toIndex);
            Assert.Equal(titles.Length, builder.SlideCount);
            builder.Save();
        }

        var texts = SlideOpsTestHelpers.ReadSlideTextsInOrder(path);

        // Exact final order: each builder-created slide's InnerText is exactly its title.
        Assert.Equal(expectedOrder, texts);
    }
}
