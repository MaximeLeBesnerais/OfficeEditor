using OfficeEditor.Api.Services;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Tests.Unit;

public class DeckServiceTests
{
    private static byte[] BuildDeckBytes()
    {
        using var builder = PresentationBuilder.Create();
        builder.AddSlide();
        builder.CurrentSlide.AddTitle("Deck Title");
        builder.CurrentSlide.AddText("Body text");
        builder.AddSlide();
        builder.CurrentSlide.AddTable(new List<List<string>> { new() { "h1", "h2" }, new() { "a", "b" } });
        return builder.SaveToBytes();
    }

    [Fact]
    public void GetAnatomy_UnknownDeck_ReturnsNull()
    {
        var service = new DeckService(new StubDeckSessionStore());
        Assert.Null(service.GetAnatomy(Guid.NewGuid()));
    }

    [Fact]
    public void GetAnatomy_ReturnsSlidesElementsAndPositions()
    {
        var store = new StubDeckSessionStore();
        var deckId = store.Add(BuildDeckBytes());
        var service = new DeckService(store);

        var anatomy = service.GetAnatomy(deckId);

        Assert.NotNull(anatomy);
        Assert.Equal(deckId, anatomy.DeckId);
        Assert.Equal(2, anatomy.SlideCount);
        Assert.Equal(2, anatomy.Slides.Count);
        Assert.Equal(1, anatomy.Slides[0].SlideIndex);
        Assert.Equal(2, anatomy.Slides[1].SlideIndex);

        var body = anatomy.Slides[0].Elements.First(e => e.Text == "Body text");
        Assert.Equal("Text", body.Type);
        Assert.True(body.Id > 0);
        Assert.NotNull(body.Position);
        // SlideBuilder.AddText: off=(0,1440000) ext=(7200000,3600000) EMU.
        Assert.Equal(0, body.Position.X);
        Assert.Equal(1440000, body.Position.Y);
        Assert.Equal(7200000, body.Position.Cx);
        Assert.Equal(3600000, body.Position.Cy);

        var table = anatomy.Slides[1].Elements.First(e => e.Type == "Table");
        Assert.NotNull(table.Position);
        Assert.NotNull(table.TableData);
        Assert.Equal("h1", table.TableData[0][0]);
    }

    [Fact]
    public void GetDeckBytes_UnknownDeck_ReturnsNull()
    {
        var service = new DeckService(new StubDeckSessionStore());
        Assert.Null(service.GetDeckBytes(Guid.NewGuid()));
    }

    [Fact]
    public void GetDeckBytes_ReturnsStoredBytes()
    {
        var store = new StubDeckSessionStore();
        var bytes = BuildDeckBytes();
        var deckId = store.Add(bytes);
        var service = new DeckService(store);

        var result = service.GetDeckBytes(deckId);

        Assert.NotNull(result);
        Assert.Equal(bytes, result);
    }
}
