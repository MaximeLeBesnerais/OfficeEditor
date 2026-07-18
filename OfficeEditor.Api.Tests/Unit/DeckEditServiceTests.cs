using OfficeEditor.Api.Services;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Tests.Unit;

public class DeckEditServiceTests
{
    private static byte[] BuildDeckBytes()
    {
        using var builder = PresentationBuilder.Create();
        builder.AddSlide();
        builder.CurrentSlide.AddTitle("Original Title");
        builder.AddSlide();
        builder.CurrentSlide.AddTitle("Second Slide");
        return builder.SaveToBytes();
    }

    private static uint FindTitleElementId(byte[] deckBytes, int slide)
    {
        using var builder = PresentationBuilder.Open(deckBytes);
        return builder.Analyze()[slide - 1].Elements.First(e => e.Type == "Text").Id;
    }

    [Fact]
    public void ApplyInstructions_UnknownDeck_ReturnsNull()
    {
        var service = new DeckEditService(new StubDeckSessionStore());
        Assert.Null(service.ApplyInstructions(Guid.NewGuid(), """{ "operations": [] }"""));
    }

    [Fact]
    public void ApplyInstructions_InvalidJson_ReturnsParseError()
    {
        var store = new StubDeckSessionStore();
        var deckId = store.Add(BuildDeckBytes());
        var service = new DeckEditService(store);

        var response = service.ApplyInstructions(deckId, "this is not json");

        Assert.NotNull(response);
        Assert.False(response.Success);
        Assert.Null(response.Revision);
        Assert.Empty(response.ChangedSlides);
        var error = Assert.Single(response.Errors);
        Assert.Equal(-1, error.Index);
        Assert.Equal("parse", error.Type);
        Assert.Equal($"/api/decks/{deckId}/file", response.DownloadUrl);
    }

    [Fact]
    public void ApplyInstructions_ValidEdit_UpdatesStoreAndReportsUrls()
    {
        var store = new StubDeckSessionStore();
        var originalBytes = BuildDeckBytes();
        var deckId = store.Add(originalBytes);
        var titleId = FindTitleElementId(originalBytes, 1);
        var service = new DeckEditService(store);

        var json = $$"""
            { "operations": [ { "type": "replaceText", "slide": 1, "elementId": {{titleId}}, "text": "Edited Title" } ] }
            """;
        var response = service.ApplyInstructions(deckId, json);

        Assert.NotNull(response);
        Assert.True(response.Success);
        Assert.NotNull(response.Revision);
        Assert.Equal($"/api/decks/{deckId}/file", response.DownloadUrl);
        Assert.Empty(response.Errors);

        var changed = Assert.Single(response.ChangedSlides);
        Assert.Equal(1, changed.SlideIndex);
        Assert.Equal($"/api/decks/{deckId}/slides/1/preview", changed.PreviewUrl);

        // Store was updated; the edit round-trips through a reopen.
        var updatedBytes = store.GetBytes(deckId);
        Assert.NotNull(updatedBytes);
        Assert.NotEqual(originalBytes, updatedBytes);
        using var builder = PresentationBuilder.Open(updatedBytes);
        Assert.Contains(builder.Analyze()[0].Elements, e => e.Text == "Edited Title");
    }

    [Fact]
    public void ApplyInstructions_ValidationFailure_LeavesStoreUntouched()
    {
        var store = new StubDeckSessionStore();
        var originalBytes = BuildDeckBytes();
        var deckId = store.Add(originalBytes);
        var titleId = FindTitleElementId(originalBytes, 1);
        var service = new DeckEditService(store);

        var json = $$"""
            { "operations": [
                { "type": "replaceText", "slide": 1, "elementId": {{titleId}}, "text": "Nope" },
                { "type": "replaceText", "slide": 42, "elementId": {{titleId}}, "text": "Nope" }
            ] }
            """;
        var response = service.ApplyInstructions(deckId, json);

        Assert.NotNull(response);
        Assert.False(response.Success);
        Assert.Null(response.Revision);
        Assert.Empty(response.ChangedSlides);
        var error = Assert.Single(response.Errors);
        Assert.Equal(1, error.Index);
        Assert.Equal("replaceText", error.Type);

        // Validate-then-execute: nothing was applied, store bytes unchanged.
        Assert.Equal(originalBytes, store.GetBytes(deckId));
    }

    [Fact]
    public void ApplyInstructions_StructuralBatch_ChangesSlideCount()
    {
        var store = new StubDeckSessionStore();
        var originalBytes = BuildDeckBytes();
        var deckId = store.Add(originalBytes);
        var service = new DeckEditService(store);

        var json = """
            { "operations": [
                { "type": "duplicateSlide", "slide": 1 },
                { "type": "moveSlide", "from": 2, "to": 1 }
            ] }
            """;
        var response = service.ApplyInstructions(deckId, json);

        Assert.NotNull(response);
        Assert.True(response.Success, string.Join("; ", response.Errors.Select(e => e.Error)));
        Assert.Equal(2, response.ChangedSlides.Count); // {2} from duplicate ∪ {1,2} from move

        var updatedBytes = store.GetBytes(deckId);
        Assert.NotNull(updatedBytes);
        using var builder = PresentationBuilder.Open(updatedBytes);
        Assert.Equal(3, builder.SlideCount);
    }

    [Fact]
    public void ApplyInstructions_DeleteOnlySlide_IsRejected()
    {
        var store = new StubDeckSessionStore();
        using (var single = PresentationBuilder.Create())
        {
            single.AddSlide();
            single.CurrentSlide.AddTitle("Only");
            var deckId = store.Add(single.SaveToBytes());
            var service = new DeckEditService(store);

            var response = service.ApplyInstructions(deckId,
                """{ "operations": [ { "type": "deleteSlide", "slide": 1 } ] }""");

            Assert.NotNull(response);
            Assert.False(response.Success);
            Assert.Contains("last remaining slide", Assert.Single(response.Errors).Error);
        }
    }
}
