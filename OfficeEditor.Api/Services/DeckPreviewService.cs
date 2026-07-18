using System.Security.Cryptography;
using PptxEditor.Core.Builders;

namespace OfficeEditor.Api.Services;

/// <summary>
/// Renders a single slide of a deck session. Internal seam so the v0 whole-deck
/// render can be swapped for W3's per-slide emission without touching the endpoint.
/// </summary>
public interface ISlideRenderer
{
    /// <summary>Renders 0-based <paramref name="slideIndex"/> to the requested format.</summary>
    byte[] RenderSlide(DeckSession session, int slideIndex, string normalizedFormat, int ppi);
}

public sealed record SlidePreview(
    byte[] Bytes,
    string ContentType,
    string ETag,
    IReadOnlyList<string> Warnings);

public interface IDeckPreviewService
{
    /// <summary>
    /// Returns the rendered slide for the session, serving from the per-deck
    /// rendered-page cache when the same (slide, format, ppi) was already rendered
    /// at the current deck revision.
    /// </summary>
    Task<SlidePreview> GetSlidePreviewAsync(
        DeckSession session, int slideNumber, string normalizedFormat, int ppi, CancellationToken ct = default);
}

/// <summary>v0 renderer: whole-deck render via the merged PresentationBuilder API.</summary>
public sealed class BuilderSlideRenderer : ISlideRenderer
{
    public byte[] RenderSlide(DeckSession session, int slideIndex, string normalizedFormat, int ppi)
    {
        // W3-INTEGRATION: switch to per-slide emission when
        // w3-request-PresentationBuilder.cs.patch lands (single-slide converter
        // overload / true per-slide compile instead of whole-deck render + slice).
        using var builder = PresentationBuilder.Open(session.SourceBytes);
        return builder.ExportThumbnail(slideIndex, new ThumbnailOptions
        {
            Ppi = ppi,
            Format = normalizedFormat
        });
    }
}

public sealed class DeckPreviewService : IDeckPreviewService
{
    private readonly ISlideRenderer _renderer;

    public DeckPreviewService(ISlideRenderer renderer)
    {
        _renderer = renderer;
    }

    public async Task<SlidePreview> GetSlidePreviewAsync(
        DeckSession session, int slideNumber, string normalizedFormat, int ppi, CancellationToken ct = default)
    {
        var slideIndex = slideNumber - 1; // API is 1-based; builder is 0-based.
        var cacheKey = new RenderedPageKey(slideIndex, normalizedFormat, ppi);

        // Serialize against concurrent edits (W7) and concurrent renders of this deck.
        await session.RenderLock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            if (!session.RenderedPages.TryGetValue(cacheKey, out var bytes))
            {
                bytes = await Task.Run(
                    () => _renderer.RenderSlide(session, slideIndex, normalizedFormat, ppi), ct)
                    .ConfigureAwait(false);
                session.RenderedPages[cacheKey] = bytes;
            }

            return new SlidePreview(
                bytes,
                DeckPreviewValidators.ContentTypeForFormat(normalizedFormat),
                ComputeETag(bytes),
                // W3-INTEGRATION: populate from per-slide conversion warnings once W3's
                // per-slide emission lands; until then the list is always empty.
                Warnings: []);
        }
        finally
        {
            session.RenderLock.Release();
        }
    }

    private static string ComputeETag(byte[] bytes)
    {
        var hash = SHA256.HashData(bytes);
        return $"\"{Convert.ToHexString(hash)[..32]}\"";
    }
}
