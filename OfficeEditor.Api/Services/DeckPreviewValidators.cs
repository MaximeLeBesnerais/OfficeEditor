namespace OfficeEditor.Api.Services;

/// <summary>
/// Pure validation helpers for the per-slide preview endpoint. Extracted from the
/// HTTP layer so they are unit-testable without a web host (no Mvc.Testing).
/// </summary>
public static class DeckPreviewValidators
{
    public const int MinPpi = 36;
    public const int MaxPpi = 600;
    public const int DefaultPpi = 150;

    public const string PngContentType = "image/png";
    public const string SvgContentType = "image/svg+xml";

    /// <summary>
    /// Validates the 1-based slide number against the deck's slide count.
    /// Returns an error message, or null when valid.
    /// </summary>
    public static string? ValidateSlideNumber(int slideNumber, int slideCount)
    {
        if (slideNumber < 1)
        {
            return $"Slide number must be 1 or greater (got {slideNumber}).";
        }

        if (slideNumber > slideCount)
        {
            return $"Slide {slideNumber} is out of range: the deck has {slideCount} slide(s).";
        }

        return null;
    }

    /// <summary>
    /// Normalizes the format query value to "png" or "svg".
    /// Returns false with an error message for anything else.
    /// </summary>
    public static bool TryNormalizeFormat(string? format, out string normalized, out string? error)
    {
        normalized = (format ?? "png").Trim().ToLowerInvariant();
        if (normalized is "png" or "svg")
        {
            error = null;
            return true;
        }

        error = $"Unsupported preview format '{format}'. Supported values: png, svg.";
        return false;
    }

    /// <summary>Clamps the requested ppi into the supported 36–600 range.</summary>
    public static int ClampPpi(int? ppi)
    {
        var value = ppi ?? DefaultPpi;
        return Math.Clamp(value, MinPpi, MaxPpi);
    }

    public static string ContentTypeForFormat(string normalizedFormat)
    {
        return normalizedFormat == "svg" ? SvgContentType : PngContentType;
    }
}
