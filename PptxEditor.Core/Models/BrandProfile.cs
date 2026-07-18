namespace PptxEditor.Core.Models;

/// <summary>
/// Immutable brand profile extracted from a presentation's first slide master:
/// theme colors, theme fonts, level-0 master text-style defaults, background color
/// and slide size in points. Read-only extraction — no document part is ever written.
/// </summary>
public record BrandProfile
{
    /// <summary>
    /// Theme color scheme slots (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) as
    /// #RRGGBB, sysClr entries resolved via their LastColor. Sorted by key for
    /// deterministic serialization. Raw scheme values — clrMap remapping not applied.
    /// </summary>
    public required IReadOnlyDictionary<string, string> ThemeColors { get; init; }

    /// <summary>Major (heading) latin theme font typeface, or null when absent.</summary>
    public string? MajorFont { get; init; }

    /// <summary>Minor (body) latin theme font typeface, or null when absent.</summary>
    public string? MinorFont { get; init; }

    /// <summary>Level-0 title text-style defaults from the master, or null when undefined.</summary>
    public BrandTextStyle? Title { get; init; }

    /// <summary>Level-0 body text-style defaults from the master, or null when undefined.</summary>
    public BrandTextStyle? Body { get; init; }

    /// <summary>Level-0 other (non-placeholder) text-style defaults from the master, or null when undefined.</summary>
    public BrandTextStyle? Other { get; init; }

    /// <summary>
    /// Resolved slide background color (#RRGGBB) from the slide/layout/master cascade,
    /// or null when no explicit solid-fill background is defined (e.g. theme bgRef backgrounds).
    /// </summary>
    public string? BackgroundColor { get; init; }

    /// <summary>Slide width in points (EMU / 12700).</summary>
    public double SlideWidthPt { get; init; }

    /// <summary>Slide height in points (EMU / 12700).</summary>
    public double SlideHeightPt { get; init; }

    /// <summary>Non-fatal extraction caveats (e.g. custom clrMap remapping detected).</summary>
    public IReadOnlyList<string> Warnings { get; init; } = [];
}

/// <summary>
/// Level-0 master text-style defaults for one placeholder class.
/// FontFamily has theme font references (+mj-lt / +mn-lt) already resolved to the
/// theme's major/minor latin typeface when the theme defines them.
/// </summary>
public record BrandTextStyle(string? FontFamily, double? FontSizePt, bool? Bold, bool? Italic, string? Color);
