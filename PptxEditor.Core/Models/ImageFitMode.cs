namespace PptxEditor.Core.Models;

/// <summary>
/// How a replacement image is fitted into the target picture frame
/// (<see cref="PptxEditor.Core.Services.PptxElementReplacer"/> / the replaceImage
/// instruction's "fit" argument). The frame's a:xfrm is never moved except for
/// <see cref="Contain"/>.
/// </summary>
public enum ImageFitMode
{
    /// <summary>Stretch the image into the frame, ignoring aspect ratio (default, legacy behavior).</summary>
    Stretch,

    /// <summary>Cover: center-crop the image via a:srcRect so it fills the frame at the frame's aspect.</summary>
    Fill,

    /// <summary>Crop: write a caller-supplied a:srcRect verbatim (null rect → treated as <see cref="Fill"/>).</summary>
    Crop,

    /// <summary>Fit inside: shrink the frame's a:ext around the frame's center so the frame matches the image aspect.</summary>
    Contain
}

/// <summary>
/// Source crop rectangle in 1/1000ths of a percent (same scale family as spcPct):
/// 100000 = 100%. Each component is the fraction cropped
/// from that edge of the source image; valid range is 0–100000.
/// </summary>
public readonly record struct SourceRect(int Left, int Top, int Right, int Bottom);
