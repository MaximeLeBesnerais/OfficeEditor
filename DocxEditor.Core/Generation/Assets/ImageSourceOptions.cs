namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Path policy for image sources. Remote HTTP(S) sources are always rejected (no network,
/// no SSRF); this controls the local-file half of resolution.
///
/// Relative sources are resolved against <see cref="AllowedRoot"/> when it is set (and must
/// stay inside it, both lexically and — when the file exists — canonically, symlink-aware).
/// When <see cref="AllowedRoot"/> is unset, relative sources resolve against the process
/// current directory. Absolute paths are rejected unless <see cref="AllowAbsolutePaths"/> is
/// enabled, and remain subject to the allowed-root boundary when one is configured.
/// </summary>
public sealed record ImageSourceOptions
{
    /// <summary>Default policy: no allowed root, absolute paths rejected.</summary>
    public static ImageSourceOptions Default { get; } = new();

    /// <summary>
    /// Optional absolute directory every local source must resolve inside. When set, both a
    /// lexical prefix check and (for existing paths) a symlink-aware canonical containment
    /// check are applied.
    /// </summary>
    public string? AllowedRoot { get; init; }

    /// <summary>
    /// When true, absolute filesystem paths are accepted (still constrained by
    /// <see cref="AllowedRoot"/> when one is set). Defaults to false.
    /// </summary>
    public bool AllowAbsolutePaths { get; init; }
}
