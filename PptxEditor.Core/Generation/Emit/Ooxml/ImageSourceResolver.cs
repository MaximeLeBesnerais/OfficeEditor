namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Shared file-system resolution for generation-pipeline image sources. One documented
/// precedence applies everywhere (OOXML emitter, API DeckGenerationService):
/// (a) the repository root — a source that would escape the root ("../../etc/passwd",
/// absolute paths outside the repo) is rejected with <see cref="ArgumentException"/>
/// before any file probe; (b) the process current directory, likewise
/// containment-checked; (c) <see cref="FileNotFoundException"/>.
/// Data URIs and remote URLs are not file sources and never pass through here.
/// </summary>
public static class ImageSourceResolver
{
    /// <summary>
    /// Locates the repository root by walking up from the application base directory
    /// looking for a ".git" entry — a directory in a normal checkout, a file (pointer
    /// to the real gitdir) in a git worktree. Null when no ".git" entry is found.
    /// </summary>
    public static string? TryFindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var gitPath = Path.Combine(directory.FullName, ".git");
            if (Directory.Exists(gitPath) || File.Exists(gitPath))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return null;
    }

    /// <summary>
    /// Resolves <paramref name="source"/> against <paramref name="root"/> and requires
    /// the normalized result to stay inside <paramref name="root"/> (prefix comparison
    /// against root + directory separator, ordinal). Returns the absolute path.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> escapes <paramref name="root"/> (e.g. "../../…" or an
    /// absolute path outside the root).
    /// </exception>
    public static string ResolveContained(string root, string source)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(root);
        ArgumentException.ThrowIfNullOrWhiteSpace(source);

        var normalizedRoot = Path.GetFullPath(root)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        // Path.Combine returns a rooted source unchanged; GetFullPath collapses ".." segments.
        var resolved = Path.GetFullPath(Path.Combine(normalizedRoot, source));
        if (!resolved.StartsWith(normalizedRoot, StringComparison.Ordinal))
        {
            throw new ArgumentException(
                $"Image source '{source}' resolves to '{resolved}', which is outside the allowed root " +
                $"'{normalizedRoot}'. Image file sources must stay inside the repository.",
                nameof(source));
        }

        return resolved;
    }
}
