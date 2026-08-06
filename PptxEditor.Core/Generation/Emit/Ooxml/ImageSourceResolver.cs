namespace PptxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Shared file-system resolution for generation-pipeline image sources. Resolution is
/// rooted at one directory only — the document directory (the directory of the JSON
/// deck). A relative source that would escape that directory ("../../…", or a symlink
/// inside the directory pointing outside it) is rejected with
/// <see cref="ArgumentException"/> before any file probe; a source that stays inside but
/// does not exist surfaces as <see cref="FileNotFoundException"/> at the call site.
/// Data URIs and remote URLs are not file sources and never pass through here.
/// </summary>
public static class ImageSourceResolver
{
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
        var (normalizedRoot, resolved) = ResolveRelative(root, source);
        if (!IsContained(resolved, normalizedRoot))
        {
            throw new ArgumentException(
                $"Image source '{source}' resolves to '{resolved}', which is outside the allowed root " +
                $"'{normalizedRoot}'. Image file sources must stay inside the repository.",
                nameof(source));
        }

        return resolved;
    }

    /// <summary>
    /// Resolves <paramref name="source"/> against <paramref name="root"/> with full
    /// containment: a lexical prefix check first, then — when the file exists — a
    /// symlink-aware canonical check on the real paths of both the root and the resolved
    /// file, so a link inside the root that points outside it is rejected too. Returns
    /// the canonical absolute path of the existing file (or the lexical path when the
    /// file does not exist, so the caller can report it as missing).
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> escapes <paramref name="root"/> lexically or through a
    /// symbolic link, or its path cannot be canonicalized safely.
    /// </exception>
    public static string ResolveContainedCanonical(string root, string source)
    {
        var (normalizedRoot, resolved) = ResolveRelative(root, source);
        if (!IsContained(resolved, normalizedRoot))
        {
            throw new ArgumentException(
                $"Image source '{source}' resolves to '{resolved}', which is outside the document directory " +
                $"'{normalizedRoot}'. Relative image sources must stay inside the directory of the JSON document.",
                nameof(source));
        }

        if (!File.Exists(resolved))
        {
            return resolved;
        }

        var canonicalRoot = CanonicalizeExistingPath(root);
        var canonicalPath = CanonicalizeExistingPath(resolved);
        if (!IsContained(canonicalPath, canonicalRoot))
        {
            throw new ArgumentException(
                $"Image source '{source}' escapes the document directory '{root}' through a symbolic link " +
                $"(resolves to '{canonicalPath}').",
                nameof(source));
        }

        // The caller reads this validated target rather than reopening the original
        // symlink-bearing path, narrowing the check/read race.
        return canonicalPath;
    }

    private static (string NormalizedRoot, string Resolved) ResolveRelative(string root, string source)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(root);
        ArgumentException.ThrowIfNullOrWhiteSpace(source);

        var normalizedRoot = Path.GetFullPath(root)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        // Path.Combine returns a rooted source unchanged; GetFullPath collapses ".." segments.
        var resolved = Path.GetFullPath(Path.Combine(normalizedRoot, source));
        return (normalizedRoot, resolved);
    }

    private static bool IsContained(string path, string root) =>
        path.StartsWith(root, StringComparison.Ordinal);

    /// <summary>
    /// Symlink-aware canonicalization for an existing path. Walks every component and
    /// resolves links to their final targets, repeating the walk until the path is stable
    /// so that a resolved target's own links are canonicalized too (on macOS the temp
    /// root is reached through /var → /private/var, and a stored link target can keep
    /// the /var spelling). Any inspection or link-resolution failure — or a symlink
    /// cycle — is surfaced as an <see cref="ArgumentException"/> rather than falling back
    /// to an unvalidated lexical path.
    /// </summary>
    private static string CanonicalizeExistingPath(string path)
    {
        try
        {
            var current = Path.GetFullPath(path);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            while (seen.Add(current))
            {
                var resolved = ResolvePathLinks(current);
                if (string.Equals(resolved, current, StringComparison.Ordinal))
                {
                    return current;
                }
                current = resolved;
            }
            throw new IOException($"Symbolic link cycle detected while canonicalizing '{path}'.");
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new ArgumentException($"Image path could not be canonicalized safely: '{path}'.", ex);
        }
    }

    private static string ResolvePathLinks(string path)
    {
        var full = Path.GetFullPath(path);
        var root = Path.GetPathRoot(full) ?? string.Empty;
        var components = full[root.Length..].Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);

        var current = root;
        foreach (var component in components)
        {
            current = Path.Combine(current, component);
            var attributes = File.GetAttributes(current);
            if ((attributes & FileAttributes.ReparsePoint) != 0)
            {
                FileSystemInfo info = (attributes & FileAttributes.Directory) != 0
                    ? new DirectoryInfo(current)
                    : new FileInfo(current);
                current = info.ResolveLinkTarget(true)?.FullName
                    ?? throw new IOException($"Symbolic link target could not be resolved: '{current}'.");
            }
        }
        return Path.GetFullPath(current);
    }
}
