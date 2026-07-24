namespace OfficeEditor.Api.Services;

/// <summary>
/// Locates the repository root by walking up from the application base directory looking
/// for a ".git" entry. In a normal checkout ".git" is a directory; in a git worktree it is
/// a file (a pointer to the real gitdir), so both forms must be accepted — the same
/// approach as PptxEditor.Core's OoxmlEmitter.TryFindRepositoryRoot.
/// </summary>
internal static class RepositoryRootLocator
{
    /// <summary>Repository root, or null when no ".git" entry is found walking up.</summary>
    public static string? FindOrNull()
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
    /// Repository root, falling back to the parent of the API project folder
    /// (bin/Debug/net9.0 → OfficeEditor.Api → repo root) when no ".git" entry is found.
    /// </summary>
    public static string FindOrFallback()
    {
        if (FindOrNull() is { } root)
        {
            return root;
        }

        var fallback = new DirectoryInfo(AppContext.BaseDirectory);
        for (var i = 0; i < 4 && fallback is not null; i++)
        {
            fallback = fallback.Parent;
        }

        return fallback?.FullName ?? AppContext.BaseDirectory;
    }
}
