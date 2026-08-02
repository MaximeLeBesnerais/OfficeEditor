using OfficeEditor.Api.Services;

namespace OfficeEditor.Api.Tests.Unit;

/// <summary>
/// Repository-root discovery used by DemoDeckService, SampleFileService and the
/// OfficialRenderService to resolve repo-relative REF/template paths. The tests run
/// inside a git checkout, so the walk-up from the test bin directory must find the
/// repo's ".git" entry — the same environment-invariance assumption the other tests
/// rely on when reading examples/REF/ and demo/demo-deck.json.
/// </summary>
public sealed class RepositoryRootLocatorTests
{
    [Fact]
    public void FindOrNull_ReturnsDirectoryContainingGitMarker()
    {
        var root = RepositoryRootLocator.FindOrNull();

        Assert.NotNull(root);
        Assert.True(
            Directory.Exists(Path.Combine(root!, ".git")) || File.Exists(Path.Combine(root!, ".git")),
            $"'{root}' has no .git entry (directory or worktree pointer)");
    }

    [Fact]
    public void FindOrNull_ReturnsAncestorOfTheTestBinDirectory()
    {
        var root = RepositoryRootLocator.FindOrNull();

        Assert.NotNull(root);
        // The test assembly lives under <repo>/OfficeEditor.Api.Tests/bin/<config>/net9.0,
        // so the located root must be a strict ancestor of the base directory.
        var baseDir = Path.GetFullPath(AppContext.BaseDirectory);
        var rootFull = Path.GetFullPath(root!);
        Assert.StartsWith(rootFull + Path.DirectorySeparatorChar, baseDir);
    }

    [Fact]
    public void FindOrFallback_MatchesFindOrNull_WhenGitCheckout()
    {
        // In a git checkout both methods resolve to the same repo root; FindOrFallback
        // only diverges when no ".git" entry exists at all.
        Assert.Equal(RepositoryRootLocator.FindOrNull(), RepositoryRootLocator.FindOrFallback());
    }

    [Fact]
    public void FindOrFallback_ReturnsNonNullPath()
    {
        Assert.NotNull(RepositoryRootLocator.FindOrFallback());
    }
}
