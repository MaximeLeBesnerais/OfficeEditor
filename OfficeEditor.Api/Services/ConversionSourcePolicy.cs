using DocxEditor.Core.Generation.Assets;

namespace OfficeEditor.Api.Services;

/// <summary>
/// API-owned source policy for HTTP conversions. Uploaded Markdown and generation JSON are
/// untrusted, so their image sources must never resolve to server-local filesystem paths
/// (the process working directory, /etc, …). Only bounded data URIs are accepted; the
/// restriction is a deployment policy of the API conversion service, not a library invariant
/// — trusted callers keep the full <see cref="ImageSourceResolver"/> surface.
///
/// The policy is implemented purely through the library's existing
/// <see cref="ImageSourceOptions"/> path policy, with no resolver changes: an allowed root
/// that is never created (so every relative source resolves to a file that cannot exist, and
/// any <c>../</c> traversal escapes the root) plus absolute paths rejected. Data URIs are
/// unaffected by the path policy and remain capped by the default asset limits
/// (<see cref="ImageAssetOptions.Default"/>).
/// </summary>
internal static class ConversionSourcePolicy
{
    // A fixed, unpredictable sentinel directory under the process temp root. It is never
    // created and never read, so no local-file source can ever resolve inside it; a caller
    // who does not already have local file access cannot predict or control it.
    private const string SentinelRootName = "officeeditor-api-images-3f0a1c2e-9b4d-4e5f-8a6b-7c8d9e0f1a2b";

    /// <summary>
    /// Image source options that disable every local-file source: absolute paths are rejected
    /// and relative paths resolve against a sentinel allowed root that does not exist, so all
    /// local-file resolution fails (file-not-found or outside-the-root). Data URIs remain the
    /// only accepted source.
    /// </summary>
    public static ImageSourceOptions DataUriOnlyImageSources { get; } = new()
    {
        AllowedRoot = Path.Combine(Path.GetTempPath(), SentinelRootName),
        AllowAbsolutePaths = false
    };
}
