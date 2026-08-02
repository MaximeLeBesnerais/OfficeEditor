namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Resolves an image source string into an <see cref="ImageSourceResolution"/>. Accepts
/// data URIs and local files; HTTP(S) and any other absolute-URI scheme are rejected
/// explicitly (no network, no SSRF). Local paths are validated against the
/// <see cref="ImageSourceOptions"/> path policy: absolute paths only when enabled, and —
/// when an allowed root is configured — lexical containment plus symlink-aware canonical
/// containment for paths that exist.
///
/// Resolution is pure path policy; the loader reads and decodes the payload.
/// </summary>
public static class ImageSourceResolver
{
    /// <summary>
    /// Resolves a source string. Throws <see cref="ImageSourceException"/> with a typed
    /// <see cref="ImageSourceErrorCode"/> for every rejection. The returned descriptor is
    /// validated but its bytes are not yet read.
    /// </summary>
    public static ImageSourceResolution Resolve(string source, ImageSourceOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        var opts = options ?? ImageSourceOptions.Default;

        if (source.Length == 0 || string.IsNullOrWhiteSpace(source))
        {
            throw new ImageSourceException(ImageSourceErrorCode.SourceIsEmpty, "The image source is empty.");
        }

        if (IsDataUri(source))
        {
            return ResolveDataUri(source);
        }

        if (Path.IsPathRooted(source))
        {
            // Absolute filesystem path (handled before URI parsing so Windows drive letters
            // such as C:\… are never mistaken for URI schemes).
            return ResolveLocalFile(source, source, opts);
        }

        if (Uri.TryCreate(source, UriKind.Absolute, out var uri))
        {
            if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.RemoteSourceNotAllowed,
                    $"Remote image sources are not allowed (network access is disabled): '{source}'.");
            }
            if (uri.Scheme == Uri.UriSchemeFile)
            {
                return ResolveLocalFile(uri.LocalPath, source, opts);
            }
            throw new ImageSourceException(
                ImageSourceErrorCode.UnsupportedUriScheme,
                $"Unsupported image source scheme '{uri.Scheme}' in '{source}'. Use a data URI or a local file path.");
        }

        return ResolveLocalFile(source, source, opts);
    }

    private static bool IsDataUri(string source) =>
        source.StartsWith("data:", StringComparison.OrdinalIgnoreCase);

    private static ImageSourceResolution ResolveDataUri(string source)
    {
        var commaIndex = source.IndexOf(',');
        if (commaIndex < 0)
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.InvalidDataUri,
                $"Invalid data URI '{source}': expected 'data:[<media-type>][;base64],<payload>'.");
        }

        var header = source[5..commaIndex];
        var payload = source[(commaIndex + 1)..];
        var isBase64 = false;
        string? mediaType = null;

        foreach (var token in header.Split(';'))
        {
            var trimmed = token.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }
            if (string.Equals(trimmed, "base64", StringComparison.OrdinalIgnoreCase))
            {
                isBase64 = true;
            }
            else if (trimmed.Contains('/'))
            {
                mediaType = trimmed;
            }
        }

        return new ImageSourceResolution
        {
            Kind = ImageSourceKind.DataUri,
            Source = source,
            DeclaredMediaType = mediaType,
            Payload = payload,
            IsBase64 = isBase64
        };
    }

    private static ImageSourceResolution ResolveLocalFile(string path, string source, ImageSourceOptions opts)
    {
        if (Path.IsPathRooted(path) && !opts.AllowAbsolutePaths)
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.AbsolutePathNotAllowed,
                $"Absolute image path '{path}' is not allowed; enable AbsolutePaths in ImageSourceOptions or use a relative path inside the allowed root.");
        }

        string fullPath;
        string sourceFullPath;
        if (opts.AllowedRoot is { } root)
        {
            var rootFull = GetFullRootPath(root);
            fullPath = Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(rootFull, path));
            sourceFullPath = fullPath;

            if (!IsLexicallyContained(fullPath, rootFull))
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.OutsideAllowedRoot,
                    $"Image path '{path}' resolves outside the allowed root '{root}'.");
            }

            if (!File.Exists(fullPath))
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.FileNotFound,
                    $"Image file not found: '{path}'.");
            }

            var canonicalRoot = CanonicalizeExistingPath(rootFull, root);
            var canonicalPath = CanonicalizeExistingPath(fullPath, path);
            if (!IsLexicallyContained(canonicalPath, canonicalRoot))
            {
                throw new ImageSourceException(
                    ImageSourceErrorCode.OutsideAllowedRoot,
                    $"Image path '{path}' escapes the allowed root '{root}' through a symbolic link.");
            }

            // The loader reads this validated target rather than reopening the original
            // symlink-bearing path, narrowing the check/read race.
            fullPath = canonicalPath;
        }
        else
        {
            fullPath = Path.GetFullPath(path);
            sourceFullPath = fullPath;
        }

        if (!File.Exists(fullPath))
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.FileNotFound,
                $"Image file not found: '{path}'.");
        }

        return new ImageSourceResolution
        {
            Kind = ImageSourceKind.LocalFile,
            Source = source,
            DeclaredMediaType = DocxImageMediaTypes.ParseExtension(sourceFullPath) is { } hint
                ? DocxImageMediaTypes.GetContentType(hint)
                : null,
            FilePath = fullPath,
            FileName = Path.GetFileName(sourceFullPath)
        };
    }

    private static string GetFullRootPath(string root)
    {
        var full = Path.GetFullPath(root);
        return Path.TrimEndingDirectorySeparator(full);
    }

    private static bool IsLexicallyContained(string path, string root)
    {
        var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (string.Equals(path, root, comparison))
        {
            return true;
        }
        return path.StartsWith(root + Path.DirectorySeparatorChar, comparison)
            || path.StartsWith(root + Path.AltDirectorySeparatorChar, comparison);
    }

    /// <summary>
    /// Symlink-aware canonicalization for an existing path. Walks every component and resolves
    /// links to their final targets. Any inspection or link-resolution failure is surfaced as
    /// a typed asset error rather than falling back to an unvalidated lexical path.
    /// </summary>
    private static string CanonicalizeExistingPath(string path, string displayPath)
    {
        try
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
        catch (Exception ex) when (ex is not ImageSourceException)
        {
            throw new ImageSourceException(
                ImageSourceErrorCode.PathNotReadable,
                $"Image path could not be canonicalized safely: '{displayPath}'.",
                ex);
        }
    }
}
