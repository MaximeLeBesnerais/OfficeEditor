using System.Security.Cryptography;
using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Image source path policy and limits: data-URI parsing, remote-source rejection, absolute
/// path policy, allowed-root lexical and symlink-aware containment, file-not-found handling,
/// and the loader's byte limits.
/// </summary>
public class ImageSourcePolicyTests
{
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Resolve_EmptyOrWhitespace_ThrowsSourceIsEmpty(string source)
    {
        var exception = Assert.Throws<ImageSourceException>(() => ImageSourceResolver.Resolve(source));
        Assert.Equal(ImageSourceErrorCode.SourceIsEmpty, exception.Code);
    }

    [Theory]
    [InlineData("http://example.com/a.png")]
    [InlineData("https://example.com/a.png")]
    public void Resolve_RemoteHttp_IsRejected(string source)
    {
        var exception = Assert.Throws<ImageSourceException>(() => ImageSourceResolver.Resolve(source));
        Assert.Equal(ImageSourceErrorCode.RemoteSourceNotAllowed, exception.Code);
        Assert.Contains("network access is disabled", exception.Message);
    }

    [Fact]
    public void Resolve_UnsupportedScheme_IsRejected()
    {
        var exception = Assert.Throws<ImageSourceException>(() => ImageSourceResolver.Resolve("ftp://example.com/a.png"));
        Assert.Equal(ImageSourceErrorCode.UnsupportedUriScheme, exception.Code);
    }

    [Fact]
    public void Resolve_ValidDataUri_ParsesMetadata()
    {
        var resolution = ImageSourceResolver.Resolve(TestImages.DataUriBase64(TestImages.Png(4, 4)));

        Assert.Equal(ImageSourceKind.DataUri, resolution.Kind);
        Assert.Equal("image/png", resolution.DeclaredMediaType);
        Assert.True(resolution.IsBase64);
        Assert.False(string.IsNullOrEmpty(resolution.Payload));
        Assert.Null(resolution.FilePath);
    }

    [Fact]
    public void Resolve_DataUriWithoutBase64_MarksPayload()
    {
        var resolution = ImageSourceResolver.Resolve("data:image/png,%89PNG%0D%0A");

        Assert.Equal(ImageSourceKind.DataUri, resolution.Kind);
        Assert.False(resolution.IsBase64);
        Assert.Equal("%89PNG%0D%0A", resolution.Payload);
    }

    [Fact]
    public void Resolve_DataUriWithoutMediaType_IsAccepted()
    {
        var resolution = ImageSourceResolver.Resolve("data:,%89PNG");
        Assert.Null(resolution.DeclaredMediaType);
        Assert.False(resolution.IsBase64);
    }

    [Fact]
    public void Resolve_MalformedDataUri_IsRejected()
    {
        var exception = Assert.Throws<ImageSourceException>(() => ImageSourceResolver.Resolve("data:image/png"));
        Assert.Equal(ImageSourceErrorCode.InvalidDataUri, exception.Code);
    }

    [Fact]
    public void Resolve_AbsolutePath_RejectedByDefault()
    {
        var absolute = Path.Combine(Path.GetTempPath(), "pic.png");
        var exception = Assert.Throws<ImageSourceException>(() => ImageSourceResolver.Resolve(absolute));
        Assert.Equal(ImageSourceErrorCode.AbsolutePathNotAllowed, exception.Code);
    }

    [Fact]
    public void Resolve_AbsolutePath_AllowedWhenEnabled()
    {
        using var temp = new TempDirectory();
        var file = temp.File("pic.png");
        File.WriteAllBytes(file, TestImages.Png(4, 4));

        var resolution = ImageSourceResolver.Resolve(
            file,
            new ImageSourceOptions { AllowAbsolutePaths = true });

        Assert.Equal(ImageSourceKind.LocalFile, resolution.Kind);
        Assert.True(File.Exists(resolution.FilePath));
        Assert.EndsWith("pic.png", resolution.FilePath);
    }

    [Fact]
    public void Resolve_FileUri_ResolvesWhenEnabled()
    {
        using var temp = new TempDirectory();
        var file = temp.File("pic.png");
        File.WriteAllBytes(file, TestImages.Png(4, 4));

        var resolution = ImageSourceResolver.Resolve(
            new Uri(file).AbsoluteUri,
            new ImageSourceOptions { AllowAbsolutePaths = true });

        Assert.Equal(ImageSourceKind.LocalFile, resolution.Kind);
        Assert.True(File.Exists(resolution.FilePath));
        Assert.EndsWith("pic.png", resolution.FilePath);
    }

    [Fact]
    public void Resolve_RelativeInsideRoot_Resolves()
    {
        using var temp = new TempDirectory();
        var file = temp.File("pic.png");
        File.WriteAllBytes(file, TestImages.Png(4, 4));

        var resolution = ImageSourceResolver.Resolve(
            "pic.png",
            new ImageSourceOptions { AllowedRoot = temp.Path });

        Assert.Equal(ImageSourceKind.LocalFile, resolution.Kind);
        Assert.True(File.Exists(resolution.FilePath));
        Assert.EndsWith("pic.png", resolution.FilePath);
    }

    [Fact]
    public void Resolve_TraversalOutsideRoot_IsRejected()
    {
        using var temp = new TempDirectory();
        var exception = Assert.Throws<ImageSourceException>(() =>
            ImageSourceResolver.Resolve(
                "../escape.png",
                new ImageSourceOptions { AllowedRoot = temp.Path }));
        Assert.Equal(ImageSourceErrorCode.OutsideAllowedRoot, exception.Code);
    }

    [Fact]
    public void Resolve_AbsolutePathInsideRoot_RequiresAbsoluteFlag()
    {
        using var temp = new TempDirectory();
        var file = temp.File("pic.png");
        File.WriteAllBytes(file, TestImages.Png(4, 4));

        // Absolute path inside the root but the flag is off → rejected before containment.
        var exception = Assert.Throws<ImageSourceException>(() =>
            ImageSourceResolver.Resolve(file, new ImageSourceOptions { AllowedRoot = temp.Path }));
        Assert.Equal(ImageSourceErrorCode.AbsolutePathNotAllowed, exception.Code);

        // With the flag on, containment applies and it resolves.
        var resolution = ImageSourceResolver.Resolve(file, new ImageSourceOptions
        {
            AllowedRoot = temp.Path,
            AllowAbsolutePaths = true
        });
        Assert.Equal(ImageSourceKind.LocalFile, resolution.Kind);
        Assert.True(File.Exists(resolution.FilePath));
        Assert.EndsWith("pic.png", resolution.FilePath);
    }

    [Fact]
    public void Resolve_MissingFileInsideRoot_IsRejected()
    {
        using var temp = new TempDirectory();
        var exception = Assert.Throws<ImageSourceException>(() =>
            ImageSourceResolver.Resolve("missing.png", new ImageSourceOptions { AllowedRoot = temp.Path }));
        Assert.Equal(ImageSourceErrorCode.FileNotFound, exception.Code);
    }

    [Fact]
    public void Resolve_SymlinkEscapingRoot_IsRejected()
    {
        using var temp = new TempDirectory();
        var root = Path.Combine(temp.Path, "root");
        var outside = Path.Combine(temp.Path, "outside");
        Directory.CreateDirectory(root);
        Directory.CreateDirectory(outside);

        var secret = Path.Combine(outside, "secret.png");
        File.WriteAllBytes(secret, TestImages.Png(4, 4));

        var link = Path.Combine(root, "link.png");
        File.CreateSymbolicLink(link, secret);

        var exception = Assert.Throws<ImageSourceException>(() =>
            ImageSourceResolver.Resolve("link.png", new ImageSourceOptions { AllowedRoot = root }));
        Assert.Equal(ImageSourceErrorCode.OutsideAllowedRoot, exception.Code);
        Assert.Contains("symbolic link", exception.Message);
    }

    [Fact]
    public void Resolve_SymlinkInsideRoot_IsAllowed()
    {
        using var temp = new TempDirectory();
        var root = Path.Combine(temp.Path, "root");
        Directory.CreateDirectory(root);

        var real = Path.Combine(root, "real.png");
        File.WriteAllBytes(real, TestImages.Png(4, 4));

        var link = Path.Combine(root, "link.png");
        File.CreateSymbolicLink(link, real);

        var resolution = ImageSourceResolver.Resolve("link.png", new ImageSourceOptions { AllowedRoot = root });
        Assert.Equal(ImageSourceKind.LocalFile, resolution.Kind);
    }

    [Fact]
    public void Loader_ReadsLocalFileUnderRoot()
    {
        using var temp = new TempDirectory();
        var file = temp.File("pic.png");
        var bytes = TestImages.Png(8, 8);
        File.WriteAllBytes(file, bytes);

        var asset = new ImageAssetLoader().Load(
            "pic.png",
            new ImageSourceOptions { AllowedRoot = temp.Path });

        Assert.Equal(bytes, asset.Bytes);
        Assert.Equal(DocxImageMediaType.Png, asset.MediaType);
        Assert.Equal(Path.GetFileName(file), asset.DisplayName);
    }

    [Fact]
    public void Loader_FileExceedingByteLimit_IsRejected()
    {
        using var temp = new TempDirectory();
        var file = temp.File("big.png");
        File.WriteAllBytes(file, TestImages.Png(64, 64));

        var exception = Assert.Throws<ImageSourceException>(() =>
            new ImageAssetLoader().Load(
                "big.png",
                new ImageSourceOptions { AllowedRoot = temp.Path },
                new ImageAssetOptions { MaxEncodedBytes = 64 }));
        Assert.Equal(ImageSourceErrorCode.EncodedBytesExceededLimit, exception.Code);
    }

    [Fact]
    public void Loader_PercentEncodedPayload_DecodesAndSniffs()
    {
        var dataUri = TestImages.DataUriPercentEncoded(TestImages.Png(6, 6));
        var asset = new ImageAssetLoader().Load(dataUri);

        Assert.Equal(6, asset.Width);
        Assert.Equal(6, asset.Height);
    }

    [Fact]
    public void Loader_MalformedPercentEscape_IsRejected()
    {
        var exception = Assert.Throws<ImageSourceException>(() =>
            new ImageAssetLoader().Load("data:image/png,%zz"));
        Assert.Equal(ImageSourceErrorCode.InvalidDataUri, exception.Code);
    }

    [Fact]
    public void Loader_PayloadAfterLengthLimit_IsRejected()
    {
        var bytes = TestImages.Png(64, 64);
        var dataUri = TestImages.DataUriBase64(bytes);

        var exception = Assert.Throws<ImageSourceException>(() =>
            new ImageAssetLoader().Load(dataUri, assetOptions: new ImageAssetOptions { MaxEncodedBytes = bytes.Length - 4 }));
        Assert.Equal(ImageSourceErrorCode.EncodedBytesExceededLimit, exception.Code);
    }

    [Fact]
    public void ImageAsset_BytesAreDefensivelyCopied()
    {
        var original = TestImages.Png(4, 4);
        var bytes = (byte[])original.Clone();
        var asset = new ImageAssetLoader().LoadFromBytes(bytes);

        // Mutating the caller's buffer cannot change the asset's stored bytes or hash.
        bytes[0] = 0xFF;

        Assert.Equal(original, asset.Bytes);
        Assert.Equal(Convert.ToHexStringLower(SHA256.HashData(original)), asset.ContentHash);
    }
}
