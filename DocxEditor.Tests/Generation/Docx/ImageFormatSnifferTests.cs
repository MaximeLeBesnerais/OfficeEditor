using System.Buffers.Binary;
using System.Security.Cryptography;
using DocxEditor.Core.Generation.Assets;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Content-based image format sniffing and asset loading: detection and header metadata for
/// every supported format (PNG/JPEG/GIF/BMP/TIFF/SVG), truncation and hostile-header
/// rejection, DPI inference and fallback, and the loader's limits/warnings/hashing.
/// </summary>
public class ImageFormatSnifferTests
{
    [Theory]
    [InlineData(DocxImageMediaType.Png)]
    [InlineData(DocxImageMediaType.Jpeg)]
    [InlineData(DocxImageMediaType.Gif)]
    [InlineData(DocxImageMediaType.Bmp)]
    [InlineData(DocxImageMediaType.Tiff)]
    [InlineData(DocxImageMediaType.Svg)]
    public void Detect_RecognizesEverySupportedFormat(DocxImageMediaType expected)
    {
        var bytes = expected switch
        {
            DocxImageMediaType.Png => TestImages.Png(4, 4),
            DocxImageMediaType.Jpeg => TestImages.Jpeg(4, 4),
            DocxImageMediaType.Gif => TestImages.Gif(4, 4),
            DocxImageMediaType.Bmp => TestImages.Bmp(4, 4),
            DocxImageMediaType.Tiff => TestImages.Tiff(4, 4),
            DocxImageMediaType.Svg => TestImages.Svg(4, 4),
            _ => throw new ArgumentOutOfRangeException()
        };

        Assert.Equal(expected, ImageFormatSniffer.TryDetectFormat(bytes));
    }

    [Fact]
    public void Png_ReadsDimensionsAndPhyDpi()
    {
        var metadata = ImageFormatSniffer.TryReadMetadata(TestImages.Png(200, 100, dpiX: 72, dpiY: 72));

        Assert.NotNull(metadata);
        Assert.Equal(DocxImageMediaType.Png, metadata!.MediaType);
        Assert.Equal(200, metadata.Width);
        Assert.Equal(100, metadata.Height);
        Assert.NotNull(metadata.DpiX);
        Assert.InRange(metadata.DpiX!.Value, 71.9, 72.1);
        Assert.InRange(metadata.DpiY!.Value, 71.9, 72.1);
        Assert.Empty(metadata.Warnings);
    }

    [Fact]
    public void Png_WithoutPhy_HasNoDpiAndWarnsOnLoad()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Png(200, 100));

        Assert.Equal(DocxImageMediaType.Png, asset.MediaType);
        Assert.Null(asset.DpiX);
        Assert.Equal(150.0, asset.NaturalWidthPt, 3);  // 200 / 96 × 72
        Assert.Equal(75.0, asset.NaturalHeightPt, 3);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.IntrinsicDpiUnavailable);
    }

    [Fact]
    public void Png_WithPhy_ComputesIntrinsicSize()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Png(200, 100, dpiX: 72, dpiY: 72));

        Assert.Equal(200.0, asset.NaturalWidthPt, 1);
        Assert.Equal(100.0, asset.NaturalHeightPt, 1);
        Assert.DoesNotContain(asset.Warnings, w => w.Code == ImageAssetWarningCode.IntrinsicDpiUnavailable);
    }

    [Fact]
    public void Png_InconsistentDpi_WarnsAspectInconsistency()
    {
        var bytes = TestImages.Png(200, 100, dpiX: 72, dpiY: 36);
        var asset = new ImageAssetLoader().LoadFromBytes(bytes);

        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.AspectDpiInconsistent);
    }

    [Theory]
    [InlineData(1, 72.0, 200.0)]   // dots per inch
    [InlineData(2, 254.0, 56.69)]  // dots per centimeter → ×2.54
    public void Jpeg_DensityUnits_MapToDpi(int units, double expectedDpi, double expectedNaturalWidth)
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Jpeg(200, 100, density: units == 1 ? 72 : 100, units: units));

        Assert.Equal(DocxImageMediaType.Jpeg, asset.MediaType);
        Assert.Equal(expectedDpi, asset.DpiX!.Value, 1);
        Assert.Equal(expectedNaturalWidth, asset.NaturalWidthPt, 1);
    }

    [Fact]
    public void Jpeg_UnspecifiedDensity_WarnsAndAssumes96Dpi()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Jpeg(200, 100, density: 0, units: 0));

        Assert.Null(asset.DpiX);
        Assert.Equal(150.0, asset.NaturalWidthPt, 3);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.DensityUnitsUnspecified);
    }

    [Fact]
    public void Gif_CarriesNoDpi_AndAssumes96()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Gif(100, 50));

        Assert.Equal(DocxImageMediaType.Gif, asset.MediaType);
        Assert.Equal(100, asset.Width);
        Assert.Equal(50, asset.Height);
        Assert.Null(asset.DpiX);
        Assert.Equal(75.0, asset.NaturalWidthPt, 3);
    }

    [Fact]
    public void Bmp_ReadsResolutionAndTopDownHeight()
    {
        var metadata = ImageFormatSniffer.TryReadMetadata(TestImages.Bmp(200, -100, pixelsPerMeterX: 2835, pixelsPerMeterY: 2835));

        Assert.NotNull(metadata);
        Assert.Equal(DocxImageMediaType.Bmp, metadata!.MediaType);
        Assert.Equal(200, metadata.Width);
        Assert.Equal(100, metadata.Height); // absolute value for top-down bitmaps
        Assert.InRange(metadata.DpiX!.Value, 71.9, 72.1);
    }

    [Fact]
    public void Bmp_CoreHeader_HasNoResolutionAndWarns()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.BmpCore(100, 50));

        Assert.Equal(DocxImageMediaType.Bmp, asset.MediaType);
        Assert.Null(asset.DpiX);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.IntrinsicDpiUnavailable);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Tiff_BothByteOrders_ReadDimensionsAndDpi(bool littleEndian)
    {
        var metadata = ImageFormatSniffer.TryReadMetadata(TestImages.Tiff(200, 100, littleEndian, dpiX: 72, dpiY: 72, resolutionUnit: 2));

        Assert.NotNull(metadata);
        Assert.Equal(DocxImageMediaType.Tiff, metadata!.MediaType);
        Assert.Equal(200, metadata.Width);
        Assert.Equal(100, metadata.Height);
        Assert.Equal(72.0, metadata.DpiX!.Value, 1);
        Assert.Equal(72.0, metadata.DpiY!.Value, 1);
        Assert.Empty(metadata.Warnings);
    }

    [Fact]
    public void Tiff_CentimeterUnit_ConvertsToDpi()
    {
        var metadata = ImageFormatSniffer.TryReadMetadata(TestImages.Tiff(200, 100, littleEndian: true, dpiX: 100, dpiY: 100, resolutionUnit: 3));

        Assert.NotNull(metadata);
        Assert.Equal(254.0, metadata!.DpiX!.Value, 1);
    }

    [Fact]
    public void Tiff_UnspecifiedResolutionUnit_Warns()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Tiff(200, 100, littleEndian: true, dpiX: 72, dpiY: 72, resolutionUnit: 1));

        Assert.Null(asset.DpiX);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.ResolutionUnitUnspecified);
    }

    [Fact]
    public void Svg_ReadsDimensionsAt96Dpi()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Svg(200, 100));

        Assert.Equal(DocxImageMediaType.Svg, asset.MediaType);
        Assert.Equal(200, asset.Width);
        Assert.Equal(100, asset.Height);
        Assert.Equal(96.0, asset.DpiX!.Value);
        Assert.Equal(150.0, asset.NaturalWidthPt, 3);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.SvgRasterFallbackMissing);
    }

    [Fact]
    public void Svg_ViewBoxOnly_InferDimensions()
    {
        var metadata = ImageFormatSniffer.TryReadMetadata(TestImages.SvgViewBox(300, 200));

        Assert.NotNull(metadata);
        Assert.Equal(300, metadata!.Width);
        Assert.Equal(200, metadata.Height);
    }

    [Fact]
    public void Svg_AbsoluteUnits_ConvertToPixels()
    {
        var svg = """
            <svg xmlns="http://www.w3.org/2000/svg" width="72pt" height="36pt"/>
            """u8.ToArray();

        var metadata = ImageFormatSniffer.TryReadMetadata(svg);

        Assert.NotNull(metadata);
        Assert.Equal(96, metadata!.Width);   // 72pt → 96 CSS px
        Assert.Equal(48, metadata.Height);   // 36pt → 48 CSS px
    }

    [Theory]
    [InlineData("truncated png signature")]
    [InlineData("")]
    public void UnrecognizedBytes_ReturnNull(string payload)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(payload);
        Assert.Null(ImageFormatSniffer.TryReadMetadata(bytes));
        Assert.Null(ImageFormatSniffer.TryDetectFormat(bytes));
    }

    [Fact]
    public void TruncatedPng_ReturnsNull()
    {
        var full = TestImages.Png(8, 8);
        Assert.Null(ImageFormatSniffer.TryReadMetadata(full[..8]));   // signature only
        Assert.Null(ImageFormatSniffer.TryReadMetadata(full[..16]));  // no IHDR dimensions
    }

    [Fact]
    public void PngWithZeroDimensions_ReturnsNull()
    {
        var bytes = TestImages.Png(4, 4);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), 0);
        Assert.Null(ImageFormatSniffer.TryReadMetadata(bytes));
    }

    [Fact]
    public void TruncatedJpeg_ReturnsNull()
    {
        Assert.Null(ImageFormatSniffer.TryReadMetadata(TestImages.Jpeg(8, 8)[..2])); // SOI only

        // A JPEG whose APP0 length runs past the buffer end.
        var hostile = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x40 };
        Assert.Null(ImageFormatSniffer.TryReadMetadata(hostile));
    }

    [Fact]
    public void TruncatedGifAndBmp_ReturnNull()
    {
        Assert.Null(ImageFormatSniffer.TryReadMetadata(TestImages.Gif(4, 4)[..6]));
        Assert.Null(ImageFormatSniffer.TryReadMetadata(TestImages.Bmp(4, 4)[..2]));
    }

    [Fact]
    public void HostileTiff_IfdOffsetBeyondBuffer_ReturnsNull()
    {
        var bytes = new byte[8];
        bytes[0] = 0x49;
        bytes[1] = 0x49;
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(2, 2), 42);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4, 4), 100_000);

        Assert.Null(ImageFormatSniffer.TryReadMetadata(bytes));
    }

    [Fact]
    public void HostileTiff_HugeEntryCount_ReturnsNull()
    {
        var bytes = new byte[10];
        bytes[0] = 0x49;
        bytes[1] = 0x49;
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(2, 2), 42);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4, 4), 8);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(8, 2), 0xFFFF);

        Assert.Null(ImageFormatSniffer.TryReadMetadata(bytes));
    }

    [Fact]
    public void TiffWithoutDimensions_ReturnsNull()
    {
        // Big-endian TIFF with an empty IFD: entry count 0 → no dimensions.
        var bytes = new byte[14];
        bytes[0] = 0x4D;
        bytes[1] = 0x4D;
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(2, 2), 42);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(4, 4), 8);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(8, 2), 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(10, 4), 0);

        Assert.Null(ImageFormatSniffer.TryReadMetadata(bytes));
    }

    [Fact]
    public void Load_ComputesHashAndDeterministicDisplayName()
    {
        var bytes = TestImages.Png(4, 4);
        var asset = new ImageAssetLoader().LoadFromBytes(bytes);

        Assert.Equal(Convert.ToHexStringLower(SHA256.HashData(bytes)), asset.ContentHash);
        Assert.Equal($"image_{asset.ContentHash[..12]}.png", asset.DisplayName);
        Assert.Equal(bytes, asset.Bytes);
    }

    [Fact]
    public void Load_DeclaredMediaTypeMismatch_WarnsButLoads()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Png(4, 4), declaredMediaType: "image/jpeg");

        Assert.Equal(DocxImageMediaType.Png, asset.MediaType);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.DeclaredMediaTypeMismatch);
    }

    [Fact]
    public void Load_MatchingDeclaredMediaType_DoesNotWarn()
    {
        var asset = new ImageAssetLoader().LoadFromBytes(TestImages.Png(4, 4), declaredMediaType: "image/png");
        Assert.DoesNotContain(asset.Warnings, w => w.Code == ImageAssetWarningCode.DeclaredMediaTypeMismatch);
    }

    [Fact]
    public void Load_DataUri_ProducesSameAsset()
    {
        var bytes = TestImages.Png(4, 4);
        var fromUri = new ImageAssetLoader().Load(TestImages.DataUriBase64(bytes));
        var fromBytes = new ImageAssetLoader().LoadFromBytes(bytes);

        Assert.Equal(fromBytes.ContentHash, fromUri.ContentHash);
        Assert.Equal(fromBytes.MediaType, fromUri.MediaType);
        Assert.Equal(fromBytes.Width, fromUri.Width);
    }

    [Fact]
    public void Load_EmptyBytes_ThrowsInvalidImageData()
    {
        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().LoadFromBytes(Array.Empty<byte>()));
        Assert.Equal(ImageSourceErrorCode.InvalidImageData, exception.Code);
    }

    [Fact]
    public void Load_UnsupportedPayload_ThrowsUnsupportedMediaType()
    {
        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().LoadFromBytes(System.Text.Encoding.UTF8.GetBytes("hello world")));
        Assert.Equal(ImageSourceErrorCode.UnsupportedMediaType, exception.Code);
    }

    [Fact]
    public void Load_PixelDimensionLimit_ThrowsDimensionsExceedLimits()
    {
        var options = new ImageAssetOptions { MaxPixelDimension = 100, MaxPixelArea = 100_000 };
        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().LoadFromBytes(TestImages.Png(200, 100), assetOptions: options));
        Assert.Equal(ImageSourceErrorCode.DimensionsExceedLimits, exception.Code);
    }

    [Fact]
    public void Load_PixelAreaLimit_ThrowsDimensionsExceedLimits()
    {
        var options = new ImageAssetOptions { MaxPixelDimension = 100_000, MaxPixelArea = 10_000 };
        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().LoadFromBytes(TestImages.Png(200, 100), assetOptions: options));
        Assert.Equal(ImageSourceErrorCode.DimensionsExceedLimits, exception.Code);
    }

    [Fact]
    public void Load_DisabledMediaType_ThrowsUnsupportedMediaType()
    {
        var options = new ImageAssetOptions { SupportedMediaTypes = new HashSet<DocxImageMediaType> { DocxImageMediaType.Jpeg } };
        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().LoadFromBytes(TestImages.Png(4, 4), assetOptions: options));
        Assert.Equal(ImageSourceErrorCode.UnsupportedMediaType, exception.Code);
    }

    [Fact]
    public void Load_EncodedBytesLimit_RejectsBeforeDecoding()
    {
        var options = new ImageAssetOptions { MaxEncodedBytes = 16 };
        var dataUri = TestImages.DataUriBase64(TestImages.Png(32, 32));

        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().Load(dataUri, assetOptions: options));
        Assert.Equal(ImageSourceErrorCode.EncodedBytesExceededLimit, exception.Code);
    }

    [Fact]
    public void Load_MalformedBase64_ThrowsInvalidDataUri()
    {
        var exception = Assert.Throws<ImageSourceException>(
            () => new ImageAssetLoader().Load("data:image/png;base64,@@@not-base64@@@"));
        Assert.Equal(ImageSourceErrorCode.InvalidDataUri, exception.Code);
    }

    [Fact]
    public void Load_FileExtensionIsNotTrusted_ContentWins()
    {
        // A JPEG payload stored under a .png name must be detected as JPEG.
        var asset = new ImageAssetLoader().LoadFromBytes(
            TestImages.Jpeg(8, 8),
            declaredMediaType: DocxImageMediaTypes.GetContentType(DocxImageMediaType.Png),
            displayName: "fake.png");

        Assert.Equal(DocxImageMediaType.Jpeg, asset.MediaType);
        Assert.Contains(asset.Warnings, w => w.Code == ImageAssetWarningCode.DeclaredMediaTypeMismatch);
    }
}
