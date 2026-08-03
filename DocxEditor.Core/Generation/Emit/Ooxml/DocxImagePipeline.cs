using DocumentFormat.OpenXml.Packaging;
using DocxEditor.Core.Generation.Assets;
using DocxEditor.Core.Generation.Emit.Ooxml.Images;
using DocxEditor.Core.Generation.Emit.Ooxml.Positioned;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Core.Generation.Emit.Ooxml;

/// <summary>
/// Per-document adapter between asset loading and OOXML image relationships. It caches loaded
/// sources, delegates package deduplication to <see cref="DocxImagePartManager"/>, and maps asset
/// warnings into the generation issue channel.
/// </summary>
internal sealed class DocxImagePipeline : IPositionedImageResolver
{
    private readonly ImageAssetLoader _loader;
    private readonly ImageSourceOptions _sourceOptions;
    private readonly ImageAssetOptions _assetOptions;
    private readonly IList<DocxGenerationIssue> _warnings;
    private readonly Dictionary<string, ImageAsset> _assets = new(StringComparer.Ordinal);
    private readonly Dictionary<OpenXmlPart, DocxImagePartManager> _partManagers = [];
    private readonly HashSet<string> _reportedAssetWarnings = new(StringComparer.Ordinal);

    public DocxImagePipeline(
        ImageAssetLoader loader,
        ImageSourceOptions sourceOptions,
        ImageAssetOptions assetOptions,
        IList<DocxGenerationIssue> warnings)
    {
        _loader = loader;
        _sourceOptions = sourceOptions;
        _assetOptions = assetOptions;
        _warnings = warnings;
    }

    public (ImageAsset Asset, RegisteredImagePart Part) Resolve(OpenXmlPart owningPart, string source, string path)
    {
        ArgumentNullException.ThrowIfNull(owningPart);
        var asset = Load(source, path);
        return (asset, GetPartManager(owningPart).Register(asset));
    }

    public PositionedImagePlacement? Resolve(OpenXmlPart owningPart, PositionedImage image, string? path)
    {
        ArgumentNullException.ThrowIfNull(owningPart);
        ArgumentNullException.ThrowIfNull(image);
        var (asset, part) = Resolve(owningPart, image.Source, path ?? "$.positioned.image");
        var geometry = ResolvedImageGeometry.Resolve(
            asset, image.Fit, image.Crop, image.Position.WidthPt, image.Position.HeightPt);
        return new PositionedImagePlacement { EmbedId = part.RelationshipId, Geometry = geometry };
    }

    private ImageAsset Load(string source, string path)
    {
        if (_assets.TryGetValue(source, out var cached))
        {
            return cached;
        }

        var asset = _loader.Load(source, _sourceOptions, _assetOptions);
        _assets[source] = asset;
        if (_reportedAssetWarnings.Add(source))
        {
            foreach (var warning in asset.Warnings)
            {
                _warnings.Add(new DocxGenerationIssue(
                    path,
                    $"[{warning.Code}] {warning.Message}",
                    null,
                    DocxGenerationIssueSeverity.Warning));
            }
        }
        return asset;
    }

    private DocxImagePartManager GetPartManager(OpenXmlPart part)
    {
        if (_partManagers.TryGetValue(part, out var manager))
        {
            return manager;
        }

        manager = part switch
        {
            MainDocumentPart mainPart => new DocxImagePartManager(mainPart),
            HeaderPart headerPart => new DocxImagePartManager(headerPart),
            FooterPart footerPart => new DocxImagePartManager(footerPart),
            _ => throw new NotSupportedException(
                $"Image relationships are not supported for part type '{part.GetType().Name}'.")
        };
        _partManagers.Add(part, manager);
        return manager;
    }
}
