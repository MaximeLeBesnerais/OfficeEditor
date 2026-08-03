using System.Reflection;
using System.Text.Json;

namespace OfficeEditor.Core.Rendering;

/// <summary>Single entry point for rendering any supported document source to PDF/PNG/SVG/Typst.</summary>
public interface IDocumentRenderer
{
    DocumentRenderResult Render(DocumentRenderRequest request);
}

/// <summary>
/// The shared render facade. All surfaces (CLI/API/MCP) call this one library API instead of
/// per-format bespoke handlers. It dispatches by source extension to format-specific renderers
/// that live in the format core assemblies and are discovered at construction from loaded
/// assemblies via <see cref="DocumentSourceFormatAttribute"/> — no explicit wiring required.
///
/// JSON dispatch: a .json source has no intrinsic format marker, so the facade probes each
/// format renderer's own generation-vocabulary validator in priority order — XLSX
/// ("worksheets") first, then DOCX ("sections"), then PPTX ("slides") — and delegates to the
/// first match. When no vocabulary matches, a clear error is returned.
///
/// The facade never throws for request or rendering failures; they surface as a failed
/// <see cref="DocumentRenderResult"/> with <see cref="DocumentRenderResult.ErrorMessage"/>.
/// </summary>
public sealed class DocumentRenderer : IDocumentRenderer
{
    private static readonly string[] KnownFormatAssemblies = ["PptxEditor.Core", "DocxEditor.Core", "XlsxEditor.Core"];

    private readonly IReadOnlyList<RendererRegistration> _renderers;

    public DocumentRenderer()
        : this(DiscoverRenderers())
    {
    }

    private DocumentRenderer(IReadOnlyList<RendererRegistration> registrations)
    {
        _renderers = registrations
            .OrderBy(registration => registration.JsonPriority)
            .ToList();
    }

    /// <summary>
    /// Creates a facade over an explicit renderer list (ordered by JSON priority). The parameterless
    /// constructor discovers all renderers from loaded assemblies; this overload exists for
    /// deterministic wiring in hosts that prefer explicit composition.
    /// </summary>
    internal DocumentRenderer(IEnumerable<IFormatRenderer> renderers)
        : this(BuildRegistrations(renderers))
    {
    }

    private static IReadOnlyList<RendererRegistration> BuildRegistrations(IEnumerable<IFormatRenderer> renderers)
    {
        ArgumentNullException.ThrowIfNull(renderers);
        return renderers.Select(CreateRegistration).ToList();
    }

    public DocumentRenderResult Render(DocumentRenderRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (string.IsNullOrWhiteSpace(request.SourcePath))
        {
            return Failure("SourcePath must be a non-empty file path.");
        }

        if (!File.Exists(request.SourcePath))
        {
            return Failure($"Source file not found: '{request.SourcePath}'.");
        }

        var extension = Path.GetExtension(request.SourcePath);
        if (string.Equals(extension, ".json", StringComparison.OrdinalIgnoreCase))
        {
            return RenderJson(request);
        }

        foreach (var registration in _renderers)
        {
            if (string.Equals(registration.SourceExtension, extension, StringComparison.OrdinalIgnoreCase))
            {
                return SafeRender(registration.Renderer, request);
            }
        }

        return Failure($"Unsupported source type '{extension}'. Supported sources: .pptx, .docx, .xlsx, .json.");
    }

    /// <summary>
    /// Routes a .json source to the format whose generation vocabulary matches. Probes XLSX first,
    /// then DOCX, then PPTX (the documented heuristic; see the class summary).
    /// </summary>
    private DocumentRenderResult RenderJson(DocumentRenderRequest request)
    {
        string json;
        try
        {
            json = File.ReadAllText(request.SourcePath);
        }
        catch (Exception ex)
        {
            return Failure($"Cannot read JSON source '{request.SourcePath}': {ex.Message}");
        }

        try
        {
            using var _ = JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            return Failure($"Malformed JSON source: {ex.Message}");
        }

        foreach (var registration in _renderers)
        {
            try
            {
                if (registration.Renderer.CanRenderSource(request.SourcePath))
                {
                    return SafeRender(registration.Renderer, request);
                }
            }
            catch (Exception ex)
            {
                // A probing failure must never mask a later vocabulary that could still match.
                _ = ex;
            }
        }

        return Failure(
            "Unrecognized JSON source: the document does not match any supported generation vocabulary " +
            "(XLSX 'worksheets', DOCX 'sections', or PPTX 'slides' generation JSON).");
    }

    private static DocumentRenderResult SafeRender(IFormatRenderer renderer, DocumentRenderRequest request)
    {
        try
        {
            return renderer.Render(request);
        }
        catch (Exception ex)
        {
            return Failure($"Render failed: {ex.Message}");
        }
    }

    private static DocumentRenderResult Failure(string message) => new()
    {
        Success = false,
        ErrorMessage = message
    };

    private static IReadOnlyList<RendererRegistration> DiscoverRenderers()
    {
        var registrations = new List<RendererRegistration>();
        foreach (var type in EnumerateRendererTypes())
        {
            if (type.GetCustomAttribute<DocumentSourceFormatAttribute>(inherit: false) is null)
            {
                continue;
            }

            IFormatRenderer renderer;
            try
            {
                renderer = (IFormatRenderer)Activator.CreateInstance(type)!;
            }
            catch
            {
                // A renderer whose constructor throws cannot serve this facade; skip it.
                continue;
            }

            registrations.Add(CreateRegistration(renderer));
        }

        return registrations;
    }

    private static IEnumerable<Type> EnumerateRendererTypes()
    {
        // Load the format-core assemblies up front so discovery is independent of the order in
        // which the host happens to touch types from them. Load is a no-op when already loaded.
        foreach (var assemblyName in KnownFormatAssemblies)
        {
            try
            {
                Assembly.Load(new AssemblyName(assemblyName));
            }
            catch
            {
                // The format core is optional; without it that format simply cannot render.
            }
        }

        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            foreach (var type in SafeGetTypes(assembly))
            {
                if (type is { IsClass: true, IsAbstract: false } && typeof(IFormatRenderer).IsAssignableFrom(type))
                {
                    yield return type;
                }
            }
        }
    }

    private static IEnumerable<Type> SafeGetTypes(Assembly assembly)
    {
        try
        {
            return assembly.GetTypes();
        }
        catch (ReflectionTypeLoadException ex)
        {
            return ex.Types.Where(type => type is not null)!;
        }
    }

    private static RendererRegistration CreateRegistration(IFormatRenderer renderer)
    {
        var attribute = renderer.GetType().GetCustomAttribute<DocumentSourceFormatAttribute>(inherit: false);
        return new RendererRegistration(
            renderer,
            attribute?.SourceExtension,
            attribute?.JsonPriority ?? int.MaxValue);
    }

    private sealed record RendererRegistration(IFormatRenderer Renderer, string? SourceExtension, int JsonPriority);
}
