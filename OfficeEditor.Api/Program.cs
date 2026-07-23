using System.Text.Json;
using OfficeEditor.Api.Models;
using OfficeEditor.Api.Services;
using PptxEditor.Core.Builders;

var builder = WebApplication.CreateBuilder(args);

const string ReactCorsPolicy = "ReactCorsPolicy";

builder.Services.AddCors(options =>
{
    options.AddPolicy(ReactCorsPolicy, policy =>
    {
        policy
            .WithOrigins(
                "http://localhost:5173",
                "https://localhost:5173")
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

builder.Services.AddMemoryCache();
builder.Services.AddSingleton<IConversionResultStore, InMemoryConversionResultStore>();
builder.Services.AddSingleton<IConversionService, ConversionService>();
builder.Services.AddSingleton<ISampleFileService, SampleFileService>();
builder.Services.AddSingleton<IPreviewService, PreviewService>();
builder.Services.AddSingleton<IDeckSessionStore, InMemoryDeckSessionStore>();
builder.Services.AddSingleton<ISlideRenderer, BuilderSlideRenderer>();
builder.Services.AddSingleton<IDeckPreviewService, DeckPreviewService>();
builder.Services.AddSingleton<IDeckService, DeckService>();
builder.Services.AddSingleton<IDeckEditService, DeckEditService>();
builder.Services.AddSingleton<IDeckGenerationService>(
    _ => new DeckGenerationService(builder.Configuration["Demo:FontDirectory"]));
builder.Services.AddSingleton<IDemoDeckService>(sp => new DemoDeckService(
    sp.GetRequiredService<IDeckSessionStore>(),
    builder.Configuration["Demo:FontDirectory"]));
builder.Services.AddSingleton<ILibreOfficeCompareService, LibreOfficeCompareService>();
builder.Services.AddSingleton<IOfficialRenderService, OfficialRenderService>();
builder.Services.AddSingleton(sp => new RenderWarmupService(
    new OfficeEditor.Core.Services.TypstCompilerService(),
    builder.Configuration["Demo:FontDirectory"],
    sp.GetRequiredService<ILogger<RenderWarmupService>>()));
builder.Services.AddHostedService(sp => sp.GetRequiredService<RenderWarmupService>());

var app = builder.Build();

app.UseCors(ReactCorsPolicy);

app.MapGet("/api/health", () => Results.Json(new { status = "ok" }));

app.MapGet("/api/samples", (ISampleFileService sampleFileService) =>
{
    var samples = sampleFileService.GetSampleFiles()
        .Select(s => new SampleFileDto(s.Name, s.Format.ToString().ToLowerInvariant(), s.Description))
        .ToList();

    return Results.Ok(samples);
});

app.MapPost("/api/convert", async (
    HttpContext context,
    IConversionService conversionService,
    ISampleFileService sampleFileService,
    IConversionResultStore resultStore,
    IPreviewService previewService,
    CancellationToken ct) =>
{
    var form = await context.Request.ReadFormAsync(ct);

    var targetFormatField = form["targetFormat"].FirstOrDefault();
    if (string.IsNullOrWhiteSpace(targetFormatField) ||
        !TryParseTargetFormat(targetFormatField, out var targetFormat))
    {
        return Results.BadRequest(new ConvertResponse(
            Success: false,
            ErrorMessage: $"Invalid or missing targetFormat. Supported values: pdf, png, svg, docx, pptx, xlsx."));
    }

    var sampleName = form["sampleName"].FirstOrDefault();
    var file = form.Files.GetFile("file");

    byte[] sourceBytes;
    string sourceFileName;

    if (file is not null)
    {
        await using var stream = file.OpenReadStream();
        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, ct);
        sourceBytes = memoryStream.ToArray();
        sourceFileName = file.FileName;
    }
    else if (!string.IsNullOrWhiteSpace(sampleName))
    {
        try
        {
            sourceBytes = await sampleFileService.LoadAsync(sampleName, ct);
            sourceFileName = sampleName;
        }
        catch (FileNotFoundException ex)
        {
            return Results.BadRequest(new ConvertResponse(
                Success: false,
                ErrorMessage: ex.Message));
        }
    }
    else
    {
        return Results.BadRequest(new ConvertResponse(
            Success: false,
            ErrorMessage: "Either a 'file' upload or a 'sampleName' must be provided."));
    }

    var options = ParseOptions(form["options"].FirstOrDefault());

    var request = new ConversionRequest(sourceBytes, sourceFileName, targetFormat, options);
    var result = await conversionService.ConvertAsync(request, ct);

    if (!result.Success || result.OutputBytes is null)
    {
        return Results.BadRequest(new ConvertResponse(
            Success: false,
            ErrorMessage: result.ErrorMessage ?? "Conversion produced no output.",
            Messages: result.Messages));
    }

    var id = resultStore.Store(result.OutputBytes, result.ContentType, result.OutputFileName);

    var baseUrl = $"{context.Request.Scheme}://{context.Request.Host}";
    var downloadUrl = $"{baseUrl}/api/download/{id}";
    var previewUrl = previewService.CanPreview(result.ContentType)
        ? $"{baseUrl}/api/preview/{id}"
        : null;

    return Results.Ok(new ConvertResponse(
        Success: true,
        OutputFileName: result.OutputFileName,
        ContentType: result.ContentType,
        DownloadUrl: downloadUrl,
        PreviewUrl: previewUrl,
        Messages: result.Messages));
});

app.MapGet("/api/download/{id:guid}", (Guid id, IConversionResultStore resultStore) =>
{
    if (!resultStore.TryGet(id, out var result) || result is null)
    {
        return Results.NotFound();
    }

    return Results.File(result.Bytes, result.ContentType, result.FileName);
});

app.MapGet("/api/preview/{id:guid}", (Guid id, IConversionResultStore resultStore, IPreviewService previewService) =>
{
    if (!resultStore.TryGet(id, out var result) || result is null)
    {
        return Results.NotFound();
    }

    if (!previewService.CanPreview(result.ContentType))
    {
        return Results.BadRequest(new { error = "This file type cannot be previewed in the browser." });
    }

    return Results.File(result.Bytes, result.ContentType, result.FileName, enableRangeProcessing: true);
});

app.MapPost("/api/decks", async (
    HttpContext context,
    ISampleFileService sampleFileService,
    IDeckSessionStore deckSessionStore,
    CancellationToken ct) =>
{
    var form = await context.Request.ReadFormAsync(ct);

    var sampleName = form["sampleName"].FirstOrDefault();
    var file = form.Files.GetFile("file");

    byte[] sourceBytes;
    string sourceFileName;

    if (file is not null)
    {
        if (!file.FileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            return Results.BadRequest(new { error = "Only PPTX uploads are supported for deck sessions." });
        }

        await using var stream = file.OpenReadStream();
        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, ct);
        sourceBytes = memoryStream.ToArray();
        sourceFileName = file.FileName;
    }
    else if (!string.IsNullOrWhiteSpace(sampleName))
    {
        try
        {
            sourceBytes = await sampleFileService.LoadAsync(sampleName, ct);
            sourceFileName = sampleName;
        }
        catch (FileNotFoundException ex)
        {
            return Results.BadRequest(new { error = ex.Message });
        }
    }
    else
    {
        return Results.BadRequest(new { error = "Either a 'file' upload or a 'sampleName' must be provided." });
    }

    // Open once to read the slide count, then keep only the bytes in the session
    // (memory-budget conscious: no permanently open PresentationDocument per deck).
    int slideCount;
    try
    {
        using var builder = PresentationBuilder.Open(sourceBytes);
        slideCount = builder.SlideCount;
    }
    catch (Exception ex)
    {
        // Boundary: any OpenXML/zip parse failure means the client sent a bad deck
        // (FileFormatException, InvalidDataException, ...); report 400, don't 500.
        return Results.BadRequest(new { error = $"The uploaded file is not a valid PPTX deck: {ex.Message}" });
    }

    if (slideCount < 1)
    {
        return Results.BadRequest(new { error = "The deck contains no slides." });
    }

    var deckId = deckSessionStore.Store(sourceBytes, sourceFileName, slideCount);
    return Results.Ok(new CreateDeckResponse(deckId, slideCount, sourceFileName));
});

app.MapGet("/api/decks/{deckId:guid}/slides/{n:int}/preview", async (
    Guid deckId,
    int n,
    HttpContext context,
    IDeckSessionStore deckSessionStore,
    IDeckPreviewService deckPreviewService,
    string? format,
    int? ppi,
    CancellationToken ct) =>
{
    if (!deckSessionStore.TryGet(deckId, out var session) || session is null)
    {
        return Results.NotFound(new { error = $"Unknown deck '{deckId}'." });
    }

    var slideError = DeckPreviewValidators.ValidateSlideNumber(n, session.SlideCount);
    if (slideError is not null)
    {
        return Results.BadRequest(new { error = slideError });
    }

    if (!DeckPreviewValidators.TryNormalizeFormat(format, out var normalizedFormat, out var formatError))
    {
        return Results.BadRequest(new { error = formatError });
    }

    var clampedPpi = DeckPreviewValidators.ClampPpi(ppi);

    SlidePreview preview;
    try
    {
        preview = await deckPreviewService.GetSlidePreviewAsync(session, n, normalizedFormat, clampedPpi, ct);
    }
    catch (InvalidOperationException ex)
    {
        // e.g. rendered page count != slide count inside the v0 full-render fallback.
        return Results.Problem(detail: ex.Message, statusCode: StatusCodes.Status500InternalServerError);
    }

    var headers = context.Response.Headers;
    headers.CacheControl = "private, max-age=60";
    headers.ETag = preview.ETag;
    // W3-INTEGRATION: per-slide warnings from the converter will populate this header
    // once W3's per-slide emission lands; previews return raw bytes, so headers are
    // the warnings channel. JSON-encoded string array, empty until then.
    headers["X-Slide-Warnings"] = JsonSerializer.Serialize(preview.Warnings);

    if (context.Request.Headers.IfNoneMatch.Count > 0
        && context.Request.Headers.IfNoneMatch.Any(tag => string.Equals(tag, preview.ETag, StringComparison.Ordinal)))
    {
        return Results.StatusCode(StatusCodes.Status304NotModified);
    }

    return Results.File(preview.Bytes, preview.ContentType);
});

app.MapGet("/api/decks/{id:guid}/anatomy", (Guid id, IDeckService deckService) =>
{
    var anatomy = deckService.GetAnatomy(id);
    return anatomy is null ? Results.NotFound() : Results.Ok(anatomy);
});

app.MapPost("/api/decks/generate", async (
    HttpContext context,
    IDeckGenerationService deckGenerationService,
    IDeckSessionStore deckSessionStore,
    CancellationToken ct) =>
{
    using var reader = new StreamReader(context.Request.Body);
    var body = await reader.ReadToEndAsync(ct);

    // Envelope mirrors the MCP deck_generate arguments:
    // { "document": {…generation JSON, plan.md §3.4…}, "previewFormat": "svg"|"png", "ppi": 150 }.
    string? documentJson = null;
    string? requestedFormat = null;
    int? requestedPpi = null;
    try
    {
        using var envelope = JsonDocument.Parse(body);
        if (envelope.RootElement.ValueKind == JsonValueKind.Object)
        {
            if (envelope.RootElement.TryGetProperty("document", out var document)
                && document.ValueKind == JsonValueKind.Object)
            {
                // Raw text keeps the P1 validator's error paths byte-accurate.
                documentJson = document.GetRawText();
            }
            if (envelope.RootElement.TryGetProperty("previewFormat", out var format)
                && format.ValueKind == JsonValueKind.String)
            {
                requestedFormat = format.GetString();
            }
            if (envelope.RootElement.TryGetProperty("ppi", out var ppi)
                && ppi.ValueKind == JsonValueKind.Number
                && ppi.TryGetInt32(out var ppiValue))
            {
                requestedPpi = ppiValue;
            }
        }
    }
    catch (JsonException)
    {
        // Falls through to the envelope error below.
    }

    if (documentJson is null)
    {
        return Results.BadRequest(new GenerateDeckResponse(
            Success: false,
            ErrorMessage: "Request body must be a JSON object with a 'document' property containing the generation document (plan.md §3.4)."));
    }

    // The generation surface is the SVG live-preview path (P9), so it defaults to svg —
    // unlike the deck-session preview endpoint, which defaults to png.
    if (!DeckPreviewValidators.TryNormalizeFormat(requestedFormat ?? "svg", out var normalizedFormat, out var formatError))
    {
        return Results.BadRequest(new GenerateDeckResponse(Success: false, ErrorMessage: formatError));
    }
    var clampedPpi = DeckPreviewValidators.ClampPpi(requestedPpi);

    var generation = deckGenerationService.Generate(documentJson, normalizedFormat, clampedPpi);
    if (!generation.Success)
    {
        return Results.BadRequest(new GenerateDeckResponse(
            Success: false,
            Errors: ToDtos(generation.Errors),
            Warnings: ToDtos(generation.Warnings),
            TotalMilliseconds: generation.TotalMilliseconds));
    }

    var deckId = deckSessionStore.Store(generation.PptxBytes!, "generated.pptx", generation.SlideCount);
    var downloadUrl = $"{context.Request.Scheme}://{context.Request.Host}/api/decks/{deckId}/file";

    return Results.Ok(new GenerateDeckResponse(
        Success: true,
        DeckId: deckId,
        SlideCount: generation.SlideCount,
        DownloadUrl: downloadUrl,
        Previews: generation.Previews
            .Select(p => new GeneratedSlidePreviewDto(p.Slide, p.Format, p.ContentType, Convert.ToBase64String(p.Bytes)))
            .ToList(),
        Warnings: ToDtos(generation.Warnings),
        PipelineWarnings: generation.PipelineWarnings,
        PreviewError: generation.PreviewError,
        GenerationMilliseconds: generation.GenerationMilliseconds,
        TotalMilliseconds: generation.TotalMilliseconds));
});

app.MapPost("/api/decks/{id:guid}/instructions", async (
    Guid id,
    HttpContext context,
    IDeckEditService deckEditService,
    CancellationToken ct) =>
{
    using var reader = new StreamReader(context.Request.Body);
    var instructionsJson = await reader.ReadToEndAsync(ct);

    var response = deckEditService.ApplyInstructions(id, instructionsJson);
    if (response is null)
    {
        return Results.NotFound();
    }

    return response.Success ? Results.Ok(response) : Results.BadRequest(response);
});

app.MapGet("/api/decks/{id:guid}/file", (Guid id, IDeckService deckService) =>
{
    var bytes = deckService.GetDeckBytes(id);
    if (bytes is null)
    {
        return Results.NotFound();
    }

    return Results.File(
        bytes,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        $"{id}.pptx");
});

app.MapGet("/api/demo/decks", (IDemoDeckService demoDeckService) =>
{
    var decks = demoDeckService.ListDecks()
        .Select(d => new DemoDeckDto(d.Name, d.FileName, d.Description, d.SlideCount))
        .ToList();

    return Results.Ok(new { decks });
});

app.MapGet("/api/demo/decks/{name}/official-slides", (
    string name,
    IDemoDeckService demoDeckService,
    IOfficialRenderService officialRenderService) =>
{
    if (demoDeckService is not DemoDeckService concreteDemoDeckService
        || !concreteDemoDeckService.TryGetDeckFile(name, out _))
    {
        return Results.NotFound(new { error = $"Unknown demo deck '{name}'." });
    }

    if (!officialRenderService.PdfToPpmAvailable)
    {
        return Results.Json(
            new { error = "pdftoppm not found; official-render comparison unavailable." },
            statusCode: StatusCodes.Status503ServiceUnavailable);
    }

    if (!officialRenderService.TryGetOfficialSlides(name, out var official))
    {
        return Results.NotFound(new { error = $"No official render for '{name}'." });
    }

    return Results.Ok(new OfficialSlidesResponse(
        official.Name,
        official.SlideCount,
        "PowerPoint PDF export",
        official.Pages
            .Select((bytes, index) => new GeneratedSlidePreviewDto(
                index + 1,
                "png",
                "image/png",
                Convert.ToBase64String(bytes)))
            .ToList()));
});

app.MapPost("/api/demo/render", async (
    HttpContext context,
    IDemoDeckService demoDeckService,
    CancellationToken ct) =>
{
    using var reader = new StreamReader(context.Request.Body);
    var body = await reader.ReadToEndAsync(ct);

    // Envelope: { "name": "northwind", "ppi": 110, "format": "svg" } — ppi and format
    // optional, parsed defensively like the /api/decks/generate envelope.
    string? name = null;
    int? requestedPpi = null;
    string? requestedFormat = null;
    try
    {
        using var envelope = JsonDocument.Parse(body);
        if (envelope.RootElement.ValueKind == JsonValueKind.Object)
        {
            if (envelope.RootElement.TryGetProperty("name", out var nameElement)
                && nameElement.ValueKind == JsonValueKind.String)
            {
                name = nameElement.GetString();
            }
            if (envelope.RootElement.TryGetProperty("ppi", out var ppiElement)
                && ppiElement.ValueKind == JsonValueKind.Number
                && ppiElement.TryGetInt32(out var ppiValue))
            {
                requestedPpi = ppiValue;
            }
            if (envelope.RootElement.TryGetProperty("format", out var formatElement)
                && formatElement.ValueKind == JsonValueKind.String)
            {
                requestedFormat = formatElement.GetString();
            }
        }
    }
    catch (JsonException)
    {
        // Falls through to the envelope error below.
    }

    if (string.IsNullOrWhiteSpace(name))
    {
        return Results.BadRequest(new { error = "Request body must be a JSON object with a 'name' property naming a demo deck (see GET /api/demo/decks)." });
    }

    if (!DeckPreviewValidators.TryNormalizeFormat(requestedFormat, out var normalizedFormat, out var formatError))
    {
        return Results.BadRequest(new { error = formatError });
    }

    DemoRenderResult result;
    try
    {
        result = demoDeckService.RenderDeck(name, requestedPpi ?? DeckPreviewValidators.DefaultPpi, normalizedFormat);
    }
    catch (ArgumentException ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
    catch (FileNotFoundException ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
    catch (Exception ex)
    {
        // Render failures (Typst backend errors, …) are surfaced, never swallowed.
        return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
    }

    return Results.Ok(new DemoRenderResponse(
        Success: true,
        DeckId: result.DeckId,
        SlideCount: result.SlideCount,
        TotalMilliseconds: result.TotalMilliseconds,
        Previews: result.Pages
            .Select((bytes, index) => new GeneratedSlidePreviewDto(
                index + 1,
                result.Format,
                DeckPreviewValidators.ContentTypeForFormat(result.Format),
                Convert.ToBase64String(bytes)))
            .ToList()));
});

app.MapPost("/api/demo/render-upload", async (
    HttpContext context,
    IDemoDeckService demoDeckService,
    CancellationToken ct) =>
{
    var form = await context.Request.ReadFormAsync(ct);

    var file = form.Files.GetFile("file");
    if (file is null)
    {
        return Results.BadRequest(new { error = "A 'file' form field with a .pptx upload is required." });
    }

    if (!file.FileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
    {
        return Results.BadRequest(new { error = "Only .pptx files are supported." });
    }

    if (!DeckPreviewValidators.TryNormalizeFormat(form["format"].FirstOrDefault(), out var normalizedFormat, out var formatError))
    {
        return Results.BadRequest(new { error = formatError });
    }

    var requestedPpi = int.TryParse(form["ppi"].FirstOrDefault(), out var ppiValue) ? ppiValue : (int?)null;
    var clampedPpi = DeckPreviewValidators.ClampPpi(requestedPpi);

    // Read the upload like POST /api/decks does (OpenReadStream -> MemoryStream).
    byte[] sourceBytes;
    await using (var stream = file.OpenReadStream())
    {
        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, ct);
        sourceBytes = memoryStream.ToArray();
    }

    DemoRenderResult result;
    try
    {
        result = demoDeckService.RenderUploadedDeck(sourceBytes, file.FileName, normalizedFormat, clampedPpi);
    }
    catch (ArgumentException ex)
    {
        // Client-side problems: invalid deck, empty deck, over the slide cap.
        return Results.BadRequest(new { error = ex.Message });
    }
    catch (Exception ex)
    {
        // Render failures (Typst backend errors, …) are surfaced, never swallowed.
        return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
    }

    return Results.Ok(new DemoRenderResponse(
        Success: true,
        DeckId: result.DeckId,
        SlideCount: result.SlideCount,
        TotalMilliseconds: result.TotalMilliseconds,
        Previews: result.Pages
            .Select((bytes, index) => new GeneratedSlidePreviewDto(
                index + 1,
                result.Format,
                DeckPreviewValidators.ContentTypeForFormat(result.Format),
                Convert.ToBase64String(bytes)))
            .ToList()));
});

app.MapGet("/api/demo/compare/capabilities", (ILibreOfficeCompareService libreOfficeCompareService) =>
{
    var probe = libreOfficeCompareService.Probe();
    return Results.Ok(new CompareCapabilitiesDto(
        probe.Available, probe.Version, probe.PdfToPpmAvailable, probe.SkipReason));
});

app.MapPost("/api/demo/compare/libreoffice", async (
    HttpContext context,
    IDemoDeckService demoDeckService,
    ILibreOfficeCompareService libreOfficeCompareService,
    IConversionResultStore resultStore,
    CancellationToken ct) =>
{
    using var reader = new StreamReader(context.Request.Body);
    var body = await reader.ReadToEndAsync(ct);

    // Envelope: { "name": "northwind", "ppi": 110 } — parsed defensively like /api/demo/render.
    string? name = null;
    int? requestedPpi = null;
    try
    {
        using var envelope = JsonDocument.Parse(body);
        if (envelope.RootElement.ValueKind == JsonValueKind.Object)
        {
            if (envelope.RootElement.TryGetProperty("name", out var nameElement)
                && nameElement.ValueKind == JsonValueKind.String)
            {
                name = nameElement.GetString();
            }
            if (envelope.RootElement.TryGetProperty("ppi", out var ppiElement)
                && ppiElement.ValueKind == JsonValueKind.Number
                && ppiElement.TryGetInt32(out var ppiValue))
            {
                requestedPpi = ppiValue;
            }
        }
    }
    catch (JsonException)
    {
        // Falls through to the envelope error below.
    }

    if (string.IsNullOrWhiteSpace(name))
    {
        return Results.BadRequest(new { error = "Request body must be a JSON object with a 'name' property naming a demo deck (see GET /api/demo/decks)." });
    }

    if (demoDeckService is not DemoDeckService concreteDemoDeckService
        || !concreteDemoDeckService.TryGetDeckFile(name, out var deckFile))
    {
        return Results.BadRequest(new { error = $"Unknown or missing demo deck '{name}'." });
    }

    var clampedPpi = DeckPreviewValidators.ClampPpi(requestedPpi);
    var deckBytes = await File.ReadAllBytesAsync(deckFile, ct);
    var deckFileName = Path.GetFileName(deckFile);

    // Open once for the real slide count: the client needs it for per-slide math even
    // when pdftoppm is missing and the rasterized page count is 0.
    int slideCount;
    using (var presentationBuilder = PresentationBuilder.Open(deckBytes))
    {
        slideCount = presentationBuilder.SlideCount;
    }

    // LibreOffice leg only: headless PDF conversion + pdftoppm rasterization. Never
    // throws for missing tools or conversion failures — failures ride on the result.
    var loResult = libreOfficeCompareService.RenderDeck(deckBytes, deckFileName, clampedPpi);
    return Results.Ok(ToLibreOfficeLegResponse(context, resultStore, loResult, name, slideCount));
});

app.MapPost("/api/demo/compare/libreoffice-upload", async (
    HttpContext context,
    ILibreOfficeCompareService libreOfficeCompareService,
    IConversionResultStore resultStore,
    CancellationToken ct) =>
{
    var form = await context.Request.ReadFormAsync(ct);

    // Validation mirrors POST /api/demo/render-upload.
    var file = form.Files.GetFile("file");
    if (file is null)
    {
        return Results.BadRequest(new { error = "A 'file' form field with a .pptx upload is required." });
    }

    if (!file.FileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
    {
        return Results.BadRequest(new { error = "Only .pptx files are supported." });
    }

    var requestedPpi = int.TryParse(form["ppi"].FirstOrDefault(), out var ppiValue) ? ppiValue : (int?)null;
    var clampedPpi = DeckPreviewValidators.ClampPpi(requestedPpi);

    byte[] sourceBytes;
    await using (var stream = file.OpenReadStream())
    {
        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, ct);
        sourceBytes = memoryStream.ToArray();
    }

    // Open once to reject garbage decks with a clear 400 instead of a soffice failure.
    int slideCount;
    try
    {
        using var builder = PresentationBuilder.Open(sourceBytes);
        slideCount = builder.SlideCount;
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = $"The uploaded file is not a valid PPTX deck: {ex.Message}" });
    }

    if (slideCount < 1)
    {
        return Results.BadRequest(new { error = "The deck contains no slides." });
    }

    if (slideCount > DemoDeckService.MaxUploadSlides)
    {
        return Results.BadRequest(new { error = $"Decks with more than {DemoDeckService.MaxUploadSlides} slides are not supported in the demo compare (got {slideCount})." });
    }

    var loResult = libreOfficeCompareService.RenderDeck(sourceBytes, file.FileName, clampedPpi);
    return Results.Ok(ToLibreOfficeLegResponse(context, resultStore, loResult, file.FileName, slideCount));
});

app.MapPost("/api/demo/compare/typst", async (
    HttpContext context,
    IDemoDeckService demoDeckService,
    IConversionResultStore resultStore,
    CancellationToken ct) =>
{
    using var reader = new StreamReader(context.Request.Body);
    var body = await reader.ReadToEndAsync(ct);

    // Envelope: { "name": "northwind", "ppi": 110 } — parsed defensively like the
    // /api/demo/compare/libreoffice envelope.
    string? name = null;
    int? requestedPpi = null;
    try
    {
        using var envelope = JsonDocument.Parse(body);
        if (envelope.RootElement.ValueKind == JsonValueKind.Object)
        {
            if (envelope.RootElement.TryGetProperty("name", out var nameElement)
                && nameElement.ValueKind == JsonValueKind.String)
            {
                name = nameElement.GetString();
            }
            if (envelope.RootElement.TryGetProperty("ppi", out var ppiElement)
                && ppiElement.ValueKind == JsonValueKind.Number
                && ppiElement.TryGetInt32(out var ppiValue))
            {
                requestedPpi = ppiValue;
            }
        }
    }
    catch (JsonException)
    {
        // Falls through to the envelope error below.
    }

    if (string.IsNullOrWhiteSpace(name))
    {
        return Results.BadRequest(new { error = "Request body must be a JSON object with a 'name' property naming a demo deck (see GET /api/demo/decks)." });
    }

    if (demoDeckService is not DemoDeckService concreteDemoDeckService
        || !concreteDemoDeckService.TryGetDeckFile(name, out var deckFile))
    {
        return Results.BadRequest(new { error = $"Unknown or missing demo deck '{name}'." });
    }

    var clampedPpi = DeckPreviewValidators.ClampPpi(requestedPpi);
    var deckBytes = await File.ReadAllBytesAsync(deckFile, ct);

    TypstLegResult result;
    try
    {
        result = demoDeckService.RenderTypstLeg(deckBytes, Path.GetFileName(deckFile), clampedPpi);
    }
    catch (ArgumentException ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
    catch (Exception ex)
    {
        // PNG render failures (Typst backend errors, …) are surfaced, never swallowed.
        return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
    }

    return Results.Ok(ToTypstLegResponse(context, resultStore, result, name));
});

app.MapPost("/api/demo/compare/typst-upload", async (
    HttpContext context,
    IDemoDeckService demoDeckService,
    IConversionResultStore resultStore,
    CancellationToken ct) =>
{
    var form = await context.Request.ReadFormAsync(ct);

    // Validation mirrors POST /api/demo/compare/libreoffice-upload.
    var file = form.Files.GetFile("file");
    if (file is null)
    {
        return Results.BadRequest(new { error = "A 'file' form field with a .pptx upload is required." });
    }

    if (!file.FileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
    {
        return Results.BadRequest(new { error = "Only .pptx files are supported." });
    }

    var requestedPpi = int.TryParse(form["ppi"].FirstOrDefault(), out var ppiValue) ? ppiValue : (int?)null;
    var clampedPpi = DeckPreviewValidators.ClampPpi(requestedPpi);

    byte[] sourceBytes;
    await using (var stream = file.OpenReadStream())
    {
        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream, ct);
        sourceBytes = memoryStream.ToArray();
    }

    TypstLegResult result;
    try
    {
        result = demoDeckService.RenderTypstLeg(sourceBytes, file.FileName, clampedPpi);
    }
    catch (ArgumentException ex)
    {
        // Client-side problems: invalid deck, empty deck, over the slide cap.
        return Results.BadRequest(new { error = ex.Message });
    }
    catch (Exception ex)
    {
        // PNG render failures (Typst backend errors, …) are surfaced, never swallowed.
        return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
    }

    return Results.Ok(ToTypstLegResponse(context, resultStore, result, file.FileName));
});

app.MapGet("/api/demo/deck-template", (IDemoDeckService demoDeckService) =>
{
    try
    {
        return Results.Content(demoDeckService.GetDeckTemplateJson(), "application/json");
    }
    catch (FileNotFoundException ex)
    {
        return Results.NotFound(new { error = ex.Message });
    }
});

app.Run();

static IReadOnlyList<GenerationIssueDto> ToDtos(IReadOnlyList<PptxEditor.Core.Generation.Schema.GenerationIssue> issues) =>
    issues.Select(i => new GenerationIssueDto(i.Path, i.Message, i.Suggestion, i.Severity.ToString())).ToList();

// Maps a LibreOffice compare render to the flat endpoint DTO: stores the PDF for
// download when produced. SlideCount is the real deck slide count (opened by the
// endpoint), NOT the rasterized page count — the client needs it for per-slide math
// even when pdftoppm did not run (previews are then null and the client falls back
// to the PDF).
static LibreOfficeLegResponse ToLibreOfficeLegResponse(
    HttpContext context,
    IConversionResultStore resultStore,
    LibreOfficeRenderResult loResult,
    string deck,
    int slideCount)
{
    string? pdfDownloadUrl = null;
    if (loResult.PdfBytes is not null)
    {
        var id = resultStore.Store(loResult.PdfBytes, "application/pdf", $"{deck}-libreoffice.pdf");
        pdfDownloadUrl = $"{context.Request.Scheme}://{context.Request.Host}/api/download/{id}";
    }

    var previews = loResult.PngPages.Count > 0
        ? loResult.PngPages
            .Select((bytes, index) => new GeneratedSlidePreviewDto(
                index + 1, "png", "image/png", Convert.ToBase64String(bytes)))
            .ToList()
        : null;

    return new LibreOfficeLegResponse(
        Success: loResult.Available && loResult.Error is null,
        Deck: deck,
        SlideCount: slideCount,
        Available: loResult.Available,
        Version: loResult.Version,
        PdfToPpmAvailable: loResult.PdfToPpmAvailable,
        ConversionMilliseconds: loResult.Available ? loResult.ConversionMilliseconds : null,
        RasterizationMilliseconds: loResult.RasterizationMilliseconds,
        TotalMilliseconds: loResult.Available ? loResult.TotalMilliseconds : null,
        Previews: previews,
        PdfDownloadUrl: pdfDownloadUrl,
        Error: loResult.Error);
}

// Maps a Typst compare render to the flat endpoint DTO: stores the PDF for download
// when the best-effort export produced one; PdfError carries the failure otherwise.
// Success reflects the PNG leg (a PDF failure never fails the leg).
static TypstLegResponse ToTypstLegResponse(
    HttpContext context,
    IConversionResultStore resultStore,
    TypstLegResult result,
    string deck)
{
    string? pdfDownloadUrl = null;
    if (result.PdfBytes is not null)
    {
        var id = resultStore.Store(result.PdfBytes, "application/pdf", $"{deck}-typst.pdf");
        pdfDownloadUrl = $"{context.Request.Scheme}://{context.Request.Host}/api/download/{id}";
    }

    return new TypstLegResponse(
        Success: true,
        Deck: deck,
        SlideCount: result.SlideCount,
        PngMilliseconds: result.PngMilliseconds,
        PdfMilliseconds: result.PdfMilliseconds,
        // Wall-clock of the parallel pair (PNG and PDF legs run concurrently) — NOT the
        // sum of the per-phase times.
        TotalMilliseconds: result.TotalMilliseconds,
        Previews: result.PngPages
            .Select((bytes, index) => new GeneratedSlidePreviewDto(
                index + 1, "png", "image/png", Convert.ToBase64String(bytes)))
            .ToList(),
        PdfDownloadUrl: pdfDownloadUrl,
        PdfError: result.PdfError);
}

static bool TryParseTargetFormat(string value, out ConversionTargetFormat targetFormat)
{
    targetFormat = default;
    return Enum.TryParse(value, ignoreCase: true, out targetFormat)
           && Enum.IsDefined(typeof(ConversionTargetFormat), targetFormat);
}

static IReadOnlyDictionary<string, string>? ParseOptions(string? json)
{
    if (string.IsNullOrWhiteSpace(json))
    {
        return null;
    }

    try
    {
        var dictionary = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
        return dictionary as IReadOnlyDictionary<string, string>;
    }
    catch (JsonException)
    {
        return null;
    }
}

// Marker partial class enabling WebApplicationFactory<Program> integration tests
// against the top-level-statements host (consumed by future HTTP-layer tests).
public partial class Program {}
