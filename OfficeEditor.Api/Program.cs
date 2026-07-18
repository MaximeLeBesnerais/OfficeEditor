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

app.Run();

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
