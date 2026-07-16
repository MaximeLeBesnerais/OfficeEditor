using System.Text.Json;
using OfficeEditor.Api.Models;
using OfficeEditor.Api.Services;

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
