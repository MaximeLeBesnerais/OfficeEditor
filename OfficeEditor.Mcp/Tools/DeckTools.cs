using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using OfficeEditor.Core.Services;
using OfficeEditor.Mcp.JsonRpc;
using OfficeEditor.Mcp.Sessions;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Emit.Typst;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;
using PptxEditor.Core.Instructions;
using PptxEditor.Core.Serialization;
using PptxEditor.Core.Services;

namespace OfficeEditor.Mcp.Tools;

/// <summary>
/// The three deck tools: thin wrappers over the stabilized Core services
/// (PptxAnatomizer via IPresentationBuilder.Analyze, PptxInstructionEngine,
/// PresentationBuilder.ExportThumbnail). All argument validation happens here and
/// raises <see cref="McpException"/> with client-safe messages; the host maps them
/// onto JSON-RPC error objects.
/// </summary>
public sealed class DeckTools
{
    private static readonly JsonSerializerOptions PayloadJson = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly DeckSessionStore _sessions;

    public DeckTools(DeckSessionStore sessions)
    {
        _sessions = sessions;
    }

    public JsonArray ListTools() => ToolSchemas.ListAll();

    public JsonObject Call(string name, JsonElement? arguments)
    {
        if (arguments is null || arguments.Value.ValueKind != JsonValueKind.Object)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"Tool '{name}' requires an 'arguments' object.");
        }
        var args = arguments.Value;

        return name switch
        {
            ToolSchemas.AnatomizeName => Anatomize(args),
            ToolSchemas.ReplaceElementName => ReplaceElement(args),
            ToolSchemas.RenderSlideName => RenderSlide(args),
            ToolSchemas.GenerateName => Generate(args),
            _ => throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"Unknown tool '{name}'. Available tools: {ToolSchemas.AnatomizeName}, " +
                $"{ToolSchemas.ReplaceElementName}, {ToolSchemas.RenderSlideName}, {ToolSchemas.GenerateName}.")
        };
    }

    /// <summary>
    /// deck_generate: generation JSON in → PPTX + per-slide previews out.
    /// The pipeline is the P1–P5 chain (validate → expand → layout → OOXML; previews via the
    /// Typst emitter → TypstBridge-first compiler). Preview rendering is best-effort: with no
    /// Typst backend available the deck still succeeds and previewError carries the reason.
    /// </summary>
    private JsonObject Generate(JsonElement args)
    {
        var documentElement = GetRequired(args, "document");
        if (documentElement.ValueKind != JsonValueKind.Object)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                "'document' must be a JSON object (the generation document).");
        }
        // Raw text keeps the P1 validator's error paths byte-accurate.
        var documentJson = documentElement.GetRawText();

        var format = GetOptionalString(args, "previewFormat")?.Trim().ToLowerInvariant() ?? "svg";
        if (format is not ("svg" or "png"))
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'previewFormat' must be \"svg\" or \"png\" (got '{format}').");
        }

        var ppi = GetOptionalNumber(args, "ppi") ?? 150f;
        if (ppi is < 36f or > 600f)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'ppi' must be between 36 and 600 (got {ppi}).");
        }

        var totalTimer = Stopwatch.StartNew();
        var generationTimer = Stopwatch.StartNew();

        var validation = new GenerationDocumentParser().Validate(documentJson);
        if (!validation.IsValid)
        {
            // P1's actionable errors, verbatim — one per line so an AI caller can fix them all.
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"Invalid generation document ({validation.Errors.Count} error(s)):{Environment.NewLine}" +
                string.Join(Environment.NewLine, validation.Errors.Select(e => e.ToString())));
        }

        LayoutResult layout;
        try
        {
            // Fresh stack per call: resolver/measurer/catalog carry per-run state and are
            // not thread-safe; the system-font scan behind the catalog is process-cached,
            // so this stays cheap on the warm path.
            var expanded = ComponentExpander.Expand(validation.Document!);
            layout = new LayoutResolver(new TextMeasure(new FontMetricsCatalog())).Resolve(expanded);
        }
        catch (ComponentException ex)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, $"Invalid generation document: {ex.Message}");
        }
        catch (LayoutException ex)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, $"Invalid generation document: {ex.Message}");
        }

        var emission = new OoxmlEmitter().Emit(layout);
        var generationMs = generationTimer.Elapsed.TotalMilliseconds;

        var session = _sessions.Store(emission.Bytes, "generated.pptx", layout.Slides.Count);

        JsonArray previews = [];
        string? previewError = null;
        try
        {
            var source = new TypstEmitter().Emit(layout);
            using var compiler = new TypstCompilerService();
            var result = compiler.Compile(source, new CompileOptions
            {
                Format = format == "svg" ? OutputFormat.Svg : OutputFormat.Png,
                Ppi = ppi
            });
            if (!result.Success)
            {
                previewError = result.ErrorMessage ?? "Typst preview compilation failed.";
            }
            else if (result.Pages.Length != layout.Slides.Count)
            {
                previewError = $"Typst preview produced {result.Pages.Length} page(s) for {layout.Slides.Count} slide(s).";
            }
            else
            {
                var contentType = format == "svg" ? "image/svg+xml" : "image/png";
                for (var i = 0; i < result.Pages.Length; i++)
                {
                    previews.Add(new JsonObject
                    {
                        ["slide"] = i + 1,
                        ["format"] = format,
                        ["contentType"] = contentType,
                        ["contentBase64"] = Convert.ToBase64String(result.Pages[i])
                    });
                }
            }
        }
        catch (TypstEmitException ex)
        {
            previewError = $"Typst emission failed: {ex.Message}";
        }

        totalTimer.Stop();
        var payload = new JsonObject
        {
            ["success"] = true,
            ["slideCount"] = layout.Slides.Count,
            ["pptxBase64"] = Convert.ToBase64String(emission.Bytes),
            ["previewFormat"] = format,
            ["ppi"] = ppi,
            ["previews"] = previews,
            ["warnings"] = new JsonArray(validation.Warnings.Select(w => (JsonNode)w.ToString()).ToArray()),
            ["pipelineWarnings"] = new JsonArray(layout.Warnings.Concat(emission.Warnings).Select(w => (JsonNode)w).ToArray()),
            ["generationMs"] = Math.Round(generationMs, 2),
            ["totalMs"] = Math.Round(totalTimer.Elapsed.TotalMilliseconds, 2)
        };
        if (previewError is not null)
        {
            payload["previewError"] = previewError;
        }
        payload["deckHandle"] = session.Handle.ToString();
        payload["revision"] = session.Revision;
        return payload;
    }

    private JsonObject Anatomize(JsonElement args)
    {
        var deck = ResolveDeck(args);
        List<PptxEditor.Core.Models.SlideAnatomy> anatomy;
        using (var builder = PresentationBuilder.Open(deck.Bytes))
        {
            anatomy = builder.Analyze();
        }

        var payload = new JsonObject
        {
            ["slideCount"] = deck.SlideCount,
            ["slides"] = JsonNode.Parse(JsonSerializer.Serialize(anatomy, PayloadJson))
        };
        AppendSessionInfo(payload, deck);
        return payload;
    }

    private JsonObject ReplaceElement(JsonElement args)
    {
        var operations = GetRequired(args, "operations");
        if (operations.ValueKind != JsonValueKind.Array)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, "'operations' must be an array.");
        }

        // Reuse W7's parser verbatim: the operations array is round-tripped into the
        // instruction-set envelope it expects, so validation errors name the op index.
        PptxEditor.Core.Models.PptxInstructionSet instructionSet;
        try
        {
            instructionSet = new PptxJsonInstructionParser()
                .Parse($"{{\"operations\":{operations.GetRawText()}}}");
        }
        catch (ArgumentException ex)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, ex.Message);
        }

        var deck = ResolveDeck(args);

        PptxEditResult edit;
        byte[]? newBytes;
        using (var builder = PresentationBuilder.Open(deck.Bytes))
        {
            edit = new PptxInstructionEngine().Apply(builder, instructionSet);
            newBytes = edit.AppliedOps > 0 ? builder.SaveToBytes() : null;
        }

        if (deck.Session is { } session && newBytes is not null)
        {
            // Slide ops can change the slide count — re-derive it from the new bytes
            // (mirrors the API's DeckSessionStore.UpdateBytes) before storing.
            int slideCount;
            using (var probe = PresentationBuilder.Open(newBytes))
            {
                slideCount = probe.SlideCount;
            }
            _sessions.UpdateBytes(session.Handle, newBytes, slideCount);
        }

        var payload = new JsonObject
        {
            ["success"] = edit.Success,
            ["appliedOps"] = edit.AppliedOps,
            ["changedSlides"] = new JsonArray(edit.ChangedSlides.Select(s => (JsonNode)s).ToArray()),
            ["errors"] = new JsonArray(edit.FailedOps.Select(e => (JsonNode)new JsonObject
            {
                ["index"] = e.Index,
                ["type"] = e.Type,
                ["error"] = e.Error
            }).ToArray())
        };
        // Stateless callers (pptxBase64) get the edited bytes back; the upload still
        // created a session (bytes updated above), so the returned handle stays usable.
        if (deck.Uploaded && newBytes is not null)
        {
            payload["newPptxBase64"] = Convert.ToBase64String(newBytes);
        }
        // Session decks get deckHandle + revision appended here.
        AppendSessionInfo(payload, deck);
        return payload;
    }

    private JsonObject RenderSlide(JsonElement args)
    {
        var deck = ResolveDeck(args);

        var slide = GetRequiredPositiveInt(args, "slide");
        if (slide > deck.SlideCount)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'slide' {slide} is out of range: deck has {deck.SlideCount} slide(s).");
        }

        var format = GetOptionalString(args, "format")?.Trim().ToLowerInvariant() ?? "png";
        if (format is not ("png" or "svg"))
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'format' must be \"png\" or \"svg\" (got '{format}').");
        }

        var ppi = GetOptionalNumber(args, "ppi") ?? 150f;
        if (ppi is < 36f or > 600f)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'ppi' must be between 36 and 600 (got {ppi}).");
        }

        byte[] content;
        try
        {
            using var builder = PresentationBuilder.Open(deck.Bytes);
            content = builder.ExportThumbnail(slide - 1, new ThumbnailOptions { Ppi = ppi, Format = format });
        }
        catch (ArgumentException ex)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, ex.Message);
        }

        var payload = new JsonObject
        {
            ["slide"] = slide,
            ["format"] = format,
            ["ppi"] = ppi,
            ["contentType"] = format == "svg" ? "image/svg+xml" : "image/png",
            ["contentBase64"] = Convert.ToBase64String(content)
        };
        AppendSessionInfo(payload, deck);
        return payload;
    }

    /// <summary>
    /// Resolves the deck source: exactly one of deckHandle (session lookup, slides the
    /// lifetime) or pptxBase64 (upload — creates a session so follow-up calls can use
    /// the returned handle, mirroring the API's upload → deckId semantics).
    /// </summary>
    private DeckResolution ResolveDeck(JsonElement args)
    {
        var handleText = GetOptionalString(args, "deckHandle");
        var uploadText = GetOptionalString(args, "pptxBase64");

        if ((handleText is null) == (uploadText is null))
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                "Exactly one of 'deckHandle' or 'pptxBase64' is required.");
        }

        if (handleText is not null)
        {
            if (!Guid.TryParse(handleText, out var handle))
            {
                throw new McpException(JsonRpcErrorCodes.InvalidParams,
                    $"'deckHandle' must be a GUID (got '{handleText}').");
            }
            if (!_sessions.TryGet(handle, out var session) || session is null)
            {
                throw new McpException(JsonRpcErrorCodes.DeckNotFound,
                    $"Unknown or expired deckHandle '{handle}'. Re-upload the deck with pptxBase64.");
            }
            return new DeckResolution(session.DeckBytes, session.SlideCount, session, Uploaded: false);
        }

        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(uploadText!);
        }
        catch (FormatException)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, "'pptxBase64' is not valid base64.");
        }

        // Validate the upload is a readable PPTX (and learn the slide count) before
        // admitting it to a session.
        int count;
        try
        {
            using var builder = PresentationBuilder.Open(bytes);
            count = builder.SlideCount;
        }
        catch (Exception ex) when (ex is InvalidOperationException or InvalidDataException or ArgumentException or IOException)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'pptxBase64' does not contain a readable PPTX deck: {ex.Message}");
        }

        var uploaded = _sessions.Store(bytes, "upload.pptx", count);
        return new DeckResolution(bytes, count, uploaded, Uploaded: true);
    }

    private static void AppendSessionInfo(JsonObject payload, DeckResolution deck)
    {
        if (deck.Session is { } session)
        {
            payload["deckHandle"] = session.Handle.ToString();
            payload["revision"] = session.Revision;
        }
    }

    private static JsonElement GetRequired(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, $"'{name}' is required.");
        }
        return value;
    }

    private static string? GetOptionalString(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return null;
        }
        if (value.ValueKind != JsonValueKind.String)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, $"'{name}' must be a string.");
        }
        return value.GetString();
    }

    private static int GetRequiredPositiveInt(JsonElement args, string name)
    {
        var value = GetRequired(args, name);
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var number) || number < 1)
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams,
                $"'{name}' is required and must be a positive integer (1-based).");
        }
        return number;
    }

    private static float? GetOptionalNumber(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return null;
        }
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetSingle(out var number))
        {
            throw new McpException(JsonRpcErrorCodes.InvalidParams, $"'{name}' must be a number.");
        }
        return number;
    }

    private sealed record DeckResolution(byte[] Bytes, int SlideCount, McpDeckSession? Session, bool Uploaded);
}
