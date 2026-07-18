namespace PptxEditor.Core.Models;

/// <summary>
/// PPTX edit instruction vocabulary, mirroring the DOCX instruction pattern
/// (DocxEditor.Core/Models/Instructions.cs). Slides are addressed 1-based
/// (matching PptxAnatomizer's SlideIndex); element ids are the cNvPr/@id values
/// reported by the anatomizer and are unique per slide only.
/// Names are Pptx-prefixed to avoid clashing with the DOCX instruction records
/// when both namespaces are imported (the API project references both cores).
/// </summary>
public abstract record PptxInstruction
{
    public string Type { get; init; } = string.Empty;
}

/// <summary>replaceText { slide, elementId, text }</summary>
public record PptxReplaceTextInstruction : PptxInstruction
{
    public PptxReplaceTextInstruction()
    {
        Type = "replaceText";
    }

    /// <summary>1-based slide index.</summary>
    public required int Slide { get; init; }
    public required uint ElementId { get; init; }
    public required string Text { get; init; }
}

/// <summary>
/// replaceImage { slide, elementId, image (base64), fit? }
/// <para>
/// <see cref="Fit"/> is parsed and validated ("stretch"|"fill"|"crop"|"contain")
/// but NOT applied by the engine yet — the current replacer always stretches the
/// image into the existing frame. Fit modes land with the F7 work; the field is
/// part of the vocabulary now so clients can send it forward-compatibly.
/// </para>
/// </summary>
public record PptxReplaceImageInstruction : PptxInstruction
{
    public PptxReplaceImageInstruction()
    {
        Type = "replaceImage";
    }

    public required int Slide { get; init; }
    public required uint ElementId { get; init; }

    /// <summary>Base64-encoded image bytes.</summary>
    public required string Image { get; init; }

    /// <summary>Optional fit mode; reserved for F7 (see type remarks).</summary>
    public string? Fit { get; init; }
}

/// <summary>replaceTable { slide, elementId, rows }</summary>
public record PptxReplaceTableInstruction : PptxInstruction
{
    public PptxReplaceTableInstruction()
    {
        Type = "replaceTable";
    }

    public required int Slide { get; init; }
    public required uint ElementId { get; init; }
    public required List<List<string>> Rows { get; init; }
}

/// <summary>moveSlide { from, to } — both 1-based.</summary>
public record PptxMoveSlideInstruction : PptxInstruction
{
    public PptxMoveSlideInstruction()
    {
        Type = "moveSlide";
    }

    public required int From { get; init; }
    public required int To { get; init; }
}

/// <summary>
/// duplicateSlide { slide, position? } — 1-based. Position is the 1-based slot
/// the copy occupies after the operation; defaults to right after the source.
/// </summary>
public record PptxDuplicateSlideInstruction : PptxInstruction
{
    public PptxDuplicateSlideInstruction()
    {
        Type = "duplicateSlide";
    }

    public required int Slide { get; init; }
    public int? Position { get; init; }
}

/// <summary>deleteSlide { slide } — 1-based.</summary>
public record PptxDeleteSlideInstruction : PptxInstruction
{
    public PptxDeleteSlideInstruction()
    {
        Type = "deleteSlide";
    }

    public required int Slide { get; init; }
}

/// <summary>Root document: { "operations": [...] }.</summary>
public record PptxInstructionSet
{
    public required List<PptxInstruction> Operations { get; init; }
}
