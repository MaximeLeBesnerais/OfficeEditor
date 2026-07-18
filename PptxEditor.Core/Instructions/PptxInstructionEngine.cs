using DocumentFormat.OpenXml.Packaging;
using PptxEditor.Core.Builders;
using PptxEditor.Core.Models;
using PptxEditor.Core.Services;

namespace PptxEditor.Core.Instructions;

/// <summary>Per-op failure detail: operation index (0-based), op type, message.</summary>
public record PptxOpError(int Index, string Type, string Error);

/// <summary>
/// Outcome of an instruction batch.
/// <list type="bullet">
/// <item><see cref="AppliedOps"/> — number of operations successfully applied.</item>
/// <item><see cref="FailedOps"/> — per-op errors (validation or execution).</item>
/// <item><see cref="ChangedSlides"/> — instruction-derived 1-based slide indexes, sorted.
/// moveSlide marks the whole affected range [min(from,to)..max(from,to)]; deleteSlide marks
/// the range that shifted into the deleted slot; duplicateSlide marks the inserted slot.</item>
/// <item><see cref="Revision"/> — a fresh Guid per applied batch; null when nothing was applied.</item>
/// </list>
/// </summary>
public record PptxEditResult
{
    public int AppliedOps { get; init; }
    public IReadOnlyList<PptxOpError> FailedOps { get; init; } = Array.Empty<PptxOpError>();
    public IReadOnlyList<int> ChangedSlides { get; init; } = Array.Empty<int>();
    public Guid? Revision { get; init; }

    public bool Success => FailedOps.Count == 0;
}

/// <summary>
/// Validate-then-execute executor for the PPTX instruction vocabulary.
/// <para>
/// Validation runs over ALL operations first (slide bounds, element existence and type via
/// <see cref="PptxAnatomizer"/>, image base64 decodability, last-slide delete protection).
/// Any validation failure → nothing is applied and all errors are returned.
/// </para>
/// <para>
/// Execution applies ops in order against the builder. Ops are applied sequentially, so a
/// structural op (move/duplicate/delete) shifts slide indexes for SUBSEQUENT ops in the same
/// batch; validation is done against the pre-batch deck, and an op whose target moved out
/// from under it fails with an explicit per-op error instead of editing the wrong slide.
/// Exceptions are never swallowed: any failure becomes a <see cref="PptxOpError"/> and the
/// remaining ops continue.
/// </para>
/// </summary>
public class PptxInstructionEngine
{
    public PptxEditResult Apply(IPresentationBuilder builder, PptxInstructionSet instructions)
    {
        ArgumentNullException.ThrowIfNull(builder);
        ArgumentNullException.ThrowIfNull(instructions);

        var operations = instructions.Operations;
        var validationErrors = Validate(builder, operations);
        if (validationErrors.Count > 0)
        {
            return new PptxEditResult
            {
                AppliedOps = 0,
                FailedOps = validationErrors,
                ChangedSlides = Array.Empty<int>(),
                Revision = null
            };
        }

        var appliedOps = 0;
        var failedOps = new List<PptxOpError>();
        var changedSlides = new SortedSet<int>();

        for (var index = 0; index < operations.Count; index++)
        {
            var op = operations[index];
            try
            {
                foreach (var slide in Execute(builder, op))
                {
                    changedSlides.Add(slide);
                }
                appliedOps++;
            }
            catch (Exception ex)
            {
                failedOps.Add(new PptxOpError(index, op.Type, ex.Message));
            }
        }

        return new PptxEditResult
        {
            AppliedOps = appliedOps,
            FailedOps = failedOps,
            ChangedSlides = changedSlides.ToList(),
            Revision = appliedOps > 0 ? Guid.NewGuid() : null
        };
    }

    private static List<PptxOpError> Validate(IPresentationBuilder builder, List<PptxInstruction> operations)
    {
        var errors = new List<PptxOpError>();
        var slideCount = builder.SlideCount;

        // Element id → type lookup per 1-based slide, from the anatomizer.
        var elementTypesBySlide = new Dictionary<int, Dictionary<uint, string>>();
        foreach (var slideAnatomy in builder.Analyze())
        {
            elementTypesBySlide[slideAnatomy.SlideIndex] = slideAnatomy.Elements
                .GroupBy(e => e.Id)
                .ToDictionary(g => g.Key, g => g.First().Type);
        }

        var deleteCount = 0;

        for (var index = 0; index < operations.Count; index++)
        {
            var op = operations[index];
            switch (op)
            {
                case PptxReplaceTextInstruction replaceText:
                    ValidateElementOp(errors, index, op.Type, elementTypesBySlide, slideCount,
                        replaceText.Slide, replaceText.ElementId, expectedType: "Text");
                    break;

                case PptxReplaceImageInstruction replaceImage:
                    ValidateElementOp(errors, index, op.Type, elementTypesBySlide, slideCount,
                        replaceImage.Slide, replaceImage.ElementId, expectedType: "Image");
                    try
                    {
                        Convert.FromBase64String(replaceImage.Image);
                    }
                    catch (FormatException)
                    {
                        errors.Add(new PptxOpError(index, op.Type,
                            $"'image' is not valid base64 (operations[{index}])."));
                    }
                    break;

                case PptxReplaceTableInstruction replaceTable:
                    ValidateElementOp(errors, index, op.Type, elementTypesBySlide, slideCount,
                        replaceTable.Slide, replaceTable.ElementId, expectedType: "Table");
                    break;

                case PptxMoveSlideInstruction moveSlide:
                    ValidateSlideInBounds(errors, index, op.Type, moveSlide.From, slideCount, "from");
                    ValidateSlideInBounds(errors, index, op.Type, moveSlide.To, slideCount, "to");
                    break;

                case PptxDuplicateSlideInstruction duplicateSlide:
                    ValidateSlideInBounds(errors, index, op.Type, duplicateSlide.Slide, slideCount, "slide");
                    if (duplicateSlide.Position is { } position && position > slideCount + 1)
                    {
                        errors.Add(new PptxOpError(index, op.Type,
                            $"position {position} is out of range: deck has {slideCount} slide(s), " +
                            $"so position must be between 1 and {slideCount + 1}."));
                    }
                    break;

                case PptxDeleteSlideInstruction deleteSlide:
                    ValidateSlideInBounds(errors, index, op.Type, deleteSlide.Slide, slideCount, "slide");
                    deleteCount++;
                    if (deleteCount >= slideCount)
                    {
                        errors.Add(new PptxOpError(index, op.Type,
                            "Refusing to delete the last remaining slide of the deck."));
                    }
                    break;
            }
        }

        return errors;
    }

    private static void ValidateElementOp(
        List<PptxOpError> errors,
        int index,
        string opType,
        Dictionary<int, Dictionary<uint, string>> elementTypesBySlide,
        int slideCount,
        int slide,
        uint elementId,
        string expectedType)
    {
        if (!ValidateSlideInBounds(errors, index, opType, slide, slideCount, "slide"))
        {
            return;
        }

        if (!elementTypesBySlide.TryGetValue(slide, out var elements)
            || !elements.TryGetValue(elementId, out var actualType))
        {
            errors.Add(new PptxOpError(index, opType,
                $"No element with id {elementId} on slide {slide}."));
            return;
        }

        if (!string.Equals(actualType, expectedType, StringComparison.Ordinal))
        {
            errors.Add(new PptxOpError(index, opType,
                $"Element {elementId} on slide {slide} is '{actualType}', " +
                $"but {opType} requires a '{expectedType}' element."));
        }
    }

    private static bool ValidateSlideInBounds(
        List<PptxOpError> errors, int index, string opType, int slide, int slideCount, string field)
    {
        if (slide < 1 || slide > slideCount)
        {
            errors.Add(new PptxOpError(index, opType,
                $"{field} {slide} is out of range: deck has {slideCount} slide(s)."));
            return false;
        }
        return true;
    }

    /// <summary>Executes one op; returns the 1-based changed slides it produced.</summary>
    private static IEnumerable<int> Execute(IPresentationBuilder builder, PptxInstruction op)
    {
        switch (op)
        {
            case PptxReplaceTextInstruction replaceText:
            {
                var slidePart = GetSlidePart(builder, replaceText.Slide);
                var result = new PptxElementReplacer().ReplaceText(slidePart, replaceText.ElementId, replaceText.Text);
                if (!result.Success)
                {
                    throw new InvalidOperationException(result.Error);
                }
                return new[] { replaceText.Slide };
            }

            case PptxReplaceImageInstruction replaceImage:
            {
                var slidePart = GetSlidePart(builder, replaceImage.Slide);
                var imageBytes = Convert.FromBase64String(replaceImage.Image);
                var fitMode = ParseFitMode(replaceImage.Fit);
                using var stream = new MemoryStream(imageBytes, writable: false);
                var result = new PptxElementReplacer().ReplaceImage(
                    slidePart, replaceImage.ElementId, stream, DetectImageExtension(imageBytes), fitMode);
                if (!result.Success)
                {
                    throw new InvalidOperationException(result.Error);
                }
                return new[] { replaceImage.Slide };
            }

            case PptxReplaceTableInstruction replaceTable:
            {
                var slidePart = GetSlidePart(builder, replaceTable.Slide);
                var result = new PptxElementReplacer().ReplaceTableData(slidePart, replaceTable.ElementId, replaceTable.Rows);
                if (!result.Success)
                {
                    throw new InvalidOperationException(result.Error);
                }
                return new[] { replaceTable.Slide };
            }

            case PptxMoveSlideInstruction moveSlide:
            {
                builder.ReorderSlide(moveSlide.From - 1, moveSlide.To - 1);
                // Every slide in [min..max] changes slot.
                var min = Math.Min(moveSlide.From, moveSlide.To);
                var max = Math.Max(moveSlide.From, moveSlide.To);
                return min == max ? Array.Empty<int>() : Enumerable.Range(min, max - min + 1).ToArray();
            }

            case PptxDuplicateSlideInstruction duplicateSlide:
            {
                var insertPosition1Based = duplicateSlide.Position ?? duplicateSlide.Slide + 1;
                builder.DuplicateSlide(duplicateSlide.Slide - 1, insertPosition1Based - 1);
                return new[] { insertPosition1Based };
            }

            case PptxDeleteSlideInstruction deleteSlide:
            {
                var preCount = builder.SlideCount;
                builder.RemoveSlide(deleteSlide.Slide - 1);
                var postCount = preCount - 1;
                // Slides that shifted into the deleted slot and everything after it
                // changed identity; deleting the last slide shifts nothing.
                return deleteSlide.Slide <= postCount
                    ? Enumerable.Range(deleteSlide.Slide, postCount - deleteSlide.Slide + 1).ToArray()
                    : Array.Empty<int>();
            }

            default:
                throw new NotSupportedException($"Instruction type '{op.Type}' is not supported.");
        }
    }

    private static SlidePart GetSlidePart(IPresentationBuilder builder, int slide1Based)
    {
        // The IPresentationBuilder.Replace* wrappers swallow the replacer's
        // PptxReplaceResult, so the engine drives the replacer directly to honor
        // the never-swallow rule. GetSlide(i) returns the concrete SlideBuilder.
        if (builder.GetSlide(slide1Based - 1) is not SlideBuilder slideBuilder)
        {
            throw new InvalidOperationException(
                "The underlying slide builder does not expose its SlidePart; cannot apply element edits.");
        }
        return slideBuilder.SlidePart;
    }

    /// <summary>
    /// Maps the validated fit argument to <see cref="ImageFitMode"/>. The JSON parser
    /// validates the vocabulary already; the throw below only guards instructions
    /// constructed directly in code. The vocabulary carries no explicit crop rect,
    /// so "crop" reaches the replacer with a null rect and behaves as "fill".
    /// </summary>
    private static ImageFitMode ParseFitMode(string? fit) => fit?.ToLowerInvariant() switch
    {
        null or "" or "stretch" => ImageFitMode.Stretch,
        "fill" => ImageFitMode.Fill,
        "crop" => ImageFitMode.Crop,
        "contain" => ImageFitMode.Contain,
        _ => throw new InvalidOperationException(
            $"Unknown fit mode '{fit}'; expected one of stretch|fill|crop|contain.")
    };

    private static string DetectImageExtension(byte[] bytes)
    {
        if (bytes is [0x89, 0x50, 0x4E, 0x47, ..]) return ".png";
        if (bytes is [0xFF, 0xD8, ..]) return ".jpg";
        if (bytes is [0x47, 0x49, 0x46, ..]) return ".gif";
        if (bytes is [0x42, 0x4D, ..]) return ".bmp";
        if (bytes is [0x49, 0x49, 0x2A, 0x00, ..] or [0x4D, 0x4D, 0x00, 0x2A, ..]) return ".tiff";
        return ".png"; // fallback: let the replacer's part-type default handle it
    }
}
