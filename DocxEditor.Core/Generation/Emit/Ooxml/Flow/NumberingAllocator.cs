using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Core.Generation.Emit.Ooxml.Flow;

/// <summary>
/// Allocates fresh, collision-free numbering definitions and instances for generated lists.
/// Existing template numbering is never touched or reused: one new abstract definition per
/// (ordered, start) semantic is created on first use and shared by later lists of the same
/// semantic, and every list gets its own numbering instance so separate lists restart their
/// counters. Abstract definitions are inserted before the first instance because CT_Numbering
/// requires all abstractNum elements to precede all num elements.
/// </summary>
internal sealed class NumberingAllocator
{
    private readonly Dictionary<(bool Ordered, int Start), int> _abstractNumIds = new();

    public int Allocate(MainDocumentPart mainPart, bool ordered, int start)
    {
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart is null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        }

        var numbering = numberingPart.Numbering;
        if (numbering is null)
        {
            numbering = new Numbering();
            numberingPart.Numbering = numbering;
        }

        var key = (ordered, start);
        if (!_abstractNumIds.TryGetValue(key, out var abstractNumId))
        {
            abstractNumId = numbering.Elements<AbstractNum>()
                .Select(n => n.AbstractNumberId?.Value ?? -1)
                .DefaultIfEmpty(-1)
                .Max() + 1;

            var abstractNum = ordered
                ? CreateOrderedAbstractNum(abstractNumId, start)
                : CreateBulletAbstractNum(abstractNumId);

            var firstInstance = numbering.Elements<NumberingInstance>().FirstOrDefault();
            if (firstInstance is not null)
            {
                numbering.InsertBefore(abstractNum, firstInstance);
            }
            else
            {
                numbering.Append(abstractNum);
            }

            _abstractNumIds[key] = abstractNumId;
        }

        var numberId = numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        numbering.Append(new NumberingInstance(
            new AbstractNumId { Val = abstractNumId })
        {
            NumberID = numberId
        });

        return numberId;
    }

    private static AbstractNum CreateOrderedAbstractNum(int abstractNumId, int start)
    {
        return new AbstractNum(
            new Level(
                new StartNumberingValue { Val = start },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" })
            ) { LevelIndex = 0 })
        {
            AbstractNumberId = abstractNumId
        };
    }

    private static AbstractNum CreateBulletAbstractNum(int abstractNumId)
    {
        return new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "\u2022" },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(
                    new Indentation { Left = "720", Hanging = "360" })
            ) { LevelIndex = 0 })
        {
            AbstractNumberId = abstractNumId
        };
    }
}
