using PptxEditor.Core.Generation.Model;

namespace PptxEditor.Core.Generation.Components;

/// <summary>
/// <c>table_block</c> (component layer): a header row plus body rows. The table
/// look is delegated to the existing emission path — the component expands to plain
/// row/column containers with cell texts (the same primitives every table renders as),
/// with no OOXML/Typst calls of its own. Rows take a deterministic height (the body-size
/// line estimate plus cell padding, overridable via <c>rowHeight</c>); the header row is
/// filled with the "primary" token, body rows are unfilled.
/// <para>Overflow contract: cell texts shrink; rows that do not fit the box are a
/// structural error (default container policy).</para>
/// </summary>
public static class TableBlockComponent
{
    private const double CellPadV = 6;
    private const double CellPadH = 8;

    private static readonly IReadOnlySet<string> ContentProps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "columns", "rows", "header", "columnWeights", "rowHeight" };

    /// <summary>Expands <paramref name="element"/> into its primitive subtree.</summary>
    internal static ContainerElement Expand(ComponentElement element, DesignTokens design, string path, ElementIdAllocator allocator)
    {
        ArgumentNullException.ThrowIfNull(element);
        ArgumentNullException.ThrowIfNull(design);

        var reader = ComponentContentReader.For(path, "table_block", element.Content, ContentProps, design.Palette);
        var columns = reader.StringArray("columns", required: true, minCount: 1);
        var rows = reader.StringMatrix("rows", required: true);
        var header = reader.Bool("header", defaultValue: true);
        var weights = reader.NumberArray("columnWeights");
        var metrics = design.Metrics;
        var rowHeight = reader.Number("rowHeight", min: 1) ?? DefaultRowHeight(metrics.BodySizePt);
        reader.ThrowIfInvalid();

        if (weights is not null && weights.Count != columns!.Count)
        {
            throw new ComponentException(path,
                $"content.columnWeights: must contain exactly {columns.Count} weight(s), one per column (got {weights.Count}).");
        }
        for (var r = 0; r < rows!.Count; r++)
        {
            if (rows[r].Count != columns!.Count)
            {
                throw new ComponentException(path,
                    $"content.rows[{r}]: must contain exactly {columns.Count} cell(s), one per column (got {rows[r].Count}).");
            }
        }

        ComponentStyle.RequireTokens(design, path, "table_block", "ink", header ? "primary" : null, header ? "paper" : null);

        var children = new List<GenElement>();
        if (header)
        {
            children.Add(Row(columns!, weights, rowHeight, metrics.BodySizePt,
                fill: new SolidFill("primary"), textColor: "paper", bold: true, rowId: allocator.ChildId(element.Id, "row-0"), allocator));
        }
        for (var r = 0; r < rows.Count; r++)
        {
            children.Add(Row(rows[r], weights, rowHeight, metrics.BodySizePt,
                fill: null, textColor: "ink", bold: false, rowId: allocator.ChildId(element.Id, $"row-{r + (header ? 1 : 0)}"), allocator));
        }

        return new ContainerElement
        {
            Id = element.Id,
            Size = element.Size,
            At = element.At,
            Layout = new LayoutSpec { Mode = LayoutMode.Column },
            Children = children
        };
    }

    /// <summary>Default row height: one body-size line plus the vertical cell padding.</summary>
    public static double DefaultRowHeight(double bodySizePt) => Math.Round(ComponentStyle.LineHeight(bodySizePt) + 2 * CellPadV, 3);

    private static ContainerElement Row(
        IReadOnlyList<string> cells, IReadOnlyList<double>? weights, double rowHeight, double bodySizePt,
        FillSpec? fill, string textColor, bool bold, string? rowId, ElementIdAllocator allocator)
    {
        var children = new List<GenElement>(cells.Count);
        for (var c = 0; c < cells.Count; c++)
        {
            var cellId = allocator.ChildId(rowId, $"cell-{c}");
            children.Add(new ContainerElement
            {
                Id = cellId,
                Size = new SizeSpec { Grow = weights?[c] ?? 1 },
                Padding = EdgeInsets.Symmetric(CellPadV, CellPadH),
                Layout = new LayoutSpec { Mode = LayoutMode.Row, Align = AlignItems.Center },
                Children =
                [
                    new TextElement
                    {
                        Id = allocator.ChildId(cellId, "text"),
                        Value = cells[c],
                        Font = "body",
                        FontSize = bodySizePt,
                        Color = textColor,
                        Bold = bold,
                        Anchor = TextAnchor.Middle,
                        Size = new SizeSpec { Grow = 1 }
                    }
                ]
            });
        }

        return new ContainerElement
        {
            Id = rowId,
            Size = new SizeSpec { Height = rowHeight },
            Fill = fill,
            Layout = new LayoutSpec { Mode = LayoutMode.Row },
            Children = children
        };
    }
}
