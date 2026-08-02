namespace DocxEditor.Core.Generation.Model;

/// <summary>
/// Section page geometry: page size, orientation, margins, columns and the section break
/// applied before this section. Every field is optional — nulls resolve to the design
/// token page defaults (or the built-in A4 / portrait / 1-inch defaults).
/// </summary>
public sealed record PageSetup
{
    /// <summary>Named or custom page size. Null = design or built-in default (A4).</summary>
    public PageSize? PageSize { get; init; }

    /// <summary>Page orientation. Null = design or built-in default (portrait).</summary>
    public PageOrientation? Orientation { get; init; }

    /// <summary>Page margins in points. Null = design or built-in default (1 inch).</summary>
    public Margins? Margins { get; init; }

    /// <summary>Optional column layout for the section body.</summary>
    public PageColumns? Columns { get; init; }

    /// <summary>Section break applied before this section; ignored for the first section.</summary>
    public SectionBreakType? BreakType { get; init; }
}

/// <summary>
/// A page size: either a named ISO/ANSI size (<see cref="Name"/>) or a custom size
/// (<see cref="WidthPt"/> / <see cref="HeightPt"/>). Named and custom are mutually
/// exclusive. Dimensions are always the portrait (upright) dimensions; landscape is
/// resolved by swapping via <see cref="EffectiveSize"/>.
/// </summary>
public sealed record PageSize
{
    /// <summary>Named size; mutually exclusive with <see cref="WidthPt"/> / <see cref="HeightPt"/>.</summary>
    public PageSizeName? Name { get; init; }

    /// <summary>Custom width in points (portrait dimension). Only valid without <see cref="Name"/>.</summary>
    public double? WidthPt { get; init; }

    /// <summary>Custom height in points (portrait dimension). Only valid without <see cref="Name"/>.</summary>
    public double? HeightPt { get; init; }

    /// <summary>True when a named size is used.</summary>
    public bool IsNamed => Name is not null;

    /// <summary>True when custom dimensions are used.</summary>
    public bool IsCustom => Name is null;

    /// <summary>Portrait dimensions in points for each <see cref="PageSizeName"/>.</summary>
    public static IReadOnlyDictionary<PageSizeName, PageSize> Catalog { get; } =
        new Dictionary<PageSizeName, PageSize>
        {
            [PageSizeName.A3] = new() { WidthPt = 841.9, HeightPt = 1190.6 },
            [PageSizeName.A4] = new() { WidthPt = 595.3, HeightPt = 841.9 },
            [PageSizeName.A5] = new() { WidthPt = 419.5, HeightPt = 595.3 },
            [PageSizeName.B4] = new() { WidthPt = 708.7, HeightPt = 1000.6 },
            [PageSizeName.B5] = new() { WidthPt = 498.9, HeightPt = 708.7 },
            [PageSizeName.Letter] = new() { WidthPt = 612.0, HeightPt = 792.0 },
            [PageSizeName.Legal] = new() { WidthPt = 612.0, HeightPt = 1008.0 },
            [PageSizeName.Executive] = new() { WidthPt = 522.0, HeightPt = 756.0 },
            [PageSizeName.Statement] = new() { WidthPt = 396.0, HeightPt = 612.0 },
            [PageSizeName.Tabloid] = new() { WidthPt = 792.0, HeightPt = 1224.0 }
        };

    /// <summary>Built-in fallback page size (A4 portrait).</summary>
    public static PageSize Default => Named(PageSizeName.A4);

    /// <summary>Returns the catalog entry for a named size.</summary>
    public static PageSize Named(PageSizeName name) => Catalog[name];

    /// <summary>
    /// Resolves the effective page dimensions in points for the given orientation,
    /// swapping width and height for <see cref="PageOrientation.Landscape"/>.
    /// </summary>
    public (double WidthPt, double HeightPt) EffectiveSize(PageOrientation orientation)
    {
        double width;
        double height;
        if (IsNamed)
        {
            var named = Catalog[Name!.Value];
            width = named.WidthPt!.Value;
            height = named.HeightPt!.Value;
        }
        else
        {
            width = WidthPt!.Value;
            height = HeightPt!.Value;
        }
        return orientation == PageOrientation.Landscape ? (height, width) : (width, height);
    }
}

/// <summary>Page margins in points. All four edges default to 72pt (1 inch).</summary>
public sealed record Margins
{
    /// <summary>Top margin in points (≥ 0).</summary>
    public double TopPt { get; init; }

    /// <summary>Right margin in points (≥ 0).</summary>
    public double RightPt { get; init; }

    /// <summary>Bottom margin in points (≥ 0).</summary>
    public double BottomPt { get; init; }

    /// <summary>Left margin in points (≥ 0).</summary>
    public double LeftPt { get; init; }

    /// <summary>Built-in default margins: 72pt (1 inch) on every edge.</summary>
    public static Margins Defaults => new() { TopPt = 72, RightPt = 72, BottomPt = 72, LeftPt = 72 };
}

/// <summary>Equal-width column layout for a section body (count 1 = no columns).</summary>
public sealed record PageColumns
{
    /// <summary>Column count (≥ 1). 1 means single column.</summary>
    public int Count { get; init; } = 1;

    /// <summary>Gutter between columns in points (≥ 0).</summary>
    public double SpacingPt { get; init; }

    /// <summary>Whether a vertical separator line is drawn between columns.</summary>
    public bool SeparatorLine { get; init; }
}
