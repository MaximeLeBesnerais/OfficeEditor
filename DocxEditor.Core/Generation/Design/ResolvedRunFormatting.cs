namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Fully-resolved, immutable run formatting. Emitters translate this straight to OOXML (or Typst)
/// without re-interpreting palette/font/token strings. Nullable emphasis booleans preserve
/// "not specified" vs. explicit on/off: <c>null</c> inherits from the surrounding style or token,
/// <c>true</c> forces it on. <see cref="Overlay"/> merges a base (typography token) with a more
/// specific formatting (a run), the overlay winning where non-null.
/// </summary>
public sealed record ResolvedRunFormatting
{
    /// <summary>The identity/empty formatting — nothing set.</summary>
    public static ResolvedRunFormatting Empty { get; } = new();

    /// <summary>Font family name. Null = inherit from style/document default.</summary>
    public string? FontFamily { get; init; }

    /// <summary>Font size in points. Null = inherit.</summary>
    public double? FontSizePt { get; init; }

    /// <summary>Resolved text color as a normalized #RRGGBB literal. Null = inherit.</summary>
    public string? ColorHex { get; init; }

    /// <summary>Bold. Null = inherit, true = force bold.</summary>
    public bool? Bold { get; init; }

    /// <summary>Italic. Null = inherit, true = force italic.</summary>
    public bool? Italic { get; init; }

    /// <summary>Underline. Null = inherit, true = force single underline.</summary>
    public bool? Underline { get; init; }

    /// <summary>All-caps rendering. Null = inherit, true = force <c>w:caps</c>.</summary>
    public bool? AllCaps { get; init; }

    /// <summary>True when no field is set; such a value needs no properties emitted.</summary>
    public bool IsEmpty =>
        FontFamily is null && FontSizePt is null && ColorHex is null &&
        Bold is null && Italic is null && Underline is null && AllCaps is null;

    /// <summary>
    /// Overlays <paramref name="over"/> onto this instance: each non-null field of the overlay wins,
    /// everything else is inherited. Used to compose a typography-token base with a run override.
    /// </summary>
    public ResolvedRunFormatting Overlay(ResolvedRunFormatting over) =>
        over.IsEmpty
            ? this
            : new ResolvedRunFormatting
            {
                FontFamily = over.FontFamily ?? FontFamily,
                FontSizePt = over.FontSizePt ?? FontSizePt,
                ColorHex = over.ColorHex ?? ColorHex,
                Bold = over.Bold ?? Bold,
                Italic = over.Italic ?? Italic,
                Underline = over.Underline ?? Underline,
                AllCaps = over.AllCaps ?? AllCaps
            };
}
