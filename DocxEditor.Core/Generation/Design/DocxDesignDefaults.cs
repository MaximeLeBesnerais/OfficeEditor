using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Built-in (foundation) design defaults used when a document's <c>design</c> tokens are absent or
/// incomplete. Fonts follow the Word convention (Calibri body / Calibri Light display); page
/// geometry mirrors <see cref="PageSize.Default"/> and <see cref="Margins.Defaults"/> (A4 portrait,
/// 1 inch margins). Baseline style sizes (title, headings, code) live here too so the style
/// generator and the resolver agree on one deterministic scale.
/// </summary>
public static class DocxDesignDefaults
{
    /// <summary>Built-in body font family (Word's classic default).</summary>
    public const string BodyFont = "Calibri";

    /// <summary>Built-in display/heading font family (Word's classic heading default).</summary>
    public const string DisplayFont = "Calibri Light";

    /// <summary>Built-in body text size in points.</summary>
    public const double BodyFontSizePt = 11;

    /// <summary>Built-in title size in points.</summary>
    public const double TitleFontSizePt = 28;

    /// <summary>Built-in monospace family used by the code baseline style.</summary>
    public const string CodeFont = "Consolas";

    /// <summary>Built-in code text size in points.</summary>
    public const double CodeFontSizePt = 10;

    /// <summary>Deterministic heading size scale for levels 1..6 (matches the Word heading ramp).</summary>
    public static double HeadingFontSizePt(int level) => level switch
    {
        1 => 16,
        2 => 13,
        3 => 12,
        4 => 11,
        5 => 10,
        _ => 10
    };
}
