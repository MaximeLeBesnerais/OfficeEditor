using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Core.Generation.Design;

/// <summary>
/// Raised by the strict resolution methods (<see cref="DocxDesignResolver.ResolveColor"/>,
/// <see cref="DocxDesignResolver.ResolveTypography"/>, <see cref="DocxDesignResolver.ResolveSpacing"/>)
/// when a token name cannot be resolved — not in the design palette, not a known typography token,
/// not a known spacing token. Emitters that must have a value use the strict methods and let this
/// typed exception surface; emitters that can fall back use the <c>Try*</c> methods, which record a
/// <see cref="DesignResolutionWarning"/> instead of throwing.
/// </summary>
public sealed class DocxDesignResolutionException : OfficeEditorException
{
    /// <summary>The token name that could not be resolved.</summary>
    public string TokenName { get; }

    /// <summary>What kind of token failed (e.g. "color", "typography token", "spacing token").</summary>
    public string TokenKind { get; }

    /// <summary>Optional "Did you mean …?" suggestion.</summary>
    public string? Suggestion { get; }

    public DocxDesignResolutionException(string tokenName, string tokenKind, string? suggestion = null)
        : base(BuildMessage(tokenName, tokenKind, suggestion))
    {
        TokenName = tokenName;
        TokenKind = tokenKind;
        Suggestion = suggestion;
    }

    private static string BuildMessage(string tokenName, string tokenKind, string? suggestion)
    {
        var hint = suggestion is null ? string.Empty : $" {suggestion}";
        return $"Unresolved {tokenKind} '{tokenName}' in the DOCX design tokens.{hint}";
    }
}
