using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Resolution coverage for the built-in theme catalog: editorial defaults, document-over-theme
/// override precedence, density scaling, role defaults and unknown-theme fallback.
/// </summary>
public class DocxThemeResolverTests
{
    [Fact]
    public void NullDesign_ResolvesToEditorialThemeDefaults()
    {
        var resolver = new DocxDesignResolver();
        var resolved = resolver.ResolveAll();

        Assert.Equal("editorial", resolved.ThemeName);
        Assert.Equal("Georgia", resolved.DisplayFontFamily);
        Assert.Equal("Arial", resolved.BodyFontFamily);
        Assert.Equal("#1F3A5F", resolved.Palette["primary"]);
        Assert.Equal("#E4674A", resolved.Palette["coral"]);
        Assert.Equal("#EEF2F6", resolved.Palette["pale"]);
        Assert.Empty(resolver.Warnings);
    }

    [Fact]
    public void EditorialTheme_ProvidesSemanticRoleDefaults()
    {
        var resolved = new DocxDesignResolver().ResolveAll();

        Assert.Equal(11, resolved.Roles[TextRole.Body].Run.FontSizePt);
        Assert.Equal("Arial", resolved.Roles[TextRole.Body].Run.FontFamily);
        Assert.Equal(8, resolved.Roles[TextRole.Body].Paragraph.SpaceAfterPt);
        Assert.Equal(1.3, resolved.Roles[TextRole.Body].Paragraph.LineSpacingMultiple);

        Assert.Equal(20, resolved.Roles[TextRole.Heading1].Run.FontSizePt);
        Assert.Equal("Georgia", resolved.Roles[TextRole.Heading1].Run.FontFamily);
        Assert.True(resolved.Roles[TextRole.Heading1].KeepNext);
        Assert.True(resolved.Roles[TextRole.Heading1].KeepLines);

        Assert.True(resolved.Roles[TextRole.Eyebrow].Run.AllCaps is true);
        Assert.Equal(TextAlignment.Center, resolved.Roles[TextRole.Footer].Paragraph.Alignment);
        Assert.Equal("Arial", resolved.Roles[TextRole.TableHeader].Run.FontFamily);
        Assert.Equal("#FFFFFF", resolved.Roles[TextRole.TableHeader].Run.ColorHex);
    }

    [Fact]
    public void DesignBlock_OverridesThemeFieldByField()
    {
        var design = new DesignTokens
        {
            Palette = new Dictionary<string, string> { ["primary"] = "#FF0000", ["extra"] = "#123456" },
            Fonts = new FontTokens { Display = "Times New Roman", Body = "Verdana" },
            Layout = new LayoutDefaults { MinBodySizePt = 10 }
        };
        var resolved = new DocxDesignResolver(design).ResolveAll();

        Assert.Equal("editorial", resolved.ThemeName);
        Assert.Equal("#FF0000", resolved.Palette["primary"]);
        Assert.Equal("#123456", resolved.Palette["extra"]);
        Assert.Equal("Times New Roman", resolved.DisplayFontFamily);
        Assert.Equal("Verdana", resolved.BodyFontFamily);
        Assert.Equal(10, resolved.Layout.MinBodySizePt);
    }

    [Fact]
    public void DesignPalette_OverridesThemePaletteAndKeepsUntouchedThemeTokens()
    {
        var design = new DesignTokens { Palette = new Dictionary<string, string> { ["primary"] = "#FF0000" } };
        var resolved = new DocxDesignResolver(design).ResolveAll();

        Assert.Equal("#FF0000", resolved.Palette["primary"]);
        Assert.Equal("#1F7A6E", resolved.Palette["teal"]);
        Assert.Equal("#5A6B7B", resolved.Palette["muted"]);
    }

    [Fact]
    public void ExplicitTheme_Corporate_ChangesFontsPaletteAndPage()
    {
        var design = new DesignTokens { Theme = "corporate" };
        var resolver = new DocxDesignResolver(design);
        var resolved = resolver.ResolveAll();

        Assert.Equal("corporate", resolved.ThemeName);
        Assert.Equal("Trebuchet MS", resolved.DisplayFontFamily);
        Assert.Equal("Arial", resolved.BodyFontFamily);
        Assert.Equal("#14498C", resolved.Palette["primary"]);
        // Catalog page sizes are named by their dimensions (Letter = 612 × 792pt portrait).
        Assert.Equal(612.0, resolved.Page.PageSize.WidthPt);
        Assert.Equal(792.0, resolved.Page.PageSize.HeightPt);
        Assert.Empty(resolver.Warnings);
    }

    [Fact]
    public void UnknownTheme_FallsBackToEditorial_WithWarning()
    {
        var design = new DesignTokens { Theme = "bogus" };
        var resolver = new DocxDesignResolver(design);
        var resolved = resolver.ResolveAll();

        Assert.Equal("editorial", resolved.ThemeName);
        var warning = Assert.Single(resolver.Warnings);
        Assert.Equal("UnknownTheme", warning.Code);
        Assert.Contains("bogus", warning.Message);
    }

    [Fact]
    public void DensitySpacious_ScalesRoleSpacing()
    {
        var design = new DesignTokens { Layout = new LayoutDefaults { Density = Density.Spacious } };
        var resolved = new DocxDesignResolver(design).ResolveAll();

        Assert.Equal(1.4, resolved.Layout.DensityScale);
        // Body after = 8pt * 1.4 = 11.2pt; heading1 before = 22pt * 1.4 = 30.8pt.
        Assert.Equal(11.2, resolved.Roles[TextRole.Body].Paragraph.SpaceAfterPt!.Value, 3);
        Assert.Equal(30.8, resolved.Roles[TextRole.Heading1].Paragraph.SpaceBeforePt!.Value, 3);
    }

    [Fact]
    public void DensityCompact_ScalesRoleSpacing()
    {
        var design = new DesignTokens { Layout = new LayoutDefaults { Density = Density.Compact } };
        var resolved = new DocxDesignResolver(design).ResolveAll();

        Assert.Equal(0.75, resolved.Layout.DensityScale);
        Assert.Equal(6, resolved.Roles[TextRole.Body].Paragraph.SpaceAfterPt!.Value, 3);
    }

    [Fact]
    public void TryResolveRole_ReturnsNullForNullAndValueForKnownRole()
    {
        var resolver = new DocxDesignResolver();

        Assert.Null(resolver.TryResolveRole(null));
        Assert.NotNull(resolver.TryResolveRole(TextRole.Metric));
        Assert.Equal(24, resolver.TryResolveRole(TextRole.Metric)!.Run.FontSizePt);
    }

    [Fact]
    public void ContentColor_ResolvesAgainstThemePaletteWithoutDesignPalette()
    {
        var resolver = new DocxDesignResolver();

        Assert.Equal("#1F3A5F", resolver.ResolveColor("primary"));
        Assert.Equal("#1C2733", resolver.ResolveColor("ink"));
        Assert.Empty(resolver.Warnings);
    }

    [Fact]
    public void ThemeRoleColor_IsResolvedHex()
    {
        var resolved = new DocxDesignResolver().ResolveAll();

        Assert.Equal("#1F3A5F", resolved.Roles[TextRole.Title].Run.ColorHex);
        Assert.Equal("#5A6B7B", resolved.Roles[TextRole.Muted].Run.ColorHex);
        Assert.Equal("#E4674A", resolved.Roles[TextRole.Eyebrow].Run.ColorHex);
    }

    [Fact]
    public void BodyFontAndDisplayFont_ThemeDefaultsBackSlots()
    {
        var design = new DesignTokens { Palette = new Dictionary<string, string>() };
        var resolver = new DocxDesignResolver(design);

        Assert.Equal("Georgia", resolver.TryResolveFontFamily("display"));
        Assert.Equal("Arial", resolver.TryResolveFontFamily("body"));
        Assert.Empty(resolver.Warnings);
    }
}
