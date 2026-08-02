using DocxEditor.Core.Generation.Design;
using DocxEditor.Core.Generation.Model;
using DocxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation.Docx;

/// <summary>
/// Schema/parser/validator coverage for the themed design system: optional <c>design.theme</c>,
/// semantic <c>role</c>, layout guardrails/density, <c>allCaps</c>, and v1 backward compatibility
/// (additive fields never break existing JSON).
/// </summary>
public class DocxGenerationDesignSchemaTests
{
    private readonly DocxGenerationDocumentParser _parser = new();

    [Fact]
    public void Parse_MinimalDocumentWithoutDesign_IsValidAndDefaultsToEditorialTheme()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","text":"Hello"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.NotNull(result.Document);
        Assert.Null(result.Document!.Design);
        // Content colors still validate against the editorial theme palette.
        Assert.Contains("primary", DesignThemeCatalog.Editorial.Palette);
    }

    [Fact]
    public void Parse_DesignWithoutPalette_IsValidAndThemeIsStored()
    {
        var json = """
            {"version":"1.0","design":{"theme":"editorial"},"sections":[{"blocks":[{"type":"paragraph","text":"Hello"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.Equal("editorial", result.Document!.Design!.Theme);
    }

    [Fact]
    public void Parse_EmptyDesignBlock_IsValid()
    {
        var json = """
            {"version":"1.0","design":{},"sections":[{"blocks":[{"type":"paragraph","text":"Hello"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.NotNull(result.Document!.Design);
    }

    [Fact]
    public void Parse_ContentColorReferencingThemeToken_ValidatesAgainstThemePalette()
    {
        var json = """
            {"version":"1.0","design":{"fonts":{"display":"Georgia","body":"Arial"}},
             "sections":[{"blocks":[{"type":"paragraph","runs":[{"text":"x","color":"primary"}]}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
    }

    [Fact]
    public void Parse_ContentColorReferencingThemeTokenWithoutDesignBlock_Validates()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","runs":[{"text":"x","color":"coral"}]}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
    }

    [Fact]
    public void Parse_UnknownTheme_Errors()
    {
        var json = """
            {"version":"1.0","design":{"theme":"bogus"},"sections":[{"blocks":[{"type":"paragraph","text":"Hello"}]}]}
            """;

        var result = _parser.Validate(json);

        var error = Assert.Single(result.Errors);
        Assert.Contains("unknown theme", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("$.design.theme", error.Path);
    }

    [Fact]
    public void Parse_CorporateTheme_IsAccepted()
    {
        var json = """
            {"version":"1.0","design":{"theme":"corporate"},"sections":[{"blocks":[{"type":"paragraph","text":"Hello"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.Equal("corporate", result.Document!.Design!.Theme);
    }

    [Fact]
    public void Parse_ValidRoles_AreAccepted()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[
              {"type":"paragraph","role":"eyebrow","text":"Kicker"},
              {"type":"heading","level":2,"role":"heading3","text":"Title"},
              {"type":"list","items":[{"text":"item","role":"label"}]},
              {"type":"callout","role":"callout","text":"note"},
              {"type":"table","rows":[{"header":true,"cells":[{"text":"H","role":"tableHeader"}]},{"cells":[{"text":"B","role":"tableBody"}]}]}
            ]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        var document = result.Document!;
        var blocks = document.Sections[0].Blocks;
        Assert.Equal(TextRole.Eyebrow, ((ParagraphBlock)blocks[0]).Content.Role);
        Assert.Equal(TextRole.Heading3, ((HeadingBlock)blocks[1]).Content.Role);
        Assert.Equal(TextRole.Label, ((ListBlock)blocks[2]).Items[0].Role);
        Assert.Equal(TextRole.Callout, ((CalloutBlock)blocks[3]).Content.Role);
        var table = (TableBlock)blocks[4];
        Assert.Equal(TextRole.TableHeader, table.Rows[0].Cells[0].Content!.Role);
        Assert.Equal(TextRole.TableBody, table.Rows[1].Cells[0].Content!.Role);
    }

    [Fact]
    public void Parse_UnknownRole_Errors()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","role":"hero","text":"x"}]}]}
            """;

        var result = _parser.Validate(json);

        var error = Assert.Single(result.Errors);
        Assert.Contains("not a valid text role", error.Message);
    }

    [Fact]
    public void Parse_LayoutGuardrails_AreAccepted()
    {
        var json = """
            {"version":"1.0","design":{"layout":{"density":"spacious","minBodySizePt":10,"maxTableWidthPt":420}},
             "sections":[{"blocks":[{"type":"paragraph","text":"Hello"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        var layout = result.Document!.Design!.Layout;
        Assert.Equal(Density.Spacious, layout.Density);
        Assert.Equal(10, layout.MinBodySizePt);
        Assert.Equal(420, layout.MaxTableWidthPt);
    }

    [Fact]
    public void Parse_InvalidDensity_Errors()
    {
        var json = """
            {"version":"1.0","design":{"layout":{"density":"huge"}},"sections":[{"blocks":[{"type":"paragraph","text":"x"}]}]}
            """;

        var result = _parser.Validate(json);

        var error = Assert.Single(result.Errors);
        Assert.Contains("not a valid density", error.Message);
        Assert.Equal("$.design.layout.density", error.Path);
    }

    [Fact]
    public void Parse_InvalidLayoutGuardrails_Errors()
    {
        var json = """
            {"version":"1.0","design":{"layout":{"minBodySizePt":-2}},"sections":[{"blocks":[{"type":"paragraph","text":"x"}]}]}
            """;

        var result = _parser.Validate(json);

        var error = Assert.Single(result.Errors);
        Assert.Equal("$.design.layout.minBodySizePt", error.Path);
    }

    [Fact]
    public void Parse_AllCapsOnRun_IsAccepted()
    {
        var json = """
            {"version":"1.0","sections":[{"blocks":[{"type":"paragraph","runs":[{"text":"UP","allCaps":true}]}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.True(((ParagraphBlock)result.Document!.Sections[0].Blocks[0]).Content.Runs![0].AllCaps);
    }

    [Fact]
    public void Parse_AllCapsOnTypographyToken_IsAccepted()
    {
        var json = """
            {"version":"1.0","design":{"typography":{"kicker":{"font":"body","size":10,"allCaps":true}}},
             "sections":[{"blocks":[{"type":"paragraph","token":"kicker","text":"UP"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.True(result.Document!.Design!.Typography["kicker"].AllCaps);
    }

    [Fact]
    public void Parse_UnknownPropertyOnDesign_Errors()
    {
        var json = """
            {"version":"1.0","design":{"layoutz":{}},"sections":[{"blocks":[{"type":"paragraph","text":"x"}]}]}
            """;

        var result = _parser.Validate(json);

        var error = Assert.Single(result.Errors);
        Assert.Contains("unknown property 'layoutz'", error.Message);
    }

    [Fact]
    public void Parse_FullLegacyDesignBlock_RemainsBackwardCompatible()
    {
        // The v1 vocabulary's full design block (explicit palette/fonts/typography/spacing/
        // shapes/page) must still parse with zero errors — the new fields are additive only.
        var json = """
            {"version":"1.0","design":{
              "palette":{"ink":"#1F2937","navy":"#1F4E79","blue":"#2563EB"},
              "fonts":{"display":"Aptos Display","body":"Aptos"},
              "typography":{"title":{"font":"display","size":26,"color":"navy","bold":true}},
              "spacing":{"tight":3,"normal":8},
              "shapes":{"cornerRadius":6,"defaultFill":"navy"},
              "page":{"size":"letter","orientation":"portrait","margins":{"top":54,"right":54,"bottom":54,"left":54}}
             },"sections":[{"blocks":[{"type":"heading","level":1,"text":"T","token":"title"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Parse_ThemeAndFullDesignTogether_PreservesExplicitOverrides()
    {
        var json = """
            {"version":"1.0","design":{"theme":"corporate","palette":{"primary":"#FF0000"},
              "fonts":{"display":"Times New Roman"}},
             "sections":[{"blocks":[{"type":"paragraph","text":"x"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
        var design = result.Document!.Design!;
        Assert.Equal("corporate", design.Theme);
        Assert.Equal("#FF0000", design.Palette["primary"]);
        Assert.Equal("Times New Roman", design.Fonts.Display);
    }

    [Fact]
    public void Parse_DuplicateRoleAndToken_IsAccepted()
    {
        var json = """
            {"version":"1.0","design":{"typography":{"hero":{"font":"display","size":24}}},
             "sections":[{"blocks":[{"type":"paragraph","role":"title","token":"hero","text":"x"}]}]}
            """;

        var result = _parser.Validate(json);

        Assert.Empty(result.Errors);
    }
}
