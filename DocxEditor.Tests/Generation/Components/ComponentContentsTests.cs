using System.Text.Json;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Model;
using PptxEditor.Core.Models;

namespace DocxEditor.Tests.Generation.Components;

/// <summary>
/// Direct-coverage tests for the strongly typed content payloads in
/// <c>ComponentContents.cs</c>. Each record type is tested for required-property
/// construction, optional-property defaults, and the JSON-bag → typed-content
/// validation round-trip via <see cref="ComponentContentReader"/> (the internal
/// reader that the component layer uses to bridge raw JSON to these records).
/// </summary>
public sealed class ComponentContentsTests
{
    // ------------------------------------------------------------------ CardContent

    [Fact]
    public void CardContent_Construction_WithRequiredTitleOnly()
    {
        var content = new CardContent { Title = "Hello" };

        Assert.Equal("Hello", content.Title);
        Assert.Null(content.Subtitle);
        Assert.Null(content.Body);
    }

    [Fact]
    public void CardContent_Construction_WithAllProperties()
    {
        var content = new CardContent
        {
            Title = "Revenue",
            Subtitle = "Q3 2026",
            Body = "All regions grew double digits."
        };

        Assert.Equal("Revenue", content.Title);
        Assert.Equal("Q3 2026", content.Subtitle);
        Assert.Equal("All regions grew double digits.", content.Body);
    }

    // ------------------------------------------------------------------ KpiContent

    [Fact]
    public void KpiContent_Construction_WithRequiredProperties()
    {
        var content = new KpiContent { Value = "+34%", Label = "Revenue" };

        Assert.Equal("+34%", content.Value);
        Assert.Equal("Revenue", content.Label);
        Assert.Null(content.Delta);
    }

    [Fact]
    public void KpiContent_Construction_WithOptionalDelta()
    {
        var content = new KpiContent
        {
            Value = "+34%",
            Label = "Revenue",
            Delta = "+12% vs LY"
        };

        Assert.Equal("+12% vs LY", content.Delta);
    }

    // ------------------------------------------------------------------ TitleBlockContent

    [Fact]
    public void TitleBlockContent_Construction_WithRequiredTitleOnly()
    {
        var content = new TitleBlockContent { Title = "Growth Strategy" };

        Assert.Equal("Growth Strategy", content.Title);
        Assert.Null(content.Subtitle);
        Assert.Null(content.Kicker);
    }

    [Fact]
    public void TitleBlockContent_Construction_WithAllProperties()
    {
        var content = new TitleBlockContent
        {
            Kicker = "Q3 REPORT",
            Title = "Growth",
            Subtitle = "All regions"
        };

        Assert.Equal("Q3 REPORT", content.Kicker);
        Assert.Equal("Growth", content.Title);
        Assert.Equal("All regions", content.Subtitle);
    }

    // ------------------------------------------------------------------ BulletListContent

    [Fact]
    public void BulletListContent_Construction_WithRequiredItemsOnly()
    {
        var content = new BulletListContent { Items = ["One", "Two", "Three"] };

        Assert.Equal(3, content.Items.Count);
        Assert.Equal("One", content.Items[0]);
        Assert.Null(content.Title);
        Assert.Null(content.MarkerColor);
    }

    [Fact]
    public void BulletListContent_Construction_WithTitleAndMarkerColor()
    {
        var content = new BulletListContent
        {
            Items = ["Alpha"],
            Title = "Highlights",
            MarkerColor = "primary"
        };

        Assert.Equal("Highlights", content.Title);
        Assert.Equal("primary", content.MarkerColor);
    }

    // ------------------------------------------------------------------ DividerContent

    [Fact]
    public void DividerContent_Construction_HasSensibleDefaults()
    {
        var content = new DividerContent();

        Assert.Equal("muted", content.Color);
        Assert.Equal(1, content.WidthPt);
        Assert.Equal(LineOrientation.Horizontal, content.Orientation);
    }

    [Fact]
    public void DividerContent_Construction_AllPropertiesCanBeSet()
    {
        var content = new DividerContent
        {
            Color = "primary",
            WidthPt = 2.5,
            Orientation = LineOrientation.Vertical
        };

        Assert.Equal("primary", content.Color);
        Assert.Equal(2.5, content.WidthPt);
        Assert.Equal(LineOrientation.Vertical, content.Orientation);
    }

    // ------------------------------------------------------------------ BadgeContent

    [Fact]
    public void BadgeContent_Construction_WithRequiredTextUsesDefaultsForColors()
    {
        var content = new BadgeContent { Text = "NEW" };

        Assert.Equal("NEW", content.Text);
        Assert.Equal("accent", content.Color);
        Assert.Equal("paper", content.TextColor);
    }

    [Fact]
    public void BadgeContent_Construction_WithCustomColors()
    {
        var content = new BadgeContent
        {
            Text = "SALE",
            Color = "primary",
            TextColor = "ink"
        };

        Assert.Equal("SALE", content.Text);
        Assert.Equal("primary", content.Color);
        Assert.Equal("ink", content.TextColor);
    }

    // ------------------------------------------------------------------ ImageCardContent

    [Fact]
    public void ImageCardContent_Construction_WithRequiredSource()
    {
        var content = new ImageCardContent { Source = "hero.png" };

        Assert.Equal("hero.png", content.Source);
        Assert.Null(content.Title);
        Assert.Null(content.Subtitle);
        Assert.Equal(ImageFitMode.Crop, content.Fit);
        Assert.Null(content.Alt);
        Assert.Equal(1, content.ImageGrow);
    }

    [Fact]
    public void ImageCardContent_Construction_WithAllProperties()
    {
        var content = new ImageCardContent
        {
            Source = "team.jpg",
            Title = "Leadership",
            Subtitle = "Executive team",
            Fit = ImageFitMode.Contain,
            Alt = "Team photo",
            ImageGrow = 1.5
        };

        Assert.Equal("team.jpg", content.Source);
        Assert.Equal("Leadership", content.Title);
        Assert.Equal("Executive team", content.Subtitle);
        Assert.Equal(ImageFitMode.Contain, content.Fit);
        Assert.Equal("Team photo", content.Alt);
        Assert.Equal(1.5, content.ImageGrow);
    }

    // ------------------------------------------------------------------ TableBlockContent

    [Fact]
    public void TableBlockContent_Construction_WithRequiredColumnsAndRows()
    {
        var columns = new List<string> { "Region", "Revenue" };
        var rows = new List<IReadOnlyList<string>>
        {
            new List<string> { "EU", "12M" },
            new List<string> { "US", "9M" }
        };

        var content = new TableBlockContent { Columns = columns, Rows = rows };

        Assert.Equal(2, content.Columns.Count);
        Assert.Equal("Region", content.Columns[0]);
        Assert.Equal(2, content.Rows.Count);
        Assert.True(content.Header);
        Assert.Null(content.ColumnWeights);
        Assert.Equal(0, content.RowHeight);
    }

    [Fact]
    public void TableBlockContent_Construction_WithAllProperties()
    {
        var content = new TableBlockContent
        {
            Columns = ["A", "B", "C"],
            Rows = [new List<string> { "1", "2", "3" }],
            Header = false,
            ColumnWeights = [2, 1, 1],
            RowHeight = 36
        };

        Assert.False(content.Header);
        Assert.Equal(3, content.ColumnWeights!.Count);
        Assert.Equal(2, content.ColumnWeights[0]);
        Assert.Equal(36, content.RowHeight);
    }

    // ------------------------------------------------------------------ inheritance & records

    [Fact]
    public void AllRecords_AreAssignableToComponentContent()
    {
        ComponentContent[] contents =
        [
            new CardContent { Title = "T" },
            new KpiContent { Value = "V", Label = "L" },
            new TitleBlockContent { Title = "T" },
            new BulletListContent { Items = ["a"] },
            new DividerContent(),
            new BadgeContent { Text = "X" },
            new ImageCardContent { Source = "a.png" },
            new TableBlockContent { Columns = ["A"], Rows = [] }
        ];

        Assert.Equal(8, contents.Length);
    }

    [Fact]
    public void AllRecordTypes_AreImmutable_EqualityWorksByValue()
    {
        var a = new CardContent { Title = "T", Subtitle = "S" };
        var b = new CardContent { Title = "T", Subtitle = "S" };
        var c = new CardContent { Title = "T", Subtitle = "Different" };

        Assert.Equal(a, b);
        Assert.NotEqual(a, c);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
    }

    // ------------------------------------------------------------------ ComponentContentReader record-building (production path)

    [Fact]
    public void ContentReader_ParsesCardContent()
    {
        var reader = Reader("card", """{"title":"Growth","subtitle":"Q3","body":"Strong performance"}""",
            Props("title", "subtitle", "body"));
        var content = new CardContent
        {
            Title = reader.String("title", required: true)!,
            Subtitle = reader.String("subtitle"),
            Body = reader.String("body")
        };
        reader.ThrowIfInvalid();

        Assert.Equal("Growth", content.Title);
        Assert.Equal("Q3", content.Subtitle);
        Assert.Equal("Strong performance", content.Body);
    }

    [Fact]
    public void ContentReader_ParsesKpiContent()
    {
        var reader = Reader("kpi", """{"value":"+34%","label":"Revenue","delta":"+12% vs LY"}""",
            Props("value", "label", "delta"));
        var content = new KpiContent
        {
            Value = reader.String("value", required: true)!,
            Label = reader.String("label", required: true)!,
            Delta = reader.String("delta")
        };
        reader.ThrowIfInvalid();

        Assert.Equal("+34%", content.Value);
        Assert.Equal("Revenue", content.Label);
        Assert.Equal("+12% vs LY", content.Delta);
    }

    [Fact]
    public void ContentReader_ParsesTitleBlockContent_AllProperties()
    {
        var reader = Reader("title_block", """{"kicker":"Q3 REPORT","title":"Growth","subtitle":"All regions"}""",
            Props("kicker", "title", "subtitle"));
        var content = new TitleBlockContent
        {
            Kicker = reader.String("kicker"),
            Title = reader.String("title", required: true)!,
            Subtitle = reader.String("subtitle")
        };
        reader.ThrowIfInvalid();

        Assert.Equal("Q3 REPORT", content.Kicker);
        Assert.Equal("Growth", content.Title);
        Assert.Equal("All regions", content.Subtitle);
    }

    [Fact]
    public void ContentReader_ParsesBulletListContent()
    {
        var reader = Reader("bullet_list", """{"items":["Alpha","Beta","Gamma"],"title":"Key Metrics"}""",
            Props("items", "title", "markerColor"));
        var content = new BulletListContent
        {
            Items = reader.StringArray("items", required: true, minCount: 1)!,
            Title = reader.String("title"),
            MarkerColor = reader.Color("markerColor")
        };
        reader.ThrowIfInvalid();

        Assert.Equal(3, content.Items.Count);
        Assert.Contains("Alpha", content.Items);
        Assert.Equal("Key Metrics", content.Title);
        Assert.Null(content.MarkerColor);
    }

    [Fact]
    public void ContentReader_ParsesDividerContent_Defaults()
    {
        var reader = Reader("divider", "{}", Props("color", "widthPt", "orientation"));
        var content = new DividerContent
        {
            Color = reader.Color("color") ?? "muted",
            WidthPt = reader.Number("widthPt") ?? 1,
            Orientation = reader.Enum("orientation", new Dictionary<string, LineOrientation>(StringComparer.OrdinalIgnoreCase)
            {
                ["horizontal"] = LineOrientation.Horizontal,
                ["vertical"] = LineOrientation.Vertical
            }, "orientation") ?? LineOrientation.Horizontal
        };
        reader.ThrowIfInvalid();

        Assert.Equal("muted", content.Color);
        Assert.Equal(1, content.WidthPt);
        Assert.Equal(LineOrientation.Horizontal, content.Orientation);
    }

    [Fact]
    public void ContentReader_ParsesImageCardContent_Minimal()
    {
        var reader = Reader("image_card", """{"source":"hero.png"}""",
            Props("source", "title", "subtitle", "fit", "alt", "imageGrow"));
        var content = new ImageCardContent
        {
            Source = reader.String("source", required: true)!,
            Title = reader.String("title"),
            Subtitle = reader.String("subtitle")
        };
        reader.ThrowIfInvalid();

        Assert.Equal("hero.png", content.Source);
        Assert.Null(content.Title);
        Assert.Null(content.Subtitle);
        Assert.Equal(ImageFitMode.Crop, content.Fit);
    }

    [Fact]
    public void ContentReader_ParsesTableBlockContent()
    {
        var reader = Reader("table_block",
            """{"columns":["Region","Rev"],"rows":[["EU","12"]],"header":false,"columnWeights":[2,1],"rowHeight":36}""",
            Props("columns", "rows", "header", "columnWeights", "rowHeight"));
        var content = new TableBlockContent
        {
            Columns = reader.StringArray("columns", required: true, minCount: 1)!,
            Rows = reader.StringMatrix("rows") ?? [],
            Header = reader.Bool("header", true),
            ColumnWeights = reader.NumberArray("columnWeights"),
            RowHeight = reader.Number("rowHeight") ?? 0
        };
        reader.ThrowIfInvalid();

        Assert.Equal(2, content.Columns.Count);
        Assert.Single(content.Rows);
        Assert.False(content.Header);
        Assert.Equal(new[] { 2.0, 1.0 }, (IEnumerable<double>)content.ColumnWeights!);
        Assert.Equal(36, content.RowHeight);
    }

    // ------------------------------------------------------------------ ComponentContentReader integration

    private static readonly IReadOnlyDictionary<string, string> Palette = new Dictionary<string, string>
    {
        ["primary"] = "#0B3D91",
        ["accent"] = "#FF6B00",
        ["ink"] = "#1A1A1A",
        ["paper"] = "#FFFFFF",
        ["muted"] = "#8A94A6"
    };

    private static ComponentContentReader Reader(string componentName, string contentJson, IReadOnlySet<string>? allowedProps = null)
    {
        var set = allowedProps ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var content = JsonDocument.Parse(contentJson).RootElement.Clone();
        return ComponentContentReader.For("test", componentName, content, set, Palette);
    }

    private static HashSet<string> Props(params string[] names)
        => new(names, StringComparer.OrdinalIgnoreCase);

    [Fact]
    public void ContentReader_ValidCardPayload_DoesNotThrow()
    {
        var reader = Reader("card", """{"title":"Revenue","subtitle":"Q3","body":"All regions"}""", Props("title", "subtitle", "body"));
        reader.ThrowIfInvalid();
        Assert.True(reader.HasContent);
    }

    [Fact]
    public void ContentReader_NullContent_IsInvalidWhenRequired()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var reader = ComponentContentReader.For("test", "card", null, set, Palette, required: true);

        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("'content' is required for component 'card'", ex.Message);
    }

    [Fact]
    public void ContentReader_NullContent_IsValidWhenNotRequired()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var reader = ComponentContentReader.For("test", "divider", null, set, Palette, required: false);

        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_NonObjectContent_IsInvalid()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var content = JsonDocument.Parse("42").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "card", content, set, Palette);

        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("must be an object", ex.Message);
    }

    [Fact]
    public void ContentReader_UnknownProperty_SuggestsCorrection()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "title", "subtitle", "body" };
        var content = JsonDocument.Parse("""{"title":"T","subtile":"S"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "card", content, set, Palette);

        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("unknown property 'subtile'", ex.Message);
        Assert.Contains("Did you mean 'subtitle'?", ex.Message);
    }

    [Fact]
    public void ContentReader_String_RequiredAndOptional()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "title", "subtitle" };
        var content = JsonDocument.Parse("""{"title":"Hello"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "card", content, set, Palette);

        Assert.Equal("Hello", reader.String("title", required: true));
        Assert.Null(reader.String("subtitle"));
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_String_MissingRequired_CollectsError()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "title" };
        var content = JsonDocument.Parse("""{}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "card", content, set, Palette);

        reader.String("title", required: true);
        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("'title' is required", ex.Message);
    }

    [Fact]
    public void ContentReader_Number_ValidAndOutOfRange()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "width", "tiny" };
        var content = JsonDocument.Parse("""{"width":2.5,"tiny":-1}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "divider", content, set, Palette);

        Assert.Equal(2.5, reader.Number("width"));
        Assert.Null(reader.Number("tiny", min: 0, max: 100));
        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("must be between", ex.Message);
    }

    [Fact]
    public void ContentReader_Bool_DefaultsAndOverrides()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "header" };
        var content = JsonDocument.Parse("""{"header":false}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "table_block", content, set, Palette);

        Assert.False(reader.Bool("header", true));
        Assert.True(reader.Bool("missing", true));
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_Color_ValidTokenAndHex()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "color", "hex" };
        var content = JsonDocument.Parse("""{"color":"accent","hex":"#FF6B00"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "badge", content, set, Palette);

        Assert.Equal("accent", reader.Color("color"));
        Assert.Equal("#FF6B00", reader.Color("hex"));
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_Color_Invalid_CollectsError()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "color" };
        var content = JsonDocument.Parse("""{"color":"not_a_token"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "badge", content, set, Palette);

        Assert.Null(reader.Color("color"));
        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("unknown color 'not_a_token'", ex.Message);
    }

    [Fact]
    public void ContentReader_Enum_ValidAndInvalid()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "orientation" };
        var map = new Dictionary<string, LineOrientation>(StringComparer.OrdinalIgnoreCase)
        {
            ["horizontal"] = LineOrientation.Horizontal,
            ["vertical"] = LineOrientation.Vertical
        };
        var content = JsonDocument.Parse("""{"orientation":"vertical"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "divider", content, set, Palette);

        Assert.Equal(LineOrientation.Vertical, reader.Enum("orientation", map, "orientation"));
        Assert.Null(reader.Enum<LineOrientation>("missing", map, "orientation"));
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_StringArray_ValidAndMinCount()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "items" };
        var content = JsonDocument.Parse("""{"items":["a","b","c"]}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "bullet_list", content, set, Palette);

        var items = reader.StringArray("items", required: true, minCount: 1);
        Assert.NotNull(items);
        Assert.Equal(3, items.Count);
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_StringArray_TooFew_CollectsError()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "items" };
        var content = JsonDocument.Parse("""{"items":[]}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "bullet_list", content, set, Palette);

        reader.StringArray("items", required: true, minCount: 1);
        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("must contain at least 1 item(s)", ex.Message);
    }

    [Fact]
    public void ContentReader_StringMatrix_ValidRows()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "rows" };
        var content = JsonDocument.Parse("""{"rows":[["a","b"],["c","d"]]}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "table_block", content, set, Palette);

        var rows = reader.StringMatrix("rows");
        Assert.NotNull(rows);
        Assert.Equal(2, rows.Count);
        Assert.Equal(2, rows[0].Count);
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_NumberArray_ValidWeights()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "columnWeights" };
        var content = JsonDocument.Parse("""{"columnWeights":[2,1,1]}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "table_block", content, set, Palette);

        var weights = reader.NumberArray("columnWeights");
        Assert.NotNull(weights);
        Assert.Equal(new[] { 2.0, 1.0, 1.0 }, weights);
        reader.ThrowIfInvalid();
    }

    [Fact]
    public void ContentReader_CollectsMultipleErrors()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "title", "subtitle" };
        var content = JsonDocument.Parse("""{"title":42,"subtitle":true,"badProp":"x"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "card", content, set, Palette);

        reader.String("title", required: true);
        reader.String("subtitle");
        // "badProp" was already picked up by For()

        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("must be a string", ex.Message);
        Assert.Contains("unknown property 'badProp'", ex.Message);
    }

    [Fact]
    public void ContentReader_PercentageStrings_AreRejected()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "width" };
        var content = JsonDocument.Parse("""{"width":"50%"}""").RootElement.Clone();
        var reader = ComponentContentReader.For("test", "divider", content, set, Palette);

        Assert.Null(reader.Number("width"));
        var ex = Assert.Throws<ComponentException>(reader.ThrowIfInvalid);
        Assert.Contains("percentages are not supported", ex.Message);
    }
}
