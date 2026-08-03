using PptxEditor.Core.Generation.Schema;

namespace DocxEditor.Tests.Generation;

/// <summary>
/// Adversarial documents the validator must reject with actionable errors
/// (CSS-isms, percentages, wrap, typos with near-match suggestions).
/// </summary>
public class GenerationValidatorAdversarialTests
{
    private const string RowSlide = """
        {"version":"2.0","design":{"palette":{"primary":"#0B3D91"}},"slides":[{"type":"container","layout":{"mode":"row"},"children":[ELEMENT]}]}
        """;

    private const string CanvasSlide = """
        {"version":"2.0","design":{"palette":{"primary":"#0B3D91"}},"slides":[{"type":"container","children":[ELEMENT]}]}
        """;

    private static string InRow(string element) => RowSlide.Replace("ELEMENT", element);

    private static string OnCanvas(string element) => CanvasSlide.Replace("ELEMENT", element);

    public static IEnumerable<object[]> AdversarialDocuments()
    {
        // CSS-isms — explicitly rejected v1 non-goals.
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","flex-wrap":true},"children":[]}]}""", "flex-wrap");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","flexWrap":true},"children":[]}]}""", "flex-wrap");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","z-index":2,"children":[]}]}""", "z-index");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","zIndex":2,"children":[]}]}""", "z-index");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","padding":"10%","children":[]}]}""", "percentages are not supported");
        yield return Doc(OnCanvas("""{"type":"rect","at":{"x":0,"y":0},"size":{"w":"100%","h":50}}"""), "percentages are not supported");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row"},"wrap":true,"children":[]}]}""", "never wraps");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"wrap"},"children":[]}]}""", "not a valid layout mode");
        yield return Doc(InRow("""{"type":"rect","position":"absolute"}"""), "CSS 'position'");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","margin":8,"children":[]}]}""", "margins are not supported");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","display":"flex","children":[]}]}""", "CSS 'display'");
        yield return Doc(InRow("""{"type":"rect","width":100}"""), "CSS 'width'");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","justify-content":"center"},"children":[]}]}""", "justify-content");

        // Typos — near-match suggestions.
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layot":{"mode":"row"},"children":[]}]}""", "Did you mean 'layout'?");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","childern":[]}]}""", "Did you mean 'children'?");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","justif":"center"},"children":[]}]}""", "Did you mean 'justify'?");
        yield return Doc(InRow("""{"type":"rect","fill":"primry"}"""), "Did you mean 'primary'?");
        yield return Doc(InRow("""{"type":"kard"}"""), "Did you mean 'card'?");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"stylesheet":"x","slides":[{"type":"container","children":[]}]}""", "unknown property 'stylesheet'");

        // Document root.
        yield return Doc("""{"version":"1.0","design":{"palette":{}},"slides":[{"type":"container","children":[]}]}""", "unsupported version");
        yield return Doc("""{"design":{"palette":{}},"slides":[{"type":"container","children":[]}]}""", "'version' is required");
        yield return Doc("""{"version":2.0,"design":{"palette":{}},"slides":[{"type":"container","children":[]}]}""", "must be a string");
        yield return Doc("""{"version":"2.0","slides":[{"type":"container","children":[]}]}""", "'design' is required");
        yield return Doc("""{"version":"2.0","design":{"palette":{}}}""", "'slides' is required");
        yield return Doc("""[]""", "root must be a JSON object");
        yield return Doc("""{""", "malformed JSON");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[]}""", "at least one slide");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"rect"}]}""", "root must be a 'container'");
        yield return Doc("""{"version":"2.0","design":{"palette":{"primary":"blue"}},"slides":[{"type":"container","children":[]}]}""", "must be #RRGGBB");

        // Elements and layout semantics.
        yield return Doc(InRow("""{"type":"squiggle"}"""), "unknown element type 'squiggle'");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","children":{}}]}""", "must be an array");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"grid"},"children":[]}]}""", "require 'cols'");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","rowGap":4},"children":[]}]}""", "'rowGap' is only valid for grid");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","cols":2},"children":[]}]}""", "'cols' is only valid for grid");
        yield return Doc(InRow("""{"type":"rect","at":{"x":0,"y":0}}"""), "not allowed inside a layout container");
        yield return Doc(OnCanvas("""{"type":"rect","size":{"w":10,"h":10}}"""), "require 'at'");
        yield return Doc(OnCanvas("""{"type":"rect","at":{"x":0,"y":0}}"""), "fixed 'size'");
        yield return Doc(OnCanvas("""{"type":"rect","at":{"x":0,"y":0},"size":{"w":10,"h":10,"grow":1}}"""), "requires a layout container parent");
        yield return Doc(InRow("""{"type":"text"}"""), "exactly one of 'text' (string) or 'runs' (array)");
        yield return Doc(InRow("""{"type":"text","text":"a","runs":[{"text":"b"}]}"""), "exactly one of 'text' (string) or 'runs' (array)");
        yield return Doc(InRow("""{"type":"rect","fill":"#12345"}"""), "not a valid hex color");
        yield return Doc(InRow("""{"type":"rect","size":{"aspect":"16:9:3"}}"""), "not a valid aspect ratio");
        yield return Doc(InRow("""{"type":"rect","size":{"aspect":"0:9"}}"""), "not a valid aspect ratio");
        yield return Doc(InRow("""{"type":"rect","fill":{"angle":45,"stops":[{"color":"primary","offset":0}]}}"""), "at least 2 stops");
        yield return Doc(InRow("""{"type":"rect","fill":{"angle":45,"stops":[{"color":"primary","offset":0},{"color":"primary","offset":2}]}}"""), "must be ≤ 1");
        yield return Doc(InRow("""{"type":"rect","fill":{"stops":[{"color":"primary","offset":0},{"color":"primary","offset":1}]}}"""), "'angle' is required");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","layout":{"mode":"row","gap":-4},"children":[]}]}""", "must be ≥ 0");
        yield return Doc(InRow("""{"type":"rect","size":{"grow":-1}}"""), "must be ≥ 0");
        yield return Doc(InRow("""{"type":"text","text":"a","overflow":"ignore"}"""), "not a valid overflow policy");
        yield return Doc(InRow("""{"type":"rect","overflow":"clip"}"""), "only valid on containers and text");
        yield return Doc("""{"version":"2.0","design":{"palette":{},"shape":{"cardStyle":"glow"}},"slides":[{"type":"container","children":[]}]}""", "not a valid card style");
        yield return Doc(InRow("""{"type":"line","orientation":"diagonal"}"""), "not a valid line orientation");
        yield return Doc(InRow("""{"type":"image","fit":"fill"}"""), "'src' is required");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","padding":[1,2,3],"children":[]}]}""", "2 ([v, h]) or 4");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","size":{"w":10},"children":[]}]}""", "takes its size from the slide");
        yield return Doc("""{"version":"2.0","design":{"palette":{}},"slides":[{"type":"container","at":{"x":0,"y":0},"children":[]}]}""", "fills the slide");
    }

    private static object[] Doc(string json, string expectedFragment) => [json, expectedFragment];

    [Theory]
    [MemberData(nameof(AdversarialDocuments))]
    public void Validate_AdversarialDocument_RejectedWithActionableError(string json, string expectedFragment)
    {
        var result = new GenerationDocumentParser().Validate(json);

        Assert.False(result.IsValid, $"Expected rejection of: {json}");
        Assert.Null(result.Document);
        Assert.Contains(result.Errors, e => e.ToString().Contains(expectedFragment, StringComparison.Ordinal));
        Assert.All(result.Errors, e =>
        {
            Assert.False(string.IsNullOrWhiteSpace(e.Path));
            Assert.False(string.IsNullOrWhiteSpace(e.Message));
        });
    }

    [Fact]
    public void AdversarialSuite_CoversAtLeastTwentyDocuments()
    {
        Assert.True(AdversarialDocuments().Count() >= 20);
    }
}
