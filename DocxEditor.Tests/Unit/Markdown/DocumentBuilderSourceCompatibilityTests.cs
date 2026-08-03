using DocxEditor.Core.Builders;
using DocxEditor.Core.Content;
using DocxEditor.Core.Models;
using OfficeEditor.Core.Models;

namespace DocxEditor.Tests.Unit.Markdown;

/// <summary>
/// The rich-markdown members on <see cref="IDocumentBuilder"/> (AddRichMarkdown,
/// ReplaceWithRichMarkdown, LastRichMarkdownResult) carry default interface implementations,
/// so implementations written against the pre-rich-markdown surface stay source-compatible:
/// they compile without declaring the new members. Only the defaults run for them — a
/// descriptive NotSupportedException and a null result — while <see cref="DocumentBuilder"/>
/// overrides all three and keeps full behavior.
/// </summary>
public class DocumentBuilderSourceCompatibilityTests
{
    [Fact]
    public void LegacyImplementation_CompilesWithoutRichMembers_DefaultsAreSourceCompatible()
    {
        IDocumentBuilder builder = new LegacyDocumentBuilder();

        Assert.Throws<NotSupportedException>(() => builder.AddRichMarkdown("# X"));
        Assert.Throws<NotSupportedException>(() => builder.ReplaceWithRichMarkdown("T", "# X"));
        Assert.Null(builder.LastRichMarkdownResult);
    }

    /// <summary>
    /// A deliberately minimal implementation that declares every member except the three
    /// rich-markdown ones. It compiles only because those three have default implementations.
    /// </summary>
    private sealed class LegacyDocumentBuilder : IDocumentBuilder
    {
        public IDocumentBuilder AddParagraph(string text, string? style = null) => this;
        public IDocumentBuilder InsertAfter(string targetText, string text, string? style = null) => this;
        public IDocumentBuilder InsertBefore(string targetText, string text, string? style = null) => this;
        public IDocumentBuilder ReplaceText(string find, string replace) => this;
        public IDocumentBuilder ReplaceParagraph(string targetText, string newText, string? style = null) => this;
        public IDocumentBuilder DeleteParagraph(string targetText) => this;
        public IDocumentBuilder ApplyStyle(string styleId) => this;
        public IDocumentBuilder AddRichContent(List<ContentBlock> blocks) => this;
        public IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks) => this;
        public IDocumentBuilder AddHyperlink(string url, string displayText, string? style = null) => this;
        public IDocumentBuilder AddMarkdown(string markdown, StyleMapping? styleMap = null) => this;
        public IDocumentBuilder ReplaceWithMarkdown(string targetText, string markdown, StyleMapping? styleMap = null) => this;
        public List<VariableInfo> DetectVariables() => [];
        public IDocumentBuilder MergeVariables(Dictionary<string, string> data) => this;
        public void MergeBatch(List<Dictionary<string, string>> records, string outputPattern, string? templatePath = null)
        {
        }

        public void Save(string? path = null)
        {
        }

        public void Save(Stream stream)
        {
        }

        public byte[] SaveToBytes() => [];

        public static IDocumentBuilder Create() => throw new NotSupportedException();
        public static IDocumentBuilder Open(Stream stream) => throw new NotSupportedException();
        public static IDocumentBuilder Open(byte[] bytes) => throw new NotSupportedException();

        public void Dispose()
        {
        }
    }
}
