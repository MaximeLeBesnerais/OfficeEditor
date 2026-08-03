using DocxEditor.Core.Markdown.Model;
using Markdig.Extensions.Abbreviations;
using Markdig.Extensions.Alerts;
using Markdig.Extensions.AutoIdentifiers;
using Markdig.Extensions.CustomContainers;
using Markdig.Extensions.DefinitionLists;
using Markdig.Extensions.Emoji;
using Markdig.Extensions.Footnotes;
using Markdig.Extensions.Mathematics;
using Markdig.Extensions.Tables;
using Markdig.Extensions.TaskLists;
using Markdig.Extensions.Yaml;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace DocxEditor.Core.Markdown;

/// <summary>
/// Parses markdown into the rich, recursive <see cref="Model"/> IR. Unlike the legacy
/// <see cref="MarkdownParser"/> (which flattens content into <c>Models.ContentBlock</c>),
/// this parser maps the Markdig AST recursively without flattening nested formatting, and
/// retains constructs it cannot fully structure as explicit nodes plus diagnostics.
/// </summary>
public sealed class RichMarkdownParser
{
    /// <summary>
    /// Maximum nesting depth (block inside block, inline inside inline) the parser will
    /// structure. Deeper content is flattened into a verbatim node plus a diagnostic instead of
    /// recursing, so adversarial input can never drive the conversion recursion past the
    /// process stack. The renderer enforces the same bound as a second line of defense.
    /// </summary>
    public const int MaxNestingDepth = 128;

    private readonly MarkdownParseOptions _options;

    public RichMarkdownParser()
        : this(MarkdownParseOptions.Default)
    {
    }

    public RichMarkdownParser(MarkdownParseOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    /// <summary>
    /// Parses the given markdown with the default options.
    /// </summary>
    public static MarkdownParseResult ParseDocument(string markdown, MarkdownParseOptions? options = null) =>
        new RichMarkdownParser(options ?? MarkdownParseOptions.Default).Parse(markdown);

    /// <summary>
    /// Parses the given markdown into a <see cref="MarkdownParseResult"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">When <paramref name="markdown"/> is null.</exception>
    /// <exception cref="MarkdownParseException">When Markdig fails to parse the source.</exception>
    public MarkdownParseResult Parse(string markdown)
    {
        if (markdown is null)
        {
            throw new ArgumentNullException(nameof(markdown));
        }

        DocxEditor.Core.Markdown.Model.MarkdownDocument document;
        try
        {
            var pipeline = _options.BuildPipeline();
            var ast = Markdig.Markdown.Parse(markdown, pipeline);
            var context = new ParseContext(_options, markdown);
            document = new DocxEditor.Core.Markdown.Model.MarkdownDocument
            {
                Blocks = context.ConvertBlocks(ast, depth: 0),
                Span = context.GetSpan(ast)
            };
            return new MarkdownParseResult(document, context.Diagnostics);
        }
        catch (MarkdownParseException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new MarkdownParseException("Failed to parse markdown source.", ex);
        }
    }

    /// <summary>
    /// Holds the per-parse state so a single parser instance can be reused safely
    /// and the recursion stays free of shared mutable fields.
    /// </summary>
    private sealed class ParseContext
    {
        private readonly MarkdownParseOptions _options;
        private readonly string _source;

        public ParseContext(MarkdownParseOptions options, string source)
        {
            _options = options;
            _source = source;
        }

        public List<MarkdownDiagnostic> Diagnostics { get; } = [];

        // ---- Block mapping -------------------------------------------------

        public IReadOnlyList<MarkdownBlock> ConvertBlocks(IEnumerable<Block> blocks, int depth)
        {
            var result = new List<MarkdownBlock>();
            foreach (var block in blocks)
            {
                var node = ConvertBlock(block, depth);
                if (node != null)
                {
                    result.Add(node);
                }
            }

            return result;
        }

        private MarkdownBlock? ConvertBlock(Block block, int depth)
        {
            if (depth > MaxNestingDepth)
            {
                // Deeply nested content is flattened to a verbatim node rather than recursing
                // further, so adversarial nesting can never overflow the conversion stack.
                ReportDepthOverflow(block);
                return new MarkdownUnknownBlock
                {
                    Kind = "nestingDepthOverflow",
                    Raw = RawOf(block),
                    Span = GetSpan(block)
                };
            }

            switch (block)
            {
                case YamlFrontMatterBlock yaml:
                    return new MarkdownYamlFrontMatter
                    {
                        Yaml = yaml.Lines.ToString(),
                        Span = GetSpan(yaml)
                    };

                case HeadingBlock heading:
                    return new MarkdownHeading
                    {
                        Level = heading.Level,
                        Inlines = ConvertInlines(heading.Inline, depth),
                        Span = GetSpan(heading)
                    };

                case ParagraphBlock paragraph:
                    return new MarkdownParagraph
                    {
                        Inlines = ConvertInlines(paragraph.Inline, depth),
                        Span = GetSpan(paragraph)
                    };

                case ListBlock list:
                    return ConvertList(list, depth);

                case QuoteBlock quote:
                    var alertKind = quote is AlertBlock alert ? alert.Kind.ToString().Trim() : null;
                    if (alertKind is { Length: > 0 })
                    {
                        Warn($"Alert block retained as a quote (kind '{alertKind}').", quote, nameof(MarkdownQuote));
                    }

                    return new MarkdownQuote
                    {
                        AlertKind = alertKind,
                        Blocks = ConvertBlocks(quote, depth + 1),
                        Span = GetSpan(quote)
                    };

                case MathBlock math:
                    Warn("Math block retained as a code block.", math, nameof(MarkdownCodeBlock));
                    return new MarkdownCodeBlock
                    {
                        Text = math.Lines.ToString(),
                        Language = "math",
                        Span = GetSpan(math)
                    };

                case FencedCodeBlock fenced:
                    return new MarkdownCodeBlock
                    {
                        Text = fenced.Lines.ToString(),
                        Language = string.IsNullOrEmpty(fenced.Info) ? null : fenced.Info,
                        Arguments = string.IsNullOrEmpty(fenced.Arguments) ? null : fenced.Arguments,
                        Span = GetSpan(fenced)
                    };

                case CodeBlock code:
                    return new MarkdownCodeBlock
                    {
                        Text = code.Lines.ToString(),
                        Span = GetSpan(code)
                    };

                case Markdig.Syntax.HtmlBlock html:
                    Warn("Raw HTML block retained as an explicit node.", html, nameof(MarkdownHtmlBlock));
                    return new MarkdownHtmlBlock
                    {
                        Text = html.Lines.ToString(),
                        Type = html.Type.ToString(),
                        Span = GetSpan(html)
                    };

                case Table table:
                    return ConvertTable(table, depth);

                case DefinitionList definitionList:
                    return ConvertDefinitionList(definitionList, depth);

                case FootnoteGroup footnoteGroup:
                    return ConvertFootnoteGroup(footnoteGroup, depth);

                case LinkReferenceDefinitionGroup referenceGroup:
                    return ConvertLinkReferenceDefinitions(referenceGroup);

                case ThematicBreakBlock thematic:
                    return new MarkdownThematicBreak
                    {
                        Character = thematic.ThematicChar,
                        Count = thematic.ThematicCharCount,
                        Span = GetSpan(thematic)
                    };

                case EmptyBlock:
                    return null;

                default:
                    Warn($"Unrecognized block '{block.GetType().Name}' retained verbatim.", block, nameof(MarkdownUnknownBlock));
                    return new MarkdownUnknownBlock
                    {
                        Kind = block.GetType().Name,
                        Raw = RawOf(block),
                        Span = GetSpan(block)
                    };
            }
        }

        private MarkdownList ConvertList(ListBlock list, int depth)
        {
            var items = new List<MarkdownListItem>();
            foreach (var child in list)
            {
                if (child is ListItemBlock item)
                {
                    items.Add(new MarkdownListItem
                    {
                        Order = item.Order,
                        Blocks = ConvertBlocks(item, depth + 1),
                        Span = GetSpan(item)
                    });
                }
            }

            return new MarkdownList
            {
                Ordered = list.IsOrdered,
                BulletType = list.BulletType == '\0' ? null : list.BulletType,
                OrderedStart = string.IsNullOrEmpty(list.OrderedStart) ? null : list.OrderedStart,
                IsLoose = list.IsLoose,
                Items = items,
                Span = GetSpan(list)
            };
        }

        private MarkdownTable ConvertTable(Table table, int depth)
        {
            var columns = table.ColumnDefinitions
                .Select(def => new MarkdownTableColumn
                {
                    Alignment = def.Alignment.HasValue ? (MarkdownTableAlignment)def.Alignment.Value : null,
                    Width = def.Width
                })
                .ToList();

            var rows = new List<MarkdownTableRow>();
            foreach (var child in table)
            {
                if (child is TableRow row)
                {
                    var cells = new List<MarkdownTableCell>();
                    foreach (var rowChild in row)
                    {
                        if (rowChild is TableCell cell)
                        {
                            cells.Add(new MarkdownTableCell
                            {
                                ColumnIndex = cell.ColumnIndex,
                                ColumnSpan = cell.ColumnSpan,
                                RowSpan = cell.RowSpan,
                                Blocks = ConvertBlocks(cell, depth + 1)
                            });
                        }
                    }

                    rows.Add(new MarkdownTableRow
                    {
                        IsHeader = row.IsHeader,
                        Cells = cells
                    });
                }
            }

            return new MarkdownTable
            {
                Columns = columns,
                Rows = rows,
                Span = GetSpan(table)
            };
        }

        private MarkdownDefinitionList ConvertDefinitionList(DefinitionList definitionList, int depth)
        {
            var items = new List<MarkdownDefinitionItem>();
            foreach (var child in definitionList)
            {
                if (child is DefinitionItem item)
                {
                    var terms = new List<MarkdownDefinitionTerm>();
                    var definitions = new List<MarkdownBlock>();
                    foreach (var itemChild in item)
                    {
                        if (itemChild is DefinitionTerm term)
                        {
                            terms.Add(new MarkdownDefinitionTerm
                            {
                                Inlines = ConvertInlines(term.Inline, depth),
                                Span = GetSpan(term)
                            });
                        }
                        else
                        {
                            var definition = ConvertBlock(itemChild, depth + 1);
                            if (definition != null)
                            {
                                definitions.Add(definition);
                            }
                        }
                    }

                    items.Add(new MarkdownDefinitionItem
                    {
                        Terms = terms,
                        Definitions = definitions,
                        Span = GetSpan(item)
                    });
                }
            }

            return new MarkdownDefinitionList
            {
                Items = items,
                Span = GetSpan(definitionList)
            };
        }

        private MarkdownFootnotesBlock ConvertFootnoteGroup(FootnoteGroup group, int depth)
        {
            var footnotes = new List<MarkdownFootnote>();
            foreach (var child in group)
            {
                if (child is Footnote footnote)
                {
                    footnotes.Add(new MarkdownFootnote
                    {
                        Label = footnote.Label ?? string.Empty,
                        Order = footnote.Order,
                        Blocks = ConvertBlocks(footnote, depth + 1),
                        Span = GetSpan(footnote)
                    });
                }
            }

            return new MarkdownFootnotesBlock
            {
                Footnotes = footnotes,
                Span = GetSpan(group)
            };
        }

        private MarkdownLinkReferenceDefinitions? ConvertLinkReferenceDefinitions(LinkReferenceDefinitionGroup group)
        {
            var definitions = new List<MarkdownLinkReferenceDefinition>();
            foreach (var child in group)
            {
                if (child is LinkReferenceDefinition definition
                    && definition is not FootnoteLinkReferenceDefinition
                    && definition is not HeadingLinkReferenceDefinition)
                {
                    definitions.Add(new MarkdownLinkReferenceDefinition
                    {
                        Label = definition.Label ?? string.Empty,
                        Url = definition.Url ?? string.Empty,
                        Title = definition.Title,
                        Span = GetSpan(definition)
                    });
                }
            }

            if (definitions.Count == 0)
            {
                return null;
            }

            return new MarkdownLinkReferenceDefinitions
            {
                Definitions = definitions,
                Span = GetSpan(group)
            };
        }

        // ---- Inline mapping -------------------------------------------------

        private IReadOnlyList<MarkdownInline> ConvertInlines(ContainerInline? root, int depth)
        {
            var result = new List<MarkdownInline>();
            if (root == null)
            {
                return result;
            }

            foreach (var inline in root)
            {
                if (inline is DelimiterInline delimiter)
                {
                    Warn($"Unmatched delimiter '{delimiter.GetType().Name}' flattened to its literal content.", delimiter, nameof(MarkdownUnknownInline));
                    result.AddRange(ConvertInlines(delimiter, depth + 1));
                    continue;
                }

                var node = ConvertInline(inline, depth);
                if (node != null)
                {
                    result.Add(node);
                }
            }

            return result;
        }

        private MarkdownInline? ConvertInline(Inline inline, int depth)
        {
            if (depth > MaxNestingDepth)
            {
                // Deeply nested inline content is flattened to a verbatim node rather than
                // recursing further, so adversarial nesting can never overflow the conversion stack.
                ReportDepthOverflow(inline);
                return new MarkdownUnknownInline
                {
                    Kind = "nestingDepthOverflow",
                    Raw = RawOf(inline),
                    Span = GetSpan(inline)
                };
            }

            switch (inline)
            {
                case EmojiInline emoji:
                    return new MarkdownEmoji
                    {
                        Match = emoji.Match ?? string.Empty,
                        Text = emoji.Content.ToString(),
                        Span = GetSpan(emoji)
                    };

                case LiteralInline literal:
                    return new MarkdownText
                    {
                        Text = literal.Content.ToString(),
                        IsEscaped = literal.IsFirstCharacterEscaped,
                        Span = GetSpan(literal)
                    };

                case CustomContainerInline:
                    Warn("Inline custom container retained verbatim.", inline, nameof(MarkdownUnknownInline));
                    return new MarkdownUnknownInline
                    {
                        Kind = inline.GetType().Name,
                        Raw = RawOf(inline),
                        Span = GetSpan(inline)
                    };

                case EmphasisInline emphasis:
                    return new MarkdownEmphasis
                    {
                        Kind = MapEmphasisKind(emphasis),
                        Children = ConvertInlines(emphasis, depth + 1),
                        Span = GetSpan(emphasis)
                    };

                case CodeInline code:
                    return new MarkdownCode
                    {
                        Content = code.Content,
                        Delimiter = code.Delimiter,
                        DelimiterCount = code.DelimiterCount,
                        Span = GetSpan(code)
                    };

                case LinkInline link:
                    var children = ConvertInlines(link, depth + 1);
                    if (link.IsImage)
                    {
                        return new MarkdownImage
                        {
                            Url = link.Url,
                            Title = link.Title,
                            Children = children,
                            Span = GetSpan(link)
                        };
                    }

                    return new MarkdownLink
                    {
                        Url = link.Url,
                        Title = link.Title,
                        IsAutoLink = link.IsAutoLink,
                        IsShortcut = link.IsShortcut,
                        Children = children,
                        Span = GetSpan(link)
                    };

                case AutolinkInline autolink:
                    return new MarkdownLink
                    {
                        Url = autolink.Url,
                        IsAutoLink = true,
                        Children = [new MarkdownText { Text = autolink.Url }],
                        Span = GetSpan(autolink)
                    };

                case HtmlInline html:
                    Warn("Raw inline HTML retained as an explicit node.", html, nameof(MarkdownHtml));
                    return new MarkdownHtml
                    {
                        Tag = html.Tag,
                        Span = GetSpan(html)
                    };

                case HtmlEntityInline entity:
                    return new MarkdownEntity
                    {
                        Original = entity.Original.ToString(),
                        Decoded = entity.Transcoded.ToString(),
                        Span = GetSpan(entity)
                    };

                case LineBreakInline lineBreak:
                    return new MarkdownLineBreak
                    {
                        IsHard = lineBreak.IsHard,
                        IsBackslash = lineBreak.IsBackslash,
                        Span = GetSpan(lineBreak)
                    };

                case FootnoteLink footnoteLink:
                    return new MarkdownFootnoteReference
                    {
                        Label = footnoteLink.Footnote?.Label ?? string.Empty,
                        Index = footnoteLink.Index,
                        IsBackLink = footnoteLink.IsBackLink,
                        Span = GetSpan(footnoteLink)
                    };

                case TaskList taskList:
                    return new MarkdownTaskCheckbox
                    {
                        Checked = taskList.Checked,
                        Span = GetSpan(taskList)
                    };

                case AbbreviationInline:
                    Warn("Abbreviation retained verbatim.", inline, nameof(MarkdownUnknownInline));
                    return new MarkdownUnknownInline
                    {
                        Kind = inline.GetType().Name,
                        Raw = RawOf(inline),
                        Span = GetSpan(inline)
                    };

                case MathInline math:
                    Warn("Inline math retained verbatim.", math, nameof(MarkdownUnknownInline));
                    return new MarkdownUnknownInline
                    {
                        Kind = inline.GetType().Name,
                        Raw = RawOf(math),
                        Span = GetSpan(math)
                    };

                default:
                    Warn($"Unrecognized inline '{inline.GetType().Name}' retained verbatim.", inline, nameof(MarkdownUnknownInline));
                    return new MarkdownUnknownInline
                    {
                        Kind = inline.GetType().Name,
                        Raw = RawOf(inline),
                        Span = GetSpan(inline)
                    };
            }
        }

        // ---- Helpers --------------------------------------------------------

        /// <summary>
        /// Reports a node that was flattened because its nesting exceeded
        /// <see cref="MaxNestingDepth"/>. Unlike <see cref="Warn"/> this diagnostic is emitted in
        /// every mode: it signals user-visible data loss (content escaped to text), not a
        /// retained-verbatim construct, so it must never be silently dropped.
        /// </summary>
        private void ReportDepthOverflow(Markdig.Syntax.MarkdownObject node)
        {
            Diagnostics.Add(new MarkdownDiagnostic(
                MarkdownDiagnosticSeverity.Warning,
                $"Markdown nesting exceeds the maximum supported depth of {MaxNestingDepth}; the node was flattened to escaped text.",
                GetSpan(node),
                "nestingDepthOverflow"));
        }

        private static EmphasisKind MapEmphasisKind(EmphasisInline emphasis) =>
            emphasis.DelimiterChar switch
            {
                '*' or '_' => emphasis.DelimiterCount >= 2 ? EmphasisKind.Bold : EmphasisKind.Italic,
                '~' => emphasis.DelimiterCount >= 2 ? EmphasisKind.Strikethrough : EmphasisKind.Subscript,
                '^' => EmphasisKind.Superscript,
                '+' => EmphasisKind.Inserted,
                '=' => EmphasisKind.Marked,
                _ => emphasis.DelimiterCount >= 2 ? EmphasisKind.Bold : EmphasisKind.Italic
            };

        public DocxEditor.Core.Markdown.Model.SourceSpan? GetSpan(Markdig.Syntax.MarkdownObject node)
        {
            if (!_options.UseSourceSpans)
            {
                return null;
            }

            var span = node.Span;
            if (span.Start < 0 || span.End < span.Start)
            {
                return null;
            }

            return new DocxEditor.Core.Markdown.Model.SourceSpan(span.Start, span.End, node.Line);
        }

        private string RawOf(Markdig.Syntax.MarkdownObject node)
        {
            if (node is LeafBlock leaf && leaf.Lines.Count > 0)
            {
                return leaf.Lines.ToString();
            }

            var span = node.Span;
            if (span.Start >= 0 && span.End >= span.Start && span.End < _source.Length)
            {
                return _source.Substring(span.Start, span.End - span.Start + 1);
            }

            return node.ToString() ?? node.GetType().Name;
        }

        private void Warn(string message, Markdig.Syntax.MarkdownObject node, string nodeKind)
        {
            if (!_options.Strict)
            {
                return;
            }

            Diagnostics.Add(new MarkdownDiagnostic(
                MarkdownDiagnosticSeverity.Warning,
                message,
                GetSpan(node),
                nodeKind));
        }
    }
}
