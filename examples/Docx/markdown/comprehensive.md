---
title: Markdown to DOCX — Comprehensive Feature Report
author: OfficeEditor
subject: Rich markdown renderer coverage fixture
keywords: markdown, DOCX, headings, lists, tables, footnotes, emoji
description: A self-contained test report exercising the rich markdown parser and renderer feature by feature, with expected DOCX behaviour and explicit fallback/limitation sections.
language: en-US
---

# Markdown to DOCX — Comprehensive Feature Report

This document is a **source fixture and test report** for the rich Markdown → DOCX
path (`RichMarkdownParser` + `RichMarkdownRenderer`, exercised through
`builder.AddRichMarkdown(...)` / `builder.AddMarkdown(...)`).

Every section below shows **the source**, then a short **expected DOCX behaviour** line.
A trailing legend marks each feature as **supported**, **fallback**, or **limitation**.

| Section | Feature area | Status |
|---|---|---|
| [Headings](#headings-atx-and-setext) | ATX levels 1–6 + setext `===` / `---` | Supported |
| [Inline emphasis](#inline-emphasis) | Nested bold / italic / strike / sub / sup / insert / mark | Supported |
| [Code](#code) | Inline, fenced (with language), indented | Supported |
| [Links](#links) | Inline, reference, shortcut, autolink | Supported |
| [Images](#images) | Local SVG reference, data URI | Supported (policy-bound) |
| [Blockquotes](#blockquotes) | Recursive / nested quotes | Supported |
| [Lists](#lists) | Bullet, ordered (non-1 start), nested, task | Supported |
| [Tables](#tables) | Pipe + grid, alignment, inline formatting | Supported |
| [Horizontal rules](#horizontal-rules) | `---`, `***` | Supported |
| [Line breaks](#line-breaks) | Hard and soft breaks | Supported |
| [Escapes & entities](#escapes-and-entities) | Backslash escapes, HTML entities | Supported |
| [Footnotes](#footnotes) | Footnote references + definitions | Supported |
| [Definition lists](#definition-lists) | Term + definitions | Supported |
| [Emoji](#emoji-and-shortcodes) | Unicode literals and `:shortcode:` | Supported |
| [YAML front matter](#yaml-front-matter) | Core-property mapping | Supported |
| [Generic attributes](#generic-attributes) | `{#id .class key=val}` | Accepted, not applied |
| [HTML](#html-fallback) | Raw HTML block / inline | Fallback (escaped text) |
| [Math](#math-fallback) | Block and inline math | Fallback (code / verbatim) |

The YAML block above is not rendered as content; it is mapped to the package core
properties (`title`, `author`, `subject`, `keywords`, `description`, `language`).

## Headings (ATX and setext)

Heading levels 1–6 render as Word heading paragraphs with matching outline levels.
In ATX form, a leading `#`–`######` selects the level. In setext form, an underline
of `=` produces level 1 and an underline of `-` produces level 2.

### Heading level 3

#### Heading level 4

##### Heading level 5

###### Heading level 6

Setext heading level 1
======================

Setext heading level 2
----------------------

> Expected: paragraphs above are paragraphs; `###### Heading level 6` is the deepest
> heading; the two setext lines become level-1 and level-2 headings respectively.

## Inline emphasis

Emphasis combines recursively: a bold span can contain italic, strike-through,
subscript, superscript, inserted, or marked text, and each combination is written as
run properties on the leaf text.

- **bold** and *italic* and ***bold italic***.
- ~~strike through~~, ~subscript~, ^superscript^, ++inserted++, ==marked==.
- Nested: **bold with *italic* and `code` inside**, ~~struck with **bold** inside~~.
- A single trailing pair keeps the nesting: *outer **inner** outer*.

> Expected: bold `<w:b/>`, italic `<w:i/>`, strike `<w:strike/>`, subscript and
> superscript `<w:vertAlign/>`, inserted underline `<w:u/>`, marked highlight
> `<w:highlight val="yellow"/>` on the respective runs.

## Code

Inline code is a direct-formatted run (monospace font `Consolas`, light shading).
Fenced blocks become code paragraphs carrying the `code`-mapped paragraph style; the
language tag is preserved in the parse but not used for syntax colouring. Indented
blocks (four leading spaces) behave like fenced blocks without an info string.

Inline: use `DocumentBuilder.AddRichMarkdown(...)` and `` `RichMarkdownRenderer` ``.

Fenced block with a language tag:

```csharp
var options = new MarkdownRenderOptions
{
    Strict = true,
    MaxDisplayWidthPt = 468
};
builder.AddRichMarkdown(markdown, options);
```

Indented code block (no info string):

    line one
    line two
      indented continuation

> Expected: inline code as a monospace shaded run; both block forms as `Code`-styled
> paragraphs with `xml:space="preserve"`.

## Links

Inline links to absolute URLs become real Word hyperlink relationships. Reference and
shortcut links resolve through the definition list. Bare and angle-bracket autolinks
(`https://…` and `<https://…>`) produce hyperlinks too.

- Inline: [OfficeEditor guide](https://opencode.ai)
- Reference: see the [feature report][r1] and the [shortcut link][r2].
- Shortcut: [r1] and [r2] on their own are resolved from the definitions below.
- Autolink: <https://example.org/auto> and plain https://example.org/plain.

[r1]: https://example.org/reference "Reference link title"
[r2]: https://example.org/shortcut

> Expected: every URL above renders as a clickable hyperlink run (colored, underlined)
> backed by a hyperlink relationship. Non-absolute links are not converted — see the
> fallback section.

## Images

Local image references are embedded when the source is a local file inside the allowed
root (or resolves from the working directory by default); `data:` URIs are embedded
directly. Remote HTTP(S) images are never fetched and fall back to a visible
`[image: alt]` text. The embedded SVG below is a small self-authored fixture asset in
the same directory.

![OfficeEditor badge](assets/officeeditor-badge.svg)

> Expected: the badge renders inline at its natural size, capped to the configured
> maximum display width (default 468 pt). The `alt` text is written to the drawing's
> description.

## Blockquotes

Quotes indent their content. Blockquotes nest recursively: a `>` inside a `>` deepens
the indentation, and content returns to the outer level after an inner block.

> Outer level quote.
>
> > Inner level quote.
> >
> > > Deepest level quote.
>
> Back at the outer level.

> Expected: three indentation depths using the `Quote`-mapped style, with the nested
> markers not reproduced as text.

## Lists

Ordered lists restart their counter at the declared start value. Bullet and ordered
lists each receive fresh, collision-free numbering; nested lists are indented to deeper
levels; task items render a checked/unchecked box glyph (`☑` / `☐`).

Ordered list starting at 3:

3. Third item
4. Fourth item
5. Fifth item

Unordered list:

- Alpha
- Beta
- Gamma

Nested list (bullets inside a bullet, then a numbered run):

- Outer item one
  - Inner bullet a
  - Inner bullet b
- Outer item two
  1. Inner numbered one
  2. Inner numbered two
- Outer item three

Task list:

- [x] Approved
- [ ] In review
- [ ] Pending sign-off

> Expected: the ordered list begins at **3**; nested items are indented; task items show
> `☑` and `☐` as run text; each list uses its own numbering instance so no counter
> leaks between lists.

## Tables

Pipe tables support column alignment and inline formatting inside cells. Header rows
are shaded and bolded. Grid tables (Markdig syntax) support the same rendering; the
alignment grid below shows right and center alignment in a grid table.

Pipe table with alignment and inline formatting:

| Metric | Value | Trend |
|:-------|:-----:|------:|
| Revenue | **1,240,000** | `+12%` |
| Margin | *34%* | Steady |
| Churn | 1.2 | [down](https://example.org/churn) |

Grid table (plain):

+-----------+----------+----------+
| Region    | Revenue  | Target   |
+===========+==========+==========+
| North     | 410K     | 400K     |
+-----------+----------+----------+
| South     | 380K     | 400K     |
+-----------+----------+----------+

Grid table with alignment (right / center / left):

+---+---+---+
| R | C | L |
+===+===+===+
| 1 | 2 | 3 |
+---+---+---+

> Expected: pipe table cells align left/center/right per column, the header row is
> shaded and bold, and inline emphasis renders inside cells. Both grid tables render
> with the same direct grid-bordered formatting; spans are not merged (see
> limitations).

## Horizontal rules

A line of `-`, `*`, or `_` between blank lines becomes a paragraph with a bottom
border (thematic break).

---

***

> Expected: two horizontal-rule paragraphs with a single-line bottom border.

## Line breaks

A hard break (two trailing spaces or a trailing backslash) becomes an explicit
`<w:br/>`. A soft break (single newline) renders as a space by default
(`MarkdownSoftBreakMode.Space`).

Hard break  
on the next line.
Backslash break\
also on the next line.
Soft break
stays on the same visual line as a space by default.

> Expected: two explicit line breaks and one soft space. Change
> `MarkdownRenderOptions.SoftBreakMode` to `LineBreak` or `None` to alter the soft
> break handling.

## Escapes and entities

Backslash escapes neutralise markdown syntax, and HTML entities decode to their
characters in the output.

Escapes: \*not emphasis\*, \# not a heading, \[not a link\](no), \1\2\3.

Entities: `&amp;` → &amp;, `&lt;code&gt;` → &lt;code&gt;, `&copy;` → &copy;, `&#x1F680;` → &#x1F680;.

> Expected: the escaped markers appear literally, and the entities decode to `&`, `<`,
> `>` markup-free characters and the 🚀 rocket glyph.

## Footnotes

Footnote references become superscript footnote markers wired to a `FootnotesPart`;
definitions at the end of the document are collected there and referenced back.

A claim worth a footnote[^one] and a second claim[^two].

[^one]: The first footnote body, with *formatting*.
[^two]: The second footnote body, `with code`.

> Expected: two footnote references and two footnote definitions in the generated
> package's footnotes part.

## Definition lists

Definition lists group a term with one or more indented definitions.

Term 1

:   Definition one for term 1.

Term 2

:   Definition two for term 2.

> Expected: each term renders as a bold paragraph on the `paragraph`-mapped style and
> each definition renders as an indented paragraph. (The `definitionTerm` /
> `definitionDescription` style-map keys are accepted but not consumed by the rich
> renderer.)

## Emoji and shortcodes

Unicode emoji pass through as-is; `:shortcode:` sequences expand to the matching
Unicode glyph (shortcode-only mapping, so pipe-table separators such as `:---` are not
misread as smileys).

Literals: 🚀 ✅ ⚠️
Shortcodes: :rocket: :white_check_mark: :warning:

> Expected: the literal glyphs and the three shortcodes render as their Unicode
> characters in the output text.

## YAML front matter

The `---` … `---` block at the top of this document is parsed as YAML front matter.
Recognised keys (`title`, `author`, `subject`, `keywords`, `description`, `language`)
are mapped to the package core properties and produce **no visible content**. Disable
the mapping with `MarkdownRenderOptions.MapYamlFrontMatterToCoreProperties = false` to
leave the block visible instead.

## Generic attributes

Generic attributes (`{#id .class key=value}`) are parsed by the markdown pipeline, but
the renderer does **not** apply them to OOXML — the syntax is tolerated and the block
still renders with its content.

## Generic attributes, applied heading {#custom-id .banner}

A paragraph that owns attributes.

{.lead key=value}

> Expected: the heading and paragraph render normally. The attribute tokens have no
> visible effect in DOCX (no bookmark/anchor or style is written for them). Attribute
> syntax mid-line can be mistaken for an HTML block by the parser, so prefer the
> block-trailing placement shown here.

## HTML — fallback

Raw HTML is **not** passed through. A raw HTML block is retained as an explicit node
and rendered as escaped visible text (with a diagnostic in strict mode), so the author
never loses the content and no script or markup reaches the package.

<div class="alert">
  This whole block is rendered as visible, escaped text — not as live HTML.
</div>

Inline <span class="x">HTML</span> is escaped into the surrounding paragraph.

> Expected: both the block and the inline `<span>` appear as literal text (e.g. showing
> the `<` and `>` characters), plus a `Warning` diagnostic on
> `MarkdownRenderOptions.Strict`.

## Math — fallback

Inline and block math are retained, never silently dropped, but they are **not** laid
out as equations. A block is kept as a code block labelled `math`; an inline span is
kept as verbatim text.

$$
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
$$

Inline math $E = mc^2$ stays as verbatim text in the paragraph.

> Expected: the `$$ … $$` block renders as a `Code`-styled paragraph (the parser warns
> that the math block is retained as a code block); the inline `$E = mc^2$` renders as
> escaped verbatim text in its paragraph. There is no equation layout.

## Other fallbacks

- **Non-absolute links** — `[relative](docs/guide)` is rendered as its plain text with a
  strict-mode warning, because only absolute URLs become hyperlink relationships.
- **Remote images** — `![logo](https://example.com/logo.png)` is never fetched; it falls
  back to a visible `[image: logo]` placeholder.
- **Missing local images** — an unreadable or missing local path also falls back to the
  visible `[image: alt]` placeholder instead of failing the conversion.
- **Unresolved footnote references** — a reference without a matching definition renders
  as literal `[label]` text.

## Limitations

- Table cell spans (`colspan` / `rowspan`) are parsed but rendered **without** merging;
  a strict-mode warning is emitted when a cell spans more than one column or row.
- Raw HTML, math, abbreviations, and custom-container blocks are not fully structured —
  they are retained as visible fallbacks per the sections above.
- Generic attributes are accepted but not applied.
- Images require a local or `data:` source; network fetching is disabled by design.
- List numbering is single-level per instance; multi-level schemes are not authored.
