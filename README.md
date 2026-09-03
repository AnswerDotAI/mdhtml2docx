# mdhtml2docx

Convert [MDHTML](https://github.com/AnswerDotAI/mdhtml) to Word docx files.

`mdhtml` renders Markdown to an HTML5 document format shared by format-specific exporters. This package converts its portable core to docx from scratch. It uses [fast5ever](https://github.com/AnswerDotAI/fast5ever)'s mutable WHATWG DOM for input and lxml for the generated WordprocessingML.

MDHTML accepts the full HTML vocabulary. This exporter supports the portable elements and annotations listed below. It preserves the text of unknown inline elements, recurses through block containers, and warns when an unsupported block must become a plain paragraph. HTML parsing and repair belong to `mdhtml`; this package only walks the resulting tree.

## Usage

```python
from mdhtml import md2mdhtml
from mdhtml2docx import mdhtml2docx

html = md2mdhtml(markdown_text)
warnings = mdhtml2docx(html, 'out.docx')
```

`mdhtml2docx` takes an MDHTML string, writes a docx, and returns a list of warning strings. It parses strings with `mdhtml.mdhtml2dom`, so normal HTML5 repair applies and XML well-formedness is irrelevant. It also accepts an existing mutable fast5ever DOM:

```python
from mdhtml import mdhtml2dom

document = mdhtml2dom(html)
document.children[0].attrs['custom-style'] = 'Contract Title'
warnings = mdhtml2docx(document, 'out.docx')
```

Body-position text and phrasing content become implicit Word paragraphs. Whitespace between blocks remains inert. HTML templates remain inert and are not rendered.

Supported: headings, paragraphs, bold, italic, strikethrough, underline (the `u` element), highlight, superscript and subscript, inline code, links (external and internal, via bookmarks), images (local files are embedded with correct dimensions; remote URLs degrade to links), block quotes, fenced and indented code blocks, bullet and numbered lists (including nesting, `start`, and task-list checkboxes), pipe and grid tables (including row and column spans and header rows), definition lists, footnotes, horizontal rules, and abbreviations. A `details` div (the dialect's collapsible block) degrades to its label as a bold line above the body, never a numbered heading — Word has no folding. Math spans and blocks become native Word math zones (`m:oMath`) holding the source text as-is, whatever the dialect (TeX, UnicodeMath, AsciiMath). Build-up/rendering is left downstream, and pandoc reads the zones back as math.

Code blocks with a language are syntax colored when [fastpylight](https://github.com/AnswerDotAI/fastpylight) is installed: tokens get `Hl*` character styles (e.g. `Hl Keyword`), whose colors come from the theme entry in the reference list below, so a finished document can be restyled in Word like everything else.

## Styling

The generated document uses named styles, never inline formatting, so appearance is controlled by restyling. A markdown h1 is the document title: it gets Word's Title style and shows no number, and h2 through h6 map to heading 1 through 5. The title does carry the numbering's invisible level 0, so every h1 restarts the count below it: a file holding several documents, each opening with an h1, numbers each of them from 1. Prose paragraphs get Body Text (First Paragraph directly after a heading or similar block, following pandoc's convention), and the other styles are the ones you would expect: Quote, Source Code, Verbatim Char, Hyperlink, List Paragraph, Compact (table cells), Definition Term, Definition, caption, footnote styles, and Table Grid (plus the author-selectable Borderless Table).

Pass `reference='mydoc.docx'` to use your own document's styles instead of the built-in template, exactly like pandoc's `--reference-doc`. `reference` may also be a list: the first entry supplies the document (page setup, fonts, and all base styles), and each later entry contributes just its styles, replacing same-named earlier ones - either another `.docx`, or a fastpylight theme name such as `'dracula'`, which generates the code-color styles on the fly. The default is the built-in template plus `'github_light'`; pass a bare reference for plain uncolored code, or `mdhtml2docx.styles.theme_ref('dracula', 'dracula.docx')` to write a theme's styles as a standalone docx you can inspect or tweak. A `custom-style="Name"` attribute (from `{custom-style="Name"}` in Markdown) applies that style from your reference doc; if the style is missing, a stub is injected and a warning returned. A plain class like `{.note}` applies a style only when your reference doc defines one named `note`, and is otherwise ignored. Both work on tables too: a table whose `custom-style` or class names a table style in the reference doc uses it in place of Table Grid - the built-in template ships `Borderless Table` (no gridlines, for signature blocks and other layout tables).

Tables can mix fixed and proportional column widths with a `colwidths` attribute, written in Markdown as an attribute list after the table: `{: colwidths="10em 2fr 1fr"}`. Lengths fix a column; `fr` values share the remaining width, as in CSS grid.

The built-in template is generated by `tools/createref.py` from a stock Word document; it defines exactly the styles the converter emits, plus next-paragraph chains so documents stay well-styled while edited by hand in Word.

## Raw docx

This converter consumes MDHTML raw data whose `data-format` is `docx`:

```html
<script type="application/vnd.mdhtml.raw" data-format="docx">…</script>
```

A ```` ```{=docx} ```` fenced block in Markdown, or inline code followed by `{=docx}`, produces that carrier. The payload is parsed as WordprocessingML and inserted verbatim. Block payloads supply content such as `w:p` or `w:tbl`; inline payloads supply content such as `w:r`. The prefixes `w`, `r`, `wp`, `a`, `pic`, and `m` are predeclared, so no namespace boilerplate is needed.

Literal payloads have no `data-encoding`. The exporter performs one character-reference decoding pass for `data-encoding="html"`, and accepts UTF-8 payloads encoded with `data-encoding="base64"`. Malformed payloads and unknown encodings are dropped with a warning. Raw data for other formats is skipped silently.

The canonical example is a page break:

    ```{=docx}
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    ```

## Cross-references

Markdown references like `[@sec-payment]` become MDHTML `a` elements marked with `data-ref`, which this exporter turns into live Word REF fields rather than baked text. The default field is `REF <bookmark> \w \h`, the full-context paragraph number ("3.(c)(iii)") as a hyperlink; `settings.xml` gets `updateFields`, so Word refreshes on open. The `leaf`, `rel`, `text`, and `page` tokens select other fields, while the independent `bare` token suppresses the prefix word. Unknown or conflicting tokens are conversion errors.

The word before the number comes from the reference's type, the id up to its first `-`: `sec` maps to Section/Sections out of the box, and `reftypes=dict(exh=('Exhibit', 'Exhibits'))` adds more. `[Clause @sec-x]` overrides the word for one reference; `[-@sec-x]` suppresses it. Grouped references use a `span` marked with `data-refs` and join as "Sections 3.1 and 4.2" with one field per number. They are never collapsed into ranges, because "3.1-3.3" is static text whose meaning silently changes when a clause is inserted. A reference whose target id does not exist, or whose type has no prefix defined when one is needed, raises rather than warning: a lawyer's document must not open showing "Error! Reference source not found."

Number fields need numbered headings. If your reference docx already numbers its heading styles (most firm templates do), nothing more is required, and the converter leaves that numbering alone. Otherwise pass `number_headings='legal'` for 1. / (a) / (i) numbering or `'decimal'` for 1. / 1.1. / 1.1.1. (the names index mdhtml's `SCHEMES`), or pass your own scheme as a `{lvlText: numFmt}` dict, one entry per heading level from h1 down, the same shape as mdhtml's: level 0 is the h1 title with an empty lvlText, and `%2` is the h2 counter. A reference-list entry ending in `.xml` is a third route: a raw file of `w:style`, `w:abstractNum`, and `w:num` elements contributing styles and numbering together, with ids remapped to avoid collisions.

Figures and captioned tables number themselves with SEQ fields: a figure renders as its image plus a "Figure 1: caption" paragraph below (caption style), a table caption as "Table 1: caption" above the table, both live. When the element has an id, the label-and-number span is bookmarked, so `[@fig-plot]` inserts a live "Figure 1" (no extra prefix word; the label is part of the bookmarked text) and `[-@tbl-stages]` the bare number via a second number-only bookmark. `fig` and `tbl` are built-in reftypes alongside `sec`, and their label words come from the same table. Mixed-type groups render each item with its own singular prefix ("Figure 1 and Table 2"); same-type groups pluralize once. Reference targets must be things that get bookmarks - headings, paragraphs, figures, and tables with ids - and a ref to anything else is an error at conversion time.

## Validation

Every generated file is checked three ways in the test suite: lxml validation against the ECMA-376 schemas (container, CRC, XML, XSD), a semantic round trip through pandoc's independent docx reader, and occasional acceptance runs in Microsoft Word itself, driven live via AppleScript.
