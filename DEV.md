# Development

## Input and output trees

`mdhtml2docx` accepts an MDHTML string or an existing fast5ever node. Strings are parsed once by `mdhtml.mdhtml2dom`. Converter code walks fast5ever's element, text, and comment nodes in document order; it does not project them into an XML tree.

lxml remains the output library. It builds WordprocessingML parts, reads the reference archive's XML, and validates the resulting docx. Keeping the input and output types distinct prevents XML name, comment, namespace, and well-formedness rules from leaking into MDHTML handling.

Body-level phrasing runs are grouped into implicit Word paragraphs according to the MDHTML dialect. HTML template contents remain outside the ordinary child sequence and are inert. Raw docx scripts are decoded at the point where they are converted into WordprocessingML.

## Tests

Install the workspace after dependency or package-name changes, then run:

```bash
pytest -q
```

The default suite validates generated archives and reads them back through pandoc. Tests marked `slow` exercise live GUI applications or other long-running checks:

```bash
pytest -q -m slow
```

`tests/test_validate.py` checks the zip container, XML, relationships, and ECMA-376 schemas. The committed reference document and schemas remain binary/XML assets; input MDHTML fixtures should use ordinary HTML serialization rather than XHTML spellings.


## Table pagination

Table rows stay on one page by default. A row that does not fit in the remaining space moves to the next page. Tables can span pages. Rows taller than a page can still split.

Add `{: keep-rows=false}` directly below a Markdown table to allow its rows to split. `{: keep-rows=true}` explicitly selects the default. The converter writes Word's `cantSplit` setting on each row. The false override disables that setting even when a table style enables it.

## Included documents

A `div.include` with `scope="mic"` prefixes each local ID verbatim. For example, `sec-setup` becomes `mic:sec-setup`. Links within the include use the prefixed ID. A link outside the include can use that full ID. Scope names must be unique.

Each document starts with an h1 title. Word's existing heading numbering restarts after each h1. All included documents use the same numbering scheme selected for the export.

Grouped references retain their type across scopes. References to `mic:sec-setup` and `speaker:sec-setup` share the prefix "Sections". Paragraph IDs in lists remain available as Word bookmarks. Long bookmark names are shortened to fit Word's limit.
