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

A `div.include` with a unique `scope` attribute gives local reference names to an included document. IDs retain their type prefix. For example, `sec-notices` under scope `rsa.` becomes `sec-rsa.notices`. Links to targets inside the include follow the renamed IDs. External targets are unchanged.

A div's `number-headings` attribute selects `decimal`, `legal`, or `false` for its contents. Each numbered div owns a Word numbering instance. The enclosing numbering resumes after the div. Heading pagination remains controlled by the reference styles.

A `.keep-together` div keeps its consecutive paragraphs together. A page-break-only paragraph after a table is transferred to the following paragraph to avoid an empty page. List continuation paragraphs receive independent indentation properties.
