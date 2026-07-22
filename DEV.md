# Development

## Input and output trees

`convert` accepts an MDHTML string or an existing JustHTML `DocumentFragment`. Strings are parsed once by `mdhtml.parse_mdhtml`. Converter code reads JustHTML `Element`, `Text`, and `Comment` nodes in document order; it does not project them into an XML tree.

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
