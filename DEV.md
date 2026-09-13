# Development

## Input and output trees

`mdhtml2docx` accepts an MDHTML string or an existing fast5ever node. Strings are parsed once by `mdhtml.mdhtml2dom`. Converter code walks fast5ever's element, text, and comment nodes in document order; it does not project them into an XML tree.

Construction functions use `oxml.e` for WordprocessingML and configured `oxml.E` factories for math and custom namespaces. Factories create detached `oxml.XML` expressions backed by fastcore's namespace-aware builder. Completed blocks are live `oxml.Element` objects. Table cells end with paragraphs, and footnote blocks already contain their reference marks. Final document and footnote assembly serializes those blocks without reparsing or repairing whole parts. Parsed raw payloads and reference elements compose into expressions as namespace-preserving snapshots.

Reference styles, numbering, and settings stay in the package's cached live trees; `Document.save()` flushes their edits. Only generated parts are replaced with new bytes. `oxml.Document` owns content types, relationships, and atomic saving. Keeping the HTML input and XML output types distinct prevents XML rules from leaking into MDHTML handling.

Construction uses keyword attributes: `e.tcW(type='dxa', w=width)` qualifies both names with `w`. Cross-prefix attributes use `r__id` or `xml__space`; copied literal attribute mappings use `attrs_`. Namespace bindings belong to factories, not parent expressions. Run properties are built incrementally by calling a detached expression, such as `props(e.b(), e.bCs())`; an empty properties container is omitted. Calling a live parent, such as `root(e.num(...))`, attaches an expression in schema order and returns its live child. Word-specific content rules remain in the exporter: table cells end with paragraphs and footnotes receive reference marks. Explicit text positions use `parent(expression, index=n)`, counting all XML child nodes, including comments and whitespace. No lxml-compatible tree facade is used.

Generated document and footnote roots explicitly declare the `NS` vocabulary through `_partxml`. Pandoc drops images when their drawing namespaces are declared only on the nested drawing, even though that XML is namespace-valid. Preserve these root declarations when changing construction.

Resolve existing styles, numbering, settings, and footnotes through their relationships. Add internal targets before their relationships and use returned IDs, scoped to the main document or footnotes. New part names must avoid reference names case-insensitively. Style/numbering ID remapping, layout, fields, and all conversion policy remain here rather than in `oxml`. Theme reference generation uses the same package API.

Load the template's numbering IDs before processing contributors. Each XML contributor allocates fresh IDs and remaps its references before merging; contributors can reuse original IDs without cross-wiring styles. Reserve generated list and heading IDs after the contributors. Insert abstract definitions before numbering instances, and instances before `numIdMacAtCleanup`.

Bound controls retain the exporter's fixed datastore ID for reproducible builds. The converter supplies the field inventory to `Document.set_custom_xml()`; `oxml` owns the datastore properties, part names, and relationships. When a generated reference already contains that datastore, replace its field inventory rather than creating a second store with the same ID; unrelated custom XML stores are preserved.

Body-level phrasing runs are grouped into implicit Word paragraphs according to the MDHTML dialect. Explicit and implicit prose share first-paragraph styling and quote indentation; definition descriptions use the same block walker. Field construction also marks the document for field updates on open. HTML template contents remain outside the ordinary child sequence and are inert. Raw docx scripts are decoded at the point where they are converted into WordprocessingML.

## Tests

The `dev` extra includes `validation` (lxml) and pytest. The suite also uses Pandoc, fastcore, Pillow, and fastpylight from the development environment. Install the workspace after dependency or package-name changes, then run:

```bash
pytest -q
```

The default suite validates generated archives and reads them back through Pandoc. With the sibling `mdhtml` checkout available, include its full feature sample:

```bash
pytest -q -m 'not slow'
```

`tests/test_validate.py` checks the zip container, main-part XML, and its ECMA-376 schema independently of `oxml`, resolving the main part from the root relationship. Converter tests check numbering and footnote schemas, actual numbering bindings, renamed reference parts, relationship scopes, media collisions, and untouched reference content. The checkout sample also validates styles and settings. Keep these schema checks and Pandoc round trips: `oxml`'s partial validation does not replace their independent coverage. The committed reference document and schemas remain binary/XML assets; input MDHTML fixtures should use ordinary HTML serialization rather than XHTML spellings.

The no-lxml test checks only the dependency boundary: a minimal conversion in a fresh interpreter with lxml imports blocked. Feature behavior belongs in the normal converter tests, not a second suite embedded in a subprocess. `tools/createref.py` remains a development-only lxml user and requires the validation extra; it is not part of conversion.

Regenerate the bundled reference with `python tools/createref.py` using `_data/empty.docx` as the seed. The generator orders the seed's style children from the bundled XSD and validates the resulting styles as well as the main document. Authored and theme styles emit their properties in schema order directly; conversion does not repair reference styles.


## Table pagination

Table rows stay on one page by default. A row that does not fit in the remaining space moves to the next page. Tables can span pages. Rows taller than a page can still split.

Add `{: keep-rows=false}` directly below a Markdown table to allow its rows to split. `{: keep-rows=true}` explicitly selects the default. The converter writes Word's `cantSplit` setting on each row. The false override disables that setting even when a table style enables it.
