"Fast validity checks for docx: container, per-entry CRC, XML parse, then ECMA-376 schema validation after MCE stripping."
import zipfile
from functools import cache
from importlib.resources import files
from lxml import etree

__all__ = ['wml_schema', 'mce_strip', 'fast_checks']

MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

@cache
def wml_schema():
    "Compiled transitional WordprocessingML schema (see schemas/SOURCES.md); ~44ms, hence the cache"
    return etree.XMLSchema(etree.parse(str(files('xhtml2docx')/'schemas'/'wml.xsd')))

def mce_strip(root):
    "Minimal MCE consumer processing, in place: drop attributes and elements in `mc:Ignorable` namespaces, and mc:* attributes"
    ign = set()
    for e in root.iter():
        if (v := e.get(f'{{{MC}}}Ignorable')): ign |= {e.nsmap[p] for p in v.split() if p in e.nsmap}
    bad = lambda ns: ns in ign or ns == MC
    for e in root.iter():
        for a in list(e.attrib):
            if a.startswith('{') and bad(a[1:a.index('}')]): del e.attrib[a]
    for e in [e for e in root.iter() if isinstance(e.tag, str) and bad(etree.QName(e).namespace or '')]:
        e.getparent().remove(e)
    return root

def fast_checks(path):
    "Cheapest-first verdict on the docx at `path`: 'valid', or the first failure as 'container|crc|xml|schema: detail'"
    try: z = zipfile.ZipFile(path)
    except Exception as e: return f'container: {e}'
    try:
        if (b := z.testzip()): return f'crc: bad entry {b}'
    except Exception as e: return f'crc: {e}'
    try: doc = mce_strip(etree.fromstring(z.read('word/document.xml')))
    except Exception as e: return f'xml: {e}'
    s = wml_schema()
    return 'valid' if s.validate(doc) else f'schema: {s.error_log.filter_from_errors()[0]}'
