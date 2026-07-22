"Style names and the reference template. STYLE_MAP is the single source of truth: the template defines exactly these styles, the converter emits exactly these names, and tests close the loop."
from importlib.resources import files
from mdhtml.export import SCHEMES

__all__ = ['STYLE_MAP', 'style_id', 'ref_path', 'theme_styles', 'theme_ref']

STYLE_MAP = dict(
    body='Body Text', firstpara='First Paragraph', blockquote='Quote', codeblock='Source Code', codeinline='Verbatim Char',
    h1='heading 1', h2='heading 2', h3='heading 3', h4='heading 4', h5='heading 5', h6='heading 6',
    compact='Compact', hyperlink='Hyperlink', list='List Paragraph', dt='Definition Term', dd='Definition',
    caption='caption', footnotetext='footnote text', footnoteref='footnote reference', table='Table Grid')

def style_id(name):
    "Word style id for style `name`: title-cased, spaces dropped ('Body Text' -> 'BodyText', 'heading 1' -> 'Heading1')"
    return name.title().replace(' ', '')

def ref_path(): return files('mdhtml2docx')/'templates'/'reference.docx'



# Word underline values for fastpylight/Lumis underline styles
_ULINE = dict(solid='single', wavy='wave', double='double', dotted='dotted', dashed='dash')

def _hx(c): return c.lstrip('#').upper()

def theme_styles(theme):
    """w:style elements for fastpylight `theme`: one Hl* character style per scope (id 'HlKeywordFunction'
    for scope 'keyword.function', reversibly), plus a Source Code paragraph style carrying the theme's
    code-block background and default color (mirroring the template's, so later-wins merge replaces it)"""
    from fastpylight import theme_colors
    from .wml import E
    tc = theme_colors(theme)
    nrm = tc.pop('normal', {})
    def rpr(st):
        return E('w:rPr', E('w:b') if st['bold'] else None, E('w:i') if st['italic'] else None,
                 E('w:strike') if st['strikethrough'] else None,
                 E('w:color', {'w:val': _hx(st['fg'])}) if st['fg'] else None,
                 E('w:u', {'w:val': _ULINE[st['underline']]}) if st['underline'] else None,
                 E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': _hx(st['bg'])}) if st['bg'] else None)
    def sty(scope, st):
        name = ('hl ' + scope.replace('.', ' ')).title()
        return E('w:style', {'w:type': 'character', 'w:customStyle': 1, 'w:styleId': style_id(name)},
                 E('w:name', {'w:val': name}), E('w:basedOn', {'w:val': 'DefaultParagraphFont'}), rpr(st))
    mono = {'w:ascii': 'Consolas', 'w:hAnsi': 'Consolas', 'w:cs': 'Consolas'}
    sc = E('w:style', {'w:type': 'paragraph', 'w:styleId': 'SourceCode'},
           E('w:name', {'w:val': 'Source Code'}), E('w:basedOn', {'w:val': 'Normal'}),
           E('w:next', {'w:val': 'FirstParagraph'}), E('w:uiPriority', {'w:val': 1}), E('w:qFormat'),
           E('w:pPr', E('w:keepLines'), E('w:spacing', {'w:before': 120, 'w:after': 120}),
             E('w:shd', {'w:val': 'clear', 'w:color': 'auto', 'w:fill': _hx(nrm.get('bg') or '#F5F5F5')})),
           E('w:rPr', E('w:rFonts', mono),
             E('w:color', {'w:val': _hx(nrm['fg'])}) if nrm.get('fg') else None,
             E('w:sz', {'w:val': 20}), E('w:szCs', {'w:val': 20})))
    return [sc] + [sty(s, st) for s, st in sorted(tc.items())]

_CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
_DOCRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

def theme_ref(theme, dest):
    "Write a minimal, Word-openable reference docx at `dest` carrying only `theme`'s styles, for use as a later `reference` entry"
    import zipfile
    from lxml import etree
    from .wml import W, qn
    sroot = etree.Element(qn('w:styles'), nsmap={'w': W})
    for s in theme_styles(theme): sroot.append(s)
    doc = etree.Element(qn('w:document'), nsmap={'w': W})
    etree.SubElement(doc, qn('w:body')).append(etree.Element(qn('w:p')))
    with zipfile.ZipFile(dest, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', _CT)
        z.writestr('_rels/.rels', _RELS)
        z.writestr('word/_rels/document.xml.rels', _DOCRELS)
        z.writestr('word/document.xml', etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True))
        z.writestr('word/styles.xml', etree.tostring(sroot, xml_declaration=True, encoding='UTF-8', standalone=True))
    return dest
