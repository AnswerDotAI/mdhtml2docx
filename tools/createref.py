#!/usr/bin/env python
"""Generate the committed template `mdhtml2docx/templates/reference.docx` from a seed archive.

The seed (`_data/empty.docx`) is an empty document saved by Word for the web: it supplies theme,
fonts, settings, an empty header and footer slot, and Word's own definitions for the built-in
styles we keep. The seed supplies what Word's UI can materialize; this script authors the rest.
To reproduce the seed: in a blank Word (web) document, apply Heading 1-6, Quote, and Title each
to a throwaway line and make one line a bulleted list (Word then writes out their definitions),
delete all text and set the last paragraph back to Normal, insert an empty header and footer,
and download a copy.

Building strips styles.xml to exactly what STYLE_MAP needs, patches Quote for blockquote
semantics (left indent, not Word's centering), applies the house look (Times New Roman 11pt,
1.5 spacing, justified prose, bold black headings, a centered Title, a page-number footer),
authors the definitions Word leaves latent, scrubs personal metadata, and self-verifies:
fast_checks == 'valid' and every STYLE_MAP name defined."""
import zipfile
from lxml import etree
from mdhtml2docx.styles import STYLE_MAP, style_id
from mdhtml2docx.validate import fast_checks
from mdhtml2docx.wml import E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def w(tag): return f'{{{W}}}{tag}'

# Styles the seed defines that we keep (restyled below where the house look differs from Word's).
KEEP = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList', 'Quote', 'ListParagraph', 'Title',
    'Header', 'Footer', *[f'Heading{n}' for n in range(1, 7)]}

font, size, line, space_after = 'Times New Roman', 11, 1.5, 11   # typeface, body points, line spacing multiple, space below paragraphs (points)

# Styles the seed cannot supply, authored here: our own custom styles, plus built-ins that
# Word keeps latent and the web UI has no path to materialize (caption, footnote text, ...).
# Built-in names are canonical (lowercase for heading/caption/footnote families); custom ones marked so.
NEW_STYLES = r'''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="BodyText">
  <w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="1"/><w:qFormat/>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FirstParagraph">
  <w:name w:val="First Paragraph"/><w:basedOn w:val="BodyText"/><w:next w:val="BodyText"/><w:uiPriority w:val="1"/><w:qFormat/>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Centered">
  <w:name w:val="Centered"/><w:basedOn w:val="BodyText"/><w:next w:val="BodyText"/><w:qFormat/>
  <w:pPr><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Compact">
  <w:name w:val="Compact"/><w:basedOn w:val="BodyText"/><w:uiPriority w:val="1"/><w:qFormat/>
  <w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
</w:style>
<w:style w:type="paragraph" w:styleId="SourceCode">
  <w:name w:val="Source Code"/><w:basedOn w:val="Normal"/><w:next w:val="FirstParagraph"/><w:uiPriority w:val="1"/><w:qFormat/>
  <w:pPr><w:keepLines/><w:spacing w:before="120" w:after="120"/>
    <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
</w:style>
<w:style w:type="character" w:styleId="VerbatimChar">
  <w:name w:val="Verbatim Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="1"/><w:qFormat/>
  <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
</w:style>
<w:style w:type="character" w:styleId="Hyperlink">
  <w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
  <w:rPr><w:color w:val="467886" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Caption">
  <w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="FirstParagraph"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/>
  <w:pPr><w:spacing w:after="200"/></w:pPr>
  <w:rPr><w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText">
  <w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
  <w:pPr><w:spacing w:after="0"/></w:pPr>
  <w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
</w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
  <w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
  <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="DefinitionTerm">
  <w:name w:val="Definition Term"/><w:basedOn w:val="Normal"/><w:next w:val="Definition"/><w:uiPriority w:val="1"/><w:qFormat/>
  <w:pPr><w:spacing w:after="0"/><w:keepNext/></w:pPr>
  <w:rPr><w:b/><w:bCs/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Definition">
  <w:name w:val="Definition"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="1"/><w:qFormat/>
  <w:pPr><w:ind w:left="480"/></w:pPr>
</w:style>
<w:style w:type="table" w:styleId="TableGrid">
  <w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/>
  <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
  <w:tblPr><w:tblBorders>
    <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  </w:tblBorders></w:tblPr>
</w:style>
<w:style w:type="table" w:styleId="BorderlessTable">
  <w:name w:val="Borderless Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="40"/>
  <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
</w:style>
</w:styles>'''

def _hp(points): return int(round(points * 2))
def _fonts(): return E('w:rFonts', {'w:ascii': font, 'w:hAnsi': font, 'w:eastAsia': font, 'w:cs': font})
def _spacing(): return E('w:spacing', {'w:line': int(240 * line), 'w:lineRule': 'auto', 'w:after': space_after * 20})
def _jc(val): return E('w:jc', {'w:val': val})

def _restyle(root, sid, ppr=None, rpr=None):
    "Replace style `sid`'s paragraph and run properties, dropping whatever it had"
    s = root.find(f'{w("style")}[@{w("styleId")}="{sid}"]')
    for t in ('pPr', 'rPr'):
        if (e := s.find(w(t))) is not None: s.remove(e)
    for e in (ppr, rpr):
        if e is not None: s.append(e)

def build_styles(xml):
    "Strip the seed's styles.xml to KEEP, patch Quote, apply the house look, append the authored definitions"
    root = etree.fromstring(xml)
    for s in list(root.iter(w('style'))):
        if s.get(w('styleId')) not in KEEP: root.remove(s)
    q = next(s for s in root.iter(w('style')) if s.get(w('styleId')) == 'Quote')
    qp = q.find(w('pPr'))
    qp.remove(qp.find(w('jc')))
    etree.SubElement(qp, w('ind')).set(w('left'), '720')
    for s in root.iter(w('style')):
        if s.get(w('styleId')) in ('Quote', 'Title', *[f'Heading{n}' for n in range(1, 7)]):
            s.find(w('next')).set(w('val'), 'FirstParagraph')   # typing after these continues our prose chain
    rpd = root.find(f'{w("docDefaults")}/{w("rPrDefault")}')
    rpd.replace(rpd.find(w('rPr')), E('w:rPr', _fonts(), E('w:sz', {'w:val': _hp(size)}),
        E('w:szCs', {'w:val': _hp(size)}), E('w:lang', {'w:val': 'en-US'})))
    _restyle(root, 'Normal', E('w:pPr', _spacing(), E('w:ind', {'w:firstLine': 0}), _jc('both')))
    for i in range(6):
        rpr = E('w:rPr', _fonts(), E('w:b'), E('w:color', {'w:val': 'auto'}), E('w:sz', {'w:val': _hp(size)}), E('w:szCs', {'w:val': _hp(size)}))
        _restyle(root, f'Heading{i + 1}', E('w:pPr', _spacing(), _jc('both'), E('w:outlineLvl', {'w:val': i})), rpr)
    _restyle(root, 'Title', E('w:pPr', _jc('center')), E('w:rPr', E('w:sz', {'w:val': _hp(14)}), E('w:szCs', {'w:val': _hp(14)})))
    for s in etree.fromstring(NEW_STYLES.encode()): root.append(s)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def footer_content(xml):
    "word/footer.xml: replace the seed's empty scaffold with a centered page-number field"
    root = etree.fromstring(xml)
    for e in list(root): root.remove(e)
    root.append(E('w:p', E('w:pPr', _jc('center')), E('w:fldSimple', {'w:instr': ' PAGE '}, E('w:r', E('w:t', '1')))))
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def doc_content(xml):
    "word/document.xml: move the sectPr's header/footerReference first, as the schema requires (web Word appends them last)"
    root = etree.fromstring(xml)
    sect = root.find(f"{w('body')}/{w('sectPr')}")
    refs = [e for e in sect if etree.QName(e).localname in ('headerReference', 'footerReference')]
    for i, e in enumerate(refs):
        sect.remove(e)
        sect.insert(i, e)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def scrub_props(xml):
    "Replace personal creator/lastModifiedBy in docProps/core.xml"
    root = etree.fromstring(xml)
    for tag in ('{http://purl.org/dc/elements/1.1/}creator',
        '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy'):
        e = root.find(tag)
        if e is not None: e.text = 'mdhtml2docx'
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def build(seed='_data/empty.docx', out='mdhtml2docx/templates/reference.docx'):
    z = zipfile.ZipFile(seed)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for i in z.infolist():
            data = z.read(i.filename)
            if i.filename == 'word/styles.xml': data = build_styles(data)
            elif i.filename == 'word/footer.xml': data = footer_content(data)
            elif i.filename == 'word/document.xml': data = doc_content(data)
            elif i.filename == 'docProps/core.xml': data = scrub_props(data)
            zo.writestr(i.filename, data)
    verify(out)
    print(f'{out}: ok')

def verify(path):
    "The template must pass fast_checks, define (not leave latent) every STYLE_MAP style, and carry the page-number footer"
    r = fast_checks(path)
    assert r == 'valid', r
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('word/styles.xml'))
    names = {s.find(w('name')).get(w('val')) for s in root.iter(w('style'))}
    missing = set(STYLE_MAP.values()) - names
    assert not missing, f'STYLE_MAP styles not defined: {missing}'
    ids = {s.get(w('styleId')) for s in root.iter(w('style'))}
    badid = {n for n in STYLE_MAP.values() if style_id(n) not in ids}
    assert not badid, f'style_id mismatch for: {badid}'
    assert b'PAGE' in z.read('word/footer.xml'), 'footer lacks its page-number field'
    assert z.read('word/document.xml').decode().count('footerReference') == 1, 'expected exactly one footerReference'

if __name__ == '__main__': build()
