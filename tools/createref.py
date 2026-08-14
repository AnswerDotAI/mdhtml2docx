#!/usr/bin/env python
"""Generate the committed template `mdhtml2docx/templates/reference.docx` from a seed archive.

The seed (`_data/empty.docx`, a fresh empty document saved by Word 16.111) supplies theme, fonts,
settings, and Word's own modern definitions for the styles we keep; this script strips styles.xml
to exactly what STYLE_MAP needs, patches Quote for blockquote semantics (left indent, not Word's
centering), authors the definitions Word leaves latent, scrubs personal metadata, applies the house look (see the house section below), and self-verifies:
fast_checks == 'valid' and every STYLE_MAP name defined. See meta/STATUS.md, template section.
Without the seed, running this script re-applies the house look to the committed template instead."""
import zipfile
from lxml import etree
from mdhtml2docx.styles import STYLE_MAP, style_id
from mdhtml2docx.validate import fast_checks
from mdhtml2docx.wml import E, R
from pathlib import Path

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def w(tag): return f'{{{W}}}{tag}'

# Styles the seed defines that we keep as Word authored them (plus linked Char twins, which
# w:link references require). Everything else the seed defines is dropped.
KEEP = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList', 'Quote', 'QuoteChar', 'ListParagraph',
    *[f'Heading{n}' for n in range(1, 7)], *[f'Heading{n}Char' for n in range(1, 7)]}

# Styles Word keeps latent (definitions live inside Word, absent from the file), authored here.
# Built-in names are canonical (lowercase for heading/caption/footnote families); custom ones marked so.
NEW_STYLES = r'''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Title">
  <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="FirstParagraph"/><w:uiPriority w:val="10"/><w:qFormat/>
  <w:pPr><w:jc w:val="center"/></w:pPr>
  <w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="BodyText">
  <w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="1"/><w:qFormat/>
</w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="FirstParagraph">
  <w:name w:val="First Paragraph"/><w:basedOn w:val="BodyText"/><w:next w:val="BodyText"/><w:uiPriority w:val="1"/><w:qFormat/>
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

def build_styles(xml):
    "Strip the seed's styles.xml to KEEP, fix Quote, append the authored definitions"
    root = etree.fromstring(xml)
    for s in list(root.iter(w('style'))):
        if s.get(w('styleId')) not in KEEP: root.remove(s)
    q = next(s for s in root.iter(w('style')) if s.get(w('styleId')) == 'Quote')
    qp = q.find(w('pPr'))
    qp.remove(qp.find(w('jc')))
    etree.SubElement(qp, w('ind')).set(w('left'), '720')
    for s in root.iter(w('style')):
        if s.get(w('styleId')) in ('Quote', *[f'Heading{n}' for n in range(1, 7)]):
            s.find(w('next')).set(w('val'), 'FirstParagraph')   # typing after these continues our prose chain
    for s in etree.fromstring(NEW_STYLES.encode()): root.append(s)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def scrub_props(xml):
    "Replace personal creator/lastModifiedBy in docProps/core.xml"
    root = etree.fromstring(xml)
    for tag in ('{http://purl.org/dc/elements/1.1/}creator',
        '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy'):
        e = root.find(tag)
        if e is not None: e.text = 'mdhtml2docx'
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---- house look ----------------------------------------------------------
# The default reference is opinionated: serif book typography with justified prose and a
# page-number footer. Callers wanting another look pass their own reference docx.

FONT, SIZE, LINE, SPACE_AFTER = 'Times New Roman', 11, 1.5, 11   # typeface, body points, line spacing multiple, space below paragraphs (points)

def _hp(points): return int(round(points * 2))
def _fonts(): return E('w:rFonts', {'w:ascii': FONT, 'w:hAnsi': FONT, 'w:eastAsia': FONT, 'w:cs': FONT})
def _spacing(): return E('w:spacing', {'w:line': int(240 * LINE), 'w:lineRule': 'auto', 'w:after': SPACE_AFTER * 20})
def _jc(val): return E('w:jc', {'w:val': val})

def _restyle(root, sid, ppr=None, rpr=None):
    "Replace style `sid`'s paragraph and run properties, dropping whatever it had"
    s = root.find(f'{w("style")}[@{w("styleId")}="{sid}"]')
    for t in ('pPr', 'rPr'):
        if (e := s.find(w(t))) is not None: s.remove(e)
    for e in (ppr, rpr):
        if e is not None: s.append(e)

def house_styles(xml):
    "House typography: `FONT` at `SIZE`pt for everything, justified prose, bold black headings at body size, plus the Centered block style"
    root = etree.fromstring(xml)
    rpd = root.find(f'{w("docDefaults")}/{w("rPrDefault")}')
    rpd.replace(rpd.find(w('rPr')), E('w:rPr', _fonts(), E('w:sz', {'w:val': _hp(SIZE)}),
        E('w:szCs', {'w:val': _hp(SIZE)}), E('w:lang', {'w:val': 'en-US'})))
    _restyle(root, 'Normal', E('w:pPr', _spacing(), E('w:ind', {'w:firstLine': 0}), _jc('both')))
    for i in range(6):
        rpr = E('w:rPr', _fonts(), E('w:b'), E('w:color', {'w:val': 'auto'}), E('w:sz', {'w:val': _hp(SIZE)}), E('w:szCs', {'w:val': _hp(SIZE)}))
        _restyle(root, f'Heading{i + 1}', E('w:pPr', _spacing(), _jc('both'), E('w:outlineLvl', {'w:val': i})), rpr)
    if root.find(f'{w("style")}[@{w("styleId")}="Centered"]') is None:
        root.append(E('w:style', {'w:type': 'paragraph', 'w:customStyle': 1, 'w:styleId': 'Centered'},
            E('w:name', {'w:val': 'Centered'}), E('w:basedOn', {'w:val': 'BodyText'}), E('w:next', {'w:val': 'BodyText'}),
            E('w:qFormat'), E('w:pPr', E('w:ind', {'w:firstLine': 0}), _jc('center'))))
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def footer_xml():
    "word/footer1.xml: a centered page-number field"
    root = etree.Element(w('ftr'), nsmap=dict(w=W))
    root.append(E('w:p', E('w:pPr', _jc('center')), E('w:fldSimple', {'w:instr': ' PAGE '}, E('w:r', E('w:t', '1')))))
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

CTNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
RNS = 'http://schemas.openxmlformats.org/package/2006/relationships'

def add_footer(parts):
    "Wire the page-number footer into archive `parts` (name -> bytes), unless one is already present"
    if 'word/footer1.xml' in parts: return
    parts['word/footer1.xml'] = footer_xml()
    ct = etree.fromstring(parts['[Content_Types].xml'])
    etree.SubElement(ct, f'{{{CTNS}}}Override', PartName='/word/footer1.xml',
        ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
    parts['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)
    rels = etree.fromstring(parts['word/_rels/document.xml.rels'])
    rid = f"rId{max((int(r.get('Id')[3:]) for r in rels if r.get('Id', '').startswith('rId')), default=0) + 1}"
    etree.SubElement(rels, f'{{{RNS}}}Relationship', Id=rid,
        Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', Target='footer1.xml')
    parts['word/_rels/document.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    droot = etree.fromstring(parts['word/document.xml'])
    sect = droot.find(f'{w("body")}/{w("sectPr")}')
    fr = E('w:footerReference', {'w:type': 'default'})
    fr.set(f'{{{R}}}id', rid)
    sect.insert(0, fr)
    parts['word/document.xml'] = etree.tostring(droot, xml_declaration=True, encoding='UTF-8', standalone=True)

def house(parts):
    "Apply the house look to reference `parts` in place"
    parts['word/styles.xml'] = house_styles(parts['word/styles.xml'])
    add_footer(parts)

def _parts(path):
    with zipfile.ZipFile(path) as z: return {i.filename: z.read(i.filename) for i in z.infolist()}

def _write(parts, out):
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for name, data in parts.items(): zo.writestr(name, data)

def build(seed='_data/empty.docx', out='mdhtml2docx/templates/reference.docx'):
    parts = _parts(seed)
    parts['word/styles.xml'] = build_styles(parts['word/styles.xml'])
    parts['docProps/core.xml'] = scrub_props(parts['docProps/core.xml'])
    house(parts)
    _write(parts, out)
    verify(out)
    print(f'{out}: ok')

def restyle(path='mdhtml2docx/templates/reference.docx'):
    "Re-apply the house look to the committed template, for machines without the seed"
    parts = _parts(path)
    house(parts)
    _write(parts, path)
    verify(path)
    print(f'{path}: restyled')

def verify(path):
    "The template must pass fast_checks and define (not leave latent) every STYLE_MAP style"
    r = fast_checks(path)
    assert r == 'valid', r
    root = etree.fromstring(zipfile.ZipFile(path).read('word/styles.xml'))
    names = {s.find(w('name')).get(w('val')) for s in root.iter(w('style'))}
    missing = set(STYLE_MAP.values()) - names
    assert not missing, f'STYLE_MAP styles not defined: {missing}'
    ids = {s.get(w('styleId')) for s in root.iter(w('style'))}
    badid = {n for n in STYLE_MAP.values() if style_id(n) not in ids}
    assert not badid, f'style_id mismatch for: {badid}'

if __name__ == '__main__': build() if Path('_data/empty.docx').exists() else restyle()
