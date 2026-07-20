#!/usr/bin/env python
"""Generate the committed template `xhtml2docx/templates/reference.docx` from a seed archive.

The seed (`_data/empty.docx`, a fresh empty document saved by Word 16.111) supplies theme, fonts,
settings, and Word's own modern definitions for the styles we keep; this script strips styles.xml
to exactly what STYLE_MAP needs, patches Quote for blockquote semantics (left indent, not Word's
centering), authors the definitions Word leaves latent, scrubs personal metadata, and self-verifies:
fast_checks == 'valid' and every STYLE_MAP name defined. See meta/STATUS.md, template section."""
import zipfile
from lxml import etree
from xhtml2docx.styles import STYLE_MAP, style_id
from xhtml2docx.validate import fast_checks

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def w(tag): return f'{{{W}}}{tag}'

# Styles the seed defines that we keep as Word authored them (plus linked Char twins, which
# w:link references require). Everything else the seed defines is dropped.
KEEP = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList', 'Quote', 'QuoteChar', 'ListParagraph',
        *[f'Heading{n}' for n in range(1, 7)], *[f'Heading{n}Char' for n in range(1, 7)]}

# Styles Word keeps latent (definitions live inside Word, absent from the file), authored here.
# Built-in names are canonical (lowercase for heading/caption/footnote families); custom ones marked so.
NEW_STYLES = r'''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
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
        if e is not None: e.text = 'xhtml2docx'
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def build(seed='_data/empty.docx', out='xhtml2docx/templates/reference.docx'):
    z = zipfile.ZipFile(seed)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for i in z.infolist():
            data = z.read(i.filename)
            if i.filename == 'word/styles.xml': data = build_styles(data)
            elif i.filename == 'docProps/core.xml': data = scrub_props(data)
            zo.writestr(i.filename, data)
    verify(out)
    print(f'{out}: ok')

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

if __name__ == '__main__': build()
