#!/usr/bin/env python
"""Generate the committed template `mdhtml2docx/templates/reference.docx` from a seed archive.

The seed (`_data/empty.docx`, an empty document saved by Word for the web, with header/footer
inserted and each kept style applied once so Word writes out its definitions) supplies theme,
fonts, settings, and Word's own definitions for the styles we keep; this script strips styles.xml
to exactly what STYLE_MAP needs, patches Quote for blockquote semantics (left indent, not Word's
centering), authors the definitions Word leaves latent, scrubs personal metadata, and
self-verifies: document validation and every STYLE_MAP name defined."""
from oxml import Document, Tree, e, w
from mdhtml2docx.styles import STYLE_MAP, style_id
from mdhtml2docx.wml import W

# Keep the seed's style definitions, correcting XML child order below; drop the rest.
KEEP = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'NoList', 'Quote', 'ListParagraph', 'Title',
    'Header', 'Footer', *[f'Heading{n}' for n in range(1, 7)]}

# Styles the seed cannot supply: our custom styles, plus built-ins the web UI cannot materialize.
# Built-in names are canonical (lowercase for heading/caption/footnote families); custom ones marked so.
NEW_STYLES = f'''<w:styles xmlns:w="{W}">
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
  <w:pPr><w:keepLines/><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
    <w:spacing w:before="120" w:after="120"/></w:pPr>
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
  <w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr>
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


def build_styles(tree):
    "Keep and schema-order the seed's styles, patch Quote, append the authored definitions."
    for style in list(tree.elements(w.Style)):
        if style.style_id not in KEEP: style.delete()
    for style in tree.elements(w.Style):
        style.reorder()
        if style.style_id in ('Quote', 'Title', *[f'Heading{n}' for n in range(1, 7)]): style.child(w.NextParagraphStyle).val = 'FirstParagraph'
    quote = next(s for s in tree.elements(w.Style) if s.style_id == 'Quote')
    props = quote.child(w.StyleParagraphProperties)
    props.child(w.Justification).delete()
    props(e.ind(left=720))
    for style in Tree(NEW_STYLES.encode()).root.children: tree.root(style)


def build(seed='_data/empty.docx', out='mdhtml2docx/templates/reference.docx'):
    doc = Document.open(seed)
    build_styles(doc.part('StyleDefinitionsPart').xml)
    sect, = doc.main.xml.elements(w.SectionProperties)
    sect.reorder()  # web Word appends header/footerReference last
    # Raw attributes: the typed `val` refuses Word for the web's 0/1 spellings, which CT_OnOffOnly does not allow.
    for name in ('HeaderPart', 'FooterPart'):
        for node in doc.part(name).xml.elements(w.BiDiVisual):
            val = node.attribute(W, 'val')
            if val in ('0', '1'): node.set_attribute(W, 'val', 'on' if val == '1' else 'off')
    doc.properties.update(creator='mdhtml2docx', lastModifiedBy='mdhtml2docx')
    doc.save(out)
    verify(out)
    print(f'{out}: ok')

def verify(path):
    "Validate the template, its STYLE_MAP definitions and the linked footer."
    doc = Document.open(path)
    assert not (issues := doc.validate()['issues']), issues
    styles = list(doc.styles)
    names = {style.name for style in styles}
    missing = set(STYLE_MAP.values()) - names
    assert not missing, f'STYLE_MAP styles not defined: {missing}'
    ids = {style.id for style in styles}
    badid = {n for n in STYLE_MAP.values() if style_id(n) not in ids}
    assert not badid, f'style_id mismatch for: {badid}'
    footer, = doc.main.xml.elements(w.FooterReference)

if __name__ == '__main__': build()
