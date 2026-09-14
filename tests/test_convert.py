import base64, pytest, subprocess, sys, zipfile
from pathlib import Path
from lxml import etree
from oxml import Document, e, w, namespaces

from fastcore.test import test_eq as teq, test as tt, test_fail as tfail
from fastcore.utils import in_

from mdhtml import DASHES, replacements, mdhtml2dom, md2mdhtml
from mdhtml.mustache import MUSTACHE
from mdhtml.tools import SAMPLE_MD, sample_md
from mdhtml2docx.mdhtml2docx import mdhtml2docx, mustache_fields
from mdhtml2docx.styles import ref_path
from mdhtml2docx.wml import NS, W


def pandoc(path, to='markdown'):
    "Read `path` back through pandoc's independent docx reader"
    r = subprocess.run(['pandoc', '-f', 'docx', '-t', to, '--wrap=none', str(path)],
        capture_output=True, text=True, check=True)
    return r.stdout


def xmlpart(path, part='document'):
    with zipfile.ZipFile(path) as z: return etree.fromstring(z.read(f'word/{part}.xml'))

def assert_valid(path): assert not (issues := Document.open(path).validate()['issues']), issues

def fields(doc):
    "Each simple field's instruction in the main document, without its padding"
    return [f.instruction.strip() for f in doc.main.xml.elements(w.SimpleField)]


def test_conversion_without_lxml(tmp_path):
    code = ("import sys\nsys.modules['lxml'] = None\nfrom mdhtml2docx import mdhtml2docx\n"
        "mdhtml2docx('<p>Hello</p>', sys.argv[1])")
    subprocess.run([sys.executable, '-c', code, str(tmp_path/'t.docx')], check=True, timeout=60)


@pytest.mark.parametrize('tag', ['p', 'div'])
def test_skeleton(tmp_path, tag):
    out = tmp_path/'t.docx'
    mdhtml2docx(f'<{tag}>Hello <em>world</em> with <strong>bold</strong> and <strong><em>both</em></strong>.</{tag}>\n'
        f'<{tag}>Second para: café – 東京…</{tag}>', out)
    assert_valid(out)
    md = pandoc(out)
    tt('Hello *world* with **bold** and ***both***.', md, in_)
    tt('Second para: café', md, in_)
    doc = Document.open(out)
    tt('café – 東京…', doc.story.text, in_)
    # pandoc-compatible First Paragraph convention: doc-start paragraph, then Body Text
    teq([s.val for s in doc.main.xml.elements(w.ParagraphStyleId)], ['FirstParagraph', 'BodyText'])


def test_html5_fragment_and_mutable_dom_input(tmp_path):
    out = tmp_path/'t.docx'
    dom = mdhtml2dom('<p>Before <div>middle</div> after.</p><span zoop:33="x">tail</span>'
        '<!-- this is a -- comment --><input type="date"><template><p>hidden</p></template>')
    dom.children[0].attrs['custom-style'] = 'Centered'
    warns = mdhtml2docx(dom, out)
    teq(warns, ["unhandled inline <input type='date'>; dropped"])
    tree = Document.open(out).main.xml
    assert [t.text for t in tree.elements(w.Text)] == ['Before ', 'middle', ' after.', 'tail']
    assert next(tree.elements(w.ParagraphStyleId)).val == 'Centered'
    mdhtml2docx('<template><p>still hidden</p></template>', out)
    assert not list(Document.open(out).main.xml.elements(w.Paragraph))


def test_basic_blocks(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(
        '<h1 id="top">Title</h1>\n<h2>Sub <em>title</em></h2>\n'
        '<p>Call <code>f(x)</code> or see <a href="https://fast.ai/">fast.ai</a> '
        'and <a href="#top">the title</a>.</p>\n'
        '<blockquote>\n<p>Quoted line.</p>\n<blockquote>Deeper.</blockquote>\n</blockquote>\n'
        '<pre><code class="language-python">def f(x):\n    return x\n</code></pre>', out)
    teq(warns, [])
    assert_valid(out)
    md = pandoc(out)
    for s in ('# Sub *title*', '`f(x)`', '[fast.ai](https://fast.ai/)',   # no '# Title': the h1 is Title style, which pandoc lifts to metadata (its bookmark still resolves below)
        '> Quoted line.', 'def f(x):', 'return x'): tt(s, md, in_)
    # the internal link survived the round trip: pandoc keeps the raw bookmark anchor, since a
    # Title-styled target is not a heading (heading targets get rewritten to auto-identifiers)
    tt('[the title](#top)', md, in_)
    assert xmlpart(out).xpath('//w:p[w:r/w:t="Deeper."]/w:pPr/w:ind/@w:left', namespaces=NS) == ['1440']


def test_lists(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(
        '<ul>\n<li>one</li>\n<li>two\n<ul>\n<li>deep</li>\n</ul>\n</li>\n</ul>\n'
        '<ol>\n<li>first</li>\n<li>second</li>\n</ol>\n'
        '<p>Another list:</p>\n'  # Pandoc merges adjacent lists sharing a numbering definition, losing the restart.
        '<ol start="5">\n<li>fifth</li>\n</ol>\n'
        '<ul class="task-list">\n'
        '<li><input type="checkbox" disabled="disabled" checked="checked"> done</li>\n'
        '<li><input type="checkbox" disabled="disabled"> todo</li>\n</ul>', out)
    teq(warns, [])
    assert_valid(out)
    lines = [' '.join(l.split()) for l in pandoc(out).splitlines()]
    for s in ('- one', '- deep', '1. first', '2. second',
        '5. fifth'): tt(s, lines, in_)   # 5: start attr honored; restart proves per-list numbering
    # pandoc's reader recognizes the ballot-box glyphs and reconstructs markdown task items
    tt('- [x] done', lines, in_)
    tt('- [ ] todo', lines, in_)


def test_adjacent_list_numbering_instances(tmp_path):
    "Check adjacent-list restarts directly in OOXML, independently of Pandoc's list grouping."
    out = tmp_path/'lists.docx'
    teq(mdhtml2docx('<ol><li>first</li><li>second</li></ol><ol start="5"><li>fifth</li></ol>', out), [])
    assert_valid(out)
    doc = Document.open(out)
    first, second, fifth = [n.val for n in doc.main.xml.elements(w.NumberingId)]
    teq(first, second)
    assert first != fifth
    for nid, start in ((first, 1), (fifth, 5)):
        overrides = doc.numbering[nid].element.elements(w.LevelOverride)
        level0, = (level for level in overrides if level.level_index == 0)
        override, = level0.elements(w.StartOverrideNumberingValue)
        teq(override.val, start)


def test_default_reference():
    "The bundled reference defines the Title and Centered styles and keeps an empty footer wired for page numbers"
    assert_valid(ref_path())
    doc = Document.open(ref_path())
    assert {s.id for s in doc.styles} >= {'Title', 'Centered'}
    assert doc.part('FooterPart') is not None
    footer, = doc.main.xml.elements(w.FooterReference)


def test_reference_generator_uses_native_schema_order(tmp_path):
    from tools.createref import build
    seed = Path(__file__).parents[1]/'_data'/'empty.docx'
    out = tmp_path/'reference.docx'
    build(seed, out)
    doc = Document.open(out)
    indent, = doc.styles['Quote'].element.elements(w.Indentation)
    assert indent.left == '720'
    for name in ('HeaderPart', 'FooterPart'):
        bidi, = doc.part(name).xml.elements(w.BiDiVisual)
        assert bidi.val == 'off'


@pytest.mark.parametrize('scheme,level2', [('legal', '(%3)'), ('decimal', '%2.%3.')])
def test_heading_numbering(tmp_path, scheme, level2):
    "Title and headings share a multilevel definition, with an invisible Title at level 0."
    out = tmp_path/'t.docx'
    warns = mdhtml2docx('<h1 id="ttl">EXHIBIT A: Assignment Agreement</h1>\n<h2 id="sec-conf">Confidentiality</h2>\n'
        '<h3>Confidential Information</h3>\n'
        '<p>See <a href="#sec-conf" data-ref=""></a> and <a href="#ttl" data-ref="bare text"></a>.</p>', out, number_headings=scheme)
    teq(warns, [])
    assert_valid(out)
    doc = Document.open(out)
    teq([s.val for s in doc.main.xml.elements(w.ParagraphStyleId)][:3], ['Title', 'Heading1', 'Heading2'])
    for s in (r'REF sec_conf \w \h', r'REF ttl \h'): tt(s, fields(doc), in_)
    nids = []
    for level, sid in enumerate(('Title', 'Heading1', 'Heading2')):
        style = doc.styles[sid].element
        ilvl, = style.elements(w.NumberingLevelReference)
        nid, = style.elements(w.NumberingId)
        assert ilvl.val == level
        nids.append(nid.val)
    assert len(set(nids)) == 1
    levels = doc.numbering[nids[0]].definition
    assert [v.val for v in levels.elements(w.LevelText)][:3] == ['', '%2.', level2]
    level0, = (level for level in levels.elements(w.Level) if level.level_index == 0)
    assert next(level0.elements(w.LevelSuffix)).val == 'nothing'


def test_tables(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(
        '<table colwidths="10em 2fr 1fr">\n<thead>\n'
        '<tr><th align="left">Feature</th><th align="center">Status</th><th align="right">Notes</th></tr>\n'
        '</thead>\n<tbody>\n'
        '<tr><td align="left">Tables</td><td align="center">ready</td><td align="right">ok</td></tr>\n'
        '</tbody>\n</table>\n'
        '<table>\n<tbody>\n'
        '<tr><td rowspan="2">a</td><td colspan="2">bc</td></tr>\n'
        '<tr><td>b</td><td>c</td></tr>\n'
        '</tbody>\n</table>', out)
    teq(warns, [])
    assert_valid(out)
    md = pandoc(out)
    for s in ('Feature', 'Status', 'Notes', 'Tables', 'ready', 'bc'): tt(s, md, in_)
    first, second = Document.open(out).main.xml.elements(w.Table)
    # 10em -> 2200 twips; 2fr/1fr share the remaining 7160 of the template's 9360 content width
    teq([c.width for c in first.elements(w.GridColumn)], ['2200', '4773', '2387'])
    header, = first.elements(w.TableHeader)                                    # thead row repeats across pages
    teq([m.val for m in second.elements(w.VerticalMerge)], ['restart', None])   # rowspan opened, then continued
    span, = second.elements(w.GridSpan)                                        # colspan encoded
    teq(span.val, 2)


def test_br_page(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(md2mdhtml('The Signature Page follows.<br type="page">\n\nNext page text.\n\nplain<br>break'), out)
    teq(warns, [])
    assert_valid(out)
    tree = Document.open(out).main.xml
    # the page break rides inside the marker's own paragraph, so it never needs a line of its own
    nodes = [n.text if isinstance(n, w.Text) else n.type for n in tree.elements() if isinstance(n, (w.Text, w.Break))]
    teq(nodes, ['The Signature Page follows.', 'page', 'Next page text.', 'plain', None, 'break'])   # a plain <br> stays a plain line break


def test_table_empty_header(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(md2mdhtml('|  |  |\n|---|---|\n| **A:** | one |\n| **B:** | two |'), out)
    teq(warns, [])
    assert_valid(out)
    tree = Document.open(out).main.xml
    teq(tree.count(w.TableHeader), 0)     # the all-empty thead is dropped, not rendered as a blank row
    teq(tree.count(w.TableRow), 2)        # only the two body rows remain
    teq([t.text for t in tree.elements(w.Text)], ['A:', 'one', 'B:', 'two'])


def test_more_features(tmp_path):
    out = tmp_path/'t.docx'
    img = Path(__file__).parent/'fixtures'/'tiny.png'
    warns = mdhtml2docx(
        '<p>H<sub>2</sub>O, E=mc<sup>2</sup>, <del>gone</del>, <mark>hot</mark>, <u>under</u>.</p>\n'
        '<hr>\n'
        '<dl>\n<dt>term</dt>\n<dd>definition here</dd><dd><p>continued definition</p></dd>\n</dl>\n'
        f'<p><img src="{img}" alt="tiny pic"></p>\n'
        '<p>Noted.<sup id="fnref-a"><a href="#fn-a" class="footnote-ref" role="doc-noteref">1</a></sup></p>\n'
        '<p>A <span custom-style="Fancy">styled bit</span>.</p>\n'
        '<section class="footnotes" role="doc-endnotes">\n<ol>\n<li id="fn-a">\n'
        '<p>The note text with a <a href="https://example.org/">link</a>.</p>\n'
        '<a href="#fnref-a" class="footnote-backref" role="doc-backlink">↩</a>\n'
        '</li>\n</ol>\n</section>', out)
    teq(warns, ["custom style 'Fancy' not in reference doc; stub injected"])
    assert_valid(out)
    md = pandoc(out)
    for s in ('H~2~O', 'E=mc^2^', '~~gone~~', 'hot', '[under]{.underline}', 'term', 'definition here',
        'continued definition', '![tiny pic](media/', '[^1]', 'The note text'): tt(s, md, in_)
    assert '↩' not in md            # backref stripped: the endnote became a real footnote
    doc = Document.open(out)
    tree = doc.main.xml
    border, = tree.elements(w.ParagraphBorders)                                    # hr
    teq([u.val for u in tree.elements(w.Underline)], ['single'])
    assert {s.val for s in tree.elements(w.ParagraphStyleId)} >= {'DefinitionTerm', 'Definition'}
    extent, = tree.elements(namespaces['wp'].Extent)
    teq((extent.cx, extent.cy), (38100, 19050))   # 4x2 px at 96dpi in EMU
    teq(doc.styles['Fancy'].name, 'Fancy')        # stub injected into the archive


def test_math(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx('<p>Euler: <span class="math inline">e^{i\\pi} + 1 = 0</span> inline.</p>\n'
        '<div class="math display">E = mc^2</div>', out)
    teq(warns, [])
    assert_valid(out)
    inline, display = xmlpart(out).findall('w:body/w:p', NS)
    assert inline.xpath('w:r/w:t/text() | m:oMath/m:r/m:t/text()', namespaces=NS) == ['Euler: ', 'e^{i\\pi} + 1 = 0', ' inline.']
    assert display.find('m:oMathPara/m:oMath/m:r/m:t', NS).text == 'E = mc^2'
    assert r'mc\hat{}2$$' in pandoc(out)


@pytest.mark.parametrize('body', [
    '<table><tr><td><div></div></td><td><table><tr><td>nested</td></tr></table></td></tr></table>',
    '<script type="application/vnd.mdhtml.raw" data-format="docx"><w:p><w:r><w:t>raw note</w:t></w:r></w:p></script>'])
def test_footnote_blocks(tmp_path, body):
    "A table-first note needs a reference paragraph; empty block containers still need cell paragraphs."
    out = tmp_path/'notes.docx'
    mdhtml2docx('<p>Note<sup><a href="#fn-a" class="footnote-ref">1</a></sup></p>'
        f'<section class="footnotes"><ol><li id="fn-a">{body}</li></ol></section>', out)
    notes = xmlpart(out, 'footnotes')
    assert_valid(out)
    note = notes.find('w:footnote[@w:id="1"]', NS)
    assert note[0].tag == f'{{{W}}}p' and note[0].find('w:r/w:footnoteRef', NS) is not None
    assert len(note.findall('.//w:footnoteRef', NS)) == 1
    assert all(cell[-1].tag == f'{{{W}}}p' for cell in note.findall('.//w:tc', NS))
    assert ('nested' if body.startswith('<table>') else 'raw note') in pandoc(out)


@pytest.mark.checkout
def test_sample(tmp_path):
    "The loop-closer: mdhtml's packaged sample -> docx (smart, legally numbered); every emitted style resolves"
    from mdhtml2docx.styles import STYLE_MAP, style_id
    out = tmp_path/'sample.docx'
    warns = mdhtml2docx(md2mdhtml(sample_md(), callbacks={'text': replacements(*DASHES)}, implicit_figures=True), out, base=SAMPLE_MD.parent, number_headings='legal')
    teq(warns, [])
    assert_valid(out)
    z = zipfile.ZipFile(out)
    used = set()
    for part in ('word/document.xml', 'word/footnotes.xml'):
        used |= {e.get(f'{{{W}}}val') for e in etree.fromstring(z.read(part)).iter()
            if etree.QName(e).localname in ('pStyle', 'rStyle', 'tblStyle')}
    defined = {s.get(f'{{{W}}}styleId') for s in etree.fromstring(z.read('word/styles.xml')).iter(f'{{{W}}}style')}
    assert used <= defined, f'undefined styles referenced: {used - defined}'
    hl = {u for u in used if u.startswith('Hl')}
    assert hl                                   # the sample's code blocks exercise the theme styles
    teq(used - hl, {style_id(v) for v in STYLE_MAP.values()})
    doc = Document.open(out)
    for s in (r'REF sec_late \w \h', r'REF sec_payment \w \h', r'PAGEREF sec_late \h', r'REF fig_diagram \h', r'REF tbl_stages_n \h',
        r'SEQ Figure \* ARABIC', r'SEQ Table \* ARABIC'): tt(s, fields(doc), in_)
    tt('page', [b.type for b in doc.main.xml.elements(w.Break)], in_)
    tt('Delivery stages', doc.story.text, in_)
    update, = doc.part('DocumentSettingsPart').xml.elements(w.UpdateFieldsOnOpen)
    md = pandoc(out)
    for s in ('mdhtml feature sample', 'Week one', 'Temperature 1961-1990', '-89.2', r'mc\hat{}2$$', '[^1]:'): tt(s, md, in_)


def test_code_highlight(tmp_path):
    "fastpylight scopes: language-classed blocks get Hl* character styles from the theme ref; text still round-trips"
    out = tmp_path/'t.docx'
    warns = mdhtml2docx('<pre><code class="language-python">def f(x):\n    return "café 東京"\n</code></pre>\n'
        '<pre><code>no language here</code></pre>', out)
    teq(warns, [])
    assert_valid(out)
    doc = Document.open(out)
    used = {s.val for s in doc.main.xml.elements(w.RunStyle)}
    assert any(s.startswith('HlKeyword') for s in used)    # 'def'/'return' (prefix: exact id depends on theme scopes)
    assert any(s.startswith('HlString') for s in used)
    tt('HlKeyword', {s.id for s in doc.styles}, in_)        # referenced styles are defined in the merged part
    md = pandoc(out)
    tt('    def f(x):\n        return "café 東京"', md, in_)
    tt('no language here', md, in_)


def test_theme_refs(tmp_path):
    "Multi-reference composition: later entries contribute styles later-wins; theme_ref writes an equivalent artifact"
    from fastpylight import theme_colors
    from mdhtml2docx.styles import ref_path, theme_ref
    mdhtml = '<pre><code class="language-python">def f(x):\n    return x\n</code></pre>'
    out = tmp_path/'t.docx'
    mdhtml2docx(mdhtml, out, reference=[ref_path(), 'dracula'])
    assert_valid(out)
    doc = Document.open(out)
    ids = [s.id for s in doc.styles]
    tt('HlKeyword', ids, in_)
    teq(ids.count('SourceCode'), 1)     # later-wins replaced the template's, not duplicated
    fill = theme_colors('dracula')['normal']['bg'].lstrip('#').upper()
    fills = {s.fill for s in doc.part('StyleDefinitionsPart').xml.elements(w.Shading)}
    tt(fill, fills, in_)                # dracula's dark code background won
    assert 'F5F5F5' not in fills
    # a theme_ref docx contributes the same styles as the bare theme name
    ref = theme_ref('dracula', tmp_path/'dracula.docx')
    assert_valid(ref)
    out2 = tmp_path/'t2.docx'
    mdhtml2docx(mdhtml, out2, reference=[ref_path(), ref])
    teq(etree.tostring(xmlpart(out2, 'styles'), method='c14n', exclusive=True),
        etree.tostring(xmlpart(out, 'styles'), method='c14n', exclusive=True))
    # single custom reference (no theme entry) = plain code
    out3 = tmp_path/'t3.docx'
    mdhtml2docx(mdhtml, out3, reference=ref_path())
    assert not any(s.val.startswith('Hl') for s in Document.open(out3).main.xml.elements(w.RunStyle))


def test_raw_docx(tmp_path):
    out = tmp_path/'t.docx'
    b64 = base64.b64encode(b'<w:r><w:t>B64</w:t></w:r>').decode()
    warns = mdhtml2docx(
        '<p>before</p>\n'
        '<script type="application/vnd.mdhtml.raw" data-format="docx"><w:p><w:r><w:br w:type="page"/></w:r></w:p></script>\n'
        '<p>a <script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="html">&lt;w:r&gt;&lt;w:t&gt;RAW&lt;/w:t&gt;&lt;/w:r&gt;</script> b</p>\n'
        f'<p><script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="base64">{b64}</script></p>\n'
        '<script type="application/vnd.mdhtml.raw" data-format="latex">\\newpage</script>', out)
    teq(warns, [])
    assert_valid(out)
    tree = Document.open(out).main.xml
    teq([b.type for b in tree.elements(w.Break)], ['page'])
    teq([t.text for t in tree.elements(w.Text)], ['before', 'a ', 'RAW', ' b', 'B64'])   # the unrecognized latex format was dropped
    warns = mdhtml2docx('<script type="application/vnd.mdhtml.raw" data-format="docx"><w:oops></script>', out)
    teq(len(warns), 1)
    tt('malformed', warns[0], in_)
    warns = mdhtml2docx('<script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="base64">!!!</script>', out)
    teq(len(warns), 1)
    tt('malformed base64', warns[0], in_)


def test_xrefs(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(
        '<h1 id="sec-intro">Intro</h1>\n<h2 id="sec-pay">Payment terms</h2>\n'
        '<p>See <a href="#sec-pay" data-ref=""></a> and <a href="#sec-intro" data-ref="bare"></a>, '
        'per <a href="#sec-pay" data-ref="">Clause</a>, also '
        '<span data-refs=""><a href="#sec-intro" data-ref=""></a>'
        '<a href="#sec-pay" data-ref=""></a></span>, the '
        '<a href="#sec-pay" data-ref="bare text"></a> clause, page '
        '<a href="#sec-pay" data-ref="bare page"></a>.</p>', out, number_headings='legal')
    teq(warns, [])
    assert_valid(out)
    doc = Document.open(out)
    for s in (r'REF sec_pay \w \h', r'REF sec_intro \w \h', r'PAGEREF sec_pay \h'): tt(s, fields(doc), in_)
    for s in ('Section ', 'Sections ', 'Clause ', ' and ', 'Payment terms'): tt(s, doc.story.text, in_)
    nums = doc.part('NumberingDefinitionsPart').xml
    tt('lowerLetter', [f.val for f in nums.elements(w.NumberingFormat)], in_)
    tt('(%3)', [t.val for t in nums.elements(w.LevelText)], in_)   # legal's '(%2)' sits on Word level 3 under the Title level
    tt('Heading1', [s.val for s in nums.elements(w.ParagraphStyleIdInLevel)], in_)
    update, = doc.part('DocumentSettingsPart').xml.elements(w.UpdateFieldsOnOpen)
    assert doc.part('StyleDefinitionsPart').xml.count(w.NumberingProperties)
    tfail(lambda: mdhtml2docx('<p><a href="#nope" data-ref=""></a></p>', out), contains='#nope')
    tfail(lambda: mdhtml2docx('<p id="z-a"><a href="#z-a" data-ref=""></a></p>', out),
        contains="reference type 'z'")
    tfail(lambda: mdhtml2docx('<p><a href="#sec-pay" data-ref="page text"></a></p>', out), contains='conflicting data-ref')


def test_text_targets(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(
        '<h1 id="sec-d">Defs</h1>\n'
        '<dl><dt id="def-term"><strong>"Term"</strong></dt><dt id="def-ap">Agreement Period</dt>'
        '<dd>the deal period</dd></dl>\n'
        '<p>The <span id="def-eff">Effective Date</span> is set.</p>\n'
        '<p>The <a href="#def-term" data-ref=""></a> starts on the <a href="#def-eff" data-ref=""></a>; '
        'see <a href="#def-ap" data-ref="page"></a>.</p>', out)
    teq(warns, [])
    assert_valid(out)
    doc = Document.open(out)
    for s in (r'REF def_term \h', r'REF def_eff \h', r'PAGEREF def_ap \h'): tt(s, fields(doc), in_)
    assert {b.name for b in doc.main.xml.elements(w.BookmarkStart)} >= {'def_term', 'def_eff'}
    for s in ('"Term"', 'Effective Date'): tt(s, doc.story.text, in_)

def test_xrefs_numbered_reference_doc(tmp_path):
    "Keep reference numbering unchanged and insert new definitions before numIdMacAtCleanup."
    ref = tmp_path/'ref.docx'
    mdhtml2docx('<h1 id="a">A</h1>', ref, number_headings='legal')
    doc = Document.open(ref)
    root = doc.package.part('/word/numbering.xml').xml.root
    root.append_xml(b'\n<!-- retain cleanup marker -->\n')
    root(e.numIdMacAtCleanup(val=42))
    doc.save(ref)
    original = xmlpart(ref, 'numbering')
    out = tmp_path/'t.docx'
    warns = mdhtml2docx('<h1 id="sec-a">A</h1>\n<ul>\n<li>one</li>\n</ul>\n'
        '<p>See <a href="#sec-a" data-ref=""></a>.</p>', out, reference=ref, number_headings='decimal')
    teq(warns, [])
    assert_valid(out)
    nums = xmlpart(out, 'numbering')
    for tag, added in (('abstractNum', 2), ('num', 1)):
        assert len(nums.findall(f'w:{tag}', NS)) == len(original.findall(f'w:{tag}', NS)) + added
        for old in original.findall(f'w:{tag}', NS):
            nid = old.get(f'{{{W}}}{tag}Id')
            new, = nums.findall(f'w:{tag}[@w:{tag}Id="{nid}"]', NS)
            assert etree.tostring(new, method='c14n', exclusive=True) == etree.tostring(old, method='c14n', exclusive=True)
    assert etree.tostring(xmlpart(out, 'styles')) == etree.tostring(xmlpart(ref, 'styles'))
    assert nums[-1].tag == f'{{{W}}}numIdMacAtCleanup' and nums[-1].get(f'{{{W}}}val') == '42'
    tt(r'REF sec_a \w \h', fields(Document.open(out)), in_)


def test_relocated_reference_parts(tmp_path):
    "Read and update reference parts through their relationships, not conventional filenames."
    original, reference, out = [tmp_path/f'{n}.docx' for n in ('original', 'reference', 'out')]
    src = ('<p id="sec-a">Before<a href="#sec-a" data-ref="text"></a>'
        '<sup><a href="#fn-a" class="footnote-ref">1</a></sup></p><ul><li>listed</li></ul>'
        '<section class="footnotes"><ol><li id="fn-a"><p>Before</p></li></ol></section>')
    mdhtml2docx(src, original)
    renames = {'word/': 'content/', 'document.xml': 'main.xml', 'styles.xml': 'styles-custom.xml',
        'numbering.xml': 'numbering-custom.xml', 'settings.xml': 'settings-custom.xml', 'footnotes.xml': 'footnotes-custom.xml'}
    def renamed(s):
        for old, new in renames.items(): s = s.replace(old, new)
        return s
    with zipfile.ZipFile(original) as zi, zipfile.ZipFile(reference, 'w', zipfile.ZIP_DEFLATED) as zo:
        for name in zi.namelist():
            data = zi.read(name)
            if name.endswith('.rels') or name == '[Content_Types].xml': data = renamed(data.decode('utf-8-sig')).encode()
            zo.writestr(renamed(name), data)
    assert not mdhtml2docx(src.replace('Before', 'After'), out, reference=reference)
    assert_valid(out)
    with zipfile.ZipFile(reference) as zi, zipfile.ZipFile(out) as zo:
        assert set(zo.namelist()) == set(zi.namelist())  # No replacement parts at conventional paths.
        for name in ('content/main.xml', 'content/footnotes-custom.xml'): assert b'>After<' in zo.read(name)


def test_reference_media_and_story_relationships(tmp_path):
    "New media avoids reference names; body and footnote links resolve in their own relationship scopes."
    img = Path(__file__).parent/'fixtures'/'tiny.png'
    reference, out = tmp_path/'reference.docx', tmp_path/'out.docx'
    content = f'<a href="https://example.org/">link</a><img src="{img}">'
    src = (f'<p>{content}<sup><a href="#fn-a" class="footnote-ref">1</a></sup></p>'
        f'<section class="footnotes"><ol><li id="fn-a"><p>{content}</p></li></ol></section>')
    mdhtml2docx(src, reference)
    assert not mdhtml2docx(src, out, reference=reference)
    package = Document.open(out).package
    with zipfile.ZipFile(reference) as zi, zipfile.ZipFile(out) as zo:
        for name in zi.namelist():
            if name.startswith('word/media/'): assert zo.read(name) == zi.read(name)
        for part in ('document', 'footnotes'):
            tree = package.part(f'/word/{part}.xml').xml
            targets = {r['id']: r['target'] for r in package.relationships(f'/word/{part}.xml')}
            link, = tree.elements(w.Hyperlink)
            assert targets[link.id] == 'https://example.org/'
            blip, = tree.elements(namespaces['a'].Blip)
            target = package.relationship_part(f'/word/{part}.xml', blip.embed).lstrip('/')
            assert target not in zi.namelist() and zo.read(target) == img.read_bytes()


def test_xml_contributor(tmp_path):
    "Each contributor has its own numbering ID scope; later styles win without cross-wiring references."
    refs = []
    for i, (name, fmt) in enumerate((('Fancy', 'upperRoman'), ('Other', 'lowerLetter'), ('Fancy', 'lowerRoman'))):
        xml = tmp_path/f'extra{i}.xml'
        xml.write_text(  # chkstyle: ignore-node
            f'<w:styles xmlns:w="{W}"><w:abstractNum w:abstractNumId="0">'
            f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>'
            '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            '</w:style></w:styles>')
        refs.append(xml)
    out = tmp_path/'t.docx'
    warns = mdhtml2docx('<p custom-style="Fancy">Hi</p><p custom-style="Other">Bye</p><ul><li>one</li></ul>', out,
        reference=[None, *refs], number_headings='legal')
    teq(warns, [])
    assert_valid(out)
    doc = Document.open(out)
    nums = doc.part('NumberingDefinitionsPart').xml
    nids = [n.number_id for n in nums.elements(w.NumberingInstance)]
    aids = [n.abstract_number_id for n in nums.elements(w.AbstractNum)]
    for ids in (nids, aids): assert len(ids) == len(set(ids))
    for name, fmt in (('Fancy', 'lowerRoman'), ('Other', 'lowerLetter')):
        nid, = doc.styles[name].element.elements(w.NumberingId)
        numfmt, = doc.numbering[nid.val].definition.elements(w.NumberingFormat)
        assert numfmt.val == fmt


def test_caption_refs(tmp_path):
    "Figures and captioned tables: SEQ-numbered captions, label+number bookmarks, fig/tbl refs"
    out = tmp_path/'t.docx'
    img = Path(__file__).parent/'fixtures'/'tiny.png'
    warns = mdhtml2docx(
        f'<figure id="fig-plot" class="wide"><img src="{img}" alt=""><figcaption>A plot</figcaption></figure>\n'
        '<table id="tbl-r"><caption>Results <em>now</em></caption><thead><tr><th>a</th></tr></thead>'
        '<tbody><tr><td>1</td></tr></tbody></table>\n'
        '<p>See <a href="#fig-plot" data-ref=""></a> and <a href="#tbl-r" data-ref="bare"></a>, or '
        '<span data-refs=""><a href="#fig-plot" data-ref=""></a>'
        '<a href="#tbl-r" data-ref=""></a></span>.</p>', out)
    teq(warns, [])
    assert_valid(out)
    doc = Document.open(out)
    for s in (r'SEQ Figure \* ARABIC', r'SEQ Table \* ARABIC', r'REF fig_plot \h', r'REF tbl_r_n \h'): tt(s, fields(doc), in_)
    for s in ('Results ', 'A plot'): tt(s, doc.story.text, in_)
    props, = doc.main.xml.elements(namespaces['wp'].DocProperties)
    teq(props.description, 'A plot')
    assert 'Figures' not in doc.story.text          # mixed group: per-item singular prefixes
    assert r'REF fig_plot \w \h' not in fields(doc)  # caption targets never use \w
    md = pandoc(out)
    tt('A plot', md, in_)
    # a ref to an id that never becomes a bookmark is an error, not a dud field
    tfail(lambda: mdhtml2docx('<div id="d-x"><p>hi</p></div><p><a href="#d-x" data-ref=""></a></p>', out),
        contains='#d-x')


def test_template_tokens(tmp_path):
    out = tmp_path/'t.docx'
    src = md2mdhtml('Pay {{sal}} to {{name}}.\n\n{{#opt}}\n\nGranted.\n\n{{/opt}}\n', templates=MUSTACHE)
    mdhtml2docx(src, out)                                    # default: vars dropped, markers literal
    doc = Document.open(out)
    text = doc.story.text
    assert 'sal' not in text and '{{' not in text and '«#opt»' in text and not fields(doc)   # vars dropped; markers stay visible
    warns = mdhtml2docx(src, out, tmpl=mustache_fields)
    doc = Document.open(out)
    assert 'MERGEFIELD sal' in fields(doc) and 'MERGEFIELD name' in fields(doc)
    assert '«#opt»' in doc.story.text and '«/opt»' in doc.story.text          # markers: converter-rendered literals, own paragraphs
    teq(warns, [])
    assert_valid(out)

def test_table_custom_style(tmp_path):
    out = tmp_path/'t.docx'
    mdhtml2docx(md2mdhtml('| A |\n|---|\n| b |\n{: custom-style="Borderless Table"}\n\n| C |\n|---|\n| d |\n'), out)
    assert [s.val for s in Document.open(out).main.xml.elements(w.TableStyle)] == ['BorderlessTable', 'TableGrid']   # styled table picks the reference style; the plain one keeps the default
    assert_valid(out)



def test_template_controls(tmp_path):
    def controls(node): return 'control', node['value']
    out = tmp_path/'t.docx'
    mdhtml2docx(md2mdhtml('Pay {{sal}} to {{name}}.\n\n{{#opt}}\n', templates=MUSTACHE), out, tmpl=controls)
    doc = Document.open(out)
    tree = doc.main.xml
    teq(tree.count(w.SdtRun), 2)
    teq({t.val for t in tree.elements(w.Tag)}, {'sal', 'name'})
    teq(tree.count(w.ShowingPlaceholder), 2)
    teq({s.val for s in tree.elements(w.RunStyle)}, {'PlaceholderText'})
    color, = doc.styles['PlaceholderText'].element.elements(w.Color)
    teq(color.val, '808080')
    assert '«#opt»' in doc.story.text                     # marker: converter-rendered literal
    assert_valid(out)


def test_template_bound(tmp_path):
    def bound(node): return 'bound', node['value']
    out = tmp_path/'t.docx'
    mdhtml2docx(md2mdhtml('Pay {{sal}} to {{sal}} and {{name}}.\n', templates=MUSTACHE), out, tmpl=bound)
    z = zipfile.ZipFile(out)
    bindings = [b.x_path for b in Document.open(out).main.xml.elements(w.DataBinding)]
    teq(len(bindings), 3)
    teq(bindings.count('/ns0:fields[1]/ns0:sal[1]'), 2)
    cx = z.read('customXml/item1.xml').decode()
    assert cx.count('<ns0:sal/>') == 1 and '<ns0:name/>' in cx   # deduped inventory, one element per variable
    assert 'customXmlProperties' in z.read('[Content_Types].xml').decode()
    assert 'customXml' in z.read('word/_rels/document.xml.rels').decode()
    assert_valid(out)
    # A generated reference retains one datastore with the fixed binding ID, not an ambiguous duplicate.
    again = tmp_path/'again.docx'
    mdhtml2docx(md2mdhtml('Now {{name}}.\n', templates=MUSTACHE), again, reference=out, tmpl=bound)
    with zipfile.ZipFile(again) as za:
        assert sorted(n for n in za.namelist() if n.startswith('customXml/')) == sorted(n for n in z.namelist() if n.startswith('customXml/'))
        assert '<ns0:name/>' in za.read('customXml/item1.xml').decode() and '<ns0:sal/>' not in za.read('customXml/item1.xml').decode()
        assert za.read('customXml/itemProps1.xml') == z.read('customXml/itemProps1.xml')


def test_details_degrades_to_bold_label(tmp_path):
    out = tmp_path/'t.docx'
    warns = mdhtml2docx(md2mdhtml('::: {.details .tool-usage-details}\n## the label\n\nbody text\n:::\n'), out)
    assert not warns
    md = pandoc(out)
    tt('**the label**', md, in_)
    tt('body text', md, in_)
    assert '# the label' not in md   # label is a bold line, not a heading


@pytest.mark.parametrize('attr,keep', [('', 'on'), ('{: keep-rows=true}', 'on'), ('{: keep-rows=false}', 'off')])
def test_table_row_page_breaks(tmp_path, attr, keep):
    out = tmp_path/'rows.docx'
    md = '| Name | Value |\n|---|---|\n| Award | Blank |\n' + attr
    teq(mdhtml2docx(md2mdhtml(md), out), [])
    assert_valid(out)
    tree = Document.open(out).main.xml
    teq([flag.val for flag in tree.elements(w.CantSplit)], [keep, keep])
    header, = tree.elements(w.TableHeader)


@pytest.mark.parametrize('state', ['fixed', 'open', 'closed'])
def test_panel_contract(tmp_path, state):
    out = tmp_path/'panel.docx'
    src = (f'<div data-panel data-callout="warning" data-disclosure="{state}">'
        '<header id="label">The <em>label</em></header>Body text<p>More body</p><h3>Actual heading</h3></div>')
    assert not mdhtml2docx(src, out)
    assert_valid(out)
    text = pandoc(out)
    assert 'label' in text and 'Body text' in text and 'More body' in text
    assert '## Actual heading' in text  # h1 maps to Word's Title; h3 maps to Heading 2
    assert not any(line.startswith('#') and 'label' in line for line in text.splitlines())
