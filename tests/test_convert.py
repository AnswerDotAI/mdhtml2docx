import base64, subprocess, zipfile
from pathlib import Path

from fastcore.test import test_eq as teq, test as tt, test_fail as tfail
from fastcore.utils import in_

from mdhtml import JINJA, MUSTACHE, mustache_kind, parse_mdhtml, sample_md, to_mdhtml
from mdhtml2docx.convert import convert, jinja_literal, mustache_fields
from mdhtml2docx.validate import fast_checks


def pandoc(path, to='markdown'):
    "Read `path` back through pandoc's independent docx reader"
    r = subprocess.run(['pandoc', '-f', 'docx', '-t', to, '--wrap=none', str(path)],
        capture_output=True, text=True, check=True)
    return r.stdout


def test_skeleton(tmp_path):
    out = tmp_path/'t.docx'
    convert('<p>Hello <em>world</em> with <strong>bold</strong> and <strong><em>both</em></strong>.</p>\n'
        '<p>Second para.</p>', out)
    teq(fast_checks(out), 'valid')
    md = pandoc(out)
    tt('Hello *world* with **bold** and ***both***.', md, in_)
    tt('Second para.', md, in_)
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    # pandoc-compatible First Paragraph convention: doc-start paragraph, then Body Text
    assert doc.index('w:val="FirstParagraph"') < doc.index('w:val="BodyText"')


def test_html5_fragment_and_mutable_dom_input(tmp_path):
    out = tmp_path/'t.docx'
    dom = parse_mdhtml('<p>Before <div>middle</div> after.</p><span zoop:33="x">tail</span>'
        '<!-- this is a -- comment --><input type="date"><template><p>hidden</p></template>')
    dom.children[1].attrs['data-kind'] = 'changed'
    warns = convert(dom, out)
    teq(warns, ["unhandled inline <input type='date'>; dropped"])
    md = pandoc(out)
    for s in ('Before', 'middle', 'after.', 'tail'): tt(s, md, in_)
    assert 'hidden' not in md
    convert('<template><p>still hidden</p></template>', out)
    assert '<w:p>' not in zipfile.ZipFile(out).read('word/document.xml').decode()


def test_basic_blocks(tmp_path):
    out = tmp_path/'t.docx'
    warns = convert(
        '<h1 id="top">Title</h1>\n<h2>Sub <em>title</em></h2>\n'
        '<p>Call <code>f(x)</code> or see <a href="https://fast.ai/">fast.ai</a> '
        'and <a href="#top">the title</a>.</p>\n'
        '<blockquote>\n<p>Quoted line.</p>\n<blockquote><p>Deeper.</p></blockquote>\n</blockquote>\n'
        '<pre><code class="language-python">def f(x):\n    return x\n</code></pre>', out)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    md = pandoc(out)
    for s in ('# Title', '## Sub *title*', '`f(x)`', '[fast.ai](https://fast.ai/)',
        '> Quoted line.', 'def f(x):', 'return x'): tt(s, md, in_)
    # pandoc resolves the anchor against the heading bookmark and rewrites it to the heading's
    # auto-identifier, so this line proves the internal link wiring survived the round trip
    tt('[the title](#title)', md, in_)


def test_lists(tmp_path):
    out = tmp_path/'t.docx'
    warns = convert(
        '<ul>\n<li>one</li>\n<li>two\n<ul>\n<li>deep</li>\n</ul>\n</li>\n</ul>\n'
        '<ol>\n<li>first</li>\n<li>second</li>\n</ol>\n'
        '<ol start="5">\n<li>fifth</li>\n</ol>\n'
        '<ul class="task-list">\n'
        '<li><input type="checkbox" disabled="disabled" checked="checked"> done</li>\n'
        '<li><input type="checkbox" disabled="disabled"> todo</li>\n</ul>', out)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    lines = [' '.join(l.split()) for l in pandoc(out).splitlines()]
    for s in ('- one', '- deep', '1. first', '2. second',
        '5. fifth'): tt(s, lines, in_)   # 5: start attr honored; restart proves per-list numbering
    # pandoc's reader recognizes the ballot-box glyphs and reconstructs markdown task items
    tt('- [x] done', lines, in_)
    tt('- [ ] todo', lines, in_)


def test_tables(tmp_path):
    out = tmp_path/'t.docx'
    warns = convert(
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
    teq(fast_checks(out), 'valid')
    md = pandoc(out)
    for s in ('Feature', 'Status', 'Notes', 'Tables', 'ready', 'bc'): tt(s, md, in_)
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    # 10em -> 2200 twips; 2fr/1fr share the remaining 7160 of the template's 9360 content width
    for s in ('<w:gridCol w:w="2200"/>', '<w:gridCol w:w="4773"/>', '<w:gridCol w:w="2387"/>',
        '<w:tblHeader/>',           # thead row repeats across pages
        'w:val="restart"',          # rowspan opened
        '<w:gridSpan w:val="2"/>'): tt(s, doc, in_)   # colspan encoded


def test_more_features(tmp_path):
    out = tmp_path/'t.docx'
    img = Path(__file__).parent/'fixtures'/'tiny.png'
    warns = convert(
        '<p>H<sub>2</sub>O, E=mc<sup>2</sup>, <del>gone</del>, <mark>hot</mark>, <u>under</u>.</p>\n'
        '<hr>\n'
        '<dl>\n<dt>term</dt>\n<dd>definition here</dd>\n</dl>\n'
        f'<p><img src="{img}" alt="tiny pic"></p>\n'
        '<p>Noted.<sup id="fnref-a"><a href="#fn-a" class="footnote-ref" role="doc-noteref">1</a></sup></p>\n'
        '<p>Math: <span class="math inline">a^2+b^2</span>.</p>\n'
        '<div class="math display">E = mc^2</div>\n'
        '<p>A <span custom-style="Fancy">styled bit</span>.</p>\n'
        '<section class="footnotes" role="doc-endnotes">\n<ol>\n<li id="fn-a">\n'
        '<p>The note text with a <a href="https://example.org/">link</a>.</p>\n'
        '<a href="#fnref-a" class="footnote-backref" role="doc-backlink">↩</a>\n'
        '</li>\n</ol>\n</section>', out)
    teq(warns, ["custom style 'Fancy' not in reference doc; stub injected"])
    teq(fast_checks(out), 'valid')
    md = pandoc(out)
    for s in ('H~2~O', 'E=mc^2^', '~~gone~~', 'hot', '[under]{.underline}', 'term', 'definition here',
        '![tiny pic](media/', '[^1]', 'The note text',
        'Math: $', r'b\hat{}2$', r'mc\hat{}2$$'): tt(s, md, in_)   # zones read back as pandoc math
    assert '↩' not in md            # backref stripped: the endnote became a real footnote
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    for s in ('<w:pBdr>', '<w:u w:val="single"/>',                            # hr
        'w:val="DefinitionTerm"', 'w:val="Definition"',
        'cx="38100" cy="19050"'): tt(s, doc, in_)   # 4x2 px at 96dpi in EMU
    styles = zipfile.ZipFile(out).read('word/styles.xml').decode()
    tt('w:val="Fancy"', styles, in_)                 # stub injected into the archive


def test_imgsize():
    "Sniffer vs real Pillow encodings, dpi included (png pHYs, jpeg JFIF density, gif)"
    import io
    from PIL import Image
    def enc(fmt, size, **kw):
        b = io.BytesIO()
        Image.new('RGB' if fmt != 'GIF' else 'P', size).save(b, fmt, **kw)
        return b.getvalue()
    from mdhtml2docx.wml import imgsize
    teq(imgsize(enc('PNG', (4, 2))), (4, 2, 96, 96))
    teq(imgsize(enc('PNG', (10, 5), dpi=(150, 150))), (10, 5, 150, 150))
    teq(imgsize(enc('JPEG', (10, 5), dpi=(200, 100))), (10, 5, 200, 100))
    teq(imgsize(enc('GIF', (7, 3))), (7, 3, 96, 96))
    teq(imgsize(b'not an image'), None)


def test_math(tmp_path):
    out = tmp_path/'t.docx'
    warns = convert('<p>Euler: <span class="math inline">e^{i\\pi} + 1 = 0</span> inline.</p>\n'
        '<div class="math display">E = mc^2</div>', out)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    # inline zone sits between its neighbouring text runs, inside the same paragraph
    assert doc.index('Euler:') < doc.index('<m:oMath>') < doc.index('e^{i\\pi} + 1 = 0') < doc.index('inline.')
    # display zone is a block-level math paragraph
    tt('<m:oMathPara>', doc, in_)
    tt('E = mc^2', doc, in_)


def test_sample(tmp_path):
    "The loop-closer: mdhtml's packaged sample -> docx (smart, legally numbered); every emitted style resolves"
    from lxml import etree
    from mdhtml2docx.styles import STYLE_MAP, style_id
    out = tmp_path/'sample.docx'
    warns = convert(to_mdhtml(sample_md(), smart=True, auto_ids=True, implicit_figures=True), out, number_headings='legal')
    teq(warns, ['remote image not embedded: https://dummyimage.com/96x48/eeeeee/333333.png&text=demo',
        'remote image not embedded: https://dummyimage.com/96x48/eeeeee/333333.png&text=fig'])
    teq(fast_checks(out), 'valid')
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
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
    doc = z.read('word/document.xml').decode()
    for s in (r'REF sec_late \w \h', r'REF sec_payment \w \h', r'PAGEREF sec_late \h',
        r'REF fig_diagram \h', r'REF tbl_stages_n \h', ' SEQ Figure ', ' SEQ Table ',
        '<w:br w:type="page"/>', 'Delivery stages', '“'): tt(s, doc, in_)
    tt('updateFields', z.read('word/settings.xml').decode(), in_)
    md = pandoc(out)
    for s in ('# mdhtml feature sample', '###### Week one', 'Temperature 1961-1990', '-89.2',
        r'mc\hat{}2$$', '[^1]:', 'A Markdown parser that renders MDHTML fragments.'): tt(s, md, in_)


def test_code_highlight(tmp_path):
    "fastpylight scopes: language-classed blocks get Hl* character styles from the theme ref; text still round-trips"
    out = tmp_path/'t.docx'
    warns = convert('<pre><code class="language-python">def f(x):\n    return "hi"\n</code></pre>\n'
        '<pre><code>no language here</code></pre>', out)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    tt('<w:rStyle w:val="HlKeyword', doc, in_)    # 'def'/'return' (prefix: exact id depends on theme scopes)
    tt('<w:rStyle w:val="HlString', doc, in_)     # "hi"
    styles = zipfile.ZipFile(out).read('word/styles.xml').decode()
    tt('w:styleId="HlKeyword"', styles, in_)      # referenced styles are defined in the merged part
    tt('w:val="CF222E"', styles, in_)             # github_light keyword red now lives in the style, not the run
    md = pandoc(out)
    tt('def f(x):', md, in_)
    tt('no language here', md, in_)


def test_theme_refs(tmp_path):
    "Multi-reference composition: later entries contribute styles later-wins; theme_ref writes an equivalent artifact"
    from fastpylight import theme_colors
    from mdhtml2docx.styles import ref_path, theme_ref
    mdhtml = '<pre><code class="language-python">def f(x):\n    return x\n</code></pre>'
    out = tmp_path/'t.docx'
    convert(mdhtml, out, reference=[ref_path(), 'dracula'])
    teq(fast_checks(out), 'valid')
    styles = zipfile.ZipFile(out).read('word/styles.xml').decode()
    tt('w:styleId="HlKeyword"', styles, in_)
    teq(styles.count('w:styleId="SourceCode"'), 1)     # later-wins replaced the template's, not duplicated
    fill = theme_colors('dracula')['normal']['bg'].lstrip('#').upper()
    tt(f'w:fill="{fill}"', styles, in_)                # dracula's dark code background won
    assert 'w:fill="F5F5F5"' not in styles
    # a theme_ref docx contributes the same styles as the bare theme name
    ref = theme_ref('dracula', tmp_path/'dracula.docx')
    teq(fast_checks(ref), 'valid')
    out2 = tmp_path/'t2.docx'
    convert(mdhtml, out2, reference=[ref_path(), ref])
    ids = lambda p: sorted(s.split('"')[1] for s in zipfile.ZipFile(p).read('word/styles.xml').decode().split('w:styleId=')[1:])
    teq(ids(out2), ids(out))
    # single custom reference (no theme entry) = plain code
    out3 = tmp_path/'t3.docx'
    convert(mdhtml, out3, reference=ref_path())
    assert 'HlKeyword' not in zipfile.ZipFile(out3).read('word/document.xml').decode()


def test_segments_scopes():
    "segments returns raw dotted scopes tiling the source; markdown quotation keeps embedded fences inert"
    from mdhtml2docx.hilite import segments
    md = 'Some `inline` and:\n\n``` rust\nfn main() {}\n```\n'
    segs = segments(md, 'markdown')
    teq(next(s for t, s in segs if 'fn main' in t), 'markup.raw.block')
    teq(next(s for t, s in segs if t == '`inline`'), 'markup.raw')
    teq(''.join(t for t, _ in segs), md)


def test_raw_docx(tmp_path):
    out = tmp_path/'t.docx'
    b64 = base64.b64encode(b'<w:r><w:t>B64</w:t></w:r>').decode()
    warns = convert(
        '<p>before</p>\n'
        '<script type="application/vnd.mdhtml.raw" data-format="docx"><w:p><w:r><w:br w:type="page"/></w:r></w:p></script>\n'
        '<p>a <script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="html">&lt;w:r&gt;&lt;w:t&gt;RAW&lt;/w:t&gt;&lt;/w:r&gt;</script> b</p>\n'
        f'<p><script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="base64">{b64}</script></p>\n'
        '<script type="application/vnd.mdhtml.raw" data-format="latex">\\newpage</script>', out)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    tt('<w:br w:type="page"/>', doc, in_)
    md = pandoc(out)
    for s in ('a RAW b', 'B64'): tt(s, md, in_)
    assert 'newpage' not in doc  # unrecognized format dropped
    warns = convert('<script type="application/vnd.mdhtml.raw" data-format="docx"><w:oops></script>', out)
    teq(len(warns), 1)
    tt('malformed', warns[0], in_)
    warns = convert('<script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="base64">!!!</script>', out)
    teq(len(warns), 1)
    tt('malformed base64', warns[0], in_)


def test_xrefs(tmp_path):
    out = tmp_path/'t.docx'
    warns = convert(
        '<h1 id="sec-intro">Intro</h1>\n<h2 id="sec-pay">Payment terms</h2>\n'
        '<p>See <a href="#sec-pay" data-ref=""></a> and <a href="#sec-intro" data-ref="bare"></a>, '
        'per <a href="#sec-pay" data-ref="">Clause</a>, also '
        '<span data-refs=""><a href="#sec-intro" data-ref=""></a>'
        '<a href="#sec-pay" data-ref=""></a></span>, the '
        '<a href="#sec-pay" data-ref="bare text"></a> clause, page '
        '<a href="#sec-pay" data-ref="bare page"></a>.</p>', out, number_headings='legal')
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    for s in (r'REF sec_pay \w \h', r'REF sec_intro \w \h', r'PAGEREF sec_pay \h',
        'Section ', 'Sections ', 'Clause ', ' and ', 'Payment terms'): tt(s, doc, in_)
    num = zipfile.ZipFile(out).read('word/numbering.xml').decode()
    for s in ('lowerLetter', '(%2)', 'Heading1'): tt(s, num, in_)
    tt('updateFields', zipfile.ZipFile(out).read('word/settings.xml').decode(), in_)
    tt('w:numPr', zipfile.ZipFile(out).read('word/styles.xml').decode(), in_)
    tfail(lambda: convert('<p><a href="#nope" data-ref=""></a></p>', out), contains='#nope')
    tfail(lambda: convert('<p id="z-a"><a href="#z-a" data-ref=""></a></p>', out),
        contains="reference type 'z'")
    tfail(lambda: convert('<p><a href="#sec-pay" data-ref="page text"></a></p>', out), contains='conflicting data-ref')


def test_xrefs_numbered_reference_doc(tmp_path):
    "A numbered docx (made by us) as reference: numbering merges with list numbering, no reinjection"
    ref = tmp_path/'ref.docx'
    convert('<h1 id="a">A</h1>', ref, number_headings='legal')
    out = tmp_path/'t.docx'
    warns = convert('<h1 id="sec-a">A</h1>\n<ul>\n<li>one</li>\n</ul>\n'
        '<p>See <a href="#sec-a" data-ref=""></a>.</p>', out, reference=ref)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    num = zipfile.ZipFile(out).read('word/numbering.xml').decode()
    tt('Heading1', num, in_)     # heading numbering survived from the reference
    tt('bullet', num, in_)       # list numbering added alongside
    tt(r'REF sec_a \w \h', zipfile.ZipFile(out).read('word/document.xml').decode(), in_)


def test_xml_contributor(tmp_path):
    "A raw styles .xml reference entry contributes styles and numbering"
    xml = tmp_path/'extra.xml'
    xml.write_text(  # chkstyle: ignore-node
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/>'
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        '</w:style></w:styles>')
    out = tmp_path/'t.docx'
    warns = convert('<p custom-style="Fancy">Hi</p>\n<ul>\n<li>one</li>\n</ul>', out,
        reference=[None, xml])
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    num = zipfile.ZipFile(out).read('word/numbering.xml').decode()
    tt('upperRoman', num, in_)
    sty = zipfile.ZipFile(out).read('word/styles.xml').decode()
    tt('Fancy', sty, in_)


def test_caption_refs(tmp_path):
    "Figures and captioned tables: SEQ-numbered captions, label+number bookmarks, fig/tbl refs"
    out = tmp_path/'t.docx'
    img = Path(__file__).parent/'fixtures'/'tiny.png'
    warns = convert(
        f'<figure id="fig-plot" class="wide"><img src="{img}" alt=""><figcaption>A plot</figcaption></figure>\n'
        '<table id="tbl-r"><caption>Results <em>now</em></caption><thead><tr><th>a</th></tr></thead>'
        '<tbody><tr><td>1</td></tr></tbody></table>\n'
        '<p>See <a href="#fig-plot" data-ref=""></a> and <a href="#tbl-r" data-ref="bare"></a>, or '
        '<span data-refs=""><a href="#fig-plot" data-ref=""></a>'
        '<a href="#tbl-r" data-ref=""></a></span>.</p>', out)
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    for s in (' SEQ Figure ', ' SEQ Table ', r'REF fig_plot \h', r'REF tbl_r_n \h', 'Results ', 'A plot'): tt(s, doc, in_)
    tt('descr="A plot"', doc, in_)
    assert 'Figures' not in doc                  # mixed group: per-item singular prefixes
    assert r'REF fig_plot \w' not in doc         # caption targets never use \w
    md = pandoc(out)
    tt('A plot', md, in_)
    # a ref to an id that never becomes a bookmark is an error, not a dud field
    tfail(lambda: convert('<div id="d-x"><p>hi</p></div><p><a href="#d-x" data-ref=""></a></p>', out),
        contains='#d-x')


def test_template_tokens(tmp_path):
    out = tmp_path/'t.docx'
    src = to_mdhtml('Pay {{sal}} to {{name}}.\n\n{{#opt}}\n\nGranted.\n\n{{/opt}}\n', templates=MUSTACHE)
    convert(src, out)                                    # default: tokens dropped, as before
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    assert 'sal' not in doc and '{{' not in doc
    warns = convert(src, out, tmpl=mustache_fields)
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    assert 'MERGEFIELD sal' in doc and 'MERGEFIELD name' in doc
    assert '{{#opt}}' in doc and '{{/opt}}' in doc       # section markers stay literal, own paragraphs
    teq(warns, [])
    teq(fast_checks(out), 'valid')
    convert(to_mdhtml('V {{ v }}.\n\n{% if x %}\n', templates=JINJA), out, tmpl=jinja_literal)
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    assert '{{ v }}' in doc and '{% if x %}' in doc
    teq(fast_checks(out), 'valid')

def test_table_custom_style(tmp_path):
    out = tmp_path/'t.docx'
    convert(to_mdhtml('| A |\n|---|\n| b |\n{: custom-style="Borderless Table"}\n\n| C |\n|---|\n| d |\n'), out)
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    tt('<w:tblStyle w:val="BorderlessTable"/>', doc, in_)          # styled table picks the reference style
    tt('<w:tblStyle w:val="TableGrid"/>', doc, in_)                # plain table keeps the default
    teq(fast_checks(out), 'valid')




def test_template_controls(tmp_path):
    def controls(body, syntax, form):
        if mustache_kind(body) == 'section': return None
        return 'control', body
    out = tmp_path/'t.docx'
    convert(to_mdhtml('Pay {{sal}} to {{name}}.\n\n{{#opt}}\n', templates=MUSTACHE), out, tmpl=controls)
    doc = zipfile.ZipFile(out).read('word/document.xml').decode()
    assert doc.count('<w:sdt>') == 2 and 'w:val="sal"' in doc and 'w:val="name"' in doc
    assert doc.count('<w:showingPlcHdr/>') == 2 and 'w:val="PlaceholderText"' in doc
    styles = zipfile.ZipFile(out).read('word/styles.xml').decode()
    assert 'w:styleId="PlaceholderText"' in styles and 'w:val="808080"' in styles
    assert '{{' not in doc                                   # section marker dropped by this callable
    teq(fast_checks(out), 'valid')


def test_template_bound(tmp_path):
    def bound(body, syntax, form):
        if mustache_kind(body) == 'section': return None
        return 'bound', body
    out = tmp_path/'t.docx'
    convert(to_mdhtml('Pay {{sal}} to {{sal}} and {{name}}.\n', templates=MUSTACHE), out, tmpl=bound)
    z = zipfile.ZipFile(out)
    doc = z.read('word/document.xml').decode()
    assert doc.count('<w:dataBinding ') == 3 and doc.count('/ns0:fields[1]/ns0:sal[1]') == 2
    cx = z.read('customXml/item1.xml').decode()
    assert cx.count('<ns0:sal/>') == 1 and '<ns0:name/>' in cx   # deduped inventory, one element per variable
    assert 'customXmlProperties' in z.read('[Content_Types].xml').decode()
    assert 'customXml' in z.read('word/_rels/document.xml.rels').decode()
    teq(fast_checks(out), 'valid')
