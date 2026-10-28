import base64, pytest, subprocess, sys, zipfile
from pathlib import Path
from lxml import etree
from oxml import Document, e, w, namespaces

from mdhtml import Document as MdDocument, md2mdhtml
from mdhtml.mustache import MUSTACHE
from mdhtml2docx.mdhtml2docx import mdhtml2docx, mustache_fields
from mdhtml2docx.styles import ref_path, theme_ref
from mdhtml2docx.wml import NS, W

IMAGE = Path(__file__).parent/'fixtures'/'tiny.png'


def pandoc(path):
    return subprocess.run(['pandoc', '-f', 'docx', '-t', 'markdown', '--wrap=none', str(path)],
        capture_output=True, text=True, check=True, timeout=60).stdout

def xmlpart(path, part='document'):
    with zipfile.ZipFile(path) as z: return etree.fromstring(z.read(f'word/{part}.xml'))

def canonical(node): return etree.tostring(node, method='c14n', exclusive=True)
def fields(doc): return [f.instruction.strip() for f in doc.main.xml.elements(w.SimpleField)]

def checked(path):
    with zipfile.ZipFile(path) as z: assert z.testzip() is None
    doc = Document.open(path)
    assert not (issues := doc.validate()['issues']), issues
    defined = {s.id for s in doc.styles}
    for part in (doc.main, doc.part('FootnotesPart')):
        if part:
            for cls in (w.ParagraphStyleId, w.RunStyle, w.TableStyle):
                assert {s.val for s in part.xml.elements(cls)} <= defined
    return doc

def convert(src, out, warnings=(), **kwargs):
    assert mdhtml2docx(src, out, **kwargs) == list(warnings)
    return checked(out)


def test_conversion_without_lxml(tmp_path):
    code = ("import sys\nsys.modules['lxml'] = None\nfrom mdhtml2docx import mdhtml2docx\n"
        "mdhtml2docx('<p>Hello</p>', sys.argv[1])")
    subprocess.run([sys.executable, '-c', code, str(tmp_path/'t.docx')], check=True, timeout=60)


def test_rich_publication(tmp_path):
    "A publication mixes prose, nested structures, images, notes and math; another reader recovers its content."
    out = tmp_path/'publication.docx'
    src = ('<p>Hello <em>world</em> with <strong>bold</strong> and <strong><em>both</em></strong>.</p>'
        '<div>Second para: café – 東京…</div><h1 id="top">Title</h1><h2>Sub <em>title</em></h2>'
        '<p>Call <code>f(x)</code> or see <a href="https://fast.ai/">fast.ai</a> and <a href="#top">the title</a>.</p>'
        '<blockquote><p>Quoted line.</p><blockquote>Deeper.</blockquote></blockquote>'
        '<pre><code class="language-python">def f(x):\n    return "café 東京"\n</code></pre><pre><code>plain code</code></pre>'
        '<ul><li>one</li><li>two<ul><li>deep</li></ul></li></ul>'
        '<ol><li>first</li><li>second</li></ol><ol start="5"><li>fifth</li></ol>'
        '<ul class="task-list"><li><input type="checkbox" checked> done</li><li><input type="checkbox"> todo</li></ul>'
        '<table colwidths="10em 2fr 1fr"><thead><tr><th>Feature</th><th>Status</th><th>Notes</th></tr></thead>'
        '<tbody><tr><td>Tables</td><td>ready</td><td>ok</td></tr></tbody></table>'
        '<table keep-rows="false"><tr><td rowspan="2">a</td></tr><tr><td colspan="2">bc</td></tr></table>'
        '<table custom-style="Borderless Table" keep-rows="true"><thead><tr><th></th><th></th></tr></thead>'
        '<tbody><tr><td>A:</td><td>one</td></tr><tr><td>B:</td><td>two</td></tr></tbody></table>'
        '<p>H<sub>2</sub>O, E=mc<sup>2</sup>, <del>gone</del>, <mark>hot</mark>, <u>under</u>.</p><hr>'
        '<dl><dt>term</dt><dd>definition here</dd><dd><p>continued definition</p></dd></dl>'
        f'<p><img src="{IMAGE}" alt="tiny pic"></p>'
        '<p>Notes<sup id="fnref-a"><a href="#fn-a" class="footnote-ref">1</a></sup>'
        '<sup><a href="#fn-b" class="footnote-ref">2</a></sup>.</p>'
        '<section class="footnotes"><ol><li id="fn-a"><table><tr><td><div></div></td>'
        '<td><table><tr><td>nested note</td></tr></table></td></tr></table>'
        '<p>The note text with a <a href="https://example.org/">link</a>.</p>'
        '<a href="#fnref-a" class="footnote-backref">↩</a></li>'
        '<li id="fn-b"><script type="application/vnd.mdhtml.raw" data-format="docx">'
        '<w:p><w:r><w:t>raw note</w:t></w:r></w:p></script></li></ol></section>'
        '<p>Euler: <span class="math inline">e^{i\\pi} + 1 = 0</span> inline.</p>'
        '<div class="math display">E = mc^2</div><p>Signature follows.<br type="page">Next page.<br>plain break</p>')
    src += md2mdhtml('::: {.details}\n## Details label\n\nDetails body\n:::\n')
    for state in ('fixed', 'open', 'closed'):
        src += (f'<div data-panel data-callout="warning" data-disclosure="{state}"><header>The <em>label</em></header>'
            'Body text<p>More body</p><h3>Actual heading</h3></div>')
    doc = convert(src, out)
    tree, text = doc.main.xml, pandoc(out)
    for fragment in ('Hello *world* with **bold** and ***both***.', 'café', '# Sub *title*', '[the title](#top)',
        '[fast.ai](https://fast.ai/)', '> Quoted line.', '    def f(x):', 'plain code', 'H~2~O', 'E=mc^2^', '~~gone~~',
        '[under]{.underline}', 'definition here', 'continued definition', '![tiny pic](media/', 'nested note', 'raw note',
        'The note text', r'mc\hat{}2$$', '**Details label**', 'Details body', '## Actual heading'):
        assert fragment in text
    assert '↩' not in text and '# Details label' not in text
    assert not any(line.startswith('#') and 'label' in line for line in text.splitlines())
    assert text.count('Body text') == text.count('More body') == text.count('Actual heading') == 3
    assert doc.story.text.count('The label') == 3
    assert [s.val for s in tree.elements(w.ParagraphStyleId)][:2] == ['FirstParagraph', 'BodyText']
    assert xmlpart(out).xpath('//w:p[w:r/w:t="Deeper."]/w:pPr/w:ind/@w:left', namespaces=NS) == ['1440']
    used = {s.val for s in tree.elements(w.RunStyle)}
    assert any(s.startswith('HlKeyword') for s in used) and any(s.startswith('HlString') for s in used)
    # Pandoc merges adjacent lists; their actual numbering instances must still restart independently.
    first, second, fifth = [n.val for n in list(tree.elements(w.NumberingId))[3:6]]
    assert first == second and first != fifth
    for nid, start in ((first, 1), (fifth, 5)):
        level, = (l for l in doc.numbering[nid].element.elements(w.LevelOverride) if l.level_index == 0)
        assert next(level.elements(w.StartOverrideNumberingValue)).val == start
    lines = [' '.join(line.split()) for line in text.splitlines()]
    assert all(s in lines for s in ('- one', '- deep', '1. first', '2. second', '- [x] done', '- [ ] todo'))
    table, merged, blank = tree.elements(w.Table)
    assert [c.width for c in table.elements(w.GridColumn)] == ['2200', '4773', '2387']
    assert len(list(table.elements(w.TableHeader))) == 1 and not list(blank.elements(w.TableHeader))
    assert len(list(blank.elements(w.TableRow))) == 2
    assert [s.val for s in tree.elements(w.TableStyle)] == ['TableGrid', 'TableGrid', 'BorderlessTable']
    assert [m.val for m in merged.elements(w.VerticalMerge)] == ['restart', None]
    assert next(merged.elements(w.GridSpan)).val == 2 and len(list(merged.elements(w.GridColumn))) == 3
    for tbl, keep in ((table, 'on'), (merged, 'off'), (blank, 'on')):
        assert [flag.val for flag in tbl.elements(w.CantSplit)] == [keep] * len(list(tbl.elements(w.TableRow)))
    signature, = (p for p in tree.elements(w.Paragraph) if 'Signature follows.' in [t.text for t in p.elements(w.Text)])
    assert [b.type for b in signature.elements(w.Break)] == ['page', None]
    extent, = tree.elements(namespaces['wp'].Extent)
    assert (extent.cx, extent.cy) == (38100, 19050)
    notes = xmlpart(out, 'footnotes')
    for note in notes.findall('w:footnote[@w:id="1"]', NS) + notes.findall('w:footnote[@w:id="2"]', NS):
        assert note[0].tag == f'{{{W}}}p' and note[0].find('w:r/w:footnoteRef', NS) is not None
        assert len(note.findall('.//w:footnoteRef', NS)) == 1
        assert all(cell[-1].tag == f'{{{W}}}p' for cell in note.findall('.//w:tc', NS))
    assert xmlpart(out).xpath('//m:oMath/m:r/m:t/text()', namespaces=NS) == ['e^{i\\pi} + 1 = 0', 'E = mc^2']


def test_html_import_and_raw_recovery(tmp_path):
    "Shared HTML import preserves semantic content; malformed raw XML warns without losing surrounding prose."
    out = tmp_path/'repaired.docx'
    b64 = base64.b64encode(b'<w:r><w:t>B64</w:t></w:r>').decode()
    source = MdDocument.from_mdhtml('<p custom-style="Centered">Before <div>middle</div> after.</p><span zoop:33="x">tail</span><!-- a -- comment -->'
        '<input type="date"><template><p>hidden</p></template><p><span custom-style="Fancy">styled bit</span></p>'
        '<script type="application/vnd.mdhtml.raw" data-format="docx"><w:p><w:r><w:br w:type="page"/></w:r></w:p></script>'
        '<p>a <script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="html">'
        '&lt;w:r&gt;&lt;w:t&gt;RAW&lt;/w:t&gt;&lt;/w:r&gt;</script> b</p>'
        f'<p><script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="base64">{b64}</script></p>'
        '<script type="application/vnd.mdhtml.raw" data-format="latex">\\newpage</script>')
    doc = convert(source, out, warnings=("unhandled inline <input type='date'>; dropped",
        "custom style 'Fancy' not in reference doc; stub injected"))
    assert [t.text for t in doc.main.xml.elements(w.Text)] == ['Before', 'middle', 'after.', 'tail', 'styled bit', 'a ', 'RAW', ' b', 'B64']
    assert next(doc.main.xml.elements(w.ParagraphStyleId)).val == 'Centered' and doc.styles['Fancy'].name == 'Fancy'
    assert [b.type for b in doc.main.xml.elements(w.Break)] == ['page']
    text = pandoc(out)
    assert 'RAW' in text and 'hidden' not in text
    warnings = mdhtml2docx('<p>Still here.</p><script type="application/vnd.mdhtml.raw" data-format="docx"><w:oops></script>'
        '<script type="application/vnd.mdhtml.raw" data-format="docx" data-encoding="base64">!!!</script>', out)
    assert len(warnings) == 2 and any('malformed docx raw payload' in w for w in warnings) and any('malformed base64' in w for w in warnings)
    assert checked(out).story.text == 'Still here.'
    assert not convert('<template><p>still hidden</p></template>', out).main.xml.count(w.Paragraph)


@pytest.mark.parametrize('scheme,level2', [('legal', '(%3)'), ('decimal', '%2.%3.')])
def test_numbered_agreement_and_caption_references(tmp_path, scheme, level2):
    "Heading, definition and caption targets remain distinct; invalid references never replace the last good document."
    out = tmp_path/'agreement.docx'
    src = ('<p><a href="#fig-plot" data-ref="text"></a> and <a href="#tbl-r" data-ref="text"></a>.</p>'
        '<p><span id="fig_plot_t">Neighbor</span></p>'
        '<h1 id="sec-intro">Intro</h1><h2 id="sec-pay">Payment terms</h2><h3>Confidential Information</h3>'
        '<dl><dt id="def-term"><strong>"Term"</strong></dt><dt id="def-ap">Agreement Period</dt><dd>the deal period</dd></dl>'
        '<p>The <span id="def-eff">Effective Date</span> is set.</p>'
        f'<figure id="fig-plot"><img src="{IMAGE}"><figcaption>A plot</figcaption></figure>'
        '<table id="tbl-r"><caption>Results <em>now</em></caption><tr><td>1</td></tr></table>'
        '<div class="math display" id="eq-force">F = ma</div><div class="math display">x = y</div>'
        '<p>See <a href="#sec-pay" data-ref=""></a>, <a href="#sec-intro" data-ref="bare text"></a>, '
        '<a href="#sec-pay" data-ref="">Clause</a>, <a href="#sec-pay" data-ref="bare page"></a> and '
        '<span data-refs=""><a href="#sec-intro" data-ref=""></a><a href="#sec-pay" data-ref=""></a></span>. '
        '<a href="#def-term" data-ref=""></a><a href="#def-eff" data-ref=""></a><a href="#def-ap" data-ref="page"></a>; '
        '<a href="#fig-plot" data-ref=""></a><a href="#tbl-r" data-ref="bare"></a> '
        '<span data-refs=""><a href="#fig-plot" data-ref=""></a><a href="#tbl-r" data-ref=""></a></span>. '
        '<a href="#eq-force" data-ref=""></a>, <a href="#eq-force" data-ref="bare"></a>, '
        'and <a href="#eq-force" data-ref="page"></a>.</p>')
    doc = convert(src, out, number_headings=scheme, reftypes=dict(fig=('Item', 'Items'), tbl=('Item', 'Items')))
    expected = {r'REF sec_pay \w \h', r'REF sec_intro \h', r'PAGEREF sec_pay \h', r'REF def_term \h', r'REF def_eff \h',
        r'PAGEREF def_ap \h', r'SEQ Figure \* ARABIC', r'SEQ Table \* ARABIC', r'REF fig_plot_n \h', r'REF tbl_r_n \h',
        r'SEQ Equation \* ARABIC', r'REF eq_force_n \h', r'PAGEREF eq_force \h', r'REF fig_plot_t \h', r'REF tbl_r_t \h'}
    assert set(fields(doc)) >= expected and r'REF fig_plot \w \h' not in fields(doc)
    for fragment in ('Section ', 'Sections ', 'Clause ', 'Effective Date', 'A plot', 'Results ', 'Item ', 'Equation '): assert fragment in doc.story.text
    assert 'Figures' not in doc.story.text
    assert {b.name for b in doc.main.xml.elements(w.BookmarkStart)} >= {'def_term', 'def_eff', 'fig_plot', 'tbl_r_n', 'eq_force', 'eq_force_n'}
    assert fields(doc).count(r'SEQ Equation \* ARABIC') == 1 and not any(f.startswith('SEQ Item ') for f in fields(doc))
    names = [b.name for b in doc.main.xml.elements(w.BookmarkStart)]
    assert len(names) == len(set(names)) and {'fig_plot_t', 'fig_plot_t_', 'tbl_r_t'} <= set(names)
    assert xmlpart(out).xpath('//w:bookmarkStart[@w:name="fig_plot_t"]/following-sibling::w:r[1]/w:t/text()', namespaces=NS) == ['A plot']
    assert next(doc.main.xml.elements(namespaces['wp'].DocProperties)).description == 'A plot'
    assert doc.part('DocumentSettingsPart').xml.count(w.UpdateFieldsOnOpen) == 1
    nids = []
    for level, sid in enumerate(('Title', 'Heading1', 'Heading2')):
        style = doc.styles[sid].element
        assert next(style.elements(w.NumberingLevelReference)).val == level
        nids.append(next(style.elements(w.NumberingId)).val)
    assert len(set(nids)) == 1
    definition = doc.numbering[nids[0]].definition
    assert [v.val for v in definition.elements(w.LevelText)][:3] == ['', '%2.', level2]
    assert next(next(definition.elements(w.Level)).elements(w.LevelSuffix)).val == 'nothing'
    text = pandoc(out)
    assert 'A plot' in text and 'Results' in text
    saved = out.read_bytes()
    for invalid, message in (
        ('<p><a href="#nope" data-ref=""></a></p>', '#nope'),
        ('<h2 id="z-a">Unknown type</h2><p><a href="#z-a" data-ref=""></a></p>', "reference type 'z'"),
        ('<p><a href="#sec-pay" data-ref="page text"></a></p>', 'conflicting data-ref'),
        ('<div id="d-x"><p>hi</p></div><p><a href="#d-x" data-ref=""></a></p>', '#d-x'),
        ('<div class="math display" id="eq-a">x</div><p><a href="#eq-a" data-ref="text"></a></p>', 'equation source')):
        with pytest.raises(ValueError, match=message): mdhtml2docx(invalid, out)
        assert out.read_bytes() == saved


def test_reference_composition_and_reuse(tmp_path):
    "Generate a reference, merge contributors, relocate its package, then reuse it without corrupting retained content."
    from tools.createref import build
    from fastpylight import theme_colors
    built, reference, relocated, out = [tmp_path/f'{name}.docx' for name in ('built', 'reference', 'relocated', 'out')]
    build(Path(__file__).parents[1]/'_data'/'empty.docx', built)
    doc = checked(built)
    assert doc.part('FooterPart') is not None and doc.main.xml.count(w.FooterReference) == 1
    assert next(doc.styles['Quote'].element.elements(w.Indentation)).left == '720'
    for name in ('HeaderPart', 'FooterPart'): assert next(doc.part(name).xml.elements(w.BiDiVisual)).val == 'off'
    refs = []
    for i, (name, fmt) in enumerate((('Fancy', 'upperRoman'), ('Other', 'lowerLetter'), ('Fancy', 'lowerRoman'))):
        path = tmp_path/f'extra{i}.xml'
        path.write_text(f'<w:styles xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
            f'<w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            f'<w:style w:type="paragraph" w:styleId="{name}"><w:name w:val="{name}"/>'
            '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>')
        refs.append(path)
    content = f'<a href="https://example.org/">link</a><img src="{IMAGE}">'
    src = ('<h1 id="sec-a">Before</h1><h2>Terms</h2><p custom-style="Fancy">Before '
        f'{content}<sup><a href="#fn-a" class="footnote-ref">1</a></sup></p><p custom-style="Other">Other</p>'
        '<ul><li>listed</li></ul><pre><code class="language-python">def f(x):\n    return x\n</code></pre>'
        '<p><a href="#sec-a" data-ref=""></a></p>'
        f'<section class="footnotes"><ol><li id="fn-a"><p>Before {content}</p></li></ol></section>')
    doc = convert(src, reference, reference=[built, 'dracula', *refs], number_headings='legal')
    for cls, attr in ((w.NumberingInstance, 'number_id'), (w.AbstractNum, 'abstract_number_id')):
        ids = [getattr(n, attr) for n in doc.part('NumberingDefinitionsPart').xml.elements(cls)]
        assert len(ids) == len(set(ids))
    for name, fmt in (('Fancy', 'lowerRoman'), ('Other', 'lowerLetter')):
        nid = next(doc.styles[name].element.elements(w.NumberingId)).val
        assert next(doc.numbering[nid].definition.elements(w.NumberingFormat)).val == fmt
    assert [s.id for s in doc.styles].count('SourceCode') == 1
    fills = {s.fill for s in doc.part('StyleDefinitionsPart').xml.elements(w.Shading)}
    assert theme_colors('dracula')['normal']['bg'].lstrip('#').upper() in fills and 'F5F5F5' not in fills
    assert any(s.val.startswith('HlKeyword') for s in doc.main.xml.elements(w.RunStyle))
    theme = theme_ref('dracula', tmp_path/'dracula.docx')
    checked(theme)
    convert(src, out, reference=[built, theme, *refs], number_headings='legal')
    assert canonical(xmlpart(out, 'styles')) == canonical(xmlpart(reference, 'styles'))
    doc.part('NumberingDefinitionsPart').xml.root.append_xml(b'\n<!-- retain cleanup marker -->\n')
    doc.part('NumberingDefinitionsPart').xml.root(e.numIdMacAtCleanup(val=42))
    doc.save(reference)
    original = xmlpart(reference, 'numbering')
    renames = {'word/': 'content/', 'document.xml': 'main.xml', 'styles.xml': 'styles-custom.xml',
        'numbering.xml': 'numbering-custom.xml', 'settings.xml': 'settings-custom.xml', 'footnotes.xml': 'footnotes-custom.xml'}
    def renamed(s):
        for old, new in renames.items(): s = s.replace(old, new)
        return s
    with zipfile.ZipFile(reference) as zi, zipfile.ZipFile(relocated, 'w', zipfile.ZIP_DEFLATED) as zo:
        for name in zi.namelist():
            data = zi.read(name)
            if name.endswith('.rels') or name == '[Content_Types].xml': data = renamed(data.decode('utf-8-sig')).encode()
            zo.writestr(renamed(name), data)
    doc = convert(src.replace('Before', 'After'), out, reference=relocated, number_headings='decimal')
    package = doc.package
    nums = etree.fromstring(doc.part('NumberingDefinitionsPart').xml.bytes())
    for tag, added in (('abstractNum', 2), ('num', 1)):
        assert len(nums.findall(f'w:{tag}', NS)) == len(original.findall(f'w:{tag}', NS)) + added
        for old in original.findall(f'w:{tag}', NS):
            nid = old.get(f'{{{W}}}{tag}Id')
            new, = nums.findall(f'w:{tag}[@w:{tag}Id="{nid}"]', NS)
            assert canonical(new) == canonical(old)
    assert nums[-1].tag == f'{{{W}}}numIdMacAtCleanup' and nums[-1].get(f'{{{W}}}val') == '42'
    assert r'REF sec_a \w \h' in fields(doc) and 'After' in pandoc(out)
    with zipfile.ZipFile(relocated) as zi, zipfile.ZipFile(out) as zo:
        assert set(zi.namelist()) <= set(zo.namelist()) and 'word/document.xml' not in zo.namelist()
        for name in zi.namelist():
            if '/media/' in name or any(s in name for s in ('styles-custom', 'header', 'footer')):
                assert zi.read(name) == zo.read(name)
        for part in (doc.main, doc.part('FootnotesPart')):
            assert 'After' in ''.join(t.text for t in part.xml.elements(w.Text))
            targets = {r['id']: r['target'] for r in package.relationships(part.uri)}
            assert targets[next(part.xml.elements(w.Hyperlink)).id] == 'https://example.org/'
            target = package.relationship_part(part.uri, next(part.xml.elements(namespaces['a'].Blip)).embed).lstrip('/')
            assert target not in zi.namelist() and zo.read(target) == IMAGE.read_bytes()
    plain = convert('<pre><code class="language-python">def f(x):\n    return x\n</code></pre>', out, reference=ref_path())
    assert not any(s.val.startswith('Hl') for s in plain.main.xml.elements(w.RunStyle))


def test_template_lifecycle(tmp_path):
    "The same form can drop variables, expose merge fields or controls, then be reused with a refreshed bound inventory."
    out, again = tmp_path/'form.docx', tmp_path/'again.docx'
    src = md2mdhtml('Pay {{sal}} to {{sal}} and {{name}}.\n\n{{#opt}}\n\nGranted.\n\n{{/opt}}\n', templates=MUSTACHE)
    doc = convert(src, out)
    assert 'sal' not in doc.story.text and '{{' not in doc.story.text and '«#opt»' in doc.story.text and not fields(doc)
    doc = convert(src, out, tmpl=mustache_fields)
    assert fields(doc).count('MERGEFIELD sal') == 2 and 'MERGEFIELD name' in fields(doc)
    assert '«#opt»' in doc.story.text and '«/opt»' in doc.story.text
    doc = convert(src, out, tmpl=lambda node: ('control', node['name']))
    assert doc.main.xml.count(w.SdtRun) == doc.main.xml.count(w.ShowingPlaceholder) == 3
    assert {t.val for t in doc.main.xml.elements(w.Tag)} == {'sal', 'name'}
    assert {s.val for s in doc.main.xml.elements(w.RunStyle)} == {'PlaceholderText'}
    assert next(doc.styles['PlaceholderText'].element.elements(w.Color)).val == '808080'
    bound = lambda node: ('bound', node['name'])
    doc = convert(src, out, tmpl=bound)
    assert [b.x_path for b in doc.main.xml.elements(w.DataBinding)].count('/ns0:fields[1]/ns0:sal[1]') == 2
    doc.set_custom_xml('{00000000-0000-0000-0000-000000000001}', b'<unrelated/>')
    doc.save(out)
    convert(md2mdhtml('Now {{name}}.\n', templates=MUSTACHE), again, reference=out, tmpl=bound)
    with zipfile.ZipFile(out) as z, zipfile.ZipFile(again) as za:
        assert z.read('customXml/item1.xml').count(b'<ns0:sal/>') == 1
        assert b'customXmlProperties' in z.read('[Content_Types].xml') and b'customXml' in z.read('word/_rels/document.xml.rels')
        names = [n for n in z.namelist() if n.startswith('customXml/')]
        assert sorted(names) == sorted(n for n in za.namelist() if n.startswith('customXml/'))
        assert b'<ns0:name/>' in za.read('customXml/item1.xml') and b'<ns0:sal/>' not in za.read('customXml/item1.xml')
        for name in names:
            if name != 'customXml/item1.xml': assert za.read(name) == z.read(name)
