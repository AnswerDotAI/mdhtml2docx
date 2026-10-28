"Write semantic mdhtml Documents as docx with live Word fields and styles."
import re
from pathlib import Path
from mdhtml.export import HeadingNums, import_mdhtml
from oxml import Document, Tree, bookmark_name, emu, field, w
from .styles import STYLE_MAP, style_id, theme_styles
from .styles import ref_path as _refpath
from .wml import *
from .hilite import segments, tokenize

__all__ = ['mdhtml2docx', 'mustache_fields']

m = E('m', attr_ns='m')
# Pandoc reads drawing prefixes from part roots, not just their local XML scope.
_partxml = E('w', attr_ns='w', ns=NS)

def _sid(key): return style_id(STYLE_MAP[key])

PARTS = dict(styles='StyleDefinitionsPart', numbering='NumberingDefinitionsPart')

HEADING_STYLE_IDS = ['Title'] + [f'Heading{i}' for i in range(1, 6)]   # h1-h6, one per numbering level: the Title carries level 0

def _classes(node): return node['attrs'].get('class', '').split()


class Converter:
    def __init__(self, reference=None, base=None, reftypes=None, number_headings=None, tmpl=None):
        self.tmpl = tmpl
        if reference is None: reference = [_refpath()] + (['github_light'] if tokenize else [])
        elif not isinstance(reference, (list, tuple)): reference = [reference]
        self.doc = Document.open(reference[0] or _refpath())
        self.package = self.doc.package
        self._source_uri = self.package.main_part
        self.base = Path(base or '.')
        self.warnings = []
        self.bq = 0          # blockquote nesting depth
        self._bkid = 0       # bookmark id counter
        self.nums = []       # (numId, abstractNumId, start) per list instance
        self.fndefs = {}     # footnote label -> semantic blocks
        self.fnids = {}      # endnote id -> Footnote
        self.stubs = {}      # undefined custom-style name -> (kind, styleId)
        self.first = True    # next body paragraph is a 'First Paragraph' (doc start; reset after FIRST_AFTER blocks)
        self._bknames = {}   # (target id, part) -> Word-legal bookmark name
        self.has_fields = False
        self.has_controls = False
        self.bound = []      # distinct bound-control names, in first-appearance order
        self.reftypes, self.number_headings = reftypes, number_headings
        self.figure_alt = None
        self.sectpr = self.doc.main.xml.root.child(w.Body).child(w.SectionProperties)
        pg, mar = self.sectpr.child(w.PageSize), self.sectpr.child(w.PageMargin)
        self.content_w = int(pg.attribute(W, 'w')) - int(mar.attribute(W, 'left')) - int(mar.attribute(W, 'right'))
        styles = self.doc.part('StyleDefinitionsPart')
        if styles is None: raise ValueError('reference doc lacks a styles relationship')
        self.styles = styles.xml
        numbering = part.xml if (part := self.doc.part('NumberingDefinitionsPart')) else None
        used = [int(v) for e in ([] if numbering is None else numbering.root.children)
            for a in ('numId', 'abstractNumId') if (v := e.attribute(W, a)) is not None]
        self._numid = max(used, default=0)
        for ref in reference[1:]: self._merge_styles(ref)
        self._absbase = self._numid + 1
        self._numid += 3     # reserve bullet, ordered-list, and heading abstract IDs after all contributors
        self.refstyles = {s.child(w.StyleName).val.lower(): s.style_id for s in self.styles.elements(w.Style)}
        if missing := [n for n in STYLE_MAP.values() if n.lower() not in self.refstyles]:
            raise ValueError(f'reference doc lacks dialect styles (map/template drift?): {missing}')
        self.hlstyles = {n.removeprefix('hl ').replace(' ', '.'): sid for n, sid in self.refstyles.items() if n.startswith('hl ')}
        self.scheme = self.headnum = None

    def _merge_styles(self, ref):
        "Merge a contributor's w:style elements, later-wins on style id or name"
        def keys(e):
            nm = e.child(w.StyleName)
            return {(e.attribute(W, 'styleId') or '').lower(), '' if nm is None else nm.attribute(W, 'val').lower()} - {''}
        if str(ref).endswith('.docx'):
            styles = Document.open(ref).part('StyleDefinitionsPart')
            if styles is None: raise ValueError(f'style contributor lacks a styles relationship: {ref}')
            new = list(styles.xml.elements(w.Style))
        elif str(ref).endswith('.xml'): new = self._xml_contrib(ref)
        else: new = list(Tree(e.styles(theme_styles(ref)).bytes()).elements(w.Style))
        existing = {key: s for s in self.styles.elements(w.Style) for key in keys(s)}
        for s in new:
            ks = keys(s)
            old_styles = {existing[k].node_id: existing[k] for k in ks if k in existing}
            for old in old_styles.values():
                for key in keys(old): existing.pop(key, None)
                old.delete()
            copied = s.copy_to(self.styles.root)
            for key in ks: existing[key] = copied

    def _xml_contrib(self, ref):
        "Remap one XML contributor in its own ID scope, then adopt its numbering and return its styles."
        tree = Tree(Path(ref).read_bytes())
        for tag in ('abstractNum', 'num'):
            attr, mapping = f'{tag}Id', {}
            for e in tree.root.children:
                if e.qname != (W, tag): continue
                self._numid += 1
                mapping[e.attribute(W, attr)] = str(self._numid)
                e.set_attribute(W, attr, str(self._numid))
            for e in tree.elements():
                if e.qname == (W, attr) and (val := e.attribute(W, 'val')) in mapping: e.set_attribute(W, 'val', mapping[val])
        for e in tree.root.children:
            if e.qname not in ((W, 'abstractNum'), (W, 'num')): continue
            self._part('numbering').xml.root(e)
        return list(tree.elements(w.Style))

    def hlsid(self, scope):
        "Hl* style id for a dotted scope: exact, else progressively shorter prefixes (tree-sitter resolution)"
        parts = (scope or '').split('.')
        while parts:
            if sid := self.hlstyles.get('.'.join(parts)): return sid
            parts.pop()

    def _part(self, kind): return self.doc.part(PARTS[kind], create=True)

    def warn(self, msg): self.warnings.append(msg)

    # ---- inline level -------------------------------------------------------
    def rpr(self, fmt):
        "w:rPr for a formatting context dict, in CT_RPr child order; None if empty"
        props = e.rPr()
        if s := fmt.get('rstyle'): props(e.rStyle(val=s))
        if fmt.get('b'): props(e.b(), e.bCs())
        if fmt.get('i'): props(e.i(), e.iCs())
        if fmt.get('strike'): props(e.strike())
        if fmt.get('mark'): props(e.highlight(val='yellow'))
        if fmt.get('u'): props(e.u(val='single'))
        if v := fmt.get('vert'): props(e.vertAlign(val=v))
        return props if props.children else None

    def text_runs(self, text, fmt):
        "Runs for a text node; newlines collapse to spaces (pre-context text never comes here)"
        text = re.sub(r'\s+', ' ', text)
        if not text: return []
        t = e.t(text, xml__space='preserve' if text != text.strip() else None)
        return [e.r(self.rpr(fmt), t)]

    def field(self, instr, text, fmt):
        "A simple Word field with cached text; every field requests an update on open."
        self.has_fields = True
        return field(instr, text, self.rpr(fmt))

    def link(self, node, fmt):
        runs = self.runs(node['children'], fmt | {'rstyle': _sid('hyperlink')})
        href = node['url']
        if href.startswith('#'): return [e.hyperlink(runs, anchor=self.bkname(href[1:]))]
        return self.external_link(href, runs)


    def external_link(self, href, runs): return [self.doc.hyperlinks.link(href, runs, part_uri=self._source_uri)]

    REFSWITCH = dict(full=r'\w', rel=r'\r', leaf=r'\n', text='', page=None)

    def bkname(self, id, part=''):
        "Reserve a target or generated part name, including before its definition."
        key = id, part
        if key not in self._bknames:
            nm = bookmark_name(id) if not part else self.bkname(id) + dict(number='_n', text='_t')[part]
            while nm in self._bknames.values(): nm += '_'
            self._bknames[key] = nm
        return self._bknames[key]

    def references(self, items, fmt):
        out = []
        for part in self.index.plan(items):
            out += self.text_runs(part['separator'] + part['prefix'], fmt)
            out += self.ref_fld(part, fmt)
        return out

    def ref_fld(self, part, fmt):
        "Map a shared reference plan to a live Word REF or PAGEREF field."
        target, variant = part['target'], part['variant']
        name, kind = self.bkname(target), part['kind']
        if kind == 'equation' and variant == 'text': raise ValueError('DOCX does not support ref=text to equation source')
        if variant == 'page': instr, cached = rf'PAGEREF {name} \h', '#'
        elif variant == 'text' and kind in ('figure', 'table'): instr, cached = rf'REF {self.bkname(target, "text")} \h', part['text']
        elif kind in ('figure', 'table', 'equation'): instr, cached = rf'REF {self.bkname(target, "number")} \h', '#'
        elif kind == 'text': instr, cached = rf'REF {name} \h', part['text']
        else:
            switch = self.REFSWITCH[variant]
            instr = rf'REF {name} {switch} \h' if switch else rf'REF {name} \h'
            cached = part['text'] if variant == 'text' else '#'
        return [self.field(instr, cached, fmt)]


    def custom_style(self, el, kind):
        "Style id for an explicit custom-style attr (stubbed + warned if undefined), else a class matching a reference style name"
        if cs := el['attrs'].get('custom-style'):
            if cs.lower() in self.refstyles: return self.refstyles[cs.lower()]
            if cs not in self.stubs:
                self.stubs[cs] = (kind, re.sub(r'\W', '', cs) or f'Custom{len(self.stubs)}')
                self.warn(f'custom style {cs!r} not in reference doc; stub injected')
            return self.stubs[cs][1]
        return next((self.refstyles[c.lower()] for c in _classes(el) if c.lower() in self.refstyles), None)

    def omath(self, tex): return m.oMath(m.r(m.t(tex, xml__space='preserve')))

    def fnref(self, key):
        if key not in self.fndefs:
            self.warn(f'footnote reference #{key} has no definition; dropped')
            return []
        if key not in self.fnids: self.fnids[key] = self.doc.footnotes.create(self.fn_blocks(self.fndefs[key]))
        return [self.fnids[key].reference()]

    def fn_blocks(self, blocks):
        saved = self._source_uri, self.first
        self._source_uri = self.doc.part('FootnotesPart', create=True).uri
        try: return self.block_nodes(blocks, 'footnotetext')
        finally: self._source_uri, self.first = saved


    def image(self, el, fmt, alt=None):
        "Embed a local image (dimensions sniffed, width/height px attrs override); remote srcs degrade to a link"
        src = el['url']
        if alt is None: alt = el['alt'] or self.figure_alt or src
        if re.match(r'[a-z][a-z0-9+.-]*://', src):
            self.warn(f'remote image not embedded: {src}')
            return self.external_link(src, self.text_runs(alt, fmt | {'rstyle': _sid('hyperlink')}))
        try: data = (self.base/src).read_bytes()
        except OSError:
            self.warn(f'image not found: {src}; alt text emitted')
            return self.text_runs(alt, fmt)
        w_, h_ = el['attrs'].get('width'), el['attrs'].get('height')
        try:
            drawing = self.package.part(self._source_uri).add_image(data, width=emu(px=float(w_)) if w_ else None,
                height=emu(px=float(h_)) if h_ else None, description=alt)
        except ValueError as ex:
            self.warn(f'{src}: {ex}; alt text emitted')
            return self.text_runs(alt, fmt)
        return [e.r(drawing)]

    INLINE_FMT = dict(emph={'i': True}, strong={'b': True}, strike={'strike': True}, highlight={'mark': True},
        superscript={'vert': 'superscript'}, subscript={'vert': 'subscript'}, u={'u': True})

    def inline(self, node, fmt):
        kind = node['type']
        if kind == 'text': return self.text_runs(node['text'], fmt)
        if kind == 'soft_break': return self.text_runs(' ', fmt)
        if kind == 'hard_break': return [e.r(self.rpr(fmt), e.br())]
        tag = node.get('tag') if kind == 'element_inline' else kind
        if tag in self.INLINE_FMT: out = self.runs(node['children'], fmt | self.INLINE_FMT[tag])
        elif kind == 'code': out = self.text_runs(node['text'], fmt | {'rstyle': _sid('codeinline')})
        elif kind == 'link': out = self.link(node, fmt)
        elif kind == 'reference': out = self.references(node['items'], fmt)
        elif kind == 'footnote_ref': out = self.fnref(node['label'])
        elif kind == 'span':
            sid = self.custom_style(node, 'character')
            out = self.runs(node['children'], fmt | {'rstyle': sid} if sid else fmt)
        elif kind == 'math_inline': out = [self.omath(node['tex'])]
        elif kind == 'image': out = self.image(node, fmt)
        elif tag == 'br': out = [e.r(self.rpr(fmt), e.br(type='page' if node['attrs'].get('type') == 'page' else None))]
        elif tag == 'input':
            if node['attrs'].get('type') != 'checkbox':
                self.warn(f'unhandled inline <input type={node["attrs"].get("type")!r}>; dropped')
                return []
            out = self.text_runs('☒ ' if 'checked' in node['attrs'] else '☐ ', fmt)
        elif kind == 'raw_inline': out = self.rawxml(node)
        elif kind == 'operation': out = self.tmpl_runs(node, fmt)
        else: out = self.runs(node.get('children', []), fmt)
        return self.bookmark(node, out)

    def runs(self, nodes, fmt): return [run for node in nodes for run in self.inline(node, fmt)]


    # ---- block level --------------------------------------------------------
    def para(self, runs, style='body', extra=None, sid=None):
        "A w:p with `style` (STYLE_MAP key, or `sid` style-id override) and optional extra pPr children (schema order!)"
        ppr = e.pPr(e.pStyle(val=sid or _sid(style)), extra)
        return Tree(e.p(ppr, runs).bytes()).root

    def prose(self, runs, style, sid):
        "Explicit and implicit prose share first-paragraph styling and nested-quote indentation."
        extra = [e.ind(left=720 * self.bq)] if style == 'blockquote' and self.bq > 1 else None
        use = 'firstpara' if self.first and style == 'body' and not sid else style
        self.first = False
        return self.para(runs, use, extra, sid)

    def bookmark(self, el, runs, part=''):
        "Wrap `runs` in a bookmark when `el` carries an id (target for internal links)"
        if not (i := el['attrs'].get('id')): return runs
        self._bkid += 1
        return [e.bookmarkStart(id=self._bkid, name=self.bkname(i, part)), *runs, e.bookmarkEnd(id=self._bkid)]

    def codeblock(self, node):
        "Source Code paragraphs with optional syntax-colored runs."
        text, lang = node['text'].rstrip('\n'), node['lang']
        segs = (segments(text, lang) if self.hlstyles else None) or [(text, None)]
        runs = []
        for text, scope in segs:
            for i, part in enumerate(text.split('\n')):
                if i: runs.append(e.r(e.br()))
                if part:
                    sid = self.hlsid(scope)
                    runs.append(e.r(e.rPr(e.rStyle(val=sid)) if sid else None, e.t(part, xml__space='preserve')))
        return [self.para(runs, 'codeblock')]


    # ---- lists --------------------------------------------------------------
    def list_el(self, node, ilvl=0):
        "Allocate a live Word list instance; ordered lists retain their start."
        self._numid += 1
        nid = self._numid
        self.nums.append((nid, 1 if node['ordered'] else 0, node['start']))
        return [block for item in node['items'] for block in self.li(item, nid, min(ilvl, 8))]

    def li(self, item, nid, ilvl):
        numpr, cont = [e.numPr(e.ilvl(val=ilvl), e.numId(val=nid))], [e.ind(left=720 * (ilvl + 1))]
        out = []
        for node in item['blocks']:
            if node['type'] == 'list': out += self.list_el(node, ilvl + 1)
            elif node['type'] == 'paragraph':
                runs = self.runs(node['children'], {})
                if not out and item['checked'] is not None: runs = self.text_runs('☒ ' if item['checked'] else '☐ ', {}) + runs
                out.append(self.para(runs, 'list', numpr if not out else cont))
            else: out += self.block(node, 'list')
        return out or [self.para([], 'list', numpr)]


    # ---- tables -------------------------------------------------------------
    def table_grid(self, rows):
        "Resolve row/colspans into per-row cell placements: ('cell', ci, el, cs, rs) / ('cont', ci, width)"
        spans, placed = {}, []
        for ri, tr in enumerate(rows):
            ci, rowcells = 0, []
            def _skip(ci):
                while (ri, ci) in spans:
                    wd = spans.pop((ri, ci))
                    rowcells.append(('cont', ci, wd))
                    ci += wd
                return ci
            ci = _skip(ci)
            for cell in tr['cells']:
                cs, rs = int(cell['attrs'].get('colspan', 1)), int(cell['attrs'].get('rowspan', 1))
                rowcells.append(('cell', ci, cell, cs, rs))
                for k in range(1, rs): spans[(ri + k, ci)] = cs
                ci = _skip(ci + cs)
            placed.append(rowcells)
        return placed

    def col_widths(self, el, ncols):
        "colwidths tracks -> (dxa list, all_fr?) or (None, False) when absent"
        s = el['attrs'].get('colwidths') or el['attrs'].get('data-colwidths')
        if not s: return None, False
        tracks = parse_tracks(s)
        if len(tracks) != ncols:
            self.warn(f'colwidths has {len(tracks)} tracks for {ncols} columns; padding with 1fr')
            tracks = tracks[:ncols] + [('fr', 1.0)] * (ncols - len(tracks))
        fixed = sum(v for k, v in tracks if k == 'dxa')
        frs = sum(v for k, v in tracks if k == 'fr')
        rem = max(self.content_w - fixed, 0)
        dxa = [round(v) if k == 'dxa' else round(v * rem / frs) for k, v in tracks]
        return dxa, fixed == 0

    def cell_blocks(self, cell, header):
        "Render cell blocks; Word requires a final paragraph."
        content = cell['content']
        if len(content) == 1 and content[0]['type'] == 'paragraph' and not content[0]['attrs']:
            jc = [e.jc(val=cell['align'])] if cell['align'] in ('center', 'right') else None
            return [self.para(self.runs(content[0]['children'], {'b': True} if header else {}), 'compact', jc)]
        blocks = self.block_nodes(content)
        if not blocks or blocks[-1].qname != (W, 'p'): blocks.append(Tree(e.p().bytes()).root)
        return blocks

    def table(self, node):
        entries = [(section == 'head', entry) for section in ('head', 'rows', 'foot') for entry in node[section]]
        head = [entry for header, entry in entries if header and entry['type'] == 'row']
        if head and all(block['type'] == 'paragraph' and not block['children'] for row in head for cell in row['cells'] for block in cell['content']):
            entries = [(header, entry) for header, entry in entries if not header or entry['type'] == 'operation']
        rows = [entry for _, entry in entries if entry['type'] == 'row']
        placed, ncols = self.table_grid(rows), node['columns']
        dxa, all_fr = self.col_widths(node, ncols)
        def _tcw(ci, cs):
            if not dxa: return None
            width = sum(dxa[ci:ci + cs])
            return e.tcW(type='pct', w=round(width / self.content_w * 5000)) if all_fr else e.tcW(type='dxa', w=width)
        tblw = e.tblW(type='auto', w=0) if not dxa else e.tblW(type='pct' if all_fr else 'dxa', w=5000 if all_fr else sum(dxa))
        tblpr = e.tblPr(e.tblStyle(val=self.custom_style(node, 'table') or _sid('table')), tblw,
            e.tblLayout(type='fixed') if dxa and not all_fr else None,
            e.tblLook(val='04A0', firstRow=1, lastRow=0, firstColumn=0, lastColumn=0, noHBand=0, noVBand=1))
        widths = dxa or [self.content_w // ncols] * ncols
        grid = e.tblGrid(e.gridCol(w=width) for width in widths)
        def _trpr(header=False):
            return e.trPr(e.cantSplit(val='off' if node['attrs'].get('keep-rows') == 'false' else 'on'), e.tblHeader() if header else None)
        trs, ri = [], 0
        for header, entry in entries:
            if entry['type'] == 'operation':
                tcpr = e.tcPr(_tcw(0, ncols), e.gridSpan(val=ncols) if ncols > 1 else None)
                trs.append(e.tr(_trpr(), e.tc(tcpr, self.para(self.tmpl_runs(entry, {})))))
                continue
            cells = []
            for item in placed[ri]:
                if item[0] == 'cont':
                    _, ci, width = item
                    cells.append(e.tc(e.tcPr(_tcw(ci, width), e.gridSpan(val=width) if width > 1 else None, e.vMerge()), e.p()))
                else:
                    _, ci, cell, cs, rs = item
                    tcpr = e.tcPr(_tcw(ci, cs), e.gridSpan(val=cs) if cs > 1 else None, e.vMerge(val='restart') if rs > 1 else None)
                    cells.append(e.tc(tcpr, self.cell_blocks(cell, header)))
            trs.append(e.tr(_trpr(header), cells))
            ri += 1
        return self.caption_para(node, 'tbl', node['caption']) + [Tree(e.tbl(tblpr, grid, trs).bytes()).root, Tree(e.p().bytes()).root]


    def caption_para(self, node, typ, caption, fmt={}):
        "A live SEQ caption with label and number-only bookmarks."
        if not caption and not node['attrs'].get('id'): return []
        label = self.reftypes[typ][0]
        sequence = dict(fig='Figure', tbl='Table', eq='Equation')[typ]
        seq = self.bookmark(node, [self.field(rf'SEQ {sequence} \* ARABIC', '#', fmt)], 'number')
        runs = self.bookmark(node, self.text_runs(label + ' ', fmt) + seq)
        if caption: runs += self.text_runs(': ', fmt)
        if typ != 'eq': runs += self.bookmark(node, self.runs(caption, fmt), 'text')
        return [self.para(runs, 'caption')]


    def figure(self, node):
        saved, self.figure_alt = self.figure_alt, node['caption_text']
        try: body = self.block_nodes(node['body'])
        finally: self.figure_alt = saved
        return body + self.caption_para(node, 'fig', node['caption'])


    # Paragraphs directly after these blocks (or at document start) take First Paragraph rather
    # than Body Text, matching pandoc's docx writer exactly, so the two agree on which is "first".
    FIRST_AFTER = {'heading', 'block_quote', 'code_block', 'list', 'table', 'definition_list', 'thematic_break', 'math_block'}

    def block(self, node, style='body', sid=None):
        out = self._block(node, style, sid)
        if node['type'] in self.FIRST_AFTER: self.first = True
        return out

    def _block(self, node, style, sid):
        kind = node['type']
        if kind == 'paragraph':
            return [self.prose(self.bookmark(node, self.runs(node['children'], {})), style, self.custom_style(node, 'paragraph') or sid)]
        if kind == 'heading': return [self.para(self.bookmark(node, self.runs(node['children'], {})), f'h{node["level"]}')]
        if kind == 'block_quote':
            self.bq += 1
            try: return self.block_nodes(node['children'], 'blockquote')
            finally: self.bq -= 1
        if kind == 'code_block': return self.codeblock(node)
        if kind == 'list': return self.list_el(node)
        if kind == 'table': return self.table(node)
        if kind == 'thematic_break': return [Tree(e.p(e.pPr(e.pBdr(e.bottom(val='single', sz=6, space=1, color='auto')))).bytes()).root]
        if kind == 'definition_list': return self.dl(node)
        if kind == 'raw_block': return self.rawxml(node)
        if kind == 'figure': return self.figure(node)
        if kind == 'operation': return [self.para(runs)] if (runs := self.tmpl_runs(node, {})) else []
        if kind == 'panel':
            title = node['title']
            runs = self.bookmark(title, self.runs(title['children'], {'b': True})) if title else self.text_runs(node['label'], {'b': True})
            return ([self.para(runs, style)] if runs else []) + self.block_nodes(node['body'], style, sid)
        if kind == 'math_block':
            zone = Tree(e.p(m.oMathPara(self.omath(node['tex']))).bytes()).root
            return [zone] + (self.caption_para(node, 'eq', []) if node['attrs'].get('id') else [])
        if kind in ('div', 'element'): return self.block_nodes(node['children'], style, self.custom_style(node, 'paragraph') or sid)
        return []


    RAWNS = ' '.join(f'xmlns:{k}="{v}"' for k, v in NS.items() if k != 'xml')

    BIND_NS = 'urn:mdhtml:fields'
    BIND_ID = '{8E2C9A44-7D31-4E5B-9C0D-1A6F2B3C4D5E}'   # fixed datastore id, so builds are reproducible

    def tmpl_runs(self, node, fmt):
        "Render a structured operation as text, a field, or a content control."
        if node['syntax'] == 'mustache' and node['action'] in ('section', 'inverted', 'end'):
            marker = dict(section='#', inverted='^', end='/')[node['action']]
            return self.text_runs(f'«{marker}{node["name"]}»', fmt)
        if self.tmpl is None or (res := self.tmpl(node)) is None: return []
        if isinstance(res, str): return self.text_runs(res, fmt)
        kind, val = res
        if kind == 'field': return [self.field(val, f'«{node["name"]}»', fmt)]
        if kind in ('control', 'bound'):
            self.has_controls = True
            binding = None
            if kind == 'bound':
                if val not in self.bound: self.bound.append(val)
                binding = e.dataBinding(prefixMappings=f"xmlns:ns0='{self.BIND_NS}'",
                    xpath=f'/ns0:fields[1]/ns0:{val}[1]', storeItemID=self.BIND_ID)
            sdtpr = e.sdtPr(e.alias(val=val), e.tag(val=val), e.showingPlcHdr(), binding, e.text())
            return [e.sdt(sdtpr, e.sdtContent(e.r(self.rpr(fmt | {'rstyle': 'PlaceholderText'}), e.t(val))))]
        raise ValueError(f'unknown template rendering {res!r}')



    def rawxml(self, node):
        "Parse a decoded raw docx payload as WordprocessingML."
        if node['format'] != 'docx': return []
        try: return Tree(f'<m2d {self.RAWNS}>{node["text"]}</m2d>'.encode()).root.children
        except ValueError as ex:
            self.warn(f'malformed docx raw payload: {ex}')
            return []

    def dl(self, node):
        out = []
        for item in node['items']:
            for term in item['terms']: out.append(self.para(self.bookmark(term, self.runs(term['children'], {})), 'dt'))
            for definition in item['definitions']: out += self.block_nodes(definition, 'dd')
        return out

    def block_nodes(self, nodes, style='body', sid=None): return [block for node in nodes for block in self.block(node, style, sid)]


    # ---- assembly -----------------------------------------------------------
    BULLETS = ['•', '◦', '▪']
    NUMFMTS = ['decimal', 'lowerLetter', 'lowerRoman']

    def _update_numbering(self):
        "Add list definitions and the heading scheme to the live numbering part, in schema order."
        root, abstracts, nums = self._part('numbering').xml.root, [], []
        if self.nums:
            for aid in (0, 1):
                levels = []
                for i in range(9):
                    fmt, txt = ('bullet', self.BULLETS[i % 3]) if aid == 0 else (self.NUMFMTS[i % 3], f'%{i+1}.')
                    levels.append(e.lvl(  # chkstyle: ignore-node
                        e.start(val=1), e.numFmt(val=fmt), e.lvlText(val=txt), e.lvlJc(val='left'),
                        e.pPr(e.ind(left=720 * (i + 1), hanging=360)), ilvl=i))
                abstracts.append(e.abstractNum(e.multiLevelType(val='hybridMultilevel'), levels, abstractNumId=self._absbase + aid))
        if self.headnum:
            levels = []
            for i in range(9):
                txt, fmt = self.scheme[i] if i < len(self.scheme) else (f'%{i+1}.', 'decimal')
                levels.append(e.lvl(  # chkstyle: ignore-node
                    e.start(val=1), e.numFmt(val=fmt), e.pStyle(val=HEADING_STYLE_IDS[i]) if i < 6 else None,
                    e.suff(val='nothing') if not txt else None,   # an empty number (the title's) takes no tab either
                    e.lvlText(val=txt), e.lvlJc(val='left'), ilvl=i))
            abstracts.append(e.abstractNum(e.multiLevelType(val='multilevel'), levels, abstractNumId=self._absbase + 2))
            nums.append(e.num(e.abstractNumId(val=self._absbase + 2), numId=self.headnum))
        for nid, aid, start in self.nums:
            nums.append(e.num(e.abstractNumId(val=self._absbase + aid),  # chkstyle: ignore-node
                (e.lvlOverride(e.startOverride(val=start if i == 0 else 1), ilvl=i) for i in range(9)), numId=nid))
        for node in abstracts: root(node)
        for node in nums: root(node)

    def bind_item_xml(self):
        "One empty element per bound variable; every same-name control is a live view of it"
        fields = E('ns0', ns={'ns0': self.BIND_NS})
        return fields.fields(fields(name) for name in self.bound).bytes()

    def _number_heading_styles(self):
        "Patch w:numPr into the Title and Heading1-5 styles, binding them to the generated heading numbering; the Title's level 0 is what restarts the count at every h1"
        for i, sid in enumerate(HEADING_STYLE_IDS):
            st = next((s for s in self.styles.elements(w.Style) if s.attribute(W, 'styleId') == sid), None)
            if st is None:
                if i == 0: self.warn('no Title style in the reference: an h1 will not restart the heading numbering')
                continue
            ppr = st.child(w.StyleParagraphProperties)
            if ppr is None: ppr = st(e.pPr())
            ppr(e.numPr(e.ilvl(val=i), e.numId(val=self.headnum)))

    def _update_styles(self):
        "Complete the live styles part with heading bindings, custom-style stubs and control placeholders."
        if self.headnum: self._number_heading_styles()
        root = self.styles.root
        for name, (kind, sid) in self.stubs.items():
            root(e.style(  # chkstyle: ignore-node
                e.name(val=name), e.basedOn(val='BodyText' if kind == 'paragraph' else 'DefaultParagraphFont'),
                e.qFormat(), type=kind, customStyle=1, styleId=sid))
        if self.has_controls and 'placeholder text' not in self.refstyles:
            root(e.style(  # chkstyle: ignore-node
                e.name(val='Placeholder Text'), e.basedOn(val='DefaultParagraphFont'), e.semiHidden(),
                e.rPr(e.color(val='808080')), type='character', styleId='PlaceholderText'))

    def to_docx(self, source, dest):
        document = import_mdhtml(source)
        self.index = document.index(reftypes=self.reftypes, number_headings=self.number_headings)
        self.reftypes = self.index.reftypes
        option = self.index.number_headings
        self.scheme = HeadingNums(option).scheme if option else None
        if self.scheme:
            heading = next(s for s in self.styles.elements(w.Style) if s.attribute(W, 'styleId') == 'Heading1')
            if next(heading.elements(w.NumberingProperties), None) is None:
                self._numid += 1
                self.headnum = self._numid
        self.warnings.extend(document.warnings)
        self.fndefs = document.footnotes
        for note in list(self.doc.footnotes): note.delete()
        blocks = self.block_nodes(document.blocks)
        self.package.replace_part(self.package.main_part, _partxml.document(e.body(blocks, self.sectpr)).bytes())
        if self.nums or self.headnum: self._update_numbering()
        if self.bound: self.doc.set_custom_xml(self.BIND_ID, self.bind_item_xml(), schema_uri=self.BIND_NS)
        self._update_styles()
        if self.has_fields: self.doc.settings['updateFields'] = True
        self.doc.save(dest)
        return self.warnings


def mustache_fields(node):
    "Render a variable operation as a live Word MERGEFIELD."
    return 'field', f'MERGEFIELD {node["name"]}'




def mdhtml2docx(
    mdhtml, # Semantic Document or symbolic MDHTML
    dest, # Output docx path
    reference=None, # Reference docx, contributor list, or default template and theme
    base=None, # Base directory for local images
    reftypes=None, # Additional singular/plural reference labels
    number_headings=None, # Heading scheme; None inherits document policy, False disables
    tmpl=None, # Structured operation renderer: text, field/control/bound tuple, or None
):
    "Write a semantic document as docx with native Word references; return conversion warnings."
    return Converter(reference, base, reftypes, number_headings, tmpl).to_docx(mdhtml, dest)
