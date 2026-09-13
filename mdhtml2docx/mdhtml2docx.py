"""Convert MDHTML fragments to docx.

The reference template supplies styles/theme/fonts; oxml builds the generated XML and
owns its package, parts, and relationships. Block and inline walkers mirror the MDHTML element
inventory; STYLE_MAP names every style we emit."""
import posixpath, re
from pathlib import Path
from fast5ever import Comment, Element, Node, Text
from mdhtml import mdhtml2dom
from mdhtml.export import REFTYPES, SCHEMES, decode_raw, tmpl_node, group_plan, ref_tokens, ref_variant, target_kind, Resolver, panel_parts
from oxml import Document, Tree, w
from .styles import STYLE_MAP, style_id, theme_styles
from .styles import ref_path as _refpath
from .wml import *
from .hilite import segments, tokenize

__all__ = ['mdhtml2docx', 'mustache_fields']

m = E('m', attr_ns='m')
# Pandoc reads drawing prefixes from part roots, not just their local XML scope.
_partxml = E('w', attr_ns='w', ns=NS)

def _sid(key): return style_id(STYLE_MAP[key])

def _related_part(package, kind, source=None):
    source = source or package.main_part
    rel = next((r for r in package.relationships(source) if r['type'] == f'{R}/{kind}'), None)
    return package.relationship_part(source, rel['id']) if rel else None

WPML = 'application/vnd.openxmlformats-officedocument.wordprocessingml'
IMAGE_TYPES = dict(png='image/png', jpeg='image/jpeg', jpg='image/jpeg', gif='image/gif', tiff='image/tiff')

BLOCK_TAGS = set(('address article aside blockquote details dialog div dl fieldset figure footer form h1 h2 h3 h4 h5 h6 '
    'header hgroup hr main menu nav ol p pre search section table ul').split())
INERT_TAGS = {'base', 'link', 'meta', 'style', 'template', 'title'}
HEADING_STYLE_IDS = ['Title'] + [f'Heading{i}' for i in range(1, 6)]   # h1-h6, one per numbering level: the Title carries level 0

def _walk(el):
    yield el
    for child in el.element_children: yield from _walk(child)
def _classes(el): return (el.attrs.get('class') or '').split()
def _is_raw(el): return el.name == 'script' and el.attrs.get('type') == 'application/vnd.mdhtml.raw'

def parse_frag(src):
    "Parse an MDHTML body fragment, or return an existing mutable fragment"
    if isinstance(src, Node): return src
    if not isinstance(src, str): raise TypeError('input must be an MDHTML string or fast5ever node')
    return mdhtml2dom(src)

class Converter:
    def __init__(self, reference=None, base=None, reftypes=None, number_headings=None, tmpl=None):
        self.tmpl = tmpl
        if reference is None: reference = [_refpath()] + (['github_light'] if tokenize else [])
        elif not isinstance(reference, (list, tuple)): reference = [reference]
        self.doc = Document.open(reference[0] or _refpath())
        self.package = self.doc.package
        self.part_uris = {kind: _related_part(self.package, kind) for kind in ('styles', 'numbering', 'settings', 'footnotes')}
        self._source_uri = self.package.main_part
        self._part_names = {p.lower() for p in self.package.part_names()}
        self.base = Path(base or '.')
        self.warnings = []
        self.bq = 0          # blockquote nesting depth
        self._bkid = 0       # bookmark id counter
        self._imgn = 0       # image part counter (doubles as docPr id)
        self._urlrids = {}   # (source part, hyperlink URL) -> rId
        self.nums = []       # (numId, abstractNumId, start) per list instance
        self.fndefs = {}     # endnote li elements by id, harvested before the walk
        self.fnids = {}      # endnote id -> footnote w:id
        self.fnotes = []     # (w:id, [footnote blocks])
        self.stubs = {}      # undefined custom-style name -> (kind, styleId)
        self.first = True    # next body paragraph is a 'First Paragraph' (doc start; reset after FIRST_AFTER blocks)
        self._bknames = {}   # element id -> Word-legal bookmark name
        self.has_fields = False
        self.has_controls = False
        self.bound = []      # distinct bound-control names, in first-appearance order
        self.reftypes = REFTYPES | (reftypes or {})
        self.sectpr = wchild(wchild(self.doc.main.xml.root, 'body'), 'sectPr')
        pg, mar = wchild(self.sectpr, 'pgSz'), wchild(self.sectpr, 'pgMar')
        self.content_w = int(pg.attribute(W, 'w')) - int(mar.attribute(W, 'left')) - int(mar.attribute(W, 'right'))
        if not self.part_uris['styles']: raise ValueError('reference doc lacks a styles relationship')
        self.styles = self.package.part(self.part_uris['styles']).xml
        numbering = self.package.part(uri).xml if (uri := self.part_uris['numbering']) else None
        used = [int(v) for e in ([] if numbering is None else numbering.root.children)
            for a in ('numId', 'abstractNumId') if (v := e.attribute(W, a)) is not None]
        self._numid = max(used, default=0)
        for ref in reference[1:]: self._merge_styles(ref)
        self._absbase = self._numid + 1
        self._numid += 3     # reserve bullet, ordered-list, and heading abstract IDs after all contributors
        self.refstyles = {wchild(s, 'name').attribute(W, 'val').lower(): s.attribute(W, 'styleId') for s in self.styles.elements(w.Style)}
        if missing := [n for n in STYLE_MAP.values() if n.lower() not in self.refstyles]:
            raise ValueError(f'reference doc lacks dialect styles (map/template drift?): {missing}')
        self.hlstyles = {n.removeprefix('hl ').replace(' ', '.'): sid for n, sid in self.refstyles.items() if n.startswith('hl ')}
        if isinstance(number_headings, str):
            if number_headings not in SCHEMES: raise ValueError(f'unknown numbering scheme {number_headings!r}')
            number_headings = SCHEMES[number_headings]
        self.scheme = list(number_headings.items()) if number_headings else None
        self.headnum = None
        heading = next((s for s in self.styles.elements(w.Style) if s.attribute(W, 'styleId') == 'Heading1'), None)
        if number_headings and wchild(wchild(heading, 'pPr'), 'numPr') is None:
            self._numid += 1
            self.headnum = self._numid

    def _merge_styles(self, ref):
        "Merge a contributor's w:style elements, later-wins on style id or name"
        def keys(e):
            nm = wchild(e, 'name')
            return {(e.attribute(W, 'styleId') or '').lower(), '' if nm is None else nm.attribute(W, 'val').lower()} - {''}
        if str(ref).endswith('.docx'):
            package = Document.open(ref).package
            uri = _related_part(package, 'styles')
            if uri is None: raise ValueError(f'style contributor lacks a styles relationship: {ref}')
            new = list(package.part(uri).xml.elements(w.Style))
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
                if e.raw['qname'] != (W, tag): continue
                self._numid += 1
                mapping[e.attribute(W, attr)] = str(self._numid)
                e.set_attribute(W, attr, str(self._numid))
            for e in tree.elements():
                if e.raw['qname'] == (W, attr) and (val := e.attribute(W, 'val')) in mapping: e.set_attribute(W, 'val', mapping[val])
        for e in tree.root.children:
            if e.raw['qname'] not in ((W, 'abstractNum'), (W, 'num')): continue
            root = self._part('numbering').xml.root
            before = {'num', 'numIdMacAtCleanup'} if e.raw['qname'] == (W, 'abstractNum') else {'numIdMacAtCleanup'}
            e.copy_to(root, wpos(root, before))
        return list(tree.elements(w.Style))

    def hlsid(self, scope):
        "Hl* style id for a dotted scope: exact, else progressively shorter prefixes (tree-sitter resolution)"
        parts = (scope or '').split('.')
        while parts:
            if sid := self.hlstyles.get('.'.join(parts)): return sid
            parts.pop()

    def _add_part(self, uri, content_type, data):
        "Add a new part without overwriting reference parts, including names differing only in case"
        stem, ext = posixpath.splitext(uri)
        n, candidate = 2, uri
        while candidate.lower() in self._part_names: candidate, n = f'{stem}_{n}{ext}', n + 1
        self.package.add_part(candidate, content_type, data)
        self._part_names.add(candidate.lower())
        return candidate

    def _part(self, kind):
        "A related Word part, creating an empty root and its relationship when absent."
        if not (uri := self.part_uris[kind]):
            uri = self._add_part(posixpath.join(posixpath.dirname(self.package.main_part), f'{kind}.xml'),
                f'{WPML}.{kind}+xml', e(kind).bytes())
            target = posixpath.relpath(uri, posixpath.dirname(self.package.main_part))
            self.package.add_relationship(self.package.main_part, f'{R}/{kind}', target)
            self.part_uris[kind] = uri
        return self.package.part(uri)

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
        return e.fldSimple(e.r(self.rpr(fmt), e.t(text)), instr=f' {instr.strip()} ')

    def link(self, el, fmt):
        "w:hyperlink for `a`: internal '#x' -> anchor, external -> relationship (deduped per URL); data-ref -> field"
        if el.attrs.get('data-ref') is not None:
            fld = self.ref_fld(el, fmt)
            return self.ref_prefix(el, fmt) + fld
        href = el.attrs.get('href')
        if not href: return self.runs(el, fmt)
        runs = self.runs(el, fmt | {'rstyle': _sid('hyperlink')})
        if href.startswith('#'): return [e.hyperlink(runs, anchor=self.bkname(href[1:]))]
        return self.external_link(href, runs)

    def external_link(self, href, runs):
        key = self._source_uri, href
        if key not in self._urlrids:
            self._urlrids[key] = self.package.add_relationship(self._source_uri, f'{R}/hyperlink', href, 'External')
        return [e.hyperlink(runs, r__id=self._urlrids[key])]

    REFSWITCH = dict(full=r'\w', rel=r'\r', leaf=r'\n', text='', page=None)

    def bkname(self, id):
        "Word-legal bookmark name for `id` (letter first, word chars only), stable within the document"
        if id not in self._bknames:
            nm = re.sub(r'\W', '_', id)
            if not nm[:1].isalpha(): nm = 'B' + nm
            while nm in self._bknames.values(): nm += '_'
            self._bknames[id] = nm
        return self._bknames[id]

    def ref_prefix(self, el, fmt, plural=False):
        "Literal runs before a reference field: the shared `Resolver.prefix` (override text, the type word, or nothing for bare, caption, and text refs)"
        tgt = (el.attrs.get('href') or '#')[1:]
        pre = self.res.prefix(el.to_text().strip(), tgt, ref_tokens(el.attrs.get('data-ref')), plural)
        return self.text_runs(pre, fmt) if pre else []

    def ref_fld(self, el, fmt):
        """REF/PAGEREF field for a cross-reference `a`, with a cached placeholder Word replaces on update.
        Heading/paragraph targets number via `\\w`; caption targets return their bookmarked 'Label N' text
        (or the number-only `_n` bookmark for bare/leaf/rel refs), and text targets (spans, definition
        terms) their bookmark text, so `\\w` never applies to either."""
        tgt = (el.attrs.get('href') or '#')[1:]
        tokens = ref_tokens(el.attrs.get('data-ref'))
        self.res.check(tgt)
        variant = ref_variant(tokens)
        nm = self.bkname(tgt)
        if variant == 'page': instr, cached = rf'PAGEREF {nm} \h', '#'
        elif self.reftarget[tgt] == 'caption':
            bare = 'bare' in tokens or variant in ('leaf', 'rel')
            instr, cached = (rf'REF {nm}_n \h' if bare else rf'REF {nm} \h'), '#'
        elif self.reftarget[tgt] == 'text': instr, cached = rf'REF {nm} \h', self.res.core(tgt, tokens)
        else:
            sw = self.REFSWITCH[variant]
            instr = rf'REF {nm} {sw} \h' if sw else rf'REF {nm} \h'
            cached = self.idtext.get(tgt, '#') if variant == 'text' else '#'
        return [self.field(instr, cached, fmt)]

    def ref_group(self, el, fmt):
        "data-refs span: one pluralized prefix for a same-type group, per-item singular prefixes for mixed types; never range-collapsed"
        refs = [c for c in el.element_children if c.name == 'a']
        types = [(a.attrs.get('href') or '#')[1:].split('-')[0] for a in refs]
        out = []
        for (sep, pre, plural), a in zip(group_plan(types), refs):
            if sep: out += self.text_runs(sep, fmt)
            if pre: out += self.ref_prefix(a, fmt, plural=plural)
            out += self.ref_fld(a, fmt)
        return out

    def custom_style(self, el, kind):
        "Style id for an explicit custom-style attr (stubbed + warned if undefined), else a class matching a reference style name"
        if cs := el.attrs.get('custom-style'):
            if cs.lower() in self.refstyles: return self.refstyles[cs.lower()]
            if cs not in self.stubs:
                self.stubs[cs] = (kind, re.sub(r'\W', '', cs) or f'Custom{len(self.stubs)}')
                self.warn(f'custom style {cs!r} not in reference doc; stub injected')
            return self.stubs[cs][1]
        return next((self.refstyles[c.lower()] for c in _classes(el) if c.lower() in self.refstyles), None)

    def span(self, el, fmt):
        "Inline span: math -> inline m:oMath zone (linear source, dialect-agnostic), custom style -> rStyle, else transparent; an id bookmarks the runs"
        if el.attrs.get('data-refs') is not None: return self.ref_group(el, fmt)
        if el.has_class('math'): return self.bookmark(el, [self.omath(el)])
        if sid := self.custom_style(el, 'character'): return self.bookmark(el, self.runs(el, fmt | {'rstyle': sid}))
        return self.bookmark(el, self.runs(el, fmt))

    def omath(self, el):
        "An m:oMath zone holding `el`'s text as linear-format math runs"
        return m.oMath(m.r(m.t(el.to_text(), xml__space='preserve')))

    def fnref(self, el, fmt):
        "Footnote-reference run for a sup>a.footnote-ref, or None when `el` is an ordinary sup"
        children = el.element_children
        a = children[0] if len(children) == 1 and children[0].name == 'a' else None
        if a is None or not a.has_class('footnote-ref'): return None
        key = (a.attrs.get('href') or '#')[1:]
        if key not in self.fndefs:
            self.warn(f'footnote reference #{key} has no definition; dropped')
            return []
        if key not in self.fnids:
            self.fnids[key] = len(self.fnids) + 1
            self.fnotes.append((self.fnids[key], self.fn_blocks(self.fndefs[key])))
        return [e.r(e.rPr(e.rStyle(val=_sid('footnoteref'))), e.footnoteReference(id=self.fnids[key]))]

    def fn_blocks(self, li):
        "Footnote body with its reference mark, in footnote-text style, backref stripped."
        save = self._source_uri, self.first
        self._source_uri = self._part('footnotes').uri
        try: blks = self.blocks(li, 'footnotetext')
        finally: self._source_uri, self.first = save
        if not blks or blks[0].raw['qname'] != (W, 'p'): blks.insert(0, self.para([], 'footnotetext'))
        p = blks[0]
        ppr = wchild(p, 'pPr')
        pos = p.raw['children'].index(ppr.node_id) + 1 if ppr is not None else 0
        p(e.r(e.rPr(e.rStyle(val=_sid('footnoteref'))), e.footnoteRef()), index=pos)
        p(e.r(e.t(' ', xml__space='preserve')), index=pos + 1)
        return blks

    def image(self, el, fmt, alt=None):
        "Embed a local image (dimensions sniffed, width/height px attrs override); remote srcs degrade to a link"
        src = el.attrs.get('src') or ''
        if alt is None: alt = el.attrs.get('alt') or src
        if re.match(r'[a-z][a-z0-9+.-]*://', src):
            self.warn(f'remote image not embedded: {src}')
            return self.external_link(src, self.text_runs(alt, fmt | {'rstyle': _sid('hyperlink')}))
        try: data = (self.base/src).read_bytes()
        except OSError:
            self.warn(f'image not found: {src}; alt text emitted')
            return self.text_runs(alt, fmt)
        pw, ph, dx, dy = imgsize(data) or (300, 200, 96, 96)
        cx, cy = round(pw * 914400 / dx), round(ph * 914400 / dy)
        w_, h_ = el.attrs.get('width'), el.attrs.get('height')
        if w_: cx = round(float(w_) * EMU_PER_PX)
        if h_: cy = round(float(h_) * EMU_PER_PX)
        if w_ and not h_: cy = round(cx * ph / pw)
        if h_ and not w_: cx = round(cy * pw / ph)
        self._imgn += 1
        ext = Path(src).suffix.lower() or '.bin'
        name = posixpath.join(posixpath.dirname(self.package.main_part), f'media/image{self._imgn}{ext}')
        uri = self._add_part(name, IMAGE_TYPES.get(ext[1:], 'application/octet-stream'), data)
        target = posixpath.relpath(uri, posixpath.dirname(self._source_uri))
        rid = self.package.add_relationship(self._source_uri, f'{R}/image', target)
        return [e.r(drawing(rid, self._imgn, cx, cy, alt))]

    INLINE_FMT = {'em': {'i': True}, 'strong': {'b': True}, 'code': {'rstyle': _sid('codeinline')},
        'del': {'strike': True}, 'mark': {'mark': True}, 'u': {'u': True}, 'sub': {'vert': 'subscript'}}

    def inline(self, el, fmt):
        "Run-level elements for inline `el` under formatting context `fmt` (dict; copied on change)"
        tag = el.name
        if tag in self.INLINE_FMT: out = self.runs(el, fmt | self.INLINE_FMT[tag])
        elif tag == 'a': out = self.link(el, fmt)
        elif tag == 'sup':
            fn = self.fnref(el, fmt)
            out = fn if fn is not None else self.runs(el, fmt | {'vert': 'superscript'})
        elif tag == 'span': out = self.span(el, fmt)
        elif tag == 'img': out = self.image(el, fmt)
        elif tag == 'br': out = [e.r(self.rpr(fmt), e.br(type='page' if el.attrs.get('type') == 'page' else None))]
        elif tag == 'input':   # task-list checkbox
            if el.attrs.get('type') != 'checkbox':
                self.warn(f'unhandled inline <input type={el.attrs.get("type")!r}>; dropped')
                return []
            g = '☒' if el.attrs.get('checked') is not None else '☐'
            out = [e.r(self.rpr(fmt), e.t(g + ' ', xml__space='preserve'))]
        elif tag == 'script': out = self.rawxml(el)
        elif tag == 'template': out = self.tmpl_runs(el, fmt, 'inline')
        elif tag in INERT_TAGS: out = []
        else:  # unknown inline (abbr etc): recurse transparently
            out = self.runs(el, fmt)
        return out

    def runs(self, el, fmt):
        "Run-level elements for an element's ordered text and element children"
        return self.group_runs(el.children, fmt)

    def group_runs(self, nodes, fmt):
        "Runs for a mixed list of text and inline nodes"
        return [r for node in nodes for r in self.inline_node(node, fmt)]

    def inline_node(self, node, fmt):
        if isinstance(node, Text): return self.text_runs(node.text, fmt)
        if isinstance(node, Element):
            if node.name == 'a' and node.has_class('footnote-backref'): return []
            return self.inline(node, fmt)
        return []

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

    def bookmark(self, el, runs):
        "Wrap `runs` in a bookmark when `el` carries an id (target for internal links)"
        if not (i := el.attrs.get('id')): return runs
        self._bkid += 1
        return [e.bookmarkStart(id=self._bkid, name=self.bkname(i)), *runs, e.bookmarkEnd(id=self._bkid)]

    def codeblock(self, el):
        "Source Code paragraph, lines joined with w:br; Hl* character styles when a language class names one"
        children = el.element_children
        code = children[0] if children and children[0].name == 'code' else el
        lang = next((c.removeprefix('language-') for c in _classes(code) if c.startswith('language-')), None)
        text = code.to_text().rstrip('\n')
        segs = (segments(text, lang) if self.hlstyles else None) or [(text, None)]
        runs = []
        for txt, scope in segs:
            for j, part in enumerate(txt.split('\n')):
                if j: runs.append(e.r(e.br()))
                if not part: continue
                sid = self.hlsid(scope)
                runs.append(e.r(e.rPr(e.rStyle(val=sid)) if sid else None, e.t(part, xml__space='preserve')))
        return [self.para(runs, 'codeblock')]

    # ---- lists --------------------------------------------------------------
    def list_el(self, el, ilvl=0):
        "A ul/ol: fresh num instance (so each ordered list restarts), items at level `ilvl`"
        self._numid += 1
        nid = self._numid
        self.nums.append((nid, 0 if el.name == 'ul' else 1, int(el.attrs.get('start', 1))))
        return [b for li in el.element_children if li.name == 'li' for b in self.li(li, nid, min(ilvl, 8))]

    def li_parts(self, el):
        "Split mixed li content into ('inline', [nodes]) groups and ('block', child) items, in order"
        parts = []
        def add(x):
            if isinstance(x, Text) and not x.text.strip() and '\n' in x.text: return
            if parts and parts[-1][0] == 'inline': parts[-1][1].append(x)
            else: parts.append(('inline', [x]))
        for node in el.children:
            if isinstance(node, Element) and node.name in BLOCK_TAGS: parts.append(('block', node))
            elif isinstance(node, (Text, Element)): add(node)
        return parts

    def li(self, li, nid, ilvl):
        "Blocks for one list item: the first paragraph carries the number, the rest continue indented"
        numpr = [e.numPr(e.ilvl(val=ilvl), e.numId(val=nid))]
        cont = [e.ind(left=720 * (ilvl + 1))]
        out = []
        for kind, val in self.li_parts(li):
            if kind == 'inline': out.append(self.para(self.group_runs(val, {}), 'list', numpr if not out else cont))
            elif val.name in ('ul', 'ol'): out += self.list_el(val, ilvl + 1)
            elif val.name == 'p': out.append(self.para(self.runs(val, {}), 'list', numpr if not out else cont))
            else: out += self.block(val, 'list')
        return out or [self.para([], 'list', numpr)]

    # ---- tables -------------------------------------------------------------
    def table_grid(self, rows):
        "Resolve row/colspans into per-row cell placements: ('cell', ci, el, cs, rs) / ('cont', ci, width)"
        spans, placed, ncols = {}, [], 0
        for ri, tr in enumerate(rows):
            ci, rowcells = 0, []
            def _skip(ci):
                while (ri, ci) in spans:
                    wd = spans.pop((ri, ci))
                    rowcells.append(('cont', ci, wd))
                    ci += wd
                return ci
            ci = _skip(ci)
            for cell in tr.element_children:
                if cell.name not in ('td', 'th'): continue
                cs, rs = int(cell.attrs.get('colspan', 1)), int(cell.attrs.get('rowspan', 1))
                rowcells.append(('cell', ci, cell, cs, rs))
                for k in range(1, rs): spans[(ri + k, ci)] = cs
                ci = _skip(ci + cs)
            placed.append(rowcells)
            ncols = max(ncols, ci)
        return placed, ncols

    def col_widths(self, el, ncols):
        "colwidths tracks -> (dxa list, all_fr?) or (None, False) when absent"
        s = el.attrs.get('colwidths') or el.attrs.get('data-colwidths')
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
        "Cell content ending in a paragraph; header cells bold, align honored for inline cells."
        if any(c.name in BLOCK_TAGS for c in cell.element_children):
            blocks = self.blocks(cell)
            if not blocks or blocks[-1].raw['qname'] != (W, 'p'): blocks.append(Tree(e.p().bytes()).root)
            return blocks
        jc = [e.jc(val=cell.attrs.get('align'))] if cell.attrs.get('align') in ('center', 'right') else None
        return [self.para(self.runs(cell, {'b': True} if header else {}), 'compact', jc)]

    def table(self, el):
        "w:tbl (+ caption paragraph before, spacer paragraph after)"
        cap = None
        rows, nhead, markers = [], 0, {}
        def _add(c):
            "Rows in order; template markers recorded at the row index they precede"
            if c.name == 'template': markers.setdefault(len(rows), []).append(c)
            else: rows.append(c)
        for sec in el.element_children:
            t = sec.name
            if t == 'caption': cap = sec
            elif t == 'thead':
                for c in sec.element_children: _add(c)
                nhead = len(rows)
            elif t in ('tbody', 'tfoot'):
                for c in sec.element_children: _add(c)
            elif t in ('tr', 'template'): _add(sec)
        if nhead and not any(tr.to_text().strip() for tr in rows[:nhead]):
            # a markdown pipe table cannot omit its header row, so an all-empty thead means "headerless"
            old, markers = markers, {}
            for ri, ms in old.items(): markers.setdefault(max(0, ri - nhead), []).extend(ms)
            rows, nhead = rows[nhead:], 0
        placed, ncols = self.table_grid(rows)
        dxa, all_fr = self.col_widths(el, ncols)
        def _tcw(ci, cs):
            if not dxa: return None
            wd = sum(dxa[ci:ci + cs])
            if all_fr: return e.tcW(type='pct', w=round(wd / self.content_w * 5000))
            return e.tcW(type='dxa', w=wd)
        tblw = (e.tblW(type='auto', w=0) if not dxa  # chkstyle: ignore-node
            else e.tblW(type='pct', w=5000) if all_fr
            else e.tblW(type='dxa', w=sum(dxa)))
        tblpr = e.tblPr(e.tblStyle(val=self.custom_style(el, 'table') or _sid('table')), tblw,  # chkstyle: ignore-node
            e.tblLayout(type='fixed') if dxa and not all_fr else None,
            e.tblLook(val='04A0', firstRow=1, lastRow=0, firstColumn=0, lastColumn=0, noHBand=0, noVBand=1))
        gw = dxa or [self.content_w // ncols] * ncols   # pandoc's docx reader drops tables whose gridCols lack w:w
        grid = e.tblGrid(e.gridCol(w=gw[i]) for i in range(ncols))
        def _trpr(header=False):
            return e.trPr(e.cantSplit(val=int(el.attrs.get('keep-rows') != 'false')), e.tblHeader() if header else None)
        def _marker_tr(mel):
            "A range marker between rows: one full-width literal cell, so forms keep their markers visible"
            tcpr = e.tcPr(_tcw(0, ncols), e.gridSpan(val=ncols) if ncols > 1 else None)
            return e.tr(_trpr(), e.tc(tcpr, self.para(self.tmpl_runs(mel, {}, 'row'))))
        trs = []
        for ri, rowcells in enumerate(placed):
            for mel in markers.get(ri, []): trs.append(_marker_tr(mel))
            tcs = []
            for item in rowcells:
                if item[0] == 'cont':
                    _, ci, wd = item
                    tcs.append(e.tc(e.tcPr(_tcw(ci, wd), e.gridSpan(val=wd) if wd > 1 else None, e.vMerge()), e.p()))
                else:
                    _, ci, cell, cs, rs = item
                    tcpr = e.tcPr(_tcw(ci, cs), e.gridSpan(val=cs) if cs > 1 else None, e.vMerge(val='restart') if rs > 1 else None)
                    body = self.cell_blocks(cell, ri < nhead)
                    tcs.append(e.tc(tcpr, body))
            trs.append(e.tr(_trpr(ri < nhead), tcs))
        for mel in markers.get(len(rows), []): trs.append(_marker_tr(mel))
        out = self.caption_para(el, 'tbl', cap)
        return out + [Tree(e.tbl(tblpr, grid, trs).bytes()).root, Tree(e.p().bytes()).root]

    def caption_para(self, el, typ, capel, fmt={}):
        """Numbered caption paragraph: 'Label N: text' with a SEQ field as N. When `el` has an id, the
        label+number span is bookmarked under it (REF target) and the number alone under `<name>_n`.
        Emitted whenever there is a caption or an id; the label word comes from reftypes[typ]."""
        if capel is None and not el.attrs.get('id'): return []
        label = self.reftypes[typ][0]
        seq = [self.field(rf'SEQ {label} \* ARABIC', '#', fmt)]
        if i := el.attrs.get('id'):
            nm = self.bkname(i)
            self._bknames[i + '\0n'] = nm + '_n'   # reserve the number-only name too
            self._bkid += 2
            seq = [e.bookmarkStart(id=self._bkid, name=nm + '_n'), *seq, e.bookmarkEnd(id=self._bkid)]
            runs = [e.bookmarkStart(id=self._bkid - 1, name=nm),  # chkstyle: ignore-node
                *self.text_runs(label + ' ', fmt), *seq, e.bookmarkEnd(id=self._bkid - 1)]
        else: runs = self.text_runs(label + ' ', fmt) + seq
        cap = [] if capel is None else self.runs(capel, fmt)
        if cap: runs += self.text_runs(': ', fmt) + cap
        return [self.para(runs, 'caption')]

    def figure(self, el):
        "Figure: image paragraph, then its numbered caption paragraph below (Word convention)"
        img = next((c for c in _walk(el) if c.name == 'img'), None)
        capel = next((c for c in el.element_children if c.name == 'figcaption'), None)
        alt = capel.to_text().strip() if capel is not None else None
        out = [] if img is None else [self.para(self.image(img, {}, alt), 'body')]
        return out + self.caption_para(el, 'fig', capel)

    # Paragraphs directly after these blocks (or at document start) take First Paragraph rather
    # than Body Text, matching pandoc's docx writer exactly, so the two agree on which is "first".
    FIRST_AFTER = {'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'ul', 'ol', 'table', 'dl', 'hr'}

    def block(self, el, style='body', sid=None):
        "Block elements for `el` (one element may yield several); `sid` is a custom-style id override for paragraphs"
        out = self._block(el, style, sid)
        tag = el.name
        if tag in self.FIRST_AFTER or (tag == 'div' and el.has_class('display')): self.first = True
        return out

    def _block(self, el, style, sid):
        tag = el.name
        if tag == 'p':
            psid = self.custom_style(el, 'paragraph') or sid
            return [self.prose(self.bookmark(el, self.runs(el, {})), style, psid)]
        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'): return [self.para(self.bookmark(el, self.runs(el, {})), tag)]
        if tag == 'blockquote':
            self.bq += 1
            try: return self.blocks(el, 'blockquote')
            finally: self.bq -= 1
        if tag == 'pre': return self.codeblock(el)
        if tag in ('ul', 'ol'): return self.list_el(el)
        if tag == 'table': return self.table(el)
        if tag == 'hr':
            return [Tree(e.p(e.pPr(e.pBdr(e.bottom(val='single', sz=6, space=1, color='auto')))).bytes()).root]
        if tag == 'dl': return self.dl(el)
        if tag == 'script': return self.rawxml(el)
        if tag == 'figure': return self.figure(el)
        if tag == 'template': return [self.para(runs)] if (runs := self.tmpl_runs(el, {}, 'block')) else []
        if tag == 'div':
            if (panel := panel_parts(el)) is not None:
                title, body, label = panel
                runs = self.bookmark(title, self.runs(title, {'b': True})) if title is not None else self.text_runs(label, {'b': True})
                return ([self.para(runs, style)] if runs else []) + self.block_nodes(body, style, sid)
            if el.has_class('math') and el.has_class('display'): return [Tree(e.p(m.oMathPara(self.omath(el))).bytes()).root]
            return self.blocks(el, style, self.custom_style(el, 'paragraph') or sid)
        if tag in BLOCK_TAGS and any(c.name in BLOCK_TAGS for c in el.element_children):
            return self.blocks(el, style, sid)   # unknown container: recurse
        self.warn(f'unhandled block <{tag}>; emitted as plain paragraph')
        return [self.para(self.runs(el, {}), style, None, sid)]

    RAWNS = ' '.join(f'xmlns:{k}="{v}"' for k, v in NS.items() if k != 'xml')

    BIND_NS = 'urn:mdhtml:fields'
    BIND_ID = '{8E2C9A44-7D31-4E5B-9C0D-1A6F2B3C4D5E}'   # fixed datastore id, so builds are reproducible

    def tmpl_runs(self, el, fmt, form):
        """Template-instruction runs. Non-value Mustache operations are converter policy: literal
        marker runs, so an unfilled form shows its structure. Value operations go through the `tmpl`
        callable with the semantic node dict (see `mdhtml.export.tmpl_node`): str is a literal text
        run, ('field', instr) a live field, ('control', name) an interactive plain-text content
        control, ('bound', name) a data-bound one, None dropped"""
        if 'data-op' not in el.attrs: return []
        node = tmpl_node(el, form)
        action = node['op'].rsplit(':', 1)[-1]
        if action != 'value':
            marker = dict(section='#', inverted='^', end='/').get(action, '')
            return self.text_runs(f'«{marker}{node["value"]}»', fmt)
        if self.tmpl is None: return []
        res = self.tmpl(node)
        if res is None: return []
        if isinstance(res, str): return self.text_runs(res, fmt)
        kind, val = res
        if kind == 'field': return [self.field(val, f'«{node["value"]}»', fmt)]
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


    def rawxml(self, el):
        "Elements parsed from a raw docx payload (`{=docx}` in Markdown); other formats skip silently"
        if el.attrs.get('type') != 'application/vnd.mdhtml.raw' or el.attrs.get('data-format') != 'docx': return []
        payload, warn = decode_raw(el)
        if warn:
            self.warn(warn)
            return []
        try: return Tree(f'<m2d {self.RAWNS}>{payload}</m2d>'.encode()).root.children
        except ValueError as e:
            self.warn(f'malformed docx raw payload: {e}')
            return []

    def dl(self, el):
        "Definition list: dt/dd paragraphs in their dialect styles"
        out = []
        for c in el.element_children:
            t = c.name
            if t == 'dt': out.append(self.para(self.bookmark(c, self.runs(c, {})), 'dt'))
            elif t == 'dd': out += self.blocks(c, 'dd')
        return out

    def blocks(self, parent, style='body', sid=None): return self.block_nodes(parent.children, style, sid)

    def block_nodes(self, nodes, style='body', sid=None):
        out, inline = [], []
        def flush():
            meaningful = [n for n in inline if not isinstance(n, Text) or n.text.strip()]
            if not meaningful:
                inline.clear()
                return
            raw = all(isinstance(n, Element) and _is_raw(n) for n in meaningful)
            if raw:
                for node in meaningful: out.extend(self.block(node, style, sid))
            else: out.append(self.prose(self.group_runs(inline, {}), style, sid))
            inline.clear()
        for node in nodes:
            if isinstance(node, Comment): continue
            if isinstance(node, Element) and node.name == 'template':
                flush()
                if runs := self.tmpl_runs(node, {}, 'block'): out.append(self.para(runs))
                continue
            if isinstance(node, Element) and (node.name in INERT_TAGS or node.name == 'script' and not _is_raw(node)): continue
            if isinstance(node, Element) and node.name in BLOCK_TAGS:
                flush()
                out.extend(self.block(node, style, sid))
            elif isinstance(node, (Text, Element)): inline.append(node)
        flush()
        return out

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

    def footnotes_xml(self):
        "word/footnotes.xml: the two Word-required separator notes plus our harvested ones"
        notes = [e.footnote(e.p(e.pPr(e.spacing(after=0)), e.r(e(typ))), type=typ, id=wid)
            for typ, wid in (('separator', -1), ('continuationSeparator', 0))]
        return _partxml.footnotes(notes, (e.footnote(blks, id=wid) for wid, blks in self.fnotes)).bytes()

    def _number_heading_styles(self):
        "Patch w:numPr into the Title and Heading1-5 styles, binding them to the generated heading numbering; the Title's level 0 is what restarts the count at every h1"
        for i, sid in enumerate(HEADING_STYLE_IDS):
            st = next((s for s in self.styles.elements(w.Style) if s.attribute(W, 'styleId') == sid), None)
            if st is None:
                if i == 0: self.warn('no Title style in the reference: an h1 will not restart the heading numbering')
                continue
            ppr = wchild(st, 'pPr')
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

    def harvest_footnotes(self, els):
        "Split out footnote endnote sections, indexing their li definitions by id; returns body elements"
        body, fn = [], []
        for el in els:
            is_notes = isinstance(el, Element) and el.name == 'section' and el.has_class('footnotes')
            (fn if is_notes else body).append(el)
        self.fndefs.update({li.attrs.get('id'): li for sec in fn for li in _walk(sec) if li.name == 'li' and li.attrs.get('id')})
        return body

    def to_docx(self, mdhtml, dest):
        root = parse_frag(mdhtml)
        nodes = self.harvest_footnotes(root.children)
        self.idtext, self.reftarget, self.res = {}, {}, Resolver(self.reftypes)
        for el in (el for node in nodes if isinstance(node, Element) for el in _walk(node)):
            if not (i := el.attrs.get('id')): continue
            self.idtext[i] = ' '.join(el.to_text().split())
            k = target_kind(el.name)
            if k: self.reftarget[i] = k
            self.res.register(i, k, self.idtext[i])
        blocks = self.block_nodes(nodes)
        self.package.replace_part(self.package.main_part, _partxml.document(e.body(blocks, self.sectpr)).bytes())
        if self.nums or self.headnum: self._update_numbering()
        if self.fnotes: self._part('footnotes').replace(self.footnotes_xml())
        if self.bound: self.doc.set_custom_xml(self.BIND_ID, self.bind_item_xml(), schema_uri=self.BIND_NS)
        self._update_styles()
        if self.has_fields: self._update_settings()
        self.doc.save(dest)
        return self.warnings

    def _update_settings(self):
        "Set updateFields in the live settings part, so Word refreshes fields on open."
        root = self._part('settings').xml.root
        if (update := wchild(root, 'updateFields')) is not None: update.set_attribute(W, 'val', 'true')
        else: root(e.updateFields(val='true'))

def mustache_fields(node):
    "Template variables as live Word `MERGEFIELD`s (markers never reach `tmpl`: the converter shows them literally)"
    return 'field', f'MERGEFIELD {node["value"]}'



def mdhtml2docx(mdhtml, dest, reference=None, base=None, reftypes=None, number_headings=None, tmpl=None):
    """Convert an MDHTML string or mutable fast5ever DOM to a docx file at `dest`; returns warnings.
    `reference` is a reference docx path, or a list of them: the first supplies the whole archive
    (default, or when None: the built-in template), later entries contribute styles only, later-wins -
    each a .docx path, a raw styles/numbering .xml path, or a fastpylight theme name (whose Hl*/Source
    Code styles are generated; see styles.theme_ref). Default adds 'github_light' when fastpylight is
    installed, so code blocks are colored; pass a bare reference for plain code. Relative image srcs
    resolve against `base` ('.'). Cross-references (`data-ref` anchors from Markdown `[@sec-x]`) become
    live REF fields; `reftypes` maps type tokens to (singular, plural) prefix words beyond the built-in
    `sec`, and `number_headings` (a styles.SCHEMES name such as 'legal', or a {lvlText: numFmt} dict, one entry per heading level)
    numbers the headings via a multilevel list so `\\w` fields resolve; scheme level 0 is the h1 document title (Title style), whose empty lvlText shows nothing and whose counter restarts the levels below. Template value instructions are dropped
    unless `tmpl` is given; other operations remain visible markers. `tmpl` is a callable taking the semantic instruction dict
    (`mdhtml.export.tmpl_node`: `op`, `value`, `form`) and returning a str for a literal text run,
    `('field', instr)` for a live field, `('control', name)` for an interactive plain-text content
    control, `('bound', name)` for a content control data-bound to a shared per-variable XML node
    (same-name controls stay in sync as one is filled), or None to drop; range markers never reach `tmpl` and render as literal «body» runs - `mustache_fields` here
    is the ready-made recipe."""
    return Converter(reference, base, reftypes, number_headings, tmpl).to_docx(mdhtml, dest)
