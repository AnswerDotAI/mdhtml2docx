"""Convert MDHTML fragments to docx.

Write-only, reference-archive architecture: the reference template supplies styles/theme/fonts,
we generate word/document.xml (plus footnotes/numbering/media parts as needed) into a copy of its
archive. Block and inline walkers mirror the MDHTML element inventory; STYLE_MAP names every
style we emit."""
import posixpath, re, zipfile
from copy import deepcopy
from pathlib import Path
from fast5ever import Comment, Element, Node, Text
from lxml import etree
from mdhtml import parse_mdhtml
from mdhtml.export import REFTYPES, SCHEMES, decode_raw, tmpl_node, group_plan, ref_tokens, ref_variant, target_kind, Resolver
from .styles import STYLE_MAP, style_id, theme_styles
from .styles import ref_path as _refpath
from .wml import *
from .wml import qn
from .hilite import segments, tokenize

__all__ = ['convert', 'mustache_fields']

def _sid(key): return style_id(STYLE_MAP[key])

BLOCK_TAGS = set(('address article aside blockquote details dialog div dl fieldset figure footer form h1 h2 h3 h4 h5 h6 '
    'header hgroup hr main menu nav ol p pre search section table ul').split())
INERT_TAGS = {'base', 'link', 'meta', 'style', 'template', 'title'}

def _tag(el): return el.name
def _get(el, key, default=None): return el.attrs.get(key, default)
def _els(el): return [n for n in el.children if isinstance(n, Element)]
def _walk(el):
    yield el
    for child in _els(el): yield from _walk(child)
def _classes(el): return (_get(el, 'class') or '').split()
def _is_raw(el): return _tag(el) == 'script' and _get(el, 'type') == 'application/vnd.mdhtml.raw'

def parse_frag(src):
    "Parse an MDHTML body fragment, or return an existing mutable fragment"
    if isinstance(src, Node): return src
    if not isinstance(src, str): raise TypeError('input must be an MDHTML string or fast5ever node')
    return parse_mdhtml(src)

class Converter:
    def __init__(self, reference=None, base=None, reftypes=None, number_headings=None, tmpl=None):
        self.tmpl = tmpl
        if reference is None: reference = [_refpath()] + (['github_light'] if tokenize else [])
        elif not isinstance(reference, (list, tuple)): reference = [reference]
        self.refz = zipfile.ZipFile(reference[0] or _refpath())
        self.contribs = reference[1:] # styles-only contributors: docx paths, styles/numbering .xml paths, theme names
        self.base = Path(base or '.')
        self.rels = []       # (rId, type, target, external?) beyond the template's
        self.fn_rels = []    # same, but for the footnotes part
        self.media = {}      # archive name -> bytes
        self.warnings = []
        self._rid = 1000     # clear of template rIds
        self.bq = 0          # blockquote nesting depth
        self._bkid = 0       # bookmark id counter
        self._imgn = 0       # image part counter (doubles as docPr id)
        self._urlrids = {}   # hyperlink URL -> rId
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
        self.contrib_numels, self.contrib_styleels = [], []
        self.reftypes = REFTYPES | (reftypes or {})
        tdoc = etree.fromstring(self.refz.read('word/document.xml'))
        self.sectpr = deepcopy(tdoc.find(f'{{{W}}}body/{{{W}}}sectPr'))
        pg, mar = self.sectpr.find(qn('w:pgSz')), self.sectpr.find(qn('w:pgMar'))
        self.content_w = int(pg.get(qn('w:w'))) - int(mar.get(qn('w:left'))) - int(mar.get(qn('w:right')))
        self.sroot = etree.fromstring(self.refz.read('word/styles.xml'))
        for r in self.contribs: self._merge_styles(r)
        self.refstyles = {s.find(qn('w:name')).get(qn('w:val')).lower(): s.get(qn('w:styleId')) for s in self.sroot.iter(qn('w:style'))}
        if missing := [n for n in STYLE_MAP.values() if n.lower() not in self.refstyles]:
            raise ValueError(f'reference doc lacks dialect styles (map/template drift?): {missing}')
        self.hlstyles = {n.removeprefix('hl ').replace(' ', '.'): sid for n, sid in self.refstyles.items() if n.startswith('hl ')}
        self.tmplnum = (etree.fromstring(self.refz.read('word/numbering.xml')) if 'word/numbering.xml' in self.refz.namelist() else None)
        used = [int(v) for e in ([] if self.tmplnum is None else self.tmplnum.iter())
            for a in ('w:numId', 'w:abstractNumId') if (v := e.get(qn(a))) and v.lstrip('-').isdigit()]
        self._numid = self._absbase = max(used, default=-1) + 1   # num and abstract ids both offset past the template's
        if isinstance(number_headings, str):
            if number_headings not in SCHEMES: raise ValueError(f'unknown numbering scheme {number_headings!r}')
            number_headings = SCHEMES[number_headings]
        self.scheme = list(number_headings.items()) if number_headings else None
        self.headnum = None
        if number_headings and self.sroot.find(f'{{{W}}}style[@{{{W}}}styleId="Heading1"]/{{{W}}}pPr/{{{W}}}numPr') is None:
            self._numid += 1
            self.headnum = self._numid
        self._adopt_contrib_nums()

    def _merge_styles(self, ref):
        "Merge a contributor's w:style elements into self.sroot, later-wins on style id or name"
        def keys(e):
            nm = e.find(qn('w:name'))
            return {(e.get(qn('w:styleId')) or '').lower(), '' if nm is None else nm.get(qn('w:val')).lower()} - {''}
        new = (etree.fromstring(zipfile.ZipFile(ref).read('word/styles.xml')).findall(qn('w:style'))
            if isinstance(ref, Path) and ref.suffix == '.docx' or str(ref).endswith('.docx')
            else self._xml_contrib(ref) if str(ref).endswith('.xml') else theme_styles(ref))
        for s in new:
            ks = keys(s)
            for old in [o for o in self.sroot.findall(qn('w:style')) if keys(o) & ks]: self.sroot.remove(old)
            self.sroot.append(s)

    def _xml_contrib(self, ref):
        "Styles from a raw .xml contributor, stashing its abstractNum/num elements for numbering adoption"
        root = etree.parse(str(ref)).getroot()
        self.contrib_numels += root.findall(qn('w:abstractNum')) + root.findall(qn('w:num'))
        styles = root.findall(qn('w:style'))
        self.contrib_styleels += styles
        return styles

    def _adopt_contrib_nums(self):
        "Renumber .xml contributors' numbering ids past ours, remapping their styles' numPr references"
        self.xabs = [e for e in self.contrib_numels if etree.QName(e).localname == 'abstractNum']
        self.xnums = [e for e in self.contrib_numels if etree.QName(e).localname == 'num']
        amap, nmap = {}, {}
        for e in self.xabs:
            self._numid += 1
            amap[e.get(qn('w:abstractNumId'))] = self._numid
            e.set(qn('w:abstractNumId'), str(self._numid))
        for e in self.xnums:
            self._numid += 1
            nmap[e.get(qn('w:numId'))] = self._numid
            e.set(qn('w:numId'), str(self._numid))
            ref = e.find(qn('w:abstractNumId'))
            if ref is not None and ref.get(qn('w:val')) in amap: ref.set(qn('w:val'), str(amap[ref.get(qn('w:val'))]))
        for s in self.contrib_styleels:
            for nid in s.iter(qn('w:numId')):
                if nid.get(qn('w:val')) in nmap: nid.set(qn('w:val'), str(nmap[nid.get(qn('w:val'))]))

    def hlsid(self, scope):
        "Hl* style id for a dotted scope: exact, else progressively shorter prefixes (tree-sitter resolution)"
        parts = (scope or '').split('.')
        while parts:
            if sid := self.hlstyles.get('.'.join(parts)): return sid
            parts.pop()

    def rid(self):
        self._rid += 1
        return f'rId{self._rid}'
    def warn(self, msg): self.warnings.append(msg)

    # ---- inline level -------------------------------------------------------
    def rpr(self, fmt):
        "w:rPr for a formatting context dict, in CT_RPr child order; None if empty"
        kids = []
        if s := fmt.get('rstyle'): kids.append(E('w:rStyle', {'w:val': s}))
        if fmt.get('b'): kids += [E('w:b'), E('w:bCs')]
        if fmt.get('i'): kids += [E('w:i'), E('w:iCs')]
        if fmt.get('strike'): kids.append(E('w:strike'))
        if fmt.get('mark'): kids.append(E('w:highlight', {'w:val': 'yellow'}))
        if fmt.get('u'): kids.append(E('w:u', {'w:val': 'single'}))
        if v := fmt.get('vert'): kids.append(E('w:vertAlign', {'w:val': v}))
        return E('w:rPr', *kids) if kids else None

    def text_runs(self, text, fmt):
        "Runs for a text node; newlines collapse to spaces (pre-context text never comes here)"
        text = re.sub(r'\s+', ' ', text)
        if not text: return []
        t = E('w:t', text)
        if text != text.strip(): t.set(qn('xml:space'), 'preserve')
        return [E('w:r', self.rpr(fmt), t)]

    def link(self, el, fmt):
        "w:hyperlink for `a`: internal '#x' -> anchor, external -> relationship (deduped per URL); data-ref -> field"
        if _get(el, 'data-ref') is not None:
            fld = self.ref_fld(el, fmt)
            return self.ref_prefix(el, fmt) + fld
        href = _get(el, 'href')
        if not href: return self.runs(el, fmt)
        runs = self.runs(el, fmt | {'rstyle': _sid('hyperlink')})
        if href.startswith('#'): return [E('w:hyperlink', {'w:anchor': self.bkname(href[1:])}, *runs)]
        return self.external_link(href, runs)

    def external_link(self, href, runs):
        if href not in self._urlrids:
            self._urlrids[href] = self.rid()
            self.rels.append((self._urlrids[href], f'{R}/hyperlink', href, True))
        return [E('w:hyperlink', {'r:id': self._urlrids[href]}, *runs)]

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
        tgt = (_get(el, 'href') or '#')[1:]
        pre = self.res.prefix(el.to_text().strip(), tgt, ref_tokens(_get(el, 'data-ref')), plural)
        return self.text_runs(pre, fmt) if pre else []

    def ref_fld(self, el, fmt):
        """REF/PAGEREF field for a cross-reference `a`, with a cached placeholder Word replaces on update.
        Heading/paragraph targets number via `\\w`; caption targets return their bookmarked 'Label N' text
        (or the number-only `_n` bookmark for bare/leaf/rel refs), and text targets (spans, definition
        terms) their bookmark text, so `\\w` never applies to either."""
        tgt = (_get(el, 'href') or '#')[1:]
        tokens = ref_tokens(_get(el, 'data-ref'))
        self.res.check(tgt)
        variant = ref_variant(tokens)
        nm, self.has_fields = self.bkname(tgt), True
        if variant == 'page': instr, cached = rf' PAGEREF {nm} \h ', '#'
        elif self.reftarget[tgt] == 'caption':
            bare = 'bare' in tokens or variant in ('leaf', 'rel')
            instr, cached = (rf' REF {nm}_n \h ' if bare else rf' REF {nm} \h '), '#'
        elif self.reftarget[tgt] == 'text': instr, cached = rf' REF {nm} \h ', self.res.core(tgt, tokens)
        else:
            sw = self.REFSWITCH[variant]
            instr = rf' REF {nm} {sw} \h ' if sw else rf' REF {nm} \h '
            cached = self.idtext.get(tgt, '#') if variant == 'text' else '#'
        return [E('w:fldSimple', {'w:instr': instr}, E('w:r', self.rpr(fmt), E('w:t', cached)))]

    def ref_group(self, el, fmt):
        "data-refs span: one pluralized prefix for a same-type group, per-item singular prefixes for mixed types; never range-collapsed"
        refs = [c for c in _els(el) if _tag(c) == 'a']
        types = [(_get(a, 'href') or '#')[1:].split('-')[0] for a in refs]
        out = []
        for (sep, pre, plural), a in zip(group_plan(types), refs):
            if sep: out += self.text_runs(sep, fmt)
            if pre: out += self.ref_prefix(a, fmt, plural=plural)
            out += self.ref_fld(a, fmt)
        return out

    def custom_style(self, el, kind):
        "Style id for an explicit custom-style attr (stubbed + warned if undefined), else a class matching a reference style name"
        if cs := _get(el, 'custom-style'):
            if cs.lower() in self.refstyles: return self.refstyles[cs.lower()]
            if cs not in self.stubs:
                self.stubs[cs] = (kind, re.sub(r'\W', '', cs) or f'Custom{len(self.stubs)}')
                self.warn(f'custom style {cs!r} not in reference doc; stub injected')
            return self.stubs[cs][1]
        return next((self.refstyles[c.lower()] for c in _classes(el) if c.lower() in self.refstyles), None)

    def span(self, el, fmt):
        "Inline span: math -> inline m:oMath zone (linear source, dialect-agnostic), custom style -> rStyle, else transparent; an id bookmarks the runs"
        if _get(el, 'data-refs') is not None: return self.ref_group(el, fmt)
        if 'math' in _classes(el): return self.bookmark(el, [self.omath(el)])
        if sid := self.custom_style(el, 'character'): return self.bookmark(el, self.runs(el, fmt | {'rstyle': sid}))
        return self.bookmark(el, self.runs(el, fmt))

    def omath(self, el):
        "An m:oMath zone holding `el`'s text as linear-format math runs"
        return E('m:oMath', E('m:r', E('m:t', el.to_text(), {'xml:space': 'preserve'})))

    def fnref(self, el, fmt):
        "Footnote-reference run for a sup>a.footnote-ref, or None when `el` is an ordinary sup"
        children = _els(el)
        a = children[0] if len(children) == 1 and _tag(children[0]) == 'a' else None
        if a is None or 'footnote-ref' not in _classes(a): return None
        key = (_get(a, 'href') or '#')[1:]
        if key not in self.fndefs:
            self.warn(f'footnote reference #{key} has no definition; dropped')
            return []
        if key not in self.fnids:
            self.fnids[key] = len(self.fnids) + 1
            self.fnotes.append((self.fnids[key], self.fn_blocks(self.fndefs[key])))
        return [E('w:r', E('w:rPr', E('w:rStyle', {'w:val': _sid('footnoteref')})),
            E('w:footnoteReference', {'w:id': self.fnids[key]}))]

    def fn_blocks(self, li):
        "Footnote body: the li's blocks in footnote-text style, backref stripped, reference mark prepended"
        save = self.rels, self._urlrids, self.first
        self.rels, self._urlrids = self.fn_rels, {}   # rel ids are per-part (see fn_relsxml)
        try:
            blks = [b for kind, val in self.li_parts(li)  # chkstyle: ignore-node
                for b in ([self.para(self.group_runs(val, {}), 'footnotetext')] if kind == 'inline'
                    else self.block(val, 'footnotetext'))]
        finally: self.rels, self._urlrids, self.first = save
        if not blks: blks = [self.para([], 'footnotetext')]
        mark = E('w:r', E('w:rPr', E('w:rStyle', {'w:val': _sid('footnoteref')})), E('w:footnoteRef'))
        blks[0].insert(1, E('w:r', E('w:t', ' ', {'xml:space': 'preserve'})))
        blks[0].insert(1, mark)
        return blks

    def image(self, el, fmt, alt=None):
        "Embed a local image (dimensions sniffed, width/height px attrs override); remote srcs degrade to a link"
        src = _get(el, 'src') or ''
        if alt is None: alt = _get(el, 'alt') or src
        if re.match(r'[a-z][a-z0-9+.-]*://', src):
            self.warn(f'remote image not embedded: {src}')
            return self.external_link(src, self.text_runs(alt, fmt | {'rstyle': _sid('hyperlink')}))
        try: data = (self.base/src).read_bytes()
        except OSError:
            self.warn(f'image not found: {src}; alt text emitted')
            return self.text_runs(alt, fmt)
        pw, ph, dx, dy = imgsize(data) or (300, 200, 96, 96)
        cx, cy = round(pw * 914400 / dx), round(ph * 914400 / dy)
        w_, h_ = _get(el, 'width'), _get(el, 'height')
        if w_: cx = round(float(w_) * EMU_PER_PX)
        if h_: cy = round(float(h_) * EMU_PER_PX)
        if w_ and not h_: cy = round(cx * ph / pw)
        if h_ and not w_: cx = round(cy * pw / ph)
        self._imgn += 1
        name = f'word/media/image{self._imgn}{Path(src).suffix.lower() or ".bin"}'
        self.media[name] = data
        rid = self.rid()
        self.rels.append((rid, f'{R}/image', name.removeprefix('word/'), False))
        return [E('w:r', drawing(rid, self._imgn, cx, cy, alt))]

    def inline(self, el, fmt):
        "Run-level elements for inline `el` under formatting context `fmt` (dict; copied on change)"
        tag = _tag(el)
        if tag == 'em': out = self.runs(el, fmt | {'i': True})
        elif tag == 'strong': out = self.runs(el, fmt | {'b': True})
        elif tag == 'code': out = self.runs(el, fmt | {'rstyle': _sid('codeinline')})
        elif tag == 'a': out = self.link(el, fmt)
        elif tag == 'del': out = self.runs(el, fmt | {'strike': True})
        elif tag == 'mark': out = self.runs(el, fmt | {'mark': True})
        elif tag == 'u': out = self.runs(el, fmt | {'u': True})
        elif tag == 'sub': out = self.runs(el, fmt | {'vert': 'subscript'})
        elif tag == 'sup':
            fn = self.fnref(el, fmt)
            out = fn if fn is not None else self.runs(el, fmt | {'vert': 'superscript'})
        elif tag == 'span': out = self.span(el, fmt)
        elif tag == 'img': out = self.image(el, fmt)
        elif tag == 'br': out = [E('w:r', self.rpr(fmt), E('w:br'))]
        elif tag == 'input':   # task-list checkbox
            if _get(el, 'type') != 'checkbox':
                self.warn(f'unhandled inline <input type={_get(el, "type")!r}>; dropped')
                return []
            g = '☒' if _get(el, 'checked') is not None else '☐'
            out = [E('w:r', self.rpr(fmt), E('w:t', g + ' ', {'xml:space': 'preserve'}))]
        elif tag == 'script': out = self.rawxml(el)
        elif tag == 'template': out = self.tmpl_runs(el, fmt, 'inline')
        elif tag in INERT_TAGS: out = []
        else:  # unknown inline (abbr etc): recurse transparently
            out = self.runs(el, fmt)
        return out

    def runs(self, el, fmt):
        "Run-level elements for an element's ordered text and element children"
        return [r for node in el.children for r in self.inline_node(node, fmt)]

    def inline_node(self, node, fmt):
        if isinstance(node, Text): return self.text_runs(node.text, fmt)
        if isinstance(node, Element):
            if _tag(node) == 'a' and 'footnote-backref' in _classes(node): return []
            return self.inline(node, fmt)
        return []

    # ---- block level --------------------------------------------------------
    def para(self, runs, style='body', extra=None, sid=None):
        "A w:p with `style` (STYLE_MAP key, or `sid` style-id override) and optional extra pPr children (schema order!)"
        ppr = E('w:pPr', E('w:pStyle', {'w:val': sid or _sid(style)}), *(extra or []))
        return E('w:p', ppr, *runs)

    def bookmark(self, el, runs):
        "Wrap `runs` in a bookmark when `el` carries an id (target for internal links)"
        if not (i := _get(el, 'id')): return runs
        self._bkid += 1
        return [E('w:bookmarkStart', {'w:id': self._bkid, 'w:name': self.bkname(i)}),
            *runs, E('w:bookmarkEnd', {'w:id': self._bkid})]

    def codeblock(self, el):
        "Source Code paragraph, lines joined with w:br; Hl* character styles when a language class names one"
        children = _els(el)
        code = children[0] if children and _tag(children[0]) == 'code' else el
        lang = next((c.removeprefix('language-') for c in _classes(code) if c.startswith('language-')), None)
        text = code.to_text().rstrip('\n')
        segs = (segments(text, lang) if self.hlstyles else None) or [(text, None)]
        runs = []
        for txt, scope in segs:
            for j, part in enumerate(txt.split('\n')):
                if j: runs.append(E('w:r', E('w:br')))
                if not part: continue
                sid = self.hlsid(scope)
                runs.append(E('w:r', E('w:rPr', E('w:rStyle', {'w:val': sid})) if sid else None,
                    E('w:t', part, {'xml:space': 'preserve'})))
        return [self.para(runs, 'codeblock')]

    def qindent(self):
        "Extra indent for paragraphs in nested blockquotes (Quote style itself carries the first level)"
        return [E('w:ind', {'w:left': 720 * self.bq})] if self.bq > 1 else None

    # ---- lists --------------------------------------------------------------
    def list_el(self, el, ilvl=0):
        "A ul/ol: fresh num instance (so each ordered list restarts), items at level `ilvl`"
        self._numid += 1
        nid = self._numid
        self.nums.append((nid, 0 if _tag(el) == 'ul' else 1, int(_get(el, 'start', 1))))
        return [b for li in _els(el) if _tag(li) == 'li' for b in self.li(li, nid, min(ilvl, 8))]

    def li_parts(self, el):
        "Split mixed li content into ('inline', [nodes]) groups and ('block', child) items, in order"
        parts = []
        def add(x):
            if isinstance(x, Text) and not x.text.strip() and '\n' in x.text: return
            if parts and parts[-1][0] == 'inline': parts[-1][1].append(x)
            else: parts.append(('inline', [x]))
        for node in el.children:
            if isinstance(node, Element) and _tag(node) in BLOCK_TAGS: parts.append(('block', node))
            elif isinstance(node, (Text, Element)): add(node)
        return parts

    def group_runs(self, nodes, fmt):
        "Runs for a mixed list of text and inline nodes"
        return [r for node in nodes for r in self.inline_node(node, fmt)]

    def li(self, li, nid, ilvl):
        "Blocks for one list item: the first paragraph carries the number, the rest continue indented"
        numpr = [E('w:numPr', E('w:ilvl', {'w:val': ilvl}), E('w:numId', {'w:val': nid}))]
        cont = [E('w:ind', {'w:left': 720 * (ilvl + 1)})]
        out = []
        for kind, val in self.li_parts(li):
            if kind == 'inline': out.append(self.para(self.group_runs(val, {}), 'list', numpr if not out else cont))
            elif _tag(val) in ('ul', 'ol'): out += self.list_el(val, ilvl + 1)
            elif _tag(val) == 'p': out.append(self.para(self.runs(val, {}), 'list', numpr if not out else cont))
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
            for cell in _els(tr):
                if _tag(cell) not in ('td', 'th'): continue
                cs, rs = int(_get(cell, 'colspan', 1)), int(_get(cell, 'rowspan', 1))
                rowcells.append(('cell', ci, cell, cs, rs))
                for k in range(1, rs): spans[(ri + k, ci)] = cs
                ci = _skip(ci + cs)
            placed.append(rowcells)
            ncols = max(ncols, ci)
        return placed, ncols

    def col_widths(self, el, ncols):
        "colwidths tracks -> (dxa list, all_fr?) or (None, False) when absent"
        s = _get(el, 'colwidths') or _get(el, 'data-colwidths')
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
        "Block content of one table cell; header cells bold, align attr honored for inline cells"
        if any(_tag(c) in BLOCK_TAGS for c in _els(cell)): return self.blocks(cell)
        jc = [E('w:jc', {'w:val': _get(cell, 'align')})] if _get(cell, 'align') in ('center', 'right') else None
        return [self.para(self.runs(cell, {'b': True} if header else {}), 'compact', jc)]

    def table(self, el):
        "w:tbl (+ caption paragraph before, spacer paragraph after)"
        cap = None
        rows, nhead, markers = [], 0, {}
        def _add(c):
            "Rows in order; template markers recorded at the row index they precede"
            if _tag(c) == 'template': markers.setdefault(len(rows), []).append(c)
            else: rows.append(c)
        for sec in _els(el):
            t = _tag(sec)
            if t == 'caption': cap = sec
            elif t == 'thead':
                for c in _els(sec): _add(c)
                nhead = len(rows)
            elif t in ('tbody', 'tfoot'):
                for c in _els(sec): _add(c)
            elif t in ('tr', 'template'): _add(sec)
        placed, ncols = self.table_grid(rows)
        dxa, all_fr = self.col_widths(el, ncols)
        def _tcw(ci, cs):
            if not dxa: return None
            wd = sum(dxa[ci:ci + cs])
            if all_fr: return E('w:tcW', {'w:type': 'pct', 'w:w': round(wd / self.content_w * 5000)})
            return E('w:tcW', {'w:type': 'dxa', 'w:w': wd})
        tblw = (E('w:tblW', {'w:type': 'auto', 'w:w': 0}) if not dxa  # chkstyle: ignore-node
            else E('w:tblW', {'w:type': 'pct', 'w:w': 5000}) if all_fr
            else E('w:tblW', {'w:type': 'dxa', 'w:w': sum(dxa)}))
        tblpr = E('w:tblPr', E('w:tblStyle', {'w:val': self.custom_style(el, 'table') or _sid('table')}), tblw,  # chkstyle: ignore-node
            E('w:tblLayout', {'w:type': 'fixed'}) if dxa and not all_fr else None,
            E('w:tblLook', {'w:val': '04A0', 'w:firstRow': 1, 'w:lastRow': 0,
                'w:firstColumn': 0, 'w:lastColumn': 0, 'w:noHBand': 0, 'w:noVBand': 1}))
        gw = dxa or [self.content_w // ncols] * ncols   # pandoc's docx reader drops tables whose gridCols lack w:w
        grid = E('w:tblGrid', *[E('w:gridCol', {'w:w': gw[i]}) for i in range(ncols)])
        def _marker_tr(mel):
            "A range marker between rows: one full-width literal cell, so forms keep their markers visible"
            tcpr = E('w:tcPr', _tcw(0, ncols), E('w:gridSpan', {'w:val': ncols}) if ncols > 1 else None)
            return E('w:tr', E('w:tc', tcpr, self.para(self.tmpl_runs(mel, {}, 'row'))))
        trs = []
        for ri, rowcells in enumerate(placed):
            for mel in markers.get(ri, []): trs.append(_marker_tr(mel))
            tcs = []
            for item in rowcells:
                if item[0] == 'cont':
                    _, ci, wd = item
                    tcs.append(E('w:tc', E('w:tcPr', _tcw(ci, wd),  # chkstyle: ignore-node
                        E('w:gridSpan', {'w:val': wd}) if wd > 1 else None,
                        E('w:vMerge')), E('w:p')))
                else:
                    _, ci, cell, cs, rs = item
                    tcpr = E('w:tcPr', _tcw(ci, cs),  # chkstyle: ignore-node
                        E('w:gridSpan', {'w:val': cs}) if cs > 1 else None,
                        E('w:vMerge', {'w:val': 'restart'}) if rs > 1 else None)
                    body = self.cell_blocks(cell, ri < nhead)
                    if not len(body) or etree.QName(body[-1]).localname != 'p': body.append(E('w:p'))
                    tcs.append(E('w:tc', tcpr, *body))
            trs.append(E('w:tr', E('w:trPr', E('w:tblHeader')) if ri < nhead else None, *tcs))
        for mel in markers.get(len(rows), []): trs.append(_marker_tr(mel))
        out = self.caption_para(el, 'tbl', cap)
        return out + [E('w:tbl', tblpr, grid, *trs), E('w:p')]

    def caption_para(self, el, typ, capel, fmt={}):
        """Numbered caption paragraph: 'Label N: text' with a SEQ field as N. When `el` has an id, the
        label+number span is bookmarked under it (REF target) and the number alone under `<name>_n`.
        Emitted whenever there is a caption or an id; the label word comes from reftypes[typ]."""
        if capel is None and not _get(el, 'id'): return []
        label = self.reftypes[typ][0]
        seq = [E('w:fldSimple', {'w:instr': rf' SEQ {label} \* ARABIC '}, E('w:r', self.rpr(fmt), E('w:t', '#')))]
        self.has_fields = True
        if i := _get(el, 'id'):
            nm = self.bkname(i)
            self._bknames[i + '\0n'] = nm + '_n'   # reserve the number-only name too
            self._bkid += 2
            seq = [E('w:bookmarkStart', {'w:id': self._bkid, 'w:name': nm + '_n'}), *seq,
                E('w:bookmarkEnd', {'w:id': self._bkid})]
            runs = [E('w:bookmarkStart', {'w:id': self._bkid - 1, 'w:name': nm}),  # chkstyle: ignore-node
                *self.text_runs(label + ' ', fmt), *seq,
                E('w:bookmarkEnd', {'w:id': self._bkid - 1})]
        else: runs = self.text_runs(label + ' ', fmt) + seq
        cap = [] if capel is None else self.runs(capel, fmt)
        if cap: runs += self.text_runs(': ', fmt) + cap
        return [self.para(runs, 'caption')]

    def figure(self, el):
        "Figure: image paragraph, then its numbered caption paragraph below (Word convention)"
        img = next((c for c in _walk(el) if _tag(c) == 'img'), None)
        capel = next((c for c in _els(el) if _tag(c) == 'figcaption'), None)
        alt = capel.to_text().strip() if capel is not None else None
        out = [] if img is None else [self.para(self.image(img, {}, alt), 'body')]
        return out + self.caption_para(el, 'fig', capel)

    # Paragraphs directly after these blocks (or at document start) take First Paragraph rather
    # than Body Text, matching pandoc's docx writer exactly, so the two agree on which is "first".
    FIRST_AFTER = {'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'ul', 'ol', 'table', 'dl', 'hr'}

    def block(self, el, style='body', sid=None):
        "Block elements for `el` (one element may yield several); `sid` is a custom-style id override for paragraphs"
        out = self._block(el, style, sid)
        tag = _tag(el)
        if tag in self.FIRST_AFTER or (tag == 'div' and 'display' in _classes(el)): self.first = True
        return out

    def _block(self, el, style, sid):
        tag = _tag(el)
        if tag == 'p':
            ex = self.qindent() if style == 'blockquote' else None
            psid = self.custom_style(el, 'paragraph') or sid
            use = 'firstpara' if self.first and style == 'body' and not psid else style
            self.first = False
            return [self.para(self.bookmark(el, self.runs(el, {})), use, ex, psid)]
        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'): return [self.para(self.bookmark(el, self.runs(el, {})), tag)]
        if tag == 'blockquote':
            self.bq += 1
            try: return self.blocks(el, 'blockquote')
            finally: self.bq -= 1
        if tag == 'pre': return self.codeblock(el)
        if tag in ('ul', 'ol'): return self.list_el(el)
        if tag == 'table': return self.table(el)
        if tag == 'hr':
            return [E('w:p', E('w:pPr', E('w:pBdr',
                E('w:bottom', {'w:val': 'single', 'w:sz': 6, 'w:space': 1, 'w:color': 'auto'}))))]
        if tag == 'dl': return self.dl(el)
        if tag == 'script': return self.rawxml(el)
        if tag == 'figure': return self.figure(el)
        if tag == 'template': return [self.para(runs)] if (runs := self.tmpl_runs(el, {}, 'block')) else []
        if tag == 'div':
            cls = _classes(el)
            if 'math' in cls and 'display' in cls: return [E('w:p', E('m:oMathPara', self.omath(el)))]
            if 'details' in cls and (kids := _els(el)) and _tag(kids[0]) in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
                label = self.para(self.runs(kids[0], {'b': True}), style)   # the dialect's collapsible block: label as a bold line, never a numbered heading
                return [label, *self.block_nodes(kids[1:], style, sid)]
            return self.blocks(el, style, self.custom_style(el, 'paragraph') or sid)
        if tag in BLOCK_TAGS and any(_tag(c) in BLOCK_TAGS for c in _els(el)):
            return self.blocks(el, style, sid)   # unknown container: recurse
        self.warn(f'unhandled block <{tag}>; emitted as plain paragraph')
        return [self.para(self.runs(el, {}), style, None, sid)]

    RAWNS = ' '.join(f'xmlns:{k}="{v}"' for k, v in NS.items() if k != 'xml')

    BIND_NS = 'urn:mdhtml:fields'
    BIND_ID = '{8E2C9A44-7D31-4E5B-9C0D-1A6F2B3C4D5E}'   # fixed datastore id, so builds are reproducible

    def tmpl_runs(self, el, fmt, form):
        """Template-token runs. Range markers and unknown tokens are converter policy: literal
        `«body»` runs, so an unfilled form shows its markers. Var tokens go through the `tmpl`
        callable with the token node dict (see `mdhtml.export.tmpl_node`): str is a literal text
        run, ('field', instr) a live field, ('control', name) an interactive plain-text content
        control, ('bound', name) a data-bound one, None dropped"""
        node = tmpl_node(el, form)
        if node["kind"] != "var": return self.text_runs(f'«{node["body"].strip()}»', fmt)
        if self.tmpl is None: return []
        res = self.tmpl(node)
        if res is None: return []
        if isinstance(res, str): return self.text_runs(res, fmt)
        kind, val = res
        if kind == 'field':
            self.has_fields = True
            return [E('w:fldSimple', {'w:instr': f' {val.strip()} '}, E('w:r', self.rpr(fmt), E('w:t', f'«{node["name"]}»')))]
        if kind in ('control', 'bound'):
            self.has_controls = True
            sdtpr = E('w:sdtPr', E('w:alias', {'w:val': val}), E('w:tag', {'w:val': val}), E('w:showingPlcHdr'))
            if kind == 'bound':
                if val not in self.bound: self.bound.append(val)
                sdtpr.append(E('w:dataBinding', {'w:prefixMappings': f"xmlns:ns0='{self.BIND_NS}'",
                    'w:xpath': f'/ns0:fields[1]/ns0:{val}[1]', 'w:storeItemID': self.BIND_ID}))
            sdtpr.append(E('w:text'))
            return [E('w:sdt', sdtpr,
                E('w:sdtContent', E('w:r', self.rpr(fmt | {'rstyle': 'PlaceholderText'}), E('w:t', val))))]
        raise ValueError(f'unknown template rendering {res!r}')


    def rawxml(self, el):
        "Elements parsed from a raw docx payload (`{=docx}` in Markdown); other formats skip silently"
        if _get(el, 'type') != 'application/vnd.mdhtml.raw' or _get(el, 'data-format') != 'docx': return []
        payload, warn = decode_raw(el)
        if warn:
            self.warn(warn)
            return []
        try: return list(etree.fromstring(f'<m2d {self.RAWNS}>{payload}</m2d>'))
        except etree.XMLSyntaxError as e:
            self.warn(f'malformed docx raw payload: {e}')
            return []

    def dl(self, el):
        "Definition list: dt/dd paragraphs in their dialect styles"
        out = []
        for c in _els(el):
            t = _tag(c)
            if t == 'dt': out.append(self.para(self.bookmark(c, self.runs(c, {})), 'dt'))
            elif t == 'dd':
                blocky = any(_tag(k) in BLOCK_TAGS for k in _els(c))
                out += self.blocks(c, 'dd') if blocky else [self.para(self.runs(c, {}), 'dd')]
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
            else:
                extra = self.qindent() if style == 'blockquote' else None
                use = 'firstpara' if self.first and style == 'body' and not sid else style
                self.first = False
                out.append(self.para(self.group_runs(inline, {}), use, extra, sid))
            inline.clear()
        for node in nodes:
            if isinstance(node, Comment): continue
            if isinstance(node, Element) and _tag(node) == 'template':
                flush()
                if runs := self.tmpl_runs(node, {}, 'block'): out.append(self.para(runs))
                continue
            if isinstance(node, Element) and (_tag(node) in INERT_TAGS or _tag(node) == 'script' and not _is_raw(node)): continue
            if isinstance(node, Element) and _tag(node) in BLOCK_TAGS:
                flush()
                out.extend(self.block(node, style, sid))
            elif isinstance(node, (Text, Element)): inline.append(node)
        flush()
        return out

    # ---- assembly -----------------------------------------------------------
    def document(self, body_blocks):
        "word/document.xml bytes: our blocks + the template's sectPr"
        root = etree.Element(qn('w:document'), nsmap=NS)
        body = etree.SubElement(root, qn('w:body'))
        for b in body_blocks: body.append(b)
        body.append(self.sectpr)
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    BULLETS = ['•', '◦', '▪']
    NUMFMTS = ['decimal', 'lowerLetter', 'lowerRoman']

    def numbering_xml(self):
        """word/numbering.xml: the template's part (if any) merged with our list definitions (bullet and
        decimal abstracts, one num per list with startOverrides so each restarts), the heading scheme,
        and .xml contributors' definitions - all abstractNum before all num, as the schema requires"""
        root = self.tmplnum if self.tmplnum is not None else etree.Element(qn('w:numbering'), nsmap=dict(w=W))
        tnums = [e for e in root if etree.QName(e).localname == 'num']
        for e in tnums: root.remove(e)
        if self.nums:
            for aid in (0, 1):
                an = E('w:abstractNum', {'w:abstractNumId': self._absbase + aid},
                    E('w:multiLevelType', {'w:val': 'hybridMultilevel'}))
                for i in range(9):
                    fmt, txt = ('bullet', self.BULLETS[i % 3]) if aid == 0 else (self.NUMFMTS[i % 3], f'%{i+1}.')
                    an.append(E('w:lvl', {'w:ilvl': i},  # chkstyle: ignore-node
                        E('w:start', {'w:val': 1}), E('w:numFmt', {'w:val': fmt}),
                        E('w:lvlText', {'w:val': txt}), E('w:lvlJc', {'w:val': 'left'}),
                        E('w:pPr', E('w:ind', {'w:left': 720 * (i + 1), 'w:hanging': 360}))))
                root.append(an)
        if self.headnum:
            an = E('w:abstractNum', {'w:abstractNumId': self._absbase + 2}, E('w:multiLevelType', {'w:val': 'multilevel'}))
            for i in range(9):
                txt, fmt = self.scheme[i] if i < len(self.scheme) else (f'%{i+1}.', 'decimal')
                an.append(E('w:lvl', {'w:ilvl': i},  # chkstyle: ignore-node
                    E('w:start', {'w:val': 1}), E('w:numFmt', {'w:val': fmt}),
                    E('w:pStyle', {'w:val': f'Heading{i + 1}'}) if i < 6 else None,
                    E('w:lvlText', {'w:val': txt}), E('w:lvlJc', {'w:val': 'left'})))
            root.append(an)
        for e in self.xabs: root.append(e)
        for e in tnums: root.append(e)
        if self.headnum: root.append(E('w:num', {'w:numId': self.headnum}, E('w:abstractNumId', {'w:val': self._absbase + 2})))
        for e in self.xnums: root.append(e)
        for nid, aid, start in self.nums:
            root.append(E('w:num', {'w:numId': nid}, E('w:abstractNumId', {'w:val': self._absbase + aid}),  # chkstyle: ignore-node
                *[E('w:lvlOverride', {'w:ilvl': i}, E('w:startOverride', {'w:val': start if i == 0 else 1}))
                    for i in range(9)]))
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    RNS = 'http://schemas.openxmlformats.org/package/2006/relationships'

    def _add_rels(self, root, rels):
        for rid, typ, target, ext in rels:
            rel = etree.SubElement(root, f'{{{self.RNS}}}Relationship', Id=rid, Type=typ, Target=target)
            if ext: rel.set('TargetMode', 'External')
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def relsxml(self):
        "word/_rels/document.xml.rels: template rels plus ours"
        return self._add_rels(etree.fromstring(self.refz.read('word/_rels/document.xml.rels')), self.rels)

    def fn_relsxml(self):
        "word/_rels/footnotes.xml.rels: relationship ids are per-part, so footnote links/images get their own file"
        return self._add_rels(etree.Element(f'{{{self.RNS}}}Relationships', nsmap={None: self.RNS}), self.fn_rels)

    DSNS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml'

    def bind_item_xml(self):
        "customXml/item1.xml: one empty element per bound variable; every same-name control is a live view of it"
        root = etree.Element(f'{{{self.BIND_NS}}}fields', nsmap=dict(ns0=self.BIND_NS))
        for name in self.bound: etree.SubElement(root, f'{{{self.BIND_NS}}}{name}')
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def bind_props_xml(self):
        "customXml/itemProps1.xml: the datastore id that the controls' `storeItemID` points at"
        root = etree.Element(f'{{{self.DSNS}}}datastoreItem', nsmap=dict(ds=self.DSNS))
        root.set(f'{{{self.DSNS}}}itemID', self.BIND_ID)
        srs = etree.SubElement(root, f'{{{self.DSNS}}}schemaRefs')
        etree.SubElement(srs, f'{{{self.DSNS}}}schemaRef').set(f'{{{self.DSNS}}}uri', self.BIND_NS)
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def bind_rels_xml(self):
        "customXml/_rels/item1.xml.rels: the item's link to its datastore properties"
        root = etree.Element(f'{{{self.RNS}}}Relationships', nsmap={None: self.RNS})
        return self._add_rels(root, [('rId1', f'{R}/customXmlProps', 'itemProps1.xml', False)])


    def footnotes_xml(self):
        "word/footnotes.xml: the two Word-required separator notes plus our harvested ones"
        root = etree.Element(qn('w:footnotes'), nsmap=dict(w=W, r=NS['r']))
        for typ, wid in (('separator', -1), ('continuationSeparator', 0)):
            root.append(E('w:footnote', {'w:type': typ, 'w:id': wid},
                E('w:p', E('w:pPr', E('w:spacing', {'w:after': 0})), E('w:r', E(f'w:{typ}')))))
        for wid, blks in self.fnotes: root.append(E('w:footnote', {'w:id': wid}, *blks))
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def content_types(self, extra_parts):
        "[Content_Types].xml with defaults for media extensions and overrides for added parts"
        CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
        root = etree.fromstring(self.refz.read('[Content_Types].xml'))
        have = {d.get('Extension') for d in root if d.tag == f'{{{CT}}}Default'}
        MIME = dict(png='image/png', jpeg='image/jpeg', jpg='image/jpeg', gif='image/gif', tiff='image/tiff')
        for name in self.media:
            ext = posixpath.splitext(name)[1][1:].lower()
            if ext not in have:
                etree.SubElement(root, f'{{{CT}}}Default', Extension=ext, ContentType=MIME.get(ext, 'application/octet-stream'))
                have.add(ext)
        WPML = 'application/vnd.openxmlformats-officedocument.wordprocessingml'
        for part, kind in extra_parts:
            if any(d.get('PartName') == '/' + part for d in root if d.tag == f'{{{CT}}}Override'): continue
            etree.SubElement(root, f'{{{CT}}}Override', PartName='/'+part, ContentType=f'{WPML}.{kind}+xml')
        if self.bound:
            if 'xml' not in have: etree.SubElement(root, f'{{{CT}}}Default', Extension='xml', ContentType='application/xml')
            pn = '/customXml/itemProps1.xml'
            if not any(d.get('PartName') == pn for d in root if d.tag == f'{{{CT}}}Override'):
                etree.SubElement(root, f'{{{CT}}}Override', PartName=pn,
                    ContentType='application/vnd.openxmlformats-officedocument.customXmlProperties+xml')
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    PPR_PRE_NUMPR = {'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl'}

    def _number_heading_styles(self):
        "Patch w:numPr into Heading1-6 styles, binding them to the generated heading numbering"
        for i in range(6):
            st = self.sroot.find(f'{{{W}}}style[@{{{W}}}styleId="Heading{i + 1}"]')
            if st is None: continue
            ppr = st.find(qn('w:pPr'))
            if ppr is None:
                ppr = E('w:pPr')
                rpr = st.find(qn('w:rPr'))
                st.insert(list(st).index(rpr) if rpr is not None else len(st), ppr)
            np = E('w:numPr', E('w:ilvl', {'w:val': i}), E('w:numId', {'w:val': self.headnum}))
            pos = next((j for j, c in enumerate(ppr) if etree.QName(c).localname not in self.PPR_PRE_NUMPR), len(ppr))
            ppr.insert(pos, np)

    def styles_xml(self):
        "word/styles.xml: the merged reference styles, plus stub definitions for undefined custom styles (pandoc's mechanism, plus our warning)"
        if self.headnum: self._number_heading_styles()
        root = self.sroot
        for name, (kind, sid) in self.stubs.items():
            root.append(E('w:style', {'w:type': kind, 'w:customStyle': 1, 'w:styleId': sid},  # chkstyle: ignore-node
                E('w:name', {'w:val': name}),
                E('w:basedOn', {'w:val': 'BodyText' if kind == 'paragraph' else 'DefaultParagraphFont'}),
                E('w:qFormat')))
        if self.has_controls and 'placeholder text' not in self.refstyles:
            root.append(E('w:style', {'w:type': 'character', 'w:styleId': 'PlaceholderText'},  # chkstyle: ignore-node
                E('w:name', {'w:val': 'Placeholder Text'}),
                E('w:basedOn', {'w:val': 'DefaultParagraphFont'}), E('w:semiHidden'),
                E('w:rPr', E('w:color', {'w:val': '808080'}))))
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def harvest_footnotes(self, els):
        "Split out footnote endnote sections, indexing their li definitions by id; returns body elements"
        body, fn = [], []
        for el in els:
            is_notes = isinstance(el, Element) and _tag(el) == 'section' and 'footnotes' in _classes(el)
            (fn if is_notes else body).append(el)
        self.fndefs.update({_get(li, 'id'): li for sec in fn for li in _walk(sec) if _tag(li) == 'li' and _get(li, 'id')})
        return body

    def to_docx(self, mdhtml, dest):
        root = parse_frag(mdhtml)
        nodes = self.harvest_footnotes(root.children)
        self.idtext, self.reftarget, self.res = {}, {}, Resolver(self.reftypes)
        for e in (e for node in nodes if isinstance(node, Element) for e in _walk(node)):
            if not (i := _get(e, 'id')): continue
            self.idtext[i] = ' '.join(e.to_text().split())
            k = target_kind(_tag(e))
            if k: self.reftarget[i] = k
            self.res.register(i, k, self.idtext[i])
        self.ids = set(self.idtext)
        blocks = self.block_nodes(nodes)
        docxml = self.document(blocks)
        parts = {}   # archive name -> (bytes, content-type kind); each also gets a document rel
        if self.nums or self.headnum or self.xnums:
            parts['word/numbering.xml'] = (self.numbering_xml(), 'numbering')
            if self.tmplnum is None: self.rels.append((self.rid(), f'{R}/numbering', 'numbering.xml', False))
        if self.fnotes:
            parts['word/footnotes.xml'] = (self.footnotes_xml(), 'footnotes')
            self.rels.append((self.rid(), f'{R}/footnotes', 'footnotes.xml', False))
        if self.bound: self.rels.append((self.rid(), f'{R}/customXml', '../customXml/item1.xml', False))
        extra = [(name, kind) for name, (data, kind) in parts.items()]
        with zipfile.ZipFile(dest, 'w', zipfile.ZIP_DEFLATED) as zo:
            for i in self.refz.infolist():
                if i.filename == 'word/document.xml': zo.writestr(i.filename, docxml)
                elif i.filename in parts: pass   # replaced below (e.g. a template numbering.xml we merged)
                elif i.filename == 'word/_rels/document.xml.rels': zo.writestr(i.filename, self.relsxml())
                elif i.filename == '[Content_Types].xml': zo.writestr(i.filename, self.content_types(extra))
                elif i.filename == 'word/styles.xml': zo.writestr(i.filename, self.styles_xml())
                elif i.filename == 'word/settings.xml' and self.has_fields: zo.writestr(i.filename, self.settings_xml())
                else: zo.writestr(i.filename, self.refz.read(i.filename))
            for name, (data, kind) in parts.items(): zo.writestr(name, data)
            if self.fn_rels: zo.writestr('word/_rels/footnotes.xml.rels', self.fn_relsxml())
            if self.bound:
                zo.writestr('customXml/item1.xml', self.bind_item_xml())
                zo.writestr('customXml/itemProps1.xml', self.bind_props_xml())
                zo.writestr('customXml/_rels/item1.xml.rels', self.bind_rels_xml())
            for name, data in self.media.items(): zo.writestr(name, data)
        return self.warnings

    SETT_AFTER_UPDATE = {'footnotePr', 'endnotePr', 'compat', 'rsids', 'mathPr', 'attachedSchema',  # chkstyle: ignore-node
        'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats',
        'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown',
        'smartTagType', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol',
        'listSeparator', 'docId', 'discardImageEditingData', 'defaultImageDpi',
        'docVars', 'chartTrackingRefBased'}

    def settings_xml(self):
        "word/settings.xml with w:updateFields added (schema-ordered), so Word refreshes REF fields on open"
        root = etree.fromstring(self.refz.read('word/settings.xml'))
        pos = next((i for i, c in enumerate(root) if etree.QName(c).localname in self.SETT_AFTER_UPDATE), len(root))
        root.insert(pos, E('w:updateFields', {'w:val': 'true'}))
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def mustache_fields(node):
    "Template variables as live Word `MERGEFIELD`s (markers never reach `tmpl`: the converter shows them literally)"
    return 'field', f'MERGEFIELD {node["name"]}'



def convert(mdhtml, dest, reference=None, base=None, reftypes=None, number_headings=None, tmpl=None):
    """Convert an MDHTML string or mutable fast5ever DOM to a docx file at `dest`; returns warnings.
    `reference` is a reference docx path, or a list of them: the first supplies the whole archive
    (default, or when None: the built-in template), later entries contribute styles only, later-wins -
    each a .docx path, a raw styles/numbering .xml path, or a fastpylight theme name (whose Hl*/Source
    Code styles are generated; see styles.theme_ref). Default adds 'github_light' when fastpylight is
    installed, so code blocks are colored; pass a bare reference for plain code. Relative image srcs
    resolve against `base` ('.'). Cross-references (`data-ref` anchors from Markdown `[@sec-x]`) become
    live REF fields; `reftypes` maps type tokens to (singular, plural) prefix words beyond the built-in
    `sec`, and `number_headings` (a styles.SCHEMES name such as 'legal', or a {lvlText: numFmt} dict, one entry per heading level)
    numbers the headings via a multilevel list so `\\w` fields resolve. Template tokens are dropped
    unless `tmpl` is given: a callable taking the token node dict (`mdhtml.export.tmpl_node`: `body`,
    `syntax`, `form`, `kind`, `name`, `inverted`) and returning a str for a literal text run,
    `('field', instr)` for a live field, `('control', name)` for an interactive plain-text content
    control, `('bound', name)` for a content control data-bound to a shared per-variable XML node
    (same-name controls stay in sync as one is filled), or None to drop; range markers never reach `tmpl` and render as literal «body» runs - `mustache_fields` here
    is the ready-made recipe."""
    return Converter(reference, base, reftypes, number_headings, tmpl).to_docx(mdhtml, dest)
