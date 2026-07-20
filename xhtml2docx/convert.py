"""Convert xhtmlmd XHTML fragments to docx.

Write-only, reference-archive architecture: the reference template supplies styles/theme/fonts,
we generate word/document.xml (plus footnotes/numbering/media parts as needed) into a copy of its
archive. Block and inline walkers mirror the xhtmlmd element inventory; STYLE_MAP names every
style we emit."""
import posixpath, re, zipfile
from copy import deepcopy
from pathlib import Path
from lxml import etree
from .styles import SCHEMES, STYLE_MAP, style_id, theme_styles
from .styles import ref_path as _refpath
from .wml import *
from .wml import qn
from .hilite import segments, tokenize

__all__ = ['convert']

def _sid(key): return style_id(STYLE_MAP[key])

BLOCK_TAGS = {'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'ul', 'ol', 'pre', 'table',
              'dl', 'div', 'hr', 'section', 'figure'}

def parse_frag(xhtml):
    "Parse an XHTML fragment (any number of top-level elements) into a list of elements"
    try: return list(etree.fromstring(f'<root>{xhtml}</root>'))
    except etree.XMLSyntaxError as e:
        raise ValueError(f'input is not well-formed XML ({e}); raw HTML in markdown needs '
                         'xhtmlmd.to_xhtml(..., balance=True)') from e

class Converter:
    def __init__(self, reference=None, base=None, reftypes=None, number_headings=None):
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
        self.contrib_numels, self.contrib_styleels = [], []
        self.reftypes = (dict(sec=('Section', 'Sections'), fig=('Figure', 'Figures'),
                              tbl=('Table', 'Tables')) | (reftypes or {}))
        tdoc = etree.fromstring(self.refz.read('word/document.xml'))
        self.sectpr = deepcopy(tdoc.find(f'{{{W}}}body/{{{W}}}sectPr'))
        pg, mar = self.sectpr.find(qn('w:pgSz')), self.sectpr.find(qn('w:pgMar'))
        self.content_w = int(pg.get(qn('w:w'))) - int(mar.get(qn('w:left'))) - int(mar.get(qn('w:right')))
        self.sroot = etree.fromstring(self.refz.read('word/styles.xml'))
        for r in self.contribs: self._merge_styles(r)
        self.refstyles = {s.find(qn('w:name')).get(qn('w:val')).lower(): s.get(qn('w:styleId'))
                          for s in self.sroot.iter(qn('w:style'))}
        if missing := [n for n in STYLE_MAP.values() if n.lower() not in self.refstyles]:
            raise ValueError(f'reference doc lacks dialect styles (map/template drift?): {missing}')
        self.hlstyles = {n.removeprefix('hl ').replace(' ', '.'): sid
                         for n, sid in self.refstyles.items() if n.startswith('hl ')}
        self.tmplnum = (etree.fromstring(self.refz.read('word/numbering.xml'))
                        if 'word/numbering.xml' in self.refz.namelist() else None)
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
        if el.get('data-ref') is not None:
            fld = self.ref_fld(el, fmt)
            return self.ref_prefix(el, fmt) + fld
        href = el.get('href')
        if not href: return self.runs(el, fmt)
        runs = self.runs(el, fmt | {'rstyle': _sid('hyperlink')})
        if href.startswith('#'): return [E('w:hyperlink', {'w:anchor': self.bkname(href[1:])}, *runs)]
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
        "Literal runs before a reference field: override text, the type prefix word, or nothing for bare and caption refs"
        pre = ''.join(el.itertext()).strip()
        if not pre:
            tgt = (el.get('href') or '#')[1:]
            if el.get('data-ref') == 'bare' or self.reftarget.get(tgt) == 'caption': return []
            t = tgt.split('-')[0]
            if t not in self.reftypes:
                raise ValueError(f'unknown reference type {t!r}; pass reftypes= to define its prefix')
            pre = self.reftypes[t][plural]
        return self.text_runs(pre + ' ', fmt)

    def ref_fld(self, el, fmt):
        """REF/PAGEREF field for a cross-reference `a`, with a cached placeholder Word replaces on update.
        Heading/paragraph targets number via `\\w`; caption targets return their bookmarked 'Label N' text
        (or the number-only `_n` bookmark for bare/leaf/rel refs), so `\\w` never applies to them."""
        tgt = (el.get('href') or '#')[1:]
        if tgt not in self.reftarget:
            raise ValueError(f'cross-reference target #{tgt} not found (targets are headings, paragraphs, figures, and tables with ids)')
        kind = el.get('ref', 'full')
        if kind not in self.REFSWITCH: raise ValueError(f'unknown ref= variant {kind!r}')
        nm, self.has_fields = self.bkname(tgt), True
        if kind == 'page': instr, cached = rf' PAGEREF {nm} \h ', '#'
        elif self.reftarget[tgt] == 'caption':
            bare = el.get('data-ref') == 'bare' or kind in ('leaf', 'rel')
            instr, cached = (rf' REF {nm}_n \h ' if bare else rf' REF {nm} \h '), '#'
        else:
            sw = self.REFSWITCH[kind]
            instr = rf' REF {nm} {sw} \h ' if sw else rf' REF {nm} \h '
            cached = self.idtext.get(tgt, '#') if kind == 'text' else '#'
        return [E('w:fldSimple', {'w:instr': instr}, E('w:r', self.rpr(fmt), E('w:t', cached)))]

    def ref_group(self, el, fmt):
        "span.refs: one pluralized prefix for a same-type group, per-item singular prefixes for mixed types; never range-collapsed"
        refs = [c for c in el if etree.QName(c).localname == 'a']
        flds = [self.ref_fld(a, fmt) for a in refs]
        mixed = len({(a.get('href') or '#')[1:].split('-')[0] for a in refs}) > 1
        out = []
        for i, (a, f) in enumerate(zip(refs, flds)):
            if i: out += self.text_runs(' and ' if i == len(flds) - 1 else ', ', fmt)
            if mixed or i == 0: out += self.ref_prefix(a, fmt, plural=not mixed and len(refs) > 1)
            out += f
        return out

    def custom_style(self, el, kind):
        "Style id for an explicit custom-style attr (stubbed + warned if undefined), else a class matching a reference style name"
        if cs := el.get('custom-style'):
            if cs.lower() in self.refstyles: return self.refstyles[cs.lower()]
            if cs not in self.stubs:
                self.stubs[cs] = (kind, re.sub(r'\W', '', cs) or f'Custom{len(self.stubs)}')
                self.warn(f'custom style {cs!r} not in reference doc; stub injected')
            return self.stubs[cs][1]
        return next((self.refstyles[c.lower()] for c in (el.get('class') or '').split()
                     if c.lower() in self.refstyles), None)

    def span(self, el, fmt):
        "Inline span: math -> inline m:oMath zone (linear source, dialect-agnostic), custom style -> rStyle, else transparent"
        if 'refs' in (el.get('class') or '').split(): return self.ref_group(el, fmt)
        if 'math' in (el.get('class') or '').split(): return [self.omath(el)]
        if sid := self.custom_style(el, 'character'): return self.runs(el, fmt | {'rstyle': sid})
        return self.runs(el, fmt)

    def omath(self, el):
        "An m:oMath zone holding `el`'s text as linear-format math runs"
        return E('m:oMath', E('m:r', E('m:t', ''.join(el.itertext()), {'xml:space': 'preserve'})))

    def fnref(self, el, fmt):
        "Footnote-reference run for a sup>a.footnote-ref, or None when `el` is an ordinary sup"
        a = el[0] if len(el) == 1 and etree.QName(el[0]).localname == 'a' else None
        if a is None or 'footnote-ref' not in (a.get('class') or ''): return None
        key = (a.get('href') or '#')[1:]
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
        li = deepcopy(li)
        for a in [a for a in li.iter() if etree.QName(a).localname == 'a'
                  and 'footnote-backref' in (a.get('class') or '')]:
            a.getparent().remove(a)
        save = self.rels, self._urlrids, self.first
        self.rels, self._urlrids = self.fn_rels, {}   # rel ids are per-part (see fn_relsxml)
        try:
            blks = [b for kind, val in self.li_parts(li)
                    for b in ([self.para(self.group_runs(val, {}), 'footnotetext')] if kind == 'inline'
                              else self.block(val, 'footnotetext'))]
        finally: self.rels, self._urlrids, self.first = save
        if not blks: blks = [self.para([], 'footnotetext')]
        mark = E('w:r', E('w:rPr', E('w:rStyle', {'w:val': _sid('footnoteref')})), E('w:footnoteRef'))
        blks[0].insert(1, E('w:r', E('w:t', ' ', {'xml:space': 'preserve'})))
        blks[0].insert(1, mark)
        return blks

    def image(self, el, fmt):
        "Embed a local image (dimensions sniffed, width/height px attrs override); remote srcs degrade to a link"
        src = el.get('src') or ''
        alt = el.get('alt') or src
        if re.match(r'[a-z][a-z0-9+.-]*://', src):
            self.warn(f'remote image not embedded: {src}')
            fake = E('a', href=src)
            fake.text = alt
            return self.link(fake, fmt)
        try: data = (self.base/src).read_bytes()
        except OSError:
            self.warn(f'image not found: {src}; alt text emitted')
            return self.text_runs(alt, fmt)
        pw, ph, dx, dy = imgsize(data) or (300, 200, 96, 96)
        cx, cy = round(pw * 914400 / dx), round(ph * 914400 / dy)
        w_, h_ = el.get('width'), el.get('height')
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
        tag = etree.QName(el).localname
        if tag == 'em': out = self.runs(el, fmt | {'i': True})
        elif tag == 'strong': out = self.runs(el, fmt | {'b': True})
        elif tag == 'code': out = self.runs(el, fmt | {'rstyle': _sid('codeinline')})
        elif tag == 'a': out = self.link(el, fmt)
        elif tag == 'del': out = self.runs(el, fmt | {'strike': True})
        elif tag == 'mark': out = self.runs(el, fmt | {'mark': True})
        elif tag == 'sub': out = self.runs(el, fmt | {'vert': 'subscript'})
        elif tag == 'sup':
            fn = self.fnref(el, fmt)
            out = fn if fn is not None else self.runs(el, fmt | {'vert': 'superscript'})
        elif tag == 'span': out = self.span(el, fmt)
        elif tag == 'img': out = self.image(el, fmt)
        elif tag == 'br': out = [E('w:r', self.rpr(fmt), E('w:br'))]
        elif tag == 'input':   # task-list checkbox
            g = '☒' if el.get('checked') else '☐'
            out = [E('w:r', self.rpr(fmt), E('w:t', g + ' ', {'xml:space': 'preserve'}))]
        elif tag == 'script': out = self.rawxml(el)
        else:  # unknown inline (abbr etc): recurse transparently
            out = self.runs(el, fmt)
        return out

    def runs(self, el, fmt):
        "Run-level elements for `el`'s text and children (tails included)"
        return self.text_runs(el.text or '', fmt) + [
            r for c in el for r in self.inline(c, fmt) + self.text_runs(c.tail or '', fmt)]

    # ---- block level --------------------------------------------------------
    def para(self, runs, style='body', extra=None, sid=None):
        "A w:p with `style` (STYLE_MAP key, or `sid` style-id override) and optional extra pPr children (schema order!)"
        ppr = E('w:pPr', E('w:pStyle', {'w:val': sid or _sid(style)}), *(extra or []))
        return E('w:p', ppr, *runs)

    def bookmark(self, el, runs):
        "Wrap `runs` in a bookmark when `el` carries an id (target for internal links)"
        if not (i := el.get('id')): return runs
        self._bkid += 1
        return [E('w:bookmarkStart', {'w:id': self._bkid, 'w:name': self.bkname(i)}),
                *runs, E('w:bookmarkEnd', {'w:id': self._bkid})]

    def codeblock(self, el):
        "Source Code paragraph, lines joined with w:br; Hl* character styles when a language class names one"
        code = el[0] if len(el) and etree.QName(el[0]).localname == 'code' else el
        lang = next((c.removeprefix('language-') for c in (code.get('class') or '').split()
                     if c.startswith('language-')), None)
        text = (code.text or '').rstrip('\n')
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
        self.nums.append((nid, 0 if etree.QName(el).localname == 'ul' else 1, int(el.get('start', 1))))
        return [b for li in el if etree.QName(li).localname == 'li' for b in self.li(li, nid, min(ilvl, 8))]

    def li_parts(self, el):
        "Split mixed li content into ('inline', [nodes]) groups and ('block', child) items, in order"
        parts = []
        def add(x):
            if isinstance(x, str) and not x.strip() and '\n' in x: return   # list-formatting whitespace
            if parts and parts[-1][0] == 'inline': parts[-1][1].append(x)
            else: parts.append(('inline', [x]))
        if el.text: add(el.text)
        for c in el:
            if etree.QName(c).localname in BLOCK_TAGS: parts.append(('block', c))
            else: add(c)
            if c.tail: add(c.tail)
        return parts

    def group_runs(self, nodes, fmt):
        "Runs for a mixed list of text strings and inline elements (tails are separate list entries)"
        return [r for n in nodes
                for r in (self.text_runs(n, fmt) if isinstance(n, str) else self.inline(n, fmt))]

    def li(self, li, nid, ilvl):
        "Blocks for one list item: the first paragraph carries the number, the rest continue indented"
        numpr = [E('w:numPr', E('w:ilvl', {'w:val': ilvl}), E('w:numId', {'w:val': nid}))]
        cont = [E('w:ind', {'w:left': 720 * (ilvl + 1)})]
        out = []
        for kind, val in self.li_parts(li):
            if kind == 'inline': out.append(self.para(self.group_runs(val, {}), 'list', numpr if not out else cont))
            elif etree.QName(val).localname in ('ul', 'ol'): out += self.list_el(val, ilvl + 1)
            elif etree.QName(val).localname == 'p':
                out.append(self.para(self.runs(val, {}), 'list', numpr if not out else cont))
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
            for cell in tr:
                if etree.QName(cell).localname not in ('td', 'th'): continue
                cs, rs = int(cell.get('colspan', 1)), int(cell.get('rowspan', 1))
                rowcells.append(('cell', ci, cell, cs, rs))
                for k in range(1, rs): spans[(ri + k, ci)] = cs
                ci = _skip(ci + cs)
            placed.append(rowcells)
            ncols = max(ncols, ci)
        return placed, ncols

    def col_widths(self, el, ncols):
        "colwidths tracks -> (dxa list, all_fr?) or (None, False) when absent"
        s = el.get('colwidths') or el.get('data-colwidths')
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
        if any(etree.QName(c).localname in BLOCK_TAGS for c in cell): return self.blocks(cell)
        jc = [E('w:jc', {'w:val': cell.get('align')})] if cell.get('align') in ('center', 'right') else None
        return [self.para(self.runs(cell, {'b': True} if header else {}), 'compact', jc)]

    def table(self, el):
        "w:tbl (+ caption paragraph before, spacer paragraph after)"
        cap = None
        rows, nhead = [], 0
        for sec in el:
            t = etree.QName(sec).localname
            if t == 'caption': cap = sec
            elif t == 'thead':
                rows += list(sec)
                nhead = len(rows)
            elif t in ('tbody', 'tfoot'): rows += list(sec)
            elif t == 'tr': rows.append(sec)
        placed, ncols = self.table_grid(rows)
        dxa, all_fr = self.col_widths(el, ncols)
        def _tcw(ci, cs):
            if not dxa: return None
            wd = sum(dxa[ci:ci + cs])
            if all_fr: return E('w:tcW', {'w:type': 'pct', 'w:w': round(wd / self.content_w * 5000)})
            return E('w:tcW', {'w:type': 'dxa', 'w:w': wd})
        tblw = (E('w:tblW', {'w:type': 'auto', 'w:w': 0}) if not dxa
                else E('w:tblW', {'w:type': 'pct', 'w:w': 5000}) if all_fr
                else E('w:tblW', {'w:type': 'dxa', 'w:w': sum(dxa)}))
        tblpr = E('w:tblPr', E('w:tblStyle', {'w:val': _sid('table')}), tblw,
                  E('w:tblLayout', {'w:type': 'fixed'}) if dxa and not all_fr else None,
                  E('w:tblLook', {'w:val': '04A0', 'w:firstRow': 1, 'w:lastRow': 0,
                                  'w:firstColumn': 0, 'w:lastColumn': 0, 'w:noHBand': 0, 'w:noVBand': 1}))
        grid = E('w:tblGrid', *[E('w:gridCol', {'w:w': dxa[i]} if dxa else None) for i in range(ncols)])
        trs = []
        for ri, rowcells in enumerate(placed):
            tcs = []
            for item in rowcells:
                if item[0] == 'cont':
                    _, ci, wd = item
                    tcs.append(E('w:tc', E('w:tcPr', _tcw(ci, wd),
                                           E('w:gridSpan', {'w:val': wd}) if wd > 1 else None,
                                           E('w:vMerge')), E('w:p')))
                else:
                    _, ci, cell, cs, rs = item
                    tcpr = E('w:tcPr', _tcw(ci, cs),
                             E('w:gridSpan', {'w:val': cs}) if cs > 1 else None,
                             E('w:vMerge', {'w:val': 'restart'}) if rs > 1 else None)
                    body = self.cell_blocks(cell, ri < nhead)
                    if not len(body) or etree.QName(body[-1]).localname != 'p': body.append(E('w:p'))
                    tcs.append(E('w:tc', tcpr, *body))
            trs.append(E('w:tr', E('w:trPr', E('w:tblHeader')) if ri < nhead else None, *tcs))
        out = self.caption_para(el, 'tbl', cap)
        return out + [E('w:tbl', tblpr, grid, *trs), E('w:p')]

    def caption_para(self, el, typ, capel, fmt={}):
        """Numbered caption paragraph: 'Label N: text' with a SEQ field as N. When `el` has an id, the
        label+number span is bookmarked under it (REF target) and the number alone under `<name>_n`.
        Emitted whenever there is a caption or an id; the label word comes from reftypes[typ]."""
        if capel is None and not el.get('id'): return []
        label = self.reftypes[typ][0]
        seq = [E('w:fldSimple', {'w:instr': rf' SEQ {label} \* ARABIC '}, E('w:r', self.rpr(fmt), E('w:t', '#')))]
        self.has_fields = True
        if i := el.get('id'):
            nm = self.bkname(i)
            self._bknames[i + '\0n'] = nm + '_n'   # reserve the number-only name too
            self._bkid += 2
            seq = [E('w:bookmarkStart', {'w:id': self._bkid, 'w:name': nm + '_n'}), *seq,
                   E('w:bookmarkEnd', {'w:id': self._bkid})]
            runs = [E('w:bookmarkStart', {'w:id': self._bkid - 1, 'w:name': nm}),
                    *self.text_runs(label + ' ', fmt), *seq,
                    E('w:bookmarkEnd', {'w:id': self._bkid - 1})]
        else:
            runs = self.text_runs(label + ' ', fmt) + seq
        cap = [] if capel is None else self.runs(capel, fmt)
        if cap: runs += self.text_runs(': ', fmt) + cap
        return [self.para(runs, 'caption')]

    def figure(self, el):
        "Figure: image paragraph, then its numbered caption paragraph below (Word convention)"
        img = next((c for c in el.iter() if etree.QName(c).localname == 'img'), None)
        capel = next((c for c in el if etree.QName(c).localname == 'figcaption'), None)
        out = [] if img is None else [self.para(self.image(img, {}), 'body')]
        return out + self.caption_para(el, 'fig', capel)

    # Paragraphs directly after these blocks (or at document start) take First Paragraph rather
    # than Body Text, matching pandoc's docx writer exactly, so the two agree on which is "first".
    FIRST_AFTER = {'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'ul', 'ol', 'table', 'dl', 'hr'}

    def block(self, el, style='body', sid=None):
        "Block elements for `el` (one element may yield several); `sid` is a custom-style id override for paragraphs"
        out = self._block(el, style, sid)
        tag = etree.QName(el).localname
        if tag in self.FIRST_AFTER or (tag == 'div' and 'display' in (el.get('class') or '')):
            self.first = True
        return out

    def _block(self, el, style, sid):
        tag = etree.QName(el).localname
        if tag == 'p':
            ex = self.qindent() if style == 'blockquote' else None
            psid = self.custom_style(el, 'paragraph') or sid
            use = 'firstpara' if self.first and style == 'body' and not psid else style
            self.first = False
            return [self.para(self.bookmark(el, self.runs(el, {})), use, ex, psid)]
        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            return [self.para(self.bookmark(el, self.runs(el, {})), tag)]
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
        if tag == 'div':
            cls = (el.get('class') or '').split()
            if 'math' in cls and 'display' in cls:
                return [E('w:p', E('m:oMathPara', self.omath(el)))]
            return self.blocks(el, style, self.custom_style(el, 'paragraph') or sid)
        if tag in BLOCK_TAGS and any(etree.QName(c).localname in BLOCK_TAGS for c in el):
            return self.blocks(el, style, sid)   # unknown container: recurse
        self.warn(f'unhandled block <{tag}>; emitted as plain paragraph')
        return [self.para(self.runs(el, {}), style, None, sid)]

    RAWNS = ' '.join(f'xmlns:{k}="{v}"' for k, v in NS.items() if k != 'xml')

    def rawxml(self, el):
        "Elements parsed from a raw docx payload (`{=docx}` in Markdown); other formats skip silently"
        if el.get('type') != 'text/x-docx': return []
        try: return list(etree.fromstring(f'<x2d {self.RAWNS}>{el.text or ""}</x2d>'))
        except etree.XMLSyntaxError as e:
            self.warn(f'malformed text/x-docx payload: {e}')
            return []

    def dl(self, el):
        "Definition list: dt/dd paragraphs in their dialect styles"
        out = []
        for c in el:
            t = etree.QName(c).localname
            if t == 'dt': out.append(self.para(self.runs(c, {}), 'dt'))
            elif t == 'dd':
                blocky = any(etree.QName(k).localname in BLOCK_TAGS for k in c)
                out += self.blocks(c, 'dd') if blocky else [self.para(self.runs(c, {}), 'dd')]
        return out

    def blocks(self, parent, style='body', sid=None):
        return [b for el in parent for b in self.block(el, style, sid)]

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
                    an.append(E('w:lvl', {'w:ilvl': i},
                                E('w:start', {'w:val': 1}), E('w:numFmt', {'w:val': fmt}),
                                E('w:lvlText', {'w:val': txt}), E('w:lvlJc', {'w:val': 'left'}),
                                E('w:pPr', E('w:ind', {'w:left': 720 * (i + 1), 'w:hanging': 360}))))
                root.append(an)
        if self.headnum:
            an = E('w:abstractNum', {'w:abstractNumId': self._absbase + 2},
                   E('w:multiLevelType', {'w:val': 'multilevel'}))
            for i in range(9):
                txt, fmt = self.scheme[i] if i < len(self.scheme) else (f'%{i+1}.', 'decimal')
                an.append(E('w:lvl', {'w:ilvl': i},
                            E('w:start', {'w:val': 1}), E('w:numFmt', {'w:val': fmt}),
                            E('w:pStyle', {'w:val': f'Heading{i + 1}'}) if i < 6 else None,
                            E('w:lvlText', {'w:val': txt}), E('w:lvlJc', {'w:val': 'left'})))
            root.append(an)
        for e in self.xabs: root.append(e)
        for e in tnums: root.append(e)
        if self.headnum:
            root.append(E('w:num', {'w:numId': self.headnum}, E('w:abstractNumId', {'w:val': self._absbase + 2})))
        for e in self.xnums: root.append(e)
        for nid, aid, start in self.nums:
            root.append(E('w:num', {'w:numId': nid}, E('w:abstractNumId', {'w:val': self._absbase + aid}),
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
            root.append(E('w:style', {'w:type': kind, 'w:customStyle': 1, 'w:styleId': sid},
                          E('w:name', {'w:val': name}),
                          E('w:basedOn', {'w:val': 'BodyText' if kind == 'paragraph' else 'DefaultParagraphFont'}),
                          E('w:qFormat')))
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def harvest_footnotes(self, els):
        "Split out footnote endnote sections, indexing their li definitions by id; returns body elements"
        body, fn = [], []
        for el in els:
            (fn if etree.QName(el).localname == 'section' and 'footnotes' in (el.get('class') or '') else body).append(el)
        self.fndefs.update({li.get('id'): li for sec in fn for li in sec.iter()
                            if etree.QName(li).localname == 'li' and li.get('id')})
        return body

    BOOKMARKABLE = {'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}

    def to_docx(self, xhtml, dest):
        els = self.harvest_footnotes(parse_frag(xhtml))
        self.idtext, self.reftarget = {}, {}
        for e in (e for el in els for e in el.iter()):
            if not (i := e.get('id')): continue
            self.idtext[i] = ' '.join(''.join(e.itertext()).split())
            t = etree.QName(e).localname
            if t in ('figure', 'table'): self.reftarget[i] = 'caption'
            elif t in self.BOOKMARKABLE: self.reftarget[i] = 'block'
        self.ids = set(self.idtext)
        blocks = [b for el in els for b in self.block(el)]
        docxml = self.document(blocks)
        parts = {}   # archive name -> (bytes, content-type kind); each also gets a document rel
        if self.nums or self.headnum or self.xnums:
            parts['word/numbering.xml'] = (self.numbering_xml(), 'numbering')
            if self.tmplnum is None: self.rels.append((self.rid(), f'{R}/numbering', 'numbering.xml', False))
        if self.fnotes:
            parts['word/footnotes.xml'] = (self.footnotes_xml(), 'footnotes')
            self.rels.append((self.rid(), f'{R}/footnotes', 'footnotes.xml', False))
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
            for name, data in self.media.items(): zo.writestr(name, data)
        return self.warnings

    SETT_AFTER_UPDATE = {'footnotePr', 'endnotePr', 'compat', 'rsids', 'mathPr', 'attachedSchema',
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

def convert(xhtml, dest, reference=None, base=None, reftypes=None, number_headings=None):
    """Convert an xhtmlmd XHTML fragment to a docx file at `dest`; returns a list of warnings.
    `reference` is a reference docx path, or a list of them: the first supplies the whole archive
    (default, or when None: the built-in template), later entries contribute styles only, later-wins -
    each a .docx path, a raw styles/numbering .xml path, or a fastpylight theme name (whose Hl*/Source
    Code styles are generated; see styles.theme_ref). Default adds 'github_light' when fastpylight is
    installed, so code blocks are colored; pass a bare reference for plain code. Relative image srcs
    resolve against `base` ('.'). Cross-references (`data-ref` anchors from Markdown `[@sec-x]`) become
    live REF fields; `reftypes` maps type tokens to (singular, plural) prefix words beyond the built-in
    `sec`, and `number_headings` (a styles.SCHEMES name such as 'legal', or a {lvlText: numFmt} dict, one entry per heading level)
    numbers the headings via a multilevel list so `\\w` fields resolve."""
    return Converter(reference, base, reftypes, number_headings).to_docx(xhtml, dest)
