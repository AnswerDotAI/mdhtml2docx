"""WordprocessingML construction helpers: namespaces, oxml expressions, unit
conversions, and the colwidths track-list parser. Child-element order inside w:pPr/w:rPr/w:tcPr
etc follows the ECMA-376 content models; builders here emit children in the order callers pass
them, so callers are responsible for schema order (the mdhtml2docx module's helpers encode it)."""
import re, struct
from xml.sax.saxutils import escape
from oxml import E, e, Tree

__all__ = ['W', 'R', 'WP', 'A', 'PIC', 'NS', 'E', 'e', 'wchild', 'wpos', 'twips', 'parse_tracks', 'EMU_PER_PX', 'imgsize', 'drawing']

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
XML = 'http://www.w3.org/XML/1998/namespace'
NS = dict(w=W, r=R, wp=WP, a=A, pic=PIC, m=M, xml=XML)

def wchild(el, name):
    "First direct WordprocessingML child named `name`, or None"
    return next((c for c in el.children if c.raw['qname'] == (W, name)), None) if el is not None else None

def wpos(el, names):
    "Content index before the first WordprocessingML child in `names`, or at the end"
    ids = el.raw['children']
    return next((ids.index(c.node_id) for c in el.children if c.raw['qname'][0] == W and c.raw['qname'][1] in names), len(ids))

# CSS length units in twips (1/20 pt). em/rem/ch use the template body size (11pt Aptos/Calibri).
UNITS = dict(pt=20, px=15, pc=240, em=220, rem=220, ch=110)
UNITS['in'] = 1440
UNITS['cm'] = 566.9
UNITS['mm'] = 56.69

def twips(s):
    "CSS length string -> twips (int), e.g. '10em' -> 2200; raises ValueError on unknown units"
    m = re.fullmatch(r'([\d.]+)\s*([a-z]+)', s.strip())
    if not m or m[2] not in UNITS: raise ValueError(f'unsupported length: {s!r}')
    return round(float(m[1]) * UNITS[m[2]])

def parse_tracks(s):
    """Parse a colwidths track list into (kind, value) pairs: ('dxa', twips) for lengths,
    ('fr', share) for fr units. E.g. '10em 3fr 7fr' -> [('dxa',2200),('fr',3.0),('fr',7.0)]"""
    def _track(tok):
        m = re.fullmatch(r'([\d.]+)fr', tok)
        return ('fr', float(m[1])) if m else ('dxa', twips(tok))
    return [_track(t) for t in s.split()]

EMU_PER_PX = 914400 // 96  # pixel at 96dpi in EMUs


def imgsize(data):
    """Sniff (px_w, px_h, dpi_x, dpi_y) from png/jpeg/gif bytes, or None if unrecognized.
    (Header layouts cribbed from python-docx's image parsers, MIT.)"""
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        pw, ph = struct.unpack('>II', data[16:24])
        dx = dy = 96
        i = 8
        while i + 8 <= len(data):
            ln, typ = struct.unpack('>I4s', data[i:i+8])
            if typ == b'pHYs':
                px, py, unit = struct.unpack('>IIB', data[i+8:i+17])
                if unit == 1: dx, dy = round(px * 0.0254), round(py * 0.0254)
                break
            i += ln + 12
        return pw, ph, dx, dy
    if data[:2] == b'\xff\xd8':
        dx = dy = 96
        i = 2
        while i + 4 <= len(data) and data[i] == 0xFF:
            m, ln = data[i+1], struct.unpack('>H', data[i+2:i+4])[0]
            if m == 0xE0 and data[i+4:i+9] == b'JFIF\x00':
                unit, xd, yd = data[i+11], *struct.unpack('>HH', data[i+12:i+16])
                if unit == 1: dx, dy = xd, yd
                elif unit == 2: dx, dy = round(xd * 2.54), round(yd * 2.54)
            elif m in (0xC0, 0xC1, 0xC2, 0xC3):
                ph, pw = struct.unpack('>HH', data[i+5:i+9])
                return pw, ph, dx, dy
            i += 2 + ln
    if data[:6] in (b'GIF87a', b'GIF89a'):
        pw, ph = struct.unpack('<HH', data[6:10])
        return pw, ph, 96, 96
    return None

DRAWING = r'''<w:drawing xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{P}" xmlns:r="{R}">
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="{n}" name="Picture {n}" descr="{descr}"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic><a:graphicData uri="{P}">
<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="Picture {n}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'''

def drawing(rid, n, cx, cy, descr=''):
    "A w:drawing (inline picture) element: relationship `rid`, unique docPr id `n`, extent in EMU"
    return Tree(DRAWING.format(W=W, WP=WP, A=A, P=PIC, R=R, rid=rid, n=n, cx=cx, cy=cy,
        descr=escape(descr, {'"': '&quot;'})).encode()).root
