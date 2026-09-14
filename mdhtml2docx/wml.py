"""WordprocessingML construction helpers: namespaces, oxml expressions, unit
conversions, and the colwidths track-list parser. Builders emit children in the order callers pass
them; oxml sorts them into schema order when they are attached."""
import re
from oxml import E, e, namespace_uris

__all__ = ['W', 'R', 'WP', 'A', 'PIC', 'NS', 'E', 'e', 'twips', 'parse_tracks']

NS = {p: namespace_uris[p] for p in ('w', 'r', 'wp', 'a', 'pic', 'm', 'xml')}
W, R, WP, A, PIC = (NS[p] for p in ('w', 'r', 'wp', 'a', 'pic'))

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
