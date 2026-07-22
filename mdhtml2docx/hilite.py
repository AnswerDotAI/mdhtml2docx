"""Optional syntax scopes for code blocks, via fastpylight (Jeremy's tree-sitter highlighter).
When fastpylight is absent or the language unknown, callers fall back to plain runs, so the
package keeps its lxml-only hard dependency. Colors live in the reference doc's Hl* character
styles (see styles.theme_styles); this module only tokenizes."""

# theme_colors is probed too: a pre-theme_colors fastpylight is treated as absent, not half-working
try: from fastpylight import tokenize, theme_colors
except ImportError: tokenize = None

__all__ = ['segments']

def segments(code, lang):
    """Split `code` into [(text, scope)] using fastpylight's byte-offset tokens, or None when
    fastpylight or `lang` is unavailable (caller then emits plain runs)"""
    if tokenize is None or not lang: return None
    try: toks = tokenize(code, lang)
    except ValueError: return None
    b = code.encode()
    out, pos = [], 0
    def emit(s, e, scope=None):
        if e > s: out.append((b[s:e].decode(), scope))
    for s, e, scope in toks:
        emit(pos, s)
        emit(s, e, scope)
        pos = e
    emit(pos, len(b))
    return out
