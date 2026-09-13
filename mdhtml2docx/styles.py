"Style names and the reference template. STYLE_MAP is the single source of truth for converter-emitted styles: the template defines them all, the converter emits exactly these names, and tests close the loop. The template also carries author-selectable extras (`Borderless Table`), reached only via `custom-style`/class."
from importlib.resources import files

__all__ = ['STYLE_MAP', 'style_id', 'ref_path', 'theme_styles', 'theme_ref']

STYLE_MAP = dict(  # chkstyle: ignore-node
    body='Body Text', firstpara='First Paragraph', blockquote='Quote', codeblock='Source Code', codeinline='Verbatim Char',
    h1='Title', h2='heading 1', h3='heading 2', h4='heading 3', h5='heading 4', h6='heading 5',
    compact='Compact', hyperlink='Hyperlink', list='List Paragraph', dt='Definition Term', dd='Definition',
    caption='caption', footnotetext='footnote text', footnoteref='footnote reference', table='Table Grid')

def style_id(name):
    "Word style id for style `name`: title-cased, spaces dropped ('Body Text' -> 'BodyText', 'heading 1' -> 'Heading1')"
    return name.title().replace(' ', '')

def ref_path(): return files('mdhtml2docx')/'templates'/'reference.docx'



# Word underline values for fastpylight/Lumis underline styles
_ULINE = dict(solid='single', wavy='wave', double='double', dotted='dotted', dashed='dash')

def _hx(c): return c.lstrip('#').upper()

def theme_styles(theme):
    """w:style elements for fastpylight `theme`: one Hl* character style per scope (id 'HlKeywordFunction'
    for scope 'keyword.function', reversibly), plus a Source Code paragraph style carrying the theme's
    code-block background and default color (mirroring the template's, so later-wins merge replaces it)"""
    from fastpylight import theme_colors
    from .wml import e
    tc = theme_colors(theme)
    nrm = tc.pop('normal', {})
    def rpr(st):
        return e.rPr(e.b() if st['bold'] else None, e.i() if st['italic'] else None,  # chkstyle: ignore-node
            e.strike() if st['strikethrough'] else None,
            e.color(val=_hx(st['fg'])) if st['fg'] else None,
            e.u(val=_ULINE[st['underline']]) if st['underline'] else None,
            e.shd(val='clear', color='auto', fill=_hx(st['bg'])) if st['bg'] else None)
    def sty(scope, st):
        name = ('hl ' + scope.replace('.', ' ')).title()
        return e.style(e.name(val=name), e.basedOn(val='DefaultParagraphFont'),
            rpr(st), type='character', customStyle=1, styleId=style_id(name))
    sc = e.style(
        e.name(val='Source Code'), e.basedOn(val='Normal'),
        e.next(val='FirstParagraph'), e.uiPriority(val=1), e.qFormat(),
        e.pPr(e.keepLines(),
            e.shd(val='clear', color='auto', fill=_hx(nrm.get('bg') or '#F5F5F5')),
            e.spacing(before=120, after=120)),
        e.rPr(e.rFonts(ascii='Consolas', hAnsi='Consolas', cs='Consolas'),
            e.color(val=_hx(nrm['fg'])) if nrm.get('fg') else None,
            e.sz(val=20), e.szCs(val=20)), type='paragraph', styleId='SourceCode')
    return [sc] + [sty(s, st) for s, st in sorted(tc.items())]

def theme_ref(theme, dest):
    "Write a minimal, Word-openable reference docx at `dest` carrying only `theme`'s styles, for use as a later `reference` entry"
    from oxml import Document
    from .wml import R, e
    document = Document.new()
    document.main.replace(e.document(e.body(e.p())).bytes())
    document.package.add_part('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
        e.styles(*theme_styles(theme)).bytes())
    document.package.add_relationship(document.package.main_part, f'{R}/styles', 'styles.xml')
    document.save(dest)
    return dest
