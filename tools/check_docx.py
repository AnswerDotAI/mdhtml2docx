"Development-only Word acceptance check; requires macscript and Microsoft Word on macOS."
from pathlib import Path
from oxml import Document
from macscript.word import TIMEOUT, _val, word, open_doc, doc_text, close_doc, dismiss_dlgs


def check_docx(path, timeout=10):
    """Word acceptance check for a docx in a folder Word can access, such as _data.
    Return (status, details): invalid for oxml errors, clean for an ordinary open, recovered for a
    repaired copy, or corrupt for a blocked open. Details are errors/dialog messages or body text.
    Closes what it opened."""
    try: issues = Document.open(path).validate()['issues']
    except (ValueError, OSError) as error: return 'invalid', str(error)
    if issues: return 'invalid', issues
    p = Path(path).expanduser().resolve()
    before = {_val(n) for n in (word().documents.name.get(timeout=TIMEOUT) or []) if _val(n)}
    try: d = open_doc(p, timeout=timeout)
    except Exception: return 'corrupt', dismiss_dlgs()
    names = {_val(n) for n in (word().documents.name.get(timeout=TIMEOUT) or []) if _val(n)}
    if p.name in names:
        txt = doc_text(d)
        close_doc(d)
        return 'clean', txt
    new = names - before
    if new:   # recovery renamed it (observed: fresh unsaved Document1)
        d2 = word().documents[next(iter(new))]
        txt = doc_text(d2)
        close_doc(d2)
        return 'recovered', txt
    return 'corrupt', dismiss_dlgs()
