import pytest
from pathlib import Path

pytest.importorskip('appscript')
from xhtml2docx.word import *

DATA = Path(__file__).parent.parent/'_data'

pytestmark = pytest.mark.slow

@pytest.fixture(scope='module')
def wd():
    try: word().version.get(timeout=5)
    except Exception as e: pytest.skip(f'Microsoft Word not reachable: {e}')

def test_word_loop(wd):
    d = new_doc()
    set_text(d, 'Line one.\rLine two.')
    d = save_docx(d, DATA/'test_word.docx')
    pdf = save_pdf(d, DATA/'test_word.pdf')
    close_doc(d)
    d = open_doc(DATA/'test_word.docx')
    txt = doc_text(d)
    close_doc(d)
    assert txt == 'Line one.\nLine two.\n'
    assert pdf.stat().st_size
