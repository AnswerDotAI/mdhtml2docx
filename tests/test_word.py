import pytest
from pathlib import Path

from mdhtml import sample_md, to_mdhtml
from mdhtml2docx import convert

DATA = Path(__file__).parent.parent/'_data'

pytestmark = pytest.mark.slow


def test_sample():
    docx = DATA/'sample.docx'
    warns = convert(to_mdhtml(sample_md(), smart=True, auto_ids=True, implicit_figures=True), docx, number_headings='legal')
    assert len(warns) == 2 and all('remote image not embedded' in w for w in warns)
    assert docx.stat().st_size
