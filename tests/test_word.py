import pytest
from pathlib import Path

from mdhtml import to_mdhtml
from mdhtml.tools import SAMPLE_MD, sample_md
from mdhtml2docx import convert

DATA = Path(__file__).parent.parent/'_data'

pytestmark = [pytest.mark.slow, pytest.mark.checkout]


def test_sample():
    docx = DATA/'sample.docx'
    warns = convert(to_mdhtml(sample_md(), smart=True, auto_ids=True, implicit_figures=True), docx, base=SAMPLE_MD.parent, number_headings='legal')
    assert not warns
    assert docx.stat().st_size
