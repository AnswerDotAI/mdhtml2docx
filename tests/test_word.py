import pytest
from pathlib import Path

from mdhtml import DASHES, replacements, md2mdhtml
from mdhtml.tools import SAMPLE_MD, sample_md
from mdhtml2docx import mdhtml2docx

DATA = Path(__file__).parent.parent/'_data'

pytestmark = [pytest.mark.slow, pytest.mark.checkout]


def test_sample():
    docx = DATA/'sample.docx'
    warns = mdhtml2docx(md2mdhtml(sample_md(), callbacks={'text': replacements(*DASHES)}, implicit_figures=True), docx, base=SAMPLE_MD.parent, number_headings='legal')
    assert not warns
    assert docx.stat().st_size
