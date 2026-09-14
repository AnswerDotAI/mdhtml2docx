import sys, pytest


@pytest.mark.skipif(sys.platform != 'darwin', reason='Word scripting requires macOS')
def test_invalid_docx(tmp_path):
    pytest.importorskip('macscript.word')
    from tools.check_docx import check_docx
    path = tmp_path/'invalid.docx'
    path.write_bytes(b'not a docx')
    status, details = check_docx(path)
    assert status == 'invalid' and details
