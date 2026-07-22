import re, zipfile
from pathlib import Path

from mdhtml2docx.validate import fast_checks

GOOD = Path(__file__).parent/'fixtures'/'good.docx'

def test_fast_checks(tmp_path):
    assert fast_checks(GOOD) == 'valid'
    z = zipfile.ZipFile(GOOD)
    xml = z.read('word/document.xml').decode()

    def rezip(name, doc):
        p = tmp_path/name
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as zo:
            for i in z.infolist(): zo.writestr(i.filename, doc if i.filename=='word/document.xml' else z.read(i.filename))
        return p

    r = fast_checks(rezip('badxml.docx', xml.replace('</w:body>', '</w:borked>')))
    assert r.startswith('xml:')
    trunc = tmp_path/'trunc.docx'
    trunc.write_bytes(GOOD.read_bytes()[:2000])
    assert fast_checks(trunc).startswith('container:')
    r = fast_checks(rezip('bogus.docx', xml.replace('<w:body>', '<w:body><w:bogusElement/>')))
    assert r.startswith('schema:') and 'bogusElement' in r
    r = fast_checks(rezip('banana.docx', re.sub(r'(<w:p [^>]*>)', r'\1<w:pPr><w:ind w:left="banana"/></w:pPr>', xml, count=1)))
    assert r.startswith('schema:') and 'banana' in r
    raw = bytearray(GOOD.read_bytes())
    ft = next(i for i in z.infolist() if i.filename=='word/fontTable.xml')
    off = raw.index(b'word/fontTable.xml', ft.header_offset) + len('word/fontTable.xml') + ft.compress_size//2
    raw[off:off+8] = b'\xde\xad\xbe\xef'*2
    crc = tmp_path/'crc.docx'
    crc.write_bytes(bytes(raw))
    assert fast_checks(crc).startswith('crc:')
