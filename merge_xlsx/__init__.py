from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from path import Path
from merge_xlsx.util import TemporaryDirectory


def merge(template, result, **kwargs):
    with TemporaryDirectory() as d:
        ns = dict(r='http://schemas.openxmlformats.org/package/2006/relationships')
        dp = Path(d)
        z = ZipFile(template)
        z.extractall(d)
        ET = etree.parse(dp/'xl/_rels/workbook.xml.rels')
        ws_path = dp / 'xl' / ET.xpath('''//r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet']/@Target''', namespaces=ns)[0]
        ET = etree.parse(ws_path)
        for k, v in kwargs.items():
            cv = ET.xpath('''//oo:c[@r='%s']/oo:v''' % k, namespaces=dict(oo='http://schemas.openxmlformats.org/spreadsheetml/2006/main'))[0]
            cv.text = str(v)
        ws_path.write_bytes(etree.tostring(ET))
        nz = ZipFile(result, 'w', ZIP_DEFLATED)
        for f in z.infolist():
            nz.write(dp / f.filename, f.filename)
    return nz
    