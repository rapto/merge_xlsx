from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from path import Path
from merge_xlsx.util import TemporaryDirectory
from contextlib import contextmanager
import datetime
from merge_xlsx.openpyxl_datetime import to_excel

ns = dict(r='http://schemas.openxmlformats.org/package/2006/relationships',
          oo='http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          )


def merge(template, result, **kwargs):
    with excel_recompressor(template, result) as dp:
        ET = etree.parse(dp/'xl/_rels/workbook.xml.rels')
        ws_path = dp / 'xl' / ET.xpath('''//r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet']/@Target''', namespaces=ns)[0]
        ET = etree.parse(ws_path)
        for k, v in kwargs.items():
            if isinstance(v, basestring):
                c = ET.xpath('''//oo:c[@r='%s']''' % k, namespaces=ns)[0]
                cv = ET.xpath('''//oo:c[@r='%s']/oo:v''' % k, namespaces=ns)[0]
                c.attrib['t'] = 'inlineStr'
                c.remove(cv)
                del c.attrib['s']
                # FIXME
                c.append(etree.fromstring('<is xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ><t>%s</t></is>' % v))
            else:
                if isinstance(v, (datetime.date, datetime.datetime)):
                    v = to_excel(v)
                if isinstance(v, (int, float, long)):
                    v = str(v)
                cv = ET.xpath('''//oo:c[@r='%s']/oo:v''' % k, namespaces=ns)[0]
                cv.text = v
        ws_path.write_bytes(etree.tostring(ET))
        
@contextmanager
def excel_recompressor(template, result):
    with TemporaryDirectory() as d:
        dp = Path(d)
        z = ZipFile(template)
        z.extractall(d)
        yield dp
        nz = ZipFile(result, 'w', ZIP_DEFLATED)
        for f in z.infolist():
            nz.write(dp / f.filename, f.filename)
        nz.close()
