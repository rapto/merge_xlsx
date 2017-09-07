from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import sys
if sys.version_info >= (2, 7):
    from path import Path
else:
    from path import path as Path
from merge_xlsx.util import TemporaryDirectory
from contextlib import contextmanager
import datetime
from merge_xlsx.openpyxl_datetime import to_excel

ns = dict(r='http://schemas.openxmlformats.org/package/2006/relationships',
          oo='http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
          a='http://schemas.openxmlformats.org/drawingml/2006/main',
          )



def update_cell_value(dp, k, v):
    ET = etree.parse(dp/'xl/_rels/workbook.xml.rels')
    ws_path = dp / 'xl' / (sorted(ET.xpath('''//r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet']/@Target''', namespaces=ns))[0])
    ET = etree.parse(ws_path)
    if isinstance(v, basestring):
        c = ET.xpath('''//oo:c[@r='%s']''' % k, namespaces=ns)[0]
        cv = ET.xpath('''//oo:c[@r='%s']/oo:v''' % k, namespaces=ns)[0]
        c.attrib['t'] = 'inlineStr'
        c.remove(cv)
        del c.attrib['s'] # FIXME
        c.append(etree.fromstring('<is xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ><t>%s</t></is>' % v))
    else:
        if isinstance(v, (datetime.date, datetime.datetime)):
            v = to_excel(v)
        if isinstance(v, (int, float, long)):
            v = str(v)
        cv = ET.xpath('''//oo:c[@r='%s']/oo:v''' % k, namespaces=ns)[0]
        cv.text = v
    ws_path.write_bytes(etree.tostring(ET))


def replace_image(dp, k, v):
    dr_path = dp / 'xl/drawings/drawing1.xml'
    drET = etree.parse(dr_path)
    drET.findall("//xdr:twoCellAnchor/xdr:pic/xdr:nvPicPr/xdr:cNvPr[@name='%s']" % k, namespaces=ns)[0]
    blip = drET.xpath("//xdr:twoCellAnchor/xdr:pic/xdr:nvPicPr/xdr:cNvPr[@name='%s']/../../xdr:blipFill/a:blip" % k, namespaces=ns)[0]
    rId = blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
    
    rels_path =  dp / 'xl/drawings/_rels/drawing1.xml.rels'
    relsET = etree.parse(rels_path)
    path = Path(dp / 'xl/drawings' / relsET.xpath('r:Relationship[@Id="%s"]' % rId, namespaces=ns)[0].attrib['Target'])
    assert path.exists()
    path.write_bytes(v.read())
                
    
def merge(template, result, data=None, **kwargs):
    if data is not None:
        kwargs.update(data)
    with excel_recompressor(template, result) as dp:
        for k, v in kwargs.items():
            if ' ' not in k:
                update_cell_value(dp, k, v)
            elif 'Image' in k:
                replace_image(dp, k, v)                
        
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
