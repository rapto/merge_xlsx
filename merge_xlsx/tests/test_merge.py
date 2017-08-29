import unittest
import pkg_resources
import openpyxl

from merge_xlsx import merge
from path import Path
from merge_xlsx.util import TemporaryDirectory

class TestXlsxMerge(unittest.TestCase):
    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_merge(self):
        f = pkg_resources.resource_filename('merge_xlsx.tests', 'fixtures/template.xlsx')
        merged_data = dict(C12=666)
        with TemporaryDirectory() as td:
            result = Path(td) / 'result.xlsx'
            merge(f, result, **merged_data)
            xlsx = openpyxl.load_workbook(result)
            ws = xlsx.get_sheet_by_name('Hoja1')
            for k, v in merged_data.items(): 
                c = ws[k] 
                self.assertEqual(c.value, v)