import hashlib
import io
import tempfile
import unittest
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from xlsx_calc import NS, formula_targets, package, results, uncached_workbook
from xlsx_benchmark import compare_values, prepare_input
from xlsx_calc_reference import validate_source


def fixture():
    output = io.BytesIO()
    with zipfile.ZipFile(output,'w') as archive:
        archive.writestr('xl/workbook.xml','''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcMode="manual"/></workbook>''')
        archive.writestr('xl/_rels/workbook.xml.rels','''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>''')
        archive.writestr('xl/worksheets/sheet1.xml','''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"><sheetData><row r="1" x14ac:dyDescent="0.25"><c r="A1"><v>7</v></c><c r="B1"><f>A1*2</f><v>999</v></c><c r="C1"><f>UNSUPPORTED(A1)</f><v>14</v></c><c r="D1" t="str"><f>"new"</f><v>old</v></c><c r="F1"><f t="array" ref="F1:G1">TRANSPOSE(A1:A2)</f><v>111</v></c><c r="G1" t="str"><v>222</v></c></row></sheetData></worksheet>''')
    return output.getvalue()


class CalculationTests(unittest.TestCase):
    def test_clears_formula_and_array_caches_without_changing_literal_inputs(self):
        source = fixture()
        targets,count = formula_targets(package(source))
        self.assertEqual(count,4)
        self.assertEqual([t['cell'] for t in targets],['B1','C1','D1','F1','G1'])
        cleared = package(uncached_workbook(source,targets))
        self.assertEqual(formula_targets(cleared),(targets,count))
        self.assertTrue(all(cell['value']=={'kind':'empty'} for cell in results(cleared,targets)))
        self.assertEqual(results(cleared,[{'sheet':0,'cell':'A1'}])[0]['value'],{'kind':'number','value':7.0})
        root = ET.fromstring(cleared['xl/worksheets/sheet1.xml'])
        self.assertEqual(root.get('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable'),'x14ac')
        settings = ET.fromstring(cleared['xl/workbook.xml']).find('s:calcPr',NS)
        self.assertEqual(settings.get('calcMode'),'auto')
        self.assertEqual(settings.get('fullCalcOnLoad'),'1')
        self.assertEqual(uncached_workbook(source,targets),uncached_workbook(source,targets))

    def test_frozen_input_ignores_zip_compression_but_rejects_changed_contents(self):
        source = fixture()
        targets,count = formula_targets(package(source))
        cleared = uncached_workbook(source,targets)
        def repack(parts):
            output = io.BytesIO()
            with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as archive:
                for name,data in parts.items():
                    archive.writestr(name,data)
            return output.getvalue()
        frozen = repack(package(cleared))
        self.assertNotEqual(frozen,cleared)
        reference = dict(schema_version=1,method='excel-full-rebuild-v1',engine='Microsoft Excel',
            source_sha256=hashlib.sha256(source).hexdigest(),input_sha256=hashlib.sha256(frozen).hexdigest(),
            clear_cells=targets,formula_cells=count,cells=results(package(source),targets))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'source.xlsx'
            path.write_bytes(source)
            self.assertEqual(prepare_input(path,reference,frozen),frozen)
            poisoned = package(frozen)
            poisoned['xl/worksheets/sheet1.xml'] = poisoned['xl/worksheets/sheet1.xml'].replace(b'<v>7</v>',b'<v>8</v>')
            with self.assertRaisesRegex(ValueError,'differs'):
                prepare_input(path,reference,repack(poisoned))

    def test_accuracy_checks_types_errors_and_numeric_tolerance(self):
        def cell(kind,value):
            return {'sheet':0,'cell':'A1','value':{'kind':kind,'value':value}}
        for expected,actual,correct in [
            (cell('number',1.0),cell('number',1.0+1e-10),1),
            (cell('number',1.0),cell('number',1.0001),0),
            (cell('number',1),cell('bool',True),0),
            (cell('text','ABC'),cell('text','abc'),0),
            (cell('error','#DIV/0!'),cell('error','#DIV/0!'),1),
            (cell('error','#DIV/0!'),cell('error','#NAME?'),0),
        ]:
            score = compare_values([expected],[actual],1e-9,1e-12)
            self.assertEqual(score['correct'],correct)
            self.assertEqual(score['total'],1)

    def test_reference_rejects_volatile_functions_and_external_links(self):
        parts = package(fixture())
        validate_source(parts)
        parts['xl/worksheets/sheet1.xml'] = parts['xl/worksheets/sheet1.xml'].replace(b'A1*2',b'NOW()')
        with self.assertRaisesRegex(ValueError,'Volatile'):
            validate_source(parts)
        parts = package(fixture())
        parts['xl/externalLinks/externalLink1.xml'] = b'<externalLink/>'
        with self.assertRaisesRegex(ValueError,'self-contained'):
            validate_source(parts)

    def test_rejects_reordered_results(self):
        with self.assertRaisesRegex(ValueError,'identity'):
            compare_values([{'sheet':0,'cell':'A1'}],[{'sheet':1,'cell':'A1'}],1e-9,1e-12)


if __name__ == '__main__':
    unittest.main()
