import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from xml.dom import minidom

import fitz

from test_xlsx_calc import fixture
from xlsx_calc import NS, package, results
from xlsx_fidelity import children, libreoffice_xlsx_fidelity, range_workbook

PROFILE = dict(scale_percent=75, margin_pt=18)
PAGE = dict(sheet=0, range='A1:G1', width_pt=841.89, height_pt=595.28, width_px=1754, height_px=1241)


def repack(parts):
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


class XlsxFidelityTests(unittest.TestCase):
    def test_print_settings_preserve_cells_extensions_and_other_parts(self):
        source = fixture()
        before, after = package(source), package(range_workbook(source, PROFILE, PAGE))
        for name in before.keys() - {'xl/workbook.xml', 'xl/worksheets/sheet1.xml'}:
            self.assertEqual(before[name], after[name])
        targets = [dict(sheet=0, cell=f'{column}1') for column in 'ABCDFG']
        self.assertEqual(results(before, targets), results(after, targets))
        sheet = minidom.parseString(after['xl/worksheets/sheet1.xml']).documentElement
        self.assertEqual(sheet.getAttribute('xmlns:x14ac'), 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
        self.assertEqual(sheet.getAttribute('mc:Ignorable'), 'x14ac')
        setup = children(sheet, 'pageSetup')[0]
        self.assertEqual(setup.getAttribute('scale'), '75')
        self.assertEqual(setup.getAttribute('paperSize'), '9')
        self.assertEqual(children(sheet, 'pageMargins')[0].getAttribute('left'), '0.25')
        book = minidom.parseString(after['xl/workbook.xml']).documentElement
        area = children(children(book, 'definedNames')[0], 'definedName')[0]
        self.assertEqual(area.firstChild.data, "'Sheet1'!$A$1:$G$1")

    def test_each_range_replaces_print_area_and_hides_other_sheets(self):
        parts = package(fixture())
        parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(b'</sheets>', b'<sheet name="O\'Brien" sheetId="2" r:id="rId2" state="hidden"/></sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$2</definedName><definedName name="KeepMe">42</definedName></definedNames>')
        parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace(b'</Relationships>', b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
        parts['xl/worksheets/sheet2.xml'] = parts['xl/worksheets/sheet1.xml']
        prepared = package(range_workbook(repack(parts), PROFILE, dict(PAGE, sheet=1, range='C4:D7')))
        book = minidom.parseString(prepared['xl/workbook.xml']).documentElement
        tabs = children(children(book, 'sheets')[0], 'sheet')
        self.assertEqual([tab.getAttribute('state') for tab in tabs], ['hidden', 'visible'])
        names = children(children(book, 'definedNames')[0], 'definedName')
        self.assertEqual([node.firstChild.data for node in names], ['42', "'O''Brien'!$C$4:$D$7"])
        self.assertEqual(names[-1].getAttribute('localSheetId'), '1')
        self.assertEqual(prepared['xl/worksheets/sheet1.xml'], parts['xl/worksheets/sheet1.xml'])

    def test_worksheet_index_skips_chart_tabs_but_print_area_uses_workbook_index(self):
        parts = package(fixture())
        parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(b'<sheets>', b'<sheets><sheet name="Chart" sheetId="2" r:id="chart"/>')
        parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace(b'</Relationships>', b'<Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>')
        parts['xl/chartsheets/sheet1.xml'] = b'<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
        prepared = package(range_workbook(repack(parts), PROFILE, PAGE))
        self.assertEqual(prepared['xl/chartsheets/sheet1.xml'], parts['xl/chartsheets/sheet1.xml'])
        book = minidom.parseString(prepared['xl/workbook.xml']).documentElement
        tabs = children(children(book, 'sheets')[0], 'sheet')
        self.assertEqual([tab.getAttribute('state') for tab in tabs], ['hidden', 'visible'])
        area = children(children(book, 'definedNames')[0], 'definedName')[0]
        self.assertEqual(area.getAttribute('localSheetId'), '1')
        self.assertEqual(area.firstChild.data, "'Sheet1'!$A$1:$G$1")

    def test_rejects_unknown_paper_and_invalid_ranges_instead_of_fitting(self):
        for page in [dict(PAGE, width_pt=100), dict(PAGE, range='B2:A1'), dict(PAGE, range='XFE1:XFE2'), dict(PAGE, sheet=3)]:
            with self.assertRaises(ValueError):
                range_workbook(fixture(), PROFILE, page)
        for profile in [dict(PROFILE, scale_percent=0), dict(PROFILE, margin_pt=float('nan'))]:
            with self.assertRaises(ValueError):
                range_workbook(fixture(), profile, PAGE)

    def test_extra_pages_remain_failed_comparisons_with_no_score(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'source.xlsx').write_bytes(fixture())
            sample = dict(id='sample', metadata=dict(source=dict(sha256='a' * 64), reference=dict(pages=1, capture_profile=dict(PROFILE, pages=[PAGE]))))
            def export(command, **kwargs):
                with fitz.open() as pdf:
                    pdf.new_page()
                    pdf.new_page()
                    pdf.save(Path(command[-1]).with_suffix('.pdf'))
            with patch('xlsx_fidelity.execute', side_effect=export):
                result = libreoffice_xlsx_fidelity(sample, root, 'unused', '26.2.3.2')
            self.assertEqual(result['status'], 'failed')
            self.assertEqual(result['stage'], 'capture')
            self.assertIn('produced 2 pages', result['error'])
            self.assertNotIn('penalized_ssim', result)


if __name__ == '__main__':
    unittest.main()
