import copy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pdf_fixtures import pdf_bytes
from raster import open_pdf
from xlsx_reference import export_profile


class XlsxReferenceTests(unittest.TestCase):
    def test_ranges_keep_order_and_source_bytes(self):
        profile = {'scale_percent': 75, 'margin_pt': 18,
                   'pages': [{'sheet': 0, 'range': 'A1:B2'}, {'sheet': 1, 'range': 'A3:C4'}]}
        original = copy.deepcopy(profile)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'source.xlsx'
            source.write_bytes(b'unchanged source')

            def export(*args, **kwargs):
                for index in range(2):
                    (root / f'sheet-{index + 1}.pdf').write_bytes(
                        pdf_bytes([(842, 595, f'Page {index + 1}')], {'Creator': 'Excel', 'CreationDate': 'D:20260907190000Z'}))

            with patch('xlsx_reference.subprocess.run', side_effect=export) as run:
                result = export_profile(source, root / 'reference.pdf', profile, 30, 150)
            self.assertEqual(source.read_bytes(), b'unchanged source')
            self.assertEqual(profile, original)
            self.assertEqual(result['pages'][1]['range'], 'A3:C4')
            self.assertEqual(result['pages'][0]['width_px'], 1755)
            self.assertEqual(len(result['office_pdf_metadata']), 2)
            self.assertEqual(result['office_pdf_metadata'][0]['creator'], 'Excel')
            self.assertEqual(result['office_pdf_metadata'][0]['creationDate'], 'D:20260907190000Z')
            self.assertEqual(result['pages'][0]['width_pt'], 842)
            self.assertIn('set visible of sh to sheet visible', run.call_args.kwargs['input'])
            with open_pdf(root / 'reference.pdf') as document:
                self.assertEqual(len(document), 2)
                page = document[1]
                text = page.get_textpage()
                self.assertIn('Page 2', text.get_text_range())
                text.close()
                page.close()

    def test_rejects_overflowing_ranges(self):
        profile = {'scale_percent': 75, 'margin_pt': 18,
                   'pages': [{'sheet': 0, 'range': 'A1:B2'}]}
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'sheet-1.pdf').write_bytes(pdf_bytes([(612, 792, None), (612, 792, None)]))
            with patch('xlsx_reference.subprocess.run'), self.assertRaisesRegex(ValueError, 'produced 2 pages'):
                export_profile(root / 'source.xlsx', root / 'reference.pdf', profile, 30, 150)

    def test_rejects_invalid_profile_before_office(self):
        with patch('xlsx_reference.subprocess.run') as run:
            with self.assertRaises(ValueError):
                export_profile(Path('unused'), Path('unused'), {'scale_percent': 0}, 30, 150)
            run.assert_not_called()


if __name__ == '__main__':
    unittest.main()
