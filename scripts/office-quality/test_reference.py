import tempfile
import unittest
from pathlib import Path

from pdf_fixtures import pdf_bytes
from raster import merge_pdfs, open_pdf, pdf_metadata
from reference import recorded_pdf_metadata


class ReferenceMetadataTests(unittest.TestCase):
    def test_a_merged_range_capture_records_the_office_export_metadata(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'sheet-1.pdf').write_bytes(pdf_bytes([(842, 595, 'Alpha')], {'Creator': 'Excel', 'CreationDate': 'D:20260907190000Z'}))
            (root / 'sheet-2.pdf').write_bytes(pdf_bytes([(842, 595, 'Beta')], {'Creator': 'Excel'}))
            exports = []
            for name in ('sheet-1.pdf', 'sheet-2.pdf'):
                with open_pdf(root / name) as document:
                    exports.append(pdf_metadata(document))
            merge_pdfs([root / 'sheet-1.pdf', root / 'sheet-2.pdf'], root / 'reference.pdf')
            with open_pdf(root / 'reference.pdf') as document:
                recorded = recorded_pdf_metadata(dict(kind='excel-range', office_pdf_metadata=exports), document)
            self.assertEqual((recorded['creator'], recorded['creationDate']), ('Excel', 'D:20260907190000Z'))

    def test_a_single_export_records_its_own_metadata(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'reference.pdf'
            path.write_bytes(pdf_bytes([(612, 792, 'Alpha')], {'Creator': 'Word'}))
            with open_pdf(path) as document:
                self.assertEqual(recorded_pdf_metadata(dict(kind='office-page-bounds', pages=[]), document)['creator'], 'Word')
                self.assertEqual(recorded_pdf_metadata(None, document)['creator'], 'Word')


if __name__ == '__main__':
    unittest.main()
