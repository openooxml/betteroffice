import tempfile
import unittest
from pathlib import Path

import pypdfium2 as pdfium

from pdf_fixtures import form_pdf_bytes, pdf_bytes
from raster import RASTERIZER, merge_pdfs, open_pdf, page_fonts, page_size_pt, pdf_metadata, render_page


class RasterTests(unittest.TestCase):
    def test_pages_render_on_white_at_the_requested_dpi(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'blank.pdf'
            document = pdfium.PdfDocument.new()
            document.new_page(612, 792).close()
            document.new_page(842, 595).close()
            document.save(str(path))
            document.close()
            with open_pdf(path) as pdf:
                self.assertEqual(len(pdf), 2)
                self.assertEqual(page_size_pt(pdf, 1), (842, 595))
                image = render_page(pdf, 0, 150)
                self.assertEqual(image.mode, 'RGB')
                self.assertEqual(image.size, (1275, 1650))
                self.assertEqual(image.getextrema(), ((255, 255), (255, 255), (255, 255)))
                self.assertEqual(render_page(pdf, 1, 150).size, (1755, 1240))
                self.assertEqual(render_page(pdf, 0, 96).size, (816, 1056))
        self.assertTrue(RASTERIZER.startswith('pypdfium2 '))

    def test_an_unreadable_file_is_a_value_error(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'broken.pdf'
            path.write_bytes(b'not a pdf')
            with self.assertRaises(ValueError):
                with open_pdf(path):
                    pass

    def test_metadata_fonts_and_merge_match_what_reference_capture_records(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'a.pdf').write_bytes(pdf_bytes([(842, 595, 'Alpha')], {'Creator': 'Excel', 'Producer': 'macOS'}))
            (root / 'b.pdf').write_bytes(pdf_bytes([(842, 595, 'Beta')]))
            with open_pdf(root / 'a.pdf') as document:
                metadata = pdf_metadata(document)
                self.assertEqual(sorted(metadata), sorted(['format', 'title', 'author', 'subject', 'keywords', 'creator',
                                                           'producer', 'creationDate', 'modDate', 'trapped', 'encryption']))
                self.assertEqual((metadata['format'], metadata['creator'], metadata['producer']), ('PDF 1.7', 'Excel', 'macOS'))
                self.assertEqual(page_fonts(document, 0), {'Helvetica'})
            merge_pdfs([root / 'a.pdf', root / 'b.pdf'], root / 'merged.pdf')
            with open_pdf(root / 'merged.pdf') as document:
                self.assertEqual(len(document), 2)
                self.assertEqual(page_size_pt(document, 1), (842, 595))

    def test_form_field_appearances_are_drawn(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'form.pdf'
            path.write_bytes(form_pdf_bytes())
            with open_pdf(path) as pdf:
                image = render_page(pdf, 0, 72).convert('L')
            self.assertEqual(image.getpixel((200, 642)), 0)
            self.assertEqual(image.getpixel((50, 50)), 255)


if __name__ == '__main__':
    unittest.main()
