import tempfile
import unittest
from pathlib import Path

import pypdfium2 as pdfium

from raster import RASTERIZER, open_pdf, page_size_pt, render_page


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
                self.assertEqual(render_page(pdf, 1, 150).size, (1754, 1240))
                self.assertEqual(render_page(pdf, 0, 96).size, (816, 1056))
        self.assertTrue(RASTERIZER.startswith('pypdfium2 '))

    def test_an_unreadable_file_is_a_value_error(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'broken.pdf'
            path.write_bytes(b'not a pdf')
            with self.assertRaises(ValueError):
                with open_pdf(path):
                    pass


if __name__ == '__main__':
    unittest.main()
