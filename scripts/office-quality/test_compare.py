import json
from pathlib import Path
import tempfile
import unittest

from PIL import Image

from compare import Pages, compare


class ComparisonTests(unittest.TestCase):
    def test_gaps_in_page_names_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.pages(root, 'pages', 2)
            (root / 'pages/page_0002.png').rename(root / 'pages/page_0003.png')
            with self.assertRaisesRegex(ValueError, 'contiguous'):
                Pages(root / 'pages', 150)

    def pages(self, root, name, count, size=(32, 32), metadata=None):
        directory = root / name
        directory.mkdir()
        for index in range(count):
            Image.new('RGB', size, 'white').save(directory / f'page_{index + 1:04d}.png')
        if metadata:
            (directory / 'result.json').write_text(json.dumps(metadata))
        return Pages(directory, 150)

    def test_missing_and_extra_pages_reduce_the_score(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            one = self.pages(root, 'one', 1)
            two = self.pages(root, 'two', 2)
            for index, (reference, actual) in enumerate([(one, two), (two, one)]):
                report = compare(reference, actual, root / f'diff-{index}')
                self.assertEqual(report['common_page_ssim'], 1.0)
                self.assertEqual(report['penalized_ssim'], 0.5)
                self.assertTrue((root / f'diff-{index}/index.html').exists())

    def test_different_dimensions_require_explicit_resize(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference = self.pages(root, 'reference', 1)
            actual = self.pages(root, 'actual', 1, (40, 40))
            with self.assertRaisesRegex(ValueError, 'resize'):
                compare(reference, actual, root / 'strict')
            report = compare(reference, actual, root / 'resized', resize=True)
            self.assertEqual(report['penalized_ssim'], 1.0)
            self.assertEqual(report['pages'][0]['actual_size'], (40, 40))

    def test_mismatched_sources_and_incomplete_renders_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference = self.pages(root, 'reference', 1, metadata={'status': 'ok', 'sha256': 'a'})
            actual = self.pages(root, 'actual', 1, metadata={'status': 'ok', 'sha256': 'b'})
            with self.assertRaisesRegex(ValueError, 'different source'):
                compare(reference, actual, root / 'diff')
            with self.assertRaisesRegex(ValueError, 'did not complete'):
                self.pages(root, 'failed', 1, metadata={'status': 'timeout'})


if __name__ == '__main__':
    unittest.main()
