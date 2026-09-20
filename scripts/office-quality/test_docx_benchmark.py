import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image, ImageDraw

from docx_benchmark import CHANNELS, edge_adjust, measure_sample, sha256, validate_png


class NativeBenchmarkTests(unittest.TestCase):
    def image(self, path, size=(96, 96)):
        image = Image.new('RGB', size, 'white')
        ImageDraw.Draw(image).rectangle((10, 10, 30, 30), fill='black')
        image.save(path)

    def fixture(self, root):
        (root / 'reference').mkdir()
        self.image(root / 'reference/page_0001.png', (150, 150))
        (root / 'source.docx').write_bytes(b'docx fixture')
        return dict(id='fixture', metadata=dict(source=dict(sha256=sha256(root / 'source.docx'))))

    def test_times_fresh_commands_rotates_order_and_discards_warmup(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sample = self.fixture(root)
            calls = []

            def execute(command, **kwargs):
                channel = command[0]
                calls.append(channel)
                output = root / 'timing' / ('source.png' if channel == 'libreoffice' else f'{channel}.png')
                self.assertFalse(output.exists())
                self.assertIn(str(root / 'source.docx'), command)
                self.image(output)
                metadata = dict(status='ok', skipped_images=0, page=1, pages=2, dpi=96,
                                source_sha256=sample['metadata']['source']['sha256'], png_sha256=sha256(output))
                return len(calls) * 100, json.dumps(metadata)

            with patch('docx_benchmark.execute', execute), patch('builtins.print'):
                result = measure_sample(sample, root, {channel: Path(channel) for channel in CHANNELS[:2]},
                                        root / 'manifest.json', 'libreoffice', 5)
            self.assertEqual(len(calls), 18)
            self.assertEqual(calls[:9], CHANNELS + ['commit', 'libreoffice', 'published'] + ['libreoffice', 'published', 'commit'])
            self.assertEqual(result['published']['elapsed_ms'], [600, 800, 1000, 1500, 1700])
            self.assertTrue(all(len(row['elapsed_ms']) == 5 for row in result.values()))

    def test_failed_or_skipped_native_render_cannot_contribute_a_fast_time(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sample = self.fixture(root)

            def execute(command, **kwargs):
                channel = command[0]
                output = root / 'timing' / ('source.png' if channel == 'libreoffice' else f'{channel}.png')
                self.image(output)
                return 1, json.dumps(dict(status='ok', skipped_images=1, page=1, pages=1, dpi=96,
                                         source_sha256=sample['metadata']['source']['sha256'], png_sha256=sha256(output)))

            with patch('docx_benchmark.execute', execute), patch('builtins.print'):
                result = measure_sample(sample, root, {channel: Path(channel) for channel in CHANNELS[:2]},
                                        root / 'manifest.json', 'libreoffice', 5)
            self.assertEqual(result['commit']['status'], 'failed')
            self.assertNotIn('elapsed_ms', result['commit'])
            self.assertEqual(len(result['libreoffice']['elapsed_ms']), 5)

    def test_dimensions_and_nonblank_output_are_required(self):
        with tempfile.TemporaryDirectory() as directory:
            png = Path(directory) / 'page.png'
            self.image(png)
            self.assertEqual(validate_png(png, (95, 96), True)['size'], [96, 96])
            with self.assertRaisesRegex(ValueError, 'dimensions'):
                validate_png(png, (90, 96), True)
            Image.new('RGB', (96, 96), 'white').save(png)
            with self.assertRaisesRegex(ValueError, 'blank'):
                validate_png(png, (96, 96), True)
            validate_png(png, (96, 96), False)

    def test_edge_adjustment_does_not_resample_or_hide_larger_differences(self):
        original = Image.new('RGB', (10, 10), 'black')
        adjusted, change = edge_adjust(original, (11, 10))
        self.assertEqual(adjusted.getpixel((9, 9)), (0, 0, 0))
        self.assertEqual(adjusted.getpixel((10, 9)), (255, 255, 255))
        self.assertEqual(change, dict(original=[10, 10], target=[11, 10]))
        untouched, change = edge_adjust(original, (12, 10))
        self.assertIs(untouched, original)
        self.assertIsNone(change)


if __name__ == '__main__':
    unittest.main()
