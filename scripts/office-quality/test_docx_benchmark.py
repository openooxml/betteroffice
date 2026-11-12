import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image, ImageDraw

from docx_benchmark import (CHANNELS, LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS, edge_adjust, font_path_notes, measure_sample,
                            semver_key, sha256, validate_font_paths, validate_png)

FACES = ['a' * 64, 'b' * 64, 'a' * 64 + '~Calibri']


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
                metadata = dict(status='ok', skipped_images=0, page=1, pages=2, dpi=96, faces=FACES,
                                source_sha256=sample['metadata']['source']['sha256'], png_sha256=sha256(output))
                return len(calls) * 100, json.dumps(metadata)

            with patch('docx_benchmark.execute', execute), patch('builtins.print'):
                result = measure_sample(sample, root, {channel: Path(channel) for channel in CHANNELS[:2]},
                                        root / 'manifest.json', 'libreoffice', 5)
            self.assertEqual(len(calls), 18)
            self.assertEqual(calls[:9], CHANNELS + ['commit', 'libreoffice', 'published'] + ['libreoffice', 'published', 'commit'])
            self.assertEqual(result['published']['elapsed_ms'], [600, 800, 1000, 1500, 1700])
            self.assertTrue(all(len(row['elapsed_ms']) == 5 for row in result.values()))
            replayed = hashlib.sha256('\n'.join(FACES).encode()).hexdigest()
            self.assertEqual([result[channel]['fonts_sha256'] for channel in CHANNELS[:2]], [replayed, replayed])
            self.assertNotIn('fonts_sha256', result['libreoffice'])

    def test_the_fonts_hash_covers_the_faces_each_native_render_replayed(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sample = self.fixture(root)
            runs = []

            def execute(command, **kwargs):
                channel = command[0]
                output = root / 'timing' / ('source.png' if channel == 'libreoffice' else f'{channel}.png')
                self.image(output)
                runs.append(channel)
                faces = FACES if channel == 'published' or runs.count(channel) == 1 else FACES[::-1]
                return 1, json.dumps(dict(status='ok', skipped_images=0, page=1, pages=1, dpi=96, faces=faces,
                                          source_sha256=sample['metadata']['source']['sha256'], png_sha256=sha256(output)))

            with patch('docx_benchmark.execute', execute), patch('builtins.print'):
                result = measure_sample(sample, root, {channel: Path(channel) for channel in CHANNELS[:2]},
                                        root / 'manifest.json', 'libreoffice', 2)
            self.assertEqual(result['published']['fonts_sha256'], hashlib.sha256('\n'.join(FACES).encode()).hexdigest())
            self.assertEqual(result['commit']['status'], 'failed')
            self.assertIn('different font faces', result['commit']['error'])

    def test_faces_are_required_for_docx_native_timings_only(self):
        for format, status in [('pptx', 'ok'), ('docx', 'failed')]:
            with self.subTest(format=format), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                sample = self.fixture(root)
                (root / f'source.{format}').write_bytes(b'fixture')
                sample = dict(sample, format=format, metadata=dict(source=dict(sha256=sha256(root / f'source.{format}'))))

                def execute(command, **kwargs):
                    channel = command[0]
                    output = root / 'timing' / ('source.png' if channel == 'libreoffice' else f'{channel}.png')
                    self.image(output)
                    return 1, json.dumps(dict(status='ok', skipped_images=0, page=1, pages=1, dpi=96,
                                              source_sha256=sample['metadata']['source']['sha256'], png_sha256=sha256(output)))

                with patch('docx_benchmark.execute', execute), patch('builtins.print'):
                    result = measure_sample(sample, root, {channel: Path(channel) for channel in CHANNELS[:2]},
                                            root / 'manifest.json', 'libreoffice', 2)
                for channel in CHANNELS[:2]:
                    self.assertEqual(result[channel]['status'], status)
                    self.assertNotIn('fonts_sha256', result[channel])
                if format == 'pptx':
                    self.assertEqual(len(result['published']['elapsed_ms']), 2)
                else:
                    self.assertIn('no list of replayed font faces', result['commit']['error'])

    def test_the_old_font_map_is_allowed_only_for_a_published_release_no_newer_than_the_last_without_the_registry_exports(self):
        self.assertEqual(LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS, '0.5.1')
        old = dict(path='bench-manifest', version='0.5.1', last_release_without_registry_exports=LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS)
        commit = dict(path='product-registry', version='0.5.1')
        self.assertEqual(validate_font_paths(dict(published=old, commit=commit), '0.5.1')['published'], old)
        self.assertLess(semver_key('0.6.0-next.1'), semver_key('0.6.0'))
        self.assertLess(semver_key('0.6.0-next.2'), semver_key('0.6.0-next.10'))
        self.assertLess(semver_key('0.5.10'), semver_key('0.6.0'))
        validate_font_paths(dict(published=dict(old, version='0.5.0'), commit=commit), '0.5.0')
        for published, version, error in [
            (dict(old, version='0.5.2'), '0.5.2', 'is newer than 0.5.1'),
            (dict(old, version='0.6.0-next.1'), '0.6.0-next.1', 'is newer than 0.5.1'),
            (dict(old, version='0.6.0'), '0.6.0', 'is newer than 0.5.1'),
            (dict(old, version='0.10.0'), '0.10.0', 'is newer than 0.5.1'),
            (dict(old, last_release_without_registry_exports='0.5.0'), '0.5.1', 'last release without'),
            (dict(old, version='latest'), 'latest', 'Not a semver'),
            (dict(old, path='system'), '0.5.1', 'Unknown font path'),
            (old, '0.6.0', 'plan measures'),
        ]:
            with self.assertRaisesRegex(ValueError, error):
                validate_font_paths(dict(published=published, commit=commit), version)
        with self.assertRaisesRegex(ValueError, 'Invalid product font chain paths'):
            validate_font_paths(dict(published=old, commit=dict(commit, path='bench-manifest')), '0.5.1')
        notes = font_path_notes(dict(published=old, commit=commit), 'docx-cli-page1-v2')
        self.assertIn('not comparable with runs before docx-cli-page1-v2', notes['commit'])
        self.assertNotIn('old font map', notes['commit'])
        self.assertIn('0.5.1 <= 0.5.1, the last release without the registry exports', notes['published'])
        self.assertIn('not like for like until the next release', notes['published'])

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
