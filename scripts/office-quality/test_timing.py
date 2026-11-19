import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import timing
from docx_benchmark import CHANNELS, sha256

LEAKS = ['host-a', 'host-b', 'Linux-test-x86_64']


class TimingReportTests(unittest.TestCase):
    def run_timing(self, environment, drop=()):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output, native = root / 'output', root / 'native'
            (output / 'fixture' / 'reference').mkdir(parents=True)
            (output / 'fixture' / 'source.docx').write_bytes(b'docx fixture')
            (output / 'fixture' / 'reference' / 'page_0001.png').write_bytes(b'page fixture')
            for channel, source in [('published', 'b' * 40), ('commit', 'a' * 40)]:
                binary = native / channel / 'docx-native-bench'
                binary.parent.mkdir(parents=True)
                binary.write_bytes(channel.encode())
                (binary.parent / 'build.json').write_text(json.dumps(dict(source_sha=source, binary_sha256=sha256(binary))))
            sample = dict(id='fixture', metadata=dict(
                source=dict(sha256=sha256(output / 'fixture' / 'source.docx')),
                reference_pages=[dict(sha256=sha256(output / 'fixture' / 'reference' / 'page_0001.png'))]))
            (output / 'job.json').write_text(json.dumps(dict(
                plan_sha256='c' * 64, format='docx', samples=[sample], source_sha='a' * 40, published_source_sha='b' * 40,
                method='sequential-paired-sample-v1', rounds=1, seed='seed', sample_size=200, published_version='0.5.1')))
            variables = dict(QUALITY_OUTPUT=str(output), QUALITY_NATIVE=str(native), QUALITY_FONTS=str(root / 'manifest.json'),
                             SOFFICE='libreoffice', **environment)
            with patch.dict(os.environ, variables), \
                    patch('platform.node', lambda *args, **kwargs: 'host-b'), \
                    patch('platform.platform', lambda *args, **kwargs: 'Linux-test-x86_64'), \
                    patch('timing.font_environment', lambda *args: 'f' * 64), \
                    patch('timing.execute', lambda command, **kwargs: (0, 'LibreOffice 26.2.3.2 ' + 'd' * 40)), \
                    patch('timing.measure_round', lambda *args: {channel: dict(status='ok', elapsed_ms=[5.0]) for channel in CHANNELS}), \
                    patch('builtins.print'):
                for name in drop:
                    os.environ.pop(name, None)
                timing.main()
            return json.loads((output / 'report.json').read_text())

    def test_the_report_records_no_host_name_and_no_kernel_string(self):
        processors = os.cpu_count()
        for environment, drop, runner in [
                (dict(RUNNER_NAME='host-a', ImageVersion='20251110.1'), (), dict(processors=processors, image_version='20251110.1')),
                ({}, ('RUNNER_NAME', 'ImageVersion'), dict(processors=processors, image_version='local')),
                (dict(ImageVersion='odd-image-7'), ('RUNNER_NAME',), dict(processors=processors))]:
            report = self.run_timing(environment, drop)
            self.assertEqual(report['runner'], runner)
            text = json.dumps(report)
            for leak in [*LEAKS, 'odd-image-7']:
                self.assertNotIn(leak, text)

    def test_an_image_version_that_is_not_a_plain_ascii_dated_string_is_left_out(self):
        for value in ['٢٠٢٥١١١٠.1', ' 20251110.1', '20251110.1 ', '20251110.1\n']:
            with self.subTest(value=value):
                report = self.run_timing(dict(ImageVersion=value), ('RUNNER_NAME',))
                self.assertEqual(report['runner'], dict(processors=os.cpu_count()))


if __name__ == '__main__':
    unittest.main()
