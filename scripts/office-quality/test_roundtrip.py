import json
import signal
import subprocess
from types import SimpleNamespace
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from roundtrip import A, W, at_path, digest, measure, package, select_probe, verify_preservation, xml
from test_xlsx_calc import fixture


def docx():
    return {'word/document.xml': f'<w:document xmlns:w="{W}" xmlns:custom="urn:custom"><w:body><w:p><w:r><w:t>Example</w:t></w:r></w:p><custom:extension preserved="yes"/><!--keep--></w:body></w:document>'.encode(),
            'custom/opaque.bin': b'original binary payload'}


def edited(parts, probe):
    result = dict(parts)
    document = xml(parts[probe['part']])
    node = at_path(document, probe['path'])
    node.firstChild.data = probe['new']
    result[probe['part']] = document.toxml(encoding='utf-8')
    return result


class PreservationTests(unittest.TestCase):
    def test_selects_and_accepts_exact_edits_in_all_three_formats(self):
        deck = {'ppt/slides/slide1.xml': f'<slide xmlns:a="{A}"><txBody><a:p><a:r><a:rPr b="1"/><a:t>Example</a:t></a:r></a:p></txBody></slide>'.encode()}
        for format, parts in [('docx', docx()), ('pptx', deck), ('xlsx', package(fixture()))]:
            probe = select_probe(parts, format)
            self.assertEqual(probe['status'], 'ok')
            self.assertNotEqual(probe['old'], probe['new'])
            result = verify_preservation(parts, edited(parts, probe), probe)
            self.assertEqual(result['status'], 'ok')
            self.assertEqual(result['identical_parts'], len(parts) - 1)
        probe = select_probe(package(fixture()), 'xlsx')
        self.assertEqual(probe['cell'], 'A1')
        self.assertEqual(probe['new'], '0')

    def test_rejects_noop_wrong_edit_and_loss_inside_the_edited_part(self):
        parts = docx()
        probe = select_probe(parts, 'docx')
        after = edited(parts, probe)
        for bad in [parts,
                    {**after, probe['part']: after[probe['part']].replace(b'preserved="yes"', b'preserved="no"')},
                    {**after, probe['part']: after[probe['part']].replace(b'<!--keep-->', b'')},
                    {**after, probe['part']: after[probe['part']].replace(b'urn:custom', b'urn:wrong')},
                    {**after, probe['part']: after[probe['part']].replace(probe['new'].encode(), b'Wrong edit')}]:
            self.assertEqual(verify_preservation(parts, bad, probe)['status'], 'failed')

    def test_rejects_changed_missing_or_added_unrelated_parts(self):
        parts = docx()
        probe = select_probe(parts, 'docx')
        after = edited(parts, probe)
        for bad in [{**after, 'custom/opaque.bin': b'corrupted'},
                    {probe['part']: after[probe['part']]},
                    {**after, 'new.xml': b'<added/>'}]:
            self.assertEqual(verify_preservation(parts, bad, probe)['status'], 'failed')

    def test_rejects_ambiguous_or_unavailable_probes(self):
        parts = docx()
        parts['word/document.xml'] = parts['word/document.xml'].replace(b'</w:body>', b'<w:p><w:r><w:t>Example</w:t></w:r></w:p></w:body>')
        self.assertEqual(select_probe(parts, 'docx')['status'], 'unavailable')
        parts = docx()
        parts['word/document.xml'] = parts['word/document.xml'].replace(b'Example', b'Ex<!--preserve-->ample')
        self.assertEqual(select_probe(parts, 'docx')['status'], 'unavailable')

    def test_parse_success_survives_a_later_save_failure(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source.docx'
            source.write_bytes(b'synthetic source')
            probe = root / 'probe.json'
            probe.write_text('{"status":"unavailable"}')
            def start(command, timeout, log):
                Path(command[-1]).write_text(json.dumps(dict(source_sha256=digest(source.read_bytes()), parse='ok', stage='save', error='save failed')))
            with patch('roundtrip.run_process', side_effect=start):
                result = measure(Path('/unused-native'), source, probe, root / 'result')
            self.assertEqual(result['parse']['status'], 'ok')
            self.assertEqual(result['roundtrip']['status'], 'failed')
            self.assertEqual(result['roundtrip']['stage'], 'save')


    def test_timeout_kills_the_process_group_and_retains_parse_success(self):
        from roundtrip import run_process
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            process = SimpleNamespace(pid=12345, returncode=None)
            process.wait = lambda **kwargs: (_ for _ in ()).throw(subprocess.TimeoutExpired('engine', 180)) if kwargs else None
            class Context:
                def __enter__(self):
                    return process
                def __exit__(self, *args):
                    pass
            with patch('roundtrip.subprocess.Popen', return_value=Context()), patch('roundtrip.os.killpg') as kill:
                with self.assertRaisesRegex(RuntimeError, 'timed out'):
                    run_process(['unused'], 180, root / 'log.txt')
                kill.assert_called_once_with(12345, signal.SIGKILL)
            source = root / 'source.docx'
            source.write_bytes(b'original')
            probe = root / 'probe.json'
            probe.write_text('{}')
            def timeout(command, *args):
                Path(command[-1]).write_text(json.dumps(dict(source_sha256=digest(source.read_bytes()), parse='ok', stage='save')))
                raise RuntimeError('Process timed out after 180s')
            with patch('roundtrip.run_process', side_effect=timeout):
                result = measure(Path('/unused'), source, probe, root / 'timeout')
            self.assertEqual(result['parse']['status'], 'ok')
            self.assertEqual(result['roundtrip']['status'], 'failed')
            self.assertIn('timed out', result['roundtrip']['error'])

    def test_checker_timeout_records_a_failure_after_a_verified_edit(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'source.docx'
            source.write_bytes(b'original')
            probe = root / 'probe.json'
            probe.write_text('{}')
            def finish(command, *args):
                Path(command[-2]).write_bytes(b'edited')
                Path(command[-1]).write_text(json.dumps(dict(source_sha256=digest(source.read_bytes()), parse='ok', stage='complete',
                    edit_verified=True, output_sha256=digest(b'edited'))))
            with patch('roundtrip.run_process', side_effect=finish), patch('roundtrip.helper', side_effect=RuntimeError('checker timed out')):
                result = measure(Path('/unused'), source, probe, root / 'checked')
            self.assertEqual(result['parse']['status'], 'ok')
            self.assertEqual(result['roundtrip']['stage'], 'preserve')
            self.assertEqual(result['roundtrip']['status'], 'failed')
            self.assertIn('checker timed out', result['roundtrip']['error'])


if __name__ == '__main__':
    unittest.main()
