import contextlib
import io
import json
import os
import sys
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from pdf_matrix import (
    ALLOWED_UNSUPPORTED_FIXTURES,
    GENERATED_TEXT_FIXTURES,
    UNSUPPORTED_MISSING_GLYPH,
    DocxInfo,
    describe_verapdf_failure,
    find_verapdf,
    group_for,
    load_manifest,
    main,
    make_large_docx,
    manifest_declares_pdfa_2u,
    is_allowed_unsupported,
    is_emoji_char,
    make_text_docx,
    normalize_text,
    parse_page_size,
    parse_pdffonts,
    parse_verapdf_json,
    parse_verapdf_raw,
    parse_verapdf_report,
    read_docx_info,
    scripts_in,
)


def write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('word/document.xml', document_xml)


def minimal_pdf() -> bytes:
    """A one-page A4 PDF with no content and a byte-accurate cross-reference."""
    objects = [
        b'<< /Type /Catalog /Pages 2 0 R >>',
        b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> >>',
    ]
    output = bytearray(b'%PDF-1.4\n')
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(output))
        output += f'{number} 0 obj\n'.encode() + body + b'\nendobj\n'
    xref = len(output)
    output += f'xref\n0 {len(objects) + 1}\n'.encode()
    output += b'0000000000 65535 f \n'
    for offset in offsets:
        output += f'{offset:010d} 00000 n \n'.encode()
    output += (
        f'trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n'
        f'startxref\n{xref}\n%%EOF\n'
    ).encode()
    return bytes(output)


def write_pdf(path: Path) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = minimal_pdf()
    path.write_bytes(data)
    return len(data)


def write_manifest(
    path: Path,
    entries: list[dict[str, object]],
    header: dict[str, object] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload: dict[str, object] = {**(header or {}), 'fixtures': entries}
    path.write_text(json.dumps(payload))


def ok_entry(relative: str, pdf_bytes: int) -> dict[str, object]:
    return {
        'path': relative,
        'status': 'ok',
        'pages': 1,
        'pdf_bytes': pdf_bytes,
        'fonts_embedded': 0,
        'text_chars': 0,
    }


def run_main(arguments: list[str]) -> tuple[int, str, str]:
    stdout, stderr = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        code = main(arguments)
    return code, stdout.getvalue(), stderr.getvalue()


class FakeTextPage:
    def __init__(self, text: str) -> None:
        self.text = text

    def get_text_range(self) -> str:
        return self.text

    def close(self) -> None:
        pass


class FakePage:
    def __init__(self, text: str) -> None:
        self.text = text

    def get_width(self) -> float:
        return 612.0

    def get_height(self) -> float:
        return 792.0

    def get_textpage(self) -> FakeTextPage:
        return FakeTextPage(self.text)

    def close(self) -> None:
        pass


class FakeDocument:
    def __init__(self, text: str) -> None:
        self.text = text

    def __len__(self) -> int:
        return 1

    def __getitem__(self, index: int) -> FakePage:
        return FakePage(self.text)

    def close(self) -> None:
        pass


def fake_pdfium(text: str) -> types.ModuleType:
    """A pypdfium2 stand-in whose one page extracts `text` verbatim."""
    module = types.ModuleType('pypdfium2')
    module.PdfDocument = lambda path: FakeDocument(text)
    module.__version__ = 'test'
    return module


class ScriptTests(unittest.TestCase):
    def test_detects_each_script_family(self):
        self.assertEqual(scripts_in('abc 123'), set())
        self.assertIn('rtl', scripts_in('مرحبا بالعالم'))
        self.assertIn('cjk', scripts_in('日本語の文章'))
        self.assertIn('cjk', scripts_in('한국어'))
        self.assertIn('emoji', scripts_in('😀🎉'))

    def test_normalize_text_removes_layout_artifacts(self):
        self.assertEqual(normalize_text('مرح با بالعالم'), 'مرحبابالعالم')
        self.assertEqual(normalize_text('שלום עולם'), 'שלוםעולם')
        self.assertEqual(normalize_text('\u202babc\u202c\u200f'), 'abc')
        self.assertTrue(is_emoji_char('🎉'))
        self.assertFalse(is_emoji_char('A'))

    def test_group_uses_staged_top_level_directory(self):
        info = DocxInfo(scripts={'cjk'})
        for group in ('table-merged', 'header-footer', 'footnotes', 'pagination', 'images',
                      'cjk-rtl-emoji', 'large'):
            self.assertEqual(group_for(Path(group) / 'nested' / 'x.docx', info, 0), group)

    def test_group_infers_from_content_before_name(self):
        pages = Path('page-fragments/pages.docx')
        self.assertEqual(group_for(pages, DocxInfo(scripts={'emoji'}), 0), 'cjk-rtl-emoji')
        self.assertEqual(group_for(Path('table-widths/a.docx'), DocxInfo(), 0), 'table-merged')
        self.assertEqual(group_for(Path('header-footer-reservation/a.docx'), DocxInfo(), 0),
                         'header-footer')
        self.assertEqual(group_for(Path('layout-request/a.docx'), DocxInfo(has_notes=True), 0),
                         'footnotes')


class ManifestTests(unittest.TestCase):
    def test_accepts_object_array_and_single_record(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'manifest.json'
            record = {'path': 'a.docx', 'status': 'ok'}
            for payload in ([record], {'fixtures': [record]}, record):
                path.write_text(json.dumps(payload))
                header, entries = load_manifest(path)
                self.assertEqual(entries, [record])
                self.assertEqual(header, {})

    def test_rejects_entries_without_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'manifest.json'
            path.write_text(json.dumps({'fixtures': [{'status': 'ok'}]}))
            with self.assertRaisesRegex(ValueError, 'string path'):
                load_manifest(path)

    def test_parse_page_size(self):
        self.assertEqual(parse_page_size('612x792'), (612.0, 792.0))
        self.assertEqual(parse_page_size(' 595.28 X 841.89 '), (595.28, 841.89))
        with self.assertRaises(Exception):
            parse_page_size('0x792')


class DocxTests(unittest.TestCase):
    def test_reads_section_size_scripts_and_notes(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'fixture.docx'
            write_docx(
                path,
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body><w:p><w:r><w:t>&#26085;&#26412;&#35486; &#128512;</w:t></w:r>'
                '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
                '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
            )
            info = read_docx_info(path)
            self.assertEqual(info.page_sizes, {(612.0, 792.0)})
            self.assertEqual(info.scripts, {'cjk', 'emoji'})
            self.assertTrue(info.has_notes)
            self.assertIsNone(info.error)

    def test_generated_text_fixtures_round_trip_through_the_parser(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name, paragraphs, expected in GENERATED_TEXT_FIXTURES:
                path = root / name
                make_text_docx(path, paragraphs)
                info = read_docx_info(path)
                self.assertIsNone(info.error)
                self.assertEqual(info.page_sizes, {(612.0, 792.0)})
                for text in expected:
                    self.assertIn(text, info.source_text)
                self.assertEqual(
                    info.scripts,
                    scripts_in(
                        ' '.join(text for paragraph in paragraphs for text, _, _ in paragraph)
                    ),
                )

    def test_missing_document_part_is_reported(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'broken.docx'
            with zipfile.ZipFile(path, 'w') as archive:
                archive.writestr('word/styles.xml', '<styles/>')
            info = read_docx_info(path)
            self.assertIn('word/document.xml', info.error)

    def test_large_document_repeats_the_body_but_keeps_one_section(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source.docx'
            body = '<w:p><w:r><w:t>row</w:t></w:r></w:p>'
            opening = ('<w:document xmlns:w="http://schemas.openxmlformats.org/'
                       'wordprocessingml/2006/main"><w:body>')
            section = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
            write_docx(source, opening + body + section + '</w:body></w:document>')
            target = Path(directory) / 'large.docx'
            make_large_docx(source, target, 7)
            with zipfile.ZipFile(target) as archive:
                xml = archive.read('word/document.xml').decode()
            self.assertEqual(xml.count('<w:p>'), 7)
            self.assertEqual(xml.count('<w:sectPr'), 1)


class PdffontsTests(unittest.TestCase):
    def test_parses_embedded_and_unembedded_rows(self):
        output = (
            'name                                 type              encoding         '
            'emb sub uni object ID\n'
            '------------------------------------ ----------------- ---------------- '
            '--- --- --- ---------\n'
            'ABCDEF+Calibri                       TrueType          WinAnsi          '
            'yes yes yes     12  0\n'
            'Helvetica                            Type1             WinAnsi          '
            'no  no  no       3  0\n'
        )
        fonts = parse_pdffonts(output)
        self.assertEqual(len(fonts), 2)
        self.assertEqual(fonts[0]['emb'], 'yes')
        self.assertEqual(fonts[0]['uni'], 'yes')
        self.assertEqual(fonts[1]['emb'], 'no')


class EmojiExpectationTests(unittest.TestCase):
    def run_case(
        self,
        extracted: str,
        expected: list[str] | None,
        *,
        source_text: str | None = None,
    ) -> tuple[int, dict[str, object]]:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'fixtures'
            source.mkdir()
            relative = 'emoji/report.docx'
            if expected is not None:
                (source / 'expected.json').write_text(
                    json.dumps({relative: {'text': expected}})
                )
            if source_text is not None:
                (source / relative).parent.mkdir(parents=True, exist_ok=True)
                write_docx(
                    source / relative,
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                    'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>'
                    f'{source_text}'
                    '</w:t></w:r></w:p><w:sectPr>'
                    '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
                )
            size = write_pdf(root / 'pdfs' / 'emoji' / 'report.pdf')
            write_manifest(root / 'pdfs' / 'manifest.json', [ok_entry(relative, size)])
            report_path = root / 'report.json'
            with mock.patch.dict(sys.modules, {'pypdfium2': fake_pdfium(extracted)}):
                code, stdout, stderr = run_main([
                    str(root / 'pdfs' / 'manifest.json'),
                    '--fixtures-dir', str(source),
                    '--json', str(report_path),
                ])
            self.assertTrue(report_path.is_file(), (code, stdout, stderr))
            return code, json.loads(report_path.read_text())

    def test_missing_explicit_emoji_fails_the_fixture(self):
        code, report = self.run_case('plain text', ['😀'])
        self.assertEqual(code, 1, report)
        self.assertEqual(len(report['failures']), 1, report['failures'])
        checks = report['failures'][0]['checks']
        self.assertIn('emoji', [check for check, _ in checks])
        message = next(message for check, message in checks if check == 'emoji')
        self.assertIn('😀', message)
        coverage = report['fixtures'][0]['emoji_glyph_coverage']
        self.assertEqual(coverage['missing'], ['😀'])

    def test_recovered_explicit_emoji_passes_the_fixture(self):
        code, report = self.run_case('recovered 😀', ['😀'])
        self.assertEqual(code, 0, report['failures'])
        self.assertEqual(
            report['fixtures'][0]['emoji_glyph_coverage'],
            {'expected': ['😀'], 'recovered': ['😀'], 'missing': []},
        )

    def test_emoji_without_an_expectation_stays_a_glyph_warning(self):
        code, report = self.run_case('plain text', None, source_text='😀')
        self.assertEqual(code, 0, report['failures'])
        self.assertFalse(
            any(check == 'emoji' for failure in report['failures']
                for check, _ in failure['checks'])
        )
        self.assertTrue(
            any('no extractor recovered one' in warning for warning in report['warnings']),
            report['warnings'],
        )


class MainTests(unittest.TestCase):
    def test_same_relative_path_in_different_export_roots_is_checked_per_root(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first, second = root / 'first', root / 'second'
            size = write_pdf(first / 'report.pdf')
            write_pdf(second / 'report.pdf')
            write_manifest(first / 'manifest.json', [ok_entry('report.docx', size)])
            write_manifest(second / 'manifest.json', [
                {'path': 'report.docx', 'status': 'error', 'error': 'export failed'},
            ])
            report_path = root / 'report.json'
            code, stdout, stderr = run_main([
                str(first / 'manifest.json'),
                str(second / 'manifest.json'),
                '--json', str(report_path),
            ])
            report = json.loads(report_path.read_text())

            self.assertEqual(code, 1, (stdout, stderr))
            self.assertEqual(len(report['fixtures']), 2)
            self.assertEqual(
                [fixture['status'] for fixture in report['fixtures']], ['ok', 'error']
            )
            self.assertEqual(len(report['failures']), 1, report['failures'])
            check, message = report['failures'][0]['checks'][0]
            self.assertEqual(check, 'export')
            self.assertIn('export failed', message)
            self.assertFalse(
                any('duplicate fixture path' in warning for warning in report['warnings'])
            )

    def test_expectations_apply_to_the_relative_path_across_export_roots(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / 'fixtures'
            source.mkdir()
            (source / 'expected.json').write_text(
                json.dumps({'report.docx': {'text': ['MUST APPEAR']}})
            )
            manifests = []
            for name in ('first', 'second'):
                size = write_pdf(root / name / 'report.pdf')
                manifest = root / name / 'manifest.json'
                write_manifest(manifest, [ok_entry('report.docx', size)])
                manifests.append(str(manifest))
            report_path = root / 'report.json'
            code, stdout, stderr = run_main([
                *manifests,
                '--fixtures-dir', str(source),
                '--json', str(report_path),
            ])
            report = json.loads(report_path.read_text())

            self.assertEqual(code, 1, (stdout, stderr))
            self.assertEqual(len(report['fixtures']), 2)
            self.assertEqual(len(report['failures']), 2, report['failures'])
            for failure in report['failures']:
                check, message = failure['checks'][0]
                self.assertEqual(check, 'text')
                self.assertIn('MUST APPEAR', message)

    def test_same_manifest_entry_is_still_deduplicated(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            size = write_pdf(root / 'report.pdf')
            manifest = root / 'manifest.json'
            write_manifest(manifest, [ok_entry('report.docx', size)])
            report_path = root / 'report.json'
            code, stdout, stderr = run_main([
                str(manifest), str(manifest), '--json', str(report_path),
            ])
            report = json.loads(report_path.read_text())

            self.assertEqual(code, 0, (stdout, stderr))
            self.assertEqual(len(report['fixtures']), 1)
            self.assertTrue(
                any('duplicate fixture path' in warning for warning in report['warnings'])
            )


def verapdf_rule(clause: str, message: str) -> dict[str, object]:
    return {
        'ruleStatus': 'FAILED',
        'specification': 'ISO 19005-2:2011',
        'clause': clause,
        'testNumber': 1,
        'status': 'failed',
        'failedChecks': 1,
        'description': 'archival rule',
        'checks': [
            {
                'status': 'failed',
                'context': 'root',
                'errorMessage': message,
                'errorArguments': [],
            }
        ],
    }


def verapdf_json(compliant: bool, rules: list[dict[str, object]] | None = None) -> str:
    return json.dumps(
        {
            'report': {
                'buildInformation': {},
                'jobs': [
                    {
                        'itemDetails': {'name': 'report.pdf', 'size': 1},
                        'validationResult': [
                            {
                                'compliant': compliant,
                                'jobEndStatus': 'normal',
                                'statement': (
                                    'PDF file is compliant with Validation Profile requirements.'
                                    if compliant
                                    else 'PDF file is not compliant with Validation Profile requirements.'
                                ),
                                'details': {
                                    'passedRules': 1,
                                    'failedRules': len(rules or []),
                                    'passedChecks': 1,
                                    'failedChecks': len(rules or []),
                                    'ruleSummaries': rules or [],
                                },
                            }
                        ],
                    }
                ],
            }
        }
    )


def verapdf_raw(compliant: bool, failures: list[tuple[str, str, str]] | None = None) -> str:
    assertions = ''.join(
        '<assertion ordinal="1" status="FAILED">'
        f'<ruleId specification="ISO_19005_2" clause="{clause}" testNumber="1"/>'
        f'<message>{description}</message><errorMessage>{message}</errorMessage>'
        '</assertion>'
        for clause, description, message in failures or []
    )
    value = 'true' if compliant else 'false'
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<rawResults><validationResult flavour="PDFA_2_U" '
        f'isCompliant="{value}" jobEndStatus="NORMAL"><assertions>'
        f'{assertions}</assertions></validationResult></rawResults>'
    )


def write_fake_verapdf(
    root: Path,
    *,
    payload: str,
    exit_code: int,
    version: str = '1.30.4',
) -> Path:
    path = root / 'verapdf'
    path.write_text(
        f'#!{sys.executable}\n'
        'import sys\n'
        "if '--version' in sys.argv:\n"
        f"    print('veraPDF {version}')\n"
        '    sys.exit(0)\n'
        f'sys.stdout.write({payload!r})\n'
        f'sys.exit({exit_code})\n'
    )
    path.chmod(0o755)
    return path


class VerapdfParsingTests(unittest.TestCase):
    def test_json_report_exposes_compliance_and_failed_checks(self):
        report = parse_verapdf_json(
            verapdf_json(
                False,
                [
                    verapdf_rule('6.1.3', 'Missing or empty ID in the document trailer'),
                    verapdf_rule('6.6.2.1', 'The catalog has no metadata stream'),
                ],
            )
        )
        self.assertIs(report['compliant'], False)
        self.assertEqual(report['failed_check_count'], 2)
        self.assertEqual(len(report['failed_checks']), 2)
        self.assertEqual(report['failed_checks'][0]['clause'], '6.1.3')
        self.assertIn('Missing or empty ID', report['failed_checks'][0]['error'])
        message = describe_verapdf_failure(report)
        self.assertIn('clause 6.1.3', message)
        self.assertIn('Missing or empty ID', message)

    def test_compliant_json_report_has_no_failed_checks(self):
        report = parse_verapdf_json(verapdf_json(True))
        self.assertIs(report['compliant'], True)
        self.assertEqual(report['failed_checks'], [])
        self.assertIsNone(report['error'])

    def test_invalid_or_incomplete_json_reports_are_errors(self):
        for payload in (
            'not json',
            '[]',
            '{"report":{}}',
            '{"report":{"jobs":[{"jobEndStatus":"ERROR"}]}}',
            json.dumps(
                {
                    'report': {
                        'jobs': [
                            {
                                'validationResult': [
                                    {'jobEndStatus': 'ERROR', 'statement': 'failed to parse PDF'}
                                ]
                            }
                        ]
                    }
                }
            ),
        ):
            with self.subTest(payload=payload):
                report = parse_verapdf_json(payload)
                self.assertIsNone(report['compliant'])
                self.assertIsNotNone(report['error'])
                self.assertIn('report', report['error'])

    def test_raw_report_parses_is_compliant_and_failed_assertions(self):
        failing = parse_verapdf_raw(
            verapdf_raw(False, [('6.1.3', 'trailer rule', 'Missing or empty ID')])
        )
        self.assertIs(failing['compliant'], False)
        self.assertEqual(failing['failed_check_count'], 1)
        self.assertEqual(failing['failed_checks'][0]['clause'], '6.1.3')
        self.assertIn('Missing or empty ID', failing['failed_checks'][0]['error'])

        passing = parse_verapdf_raw(verapdf_raw(True))
        self.assertIs(passing['compliant'], True)
        self.assertIsNone(passing['error'])

    def test_malformed_raw_report_is_an_error(self):
        report = parse_verapdf_raw('<rawResults><validationResult isCompliant="')
        self.assertIsNone(report['compliant'])
        self.assertIn('invalid raw report', report['error'])
        missing = parse_verapdf_raw('<rawResults/>')
        self.assertIsNone(missing['compliant'])
        self.assertIn('no validationResult', missing['error'])

    def test_report_format_is_auto_detected(self):
        self.assertIs(parse_verapdf_report(verapdf_json(True))['compliant'], True)
        self.assertIs(parse_verapdf_report(verapdf_raw(True))['compliant'], True)

    def test_failure_description_covers_invalid_reports(self):
        message = describe_verapdf_failure(
            {'error': 'veraPDF could not run', 'failed_checks': []}
        )
        self.assertEqual(message, 'veraPDF could not run')

    def test_manifest_pdfa_declaration_contract(self):
        self.assertTrue(manifest_declares_pdfa_2u({'conformance': 'pdfa-2u'}, {}))
        self.assertTrue(manifest_declares_pdfa_2u({}, {'conformance': 'PDF/A-2u'}))
        self.assertTrue(manifest_declares_pdfa_2u({}, {'pdfa_2u': True}))
        self.assertFalse(manifest_declares_pdfa_2u({'conformance': 'standard'}, {}))
        self.assertFalse(manifest_declares_pdfa_2u({}, {}))

    def test_find_verapdf_accepts_option_env_and_rejects_missing_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            executable = write_fake_verapdf(root, payload='', exit_code=0)
            self.assertEqual(find_verapdf(executable), str(executable))
            with mock.patch.dict(os.environ, {'VERAPDF': str(executable)}):
                self.assertEqual(find_verapdf(), str(executable))
            self.assertIsNone(find_verapdf(root / 'missing'))
            with mock.patch.dict(os.environ, {'VERAPDF': str(root / 'missing')}):
                self.assertIsNone(find_verapdf())


class PdfaGateTests(unittest.TestCase):
    def run_case(
        self,
        *,
        payload: str,
        exit_code: int,
        declared: bool = True,
        version: str = '1.30.4',
        require: bool = True,
        extra_args: list[str] | None = None,
    ) -> tuple[int, dict[str, object], str, str, list[str]]:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            size = write_pdf(root / 'pdfs' / 'report.pdf')
            header = {'conformance': 'pdfa-2u'} if declared else None
            write_manifest(
                root / 'pdfs' / 'manifest.json',
                [ok_entry('report.docx', size)],
                header,
            )
            executable = write_fake_verapdf(
                root, payload=payload, exit_code=exit_code, version=version
            )
            report_path = root / 'report.json'
            report_dir = root / 'verapdf-reports'
            arguments = [
                str(root / 'pdfs' / 'manifest.json'),
                '--verapdf', str(executable),
                '--verapdf-version', '1.30.4',
                '--verapdf-report-dir', str(report_dir),
                '--json', str(report_path),
                *(extra_args or []),
            ]
            if require:
                arguments.append('--require-pdfa-2u')
            with mock.patch.dict(sys.modules, {'pypdfium2': fake_pdfium('hello')}):
                code, stdout, stderr = run_main(arguments)
            report = (
                json.loads(report_path.read_text())
                if report_path.is_file()
                else {}
            )
            reports = sorted(
                str(path.relative_to(root))
                for path in report_dir.rglob('*')
                if path.is_file()
            )
            return code, report, stdout, stderr, reports

    def test_non_compliant_report_fails_even_when_the_exit_status_is_zero(self):
        code, report, _stdout, _stderr, reports = self.run_case(
            payload=verapdf_json(
                False, [verapdf_rule('6.1.3', 'Missing or empty ID')]
            ),
            exit_code=0,
        )
        self.assertEqual(code, 1, report['failures'])
        checks = report['failures'][0]['checks']
        pdfa_checks = [message for check, message in checks if check == 'pdfa']
        self.assertEqual(len(pdfa_checks), 1, checks)
        self.assertIn('clause 6.1.3', pdfa_checks[0])
        self.assertEqual(report['fixtures'][0]['verapdf']['compliant'], False)
        self.assertEqual(report['summary']['pdfa_checked'], 1)
        self.assertEqual(report['summary']['pdfa_compliant'], 0)
        self.assertEqual(reports, ['verapdf-reports/report.docx.verapdf.json'])

    def test_compliant_report_passes_even_when_the_exit_status_is_nonzero(self):
        code, report, _stdout, _stderr, _reports = self.run_case(
            payload=verapdf_json(True), exit_code=1
        )
        self.assertEqual(code, 0, report['failures'])
        self.assertEqual(report['fixtures'][0]['verapdf']['compliant'], True)
        self.assertEqual(report['summary']['pdfa_compliant'], 1)

    def test_invalid_report_is_a_verapdf_failure(self):
        code, report, _stdout, _stderr, _reports = self.run_case(
            payload='{"report":{"jobs":[{"jobEndStatus":"ERROR"}]}}',
            exit_code=0,
        )
        self.assertEqual(code, 1)
        checks = report['failures'][0]['checks']
        self.assertIn('verapdf', [check for check, _ in checks])
        message = next(message for check, message in checks if check == 'verapdf')
        self.assertIn('validationResult', message)

    def test_manifest_must_declare_pdfa_2u(self):
        code, report, _stdout, _stderr, _reports = self.run_case(
            payload=verapdf_json(True), exit_code=0, declared=False
        )
        self.assertEqual(code, 1)
        checks = report['failures'][0]['checks']
        message = next(message for check, message in checks if check == 'pdfa')
        self.assertIn('does not declare', message)

    def test_version_mismatch_is_rejected_before_validation(self):
        code, _report, _stdout, stderr, _reports = self.run_case(
            payload=verapdf_json(True), exit_code=0, version='1.29.1'
        )
        self.assertEqual(code, 2)
        self.assertIn('version mismatch', stderr)
        self.assertIn('1.30.4', stderr)

    def test_missing_verapdf_is_a_missing_dependency(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            size = write_pdf(root / 'pdfs' / 'report.pdf')
            write_manifest(
                root / 'pdfs' / 'manifest.json',
                [ok_entry('report.docx', size)],
                {'conformance': 'pdfa-2u'},
            )
            with mock.patch.dict(sys.modules, {'pypdfium2': fake_pdfium('hello')}):
                code, _stdout, stderr = run_main([
                    str(root / 'pdfs' / 'manifest.json'),
                    '--require-pdfa-2u',
                    '--verapdf', str(root / 'missing-verapdf'),
                ])
            self.assertEqual(code, 2)
            self.assertIn('needs veraPDF', stderr)


    def test_allow_unsupported_only_skips_reviewed_fixtures(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            size = write_pdf(root / 'pdfs' / 'report.pdf')
            reason = f'export: unsupported PDF content: {UNSUPPORTED_MISSING_GLYPH}'
            write_manifest(
                root / 'pdfs' / 'manifest.json',
                [
                    ok_entry('report.docx', size),
                    {
                        'path': 'cjk-rtl-emoji/generated-emoji.docx',
                        'status': 'unsupported',
                        'error': reason,
                    },
                ],
                {'conformance': 'pdfa-2u'},
            )
            report_path = root / 'report.json'
            with mock.patch.dict(sys.modules, {'pypdfium2': fake_pdfium('hello')}):
                code, stdout, stderr = run_main([
                    str(root / 'pdfs' / 'manifest.json'),
                    '--allow-unsupported',
                    '--json', str(report_path),
                ])
            report = json.loads(report_path.read_text())

            self.assertEqual(code, 0, (stdout, stderr))
            self.assertEqual(report['summary']['fixtures'], 1)
            self.assertEqual(report['summary']['unsupported'], 1)
            self.assertEqual(
                report['unsupported'][0]['path'],
                'cjk-rtl-emoji/generated-emoji.docx',
            )

    def test_unreviewed_unsupported_entry_still_fails(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reason = f'export: unsupported PDF content: {UNSUPPORTED_MISSING_GLYPH}'
            write_manifest(
                root / 'pdfs' / 'manifest.json',
                [{
                    'path': 'report.docx',
                    'status': 'unsupported',
                    'error': reason,
                }],
                {'conformance': 'pdfa-2u'},
            )
            report_path = root / 'report.json'
            with mock.patch.dict(sys.modules, {'pypdfium2': fake_pdfium('hello')}):
                code, stdout, stderr = run_main([
                    str(root / 'pdfs' / 'manifest.json'),
                    '--allow-unsupported',
                    '--json', str(report_path),
                ])
            report = json.loads(report_path.read_text())

            self.assertEqual(code, 1, (stdout, stderr))
            self.assertEqual(report['summary']['unsupported'], 0)
            checks = report['failures'][0]['checks']
            self.assertIn('export', [check for check, _ in checks])
            self.assertIn('pdf', [check for check, _ in checks])

    def test_unsupported_allowlist_requires_the_missing_glyph_reason(self):
        entry = {
            'status': 'unsupported',
            'error': f'export: unsupported PDF content: {UNSUPPORTED_MISSING_GLYPH}',
        }
        for path in ALLOWED_UNSUPPORTED_FIXTURES:
            self.assertTrue(is_allowed_unsupported(Path(path), entry))
        self.assertFalse(is_allowed_unsupported(Path('report.docx'), entry))
        self.assertFalse(
            is_allowed_unsupported(
                Path('page-fragments/pages.docx'),
                {'error': 'export: unsupported PDF content: PDF/A-2u text cannot address .notdef'},
            )
        )

    def test_without_the_flag_validation_is_skipped(self):
        code, report, _stdout, _stderr, reports = self.run_case(
            payload=verapdf_json(False), exit_code=1, require=False
        )
        self.assertEqual(code, 0, report['failures'])
        self.assertEqual(reports, [])
        self.assertIsNone(report['fixtures'][0]['verapdf'])
        self.assertFalse(report['pdfa']['required'])
        self.assertEqual(report['summary']['pdfa_checked'], 0)


if __name__ == '__main__':
    unittest.main()
