import contextlib
import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from pdf_matrix import (
    GENERATED_TEXT_FIXTURES,
    DocxInfo,
    group_for,
    load_manifest,
    main,
    make_large_docx,
    is_emoji_char,
    make_text_docx,
    normalize_text,
    parse_page_size,
    parse_pdffonts,
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


def write_manifest(path: Path, entries: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({'fixtures': entries}))


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


if __name__ == '__main__':
    unittest.main()
