#!/usr/bin/env python3
"""Verify every PDF recorded by `export_pdf_batch`'s manifest.

The gate is structural, not a Word-fidelity comparison. For each exported PDF
it asserts that pypdfium2 can open it, that it has at least one page, that page
sizes match the DOCX section sizes within 0.01 pt, and that text is extractable.
When the tools are installed it also runs `qpdf --check`, `pdffonts`, and
`pdftotext`.

Fixtures are grouped by their first path segment when they came from
`--prepare`; otherwise the group is inferred from the fixture directory, name,
and document content. `--prepare` also generates small CJK, RTL, and emoji
DOCX files and an `expected.json`; the CJK and RTL fixtures must recover their
exact text (whitespace and bidi layout artifacts ignored), and every
`pdffonts` row must be embedded, subset, and Unicode-mapped. Missing emoji
glyphs are reported as `emoji_glyph_coverage` warnings rather than failures.
The script writes compact JSON and a human summary, and exits non-zero when
any fixture fails.

Examples:

    # Stage a representative fixture set plus a generated large document.
    python3 scripts/office-quality/pdf_matrix.py --prepare .source/pdf-matrix/fixtures

    # Export every staged fixture.
    cargo +1.99.1 run -p betteroffice-docx --example export_pdf_batch -- \
      .source/pdf-matrix/fixtures .source/pdf-matrix/pdfs

    # Check the manifest and every PDF it names.
    python3 scripts/office-quality/pdf_matrix.py .source/pdf-matrix/pdfs/manifest.json

Only pypdfium2 is required. Missing `qpdf`, `pdffonts`, or `pdftotext` are
recorded as unavailable rather than treated as failures.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

REPO_ROOT = Path(__file__).resolve().parents[2]

GROUPS = (
    'table-merged',
    'header-footer',
    'footnotes',
    'pagination',
    'images',
    'cjk-rtl-emoji',
    'large',
    'other',
)

# Directory names seen in either the staged set or the repository fixture tree.
GROUP_ALIASES = {
    'table-merged': 'table-merged',
    'tables': 'table-merged',
    'table': 'table-merged',
    'merged-row-heights': 'table-merged',
    'table-widths': 'table-merged',
    'table-outer-borders': 'table-merged',
    'table-row-widows': 'table-merged',
    'table-row-min-height': 'table-merged',
    'header-footer': 'header-footer',
    'header-footer-reservation': 'header-footer',
    'header-float-wrap': 'header-footer',
    'footnotes': 'footnotes',
    'pagination': 'pagination',
    'keep-next': 'pagination',
    'widow-control': 'pagination',
    'page-fragments': 'pagination',
    'images': 'images',
    'image-line-spacing': 'images',
    'cjk-rtl-emoji': 'cjk-rtl-emoji',
    'cjk': 'cjk-rtl-emoji',
    'rtl': 'cjk-rtl-emoji',
    'emoji': 'cjk-rtl-emoji',
    'large': 'large',
    'large-doc': 'large',
}

# A representative subset of the checked-in DOCX fixtures. Staged paths keep
# their original fixture-relative directory so a manifest names the source.
PREPARE_FIXTURES = (
    ('table-merged', 'crates/docx-edit/tests/fixtures/merged-row-heights/same-span.docx'),
    ('table-merged', 'crates/docx-edit/tests/fixtures/merged-row-heights/repeated-header.docx'),
    ('table-merged', 'crates/docx-edit/tests/fixtures/table-widths/autofit-span.docx'),
    ('table-merged', 'crates/docx-edit/tests/fixtures/table-widths/fixed-span.docx'),
    ('table-merged', 'crates/docx-edit/tests/fixtures/table-outer-borders/tbl-border4-3col.docx'),
    ('table-merged', 'crates/docx-edit/tests/fixtures/table-row-widows/row-four-line-widow.docx'),
    ('header-footer', 'crates/docx-edit/tests/fixtures/header-footer-reservation/even-and-odd-headers.docx'),
    ('header-footer', 'crates/docx-edit/tests/fixtures/header-footer-reservation/title-page-header.docx'),
    ('header-footer', 'crates/docx-edit/tests/fixtures/header-footer-reservation/continuous-title-page-at-page-break.docx'),
    ('header-footer', 'crates/docx-edit/tests/fixtures/header-footer-reservation/even-and-odd-missing-even.docx'),
    ('header-footer', 'crates/docx-edit/tests/fixtures/header-float-wrap/header-float-top-and-bottom.docx'),
    ('header-footer', 'crates/docx-edit/tests/fixtures/header-float-wrap/footer-float-top-and-bottom.docx'),
    ('footnotes', 'crates/docx-edit/tests/fixtures/footnote-anchor.docx'),
    ('footnotes', 'crates/docx-edit/tests/fixtures/layout-request/request-edges.docx'),
    ('pagination', 'crates/docx-edit/tests/fixtures/keep-next/keep-next-page-break-before.docx'),
    ('pagination', 'crates/docx-edit/tests/fixtures/keep-next/keep-next-table-row-cant-split-moves.docx'),
    ('pagination', 'crates/docx-edit/tests/fixtures/widow-control/heading-five-lines-room-four.docx'),
    ('images', 'crates/docx-edit/tests/fixtures/image-line-spacing/img-body-240.docx'),
    ('images', 'crates/docx-edit/tests/fixtures/image-line-spacing/img-body-360.docx'),
    ('images', 'crates/docx-edit/tests/fixtures/image-line-spacing/img-cell-300.docx'),
)

# Generated fixtures. Their text must survive PDF export exactly: the checker
# asserts every expected character is present in extracted text and (for the
# `strict_fonts` entries) that every pdffonts row is embedded, subset and
# ToUnicode-mapped.
#
# Each tuple is (file name, paragraphs, expected substrings). A paragraph is
# a tuple of runs, and a run is (text, east-asian language, bidi language);
# a bidi language marks the paragraph and its runs right-to-left.
GENERATED_TEXT_FIXTURES = (
    (
        'generated-cjk.docx',
        ((('中文测试', 'zh-CN', None), (' 日本語のテスト', 'ja-JP', None)),),
        ('中文测试', '日本語のテスト'),
    ),
    (
        'generated-rtl.docx',
        ((('مرحبا بالعالم', None, 'ar-SA'), (' שלום עולם', None, 'he-IL')),),
        ('مرحبا بالعالم', 'שלום עולם'),
    ),
    (
        'generated-emoji.docx',
        ((('😀🎉🚀', None, None),),),
        ('😀', '🎉', '🚀'),
    ),
)
GENERATED_GROUP = 'cjk-rtl-emoji'
EXPECTATIONS_FILE = 'expected.json'

LARGE_SOURCE = 'crates/betteroffice-docx/tests/corpus/fixtures/betteroffice-demo.docx'
LARGE_REPEATS = 30

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TOLERANCE_PT = 0.01


class MissingDependency(RuntimeError):
    pass


@dataclass
class DocxInfo:
    """What the matrix needs to know about a source DOCX."""

    page_sizes: set[tuple[float, float]] = field(default_factory=set)
    scripts: set[str] = field(default_factory=set)
    has_notes: bool = False
    source_text: str = ''
    error: str | None = None


def scripts_in(text: str) -> set[str]:
    found = set()
    for character in text:
        code = ord(character)
        if 0x0590 <= code <= 0x08FF:
            found.add('rtl')
        elif (
            0x1100 <= code <= 0x11FF
            or 0x3040 <= code <= 0x30FF
            or 0x3130 <= code <= 0x318F
            or 0x3400 <= code <= 0x4DBF
            or 0x4E00 <= code <= 0x9FFF
            or 0xAC00 <= code <= 0xD7AF
            or 0xF900 <= code <= 0xFAFF
            or 0x20000 <= code <= 0x2FA1F
        ):
            found.add('cjk')
        elif (
            0x1F000 <= code <= 0x1FAFF
            or 0x2600 <= code <= 0x27BF
            or 0x2B00 <= code <= 0x2BFF
            or 0x1F1E6 <= code <= 0x1F1FF
        ):
            found.add('emoji')
    return found


BIDI_CONTROLS = frozenset(
    '\u061c\u200b\u200c\u200d\u200e\u200f\u202a\u202b\u202c\u202d\u202e\u2066\u2067\u2068\u2069\ufeff'
)


def is_emoji_char(character: str) -> bool:
    code = ord(character)
    return (
        0x1F000 <= code <= 0x1FAFF
        or 0x2600 <= code <= 0x27BF
        or 0x2B00 <= code <= 0x2BFF
        or 0x1F1E6 <= code <= 0x1F1FF
    )


def normalize_text(value: str) -> str:
    """Remove whitespace and bidi controls before comparing extracted text.

    Poppler can insert a space inside a shaped Arabic word and wraps runs in
    bidi control characters. The visual PDF is still correct, so comparisons
    remove those layout artifacts while keeping every letter.
    """
    without_controls = ''.join(
        character for character in value if character not in BIDI_CONTROLS
    )
    return re.sub(r'\s+', '', without_controls)


def read_docx_info(path: Path) -> DocxInfo:
    info = DocxInfo()
    try:
        with zipfile.ZipFile(path) as archive:
            if 'word/document.xml' not in archive.namelist():
                info.error = 'package has no word/document.xml'
                return info
            raw = archive.read('word/document.xml')
    except (OSError, zipfile.BadZipFile) as error:
        info.error = f'cannot read DOCX: {error}'
        return info

    text = raw.decode('utf-8', 'ignore')
    info.source_text = html.unescape(re.sub(r'<[^>]+>', ' ', text))
    info.scripts = scripts_in(info.source_text)
    info.has_notes = bool(
        re.search(r'footnote(?:Reference|Ref)|endnote(?:Reference|Ref)', text)
    )
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as error:
        info.error = f'word/document.xml is not well-formed: {error}'
        return info
    for element in root.iter(f'{W_NS}pgSz'):
        width = element.get(f'{W_NS}w')
        height = element.get(f'{W_NS}h')
        if not width or not height:
            continue
        try:
            info.page_sizes.add((int(width) / 20.0, int(height) / 20.0))
        except ValueError:
            continue
    return info


def group_for(relative: Path, info: DocxInfo | None, large_bytes: int) -> str:
    parts = relative.parts
    # A staged path starts with the canonical group directory and wins over
    # content inference. Repository fixture directories are only aliases, so
    # content can still move e.g. an emoji page-fragments fixture to its
    # CJK/RTL/emoji group.
    if parts and parts[0] in GROUPS:
        return parts[0]
    aliases = {GROUP_ALIASES[part] for part in parts[:-1] if part in GROUP_ALIASES}
    name = relative.name.lower()
    if 'large' in name or any('large' in part.lower() for part in parts[:-1]):
        return 'large'
    if info is not None and info.scripts:
        return 'cjk-rtl-emoji'
    if info is not None and info.has_notes:
        return 'footnotes'
    if 'footnote' in name or 'endnote' in name:
        return 'footnotes'
    if 'image' in name or 'img' in parts:
        return 'images'
    if aliases:
        return sorted(aliases)[0]
    if 'header' in name or 'footer' in name:
        return 'header-footer'
    if 'table' in name or 'merged' in name:
        return 'table-merged'
    try:
        if relative.stat().st_size >= large_bytes:
            return 'large'
    except OSError:
        pass
    return 'other'


def load_manifest(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    try:
        data = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError(f'{path}: cannot read manifest: {error}') from error

    if isinstance(data, list):
        entries, header = data, {}
    elif isinstance(data, dict):
        entries = data.get('fixtures', data.get('results'))
        header = {key: value for key, value in data.items() if key not in {'fixtures', 'results'}}
        if entries is None and 'path' in data:
            entries, header = [data], {}
    else:
        raise ValueError(f'{path}: manifest must be an object or an array')
    if not isinstance(entries, list) or not all(isinstance(entry, dict) for entry in entries):
        raise ValueError(f'{path}: manifest fixtures must be an array of objects')
    for entry in entries:
        if not isinstance(entry.get('path'), str):
            raise ValueError(f'{path}: every fixture needs a string path')
    return header, entries


def parse_page_size(value: str) -> tuple[float, float]:
    match = re.fullmatch(r'\s*([0-9.]+)\s*[xX]\s*([0-9.]+)\s*', value)
    if not match:
        raise argparse.ArgumentTypeError('expected WIDTHxHEIGHT in points')
    width, height = float(match.group(1)), float(match.group(2))
    if width <= 0 or height <= 0:
        raise argparse.ArgumentTypeError('page size must be positive')
    return width, height


def tool_path(name: str) -> str | None:
    return shutil.which(name)


def run_tool(command: list[str]) -> subprocess.CompletedProcess[str] | None:
    try:
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None


def parse_pdffonts(output: str) -> list[dict[str, str]]:
    lines = output.splitlines()
    header_index = None
    for index, line in enumerate(lines):
        if line.lstrip().startswith('name') and 'emb' in line and 'uni' in line:
            header_index = index
            header = line
            break
    if header_index is None:
        return []

    def column(name: str, default: int) -> int:
        position = header.find(name)
        return position if position != -1 else default

    name_end = column('type', 37)
    type_end = column('encoding', 55)
    encoding_end = column('emb', 73)
    emb_end = column('sub', 77)
    sub_end = column('uni', 81)
    uni_end = column('object', 89)

    fonts = []
    for line in lines[header_index + 1:]:
        if not line.strip() or set(line.strip()) <= {'-', ' '}:
            continue

        def field(start: int, end: int) -> str:
            return line[start:end].strip()

        name = field(0, name_end)
        if not name:
            continue
        fonts.append(
            {
                'name': name,
                'type': field(name_end, type_end),
                'encoding': field(type_end, encoding_end),
                'emb': field(encoding_end, emb_end).lower(),
                'sub': field(emb_end, sub_end).lower(),
                'uni': field(sub_end, uni_end).lower(),
            }
        )
    return fonts


def prepare(destination: Path) -> int:
    destination = destination.resolve()
    if destination.is_dir() and any(destination.iterdir()):
        print(
            f'pdf_matrix: warning: {destination} is not empty; stale fixtures may remain',
            file=sys.stderr,
        )
    missing = []
    staged = 0
    for group, relative in PREPARE_FIXTURES:
        source = REPO_ROOT / relative
        if not source.is_file():
            missing.append(relative)
            continue
        fixture_relative = Path(relative)
        for marker in ('/tests/fixtures/', '/tests/corpus/fixtures/'):
            if marker in relative:
                fixture_relative = Path(relative.split(marker, 1)[1])
                break
        target = destination / group / fixture_relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        staged += 1

    expectations = {}
    for name, paragraphs, expected in GENERATED_TEXT_FIXTURES:
        target = destination / GENERATED_GROUP / name
        make_text_docx(target, paragraphs)
        expectations[f'{GENERATED_GROUP}/{name}'] = {
            'text': list(expected),
            'strict_fonts': True,
        }
        staged += 1
    (destination / EXPECTATIONS_FILE).write_text(
        json.dumps(expectations, ensure_ascii=False, indent=2) + '\n'
    )

    large_source = REPO_ROOT / LARGE_SOURCE
    if large_source.is_file():
        make_large_docx(large_source, destination / 'large' / 'large-demo.docx', LARGE_REPEATS)
        staged += 1
    else:
        missing.append(LARGE_SOURCE)

    if missing:
        print('pdf_matrix: missing fixtures:', file=sys.stderr)
        for relative in missing:
            print(f'  {relative}', file=sys.stderr)
        return 2
    print(f'staged {staged} fixtures under {destination}')
    print(
        'next: cargo +1.99.1 run -p betteroffice-docx --example export_pdf_batch -- '
        f'{destination} {destination.parent / "pdfs"}'
    )
    return 0


def make_text_docx(
    target: Path,
    paragraphs: tuple[tuple[tuple[str, str | None, str | None], ...], ...],
) -> None:
    """Write a minimal one-page WordprocessingML package.

    A paragraph is a tuple of runs. A run is `(text, east_asian_language,
    bidi_language)`. Language properties make the script explicit so font
    fallback is not guessed from the text alone. A bidi language marks the
    paragraph and that run right-to-left.
    """
    body = []
    for runs in paragraphs:
        rtl = any(bidi is not None for _, _, bidi in runs)
        ppr = '<w:pPr><w:bidi/><w:spacing w:after="120"/></w:pPr>' if rtl else (
            '<w:pPr><w:spacing w:after="120"/></w:pPr>'
        )
        run_xml = []
        for text, east_asian, bidi in runs:
            rpr = []
            if east_asian is not None:
                rpr.append(f'<w:lang w:eastAsia="{xml_escape(east_asian)}"/>')
            if bidi is not None:
                rpr.append(f'<w:rtl/><w:lang w:bidi="{xml_escape(bidi)}"/>')
            run_properties = f'<w:rPr>{"".join(rpr)}</w:rPr>' if rpr else ''
            run_xml.append(
                f'<w:r>{run_properties}'
                f'<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'
            )
        body.append(f'<w:p>{ppr}{"".join(run_xml)}</w:p>')

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{"".join(body)}'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        '</w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    relationships = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>'
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', content_types)
        archive.writestr('_rels/.rels', relationships)
        archive.writestr('word/document.xml', document)


def make_large_docx(source: Path, target: Path, repeats: int) -> None:
    """Repeat the demo body so the large-doc group exercises real pagination."""
    with zipfile.ZipFile(source) as archive:
        entries = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    xml = entries['word/document.xml']
    body_start = xml.index(b'<w:body>') + len(b'<w:body>')
    body_end = xml.index(b'</w:body>', body_start)
    section_start = xml.rfind(b'<w:sectPr', body_start, body_end)
    split = section_start if section_start != -1 else body_end
    repeated = xml[body_start:split] * repeats
    entries['word/document.xml'] = (
        xml[:body_start] + repeated + xml[split:body_end] + xml[body_end:]
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('manifests', nargs='*', type=Path, help='manifest.json paths')
    parser.add_argument('--manifest', dest='extra_manifests', action='append', type=Path, default=[])
    parser.add_argument('--pdf-dir', type=Path, help='PDF root when it is not the manifest directory')
    parser.add_argument('--json', type=Path, help='write the compact report here')
    parser.add_argument('--out', type=Path, help=argparse.SUPPRESS)
    parser.add_argument('--fixtures-dir', type=Path, help='source DOCX root when the manifest omits input_dir')
    parser.add_argument('--page-size', type=parse_page_size, help='WIDTHxHEIGHT expected for every page')
    parser.add_argument('--large-bytes', type=int, default=512 * 1024, help='size threshold for the large group')
    parser.add_argument('--strict-fonts', action='store_true',
                        help='fail unless every pdffonts row is embedded, subset and Unicode-mapped')
    parser.add_argument('--verbose', action='store_true')
    parser.add_argument('--prepare', type=Path, metavar='DIR', help='stage the curated fixture set and exit')
    args = parser.parse_args(argv)

    if args.prepare:
        return prepare(args.prepare)

    manifests = [*args.manifests, *args.extra_manifests]
    if not manifests:
        parser.error('one or more manifest.json paths are required')
    manifests = [path.resolve() for path in manifests]

    try:
        import pypdfium2 as pdfium
    except ImportError:
        print(
            'pdf_matrix: missing dependency pypdfium2.\n'
            'Install it in the harness virtualenv:\n'
            '  python3 -m venv .source/office-quality/venv\n'
            '  .source/office-quality/venv/bin/pip install pypdfium2\n'
            '(or install scripts/office-quality/requirements.txt for the full harness)',
            file=sys.stderr,
        )
        return 2

    try:
        from importlib.metadata import version as package_version

        pdfium_version = package_version('pypdfium2')
    except Exception:
        pdfium_version = getattr(pdfium, '__version__', 'present')

    tools = {
        'pypdfium2': pdfium_version,
        'qpdf': tool_path('qpdf'),
        'pdffonts': tool_path('pdffonts'),
        'pdftotext': tool_path('pdftotext'),
    }

    report_path = args.json or args.out or (manifests[0].parent / 'pdf_matrix.json')
    report: dict[str, Any] = {
        'schema': 1,
        'manifests': [str(path) for path in manifests],
        'pdf_dir': str(args.pdf_dir.resolve()) if args.pdf_dir else None,
        'tools': tools,
        'summary': {},
        'groups': {group: {'fixtures': 0, 'failed': 0, 'pages': 0, 'pdf_bytes': 0,
                           'text_chars': 0, 'extractable_chars': 0} for group in GROUPS},
        'fixtures': [],
        'failures': [],
        'warnings': [],
    }

    seen: dict[str, int] = {}
    for manifest_path in manifests:
        try:
            header, entries = load_manifest(manifest_path)
        except ValueError as error:
            print(f'pdf_matrix: {error}', file=sys.stderr)
            return 2
        source_root = None
        if header.get('input_dir'):
            source_root = Path(header['input_dir'])
        elif args.fixtures_dir:
            source_root = args.fixtures_dir
        expectations = {}
        if source_root is not None:
            expectations_path = source_root / EXPECTATIONS_FILE
            if expectations_path.is_file():
                try:
                    expectations = json.loads(expectations_path.read_text())
                except (OSError, json.JSONDecodeError) as error:
                    report['warnings'].append(f'{expectations_path}: {error}')
        pdf_root = args.pdf_dir or manifest_path.parent
        for entry in entries:
            relative = Path(entry['path'])
            key = f"{pdf_root}\0{relative}"
            if key in seen:
                report['warnings'].append(
                    f'duplicate fixture path across manifests: {relative} (first in {seen[key]})'
                )
                continue
            seen[key] = len(report['fixtures'])
            fixture = check_fixture(
                pdfium,
                entry,
                relative,
                pdf_root,
                source_root,
                tools,
                args.page_size,
                args.large_bytes,
                expectations.get(str(relative).replace('\\', '/')),
                args.strict_fonts,
                report['warnings'],
            )
            group = fixture['group']
            bucket = report['groups'][group]
            bucket['fixtures'] += 1
            bucket['pages'] += fixture['pages'] or 0
            bucket['pdf_bytes'] += fixture['pdf_bytes'] or 0
            bucket['text_chars'] += fixture['text_chars'] or 0
            bucket['extractable_chars'] += fixture['extractable_chars'] or 0
            if fixture['checks']:
                bucket['failed'] += 1
                report['failures'].append(
                    {
                        'path': str(relative),
                        'groups': [group],
                        'checks': fixture['checks'],
                    }
                )
            report['fixtures'].append(fixture)
            if args.verbose:
                state = 'FAIL' if fixture['checks'] else 'ok'
                detail = '; '.join(f"{check}: {message}" for check, message in fixture['checks'])
                print(f'[{state}] {group}/{relative}: {detail}')

    fixtures = report['fixtures']
    if not fixtures:
        print('pdf_matrix: the manifests contain no fixtures', file=sys.stderr)
        return 2
    pages = sum(fixture['pages'] or 0 for fixture in fixtures)
    report['summary'] = {
        'fixtures': len(fixtures),
        'passed': sum(1 for fixture in fixtures if not fixture['checks']),
        'failed': len(report['failures']),
        'pages': pages,
        'pdf_bytes': sum(fixture['pdf_bytes'] or 0 for fixture in fixtures),
        'text_chars': sum(fixture['text_chars'] or 0 for fixture in fixtures),
        'extractable_chars': sum(fixture['extractable_chars'] or 0 for fixture in fixtures),
    }

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, separators=(',', ':'), sort_keys=True) + '\n')

    print(
        f"PDF matrix: {report['summary']['fixtures']} fixture(s), "
        f"{report['summary']['passed']} passed, {report['summary']['failed']} failed, "
        f"{report['summary']['pages']} page(s)"
    )
    print('Groups:')
    for group in GROUPS:
        bucket = report['groups'][group]
        if bucket['fixtures'] == 0:
            continue
        print(
            f"  {group:<16} {bucket['fixtures'] - bucket['failed']:>3}/{bucket['fixtures']:<3} ok  "
            f"pages={bucket['pages']:<4} text={bucket['extractable_chars']}"
        )
    if report['failures']:
        print('Failures:')
        for failure in report['failures']:
            checks = '; '.join(f'{check}: {message}' for check, message in failure['checks'])
            print(f"  - {failure['path']}: {checks}")
    unavailable = [name for name, path in tools.items() if path is None]
    if unavailable:
        print(f"Unavailable optional tools: {', '.join(unavailable)}")
    if report['warnings']:
        print(f"Warnings: {len(report['warnings'])} (see {report_path})")
    print(f'Report: {report_path}')
    return 1 if report['failures'] else 0


def check_fixture(
    pdfium: Any,
    entry: dict[str, Any],
    relative: Path,
    pdf_root: Path,
    source_root: Path | None,
    tools: dict[str, str | None],
    page_size_override: tuple[float, float] | None,
    large_bytes: int,
    expectation: dict[str, Any] | None,
    strict_fonts: bool,
    warnings: list[str],
) -> dict[str, Any]:
    checks: list[tuple[str, str]] = []
    manifest_status = entry.get('status')
    if manifest_status != 'ok':
        checks.append(('export', str(entry.get('error') or f'manifest status is {manifest_status!r}')))

    source = (source_root / relative) if source_root else None
    info = read_docx_info(source) if source and source.is_file() else None
    if source is not None and not source.is_file():
        warnings.append(f'{relative}: source DOCX not found under {source_root}; content checks skipped')
    if info is not None and info.error:
        warnings.append(f'{relative}: {info.error}')
    group = group_for(relative, info, large_bytes)

    pdf_path = (pdf_root / relative).with_suffix('.pdf')
    fixture: dict[str, Any] = {
        'path': str(relative),
        'group': group,
        'status': manifest_status,
        'pages': entry.get('pages'),
        'pdf_bytes': entry.get('pdf_bytes'),
        'fonts_embedded': entry.get('fonts_embedded'),
        'text_chars': entry.get('text_chars'),
        'extractable_chars': 0,
        'page_sizes': [],
        'fonts': [],
        'unembedded_fonts': [],
        'emoji_glyph_coverage': None,
        'checks': checks,
    }

    try:
        raw = pdf_path.read_bytes()
    except OSError as error:
        checks.append(('pdf', f'cannot read {pdf_path}: {error}'))
        return fixture

    fixture['pdf_bytes'] = len(raw)
    if not raw.startswith(b'%PDF-'):
        checks.append(('pdf', 'missing %PDF- header'))
    if b'%%EOF' not in raw[-2048:]:
        checks.append(('pdf', 'missing %%EOF trailer'))
    if isinstance(entry.get('pdf_bytes'), int) and entry['pdf_bytes'] != len(raw):
        checks.append(
            ('pdf_bytes', f"manifest says {entry['pdf_bytes']} bytes, file has {len(raw)}")
        )

    try:
        document = pdfium.PdfDocument(str(pdf_path))
    except Exception as error:  # pypdfium2 raises its own error type.
        checks.append(('pdfium', f'cannot open: {error}'))
        return fixture

    page_sizes: list[tuple[float, float]] = []
    texts: list[str] = []
    try:
        page_count = len(document)
        if page_count <= 0:
            checks.append(('pages', 'PDF has no pages'))
        if isinstance(entry.get('pages'), int) and entry['pages'] != page_count:
            checks.append(('pages', f"manifest says {entry['pages']}, PDF has {page_count}"))
        for index in range(page_count):
            page = document[index]
            try:
                width, height = page.get_width(), page.get_height()
                page_sizes.append((width, height))
                if width <= 0 or height <= 0:
                    checks.append(('page_size', f'page {index + 1} is {width}x{height} pt'))
                textpage = page.get_textpage()
                try:
                    texts.append(textpage.get_text_range())
                finally:
                    textpage.close()
            finally:
                page.close()
    finally:
        document.close()

    fixture['pages'] = len(page_sizes)
    fixture['page_sizes'] = [[round(width, 3), round(height, 3)] for width, height in page_sizes]
    text = '\n'.join(texts)
    if '\ufffd' in text or '(cid:' in text:
        checks.append(('text', 'extracted text contains a replacement glyph or (cid:)'))

    expected_sizes = {page_size_override} if page_size_override else (
        info.page_sizes if info is not None else set()
    )
    if expected_sizes:
        for index, (width, height) in enumerate(page_sizes):
            if not any(
                abs(width - expected[0]) <= TOLERANCE_PT
                and abs(height - expected[1]) <= TOLERANCE_PT
                for expected in expected_sizes
            ):
                expected = ', '.join(f'{w:.2f}x{h:.2f}' for w, h in sorted(expected_sizes))
                checks.append(
                    (
                        'page_size',
                        f'page {index + 1} is {width:.2f}x{height:.2f} pt, expected {expected}',
                    )
                )
    elif page_sizes:
        first = page_sizes[0]
        for index, (width, height) in enumerate(page_sizes[1:], start=2):
            if (
                abs(width - first[0]) > TOLERANCE_PT
                or abs(height - first[1]) > TOLERANCE_PT
            ):
                checks.append(
                    (
                        'page_size',
                        f'page {index} is {width:.2f}x{height:.2f} pt, '
                        f'page 1 is {first[0]:.2f}x{first[1]:.2f}',
                    )
                )

    text_chars = entry.get('text_chars')
    pdftotext_output: str | None = None
    if group == 'large' and len(page_sizes) < 10:
        warnings.append(f'{relative}: large fixture has only {len(page_sizes)} page(s)')

    if tools['qpdf']:
        result = run_tool([tools['qpdf'], '--check', str(pdf_path)])
        if result is None:
            checks.append(('qpdf', 'qpdf --check could not run'))
        else:
            if result.returncode != 0:
                detail = (result.stderr or result.stdout).strip().splitlines()
                checks.append(
                    ('qpdf', f'exit {result.returncode}: '
                             + (detail[-1] if detail else 'no detail'))
                )
            if result.stderr.strip():
                checks.append(('qpdf', f'stderr: {result.stderr.strip().splitlines()[-1]}'))
            qpdf_warnings = [
                line for line in result.stdout.splitlines()
                if re.match(r'\s*(?:WARNING|ERROR)\b', line)
            ]
            if qpdf_warnings:
                checks.append(('qpdf', '; '.join(qpdf_warnings[:3])))
    else:
        warnings.append(f'{relative}: qpdf unavailable')

    if tools['pdffonts']:
        result = run_tool([tools['pdffonts'], str(pdf_path)])
        if result is None:
            checks.append(('pdffonts', 'pdffonts could not run'))
        else:
            if result.returncode != 0:
                checks.append(('pdffonts', f'exit {result.returncode}'))
            if result.stderr.strip():
                checks.append(
                    ('pdffonts', f'stderr: {result.stderr.strip().splitlines()[-1]}')
                )
            font_warnings = [
                line for line in result.stdout.splitlines()
                if re.match(r'\s*(?:WARNING|ERROR|Syntax Error)\b', line)
            ]
            if font_warnings:
                checks.append(('pdffonts', '; '.join(font_warnings[:3])))
            fonts = parse_pdffonts(result.stdout)
            fixture['fonts'] = fonts
            unembedded = [font for font in fonts if font['emb'] != 'yes']
            fixture['unembedded_fonts'] = [font['name'] for font in unembedded]
            expects_text = isinstance(text_chars, int) and text_chars > 0
            strict = strict_fonts or bool(expectation and expectation.get('strict_fonts'))
            if not fonts and expects_text:
                checks.append(('fonts', 'text document has no fonts'))
            elif fonts and not any(font['emb'] == 'yes' for font in fonts) and expects_text:
                checks.append(
                    ('fonts', 'text document has no embedded font: '
                              + ', '.join(font['name'] for font in fonts))
                )
            if strict:
                incomplete = [
                    '{} (emb={}, sub={}, uni={})'.format(
                        font['name'], font['emb'], font['sub'], font['uni']
                    )
                    for font in fonts
                    if font['emb'] != 'yes' or font['sub'] != 'yes' or font['uni'] != 'yes'
                ]
                if incomplete:
                    checks.append(
                        ('fonts', 'not embedded/subset/Unicode-mapped: '
                                  + ', '.join(incomplete))
                    )
            # Base-14 and other resource-only fonts can be listed by pdffonts
            # without being painted. Outside a strict expectation they are a
            # warning, not a failure.
            elif unembedded:
                warnings.append(
                    f"{relative}: non-embedded font resource(s): "
                    + ', '.join(font['name'] for font in unembedded)
                )
            for font in fonts:
                if font['emb'] == 'yes' and font['uni'] != 'yes':
                    warnings.append(f"{relative}: {font['name']} has no ToUnicode map")
    else:
        warnings.append(f'{relative}: pdffonts unavailable')

    if tools['pdftotext']:
        result = run_tool([tools['pdftotext'], str(pdf_path), '-'])
        if result is None:
            checks.append(('pdftotext', 'pdftotext could not run'))
        else:
            if result.returncode != 0:
                checks.append(('pdftotext', f'exit {result.returncode}'))
            if result.stderr.strip():
                checks.append(
                    ('pdftotext', f'stderr: {result.stderr.strip().splitlines()[-1]}')
                )
            if isinstance(text_chars, int) and text_chars > 0 and not result.stdout.strip():
                checks.append(('pdftotext', 'no text extracted'))
            pdftotext_output = result.stdout
            if '\ufffd' in pdftotext_output or '(cid:' in pdftotext_output:
                checks.append(('pdftotext', 'replacement glyph or (cid:) in extracted text'))
    else:
        warnings.append(f'{relative}: pdftotext unavailable')

    combined_text = text + '\n' + (pdftotext_output or '')
    # pypdfium2 cannot always read a synthesized subset (colour emoji); poppler
    # still extracts it, so count both rather than failing on an empty first
    # extractor.
    fixture['extractable_chars'] = len(combined_text.strip())
    if isinstance(text_chars, int) and text_chars > 0 and fixture['extractable_chars'] == 0:
        checks.append(('text', f'no extractable text although the DOCX has {text_chars} characters'))
    combined_normalized = normalize_text(combined_text)
    recovered_scripts = scripts_in(combined_text)
    recovered_emoji = sorted({ch for ch in combined_text if is_emoji_char(ch)})

    if info is not None:
        required = info.scripts & {'cjk', 'rtl'}
        missing = sorted(required - recovered_scripts)
        if missing:
            checks.append(('text', f'extracted text has no {", ".join(missing)} characters'))
        if 'emoji' in info.scripts and not recovered_emoji:
            fixture['emoji_glyph_coverage'] = {
                'expected': None,
                'recovered': [],
                'missing': None,
            }
            warnings.append(
                f'{relative}: document contains emoji but no extractor recovered one'
            )

    if expectation:
        textual = []
        emoji_expected = []
        for expected in expectation.get('text', []):
            emoji_chars = [ch for ch in expected if is_emoji_char(ch)]
            if emoji_chars and len(emoji_chars) == len(expected.replace(' ', '')):
                emoji_expected.extend(emoji_chars)
            else:
                textual.append(expected)
        for expected in textual:
            if normalize_text(expected) not in combined_normalized:
                checks.append(
                    ('text', f'no extractor recovered expected text {expected!r}')
                )
        if emoji_expected:
            expected_emoji = sorted(set(emoji_expected))
            missing = [ch for ch in expected_emoji if ch not in recovered_emoji]
            fixture['emoji_glyph_coverage'] = {
                'expected': expected_emoji,
                'recovered': recovered_emoji,
                'missing': missing,
            }
            if missing:
                warnings.append(
                    f'{relative}: emoji glyph coverage is incomplete; missing '
                    + ', '.join(missing)
                )

    embedded = entry.get('fonts_embedded')
    if isinstance(embedded, int) and embedded == 0 and isinstance(text_chars, int) and text_chars > 0:
        checks.append(('fonts', 'manifest records no embedded font programs'))

    return fixture


if __name__ == '__main__':
    sys.exit(main())
