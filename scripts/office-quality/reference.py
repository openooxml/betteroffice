import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import platform
import plistlib
import shutil
import subprocess
import tempfile
import time
import zipfile

import fitz
from xlsx_reference import export_profile

OFFICE = {
    '.docx': ('Word', 'com.microsoft.Word', '''
        open file name inputPath read only true add to recent files false
        set benchmarkDocument to active document
        try
            save as benchmarkDocument file name outputPath file format format PDF add to recent files false
        on error messageText number errorNumber
            close benchmarkDocument saving no
            error messageText number errorNumber
        end try
        close benchmarkDocument saving no
    '''),
    '.pptx': ('PowerPoint', 'com.microsoft.Powerpoint', '''
        open (POSIX file inputPath)
        set benchmarkDocument to active presentation
        try
            save benchmarkDocument in (POSIX file outputPath) as save as PDF
        on error messageText number errorNumber
            close benchmarkDocument saving no
            error messageText number errorNumber
        end try
        close benchmarkDocument saving no
    '''),
    '.xlsx': ('Excel', 'com.microsoft.Excel', '''
        set benchmarkDocument to open workbook workbook file name inputPath update links do not update links read only true add to mru false
        try
            save workbook as benchmarkDocument filename outputPath file format PDF file format
        on error messageText number errorNumber
            close benchmarkDocument saving no
            error messageText number errorNumber
        end try
        close benchmarkDocument saving no
    '''),
}


def main():
    parser = argparse.ArgumentParser(description='Export one local Office document to reference PDF and PNGs on macOS.')
    parser.add_argument('source', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    parser.add_argument('--dpi', type=int, default=150)
    parser.add_argument('--timeout', type=int, default=120)
    parser.add_argument('--xlsx-profile', type=Path)
    args = parser.parse_args()
    if platform.system() != 'Darwin':
        parser.error('desktop Microsoft Office on macOS is required')
    extension = args.source.suffix.lower()
    if extension not in OFFICE or args.dpi <= 0 or args.timeout <= 0:
        parser.error('use .docx, .pptx, or .xlsx and positive DPI/timeout values')
    if args.xlsx_profile and extension != '.xlsx':
        parser.error('--xlsx-profile requires an XLSX document')
    if args.out.exists() and any(args.out.iterdir()):
        parser.error('output must be empty; use a fresh directory for each revision')
    data = args.source.read_bytes()
    with zipfile.ZipFile(args.source) as archive:
        if any('vbaproject' in name.lower() or '/macrosheets/' in name.lower() for name in archive.namelist()):
            parser.error('macro-bearing documents are not supported')
    name, bundle, commands = OFFICE[extension]
    app = Path(f'/Applications/Microsoft {name}.app/Contents/Info.plist')
    with app.open('rb') as handle:
        version = plistlib.load(handle)['CFBundleShortVersionString']
    container = Path.home() / 'Library/Containers' / bundle / 'Data/Documents/BetterOfficeBenchmark'
    container.mkdir(parents=True, exist_ok=True)
    args.out.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    record = dict(source=args.source.name, sha256=hashlib.sha256(data).hexdigest(),
                  engine=f'Microsoft {name}', version=version, os=platform.mac_ver()[0],
                  dpi=args.dpi, export='local PDF', status='error',
                  started_at_utc=datetime.now(timezone.utc).isoformat())
    stage = Path(tempfile.mkdtemp(prefix='reference-', dir=container))
    source = stage / ('source' + extension)
    pdf = stage / 'reference.pdf'
    source.write_bytes(data)
    script = f'''on run arguments
    set inputPath to item 1 of arguments
    set outputPath to item 2 of arguments
    with timeout of {args.timeout} seconds
        tell application "Microsoft {name}"
{commands}
        end tell
    end timeout
end run
'''
    try:
        if args.xlsx_profile:
            profile = json.loads(args.xlsx_profile.read_text())
            record['capture_profile'] = export_profile(source, pdf, profile, args.timeout, args.dpi)
        else:
            subprocess.run(['osascript', '-', str(source), str(pdf)], input=script,
                           text=True, capture_output=True, check=True, timeout=args.timeout + 5)
        with fitz.open(pdf) as document:
            if document.page_count == 0:
                raise ValueError('Office exported no pages')
            fonts = set()
            page_bounds = []
            for index, page in enumerate(document):
                pixmap = page.get_pixmap(dpi=args.dpi, alpha=False)
                pixmap.save(args.out / f'page_{index + 1:04d}.png')
                fonts.update(font[3] for font in page.get_fonts())
                if extension == '.docx' and args.dpi == 150:
                    page_bounds.append(dict(width_pt=page.rect.width, height_pt=page.rect.height,
                                            width_px=pixmap.width, height_px=pixmap.height))
            if page_bounds:
                record['capture_profile'] = dict(kind='office-page-bounds', pages=page_bounds)
            record.update(status='ok', pages=len(document), fonts=sorted(fonts), pdf_metadata=document.metadata)
        shutil.copy2(pdf, args.out / 'reference.pdf')
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, ValueError, OSError, fitz.FileDataError) as error:
        record['error'] = getattr(error, 'stderr', None) or str(error)
        record['staging_directory'] = str(stage)
        raise SystemExit(f'Office export failed. Check its permission/dialog state before retrying. {record["error"]}') from error
    finally:
        record['ms'] = (time.monotonic() - started) * 1000
        record['finished_at_utc'] = datetime.now(timezone.utc).isoformat()
        (args.out / 'result.json').write_text(json.dumps(record, indent=2) + '\n')
        if record['status'] == 'ok':
            shutil.rmtree(stage)
    print(json.dumps(record))


if __name__ == '__main__':
    main()
