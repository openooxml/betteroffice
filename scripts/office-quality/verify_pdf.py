#!/usr/bin/env python3
"""Structural checks for a PDF exported by betteroffice-docx-pdf.

This is a report gate, not a Word-fidelity comparison: it asserts page count,
page size, extractable text and embedded fonts. Visual fidelity still uses the
existing Word-reference comparison.
"""

import argparse
import sys
from pathlib import Path

import pypdfium2 as pdfium


def parse_size(value):
    width, _, height = value.lower().partition('x')
    if not width or not height:
        raise argparse.ArgumentTypeError('expected WIDTHxHEIGHT in points')
    return float(width), float(height)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('--pages', type=int)
    parser.add_argument('--page-size', type=parse_size)
    parser.add_argument('--contains', action='append', default=[])
    parser.add_argument('--allow-base14', action='store_true')
    args = parser.parse_args()

    raw = args.pdf.read_bytes()
    if not raw.startswith(b'%PDF-'):
        parser.error(f'{args.pdf}: not a PDF')
    if not args.allow_base14 and b'/BaseFont /Helvetica' in raw:
        parser.error(f'{args.pdf}: contains an unembedded Base-14 fallback font')
    if b'/FontFile2' not in raw and b'/FontFile3' not in raw:
        parser.error(f'{args.pdf}: no embedded font programs')
    if b'/ToUnicode' not in raw:
        parser.error(f'{args.pdf}: no ToUnicode map')

    document = pdfium.PdfDocument(str(args.pdf))
    try:
        if args.pages is not None and len(document) != args.pages:
            parser.error(f'{args.pdf}: {len(document)} pages, expected {args.pages}')
        texts = []
        for index in range(len(document)):
            page = document[index]
            try:
                width, height = page.get_width(), page.get_height()
                if args.page_size:
                    expected = args.page_size
                    if abs(width - expected[0]) > 0.01 or abs(height - expected[1]) > 0.01:
                        parser.error(
                            f'{args.pdf}: page {index + 1} is {width:.2f}x{height:.2f} pt, '
                            f'expected {expected[0]:.2f}x{expected[1]:.2f}'
                        )
                textpage = page.get_textpage()
                try:
                    texts.append(textpage.get_text_range())
                finally:
                    textpage.close()
            finally:
                page.close()
    finally:
        document.close()

    joined = '\n'.join(texts)
    for needle in args.contains:
        if needle not in joined:
            parser.error(f'{args.pdf}: extracted text does not contain {needle!r}')
    print(f'{args.pdf}: {len(texts)} page(s), embedded fonts, ToUnicode, extractable text')
    return 0


if __name__ == '__main__':
    sys.exit(main())
