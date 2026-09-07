import argparse
import html
import json
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageChops, ImageEnhance
from skimage.metrics import structural_similarity


class Pages:
    def __init__(self, path, dpi):
        self.pdf = fitz.open(path) if path.suffix.lower() == '.pdf' else None
        self.paths = [] if self.pdf else sorted(path.glob('page_*.png'))
        if any(page.name != f'page_{index + 1:04d}.png' for index, page in enumerate(self.paths)):
            raise ValueError(f'{path}: page filenames must be contiguous starting at page_0001.png')
        self.dpi = dpi
        metadata = path / 'result.json'
        self.metadata = json.loads(metadata.read_text()) if path.is_dir() and metadata.exists() else {}
        if self.metadata and self.metadata.get('status') != 'ok':
            raise ValueError(f'{path}: render did not complete successfully')
        if self.metadata.get('dpi', dpi) != dpi:
            raise ValueError(f'{path}: DPI does not match --dpi')
        if self.metadata.get('pages', len(self)) != len(self):
            raise ValueError(f'{path}: page files do not match the recorded page count')

    def __len__(self):
        return len(self.pdf) if self.pdf is not None else len(self.paths)

    def page(self, index):
        if self.pdf is not None:
            pixmap = self.pdf[index].get_pixmap(dpi=self.dpi, alpha=False)
            return Image.frombytes('RGB', (pixmap.width, pixmap.height), pixmap.samples)
        with Image.open(self.paths[index]) as image:
            rgba = image.convert('RGBA')
        white = Image.new('RGBA', rgba.size, 'white')
        return Image.alpha_composite(white, rgba).convert('RGB')


def compare(reference, actual, output, resize=False):
    if not len(reference):
        raise ValueError('reference has no pages')
    left_hash = reference.metadata.get('sha256')
    right_hash = actual.metadata.get('sha256')
    if left_hash and right_hash and left_hash != right_hash:
        raise ValueError('reference and actual come from different source documents')
    output.mkdir(parents=True, exist_ok=True)
    scores, rows = [], []
    count = max(len(reference), len(actual))
    for index in range(count):
        expected = reference.page(index) if index < len(reference) else None
        rendered = actual.page(index) if index < len(actual) else None
        score = None
        original_size = rendered.size if rendered else None
        if expected is not None and rendered is not None:
            if expected.size != rendered.size:
                if not resize:
                    raise ValueError(f'page {index + 1}: {expected.size} vs {rendered.size}; pass --resize only for a deliberate resize-to-match comparison')
                rendered = rendered.resize(expected.size, Image.Resampling.LANCZOS)
            smallest = min(expected.size)
            if smallest < 3:
                raise ValueError('SSIM needs images at least 3 pixels wide and high')
            window = min(7, smallest if smallest % 2 else smallest - 1)
            score = float(structural_similarity(np.asarray(expected.convert('L')), np.asarray(rendered.convert('L')), data_range=255, win_size=window))
            scores.append(score)
        available = expected if expected is not None else rendered
        expected = expected if expected is not None else Image.new('RGB', available.size, 'white')
        rendered = rendered if rendered is not None else Image.new('RGB', available.size, 'white')
        prefix = f'page_{index + 1:04d}'
        expected.save(output / f'{prefix}.reference.png')
        rendered.save(output / f'{prefix}.actual.png')
        ImageEnhance.Brightness(ImageChops.difference(expected, rendered)).enhance(4).save(output / f'{prefix}.diff.png')
        rows.append(dict(page=index + 1, ssim=score, reference_size=list(expected.size), actual_size=original_size))
    report = dict(reference_pages=len(reference), actual_pages=len(actual),
                  common_page_ssim=sum(scores) / len(scores) if scores else 0.0,
                  penalized_ssim=sum(scores) / count, source_verified=bool(left_hash and right_hash),
                  resized=resize, reference=reference.metadata, actual=actual.metadata, pages=rows)
    (output / 'score.json').write_text(json.dumps(report, indent=2) + '\n')
    cards = []
    for row in rows:
        prefix = f'page_{row["page"]:04d}'
        caption = f'Page {row["page"]}: SSIM {row["ssim"]:.4f}' if row['ssim'] is not None else f'Page {row["page"]}: missing or extra'
        images = ''.join(f'<figure><figcaption>{label}</figcaption><img loading="lazy" src="{prefix}.{kind}.png"></figure>' for kind, label in [('reference', 'Office reference'), ('actual', 'Actual'), ('diff', 'Difference ×4')])
        cards.append(f'<h2>{html.escape(caption)}</h2><section>{images}</section>')
    summary = f'<p>Reference: {len(reference)} pages · Actual: {len(actual)} pages · Common-page SSIM: {report["common_page_ssim"]:.4f} · Penalized SSIM: {report["penalized_ssim"]:.4f}</p><p>Source hash verified: {report["source_verified"]} · Resize allowed: {resize}</p>'
    (output / 'index.html').write_text('<!doctype html><meta charset="utf-8"><title>Office comparison</title><style>body{font:16px system-ui;margin:24px}section{display:flex;gap:16px}figure{margin:0;width:33%}img{width:100%;border:1px solid #ccc}figcaption{margin:8px 0}</style><h1>Office comparison</h1>' + summary + ''.join(cards))
    return report


def main():
    parser = argparse.ArgumentParser(description='Diff Office/renderer PDFs or page_0001.png directories locally.')
    parser.add_argument('reference', type=Path)
    parser.add_argument('actual', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    parser.add_argument('--dpi', type=int, default=150)
    parser.add_argument('--resize', action='store_true')
    args = parser.parse_args()
    if args.dpi <= 0 or not args.reference.exists() or not args.actual.exists():
        parser.error('inputs must exist and DPI must be positive')
    if args.out.exists() and any(args.out.iterdir()):
        parser.error('output must be empty')
    try:
        report = compare(Pages(args.reference, args.dpi), Pages(args.actual, args.dpi), args.out, args.resize)
    except ValueError as error:
        parser.error(str(error))
    print(json.dumps({key: report[key] for key in ['reference_pages', 'actual_pages', 'common_page_ssim', 'penalized_ssim', 'source_verified', 'resized']}))


if __name__ == '__main__':
    main()
