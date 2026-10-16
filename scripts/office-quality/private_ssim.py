import argparse
from bisect import bisect_right
from collections import Counter
from difflib import SequenceMatcher
import hashlib
import json
import math
from multiprocessing import Pool
from pathlib import Path
import re

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity


def read_json(path):
    return json.loads(path.read_text())


def write_json(path, value):
    path.write_text(json.dumps(value, indent=2, allow_nan=False) + '\n')


def sha256(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def pdf_lines(page):
    text = page.get_textpage()
    try:
        rects = [text.get_rect(index) for index in range(text.count_rects())]
        groups = []
        for left, bottom, right, top in rects:
            value = text.get_text_bounded(left, bottom, right, top).strip()
            if not value:
                continue
            if not groups or abs(groups[-1]['baseline'] - bottom) > 2:
                groups.append(dict(baseline=bottom, top=top, parts=[]))
            group = groups[-1]
            group['top'] = max(group['top'], top)
            group['parts'].append(' '.join(value.split()))
        return [[page.get_height() - group['top'],
                 ' '.join(group['parts'])] for group in groups]
    finally:
        text.close()


def rasterize(source, output, dpi=150):
    digest = sha256(source)
    result = output / 'result.json'
    if result.exists():
        metadata = read_json(result)
        if metadata.get('status') == 'ok' and metadata.get('sha256') == digest and metadata.get('dpi') == dpi:
            return metadata
    import pypdfium2 as pdfium

    output.mkdir(parents=True, exist_ok=True)
    write_json(result, dict(source=str(source), sha256=digest, dpi=dpi, status='running'))
    for path in output.glob('page_*.png'):
        path.unlink()
    pdf = pdfium.PdfDocument(str(source))
    pages = []
    try:
        for index in range(len(pdf)):
            page = pdf[index]
            try:
                bitmap = page.render(scale=dpi / 72, fill_color=(255, 255, 255, 255))
                try:
                    image = bitmap.to_pil().convert('RGB')
                    image.save(output / f'page_{index + 1:04d}.png')
                    pages.append(dict(page=index + 1, width_pt=page.get_width(), height_pt=page.get_height(),
                                      width_px=image.width, height_px=image.height, lines=pdf_lines(page)))
                    image.close()
                finally:
                    bitmap.close()
            finally:
                page.close()
    finally:
        pdf.close()
    if not pages:
        raise ValueError('reference has no pages')
    write_json(output / 'pages.json', pages)
    metadata = dict(source=str(source), sha256=digest, dpi=dpi, pages=len(pages), status='ok')
    write_json(result, metadata)
    return metadata


def profile(directory):
    metadata = read_json(directory / 'result.json')
    if metadata['status'] != 'ok':
        raise ValueError('reference rasterization did not complete')
    dpi = metadata['dpi']
    return dict(kind='office-page-bounds', pages=[
        dict(width_pt=page['width_pt'], height_pt=page['height_pt'],
             width_px=math.ceil(page['width_pt'] * dpi / 72),
             height_px=math.ceil(page['height_pt'] * dpi / 72))
        for page in read_json(directory / 'pages.json')])


def page_paths(directory):
    paths = sorted(directory.glob('page_*.png'))
    if any(path.name != f'page_{index + 1:04d}.png' for index, path in enumerate(paths)):
        raise ValueError(f'{directory}: page filenames must be contiguous starting at page_0001.png')
    metadata = read_json(directory / 'result.json')
    if metadata.get('status') != 'ok' or metadata.get('pages') != len(paths):
        raise ValueError(f'{directory}: render did not complete or page count differs')
    if metadata.get('dpi') != 150:
        raise ValueError(f'{directory}: comparison requires 150 DPI')
    return paths


def words(line):
    return re.findall(r'[^\W_]+', re.sub(r'\d+', '0', line.lower()))


def body_stream(pages):
    normalized = [[words(line) for line in page] for page in pages]
    counts = Counter(line for page in normalized for line in {tuple(line) for line in page if line})
    repeated = {line for line, count in counts.items() if count > max(4, len(pages) / 8)}
    stream, positions, page_words = [], [], []
    for page_index, page in enumerate(normalized):
        lines = [line for line in page if line and tuple(line) not in repeated]
        flattened = []
        for line_index, line in enumerate(lines):
            stream.extend(line)
            flattened.extend(line)
            positions.extend([(page_index + 1, line_index)] * len(line))
        page_words.append(flattened)
    return stream, positions, page_words


def align(reference, actual):
    reference_stream, _, starts = body_stream(reference)
    stream, positions, _ = body_stream(actual)
    blocks = [block for block in SequenceMatcher(None, reference_stream, stream, autojunk=False).get_matching_blocks()
              if block.size >= 5]
    heads = [block.a for block in blocks]
    start, matches = 0, []
    for page in starts:
        match = None
        for index in range(start, start + min(41, len(page))):
            block = bisect_right(heads, index) - 1
            if block >= 0 and index < blocks[block].a + blocks[block].size:
                match = positions[blocks[block].b + index - blocks[block].a]
                break
        matches.append((match[0], match[1] == 0) if match else (None, False))
        start += len(page)
    return matches


def ssim_pair(pair):
    images = []
    for path in pair:
        with Image.open(path) as image:
            rgba = image.convert('RGBA')
        white = Image.new('RGBA', rgba.size, 'white')
        images.append(Image.alpha_composite(white, rgba).convert('L'))
    expected, rendered = images
    note = None
    if expected.size != rendered.size:
        if any(abs(left - right) > 1 for left, right in zip(expected.size, rendered.size)):
            return dict(ssim=0.0, size_note=f'size mismatch: {expected.size} vs {rendered.size}')
        adjusted = Image.new('L', expected.size, 255)
        adjusted.paste(rendered, (0, 0))
        rendered = adjusted
        note = 'white pad/crop at right/bottom edge'
    smallest = min(expected.size)
    if smallest < 3:
        raise ValueError('SSIM needs images at least 3 pixels wide and high')
    window = min(7, smallest if smallest % 2 else smallest - 1)
    score = float(structural_similarity(np.asarray(expected), np.asarray(rendered), data_range=255, win_size=window))
    return dict(ssim=score, size_note=note)


def score_pairs(pairs, jobs, cache):
    pairs = list(dict.fromkeys(pairs))
    cache_file = cache / 'ssim.json' if cache else None
    stored = read_json(cache_file) if cache_file and cache_file.exists() else {}
    hashes = {path: sha256(path) for path in {path for pair in pairs for path in pair}} if cache else {}
    keys = {pair: f'{hashes[pair[0]]}:{hashes[pair[1]]}' for pair in pairs} if cache else {}
    scores = {pair: stored[keys[pair]] for pair in pairs if cache and keys[pair] in stored}
    pending = [pair for pair in pairs if pair not in scores]
    if pending:
        if jobs == 1:
            scores.update(zip(pending, map(ssim_pair, pending)))
        else:
            with Pool(jobs) as pool:
                scores.update(zip(pending, pool.imap(ssim_pair, pending, chunksize=1)))
    if cache:
        cache.mkdir(parents=True, exist_ok=True)
        stored.update({keys[pair]: score for pair, score in scores.items()})
        temporary = cache_file.with_suffix('.tmp')
        write_json(temporary, stored)
        temporary.replace(cache_file)
    return scores


def baseline_delta(report, baseline):
    if report['reference_pages'] != baseline['reference_pages']:
        raise ValueError('baseline reference page count differs')
    delta = {key: report[key] - baseline[key] for key in
             ['penalized_ssim', 'aligned_mean_ssim', 'common_page_ssim', 'actual_pages', 'page_agreement']}
    delta['first_divergence_page'] = dict(before=baseline['first_divergence_page'], after=report['first_divergence_page'])
    old = {row['page']: row['ssim'] or 0.0 for row in baseline['pages']}
    delta['dropped_pages'] = [row['page'] for row in report['pages']
                              if old.get(row['page'], 0.0) - row['ssim'] > 0.005]
    regression = (delta['penalized_ssim'] < -0.001 or delta['page_agreement'] < 0 or
                  abs(report['actual_pages'] - report['reference_pages']) >
                  abs(baseline['actual_pages'] - baseline['reference_pages']))
    return delta, regression


def compare(reference, actual, output, jobs=4, cache=None, baseline=None):
    if jobs < 1:
        raise ValueError('jobs must be positive')
    left, right = page_paths(reference), page_paths(actual)
    if not left:
        raise ValueError('reference has no pages')
    reference_text = [[line[1] for line in page['lines']] for page in read_json(reference / 'pages.json')]
    actual_text = read_json(actual / 'pages_text.json')
    if len(reference_text) != len(left) or len(actual_text) != len(right):
        raise ValueError('text and image page counts differ')
    matches = align(reference_text, actual_text)
    common = min(len(left), len(right))
    pairs = [(left[index], right[index]) for index in range(common)]
    pairs.extend((left[index], right[page - 1]) for index, (page, _) in enumerate(matches) if page)
    scores = score_pairs(pairs, jobs, cache)
    rows = []
    for index in range(max(len(left), len(right))):
        score = scores[(left[index], right[index])] if index < common else dict(ssim=0.0, size_note='missing or extra page')
        page, agrees = matches[index] if index < len(matches) else (None, False)
        aligned = scores[(left[index], right[page - 1])]['ssim'] if page else None
        rows.append(dict(page=index + 1, **score, aligned_page=page, break_agrees=agrees, aligned_ssim=aligned))
    aligned_scores = [row['aligned_ssim'] for row in rows if row['aligned_ssim'] is not None]
    agreement = sum(row['aligned_page'] == row['page'] and row['break_agrees'] for row in rows[:len(left)])
    total = sum(row['ssim'] for row in rows)
    report = dict(reference_pages=len(left), actual_pages=len(right), common_page_ssim=total / common if common else 0.0,
                  penalized_ssim=total / len(rows), aligned_mean_ssim=sum(aligned_scores) / len(aligned_scores) if aligned_scores else 0.0,
                  aligned_pages=len(aligned_scores), first_divergence_page=next((row['page'] for row in rows[:len(left)] if row['page'] > len(right) or (
                      row['aligned_page'] and (row['aligned_page'] != row['page'] or not row['break_agrees']))), None),
                  page_agreement=agreement, page_agreement_fraction=agreement / len(left),
                  break_agreement=sum(row['break_agrees'] for row in rows[:len(left)]), pages=rows)
    if baseline:
        report['delta'], report['regression'] = baseline_delta(report, read_json(baseline))
    output.mkdir(parents=True, exist_ok=True)
    write_json(output / 'score.json', report)
    return report


def main():
    parser = argparse.ArgumentParser(description='Local Word-reference SSIM and page-break gate.')
    commands = parser.add_subparsers(dest='command', required=True)
    raster = commands.add_parser('rasterize')
    raster.add_argument('reference', type=Path)
    raster.add_argument('--out', type=Path, required=True)
    raster.add_argument('--dpi', type=int, default=150)
    bounds = commands.add_parser('profile')
    bounds.add_argument('reference', type=Path)
    scoring = commands.add_parser('compare')
    scoring.add_argument('reference', type=Path)
    scoring.add_argument('actual', type=Path)
    scoring.add_argument('--out', type=Path, required=True)
    scoring.add_argument('--jobs', type=int, default=4)
    scoring.add_argument('--cache', type=Path)
    scoring.add_argument('--baseline', type=Path)
    args = parser.parse_args()
    try:
        if args.command == 'rasterize':
            if args.dpi <= 0:
                parser.error('DPI must be positive')
            result = rasterize(args.reference, args.out, args.dpi)
            print(json.dumps(dict(dpi=result['dpi'], pages=result['pages'])))
        elif args.command == 'profile':
            print(json.dumps(profile(args.reference)))
        else:
            report = compare(args.reference, args.actual, args.out, args.jobs, args.cache, args.baseline)
            print(json.dumps({key: value for key, value in report.items() if key != 'pages'}, allow_nan=False))
            if report.get('regression'):
                raise SystemExit(1)
    except (ValueError, OSError, KeyError) as error:
        parser.error(str(error))


if __name__ == '__main__':
    main()
