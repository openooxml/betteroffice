import hashlib
import json
import os
import platform
import re
import signal
import subprocess
import time
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy
import PIL
import skimage
from PIL import Image, ImageStat

from compare import Pages, compare
from raster import RASTERIZER, open_pdf, render_page

CHANNELS = ['published', 'commit', 'libreoffice']


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path, value):
    path.write_text(json.dumps(value, indent=2, allow_nan=False) + '\n')


def execute(command, timeout=180, log=None):
    started = time.perf_counter_ns()
    with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                          text=True, start_new_session=True) as process:
        try:
            stdout, stderr = process.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            os.killpg(process.pid, signal.SIGKILL)
            process.communicate()
            raise RuntimeError(f'Command timed out after {timeout}s: {command[0]}')
    elapsed = (time.perf_counter_ns() - started) / 1e6
    if log:
        log.write_text(stdout + stderr)
    if process.returncode:
        raise RuntimeError(f'{command[0]} exited {process.returncode}: {stderr[-4000:]}')
    return elapsed, stdout


def font_environment(manifest_path, output):
    manifest = json.loads(manifest_path.read_text())
    families = {}
    for face in manifest['faces']:
        file = manifest_path.parent / face['file']
        if file.parent != manifest_path.parent or sha256(file) != face['sha256']:
            raise ValueError('Bundled font hash mismatch')
        families[face['family'].lower()] = face['family']
        if face.get('metricCompatWith'):
            families[face['metricCompatWith'].lower()] = face['family']
    config = ET.Element('fontconfig')
    ET.SubElement(config, 'dir').text = str(manifest_path.parent)
    ET.SubElement(config, 'cachedir').text = str(output / 'font-cache')
    aliases = {**families, **{key: families[value] for key, value in manifest['aliases'].items() if value in families},
               'Calibri Light': 'Carlito', 'sans-serif': 'Liberation Sans',
               'serif': 'Liberation Serif', 'monospace': 'Liberation Mono'}
    for requested, substitute in sorted(aliases.items()):
        alias = ET.SubElement(config, 'alias', binding='strong')
        ET.SubElement(alias, 'family').text = requested
        ET.SubElement(ET.SubElement(alias, 'prefer'), 'family').text = substitute
    path = output / 'fonts.conf'
    ET.ElementTree(config).write(path, encoding='utf-8', xml_declaration=True)
    if platform.system() == 'Linux':
        os.environ['FONTCONFIG_FILE'] = str(path)
        execute(['fc-cache', '-f'])
    return sha256(manifest_path)


def lo_command(soffice, profile, output, source, filter_name):
    return [soffice, '-env:UserInstallation=' + profile.as_uri(), '--headless',
            '--nologo', '--nodefault', '--nofirststartwizard', '--convert-to',
            filter_name, '--outdir', str(output), str(source)]


def edge_adjust(image, target):
    if image.size == target:
        return image, None
    if any(abs(a - b) > 1 for a, b in zip(image.size, target)):
        return image, None
    adjustment = dict(original=list(image.size), target=list(target))
    edged = Image.new('RGB', target, 'white')
    edged.paste(image, (0, 0))
    return edged, adjustment


def libreoffice_fidelity(sample, root, soffice, version):
    format = sample.get('format', 'docx')
    actual = root / 'libreoffice'
    actual.mkdir()
    result = dict(channel='libreoffice', version=version)
    stage = 'capture'
    try:
        execute(lo_command(soffice, root / 'profile-fidelity', actual, root / f'source.{format}',
                           'pdf:impress_pdf_Export' if format == 'pptx' else 'pdf:writer_pdf_Export'), timeout=600, log=actual / 'convert.log')
        adjustments = []
        with open_pdf(actual / 'source.pdf') as pdf:
            if not len(pdf):
                raise ValueError('LibreOffice produced no pages')
            for index in range(len(pdf)):
                image = render_page(pdf, index, 150)
                reference = root / 'reference' / f'page_{index + 1:04d}.png'
                if reference.exists():
                    with Image.open(reference) as expected:
                        image, adjustment = edge_adjust(image, expected.size)
                    if adjustment:
                        adjustments.append(dict(page=index + 1, **adjustment))
                image.save(actual / f'page_{index + 1:04d}.png')
            write_json(actual / 'result.json', dict(status='ok', sha256=sample['metadata']['source']['sha256'],
                       dpi=150, pages=len(pdf), engine='LibreOffice ' + version,
                       rasterizer=RASTERIZER, edge_adjustments=adjustments))
        stage = 'compare'
        def progress(page, total):
            if page % 10 == 0 or page == total:
                print(f"{sample['id']}: scored {page}/{total} LibreOffice pages", flush=True)
        result.update(compare(Pages(root / 'reference', 150), Pages(actual, 150), root / 'lo-comparison',
                              gallery=False, progress=progress))
        result['status'] = 'ok'
    except Exception as error:
        result.update(status='failed', stage=stage, error=str(error))
    return result


def validate_png(path, size, require_ink):
    with Image.open(path) as image:
        image.load()
        if image.format != 'PNG' or any(abs(a - b) > 1 for a, b in zip(image.size, size)):
            raise ValueError(f'Invalid PNG dimensions: {image.size}, expected {size} ±1px')
        variance = ImageStat.Stat(image.convert('L')).var[0]
        if require_ink and variance < 0.1:
            raise ValueError('Renderer produced a blank first page')
        return dict(size=list(image.size), sha256=sha256(path), bytes=path.stat().st_size,
                    grayscale_variance=variance)


def measure_sample(sample, root, binaries, manifest, soffice, trials, offset=0, directory='timing', chains=None):
    format = sample.get('format', 'docx')
    output = root / directory
    output.mkdir()
    with Image.open(root / 'reference/page_0001.png') as reference:
        size = tuple(round(value * 96 / 150) for value in reference.size)
        require_ink = ImageStat.Stat(reference.convert('L')).var[0] >= 1
    options = {key: dict(type='long', value=value) for key, value in
               [('PixelWidth', size[0]), ('PixelHeight', size[1]), ('PageNumber', 1)]}
    options['Translucent'] = dict(type='boolean', value=False)
    commands = {channel: [str(binary), str(root / f'source.{format}'), str(output / f'{channel}.png'), str(manifest)]
                + ([str(chains[channel])] if chains and channel in chains else [])
                for channel, binary in binaries.items()}
    font_paths = {channel: 'product-registry' if chains and channel in chains else 'bench-manifest' for channel in binaries}
    commands['libreoffice'] = lo_command(soffice, root / 'profile-timing', output, root / f'source.{format}',
                                       f'png:{"impress" if format == "pptx" else "writer"}_png_Export:' + json.dumps(options, separators=(',', ':')))
    paths = {channel: output / (f'{channel}.png' if channel != 'libreoffice' else 'source.png') for channel in CHANNELS}
    results = {channel: dict(status='ok', elapsed_ms=[], command=commands[channel]) for channel in CHANNELS}
    for repeat in range(trials + 1):
        rotation = (repeat + offset) % len(CHANNELS)
        order = CHANNELS[rotation:] + CHANNELS[:rotation]
        for channel in order:
            if results[channel]['status'] == 'failed':
                continue
            print(f"{sample['id']}: {channel} {'warmup' if repeat == 0 else f'trial {repeat}/{trials}'}", flush=True)
            try:
                paths[channel].unlink(missing_ok=True)
                elapsed, stdout = execute(commands[channel], log=output / f'{channel}-{repeat}.log')
                png = validate_png(paths[channel], size, require_ink)
                if channel != 'libreoffice':
                    metadata = json.loads(stdout)
                    if (metadata.get('status') != 'ok' or metadata.get('skipped_images') != 0
                            or metadata.get('source_sha256') != sample['metadata']['source']['sha256']
                            or metadata.get('png_sha256') != png['sha256'] or metadata.get('page') != 1
                            or metadata.get('dpi') != 96 or metadata.get('pages', 0) < 1):
                        raise ValueError('Native CLI returned invalid render metadata')
                    if chains is not None and metadata.get('font_path') != font_paths[channel]:
                        raise ValueError(f"Native CLI took the {metadata.get('font_path')} font path, expected {font_paths[channel]}")
                    faces = metadata.get('faces')
                    if not isinstance(faces, list) or not all(isinstance(face, str) and face for face in faces):
                        raise ValueError('Native CLI reported no list of replayed font faces')
                    replayed = hashlib.sha256('\n'.join(faces).encode()).hexdigest()
                    if results[channel].setdefault('fonts_sha256', replayed) != replayed:
                        raise ValueError('Native CLI replayed different font faces across trials')
                    results[channel]['native'] = metadata
                results[channel]['png'] = png
                if repeat:
                    results[channel]['elapsed_ms'].append(elapsed)
                print(f"{sample['id']}: {channel} {elapsed:.1f} ms", flush=True)
            except Exception as error:
                results[channel] = dict(status='failed', error=str(error), command=commands[channel])
                print(f"{sample['id']}: {channel} FAILED: {error}", flush=True)
    if chains is not None:
        for channel in binaries:
            results[channel]['font_path'] = font_paths[channel]
    return results


LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS = '0.5.1'


def semver_key(version):
    match = re.fullmatch(r'(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?', version if isinstance(version, str) else '')
    if not match:
        raise ValueError(f'Not a semver version: {version!r}')
    prerelease = tuple((0, int(part), '') if part.isdigit() else (1, 0, part) for part in match[4].split('.')) if match[4] else ()
    return int(match[1]), int(match[2]), int(match[3]), not match[4], prerelease


def validate_font_paths(paths, published_version):
    if not isinstance(paths, dict) or set(paths) != {'published', 'commit'} or paths['commit'].get('path') != 'product-registry':
        raise ValueError(f'Invalid product font chain paths: {paths}')
    if paths['published'].get('version') != published_version:
        raise ValueError(f"Font chains were recorded for published {paths['published'].get('version')}, the plan measures {published_version}")
    entry = paths['published']
    if entry.get('path') == 'bench-manifest':
        last = LAST_RELEASE_WITHOUT_REGISTRY_EXPORTS
        if entry.get('last_release_without_registry_exports') != last:
            raise ValueError(f"Font chains name {entry.get('last_release_without_registry_exports')} as the last release without the registry exports, expected {last}")
        if semver_key(entry['version']) > semver_key(last):
            raise ValueError(f"Published {entry['version']} is newer than {last}, the last release without the registry exports, so it must replay the product registry, not the bench manifest")
    elif entry.get('path') != 'product-registry':
        raise ValueError(f"Unknown font path for the published channel: {entry.get('path')}")
    return paths


def product_chains(chains_root, paths, sample):
    chains = {}
    for channel, entry in paths.items():
        if entry['path'] == 'product-registry':
            path = chains_root / channel / f"{sample['id']}.chains.json"
            if not path.is_file():
                raise ValueError(f'Missing product font chains for the {channel} channel: {path.name}')
            chains[channel] = path
    return chains


def font_path_notes(paths, method):
    notes = {}
    for channel in CHANNELS[:2]:
        entry = (paths or {}).get(channel, dict(path='bench-manifest'))
        note = f'native column not comparable with runs before {method}'
        if entry['path'] == 'bench-manifest' and entry.get('last_release_without_registry_exports'):
            note += (f"; {channel} {entry['version']} <= {entry['last_release_without_registry_exports']}, the last release without the registry exports,"
                     ' so its native column is still on the old font map'
                     ' and published against commit on the native column is not like for like until the next release')
        notes[channel] = note
    return notes


def summarize_font_paths(samples, method, paths=None):
    notes = font_path_notes(paths, method)
    lines = [f"- {channel}: {', '.join(sorted({row['timings'][channel].get('font_path', 'bench-manifest') for row in samples})) or 'no documents'}; {notes[channel]}"
             for channel in CHANNELS[:2]]
    text = '### DOCX native font path\n\n' + '\n'.join(lines) + '\n'
    print(text, flush=True)
    if os.environ.get('GITHUB_STEP_SUMMARY'):
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a') as summary:
            summary.write(text)


def main():
    root = Path(os.environ['QUALITY_OUTPUT']).resolve()
    job = json.loads((root / 'job.json').read_text())
    native = Path(os.environ['QUALITY_NATIVE']).resolve()
    manifest = Path(os.environ['QUALITY_FONTS']).resolve()
    chains_root = Path(os.environ['QUALITY_CHAINS']).resolve() if os.environ.get('QUALITY_CHAINS') else None
    if chains_root is None:
        if os.environ.get('QUALITY_NATIVE_FONT_MAP') != 'bench-manifest':
            raise ValueError('QUALITY_CHAINS is unset; set QUALITY_NATIVE_FONT_MAP=bench-manifest to measure with the bench font manifest')
        print('DOCX native font path: bench-manifest (QUALITY_CHAINS unset, QUALITY_NATIVE_FONT_MAP=bench-manifest)', flush=True)
    else:
        print(f'DOCX native font path: product registry chains from {chains_root}', flush=True)
    font_paths = validate_font_paths(json.loads((chains_root / 'paths.json').read_text()), job['published_version']) if chains_root else None
    for channel, note in font_path_notes(font_paths, job['method']).items():
        print(f'DOCX native font path, {channel}: {note}', flush=True)
    soffice = os.environ.get('SOFFICE', 'libreoffice')
    os.environ['SAL_USE_VCLPLUGIN'] = 'svp'
    libreoffice_fonts_hash = font_environment(manifest, root)
    _, version_text = execute([soffice, '--version'])
    version = re.search(r'LibreOffice (\d+(?:\.\d+){2,3})\b', version_text)[1]
    builds, binaries = {}, {}
    for channel, source in [('published', job['published_source_sha']), ('commit', job['source_sha'])]:
        binary = native / channel / 'docx-native-bench'
        build = json.loads((binary.parent / 'build.json').read_text())
        if build['source_sha'] != source or sha256(binary) != build['binary_sha256']:
            raise ValueError('Native executable does not match the frozen source')
        builds[channel], binaries[channel] = build, binary
    report = dict(schema_version=1, plan_sha256=job['plan_sha256'], shard=job['shard'], samples=[],
                  benchmark=dict(method=job['method'], trials=job['trials'], dpi=96,
                                 published_version=job['published_version'],
                                 libreoffice_version=version, libreoffice_build=version_text.strip(),
                                 libreoffice_fonts_sha256=libreoffice_fonts_hash, font_paths=font_paths, builds=builds, environment=dict(
                                     os=platform.system(), architecture=platform.machine(),
                                     image_os=os.environ.get('ImageOS', 'local'),
                                     image_version=os.environ.get('ImageVersion', 'local'),
                                     python=platform.python_version(), rasterizer=RASTERIZER,
                                     numpy=numpy.__version__, pillow=PIL.__version__, skimage=skimage.__version__,
                                     font_policy='bundled-fontconfig' if platform.system() == 'Linux' else 'system-fonts')),
                  runner=dict(name=os.environ.get('RUNNER_NAME', platform.node()),
                              platform=platform.platform(), processors=os.cpu_count()))
    for index, sample in enumerate(job['samples']):
        directory = root / sample['id']
        if sha256(directory / 'source.docx') != sample['metadata']['source']['sha256']:
            raise ValueError('Source hash mismatch')
        for page, asset in enumerate(sample['metadata']['reference_pages'], 1):
            if sha256(directory / 'reference' / f'page_{page:04d}.png') != asset['sha256']:
                raise ValueError('Reference hash mismatch')
        print(f"DOCX shard {job['shard']}: {index + 1}/{len(job['samples'])} {sample['id']}", flush=True)
        chains = product_chains(chains_root, font_paths, sample) if chains_root else {}
        timings = measure_sample(sample, directory, binaries, manifest, soffice, job['trials'], index, chains=chains)
        print(f"{sample['id']}: LibreOffice fidelity", flush=True)
        fidelity = libreoffice_fidelity(sample, directory, soffice, version)
        report['samples'].append(dict(id=sample['id'], source_sha256=sample['metadata']['source']['sha256'],
                                      libreoffice=fidelity, timings=timings))
        write_json(root / 'report.json', report)
        print(f"{sample['id']}: fidelity {fidelity['status']}, SSIM {fidelity.get('penalized_ssim', '—')}", flush=True)
    summarize_font_paths(report['samples'], job['method'], font_paths)


if __name__ == '__main__':
    main()
