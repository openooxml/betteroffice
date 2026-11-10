import json
import os
import platform
import re
from pathlib import Path

import docx_benchmark
import xlsx_benchmark
from docx_benchmark import CHANNELS, execute, font_environment, sha256, write_json

ABS_TOLERANCE = 1e-9
REL_TOLERANCE = 1e-12
BINARIES = dict(docx='docx-native-bench', pptx='pptx-native-bench', xlsx='xlsx-native-bench')


def verify(sample, root, format):
    if sha256(root / f'source.{format}') != sample['metadata']['source']['sha256']:
        raise ValueError('Source hash mismatch')
    if format != 'xlsx' and sha256(root / 'reference' / 'page_0001.png') != sample['metadata']['reference_pages'][0]['sha256']:
        raise ValueError('Reference hash mismatch')


def measure_round(format, sample, root, binaries, manifest, soffice, offset, directory):
    if format == 'xlsx':
        config = dict(trials=1, absolute_tolerance=ABS_TOLERANCE, relative_tolerance=REL_TOLERANCE)
        return xlsx_benchmark.measure_sample(sample, root, binaries, soffice, config, offset, directory)
    return docx_benchmark.measure_sample(sample, root, binaries, manifest, soffice, 1, offset, directory)


def main():
    root = Path(os.environ['QUALITY_OUTPUT']).resolve()
    job = json.loads((root / 'job.json').read_text())
    format = job['format']
    native = Path(os.environ['QUALITY_NATIVE']).resolve()
    soffice = os.environ.get('SOFFICE', 'libreoffice')
    os.environ['SAL_USE_VCLPLUGIN'] = 'svp'
    manifest = None
    fonts_hash = None
    if format != 'xlsx':
        manifest = Path(os.environ['QUALITY_FONTS']).resolve()
        fonts_hash = font_environment(manifest, root)
    _, version_text = execute([soffice, '--version'])
    version = re.search(r'LibreOffice (\d+(?:\.\d+){2,3})\b', version_text)[1]
    builds, binaries = {}, {}
    for channel, source in [('published', job['published_source_sha']), ('commit', job['source_sha'])]:
        binary = native / channel / BINARIES[format]
        build = json.loads((binary.parent / 'build.json').read_text())
        if build['source_sha'] != source or sha256(binary) != build['binary_sha256']:
            raise ValueError('Native executable does not match the frozen source')
        builds[channel], binaries[channel] = build, binary
    samples = job['samples']
    for sample in samples:
        verify(sample, root / sample['id'], format)
    rows = [dict(id=sample['id'], source_sha256=sample['metadata']['source']['sha256'],
                 timings={channel: dict(status='ok', elapsed_ms=[]) for channel in CHANNELS}) for sample in samples]
    report = dict(schema_version=1, plan_sha256=job['plan_sha256'], format=format, samples=rows,
                  benchmark=dict(method=job['method'], rounds=job['rounds'], seed=job['seed'], sample_size=len(samples),
                                 published_version=job['published_version'], libreoffice_version=version,
                                 libreoffice_build=version_text.strip(), fonts_sha256=fonts_hash, builds=builds,
                                 order='sequential; engine order rotates per document and round; one warm-up per round'),
                  runner=dict(name=os.environ.get('RUNNER_NAME', platform.node()), platform=platform.platform(),
                              processors=os.cpu_count(), image_version=os.environ.get('ImageVersion', 'local')))
    for turn in range(job['rounds']):
        for index, (sample, row) in enumerate(zip(samples, rows)):
            print(f"{format.upper()} timing round {turn + 1}/{job['rounds']}: {index + 1}/{len(samples)} {sample['id']}", flush=True)
            measured = measure_round(format, sample, root / sample['id'], binaries, manifest, soffice,
                                     index + turn, f'timing-{turn + 1}')
            for channel in CHANNELS:
                current = row['timings'][channel]
                if current['status'] == 'failed':
                    continue
                result = measured[channel]
                if result['status'] != 'ok' or len(result['elapsed_ms']) != 1:
                    row['timings'][channel] = dict(status='failed', error=f"round {turn + 1}: {result.get('error', 'no timing')}")
                else:
                    current['elapsed_ms'].append(result['elapsed_ms'][0])
        write_json(root / 'report.json', report)
    complete = sum(all(row['timings'][channel]['status'] == 'ok' for channel in CHANNELS) for row in rows)
    print(f'{format.upper()} timing: {complete}/{len(rows)} documents timed by every engine in every round', flush=True)
    if not complete:
        raise SystemExit('No document was timed by every engine')


if __name__ == '__main__':
    main()
