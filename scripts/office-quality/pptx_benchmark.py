import json
import os
import platform
import re
from pathlib import Path

from docx_benchmark import execute, font_environment, libreoffice_fidelity, measure_sample, sha256, write_json


def main():
    root = Path(os.environ['QUALITY_OUTPUT']).resolve()
    job = json.loads((root/'job.json').read_text())
    soffice = os.environ.get('SOFFICE','libreoffice')
    os.environ['SAL_USE_VCLPLUGIN'] = 'svp'
    manifest = Path(os.environ['QUALITY_FONTS']).resolve()
    fonts_hash = font_environment(manifest,root)
    native = Path(os.environ['QUALITY_NATIVE']).resolve()
    builds, binaries = {}, {}
    for channel, source in [('published',job['published_source_sha']),('commit',job['source_sha'])]:
        binary = native/channel/'pptx-native-bench'
        build = json.loads((binary.parent/'build.json').read_text())
        if build['source_sha'] != source or sha256(binary) != build['binary_sha256']:
            raise ValueError('Native executable does not match the frozen source')
        builds[channel], binaries[channel] = build, binary
    _, version_text = execute([soffice,'--version'])
    version = re.search(r'LibreOffice (\d+(?:\.\d+){2,3})\b',version_text)[1]
    report = dict(schema_version=1,plan_sha256=job['plan_sha256'],samples=[],benchmark=dict(
        method=job['method'],dpi=96,trials=job['trials'],published_version=job['published_version'],builds=builds,libreoffice_version=version,libreoffice_build=version_text.strip(),
        fonts_sha256=fonts_hash,platform=platform.platform(),
        font_policy='bundled-fontconfig' if platform.system() == 'Linux' else 'system-fonts'))
    for index,sample in enumerate(job['samples']):
        directory = root/sample['id']
        if sha256(directory/'source.pptx') != sample['metadata']['source']['sha256']:
            raise ValueError('Source hash mismatch')
        for page,asset in enumerate(sample['metadata']['reference_pages'],1):
            if sha256(directory/'reference'/f'page_{page:04d}.png') != asset['sha256']:
                raise ValueError('Reference hash mismatch')
        print(f"PPTX LibreOffice: {index+1}/{len(job['samples'])} {sample['id']}",flush=True)
        timings = measure_sample(sample,directory,binaries,manifest,soffice,job['trials'],index)
        result = libreoffice_fidelity(sample,directory,soffice,version)
        report['samples'].append(dict(id=sample['id'],source_sha256=sample['metadata']['source']['sha256'],libreoffice=result,timings=timings))
        write_json(root/'report.json',report)
        print(f"{sample['id']}: {result['status']}, SSIM {result.get('penalized_ssim','—')}",flush=True)


if __name__ == '__main__':
    main()
