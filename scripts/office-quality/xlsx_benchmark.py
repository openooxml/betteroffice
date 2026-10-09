import hashlib
import json
import math
import os
import platform
import re
from pathlib import Path

from docx_benchmark import CHANNELS, execute, lo_command, sha256, write_json
from xlsx_calc import formula_targets, package, results, sheets, uncached_workbook


def compare_values(expected, actual, absolute_tolerance, relative_tolerance):
    if len(expected) != len(actual):
        raise ValueError('Formula result count mismatch')
    mismatches = []
    correct = 0
    for reference, observed in zip(expected, actual):
        if (reference['sheet'],reference['cell']) != (observed['sheet'],observed['cell']):
            raise ValueError('Formula cell identity mismatch')
        left,right = reference['value'],observed['value']
        matches = left == right
        if left['kind'] == right['kind'] == 'number':
            matches = math.isclose(left['value'],right['value'],abs_tol=absolute_tolerance,rel_tol=relative_tolerance)
        if matches:
            correct += 1
        elif len(mismatches) < 100:
            mismatches.append(dict(sheet=reference['sheet'],cell=reference['cell'],expected=left,actual=right))
    return dict(correct=correct,total=len(expected),mismatches=mismatches)


def prepare_input(source, reference, frozen=None):
    data = source.read_bytes()
    if reference.get('schema_version') != 1 or reference.get('method') != 'excel-full-rebuild-v1' or reference.get('engine') != 'Microsoft Excel':
        raise ValueError('Invalid calculation oracle')
    if hashlib.sha256(data).hexdigest() != reference['source_sha256']:
        raise ValueError('Calculation source hash mismatch')
    parts = package(data)
    targets,formula_count = formula_targets(parts)
    if targets != reference['clear_cells'] or formula_count != reference['formula_cells']:
        raise ValueError('Formula targets do not match the source workbook')
    cleared = uncached_workbook(data,targets)
    if frozen is not None:
        if package(frozen) != package(cleared):
            raise ValueError('Frozen calculation input differs from the uncached workbook')
        cleared = frozen
    if hashlib.sha256(cleared).hexdigest() != reference['input_sha256']:
        raise ValueError('Uncached input differs from the Excel oracle input')
    if any(cell['value']['kind'] != 'empty' for cell in results(package(cleared),targets)):
        raise ValueError('Formula caches were not fully cleared')
    known = {(cell['sheet'],cell['cell']) for cell in targets}
    expected = [(cell['sheet'],cell['cell']) for cell in reference['cells']]
    if not expected or len(set(expected)) != len(expected) or not set(expected).issubset(known):
        raise ValueError('Invalid scored formula results')
    return cleared


def measure_sample(sample, root, binaries, soffice, config, offset=0):
    reference = json.loads((root/'expected.json').read_text())
    if sha256(root/'expected.json') != sample['metadata']['calculation']['reference']['sha256']:
        raise ValueError('Calculation reference hash mismatch')
    input_path = root/'input.xlsx'
    if sha256(input_path) != sample['metadata']['calculation']['input']['sha256']:
        raise ValueError('Frozen calculation input hash mismatch')
    prepare_input(root/'source.xlsx',reference,input_path.read_bytes())
    output = root/'timing'
    output.mkdir()
    lo_output = output/'libreoffice'
    lo_output.mkdir()
    paths = {channel:output/f'{channel}.xlsx' for channel in binaries}
    paths['libreoffice'] = lo_output/'input.xlsx'
    commands = {channel:[str(binary),str(input_path),str(paths[channel])] for channel,binary in binaries.items()}
    commands['libreoffice'] = lo_command(soffice,root/'profile-calculation',lo_output,input_path,'xlsx:Calc MS Excel 2007 XML')
    measured = {channel:dict(status='ok',elapsed_ms=[]) for channel in CHANNELS}
    previous = {}
    for repeat in range(config['trials']+1):
        rotation = (repeat+offset) % len(CHANNELS)
        for channel in CHANNELS[rotation:]+CHANNELS[:rotation]:
            if measured[channel]['status'] == 'failed':
                continue
            try:
                paths[channel].unlink(missing_ok=True)
                elapsed,stdout = execute(commands[channel],timeout=120,log=output/f'{channel}-{repeat}.log')
                if channel != 'libreoffice':
                    record = json.loads(stdout)
                    if record.get('status') != 'ok' or record.get('input_sha256') != reference['input_sha256'] or record.get('output_sha256') != sha256(paths[channel]):
                        raise ValueError('Native executable returned invalid identity')
                parts = package(paths[channel].read_bytes())
                if [name for name,_,_ in sheets(parts)] != reference['sheets']:
                    raise ValueError('Recalculation changed worksheet identity')
                actual = results(parts,reference['cells'])
                if channel in previous and actual != previous[channel]:
                    raise ValueError('Recalculation results changed between trials')
                previous[channel] = actual
                accuracy = compare_values(reference['cells'],actual,config['absolute_tolerance'],config['relative_tolerance'])
                measured[channel].update(accuracy)
                if repeat:
                    measured[channel]['elapsed_ms'].append(elapsed)
                print(f"{sample['id']}: {channel} trial {repeat}/{config['trials']}, {accuracy['correct']}/{accuracy['total']} correct, {elapsed:.1f} ms",flush=True)
            except Exception as error:
                measured[channel] = dict(status='failed',correct=0,total=len(reference['cells']),error=str(error))
                print(f"{sample['id']}: {channel} FAILED: {error}",flush=True)
    return measured


def main():
    root = Path(os.environ['QUALITY_OUTPUT']).resolve()
    job = json.loads((root/'job.json').read_text())
    native = Path(os.environ['QUALITY_NATIVE']).resolve()
    soffice = os.environ.get('SOFFICE','libreoffice')
    os.environ['SAL_USE_VCLPLUGIN'] = 'svp'
    _,version_text = execute([soffice,'--version'])
    version = re.search(r'LibreOffice (\d+(?:\.\d+){2,3})\b',version_text)[1]
    builds,binaries = {},{}
    for channel,source in [('published',job['published_source_sha']),('commit',job['source_sha'])]:
        binary = native/channel/'xlsx-native-bench'
        build = json.loads((binary.parent/'build.json').read_text())
        if build['source_sha'] != source or sha256(binary) != build['binary_sha256']:
            raise ValueError('Native calculation executable does not match the frozen source')
        builds[channel],binaries[channel] = build,binary
    config = dict(method=job['method'],trials=job['trials'],absolute_tolerance=job['absolute_tolerance'],relative_tolerance=job['relative_tolerance'],
        published_version=job['published_version'],libreoffice_version=version,libreoffice_build=version_text.strip(),builds=builds)
    report = dict(schema_version=1,plan_sha256=job['plan_sha256'],shard=job['shard'],benchmark=config,samples=[],
        runner=dict(platform=platform.platform(),processors=os.cpu_count()))
    for index,sample in enumerate(job['samples']):
        print(f"XLSX calculation shard {job['shard']}: {index+1}/{len(job['samples'])} {sample['id']}",flush=True)
        calculation = measure_sample(sample,root/sample['id'],binaries,soffice,config,index)
        report['samples'].append(dict(id=sample['id'],source_sha256=sample['metadata']['source']['sha256'],
            reference_sha256=sample['metadata']['calculation']['reference']['sha256'],calculations=calculation))
        write_json(root/'report.json',report)


if __name__ == '__main__':
    main()
