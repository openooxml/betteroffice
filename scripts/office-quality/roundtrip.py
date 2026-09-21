import hashlib
import io
import json
import math
import os
import re
import signal
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import suppress
import zipfile
from collections import Counter
from pathlib import Path
from xml.dom import minidom

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
MARKER = '_BetterOfficeRoundtrip_'


def digest(data):
    return hashlib.sha256(data).hexdigest()


def write_json(path, value):
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(value, indent=2, allow_nan=False) + '\n')
    temporary.replace(path)


def package(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = [entry for entry in archive.infolist() if not entry.is_dir()]
        if len(entries) > 20000 or sum(entry.file_size for entry in entries) > 512 * 1024 * 1024:
            raise ValueError('Package exceeds the preservation checker budget')
        if len({entry.filename for entry in entries}) != len(entries):
            raise ValueError('Duplicate package parts')
        return {entry.filename: archive.read(entry) for entry in entries}


def xml(data):
    document = minidom.parseString(data)
    if document.doctype:
        raise ValueError('DTD-bearing XML is outside the preservation probe')
    return document


def elements(node):
    return [child for child in node.childNodes if child.nodeType == child.ELEMENT_NODE]


def text(node):
    return ''.join(child.data for child in node.childNodes if child.nodeType in (child.TEXT_NODE, child.CDATA_SECTION_NODE))


def path_to(node):
    path = []
    while node.parentNode:
        path.append(elements(node.parentNode).index(node))
        node = node.parentNode
    return list(reversed(path))


def at_path(document, path):
    node = document
    for index in path:
        node = elements(node)[index]
    return node


def select_probe(parts, format):
    if format == 'xlsx':
        from xlsx_calc import sheets
        for name, part, _ in sheets(parts, worksheets_only=True):
            sheet = xml(parts[part])
            for cell in sheet.getElementsByTagNameNS(S, 'c'):
                values = [node for node in elements(cell) if node.namespaceURI == S and node.localName == 'v']
                if (cell.getAttribute('t') not in ('', 'n') or len(values) != 1
                        or any(node.namespaceURI == S and node.localName == 'f' for node in elements(cell))):
                    continue
                if len(values[0].childNodes) != 1 or values[0].firstChild.nodeType not in (values[0].TEXT_NODE, values[0].CDATA_SECTION_NODE):
                    continue
                old = text(values[0])
                try:
                    number = float(old)
                except ValueError:
                    continue
                if not math.isfinite(number):
                    continue
                return dict(status='ok', part=part, path=path_to(values[0]), old=old,
                            new='1' if number == 0 else '0', sheet=name, cell=cell.getAttribute('r'))
    else:
        namespace = W if format == 'docx' else A
        names = ['word/document.xml'] if format == 'docx' else sorted(
            name for name in parts if re.fullmatch(r'ppt/slides/slide[0-9]+\.xml', name))
        candidates, contents = [], []
        for part in names:
            document = xml(parts[part])
            for paragraph in document.getElementsByTagNameNS(namespace, 'p'):
                runs = [node for node in elements(paragraph) if node.namespaceURI == namespace and node.localName == 'r']
                texts = paragraph.getElementsByTagNameNS(namespace, 't')
                content = ''.join(text(node) for node in texts)
                contents.append(content)
                if len(runs) != 1 or len(texts) != 1 or texts[0].parentNode is not runs[0]:
                    continue
                allowed = {'pPr', 'r'} if format == 'docx' else {'pPr', 'r', 'endParaRPr'}
                if any(node.namespaceURI != namespace or node.localName not in allowed for node in elements(paragraph)):
                    continue
                if any(node.namespaceURI != namespace or node.localName not in ('rPr', 't') for node in elements(runs[0])):
                    continue
                if format == 'pptx' and len(paragraph.parentNode.getElementsByTagNameNS(A, 'p')) != 1:
                    continue
                if len(texts[0].childNodes) != 1 or texts[0].firstChild.nodeType not in (texts[0].TEXT_NODE, texts[0].CDATA_SECTION_NODE):
                    continue
                if not content or content != content.strip() or MARKER in content:
                    continue
                candidates.append(dict(status='ok', part=part, path=path_to(texts[0]), old=content, new=content + MARKER))
        counts = Counter(contents)
        for candidate in candidates:
            if counts[candidate['old']] == 1 and sum(content.count(candidate['old']) for content in contents) == 1:
                return candidate
    return dict(status='unavailable', error='No unambiguous plain-text or numeric-literal edit candidate')


def signature(node):
    if node.nodeType == node.DOCUMENT_NODE:
        return ['document', [signature(child) for child in node.childNodes]]
    if node.nodeType == node.ELEMENT_NODE:
        attrs = sorted((attr.namespaceURI or '', attr.name, attr.value) for attr in node.attributes.values())
        return ['element', node.namespaceURI, node.nodeName, attrs, [signature(child) for child in node.childNodes]]
    if node.nodeType in (node.TEXT_NODE, node.CDATA_SECTION_NODE):
        return ['text', node.data]
    if node.nodeType == node.COMMENT_NODE:
        return ['comment', node.data]
    if node.nodeType == node.PROCESSING_INSTRUCTION_NODE:
        return ['instruction', node.target, node.data]
    raise ValueError('Unsupported XML node in preservation comparison')


def verify_preservation(before, after, probe):
    expected = xml(before[probe['part']])
    node = at_path(expected, probe['path'])
    if (text(node) != probe['old'] or len(node.childNodes) != 1
            or node.firstChild.nodeType not in (node.TEXT_NODE, node.CDATA_SECTION_NODE) or probe['old'] == probe['new']):
        raise ValueError('Invalid edit footprint')
    for child in list(node.childNodes):
        node.removeChild(child)
    node.appendChild(expected.createTextNode(probe['new']))
    added, removed = sorted(after.keys() - before.keys()), sorted(before.keys() - after.keys())
    changed = sorted(name for name in before.keys() & after.keys() if before[name] != after[name])
    unrelated = [name for name in changed if name != probe['part']]
    edit_matches = probe['part'] in after and signature(expected) == signature(xml(after[probe['part']]))
    preserved = not added and not removed and not unrelated and edit_matches
    return dict(status='ok' if preserved else 'failed', stage='preserve', edit_matches=edit_matches,
                original_parts=len(before), identical_parts=len(before.keys() & after.keys()) - len(changed),
                added_parts=added, removed_parts=removed, changed_parts=changed,
                unrelated_changed_parts=unrelated,
                **({} if preserved else dict(error='Package differs outside the exact intended edit')))


def run_process(command, timeout, log):
    with log.open('w') as stream:
        with subprocess.Popen(command, stdout=stream, stderr=stream, start_new_session=True) as process:
            try:
                process.wait(timeout=timeout)
                if process.returncode:
                    raise RuntimeError(f'Process exited {process.returncode}')
            except subprocess.TimeoutExpired as error:
                raise RuntimeError(f'Process timed out after {timeout}s') from error
            finally:
                with suppress(ProcessLookupError):
                    os.killpg(process.pid, signal.SIGKILL)
                process.wait()


def helper(mode, arguments, result, timeout):
    run_process([sys.executable, str(Path(__file__).resolve()), mode, *map(str, arguments), str(result)],
                timeout, result.with_suffix('.log'))
    return json.loads(result.read_text())


def measure(binary, source, probe, output, timeout=180, helper_timeout=30):
    output.mkdir()
    status_path = output / 'status.json'
    saved = output / source.name
    prefix = list(map(str, binary)) if isinstance(binary, (list, tuple)) else [str(binary)]
    command = [*prefix, str(source), str(probe), str(saved), str(status_path)]
    failure = None
    try:
        run_process(command, timeout, output / 'native.log')
    except Exception as error:
        failure = str(error)
    try:
        native = json.loads(status_path.read_text())
    except (OSError, ValueError):
        native = {}
    if native.get('source_sha256') != digest(source.read_bytes()):
        native = {}
        failure = failure or 'Native source identity is missing or mismatched'
    parsed = native.get('parse') == 'ok'
    error = failure or native.get('error') or 'Native process did not complete the edit/save probe'
    result = dict(parse=dict(status='ok') if parsed else dict(status='failed', error=error),
                  roundtrip=dict(status='failed', stage=native.get('stage', 'parse'), error=error), native=native)
    if not failure and parsed and native.get('stage') == 'complete' and native.get('edit_verified') is True:
        try:
            if native.get('output_sha256') != digest(saved.read_bytes()):
                raise ValueError('Native output hash mismatch')
            result['roundtrip'] = helper('verify', [source, saved, probe], output / 'preservation.json', helper_timeout)
        except Exception as error:
            result['roundtrip'] = dict(status='failed', stage='preserve', error=str(error))
    return result


def measure_sample(sample, root, job, binaries):
    directory = root / sample['id']
    source = directory / f"source.{job['format']}"
    if digest(source.read_bytes()) != sample['metadata']['source']['sha256']:
        raise ValueError('Source hash mismatch')
    try:
        probe = helper('probe', [source, job['format']], directory / 'probe.json', job['helper_timeout_seconds'])
    except Exception as error:
        probe = dict(status='unavailable', error=str(error))
        write_json(directory / 'probe.json', probe)
    row = dict(id=sample['id'], source_sha256=sample['metadata']['source']['sha256'], probe=probe, channels={})
    for channel, binary in binaries.items():
        print(f"{job['format'].upper()} roundtrip: {sample['id']} {channel}", flush=True)
        row['channels'][channel] = measure(binary, source, directory / 'probe.json', directory / channel,
                                           job['timeout_seconds'], job['helper_timeout_seconds'])
    return row


def main():
    root = Path(os.environ['QUALITY_OUTPUT']).resolve()
    job = json.loads((root / 'job.json').read_text())
    native = Path(os.environ['QUALITY_NATIVE']).resolve()
    builds, binaries = {}, {}
    for channel, revision in [('published', job['published_source_sha']), ('commit', job['source_sha'])]:
        binary = native / channel / f"{job['format']}-roundtrip-bench"
        build = json.loads((binary.parent / 'build.json').read_text())
        if build['source_sha'] != revision or build['binary_sha256'] != digest(binary.read_bytes()):
            raise ValueError('Native build does not match the frozen source')
        builds[channel], binaries[channel] = build, binary
    if not 1 <= len(job['samples']) <= job['shard_size'] or job['shard_size'] != 16 or job['parallelism'] != 4:
        raise ValueError('Invalid bounded roundtrip shard')
    if job['timeout_seconds'] != 180 or job['helper_timeout_seconds'] != 30:
        raise ValueError('Invalid roundtrip deadlines')
    soffice = os.environ['SOFFICE']
    lo_python = os.environ['LO_PYTHON']
    host = Path(__file__).with_name('roundtrip_libreoffice.py')
    subprocess.run([lo_python, str(host), '--check'], check=True, timeout=30)
    version_text = subprocess.run([soffice, '--version'], check=True, capture_output=True, text=True, timeout=30).stdout.strip()
    version = re.search(r'LibreOffice (\d+(?:\.\d+){2,3})\b', version_text)[1]
    binaries['libreoffice'] = [lo_python, str(host)]
    report = dict(schema_version=1, plan_sha256=job['plan_sha256'], format=job['format'], shard=job['shard'], samples=[], benchmark=dict(
        method=job['method'], published_version=job['published_version'], builds=builds,
        checker_sha256=digest(Path(__file__).read_bytes()), timeout_seconds=job['timeout_seconds'],
        helper_timeout_seconds=job['helper_timeout_seconds'], parallelism=job['parallelism'], shard_size=job['shard_size'],
        libreoffice_version=version, libreoffice_build=version_text, libreoffice_host_sha256=digest(host.read_bytes())))
    rows = {}
    with ThreadPoolExecutor(max_workers=job['parallelism']) as pool:
        futures = [pool.submit(measure_sample, sample, root, job, binaries) for sample in job['samples']]
        for future in as_completed(futures):
            row = future.result()
            rows[row['id']] = row
            report['samples'] = [rows[sample['id']] for sample in job['samples'] if sample['id'] in rows]
            write_json(root / 'report.json', report)
            print(f"{job['format'].upper()} roundtrip: completed {len(rows)}/{len(job['samples'])}", flush=True)


if __name__ == '__main__':
    if len(sys.argv) == 5 and sys.argv[1] == 'probe':
        write_json(Path(sys.argv[4]), select_probe(package(Path(sys.argv[2]).read_bytes()), sys.argv[3]))
    elif len(sys.argv) == 6 and sys.argv[1] == 'verify':
        write_json(Path(sys.argv[5]), verify_preservation(package(Path(sys.argv[2]).read_bytes()),
                   package(Path(sys.argv[3]).read_bytes()), json.loads(Path(sys.argv[4]).read_text())))
    elif len(sys.argv) == 1:
        main()
    else:
        raise ValueError('Invalid roundtrip helper arguments')
