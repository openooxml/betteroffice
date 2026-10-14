import argparse
import contextlib
import hashlib
import importlib.util
import json
import os
import re
import selectors
import shlex
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[1]
TIMEOUT_SECONDS = 600
STARTUP_TIMEOUT_SECONDS = 60
DISABLED_FEATURES = ('shell_tool', 'unified_exec_tty', 'shell_snapshot',
                     'browser_use', 'browser_use_external', 'browser_use_full_cdp_access',
                     'in_app_browser', 'computer_use', 'apps', 'view_image', 'in_app_local_automation',
                     'plugins', 'remote_plugin', 'multi_agent', 'hooks',
                     'skill_search', 'skill_mcp_dependency_install', 'image_generation',
                     'workspace_dependencies', 'worktrees')
FORMAT_TOOLS = {'xlsx': ('xlsx_outline', 'xlsx_read_range', 'xlsx_write_range', 'xlsx_edit_grid', 'xlsx_edit_sheets',
                         'xlsx_find_replace', 'xlsx_format', 'xlsx_sort'),
                'pptx': ('pptx_outline', 'pptx_read_slide', 'pptx_edit'),
                'docx': ('office_outline', 'office_grep', 'office_read', 'office_propose', 'office_accept', 'docx_edit')}
SHELL_FEATURES = ('shell_tool', 'unified_exec', 'unified_exec_tty')
ALLOWED_ITEMS = {'betteroffice': ('agent_message', 'reasoning', 'error'),
                 'generic': ('agent_message', 'reasoning', 'error', 'command_execution', 'file_change')}
TOOL_CALLS = {'betteroffice': 'BetterOffice MCP calls', 'generic': 'shell commands and file changes'}
module_spec = importlib.util.spec_from_file_location('format_check', ROOT / 'check.py')
checker = importlib.util.module_from_spec(module_spec)
module_spec.loader.exec_module(checker)
FORMATS = tuple(dict.fromkeys(task['format'] for task in checker.tasks()))


def format_selection(value):
    names = {name.strip() for name in value.split(',')}
    if not names <= set(FORMATS):
        raise argparse.ArgumentTypeError(f'unknown format in {value!r}; choose from {",".join(FORMATS)}')
    return tuple(name for name in FORMATS if name in names)


def cli_command(command):
    args = shlex.split(command)
    if not args:
        raise ValueError('agents-cli-command is empty')
    executable = shutil.which(args[0])
    if executable is None:
        raise ValueError(f'cannot find MCP executable: {args[0]}')
    args[0] = str(Path(executable).resolve())
    for i in range(1, len(args)):
        path = Path(args[i]).expanduser()
        if not args[i].startswith('-') and path.exists():
            args[i] = str(path.resolve())
    return args


def required_tools(tasks):
    return sorted({'office_open', 'office_export', *(tool for task in tasks for tool in FORMAT_TOOLS[task['format']])})


def mcp_tools(process, deadline):
    selector = selectors.DefaultSelector()
    selector.register(process.stdout, selectors.EVENT_READ)
    pending = b''

    def gone(method):
        try:
            code = process.wait(timeout=1)
        except subprocess.TimeoutExpired:
            return f'closed its output before answering {method}'
        return f'exited with code {code} before answering {method}'

    def call(identity, method, params=None):
        nonlocal pending
        message = {'jsonrpc': '2.0', 'method': method, **({'params': params} if params is not None else {}),
                   **({'id': identity} if identity is not None else {})}
        try:
            process.stdin.write(json.dumps(message).encode() + b'\n')
            process.stdin.flush()
        except BrokenPipeError:
            raise RuntimeError(gone(method)) from None
        while identity is not None:
            while b'\n' in pending:
                line, pending = pending.split(b'\n', 1)
                try:
                    reply = json.loads(line)
                except ValueError:
                    continue
                if isinstance(reply, dict) and reply.get('id') == identity:
                    if 'error' in reply:
                        raise RuntimeError(f'{method} failed: {reply["error"]}')
                    return reply.get('result') or {}
            remaining = deadline - time.monotonic()
            if remaining <= 0 or not selector.select(remaining):
                raise RuntimeError(f'did not answer {method} within {STARTUP_TIMEOUT_SECONDS} seconds')
            chunk = os.read(process.stdout.fileno(), 65536)
            if not chunk:
                raise RuntimeError(gone(method))
            pending += chunk

    call(1, 'initialize', {'protocolVersion': '2025-06-18', 'capabilities': {},
                           'clientInfo': {'name': 'format-eval-preflight', 'version': '1'}})
    call(None, 'notifications/initialized')
    names, cursor, identity = set(), None, 2
    while True:
        page = call(identity, 'tools/list', {'cursor': cursor} if cursor else {})
        names.update(tool.get('name') for tool in page.get('tools', []))
        cursor, identity = page.get('nextCursor'), identity + 1
        if not cursor:
            return names


def mcp_ready(cli, required):
    with tempfile.TemporaryDirectory(prefix='office-mcp-ready-') as root, tempfile.TemporaryFile() as stderr:
        started = time.monotonic()
        process = subprocess.Popen([*cli, '--root', root], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=stderr,
                                   start_new_session=True)
        try:
            listed, failure = mcp_tools(process, started + STARTUP_TIMEOUT_SECONDS), None
            elapsed = time.monotonic() - started
        except RuntimeError as error:
            failure = error
        finally:
            stop_process(process)
            for pipe in (process.stdin, process.stdout):
                with contextlib.suppress(OSError):
                    pipe.close()
        if failure:
            stderr.seek(0)
            detail = stderr.read().decode(errors='replace').strip()[-2000:]
            raise RuntimeError(f'MCP server not ready: {failure}' + (f'; stderr: {detail}' if detail else ''))
    missing = sorted(set(required) - listed)
    if missing:
        raise RuntimeError('MCP server not ready: missing required tools: ' + ', '.join(missing))
    return round(elapsed, 3)


def python_ready(python):
    try:
        result = subprocess.run([str(python), '-c', 'import openpyxl, pptx, docx'], stdin=subprocess.DEVNULL,
                                capture_output=True, text=True, timeout=STARTUP_TIMEOUT_SECONDS)
    except (OSError, subprocess.TimeoutExpired) as error:
        raise RuntimeError(f'{python} cannot import openpyxl, pptx and docx: {error}') from None
    if result.returncode != 0:
        raise RuntimeError(f'{python} cannot import openpyxl, pptx and docx: {result.stderr.strip()[-2000:]}')


def command_line(cli, workspace, model, prompt, arm='betteroffice'):
    if arm == 'generic':
        return ['codex', 'exec', '-m', model, '--sandbox', 'workspace-write', '--skip-git-repo-check',
                '--ignore-user-config', '--ignore-rules', '--ephemeral', '--json', '-C', str(workspace),
                *[argument for feature in SHELL_FEATURES for argument in ('--enable', feature)],
                *[argument for feature in DISABLED_FEATURES if feature not in SHELL_FEATURES for argument in ('--disable', feature)],
                '-c', 'approval_policy="never"',
                '-c', 'web_search="disabled"',
                '-c', 'sandbox_workspace_write.network_access=false',
                prompt + ' Use shell commands and Python with openpyxl, python-pptx and python-docx to read and edit the documents.'
                ' There is no network access. All other tool calls disqualify the task. The documents live in the current directory.'
                ' Use only this directory, the Python interpreter and its standard library/site-packages.'
                ' Commands that reference the repository, evaluation sources or task metadata, and file changes outside this directory disqualify the task.']
    return ['codex', 'exec', '-m', model, '--sandbox', 'read-only', '--skip-git-repo-check',
            '--ignore-user-config', '--ignore-rules', '--ephemeral', '--json', '-C', str(workspace),
            *[argument for feature in DISABLED_FEATURES for argument in ('--disable', feature)],
            '-c', 'approval_policy="never"',
            '-c', 'web_search="disabled"',
            '-c', 'mcp_servers.betteroffice.command=' + json.dumps(cli[0]),
            '-c', 'mcp_servers.betteroffice.args=' + json.dumps(cli[1:] + ['--root', str(workspace)]),
            '-c', 'mcp_servers.betteroffice.required=true',
            '-c', f'mcp_servers.betteroffice.startup_timeout_sec={STARTUP_TIMEOUT_SECONDS}',
            prompt + ' Use only the betteroffice MCP tools to read and edit the documents. All other tool calls disqualify the task. The documents live in the current directory.']


def events(log):
    for line in log.read_text(errors='replace').splitlines():
        try:
            event = json.loads(line)
        except ValueError:
            continue
        if isinstance(event, dict):
            yield event


def counted(item, arm):
    if arm == 'generic':
        return item.get('type') in ('command_execution', 'file_change')
    return item.get('type') == 'mcp_tool_call' and item.get('server') == 'betteroffice'


def tool_calls(log, arm='betteroffice'):
    identities, anonymous, available = set(), 0, False
    for event in events(log):
        if event.get('type') in ('thread.started', 'turn.started', 'turn.completed'):
            available = True
        item = event.get('item', {})
        if not isinstance(item, dict) or not counted(item, arm):
            continue
        available = True
        if 'id' in item:
            identities.add(item['id'])
        elif event.get('type') == 'item.completed':
            anonymous += 1
    return len(identities) + anonymous if available else None


def referenced_paths(text):
    pending, seen, paths = [text], set(), set()
    quoted = re.compile(r'''(['"])(.*?)\1''', re.DOTALL)
    absolute = re.compile(r'''(?<![\w./:])/(?!/)[^\s'"`<>;|&(){}\[\],\\]+''')
    while pending:
        source = pending.pop()
        if source in seen:
            continue
        seen.add(source)
        try:
            tokens = shlex.split(source)
        except ValueError:
            tokens = []
        for token in tokens:
            if re.match(r'^[A-Za-z][\w+.-]*://', token):
                if token.startswith('file://'):
                    paths.add(token[7:])
                continue
            paths.add(token)
            if not token.startswith('/') and re.search(r'''[\s'"`(){}=]''', token) and token != source:
                pending.append(token)
        for match in quoted.finditer(source):
            value = match[2]
            if value.startswith('/') and '\n' not in value:
                paths.add(value)
                source = source.replace(match[0], '')
        paths.update(match[0] for match in absolute.finditer(source))
    return paths


def path_violation(text, workspace):
    names = re.finditer(r'(?<![\w.-])(?:check\.py|validate-checkers\.py|tasks\.json|fixtures\.json)(?![\w.-])', text)
    if any(not (workspace / match[0]).is_file() for match in names) or \
            re.search(r'''\beval[/\\]|\beval\.formats\b|\b(?:from|import)\s+eval\b''', text):
        return 'forbidden path access: evaluation sources or task metadata'
    expansions = {'HOME': str(Path.home()), 'PWD': str(workspace)}
    text = re.sub(r'\$(?:\{(HOME|PWD)\}|(HOME|PWD)\b)', lambda match: expansions[match[1] or match[2]], text)
    text = re.sub(r'(?<![\w.-])~(?=/)', lambda match: expansions['HOME'], text)
    repository = REPO.resolve()
    for spelling in (str(REPO), str(repository)):
        if spelling in text:
            return f'forbidden path access: repository path {spelling}'
    for reference in sorted(referenced_paths(text)):
        try:
            path = Path(reference).expanduser()
            path = (path if path.is_absolute() else workspace / path).resolve()
        except (OSError, RuntimeError, ValueError):
            continue
        if path.is_relative_to(repository):
            return f'forbidden path access: repository path {reference}'
    return None


def access_violation(log, arm='betteroffice', workspace=None):
    for event in events(log):
        item = event.get('item')
        if not isinstance(item, dict):
            continue
        kind = item.get('type')
        if kind == 'mcp_tool_call':
            if arm == 'generic' or item.get('server') != 'betteroffice':
                return f'forbidden tool call: MCP server {item.get("server", "<missing>")}'
        elif kind not in ALLOWED_ITEMS[arm]:
            return f'forbidden tool call: {kind}'
        elif arm == 'generic' and kind in ('command_execution', 'file_change'):
            if workspace is None:
                return 'forbidden path access: task directory unavailable for audit'
            if kind == 'command_execution':
                command = item.get('command')
                if not isinstance(command, str) or not command:
                    return 'forbidden path access: command text unavailable for audit'
                violation = path_violation(command, workspace)
                if violation:
                    return violation
            else:
                changes = item.get('changes')
                if not isinstance(changes, list):
                    return 'forbidden path access: file changes unavailable for audit'
                for change in changes:
                    if not isinstance(change, dict) or not isinstance(change.get('path'), str):
                        return 'forbidden path access: changed path unavailable for audit'
                    try:
                        path = Path(change['path'])
                        path = (path if path.is_absolute() else workspace / path).resolve()
                    except (OSError, RuntimeError, ValueError):
                        return f'forbidden path access: cannot resolve changed path {change["path"]}'
                    if not path.is_relative_to(workspace):
                        return f'forbidden path access: file change outside task directory: {change["path"]}'
                    violation = path_violation(json.dumps(change), workspace)
                    if violation:
                        return violation
    return None


def stop_process(process):
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    process.wait()


def run_task(task, cli, model, results, arm='betteroffice', python=None):
    environment = None if python is None else dict(os.environ, PATH=str(python.parent) + os.pathsep + os.environ.get('PATH', ''))
    directory = results / task['id']
    directory.mkdir()
    log, errors = directory / 'codex.jsonl', directory / 'codex.stderr'
    started = time.monotonic()
    exit_code, timed_out, failure = None, False, None
    temporary_root = Path(tempfile.gettempdir()).resolve()
    if temporary_root.is_relative_to(REPO):
        temporary_root = Path('/tmp').resolve()
    if temporary_root.is_relative_to(REPO):
        raise RuntimeError('task temporary directory must be outside the repository')
    with tempfile.TemporaryDirectory(prefix='office-format-', dir=temporary_root) as temporary:
        workspace = Path(temporary).resolve()
        source = workspace / task['input']
        shutil.copyfile(checker.fixture(task), source)
        original_hash = hashlib.sha256(source.read_bytes()).digest()
        with log.open('w') as stdout, errors.open('w') as stderr:
            try:
                process = subprocess.Popen(command_line(cli, workspace, model, task['prompt'], arm),
                                           cwd=workspace, stdin=subprocess.DEVNULL, stdout=stdout,
                                           stderr=stderr, env=environment, start_new_session=True)
                try:
                    exit_code = process.wait(timeout=TIMEOUT_SECONDS)
                except subprocess.TimeoutExpired:
                    timed_out = True
                    stop_process(process)
                    exit_code = process.returncode
                except BaseException:
                    stop_process(process)
                    raise
            except OSError as error:
                failure = f'agent could not start: {error}'
        try:
            checker.check(task, workspace / task['output'])
            if not source.is_file() or hashlib.sha256(source.read_bytes()).digest() != original_hash:
                raise ValueError('input fixture was changed or removed')
            passed, message = True, 'requested changes and preserved content verified; input unchanged'
        except Exception as error:
            passed, message = False, str(error)
        runtime_error = failure or ('agent timed out after 600 seconds' if timed_out else
                                   f'agent exited with code {exit_code}' if exit_code != 0 else None)
        if runtime_error:
            message = runtime_error + ('; ' + message if not passed else '')
            passed = False
        violation = access_violation(log, arm, workspace)
        if violation:
            message = violation + ('; ' + message if not passed else '')
            passed = False
        calls = tool_calls(log, arm)
        if not calls:
            message = 'agent made no tool calls' + ('; ' + message if not passed else '')
            passed = False
        output = workspace / task['output']
        if output.is_file():
            shutil.copyfile(output, directory / task['output'])
    return {'task': task['id'], 'format': task['format'], 'pass': passed, 'check_message': message,
            'tool_calls': calls, 'duration_seconds': round(time.monotonic() - started, 3),
            'agent_exit_code': exit_code, 'timed_out': timed_out}


def write_results(results, model, records, arm='betteroffice', formats=FORMATS, **metadata):
    (results / 'results.json').write_text(json.dumps({'model': model, 'arm': arm, 'formats': list(formats), **metadata,
                                                    'timeout_seconds': TIMEOUT_SECONDS, 'results': records}, indent=2) + '\n')
    passed = sum(record['pass'] for record in records)
    idle = sum(record['tool_calls'] == 0 for record in records)
    rows = [f'Model: `{model}`. Arm: `{arm}`. Formats: {", ".join(formats)}. Passed: {passed}/{len(records)}. '
            f'Tasks without tool calls: {idle}. '
            f'Tool calls count {TOOL_CALLS[arm]}; n/a means unavailable.',
            '', '| Task | Result | Check message | Tool calls | Duration (s) |',
            '| --- | --- | --- | ---: | ---: |']
    for record in records:
        message = record['check_message'].replace('|', '\\|').replace('\n', ' ')
        calls = record['tool_calls'] if record['tool_calls'] is not None else 'n/a'
        rows.append(f'| {record["task"]} | {"PASS" if record["pass"] else "FAIL"} | {message} | {calls} | {record["duration_seconds"]:.3f} |')
    (results / 'results.md').write_text('\n'.join(rows) + '\n')


def main():
    parser = argparse.ArgumentParser(usage='%(prog)s [--arm betteroffice] [--formats <list>] <agents-cli-command> <results-dir> [model]\n'
                                           '       %(prog)s --arm generic --python <python> [--formats <list>] <results-dir> [model]')
    parser.add_argument('--arm', choices=tuple(ALLOWED_ITEMS), default='betteroffice')
    parser.add_argument('--python', help='generic arm: a Python that imports openpyxl, pptx and docx')
    parser.add_argument('--formats', type=format_selection, default=FORMATS, metavar='<list>',
                        help=f'comma-separated formats to run, from {",".join(FORMATS)} (default: all)')
    parser.add_argument('arguments', nargs='+', help=argparse.SUPPRESS)
    args = parser.parse_args()
    generic = args.arm == 'generic'
    names = ('results_dir', 'model') if generic else ('agents_cli_command', 'results_dir', 'model')
    if not len(names) - 1 <= len(args.arguments) <= len(names):
        parser.error(f'--arm {args.arm} takes ' + ' '.join(f'<{name}>' for name in names[:-1]) + ' [model]')
    values = dict(zip(names, args.arguments))
    model = values.get('model') or 'gpt-6-luna'
    try:
        if generic != (args.python is not None):
            raise ValueError('--python <python> is required with --arm generic and accepted only there')
        cli = None if generic else cli_command(values['agents_cli_command'])
        python = Path(args.python).expanduser().absolute() if generic else None
        results = Path(values['results_dir']).expanduser().resolve()
        if results == REPO or REPO in results.parents:
            raise ValueError('results-dir must be outside the repository')
        if results.exists() and any(results.iterdir()):
            raise ValueError('results-dir must be new or empty')
        if shutil.which('codex') is None:
            raise ValueError('codex is not on PATH')
    except (ValueError, OSError) as error:
        parser.error(str(error))
    tasks = [task for task in checker.tasks() if task['format'] in args.formats]
    try:
        if generic:
            python_ready(python)
            metadata = {'python': str(python)}
        else:
            metadata = {'mcp_startup_seconds': mcp_ready(cli, required_tools(tasks))}
        results.mkdir(parents=True, exist_ok=True)
    except (RuntimeError, OSError) as error:
        sys.exit(f'run.py: {error}')
    records = []
    write_results(results, model, records, args.arm, args.formats, **metadata)
    for task in tasks:
        record = run_task(task, cli, model, results, args.arm, python)
        records.append(record)
        write_results(results, model, records, args.arm, args.formats, **metadata)
        print(f'{record["task"]}: {"PASS" if record["pass"] else "FAIL"} ({record["duration_seconds"]:.3f}s): {record["check_message"]}', flush=True)
    return 0 if all(record['pass'] for record in records) else 1


if __name__ == '__main__':
    raise SystemExit(main())
