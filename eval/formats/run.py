import argparse
import contextlib
import concurrent.futures
import hashlib
import importlib.util
import json
import os
import re
import selectors
import shlex
import shutil
import signal
import stat
import subprocess
import sys
import tempfile
import threading
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
TOOL_CALLS.update({'scripted-bo': 'BetterOffice MCP calls', 'scripted-lib': 'logged library calls'})
SCRIPTED = ROOT / 'scripted'
MUSE_BLOCKER = ('Muse execution is unavailable: muse exec --help and muse --help expose --json and sandbox controls, '
                'but do not establish how to ignore user configuration and configure exactly one MCP server '
                'with no other file-reading tools, or the JSONL file-access audit schema.')
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


def script_value(value, replies):
    if isinstance(value, dict):
        if set(value) == {'$ref'}:
            parts = value['$ref'].split('.')
            result = replies[int(parts[0])]
            for part in parts[1:]:
                result = result[int(part)] if isinstance(result, list) else result[part]
            return result
        return {key: script_value(item, replies) for key, item in value.items()}
    if isinstance(value, list):
        return [script_value(item, replies) for item in value]
    return value


def mcp_tools(process, deadline, requests=None, log=None):
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
                error = RuntimeError if requests is None else TimeoutError
                raise error(f'did not answer {method} within {STARTUP_TIMEOUT_SECONDS if requests is None else TIMEOUT_SECONDS} seconds')
            chunk = os.read(process.stdout.fileno(), 65536)
            if not chunk:
                raise RuntimeError(gone(method))
            pending += chunk

    try:
        call(1, 'initialize', {'protocolVersion': '2025-06-18', 'capabilities': {},
                               'clientInfo': {'name': 'format-eval-preflight', 'version': '1'}})
        call(None, 'notifications/initialized')
        names, cursor, identity = set(), None, 2
        while True:
            page = call(identity, 'tools/list', {'cursor': cursor} if cursor else {})
            names.update(tool.get('name') for tool in page.get('tools', []))
            cursor, identity = page.get('nextCursor'), identity + 1
            if not cursor:
                break
        replies = []
        for index, request in enumerate(requests or []):
            if request.get('method') != 'tools/call':
                raise ValueError('scripted-bo accepts only tools/call requests')
            params = script_value(request['params'], replies)
            if params['name'] not in names:
                raise ValueError(f'unknown scripted tool: {params["name"]}')
            item = {'id': f'script-{index}', 'type': 'mcp_tool_call', 'server': 'betteroffice', **params}
            log.write(json.dumps({'type': 'item.started', 'item': item}) + '\n')
            log.flush()
            result = call(identity + index, 'tools/call', params)
            log.write(json.dumps({'type': 'item.completed', 'item': item, 'result': result}) + '\n')
            log.flush()
            if result.get('isError'):
                raise RuntimeError(f'{params["name"]} failed: {result}')
            value = result.get('structuredContent')
            if value is None:
                value = json.loads(next(block['text'] for block in result.get('content', []) if block.get('type') == 'text'))
            replies.append(value)
        return names
    finally:
        selector.close()


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
                prompt + ' Read and edit the documents however you like: by hand (for example unzip the package, edit its XML parts and zip it again), with any shell tool, or with Python; openpyxl, python-pptx and python-docx are installed but optional.'
                ' There is no network access. All other tool calls disqualify the task. The documents live in the current directory.'
                ' Work only inside this directory; temporary files may go in it.'
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


def final_message(log):
    text = None
    for event in events(log):
        item = event.get('item')
        if event.get('type') == 'item.completed' and isinstance(item, dict) and item.get('type') == 'agent_message' \
                and isinstance(item.get('text'), str):
            text = item['text']
    return text


def reply_answer(text):
    text = text.strip()
    fenced = re.fullmatch(r'```[\w-]*\n(.*?)\n?```', text, re.DOTALL)
    return (fenced.group(1) if fenced else text).strip()


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
        if re.search(r'(?<![\w./\\-])' + re.escape(spelling) + r'(?=/|$|[^\w./\\-])', text):
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


class UnsafeWorkspaceFile(ValueError):
    pass


@contextlib.contextmanager
def workspace_file(workspace, path, mode):
    relative = path.relative_to(workspace)
    if not relative.parts or '..' in relative.parts:
        raise UnsafeWorkspaceFile(f'unsafe workspace file: {relative}')
    descriptors = []
    try:
        if not path.resolve().is_relative_to(workspace):
            raise UnsafeWorkspaceFile(f'unsafe workspace file outside task directory: {relative}')
        directory = os.open(workspace, os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW)
        descriptors.append(directory)
        for component in relative.parts[:-1]:
            if mode == 'wb':
                with contextlib.suppress(FileExistsError):
                    os.mkdir(component, dir_fd=directory)
            directory = os.open(component, os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW, dir_fd=directory)
            descriptors.append(directory)
        name = relative.name
        try:
            before = os.stat(name, dir_fd=directory, follow_symlinks=False)
        except FileNotFoundError:
            if mode != 'wb':
                raise
            before = None
        if before is not None and not stat.S_ISREG(before.st_mode):
            raise UnsafeWorkspaceFile(f'unsafe workspace file is not regular: {relative}')
        flags = os.O_WRONLY | os.O_CREAT if mode == 'wb' else os.O_RDONLY
        descriptor = os.open(name, flags | os.O_NOFOLLOW | os.O_NONBLOCK, 0o600, dir_fd=directory)
        with os.fdopen(descriptor, mode) as stream:
            opened = os.fstat(stream.fileno())
            if not stat.S_ISREG(opened.st_mode) or before is not None and (before.st_dev, before.st_ino) != (opened.st_dev, opened.st_ino):
                raise UnsafeWorkspaceFile(f'unsafe workspace file changed during open: {relative}')
            if mode == 'wb':
                stream.truncate(0)
            yield stream
    except FileNotFoundError:
        raise
    except (OSError, RuntimeError) as error:
        raise UnsafeWorkspaceFile(f'unsafe workspace file: {relative}: {error}') from error
    finally:
        for descriptor in reversed(descriptors):
            os.close(descriptor)


class AgentProcesses:
    def __init__(self):
        self.lock = threading.Lock()
        self.cancelled = threading.Event()
        self.processes = {}

    def start(self, *args, **kwargs):
        with self.lock:
            if self.cancelled.is_set():
                raise concurrent.futures.CancelledError()
            process = subprocess.Popen(*args, **kwargs)
            self.processes[id(process)] = process
            return process

    def discard(self, process):
        with self.lock:
            self.processes.pop(id(process), None)

    def cancel(self):
        with self.lock:
            self.cancelled.set()
            processes = list(self.processes.values())
        for process in processes:
            stop_process(process)


def run_task(task, cli, model, results, arm='betteroffice', python=None, run=None, processes=None):
    if processes is not None and processes.cancelled.is_set():
        raise concurrent.futures.CancelledError()
    start_process = processes.start if processes is not None else subprocess.Popen
    environment = None if python is None else dict(os.environ, PATH=str(python.parent) + os.pathsep + os.environ.get('PATH', ''))
    directory = results / task['id']
    if run is not None:
        directory /= f'run-{run}'
    directory.mkdir(parents=True, exist_ok=True)
    log, errors = directory / 'codex.jsonl', directory / 'codex.stderr'
    if arm.startswith('scripted-'):
        log, errors = directory / 'script.jsonl', directory / 'script.stderr'
    answer_task = task.get('tier') == 'complex' and 'answer' in task['expected']
    artifact_name = 'answer.json' if answer_task else task['output']
    for artifact in (log, errors, directory / artifact_name):
        artifact.unlink(missing_ok=True)
    started = time.monotonic()
    exit_code, timed_out, failure, calls = None, False, None, 0
    passed, message, failure_type = False, '', None
    temporary_root = Path(tempfile.gettempdir()).resolve()
    if temporary_root.is_relative_to(REPO):
        temporary_root = Path('/tmp').resolve()
    if temporary_root.is_relative_to(REPO):
        raise RuntimeError('task temporary directory must be outside the repository')
    with tempfile.TemporaryDirectory(prefix='office-format-', dir=temporary_root) as temporary:
        workspace = Path(temporary).resolve()
        source = workspace / task['input']
        try:
            data = checker.prepare_input(task, checker.fixture(task).read_bytes())
            with workspace_file(workspace, source, 'wb') as stream:
                stream.write(data)
            original_hash = hashlib.sha256(data).digest()
        except Exception as error:
            failure = f'fixture verification failed: {error}'
            if isinstance(error, UnsafeWorkspaceFile):
                failure_type = 'unsafe-workspace-file'
        script = SCRIPTED / ('lib' if arm == 'scripted-lib' else 'bo') / (task['id'] + ('.py' if arm == 'scripted-lib' else '.json'))
        if failure is None and arm.startswith('scripted-') and not script.is_file():
            failure = 'no script'
        if failure is None:
            with log.open('w') as stdout, errors.open('w') as stderr:
                process = None
                try:
                    if arm == 'scripted-bo':
                        requests = json.loads(script.read_text())
                        if not isinstance(requests, list):
                            raise ValueError('scripted-bo requires a list of tools/call requests')
                        process = start_process([*cli, '--root', str(workspace)], cwd=workspace,
                                                   stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=stderr,
                                                   start_new_session=True)
                        mcp_tools(process, started + TIMEOUT_SECONDS, requests, stdout)
                        exit_code = 0
                    else:
                        command = [str(python), str(script)] if arm == 'scripted-lib' else command_line(cli, workspace, model, task['prompt'], arm)
                        process = start_process(command, cwd=workspace, stdin=subprocess.DEVNULL, stdout=stdout,
                                                   stderr=stderr, env=environment, start_new_session=True)
                        exit_code = process.wait(timeout=TIMEOUT_SECONDS)
                except subprocess.TimeoutExpired:
                    timed_out = True
                except TimeoutError as error:
                    timed_out, failure = True, str(error)
                except Exception as error:
                    failure = f'agent could not start: {error}' if isinstance(error, OSError) else str(error)
                finally:
                    if process is not None:
                        if process.poll() is None or hasattr(process, 'pid'):
                            stop_process(process)
                        if processes is not None:
                            processes.discard(process)
                        if exit_code is None:
                            exit_code = process.returncode
                        for pipe in (process.stdin, process.stdout):
                            if pipe is not None:
                                pipe.close()
            reply = final_message(log) if answer_task and not arm.startswith('scripted-') and log.is_file() else None
            try:
                if reply is not None:
                    with workspace_file(workspace, workspace / 'answer.json', 'wb') as stream:
                        stream.write((reply_answer(reply) + '\n').encode())
                with workspace_file(workspace, source, 'rb') as stream:
                    unchanged = hashlib.sha256(stream.read()).digest() == original_hash
                if not unchanged:
                    raise ValueError('input fixture was changed or removed')
                with workspace_file(workspace, workspace / artifact_name, 'rb') as stream:
                    artifact = stream.read()
                saved = directory / artifact_name
                saved.parent.mkdir(parents=True, exist_ok=True)
                saved.write_bytes(artifact)
                checker.check(task, saved)
                passed, message = True, 'requested changes and preserved content verified; input unchanged'
            except Exception as error:
                passed, message = False, str(error)
                if isinstance(error, UnsafeWorkspaceFile):
                    failure_type = 'unsafe-workspace-file'
            runtime_error = failure or (f'agent timed out after {TIMEOUT_SECONDS} seconds' if timed_out else
                                       f'agent exited with code {exit_code}' if exit_code != 0 else None)
            if runtime_error:
                message = runtime_error + ('; ' + message if not passed else '')
                passed = False
            if arm == 'scripted-lib':
                calls = sum(event.get('type') == 'script.call' for event in events(log))
            else:
                audit_arm = 'betteroffice' if arm == 'scripted-bo' else arm
                violation = access_violation(log, audit_arm, workspace)
                if violation:
                    message = violation + ('; ' + message if not passed else '')
                    passed = False
                calls = tool_calls(log, audit_arm)
            if not calls:
                message = 'agent made no tool calls' + ('; ' + message if not passed else '')
                passed = False
        else:
            message = failure
    seconds = round(time.monotonic() - started, 3)
    return {'task': task['id'], 'format': task['format'], 'run': run or 1,
            'pass': passed, 'passed': passed, 'check_message': message, 'reason': message,
            'tool_calls': calls, 'duration_seconds': seconds, 'seconds': seconds,
            'agent_exit_code': exit_code, 'timed_out': timed_out, 'failure_type': failure_type}


def atomic_write(path, content):
    with tempfile.NamedTemporaryFile(mode='w', dir=path.parent, prefix='.' + path.name, delete=False) as stream:
        staging = Path(stream.name)
        try:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
            os.replace(staging, path)
        finally:
            staging.unlink(missing_ok=True)


def write_results(results, model, records, arm='betteroffice', formats=FORMATS, **metadata):
    atomic_write(results / 'results.json', json.dumps({'model': model, 'arm': arm, 'formats': list(formats),
                 'suite': 'v2', 'runs': 1, 'agent': 'codex', **metadata,
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
        label = record['task'] + (f' / run-{record["run"]}' if metadata.get('run_directories') else '')
        rows.append(f'| {label} | {"PASS" if record["pass"] else "FAIL"} | {message} | {calls} | {record["duration_seconds"]:.3f} |')
    atomic_write(results / 'results.md', '\n'.join(rows) + '\n')


def positive_integer(value):
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError('must be a positive integer')
    return number


def suite_tasks(suite, tier='basic'):
    if suite != 'v3' and tier != 'basic':
        raise ValueError('--tier complex|all requires --suite v3')
    if suite == 'v2':
        return checker.tasks()
    try:
        return checker.tasks(suite=suite) if tier == 'basic' else checker.tasks(suite=suite, tier=tier)
    except TypeError:
        raise ValueError('v3 tasks are unavailable: check.py must support tasks(suite=...)') from None


def resume_records(path, expected, pairs):
    report = json.loads(path.read_text())
    for name, value in expected.items():
        if report.get(name) != value:
            raise ValueError(f'cannot resume: {name} differs from results.json')
    records, seen = [], set()
    pairs = set(pairs)
    for record in report['results']:
        pair = (record['task'], record.get('run', 1))
        if type(pair[1]) is not int or pair[1] < 1:
            raise ValueError('cannot resume: invalid run index')
        if pair not in pairs or pair in seen:
            raise ValueError('cannot resume: unexpected or duplicate task/run in results.json')
        seen.add(pair)
        if isinstance(record.get('pass'), bool) and all(key in record for key in ('check_message', 'tool_calls', 'duration_seconds')):
            records.append(record)
    return records


def main():
    parser = argparse.ArgumentParser(usage='%(prog)s [options] <agents-cli-command> <results-dir> [model]\n'
                                           '       %(prog)s --arm generic|scripted-lib --python <python> [options] <results-dir> [model]')
    parser.add_argument('--arm', choices=(*ALLOWED_ITEMS, 'scripted-lib', 'scripted-bo'), default='betteroffice')
    parser.add_argument('--python', help='generic/scripted-lib arm: a Python that imports openpyxl, pptx and docx')
    parser.add_argument('--suite', choices=('v2', 'v3'), default='v2')
    parser.add_argument('--tier', choices=('basic', 'complex', 'all'), default='basic')
    parser.add_argument('--runs', type=positive_integer, default=None, help='independent runs per task (default: 1)')
    parser.add_argument('--jobs', type=positive_integer, default=1, help='concurrent task/runs (default: 1, maximum: 4)')
    parser.add_argument('--tasks', help='comma-separated task IDs')
    parser.add_argument('--cross-model', action='store_true', help='select only cross_model tasks')
    parser.add_argument('--resume', action='store_true', help='skip finished task/runs in results.json')
    parser.add_argument('--agent', choices=('codex', 'muse'), default='codex')
    parser.add_argument('--formats', type=format_selection, default=FORMATS, metavar='<list>',
                        help=f'comma-separated formats to run, from {",".join(FORMATS)} (default: all)')
    parser.add_argument('arguments', nargs='+', help=argparse.SUPPRESS)
    args = parser.parse_args()
    library = args.arm in ('generic', 'scripted-lib')
    scripted = args.arm.startswith('scripted-')
    names = ('results_dir', 'model') if library else ('agents_cli_command', 'results_dir', 'model')
    if not len(names) - 1 <= len(args.arguments) <= len(names):
        parser.error(f'--arm {args.arm} takes ' + ' '.join(f'<{name}>' for name in names[:-1]) + ' [model]')
    values = dict(zip(names, args.arguments))
    model = 'none' if scripted else values.get('model') or 'gpt-6-luna'
    runs = args.runs or 1
    run_directories = args.runs is not None or args.suite != 'v2'
    try:
        if args.agent == 'muse':
            raise ValueError(MUSE_BLOCKER)
        if args.jobs > 4:
            raise ValueError('--jobs must be at most 4')
        if library != (args.python is not None):
            raise ValueError('--python <python> is required with --arm generic or scripted-lib and accepted only there')
        cli = None if library else cli_command(values['agents_cli_command'])
        python = Path(args.python).expanduser().absolute() if library else None
        results = Path(values['results_dir']).expanduser().resolve()
        if results == REPO or REPO in results.parents:
            raise ValueError('results-dir must be outside the repository')
        if results.exists() and any(results.iterdir()) and not args.resume:
            raise ValueError('results-dir must be new or empty')
        tasks = suite_tasks(args.suite, args.tier)
        selected = {name.strip() for name in args.tasks.split(',')} if args.tasks is not None else None
        if selected is not None and not selected <= {task['id'] for task in tasks}:
            raise ValueError('unknown task IDs: ' + ', '.join(sorted(selected - {task['id'] for task in tasks})))
        tasks = [task for task in tasks if task['format'] in args.formats
                 and (selected is None or task['id'] in selected) and (not args.cross_model or task.get('cross_model') is True)]
        if not tasks:
            raise ValueError('no tasks selected')
        metadata = {'suite': args.suite, 'runs': runs, 'agent': 'none' if scripted else args.agent,
                    'task_ids': [task['id'] for task in tasks], 'cross_model': args.cross_model,
                    'run_directories': run_directories}
        if args.suite == 'v3':
            metadata['tier'] = args.tier
        if library:
            metadata['python'] = str(python)
        else:
            metadata['mcp_command'] = cli
        pairs = [(task['id'], index) for task in tasks for index in range(1, runs + 1)]
        expected = {'model': model, 'arm': args.arm, 'formats': list(args.formats), **metadata}
        records = resume_records(results / 'results.json', expected, pairs) if args.resume else []
        finished = {(record['task'], record.get('run', 1)) for record in records}
        pending = [(task, index) for task in tasks for index in range(1, runs + 1) if (task['id'], index) not in finished]
        if pending and not scripted and shutil.which('codex') is None:
            raise ValueError('codex is not on PATH')
    except (ValueError, OSError, KeyError) as error:
        parser.error(str(error))
    try:
        if pending:
            if library:
                python_ready(python)
                metadata['python'] = str(python)
            else:
                metadata['mcp_startup_seconds'] = mcp_ready(cli, required_tools(tasks))
        elif args.resume:
            previous = json.loads((results / 'results.json').read_text())
            metadata.update({name: previous[name] for name in ('python', 'mcp_startup_seconds') if name in previous})
        results.mkdir(parents=True, exist_ok=True)
    except (RuntimeError, OSError) as error:
        sys.exit(f'run.py: {error}')
    order = {pair: index for index, pair in enumerate(pairs)}

    def save(record=None):
        if record is not None:
            records.append(record)
        records.sort(key=lambda item: order[(item['task'], item.get('run', 1))])
        write_results(results, model, records, args.arm, args.formats, **metadata)
        if record is not None:
            label = record['task'] + (f' / run-{record["run"]}' if run_directories else '')
            print(f'{label}: {"PASS" if record["pass"] else "FAIL"} ({record["duration_seconds"]:.3f}s): {record["check_message"]}', flush=True)

    save()
    processes = AgentProcesses()
    with concurrent.futures.ThreadPoolExecutor(max_workers=args.jobs) as executor:
        futures = []
        try:
            for task, index in pending:
                futures.append(executor.submit(run_task, task, cli, model, results, args.arm, python,
                                               index if run_directories else None, processes=processes))
            for future in concurrent.futures.as_completed(futures):
                save(future.result())
        except BaseException as error:
            for future in futures:
                future.cancel()
            processes.cancel()
            if isinstance(error, KeyboardInterrupt):
                return 130
            raise
    return 0 if all(record['pass'] for record in records) else 1


if __name__ == '__main__':
    raise SystemExit(main())
