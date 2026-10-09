import argparse
import hashlib
import importlib.util
import json
import os
import shlex
import shutil
import signal
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TIMEOUT_SECONDS = 600
module_spec = importlib.util.spec_from_file_location('format_check', ROOT / 'check.py')
checker = importlib.util.module_from_spec(module_spec)
module_spec.loader.exec_module(checker)


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


def command_line(cli, workspace, model, prompt):
    return ['codex', 'exec', '-m', model, '--sandbox', 'read-only', '--skip-git-repo-check',
            '--ignore-user-config', '--ignore-rules', '--ephemeral', '--json', '-C', str(workspace),
            '-c', 'approval_policy="never"',
            '-c', 'mcp_servers.betteroffice.command=' + json.dumps(cli[0]),
            '-c', 'mcp_servers.betteroffice.args=' + json.dumps(cli[1:] + ['--root', str(workspace)]),
            prompt + ' The documents live in the current directory. The shell is read-only; read and edit them with the betteroffice MCP tools.']


def tool_calls(log):
    identities, anonymous, available = set(), 0, False
    for line in log.read_text(errors='replace').splitlines():
        try:
            event = json.loads(line)
        except ValueError:
            continue
        if not isinstance(event, dict):
            continue
        if event.get('type') in ('thread.started', 'turn.started', 'turn.completed'):
            available = True
        item = event.get('item', {})
        if not isinstance(item, dict) or item.get('type') != 'mcp_tool_call' or item.get('server') != 'betteroffice':
            continue
        available = True
        if 'id' in item:
            identities.add(item['id'])
        elif event.get('type') == 'item.completed':
            anonymous += 1
    return len(identities) + anonymous if available else None


def stop_process(process):
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    process.wait()


def run_task(task, cli, model, results):
    directory = results / task['id']
    directory.mkdir()
    log, errors = directory / 'codex.jsonl', directory / 'codex.stderr'
    started = time.monotonic()
    exit_code, timed_out, failure = None, False, None
    with tempfile.TemporaryDirectory(prefix='office-format-') as temporary:
        workspace = Path(temporary).resolve()
        source = workspace / task['input']
        shutil.copyfile(ROOT / task['fixture'], source)
        original_hash = hashlib.sha256(source.read_bytes()).digest()
        with log.open('w') as stdout, errors.open('w') as stderr:
            try:
                process = subprocess.Popen(command_line(cli, workspace, model, task['prompt']),
                                           cwd=workspace, stdin=subprocess.DEVNULL, stdout=stdout,
                                           stderr=stderr, start_new_session=True)
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
        output = workspace / task['output']
        if output.is_file():
            shutil.copyfile(output, directory / task['output'])
    return {'task': task['id'], 'format': task['format'], 'pass': passed, 'check_message': message,
            'tool_calls': tool_calls(log), 'duration_seconds': round(time.monotonic() - started, 3),
            'agent_exit_code': exit_code, 'timed_out': timed_out}


def write_results(results, model, records):
    (results / 'results.json').write_text(json.dumps({'model': model, 'timeout_seconds': TIMEOUT_SECONDS,
                                                    'results': records}, indent=2) + '\n')
    passed = sum(record['pass'] for record in records)
    rows = [f'Model: `{model}`. Passed: {passed}/{len(records)}. Tool calls count BetterOffice MCP calls; n/a means unavailable.',
            '', '| Task | Result | Check message | Tool calls | Duration (s) |',
            '| --- | --- | --- | ---: | ---: |']
    for record in records:
        message = record['check_message'].replace('|', '\\|').replace('\n', ' ')
        calls = record['tool_calls'] if record['tool_calls'] is not None else 'n/a'
        rows.append(f'| {record["task"]} | {"PASS" if record["pass"] else "FAIL"} | {message} | {calls} | {record["duration_seconds"]:.3f} |')
    (results / 'results.md').write_text('\n'.join(rows) + '\n')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('agents_cli_command')
    parser.add_argument('results_dir', type=Path)
    parser.add_argument('model', nargs='?', default='gpt-6-luna')
    args = parser.parse_args()
    args.model = args.model or 'gpt-6-luna'
    try:
        cli = cli_command(args.agents_cli_command)
        results = args.results_dir.expanduser().resolve()
        if results == REPO or REPO in results.parents:
            raise ValueError('results-dir must be outside the repository')
        if results.exists() and any(results.iterdir()):
            raise ValueError('results-dir must be new or empty')
        if shutil.which('codex') is None:
            raise ValueError('codex is not on PATH')
        results.mkdir(parents=True, exist_ok=True)
    except (ValueError, OSError) as error:
        parser.error(str(error))
    records = []
    write_results(results, args.model, records)
    for task in checker.tasks():
        record = run_task(task, cli, args.model, results)
        records.append(record)
        write_results(results, args.model, records)
        print(f'{record["task"]}: {"PASS" if record["pass"] else "FAIL"} ({record["duration_seconds"]:.3f}s): {record["check_message"]}', flush=True)
    return 0 if all(record['pass'] for record in records) else 1


if __name__ == '__main__':
    raise SystemExit(main())
