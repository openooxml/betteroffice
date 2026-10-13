import importlib.util
import json
import os
import re
import shlex
import subprocess
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


runner = load('format_runner', 'run.py')
validation = load('format_validation', 'validate-checkers.py')
FAKE_CODEX = '''#!/usr/bin/env python3
import json
import os
import shutil
import sys
import time
from pathlib import Path
args = sys.argv[1:]
assert args[:2] == ['exec', '-m']
assert args[args.index('--sandbox') + 1] == 'read-only'
assert '--skip-git-repo-check' in args and '--ignore-user-config' in args
assert '--ignore-rules' in args and '--ephemeral' in args and '--json' in args
assert sys.stdin.read() == ''
workspace = Path(args[args.index('-C') + 1])
assert Path.cwd() == workspace
assert len(list(workspace.iterdir())) == 1
assert workspace.name.startswith('office-format-')
overrides = [args[i + 1] for i, arg in enumerate(args) if arg == '-c']
assert 'approval_policy="never"' in overrides
assert 'mcp_servers.betteroffice.required=true' in overrides
assert 'mcp_servers.betteroffice.startup_timeout_sec=' + os.environ['FAKE_STARTUP_TIMEOUT'] in overrides
command = json.loads(next(value.split('=', 1)[1] for value in overrides if value.startswith('mcp_servers.betteroffice.command=')))
mcp_args = json.loads(next(value.split('=', 1)[1] for value in overrides if value.startswith('mcp_servers.betteroffice.args=')))
assert Path(command).is_absolute() and Path(command).is_file()
assert mcp_args[-2:] == ['--root', str(workspace)]
assert mcp_args[:-2] == ['argument with spaces']
assert args[-1].endswith('Use only the betteroffice MCP tools to read and edit the documents. All other tool calls disqualify the task. The documents live in the current directory.')
tasks = json.loads(Path(os.environ['FAKE_TASKS']).read_text())
task = next(task for task in tasks if args[-1].startswith(task['prompt']))
assert args[2] == os.environ['FAKE_MODEL']
if os.environ.get('FAKE_MODE') == 'hang':
    time.sleep(60)
if os.environ.get('FAKE_MODE') == 'crash':
    sys.exit(9)
if os.environ.get('FAKE_MODE') not in ('missing', 'no-tools'):
    source = workspace / task['input'] if os.environ.get('FAKE_MODE') == 'untouched' else Path(os.environ['FAKE_CORRECT']) / (task['id'] + '.' + task['format'])
    shutil.copyfile(source, workspace / task['output'])
if os.environ.get('FAKE_MODE') == 'change-source':
    (workspace / task['input']).write_bytes(b'changed')
print(json.dumps({'type': 'thread.started'}))
mode = os.environ.get('FAKE_MODE')
for event in [] if mode == 'no-tools' else ['item.started', 'item.completed']:
    print(json.dumps({'type': event, 'item': {'id': 'tool-1', 'type': 'mcp_tool_call', 'server': 'betteroffice'}}))
if mode == 'command':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'shell-1', 'type': 'command_execution', 'command': 'unzip -p input.xlsx'}}))
elif mode == 'other-mcp':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'other-tool', 'type': 'mcp_tool_call', 'server': 'other'}}))
elif mode == 'other-tool':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'browser-1', 'type': 'web_search', 'query': 'file:///input.xlsx'}}))
'''
SERVER_TOOLS = ['office_files', 'office_open', 'office_outline', 'office_grep', 'office_read', 'office_cells', 'office_render',
                'office_verify', 'office_close', 'office_propose', 'office_propose_cells', 'office_review', 'office_accept',
                'office_reject', 'office_export', 'xlsx_outline', 'xlsx_read_range', 'xlsx_preview', 'xlsx_write_range',
                'xlsx_edit_grid', 'xlsx_edit_sheets', 'xlsx_find_replace', 'xlsx_format', 'xlsx_sort',
                'pptx_outline', 'pptx_read_slide', 'pptx_preview', 'pptx_edit']
FAKE_MCP = '''#!/usr/bin/env python3
import json
import os
import sys
from pathlib import Path
args = sys.argv[1:]
assert args[:-2] == ['argument with spaces'] and args[-2] == '--root'
assert Path(args[-1]).is_dir() and not any(Path(args[-1]).iterdir())
mode = os.environ.get('FAKE_MCP_MODE')
if mode == 'exit':
    sys.exit(3)
tools = [name for name in json.loads(os.environ['FAKE_MCP_TOOLS']) if not (mode == 'missing-tool' and name == 'xlsx_sort')]
initialized = False
for line in sys.stdin:
    message = json.loads(line)
    if message['method'] == 'notifications/initialized':
        initialized = True
        continue
    if message['method'] == 'initialize':
        result = {'protocolVersion': message['params']['protocolVersion'], 'capabilities': {'tools': {}},
                  'serverInfo': {'name': 'fake', 'version': '1'}}
    else:
        assert message['method'] == 'tools/list' and initialized
        if mode == 'silent':
            continue
        start = int(message['params'].get('cursor', '0'))
        result = {'tools': [{'name': name, 'inputSchema': {'type': 'object'}} for name in tools[start:start + 10]]}
        if start + 10 < len(tools):
            result['nextCursor'] = str(start + 10)
    print(json.dumps({'jsonrpc': '2.0', 'id': message['id'], 'result': result}), flush=True)
'''
GENERIC_SUFFIX = (' Use shell commands and Python with openpyxl, python-pptx and python-docx to read and edit the documents.'
                  ' There is no network access. All other tool calls disqualify the task. The documents live in the current directory.')
FAKE_GENERIC_CODEX = '''#!/usr/bin/env python3
import json
import os
import shutil
import sys
from pathlib import Path
args = sys.argv[1:]
assert args[:3] == ['exec', '-m', os.environ['FAKE_MODEL']]
assert args[args.index('--sandbox') + 1] == 'workspace-write'
workspace = Path(args[args.index('-C') + 1])
assert Path.cwd() == workspace and len(list(workspace.iterdir())) == 1
assert os.environ['PATH'].split(os.pathsep)[0] == os.environ['FAKE_PYTHON_BIN']
overrides = [args[i + 1] for i, arg in enumerate(args) if arg == '-c']
assert 'sandbox_workspace_write.network_access=false' in overrides
assert not any(value.startswith(('mcp_servers', 'sandbox_workspace_write.writable_roots')) for value in overrides)
assert args[-1].endswith(os.environ['FAKE_GENERIC_SUFFIX'])
task = next(task for task in json.loads(Path(os.environ['FAKE_TASKS']).read_text()) if args[-1].startswith(task['prompt']))
shutil.copyfile(Path(os.environ['FAKE_CORRECT']) / (task['id'] + '.' + task['format']), workspace / task['output'])
print(json.dumps({'type': 'thread.started'}))
for item in [{'id': 'shell-1', 'type': 'command_execution', 'command': 'python3 edit.py'},
             {'id': 'patch-1', 'type': 'file_change', 'changes': [{'path': 'edit.py', 'kind': 'add'}]}]:
    for event in ['item.started', 'item.completed']:
        print(json.dumps({'type': event, 'item': item}))
mode = os.environ.get('FAKE_MODE')
if mode == 'mcp':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'tool-1', 'type': 'mcp_tool_call', 'server': 'betteroffice'}}))
elif mode == 'web':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'web-1', 'type': 'web_search', 'query': 'openpyxl docs'}}))
'''


def executable(path, text):
    path.parent.mkdir(exist_ok=True)
    path.write_text(text)
    path.chmod(0o755)
    return path


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='office-runner-test-')
        self.root = Path(self.temporary.name)
        self.bin = self.root / 'bin with spaces'
        self.bin.mkdir()
        fake = self.bin / 'codex'
        fake.write_text(FAKE_CODEX)
        fake.chmod(0o755)
        cli = self.bin / 'mcp'
        cli.write_text(FAKE_MCP)
        cli.chmod(0o755)
        self.cli = [str(cli), 'argument with spaces']
        self.correct = self.root / 'correct'
        self.correct.mkdir()
        for task in runner.checker.tasks():
            package = validation.Package(runner.checker.fixture(task))
            number = int(task['id'].split('-')[1])
            (validation.correct_xlsx if task['format'] == 'xlsx' else validation.correct_pptx)(package, number)
            package.save(self.correct / f'{task["id"]}.{task["format"]}')
        self.environment = patch.dict(os.environ, {
            'PATH': str(self.bin) + os.pathsep + os.environ['PATH'], 'FAKE_TASKS': str(ROOT / 'tasks.json'),
            'FAKE_CORRECT': str(self.correct), 'FAKE_MODEL': 'test-model', 'FAKE_STARTUP_TIMEOUT': '60',
            'FAKE_MCP_TOOLS': json.dumps(SERVER_TOOLS),
        })
        self.environment.start()
        self.generic_bin = self.root / 'generic bin'
        executable(self.generic_bin / 'codex', FAKE_GENERIC_CODEX)
        self.python = executable(self.root / 'python bin' / 'python',
                                 '#!/bin/sh\n[ "$1" = -c ] && [ "$2" = "import openpyxl, pptx, docx" ] || exit 7\n')
        self.broken_python = executable(self.root / 'broken python' / 'python',
                                        "#!/bin/sh\necho \"ModuleNotFoundError: No module named 'docx'\" >&2\nexit 1\n")

    def generic_environment(self, **values):
        return patch.dict(os.environ, {'PATH': str(self.generic_bin) + os.pathsep + os.environ['PATH'],
                                       'FAKE_PYTHON_BIN': str(self.python.parent), 'FAKE_GENERIC_SUFFIX': GENERIC_SUFFIX, **values})

    def tearDown(self):
        self.environment.stop()
        self.temporary.cleanup()

    def run_one(self, mode):
        results = self.root / mode
        results.mkdir()
        with patch.dict(os.environ, {'FAKE_MODE': mode}):
            return runner.run_task(runner.checker.tasks()[0], self.cli, 'test-model', results)

    def test_full_shell_run_and_reports(self):
        results = self.root / 'results with spaces'
        command = ' '.join(shlex.quote(value) for value in self.cli)
        result = subprocess.run([str(ROOT / 'run.sh'), command, str(results), 'test-model'],
                                stdin=subprocess.DEVNULL, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        report = json.loads((results / 'results.json').read_text())
        self.assertEqual(len(report['results']), 20)
        self.assertEqual(report['timeout_seconds'], 600)
        self.assertGreater(report['mcp_startup_seconds'], 0)
        self.assertEqual(report['arm'], 'betteroffice')
        self.assertTrue(all(item['pass'] and item['tool_calls'] == 1 and item['duration_seconds'] > 0 for item in report['results']))
        self.assertEqual((results / 'results.md').read_text().count('| PASS |'), 20)
        repeat = subprocess.run([str(ROOT / 'run.sh'), command, str(results), 'test-model'], capture_output=True)
        self.assertEqual(repeat.returncode, 2)

    def test_mcp_preflight_requires_a_ready_server_with_the_task_tools(self):
        required = runner.required_tools(runner.checker.tasks())
        self.assertTrue({'office_open', 'office_export', 'xlsx_sort', 'pptx_edit'} <= set(required) <= set(SERVER_TOOLS))
        self.assertGreater(runner.mcp_ready(self.cli, required), 0)
        for mode, reason in [('silent', 'did not answer tools/list within 1 seconds'),
                             ('missing-tool', 'missing required tools: xlsx_sort'), ('exit', 'exited with code 3')]:
            with self.subTest(mode=mode), patch.dict(os.environ, {'FAKE_MCP_MODE': mode}), \
                    patch.object(runner, 'STARTUP_TIMEOUT_SECONDS', 1):
                started = time.monotonic()
                with self.assertRaisesRegex(RuntimeError, 'MCP server not ready: .*' + re.escape(reason)):
                    runner.mcp_ready(self.cli, required)
                self.assertLess(time.monotonic() - started, 5)

    def test_run_stops_before_the_first_task_when_the_server_is_not_ready(self):
        results = self.root / 'not ready'
        command = ' '.join(shlex.quote(value) for value in self.cli)
        with patch.dict(os.environ, {'FAKE_MCP_MODE': 'missing-tool'}):
            result = subprocess.run([str(ROOT / 'run.sh'), command, str(results), 'test-model'],
                                    stdin=subprocess.DEVNULL, capture_output=True, text=True)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn('MCP server not ready: missing required tools: xlsx_sort', result.stderr)
        self.assertFalse(results.exists())

    def test_codex_waits_for_the_required_server(self):
        command = runner.command_line(self.cli, self.root, 'test-model', 'Do the task.')
        overrides = [command[i + 1] for i, arg in enumerate(command) if arg == '-c']
        self.assertIn('mcp_servers.betteroffice.required=true', overrides)
        self.assertIn('mcp_servers.betteroffice.startup_timeout_sec=60', overrides)

    def test_tasks_without_tool_calls_are_reported(self):
        record = self.run_one('no-tools')
        self.assertEqual(record['tool_calls'], 0)
        self.assertFalse(record['pass'])
        runner.write_results(self.root / 'no-tools', 'test-model', [record, self.run_one('missing')])
        report = (self.root / 'no-tools' / 'results.md').read_text()
        self.assertIn('Tasks without tool calls: 1.', report)
        self.assertRegex(report, r'\| xlsx-01-correct-value \| FAIL \| [^|]+ \| 0 \|')

    def test_missing_output_crash_and_untouched_export_fail(self):
        for mode in ['missing', 'crash', 'untouched', 'change-source']:
            record = self.run_one(mode)
            self.assertFalse(record['pass'], mode)
            if mode == 'crash':
                self.assertEqual(record['agent_exit_code'], 9)
                self.assertIn('agent exited with code 9', record['check_message'])
            if mode == 'change-source':
                self.assertIn('input fixture was changed', record['check_message'])

    def test_timeout(self):
        with patch.object(runner, 'TIMEOUT_SECONDS', 0.1):
            record = self.run_one('hang')
        self.assertTrue(record['timed_out'])
        self.assertFalse(record['pass'])
        self.assertIn('timed out', record['check_message'])
        self.assertLess(record['duration_seconds'], 5)

    def test_forbidden_tools_fail_with_correct_output(self):
        for mode, reason in [('command', 'command_execution'), ('other-mcp', 'MCP server other'),
                             ('other-tool', 'web_search')]:
            with self.subTest(mode=mode):
                record = self.run_one(mode)
                self.assertEqual(record['agent_exit_code'], 0)
                runner.checker.check(runner.checker.tasks()[0], self.root / mode / record['task'] / 'output.xlsx')
                self.assertFalse(record['pass'])
                self.assertIn(reason, record['check_message'])
                self.assertEqual(record['tool_calls'], 1)
                runner.write_results(self.root / mode, 'test-model', [record])
                self.assertIn(reason, (self.root / mode / 'results.json').read_text())
                self.assertIn(reason, (self.root / mode / 'results.md').read_text())

    def test_file_access_features_disabled(self):
        command = runner.command_line(self.cli, self.root, 'test-model', 'Do the task.')
        disabled = {command[i + 1] for i, arg in enumerate(command) if arg == '--disable'}
        self.assertTrue({'shell_tool', 'browser_use', 'browser_use_external', 'browser_use_full_cdp_access',
                         'in_app_browser', 'computer_use', 'apps', 'view_image', 'in_app_local_automation',
                         'plugins', 'remote_plugin', 'multi_agent', 'hooks',
                         'skill_search', 'skill_mcp_dependency_install', 'image_generation',
                         'unified_exec_tty', 'shell_snapshot', 'workspace_dependencies',
                         'worktrees'} <= disabled)
        self.assertNotIn('code_mode_host', disabled)
        self.assertIn('--ignore-user-config', command)
        overrides = [command[i + 1] for i, arg in enumerate(command) if arg == '-c']
        self.assertIn('web_search="disabled"', overrides)

    def test_default_model_and_unavailable_tool_count(self):
        command = runner.command_line(self.cli, self.root, 'gpt-6-luna', 'Do the task.')
        self.assertEqual(command[command.index('-m') + 1], 'gpt-6-luna')
        log = self.root / 'plain.log'
        log.write_text('Not a structured Codex event\n[]\nnull\n')
        self.assertIsNone(runner.tool_calls(log))

    def test_generic_arm_command_allows_only_the_sandboxed_shell(self):
        command = runner.command_line(None, self.root, 'test-model', 'Do the task.', arm='generic')
        self.assertEqual(command[command.index('--sandbox') + 1], 'workspace-write')
        self.assertEqual(command[command.index('-C') + 1], str(self.root))
        self.assertFalse({'danger-full-access', '--add-dir', '--dangerously-bypass-approvals-and-sandbox'} & set(command))
        self.assertTrue({'--skip-git-repo-check', '--ignore-user-config', '--ignore-rules', '--ephemeral', '--json'} <= set(command))
        overrides = [command[i + 1] for i, arg in enumerate(command) if arg == '-c']
        self.assertEqual(sorted(overrides), ['approval_policy="never"', 'sandbox_workspace_write.network_access=false',
                                             'web_search="disabled"'])
        enabled = {command[i + 1] for i, arg in enumerate(command) if arg == '--enable'}
        disabled = {command[i + 1] for i, arg in enumerate(command) if arg == '--disable'}
        self.assertEqual(enabled, {'shell_tool', 'unified_exec', 'unified_exec_tty'})
        self.assertEqual(disabled, {'shell_snapshot', 'browser_use', 'browser_use_external', 'browser_use_full_cdp_access',
                                    'in_app_browser', 'computer_use', 'apps', 'view_image', 'in_app_local_automation',
                                    'plugins', 'remote_plugin', 'multi_agent', 'hooks', 'skill_search',
                                    'skill_mcp_dependency_install', 'image_generation', 'workspace_dependencies', 'worktrees'})
        self.assertEqual(command[-1], 'Do the task.' + GENERIC_SUFFIX)

    def test_generic_arm_counts_shell_work_and_fails_mcp_and_web_search(self):
        for mode, reason in [('', None), ('mcp', 'forbidden tool call: MCP server betteroffice'),
                             ('web', 'forbidden tool call: web_search')]:
            results = self.root / f'generic-{mode or "shell"}'
            results.mkdir()
            with self.subTest(mode=mode), self.generic_environment(FAKE_MODE=mode):
                record = runner.run_task(runner.checker.tasks()[0], None, 'test-model', results, arm='generic', python=self.python)
                self.assertEqual(record['tool_calls'], 2)
                self.assertEqual(record['pass'], reason is None, record['check_message'])
                if reason:
                    self.assertIn(reason, record['check_message'])

    def test_generic_shell_run_checks_its_python(self):
        results = self.root / 'generic results'
        with self.generic_environment():
            result = subprocess.run([str(ROOT / 'run.sh'), '--arm', 'generic', '--python', str(self.python), str(results), 'test-model'],
                                    stdin=subprocess.DEVNULL, capture_output=True, text=True)
            broken = subprocess.run([str(ROOT / 'run.sh'), '--arm', 'generic', '--python', str(self.broken_python),
                                     str(self.root / 'broken'), 'test-model'], stdin=subprocess.DEVNULL, capture_output=True, text=True)
            usage = subprocess.run([str(ROOT / 'run.sh'), '--arm', 'generic', str(self.root / 'usage')],
                                   stdin=subprocess.DEVNULL, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        report = json.loads((results / 'results.json').read_text())
        self.assertEqual((report['arm'], report['python']), ('generic', str(self.python)))
        self.assertNotIn('mcp_startup_seconds', report)
        self.assertTrue(all(item['pass'] and item['tool_calls'] == 2 for item in report['results']))
        self.assertIn('Arm: `generic`', (results / 'results.md').read_text())
        self.assertEqual(broken.returncode, 1)
        self.assertIn('cannot import openpyxl, pptx and docx', broken.stderr)
        self.assertFalse((self.root / 'broken').exists())
        self.assertEqual(usage.returncode, 2)
        self.assertIn('--python', usage.stderr)


if __name__ == '__main__':
    unittest.main()
