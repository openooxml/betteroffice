import importlib.util
import json
import os
import shlex
import subprocess
import tempfile
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
if os.environ.get('FAKE_MODE') != 'missing':
    source = workspace / task['input'] if os.environ.get('FAKE_MODE') == 'untouched' else Path(os.environ['FAKE_CORRECT']) / (task['id'] + '.' + task['format'])
    shutil.copyfile(source, workspace / task['output'])
if os.environ.get('FAKE_MODE') == 'change-source':
    (workspace / task['input']).write_bytes(b'changed')
print(json.dumps({'type': 'thread.started'}))
for event in ['item.started', 'item.completed']:
    print(json.dumps({'type': event, 'item': {'id': 'tool-1', 'type': 'mcp_tool_call', 'server': 'betteroffice'}}))
mode = os.environ.get('FAKE_MODE')
if mode == 'command':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'shell-1', 'type': 'command_execution', 'command': 'unzip -p input.xlsx'}}))
elif mode == 'other-mcp':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'other-tool', 'type': 'mcp_tool_call', 'server': 'other'}}))
elif mode == 'other-tool':
    print(json.dumps({'type': 'item.completed', 'item': {'id': 'browser-1', 'type': 'web_search', 'query': 'file:///input.xlsx'}}))
'''


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
        cli.write_text('#!/bin/sh\nexit 0\n')
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
            'FAKE_CORRECT': str(self.correct), 'FAKE_MODEL': 'test-model',
        })
        self.environment.start()

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
        self.assertTrue(all(item['pass'] and item['tool_calls'] == 1 and item['duration_seconds'] > 0 for item in report['results']))
        self.assertEqual((results / 'results.md').read_text().count('| PASS |'), 20)
        repeat = subprocess.run([str(ROOT / 'run.sh'), command, str(results), 'test-model'], capture_output=True)
        self.assertEqual(repeat.returncode, 2)

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


if __name__ == '__main__':
    unittest.main()
