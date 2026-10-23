import importlib.util
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('regression_runner', ROOT / 'run.py')
runner = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(runner)


class WorkspaceTests(unittest.TestCase):
    def test_agent_paths_cannot_redirect_harness_reads_or_writes(self):
        for target in ('answer.json', 'input.xlsx', 'output.xlsx'):
            for kind in ('symlink', 'directory', 'fifo', 'parent-symlink'):
                with self.subTest(target=target, kind=kind), tempfile.TemporaryDirectory() as temporary:
                    root = Path(temporary)
                    outside = root / 'outside'
                    outside.mkdir()
                    victim = outside / Path(target).name
                    victim.write_bytes(b'host content')
                    fixture = root / 'fixture.xlsx'
                    fixture.write_bytes(b'fixture')
                    name = 'nested/' + target if kind == 'parent-symlink' else target
                    task = {'id': 'safe-path', 'format': 'xlsx', 'input': name if target == 'input.xlsx' else 'input.xlsx',
                            'output': name if target == 'output.xlsx' else 'output.xlsx', 'prompt': 'test'}
                    if target == 'answer.json':
                        task.update(tier='complex', expected={'answer': 4})
                    else:
                        task['expected'] = {}
                    def process(command, cwd, stdout, **kwargs):
                        workspace = Path(cwd)
                        path = workspace / name
                        if kind == 'parent-symlink':
                            path.unlink(missing_ok=True)
                            path.parent.rmdir() if path.parent.exists() else None
                            path.parent.symlink_to(outside, target_is_directory=True)
                        else:
                            path.unlink(missing_ok=True)
                            if kind == 'symlink':
                                path.symlink_to(victim)
                            elif kind == 'fifo':
                                os.mkfifo(path)
                            else:
                                path.mkdir()
                        if target == 'answer.json' and kind == 'parent-symlink':
                            (workspace / 'answer.json').symlink_to(path)
                        for item in ({'id': 'tool', 'type': 'command_execution', 'command': 'true'},
                                     {'id': 'final', 'type': 'agent_message', 'text': '4'}):
                            stdout.write(json.dumps({'type': 'item.completed', 'item': item}) + '\n')
                        return SimpleNamespace(returncode=0, stdin=None, stdout=None, poll=lambda: 0, wait=lambda timeout: 0)
                    with patch.object(runner.checker, 'fixture', return_value=fixture), \
                            patch.object(runner.checker, 'prepare_input', side_effect=lambda task, data: data), \
                            patch.object(runner.checker, 'check'), \
                            patch.object(runner, 'command_line', return_value=['fake']), \
                            patch.object(runner.subprocess, 'Popen', side_effect=process):
                        record = runner.run_task(task, None, 'model', root / 'results', arm='generic')
                    self.assertEqual(victim.read_bytes(), b'host content')
                    self.assertFalse(record['pass'])
                    self.assertEqual(record['failure_type'], 'unsafe-workspace-file')
                    self.assertFalse((root / 'results' / task['id'] / target).exists())


class ConditionalFormatTests(unittest.TestCase):
    def normalized(self, ranges):
        checker = runner.checker
        sheet = ET.Element(checker.q('s', 'worksheet'))
        for priority, (sqref, formula) in enumerate(ranges, 1):
            group = ET.SubElement(sheet, checker.q('s', 'conditionalFormatting'), sqref=sqref)
            rule = ET.SubElement(group, checker.q('s', 'cfRule'), type='expression', priority=str(priority))
            ET.SubElement(rule, checker.q('s', 'formula')).text = formula
        styles = ET.Element(checker.q('s', 'styleSheet'))
        package = SimpleNamespace(root=lambda name: styles if name == 'xl/styles.xml' else sheet,
                                  namelist=lambda: ['xl/styles.xml', 'xl/worksheets/sheet1.xml'])
        with patch.object(checker, 'related', return_value='xl/styles.xml'), \
                patch.object(checker, 'sheet_map', return_value={'Sheet': 'xl/worksheets/sheet1.xml'}):
            checker.normalize_complex_xlsx(package, {})
        return checker.canonical(sheet)

    def test_cx07_split_range_requires_shifted_relative_formula(self):
        task = next(t for t in runner.checker.tasks('v3', 'complex') if t['id'] == 'xlsx-v3-cx-07-green-recovery-sort')
        change = task['expected']['conditional_formats'][0]
        formula = change['formulas'][0]
        self.assertEqual(change['sqref'], 'E4:E30')
        expected = self.normalized([('E4:E30', formula)])
        self.assertNotEqual(expected, self.normalized([('E4:E17', formula), ('E18:E30', formula)]))
        shifted = runner.checker.translate_formula(formula, 'E4', 'E18')
        self.assertEqual(expected, self.normalized([('E4:E17', formula), ('E18:E30', shifted)]))

    def test_mixed_absolute_references_and_quoted_text(self):
        formula = 'AND(E4=$A$1,$B4=E$2,C3="E4")'
        expected = self.normalized([('E4:F30', formula)])
        shifted = runner.checker.translate_formula(formula, 'E4', 'F18')
        self.assertEqual(expected, self.normalized([('E4:F17', formula), ('E18:E30', runner.checker.translate_formula(formula, 'E4', 'E18')), ('F18:F30', shifted)]))
        self.assertNotEqual(expected, self.normalized([('E4:F17', formula), ('E18:F30', formula)]))

    def test_reference_before_range_origin_is_valid(self):
        self.assertEqual(self.normalized([('E4:E30', 'A1>0')]),
                         self.normalized([('E4:E17', 'A1>0'), ('E18:E30', 'A15>0')]))


class InterruptTests(unittest.TestCase):
    def test_sigint_cancels_queued_runs_and_kills_agent_groups(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = root / 'fixture.xlsx'
            fixture.write_bytes(b'fixture')
            marker = root / 'started'
            worker = root / 'worker.py'
            worker.write_text('import os, subprocess, sys, time\nfrom pathlib import Path\n'
                              'child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])\n'
                              f'with open({str(marker)!r}, "a") as stream: stream.write(str(os.getpid()) + " " + str(child.pid) + "\\n")\n'
                              'time.sleep(60)\n')
            harness = root / 'harness.py'
            harness.write_text(f'''import importlib.util, sys
from pathlib import Path
spec = importlib.util.spec_from_file_location('runner', {str(ROOT / 'run.py')!r})
r = importlib.util.module_from_spec(spec)
spec.loader.exec_module(r)
r.checker.fixture = lambda task: Path({str(fixture)!r})
r.checker.prepare_input = lambda task, data: data
r.checker.check = lambda task, output: None
r.python_ready = lambda python: None
r.shutil.which = lambda command: sys.executable
r.command_line = lambda *args: [sys.executable, {str(worker)!r}]
task = {{'id': 'interrupt', 'format': 'xlsx', 'input': 'input.xlsx', 'output': 'output.xlsx', 'expected': {{}}, 'prompt': 'test'}}
r.suite_tasks = lambda *args: [dict(task, id='interrupt-' + str(i)) for i in range(8)]
sys.argv = ['run.py', '--arm', 'generic', '--python', sys.executable, '--jobs', '2', {str(root / 'results')!r}]
sys.exit(r.main())
''')
            process = subprocess.Popen([sys.executable, str(harness)], start_new_session=True,
                                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            completed = False
            pids = []
            try:
                deadline = time.monotonic() + 10
                while time.monotonic() < deadline:
                    if marker.exists() and len(marker.read_text().splitlines()) == 2:
                        break
                    time.sleep(0.02)
                self.assertTrue(marker.exists())
                self.assertEqual(len(marker.read_text().splitlines()), 2)
                pids = [int(pid) for pid in marker.read_text().split()]
                process.send_signal(signal.SIGINT)
                try:
                    process.wait(timeout=5)
                    completed = True
                except subprocess.TimeoutExpired:
                    pass
                self.assertTrue(completed, 'interrupted batch waited for queued or running agents')
                self.assertNotEqual(process.returncode, 0)
                self.assertEqual(len(marker.read_text().splitlines()), 2)
                for pid in pids:
                    status = Path(f'/proc/{pid}/stat')
                    self.assertTrue(not status.exists() or status.read_text().split()[2] == 'Z', f'agent descendant {pid} survived')
            finally:
                if process.poll() is None:
                    os.killpg(process.pid, signal.SIGKILL)
                for pid in pids:
                    try:
                        os.kill(pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                process.wait()


if __name__ == '__main__':
    unittest.main()
