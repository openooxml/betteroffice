import { expect, spyOn, test } from 'bun:test';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { collectResults, main, parseArgs, savedOutputPath, serialize, validateSavedOutput, validateSuite, type Group, type Report, type Task } from './collect-results';
import basicTasks from './v3/tasks.json';
import complexTasks from './v3/complex.json';

const arms = [
  { id: 'generic', label: 'Generic tools', summary: 'Generic summary.', details: ['Generic sandbox.'], tool_calls: 'Shell calls.' },
  { id: 'betteroffice', label: 'BetterOffice', summary: 'MCP summary.', details: ['MCP sandbox.'], tool_calls: 'MCP calls.' },
];
const tasks = [
  { id: 'docx-example', format: 'docx', title: 'Edit text', prompt: 'Change text.', fixture: 'example.docx', bytes: 100, sha256: 'a'.repeat(64) },
  { id: 'xlsx-example', format: 'xlsx', title: 'Edit cell', prompt: 'Change a cell.', fixture: 'example.xlsx', bytes: 200, sha256: 'b'.repeat(64) },
];
const fixtures = { kind: 'real-world', count: 2, url: 'https://example.org/', metadata: 'https://example.org/metadata.json', description: ['Real documents from a pinned collection.'] };
const row = { task: tasks[0]!.id, pass: true, check_message: 'verified', tool_calls: 2, duration_seconds: 1.5 };
const group = (arm = 'generic', run = 1, model = 'model-a'): Group & { report: Report } => ({
  arm, run, date: '2026-10-15', directory: '/private/runner', report: { model, results: [{ ...row }] },
});
const options = () => ({ suite: 'v3', tasks, arms, fixtures, harness: 'custom harness', limits: ['Limited coverage.'], reproduce: ['run-custom-suite'], reasons: [], groups: [group()] });

test('three runs, third arm, and second model subset survive collection without runner metadata', () => {
  const allArms = [...arms, { id: 'third', label: 'Third arm', summary: 'Other tools.', details: ['Other sandbox.'], tool_calls: 'Calls.' }];
  const groups = allArms.flatMap(a => [1, 2, 3].map(i => {
    const g = group(a.id, i);
    g.report.results = tasks.map(t => ({ ...row, task: t.id }));
    return g;
  }));
  groups.push(group('third', 1, 'model-b'));
  const suite = collectResults({ ...options(), arms: allArms, groups });
  expect(suite.runs).toHaveLength(19);
  expect(suite.tasks).toHaveLength(2);
  expect(suite.arms).toHaveLength(3);
  expect(suite.fixtures.kind).toBe('real-world');
  expect(suite.runs.filter(r => r.model === 'model-b')).toHaveLength(1);
  expect(serialize(suite)).not.toContain('/private/runner');
  expect(serialize(suite)).toBe(serialize(collectResults({ ...options(), arms: allArms, groups: groups.toReversed() })));
});

test('reject unknown tasks, arms, duplicate identities, invalid metrics, and missing v2 reasons', () => {
  const unknown = group(); unknown.report.results[0]!.task = 'unknown';
  expect(() => collectResults({ ...options(), groups: [unknown] })).toThrow('Unknown task');
  expect(() => collectResults({ ...options(), groups: [group('missing')] })).toThrow('Unknown arm');
  expect(() => collectResults({ ...options(), groups: [group(), group()] })).toThrow('Duplicate run');
  const invalid = group(); invalid.report.results[0]!.duration_seconds = NaN;
  expect(() => collectResults({ ...options(), groups: [invalid] })).toThrow('Invalid result');
  const failed = group(); failed.report.results[0]!.pass = false;
  expect(() => collectResults({ ...options(), suite: 'v2', groups: [failed] })).toThrow('without a reason');
});

test('failure annotations are required and passes always have null annotations', () => {
  const failed = group(); failed.report.results[0]!.pass = false;
  const reason = { model: 'model-a', arm: 'generic', task: row.task, run: 1, reason: 'The model removed formatting.', cause: 'model' as const };
  expect(collectResults({ ...options(), suite: 'v2', reasons: [reason], groups: [failed] }).runs[0]!.reason).toBe(reason.reason);
  expect(collectResults({ ...options(), reasons: [reason] }).runs[0]!.cause).toBeNull();
  expect(() => collectResults({ ...options(), reasons: [reason, reason] })).toThrow('Duplicate reason');
});

test('CLI flags belong to the preceding group', () => {
  const args = parseArgs(['--keep-invalid', '--meta', 'suite.meta.json', '--suite', 'v3', '--runs', 'generic=/one', '--date', '2026-10-13', '--runs', 'betteroffice=/two', '--date', '2026-10-15', '--run-index', '3']);
  expect(args.meta).toBe('suite.meta.json');
  expect(args.keepInvalid).toBe(true);
  expect(args.tasks).toBe(resolve(import.meta.dir, 'v3/tasks.json'));
  expect(args.fixtures).toBe(resolve(import.meta.dir, 'v3/fixtures.json'));
  expect(args.groups.map(g => [g.arm, g.date, g.run])).toEqual([['generic', '2026-10-13', 1], ['betteroffice', '2026-10-15', 3]]);
  const previous = parseArgs(['--meta', 'suite.meta.json', '--suite', 'v2', '--runs', 'generic=/one', '--date', '2026-10-15', '--run-index', '3']);
  expect(previous.tasks).toBe(resolve(import.meta.dir, 'tasks.json'));
  expect(previous.fixtures).toBe(resolve(import.meta.dir, 'fixtures.json'));
  expect(previous.keepInvalid).toBe(false);
  expect(previous.groups[0]!.run).toBe(3);
  expect(() => parseArgs(['--suite', 'v2', '--runs', 'generic=/one'])).toThrow();
  expect(() => parseArgs(['--date', '2026-10-15'])).toThrow('must follow');
  expect(() => parseArgs(['--unknown', 'value'])).toThrow('Unknown option');
});

test('record run indexes override group indexes and task tiers survive collection', () => {
  const tiered: Task[] = tasks.map((t, i) => ({ ...t, tier: i ? 'complex' : 'basic' }));
  const g = group('generic', 9);
  g.report.results = tiered.flatMap(t => [1, 2, 3].map(run => ({ ...row, task: t.id, run })));
  const suite = collectResults({ ...options(), tasks: tiered, groups: [g] });
  expect(suite.tasks.map(t => t.tier)).toEqual(['basic', 'complex']);
  expect(suite.runs.map(r => r.run)).toEqual([1, 2, 3, 1, 2, 3]);
  g.report.results[0]!.run = 0;
  expect(() => collectResults({ ...options(), groups: [g] })).toThrow('Invalid run index');
  const invalid = structuredClone(suite);
  invalid.tasks[0]!.tier = 'unknown' as never;
  expect(() => validateSuite(invalid)).toThrow('tasks.tier');
});

test('invalid workflows require keep-invalid, retain runner reasons, and cannot pass', () => {
  for (const [fields, cause] of [
    [{ timed_out: true, agent_exit_code: -9, tool_calls: 0 }, 'timeout'],
    [{ agent_exit_code: 1, tool_calls: 0 }, 'agent-exit'],
    [{ agent_exit_code: null }, 'agent-exit'],
    [{ tool_calls: 0 }, 'no-tool-calls'],
  ] as const) {
    const g = group();
    Object.assign(g.report.results[0]!, { pass: false, check_message: 'recorded runner failure', ...fields });
    expect(() => collectResults({ ...options(), groups: [g] })).toThrow('valid workflow');
    const run = collectResults({ ...options(), keepInvalid: true, groups: [g] }).runs[0]!;
    expect([run.pass, run.cause, run.reason]).toEqual([false, cause, 'recorded runner failure']);
    g.report.results[0]!.pass = true;
    expect(() => collectResults({ ...options(), keepInvalid: true, groups: [g] })).toThrow('valid workflow');
  }
  const g = group();
  g.report.results[0]!.pass = false;
  expect(collectResults({ ...options(), groups: [g] }).runs[0]!.reason).toBe(row.check_message);
});

test('saved paths follow run_directories and complex answer artifacts', () => {
  const task = { ...tasks[0]!, output: 'output.docx' };
  const report = group().report;
  expect(savedOutputPath('/out', task, report, row, 'v2')).toBe('/out/docx-example/output.docx');
  expect(savedOutputPath('/out', task, { ...report, run_directories: false }, { ...row, run: 1 }, 'v2')).toBe('/out/docx-example/output.docx');
  expect(savedOutputPath('/out', task, { ...report, run_directories: true }, { ...row, run: 3 }, 'v2')).toBe('/out/docx-example/run-3/output.docx');
  expect(savedOutputPath('/out', task, { ...report, run_directories: true }, row, 'v2', 3)).toBe('/out/docx-example/run-3/output.docx');
  expect(savedOutputPath('/out', { ...task, tier: 'complex', expected: { answer: 42 } }, report, { ...row, run: 2 }, 'v3')).toBe('/out/docx-example/run-2/answer.json');
});

test('v3 results.json contains three runs and rechecks per-run artifacts with suite v3', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'collect-v3-'));
  const answer = complexTasks.find(t => 'answer' in t.expected)!;
  const selected = [basicTasks[0]!, complexTasks[0]!, answer];
  const report: Report = { model: 'gpt-6-luna', run_directories: true, results: selected.flatMap(t => [1, 2, 3].map(run => ({ ...row, task: t.id, run }))) };
  const checks: string[][] = [];
  const checker = spyOn(Bun, 'spawnSync').mockImplementation(command => {
    const args = command as string[];
    checks.push(args);
    return { exitCode: 0, stdout: Buffer.from(`PASS: ${args.at(-2)}: rescored`), stderr: Buffer.from('') } as ReturnType<typeof Bun.spawnSync>;
  });
  try {
    for (const t of selected) for (const run of [1, 2, 3]) {
      const path = join(directory, t.id, `run-${run}`);
      await mkdir(path, { recursive: true });
      await writeFile(join(path, t === answer ? 'answer.json' : t.output), '{}');
    }
    const meta = join(directory, 'meta.json'), output = join(directory, 'suite.json');
    await writeFile(meta, JSON.stringify(options()));
    await writeFile(join(directory, 'results.json'), JSON.stringify(report));
    const args = ['--suite', 'v3', '--meta', meta, '--runs', `generic=${directory}`, '--date', '2026-10-15', '--output', output];
    await main(args);
    const suite = JSON.parse(await readFile(output, 'utf8'));
    expect(suite.tasks.map((t: Task) => t.tier).sort()).toEqual(['basic', 'complex', 'complex']);
    expect(suite.runs).toHaveLength(9);
    expect(checks).toEqual(selected.flatMap(t => [1, 2, 3].map(run => ['python3', resolve(import.meta.dir, 'check.py'), '--suite', 'v3', t.id, join(directory, t.id, `run-${run}`, t === answer ? 'answer.json' : t.output)])));
    Object.assign(report.results[0]!, { pass: false, timed_out: true, check_message: 'agent timed out' });
    await writeFile(join(directory, 'results.json'), JSON.stringify(report));
    await expect(main(args)).rejects.toThrow('valid workflow');
    checks.length = 0;
    await main([...args, '--keep-invalid']);
    expect(checks).toHaveLength(8);
    const invalid = JSON.parse(await readFile(output, 'utf8')).runs.find((r: { task: string; run: number }) => r.task === selected[0]!.id && r.run === 1);
    expect([invalid.pass, invalid.cause, invalid.reason]).toEqual([false, 'timeout', 'agent timed out']);
  } finally {
    checker.mockRestore();
    await rm(directory, { recursive: true, force: true });
  }
});

test('suite descriptions, pins, harness and optional validation come from collector inputs', () => {
  const suite = collectResults(options());
  expect(suite.harness).toBe('custom harness');
  expect(suite.fixtures.description).toEqual(fixtures.description);
  expect(suite.limits).toEqual(['Limited coverage.']);
  expect(suite.reproduce).toEqual(['run-custom-suite']);
  expect(suite.arms[0]!.details).toEqual(['Generic sandbox.']);
  expect(suite.tasks[0]).toEqual(tasks[0]);
  expect(suite).not.toHaveProperty('checker_validation');
  const checker_validation = { accepted: 4, untouched_rejected: 2, corrupted_rejected: 6 };
  expect(collectResults({ ...options(), checker_validation }).checker_validation).toEqual(checker_validation);
});

test('missing required suite fields fail loudly in collection and build validation', () => {
  for (const field of ['harness', 'limits', 'reproduce'] as const) {
    const input = options();
    delete (input as Partial<typeof input>)[field];
    expect(() => collectResults(input)).toThrow(field);
    const suite = collectResults(options());
    delete (suite as Partial<typeof suite>)[field];
    expect(() => validateSuite(suite)).toThrow(field);
  }
  for (const mutate of [
    (s: ReturnType<typeof options>) => { delete (s.fixtures as Partial<typeof fixtures>).description; },
    (s: ReturnType<typeof options>) => { s.arms = [{ ...arms[0]!, details: undefined as never }]; },
    (s: ReturnType<typeof options>) => { s.tasks = [{ ...tasks[0]!, fixture: undefined as never }]; },
    (s: ReturnType<typeof options>) => { s.tasks = [{ ...tasks[0]!, bytes: undefined as never }]; },
    (s: ReturnType<typeof options>) => { s.tasks = [{ ...tasks[0]!, sha256: undefined as never }]; },
  ]) {
    const input = structuredClone(options());
    mutate(input);
    expect(() => collectResults(input)).toThrow();
  }
  expect(() => parseArgs(['--suite', 'v3', '--runs', 'generic=/one', '--date', '2026-10-15'])).toThrow('--meta');
});

test('saved outputs must agree with runner verdicts in both directions', () => {
  const check = (pass: boolean) => ({ exitCode: pass ? 0 : 1, stdout: (pass ? 'PASS' : 'FAIL') + ': ' + row.task + ': rescored', stderr: '' });
  expect(validateSavedOutput(row, check(true))).toEqual({ pass: true, check_message: 'rescored' });
  expect(validateSavedOutput({ ...row, pass: false }, check(false))).toEqual({ pass: false, check_message: 'rescored' });
  expect(() => validateSavedOutput(row, check(false))).toThrow('now fails');
  expect(() => validateSavedOutput({ ...row, pass: false }, check(true))).toThrow('now passes');
  expect(() => validateSavedOutput({ ...row, pass: false }, { exitCode: 1, stdout: '', stderr: 'curl: 403' })).toThrow('could not score');
});

test('v3 custom task and fixture inputs reach the checker', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'collect-custom-v3-'));
  const customTasks = join(directory, 'tasks.json'), customFixtures = join(directory, 'fixtures.json');
  const checker = spyOn(Bun, 'spawnSync').mockImplementation((command, options) => {
    expect(options?.env?.FORMAT_EVAL_TASKS).toBe(customTasks);
    expect(options?.env?.FORMAT_EVAL_MANIFEST).toBe(customFixtures);
    const args = command as string[];
    return { exitCode: 0, stdout: Buffer.from(`PASS: ${args.at(-2)}: custom checks`), stderr: Buffer.from('') } as ReturnType<typeof Bun.spawnSync>;
  });
  try {
    const pins = Object.fromEntries([...basicTasks, ...complexTasks].map(t => [t.fixture, { bytes: 1, sha256: 'a'.repeat(64) }]));
    await writeFile(customTasks, JSON.stringify([{ ...basicTasks[0]!, prompt: 'Custom task.', expected: { answer: 42 } }]));
    await writeFile(customFixtures, JSON.stringify({ origin: 'https://example.org/custom', fixtures: pins }));
    await writeFile(join(directory, 'meta.json'), JSON.stringify(options()));
    await writeFile(join(directory, 'results.json'), JSON.stringify({ model: 'model-a', results: [{ ...row, task: basicTasks[0]!.id }] }));
    await main(['--suite', 'v3', '--tasks', customTasks, '--fixtures', customFixtures, '--meta', join(directory, 'meta.json'),
      '--runs', `generic=${directory}`, '--date', '2026-10-15', '--output', join(directory, 'suite.json')]);
    expect(checker).toHaveBeenCalledTimes(1);
  } finally {
    checker.mockRestore();
    await rm(directory, { recursive: true, force: true });
  }
});
