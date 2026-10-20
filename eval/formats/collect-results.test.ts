import { expect, test } from 'bun:test';
import { collectResults, parseArgs, serialize, validateSavedOutput, validateSuite, type Group, type Report } from './collect-results';

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
  const args = parseArgs(['--meta', 'suite.meta.json', '--suite', 'v3', '--runs', 'generic=/one', '--date', '2026-10-13', '--runs', 'betteroffice=/two', '--date', '2026-10-15', '--run-index', '3']);
  expect(args.meta).toBe('suite.meta.json');
  expect(args.groups.map(g => [g.arm, g.date, g.run])).toEqual([['generic', '2026-10-13', 1], ['betteroffice', '2026-10-15', 3]]);
  expect(() => parseArgs(['--suite', 'v2', '--runs', 'generic=/one'])).toThrow();
  expect(() => parseArgs(['--date', '2026-10-15'])).toThrow('must follow');
  expect(() => parseArgs(['--unknown', 'value'])).toThrow('Unknown option');
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
