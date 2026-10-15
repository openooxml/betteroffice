import { expect, test } from 'bun:test';
import { collectResults, defaultArms, parseArgs, serialize, type Group, type Report } from './collect-results';

const tasks = [{ id: 'docx-example', format: 'docx', title: 'Edit text', prompt: 'Change text.' }];
const fixtures = { kind: 'real-world', count: 1, url: 'https://example.org/', metadata: 'https://example.org/metadata.json' };
const row = { task: tasks[0]!.id, pass: true, check_message: 'verified', tool_calls: 2, duration_seconds: 1.5 };
const group = (arm = 'generic', run = 1, model = 'model-a'): Group & { report: Report } => ({
  arm, run, date: '2026-10-15', directory: '/private/runner', report: { model, results: [{ ...row }] },
});
const options = () => ({ suite: 'v3', tasks, arms: defaultArms, fixtures, reasons: [], groups: [group()] });

test('three runs, third arm, and second model subset survive collection without runner metadata', () => {
  const arms = [...defaultArms, { id: 'third', label: 'Third arm', summary: 'Other tools.', tool_calls: 'Calls.' }];
  const groups = arms.flatMap(a => [1, 2, 3].map(i => group(a.id, i)));
  groups.push(group('third', 1, 'model-b'));
  const suite = collectResults({ ...options(), arms, groups });
  expect(suite.runs).toHaveLength(10);
  expect(suite.arms).toHaveLength(3);
  expect(suite.fixtures.kind).toBe('real-world');
  expect(suite.runs.filter(r => r.model === 'model-b')).toHaveLength(1);
  expect(serialize(suite)).not.toContain('/private/runner');
  expect(serialize(suite)).toBe(serialize(collectResults({ ...options(), arms, groups: groups.toReversed() })));
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
  const args = parseArgs(['--suite', 'v3', '--runs', 'generic=/one', '--date', '2026-10-13', '--runs', 'betteroffice=/two', '--date', '2026-10-15', '--run-index', '3']);
  expect(args.groups.map(g => [g.arm, g.date, g.run])).toEqual([['generic', '2026-10-13', 1], ['betteroffice', '2026-10-15', 3]]);
  expect(() => parseArgs(['--suite', 'v2', '--runs', 'generic=/one'])).toThrow();
  expect(() => parseArgs(['--date', '2026-10-15'])).toThrow('must follow');
  expect(() => parseArgs(['--unknown', 'value'])).toThrow('Unknown option');
});
