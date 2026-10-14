import { expect, test } from 'bun:test';
import { FRAME_MS, e2eUrl, groupName, parseLatestE2e, percentile, summarizeLatency } from './latency';

const sha = 'a'.repeat(40);

function run(format: string, scenarios: unknown[], overrides: Record<string, unknown> = {}) {
  return {
    schemaVersion: 3,
    format,
    commit: sha,
    dirty: false,
    recordedAt: '2026-09-22T19:48:16.248Z',
    environment: { platform: 'linux', arch: 'x64', cpu: 'AMD EPYC 7763', cpus: 4 },
    scenarios,
    ...overrides,
  };
}

const scenario = (name: string, participants: string[], ops: [string, number, string?][]) => ({
  scenario: name,
  sample: 'demo',
  status: 'passed',
  description: `${name} description`,
  participants,
  ops: ops.map(([op, e2eMs, error]) => ({ op, e2eMs, ...(error ? { error } : {}) })),
});

test('percentiles use the harness nearest-rank rule', () => {
  const values = Array.from({ length: 20 }, (_, index) => index + 1);
  expect(percentile(values, 50)).toBe(10);
  expect(percentile(values, 95)).toBe(19);
  expect(percentile([7], 95)).toBe(7);
  expect(percentile([], 50)).toBe(0);
});

test('operation variants group under their verb', () => {
  expect(groupName('applyInput:keystroke')).toBe('applyInput');
  expect(groupName('open')).toBe('open');
});

test('formats aggregate ops, frame budget, collaboration and cross-SDK cases', () => {
  const latency = summarizeLatency([
    run('docx', [
      scenario('typing-burst', ['web'], [
        ['open', 40],
        ['applyInput:keystroke', 10],
        ['applyInput:keystroke', 30],
        ['save', 12, 'disk full'],
      ]),
      scenario('two-editors-converge', ['web:a', 'web:b'], [['applyUpdate', 0.2]]),
    ]),
    run('xlsx', [scenario('python-roundtrip', ['web', 'python'], [['python:set', 20]])], {
      recordedAt: '2026-09-22T19:50:00.000Z',
    }),
  ]);
  const docx = latency.formats.docx!;
  expect(docx.ops).toBe(5);
  expect(docx.errors).toBe(1);
  expect(docx.withinFrame).toBe(3);
  expect(docx.values).toEqual([0.2, 10, 12, 30, 40]);
  expect(docx.groups[0]).toEqual({ name: 'applyInput', count: 2, p50: 10, p95: 30, max: 30 });
  expect(docx.cases[0]!.totalMs).toBe(92);
  expect(latency.formats.pptx).toBeUndefined();
  expect(latency.cases).toBe(3);
  expect(latency.passed).toBe(3);
  expect(latency.ops).toBe(6);
  expect([latency.p50, latency.p95]).toEqual([12, 40]);
  expect(latency.multiEditor).toBe(2);
  expect(latency.crossSdk).toBe(1);
  expect(latency.recordedAt).toBe('2026-09-22T19:50:00.000Z');
  expect(latency.environment?.cpus).toBe(4);
  expect(FRAME_MS).toBeCloseTo(16.667, 3);
});

test('results from different commits or schemas are refused', () => {
  expect(() => summarizeLatency([run('docx', []), run('pptx', [], { commit: 'b'.repeat(40) })])).toThrow(
    'disagree'
  );
  expect(() => summarizeLatency([run('docx', [], { schemaVersion: 2 })])).toThrow('Malformed');
  expect(() => summarizeLatency([run('docx', []), run('docx', [])])).toThrow('Malformed');
  expect(() => summarizeLatency([])).toThrow('Malformed');
});

test('the pointer names a commit and at least one format', () => {
  expect(parseLatestE2e({ sha, formats: ['xlsx', 'docx', 'vsdx'], published_at_utc: 'now' })).toEqual({
    sha,
    publishedAt: 'now',
    formats: ['docx', 'xlsx'],
  });
  expect(() => parseLatestE2e({ sha, formats: [] })).toThrow('Malformed');
  expect(() => parseLatestE2e({ sha: 'short', formats: ['docx'] })).toThrow('Malformed');
  expect(e2eUrl(sha, 'pptx')).toBe(`/e2e/${sha}/pptx.json`);
  expect(() => e2eUrl('short', 'pptx')).toThrow('SHA');
});
