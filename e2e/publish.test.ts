import { expect, test } from 'bun:test';
import { advances, planPublication } from './publish';

const sha = 'd'.repeat(40);

function run(format: string, overrides: Record<string, unknown> = {}) {
  return {
    file: `/results/${format}.json`,
    run: {
      schemaVersion: 3,
      format,
      commit: sha,
      dirty: false,
      recordedAt: `2026-09-22T19:48:0${format.length}.000Z`,
      scenarios: [{ scenario: 'editing-session', status: 'passed' }],
      ...overrides,
    },
  };
}

const complete = () => [run('xlsx'), run('docx'), run('pptx')];

test('a complete green run is keyed by commit and the pointer is written last', () => {
  const plan = planPublication(complete(), '2026-09-22T20:00:00.000Z', sha);
  expect(plan.uploads.map((upload) => upload.key)).toEqual([
    `e2e/${sha}/docx.json`,
    `e2e/${sha}/pptx.json`,
    `e2e/${sha}/xlsx.json`,
  ]);
  expect(plan.uploads[0]!.file).toBe('/results/docx.json');
  expect(plan.manifest).toEqual({
    schema_version: 1,
    sha,
    recorded_at_utc: '2026-09-22T19:48:04.000Z',
    published_at_utc: '2026-09-22T20:00:00.000Z',
    run_number: null,
    formats: ['docx', 'pptx', 'xlsx'],
  });
});

test('partial, failed, dirty or mismatched runs are never published', () => {
  const now = '2026-09-22T20:00:00.000Z';
  expect(() => planPublication(complete().slice(1), now)).toThrow('one result file per format');
  expect(() => planPublication([...complete(), run('docx')], now)).toThrow('one result file per format');
  expect(() =>
    planPublication([run('docx', { scenarios: [{ status: 'failed' }] }), run('pptx'), run('xlsx')], now)
  ).toThrow('failed or skipped');
  expect(() => planPublication([run('docx', { scenarios: [] }), run('pptx'), run('xlsx')], now)).toThrow(
    'failed or skipped'
  );
  expect(() => planPublication([run('docx', { dirty: true }), run('pptx'), run('xlsx')], now)).toThrow('dirty');
  expect(() => planPublication([run('docx', { schemaVersion: 2 }), run('pptx'), run('xlsx')], now)).toThrow(
    'schema'
  );
  expect(() =>
    planPublication([run('docx', { commit: 'e'.repeat(40) }), run('pptx'), run('xlsx')], now)
  ).toThrow('different commits');
  expect(() => planPublication(complete(), now, 'f'.repeat(40))).toThrow(`not ${'f'.repeat(40)}`);
});

test('the pointer only moves forward in time', () => {
  const next = planPublication(complete(), '2026-09-22T20:00:00.000Z', sha).manifest;
  expect(advances(null, next)).toBe(true);
  expect(advances({ recorded_at_utc: '2026-09-21T10:00:00.000Z' }, next)).toBe(true);
  expect(advances({ recorded_at_utc: next.recorded_at_utc }, next)).toBe(true);
  expect(advances({ recorded_at_utc: '2026-09-23T10:00:00.000Z' }, next)).toBe(false);
  expect(advances({ sha }, next)).toBe(true);
});

test('a full re-run of an older push keeps its run number and cannot take the pointer back', () => {
  const rerun = planPublication(complete(), '2026-09-23T09:00:00.000Z', sha, 410).manifest;
  expect(rerun.run_number).toBe(410);
  expect(advances({ run_number: 412, recorded_at_utc: '2026-09-22T19:00:00.000Z' }, rerun)).toBe(false);
  expect(advances({ run_number: 410, recorded_at_utc: '2026-09-22T19:00:00.000Z' }, rerun)).toBe(true);
  expect(advances({ run_number: 409 }, rerun)).toBe(true);
});
