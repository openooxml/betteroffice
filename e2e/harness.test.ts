import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  ScenarioRecorder,
  parseRecordedRun,
  regressions,
  summarize,
} from './harness';
import { recorded, scenario } from './test-fixtures';
import { FORMATS, promoteBaseline } from './run';

const ops = (actor: string, e2eMs: number, count = 5) =>
  Array.from({ length: count }, () => ({ op: 'edit', actor, e2eMs }));

const recorder = () => new ScenarioRecorder('xlsx', scenario());

describe('comparison contracts', () => {
  test('detects a slowdown hidden by another actor', () => {
    const before = recorded(scenario([...ops('a', 1), ...ops('b', 2)]));
    const after = recorded(scenario([...ops('a', 1), ...ops('b', 200)]));
    expect(regressions(before, after)).toHaveLength(1);
    expect(regressions(before, after)[0]).toContain(':b:edit');
  });

  test('detects a tail slowdown when enough samples exist', () => {
    const before = recorded(scenario(ops('a', 1, 20)));
    const after = recorded(scenario([...ops('a', 1, 18), ...ops('a', 100, 2)]));
    expect(regressions(before, after)[0]).toContain('p95Ms');
  });

  test('rejects fewer operations, changed actors, and missing scenarios', () => {
    const before = recorded(scenario(ops('a', 5, 10)));
    expect(() =>
      regressions(before, recorded(scenario(ops('a', 5, 1))))
    ).toThrow('count');
    expect(() =>
      regressions(before, recorded(scenario(ops('b', 5, 10))))
    ).toThrow('inventory');
    expect(() => regressions(before, { ...before, scenarios: [] })).toThrow(
      'no scenarios'
    );
  });

  test('rejects incomplete runs, changed corpus bytes, and incompatible machines', () => {
    const before = recorded();
    expect(() =>
      regressions(
        before,
        recorded({ ...scenario(), status: 'skipped', reason: 'Python missing' })
      )
    ).toThrow('incomplete');
    expect(() =>
      regressions(
        before,
        recorded({ ...scenario(), status: 'failed', reason: 'assertion' })
      )
    ).toThrow('incomplete');
    expect(() =>
      regressions(
        before,
        recorded({ ...scenario(), sampleSha256: 'b'.repeat(64) })
      )
    ).toThrow('inventory');
    expect(() =>
      regressions(before, {
        ...before,
        environment: { ...before.environment, cpu: 'another machine' },
      })
    ).toThrow('environment');
  });

  test('rejects malformed and stale result schemas', () => {
    const before = recorded();
    expect(parseRecordedRun(JSON.parse(JSON.stringify(before)))).toEqual(
      before
    );
    expect(() => parseRecordedRun({ ...before, schemaVersion: 2 })).toThrow(
      'schemaVersion'
    );
    expect(() =>
      parseRecordedRun({ ...before, scenarios: [scenario(), scenario()] })
    ).toThrow('duplicate');
    expect(() => parseRecordedRun(recorded(scenario(ops('a', -1))))).toThrow(
      'timing'
    );
    const wrong = scenario();
    wrong.summary.opCount = 99;
    expect(() => parseRecordedRun(recorded(wrong))).toThrow('summary');
  });
});

test('records both synchronous and asynchronous operation failures', async () => {
  const r = recorder();
  expect(() =>
    r.op('edit', () => {
      throw new Error('sync failure');
    })
  ).toThrow('sync failure');
  await expect(
    r.opAsync('save', async () => {
      throw new Error('async failure');
    })
  ).rejects.toThrow('async failure');
  expect(r.ops.map((op) => op.error)).toEqual([
    expect.stringContaining('sync failure'),
    expect.stringContaining('async failure'),
  ]);
  expect(r.ops.every((op) => op.e2eMs >= 0)).toBe(true);
});

test('averages each stage over the operations that report that stage', () => {
  const stats = summarize([
    { op: 'edit', e2eMs: 4, internal: { apply: 2 } },
    { op: 'edit', e2eMs: 8, internal: { apply: 4, recalc: 3 } },
  ]);
  expect(stats.byOp.edit.stagesMs).toEqual({ apply: 3, recalc: 3 });
});

test('promotes all formats together and preserves the baseline on failed or partial runs', () => {
  const root = mkdtempSync(join(tmpdir(), 'e2e-baseline-test-'));
  const destination = join(root, 'baseline');
  const complete = FORMATS.map((format) => recorded({ ...scenario(), format }));
  try {
    promoteBaseline(complete, destination);
    const original = readFileSync(join(destination, 'xlsx.json'), 'utf8');
    expect(() => promoteBaseline(complete.slice(1), destination)).toThrow(
      'every format'
    );
    const failed = complete.map((run) => ({
      ...run,
      scenarios: run.scenarios.map((s) => ({
        ...s,
        status: 'failed' as const,
        reason: 'injected',
      })),
    }));
    expect(() => promoteBaseline(failed, destination)).toThrow('incomplete');
    expect(readFileSync(join(destination, 'xlsx.json'), 'utf8')).toBe(original);
    writeFileSync(join(destination, 'stale.json'), '{}');
    promoteBaseline(complete, destination);
    expect(() => readFileSync(join(destination, 'stale.json'))).toThrow();
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
