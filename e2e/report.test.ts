import { expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { recorded, scenario } from './test-fixtures';
import { diff, readRuns, rows } from './report';

test('missing or empty results cannot pass comparison', () => {
  const dir = mkdtempSync(join(tmpdir(), 'e2e-report-test-'));
  try {
    expect(() => readRuns(join(dir, 'missing'))).toThrow('does not exist');
    expect(() => readRuns(dir)).toThrow('no results');
    expect(() => diff([recorded()], [])).toThrow('missing');
    expect(() => diff([], [])).toThrow('missing');
    expect(() =>
      diff([recorded()], [recorded({ ...scenario(), format: 'docx' })])
    ).toThrow('mismatched');
    writeFileSync(
      join(dir, 'xlsx.json'),
      JSON.stringify({ ...recorded(), schemaVersion: 2 })
    );
    expect(() => readRuns(dir)).toThrow('schemaVersion');
    writeFileSync(join(dir, 'xlsx.json'), JSON.stringify(recorded()));
    expect(readRuns(dir)).toHaveLength(1);
    writeFileSync(join(dir, 'duplicate.json'), JSON.stringify(recorded()));
    expect(() => readRuns(dir)).toThrow('duplicate');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('the diff CLI exits nonzero when current results are missing', () => {
  const child = Bun.spawnSync([
    process.execPath,
    join(import.meta.dir, 'report.ts'),
    '--diff',
    '/missing-e2e-before',
    '/missing-e2e-after',
  ]);
  expect(child.exitCode).not.toBe(0);
});

test('flat reports retain actors and failure status', () => {
  const run = recorded({
    ...scenario([
      { op: 'edit', actor: 'a', e2eMs: 1 },
      { op: 'edit', actor: 'b', e2eMs: 200 },
    ]),
    status: 'failed',
    reason: 'failure',
  });
  expect(rows([run]).map((row) => [row.actor, row.p50Ms, row.status])).toEqual([
    ['a', 1, 'failed'],
    ['b', 200, 'failed'],
  ]);
});
