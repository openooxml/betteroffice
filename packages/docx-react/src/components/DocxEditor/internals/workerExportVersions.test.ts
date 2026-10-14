import { expect, test } from 'bun:test';
import { ResidentWorkerFailureError, type YrsSession } from '@betteroffice/docx/yrs';
import { workerExportVersions } from './workerExportVersions';
import { registerWorkerOpenExport, retireWorkerOpenExport, workerOpenExport, type WorkerPageExportResult } from './workerOpenExport';

function exported(worker: string, serial: number): WorkerPageExportResult {
  return {
    ok: true, version: worker,
    content: { structured: {}, layout: { documentVersion: worker, layoutVersion: `${worker}:${serial}` } },
  } as WorkerPageExportResult;
}

test('export versions map document and layout tokens and preserve their original worker association', () => {
  const peer = {} as YrsSession;
  const owner = {};
  const adapter = workerExportVersions(peer, owner, 1);
  const first = adapter.adapt(exported('W1', 7), 'P', 'W1');
  const second = adapter.adapt(exported('W2', 8), 'P', 'W2');
  if (!first.ok || !second.ok) throw new Error('Expected success');
  expect(first).toMatchObject({
    version: 'P', content: { layout: { documentVersion: 'P' } },
  });
  expect(first.content.layout.layoutVersion).toMatch(/^P:\d+:7$/);
  expect(second.content.layout.layoutVersion).toMatch(/^P:\d+:8$/);
  expect(adapter.workerLayoutVersion(first.content.layout.layoutVersion)).toBe('W1:7');
  expect(adapter.workerLayoutVersion(second.content.layout.layoutVersion)).toBe('W2:8');
  expect(adapter.workerLayoutVersion('unknown:7')).toBe('unknown:7');
});

test('layout associations are scoped to the peer, worker owner and load', () => {
  const peer = {} as YrsSession;
  const owner = {};
  const adapter = workerExportVersions(peer, owner, 1);
  const result = adapter.adapt(exported('W', 7), 'P', 'W');
  if (!result.ok) throw new Error('Expected success');
  const token = result.content.layout.layoutVersion;
  expect(workerExportVersions(peer, owner, 1)).toBe(adapter);
  expect(workerExportVersions({} as YrsSession, owner, 1).workerLayoutVersion(token)).toBe(token);
  expect(workerExportVersions(peer, {}, 1).workerLayoutVersion(token)).toBe(token);
  expect(workerExportVersions(peer, owner, 2).workerLayoutVersion(token)).toBe(token);
});

for (const replacement of ['owner', 'load', 'registration'] as const) {
  test(`layout tokens remain unique after replacing the ${replacement} without peer edits`, () => {
    const peer = {} as YrsSession;
    const owner = {};
    const adapter = workerExportVersions(peer, owner, 1);
    const first = adapter.adapt(exported('W1', 7), 'P', 'W1');
    if (!first.ok) throw new Error('Expected success');
    const pinned = first.content.layout.layoutVersion;
    if (replacement === 'registration') retireWorkerOpenExport(peer);
    const next = workerExportVersions(peer, replacement === 'owner' ? {} : owner, replacement === 'load' ? 2 : 1);
    const second = next.adapt(exported('W2', 7), 'P', 'W2');
    if (!second.ok) throw new Error('Expected success');
    expect(second.content.layout.layoutVersion).not.toBe(pinned);
    expect(next.workerLayoutVersion(second.content.layout.layoutVersion)).toBe('W2:7');
    expect(next.workerLayoutVersion(pinned)).toBe(pinned);
    expect(adapter.workerLayoutVersion(pinned)).toBe('W1:7');
  });
}

test('retiring a worker export drops its registration and layout associations', () => {
  const peer = {} as YrsSession;
  const owner = {};
  const adapter = workerExportVersions(peer, owner, 1);
  const result = adapter.adapt(exported('W', 7), 'P', 'W');
  if (!result.ok) throw new Error('Expected success');
  const token = result.content.layout.layoutVersion;
  registerWorkerOpenExport(peer, { export: async () => exported('W', 7) });
  expect(workerOpenExport(peer)).not.toBeNull();
  retireWorkerOpenExport(peer);
  expect(workerOpenExport(peer)).toBeNull();
  expect(workerExportVersions(peer, owner, 1)).not.toBe(adapter);
  expect(workerExportVersions(peer, owner, 1).workerLayoutVersion(token)).toBe(token);
});

test('refusals use the captured peer version and mismatched reads cannot be relabeled', () => {
  const adapter = workerExportVersions({} as YrsSession, {}, 1);
  expect(adapter.adapt({ ok: false, version: 'W', failure: { code: 'stale-layout', target: null, message: 'Stale layout' } }, 'P', 'W')).toMatchObject({
    ok: false, version: 'P', failure: { code: 'stale-layout' },
  });
  expect(() => adapter.adapt(exported('later W', 7), 'P', 'W')).toThrow(ResidentWorkerFailureError);
  const wrongLayout = exported('W', 7);
  if (!wrongLayout.ok) throw new Error('Expected success');
  wrongLayout.content.layout.documentVersion = 'later W';
  expect(() => adapter.adapt(wrongLayout, 'P', 'W')).toThrow(ResidentWorkerFailureError);
});

test('a rotated worker token cannot overwrite an earlier pinned association', () => {
  const adapter = workerExportVersions({} as YrsSession, {}, 1);
  const first = adapter.adapt(exported('W1', 7), 'P', 'W1');
  if (!first.ok) throw new Error('Expected success');
  expect(() => adapter.adapt(exported('W2', 7), 'P', 'W2')).toThrow(ResidentWorkerFailureError);
  expect(adapter.workerLayoutVersion(first.content.layout.layoutVersion)).toBe('W1:7');
});
