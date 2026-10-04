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
  expect(adapter.adapt(exported('W1', 7), 'P', 'W1')).toMatchObject({
    version: 'P', content: { layout: { documentVersion: 'P', layoutVersion: 'P:7' } },
  });
  expect(adapter.adapt(exported('W2', 8), 'P', 'W2')).toMatchObject({ content: { layout: { layoutVersion: 'P:8' } } });
  expect(adapter.workerLayoutVersion('P:7')).toBe('W1:7');
  expect(adapter.workerLayoutVersion('P:8')).toBe('W2:8');
  expect(adapter.workerLayoutVersion('unknown:7')).toBe('unknown:7');
});

test('layout associations are scoped to the peer, worker owner and load', () => {
  const peer = {} as YrsSession;
  const owner = {};
  const adapter = workerExportVersions(peer, owner, 1);
  adapter.adapt(exported('W', 7), 'P', 'W');
  expect(workerExportVersions(peer, owner, 1)).toBe(adapter);
  expect(workerExportVersions({} as YrsSession, owner, 1).workerLayoutVersion('P:7')).toBe('P:7');
  expect(workerExportVersions(peer, {}, 1).workerLayoutVersion('P:7')).toBe('P:7');
  expect(workerExportVersions(peer, owner, 2).workerLayoutVersion('P:7')).toBe('P:7');
});

test('retiring a worker export drops its registration and layout associations', () => {
  const peer = {} as YrsSession;
  const owner = {};
  const adapter = workerExportVersions(peer, owner, 1);
  adapter.adapt(exported('W', 7), 'P', 'W');
  registerWorkerOpenExport(peer, { export: async () => exported('W', 7) });
  expect(workerOpenExport(peer)).not.toBeNull();
  retireWorkerOpenExport(peer);
  expect(workerOpenExport(peer)).toBeNull();
  expect(workerExportVersions(peer, owner, 1)).not.toBe(adapter);
  expect(workerExportVersions(peer, owner, 1).workerLayoutVersion('P:7')).toBe('P:7');
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
  adapter.adapt(exported('W1', 7), 'P', 'W1');
  expect(() => adapter.adapt(exported('W2', 7), 'P', 'W2')).toThrow(ResidentWorkerFailureError);
  expect(adapter.workerLayoutVersion('P:7')).toBe('W1:7');
});
