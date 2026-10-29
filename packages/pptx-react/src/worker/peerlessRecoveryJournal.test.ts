import { expect, test } from 'bun:test';
import { peerlessRecoveryFactory } from './peerlessRecoveryJournal';

const checkpoint = (throughSeq: number, recoveryUpdate?: ArrayBuffer) => ({
  session: 'session', incarnation: 'worker', epoch: 1, throughSeq, version: String(throughSeq),
  bytes: new ArrayBuffer(8), ...(recoveryUpdate ? { recoveryUpdate } : {}),
});

test('recovery adopts baseline and checkpoint buffers without copying', () => {
  const recovery = peerlessRecoveryFactory.create(new Uint8Array(1));
  const update = new ArrayBuffer(32);
  recovery.open(update, 4, '4');
  expect(recovery.baseline?.kind === 'update' ? recovery.baseline.update.buffer : undefined).toBe(update);
  expect(recovery.baseline?.throughSeq).toBe(4);
  const saved = checkpoint(5);
  recovery.checkpoint(saved);
  expect(recovery.saved.bytes.buffer).toBe(saved.bytes);
  expect(recovery.baseline?.kind === 'update' ? recovery.baseline.update.buffer : undefined).toBe(update);
});

test('recovery retains journal across checkpoints and trims only a rebase', () => {
  const recovery = peerlessRecoveryFactory.create(new Uint8Array(1), 3);
  recovery.open(new ArrayBuffer(32), 0, '0');
  recovery.append(1, '1', [{ bytes: new ArrayBuffer(4), origin: 'local' }]);
  expect(recovery.needsRebase()).toBe(true);
  recovery.checkpoint(checkpoint(1));
  expect(recovery.journal).toHaveLength(1);
  expect(recovery.hasUncheckpointedUpdates()).toBe(false);
  recovery.append(2, '2', [{ bytes: new ArrayBuffer(1), origin: 'local' }]);
  expect(recovery.hasUncheckpointedUpdates()).toBe(true);
  const update = new ArrayBuffer(64);
  recovery.checkpoint(checkpoint(1, update));
  expect(recovery.baseline?.kind === 'update' ? recovery.baseline.update.buffer : undefined).toBe(update);
  expect(recovery.journal.map((entry) => entry.throughSeq)).toEqual([2]);
  expect(recovery.needsRebase()).toBe(false);
});

test('recovery journal limit defaults to the larger of sixteen MiB and half the baseline', () => {
  for (const baselineBytes of [2, 40 * 1024 * 1024]) {
    const recovery = peerlessRecoveryFactory.create(new Uint8Array(1));
    recovery.open(new ArrayBuffer(baselineBytes), 0, '0');
    const limit = Math.max(16 * 1024 * 1024, baselineBytes / 2);
    recovery.append(1, '1', [{ bytes: new ArrayBuffer(limit), origin: 'local' }]);
    expect(recovery.needsRebase()).toBe(false);
    recovery.append(2, '2', [{ bytes: new ArrayBuffer(1), origin: 'local' }]);
    expect(recovery.needsRebase()).toBe(true);
  }
});

test('source baseline retains local and remote acknowledgements across ordinary checkpoints', () => {
  const source = new Uint8Array(32);
  const recovery = peerlessRecoveryFactory.create(source);
  recovery.open(undefined, 0, '0', 'source', 'engine');
  expect(recovery.baseline).toEqual({ kind: 'source', throughSeq: 0, engine: 'engine' });
  expect(recovery.needsRebase()).toBe(false);
  recovery.append(1, '1', []);
  recovery.append(2, '2', [{ bytes: new ArrayBuffer(3), origin: 'remote' }]);
  recovery.checkpoint(checkpoint(2));
  recovery.append(3, '3', [{ bytes: new ArrayBuffer(4), origin: 'local' }]);
  recovery.checkpoint(checkpoint(3));
  recovery.append(4, '4', [{ bytes: new ArrayBuffer(2), origin: 'local' }]);
  expect(recovery.journal.map((entry) => entry.throughSeq)).toEqual([2, 3, 4]);
  expect(recovery.validate().kind).toBe('source');
  expect(recovery.hasUncheckpointedUpdates()).toBe(true);
  expect(recovery.needsRebase()).toBe(false);
  recovery.checkpoint(checkpoint(3, new ArrayBuffer(64)));
  expect(recovery.baseline?.kind).toBe('update');
  expect(recovery.baseline?.throughSeq).toBe(3);
  expect(recovery.journal.map((entry) => entry.throughSeq)).toEqual([4]);
  expect(recovery.validate().kind).toBe('update');
});

test('source rebase bound uses the original compressed source byte length', () => {
  for (const sourceBytes of [2, 40 * 1024 * 1024]) {
    const recovery = peerlessRecoveryFactory.create(new Uint8Array(sourceBytes));
    recovery.open(undefined, 0, '0', 'source', 'engine');
    const limit = Math.max(16 * 1024 * 1024, sourceBytes / 2);
    recovery.append(1, '1', [{ bytes: new ArrayBuffer(limit), origin: 'local' }]);
    recovery.checkpoint(checkpoint(1));
    expect(recovery.needsRebase()).toBe(false);
    recovery.append(2, '2', [{ bytes: new ArrayBuffer(1), origin: 'remote' }]);
    expect(recovery.needsRebase()).toBe(true);
  }
});

test('source recovery refuses missing baseline identity, bytes and lost journal entries', () => {
  const unknown = peerlessRecoveryFactory.create(new Uint8Array(1));
  unknown.open(undefined, 0, '0');
  expect(() => unknown.validate()).toThrow('Recovery baseline is unavailable');
  for (const [bytes, engine, throughSeq] of [[0, 'engine', 0], [1, '', 0], [1, 'engine', 1]] as const) {
    const recovery = peerlessRecoveryFactory.create(new Uint8Array(bytes));
    recovery.open(undefined, throughSeq, '0', 'source', engine);
    expect(() => recovery.validate()).toThrow('Recovery baseline is unavailable');
  }
  for (const sequence of [[2], [1, 3], [1, 1], [3, 5, 6]]) {
    const recovery = peerlessRecoveryFactory.create(new Uint8Array(1));
    recovery.open(undefined, 0, '0', 'source', 'engine');
    for (const seq of sequence) recovery.append(seq, String(seq), [{ bytes: new ArrayBuffer(1), origin: 'local' }]);
    expect(recovery.validate().kind).toBe('source');
  }
  const recovery = peerlessRecoveryFactory.create(new Uint8Array(1));
  recovery.open(undefined, 0, '0', 'source', 'engine');
  recovery.append(1, '1', []);
  recovery.append(2, '2', [{ bytes: new ArrayBuffer(1), origin: 'local' }]);
  recovery.journal.shift();
  expect(() => recovery.validate()).toThrow('Recovery baseline is unavailable');
});

test('source recovery keeps its journal when a rebase update is too large to reopen', () => {
  const recovery = peerlessRecoveryFactory.create(new Uint8Array(1), 1);
  recovery.open(undefined, 0, '0', 'source', 'engine');
  recovery.append(1, '1', [{ bytes: new ArrayBuffer(2), origin: 'local' }]);
  expect(recovery.needsRebase()).toBe(true);
  recovery.checkpoint({ bytes: new ArrayBuffer(1), throughSeq: 1, version: '1',
    recoveryUpdate: new ArrayBuffer(64 * 1024 * 1024 + 1) } as Parameters<typeof recovery.checkpoint>[0]);
  expect(recovery.validate().kind).toBe('source');
  expect(recovery.needsRebase()).toBe(false);
  expect(recovery.journal.map((entry) => entry.throughSeq)).toEqual([1]);
  recovery.append(2, '2', [{ bytes: new ArrayBuffer(1), origin: 'local' }]);
  expect(recovery.journal.map((entry) => entry.throughSeq)).toEqual([1, 2]);
});
