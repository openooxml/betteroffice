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
  expect(recovery.baseline?.update.buffer).toBe(update);
  expect(recovery.baseline?.throughSeq).toBe(4);
  const saved = checkpoint(5);
  recovery.checkpoint(saved);
  expect(recovery.saved.bytes.buffer).toBe(saved.bytes);
  expect(recovery.baseline?.update.buffer).toBe(update);
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
  expect(recovery.baseline?.update.buffer).toBe(update);
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
