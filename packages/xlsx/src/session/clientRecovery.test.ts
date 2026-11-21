import { expect, spyOn, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session/types';
import { createWorkbookSessionRecovery, workbookRecoveryBackend, WorkbookRecoveryReplayError } from './clientRecovery';

function admit(recovery: ReturnType<typeof createWorkbookSessionRecovery>, seq: number) {
  recovery.admit({ session: 'book', incarnation: 1, seq, basis: { epoch: 1, version: 0 }, kind: 'setReadOnly',
    payload: { payload: { readOnly: false } } });
}

const settle = () => new Promise((resolve) => setImmediate(resolve));

test('rejected checkpoints back off and fail typed only at the retained bound', async () => {
  const attempts: number[] = [];
  const failures: unknown[] = [];
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const recovery = createWorkbookSessionRecovery(new Uint8Array([1]), {}, async () => {
    attempts.push(recovery.recovery.retainedCommands);
    throw new Error('malformed request: Recovery projection cannot be reconstructed');
  }, 1, false, undefined, (error) => failures.push(error));
  try {
    for (let seq = 1; seq < 4096; seq += 1) {
      admit(recovery, seq);
      await settle();
    }
    expect(failures).toEqual([]);
    expect(attempts.length).toBeGreaterThan(3);
    expect(attempts.length).toBeLessThan(12);
    expect(attempts[0]).toBe(256);
    expect(recovery.recovery.retainedCommands).toBe(4095);
    admit(recovery, 4096);
    await settle();
    expect(failures).toHaveLength(1);
    expect(failures[0]).toBeInstanceOf(WorkbookRecoveryReplayError);
    expect(failures[0]).toMatchObject({ reason: 'log-full' });
    expect(recovery.recovery.retainedCommands).toBe(4096);
  } finally { warn.mockRestore(); recovery.dispose(); }
});

test('a worker failure during a checkpoint still fails the session', async () => {
  const failures: unknown[] = [];
  const crash = new SessionFailure('crash', 'Worker stopped');
  const recovery = createWorkbookSessionRecovery(new Uint8Array([1]), {}, async () => { throw crash; },
    1, false, undefined, (error) => failures.push(error));
  try {
    admit(recovery, 1);
    await expect(recovery.recovery.capture(() => true)).rejects.toBe(crash);
    expect(failures).toEqual([crash]);
  } finally { recovery.dispose(); }
});

function checkpointed(open: (bytes: Uint8Array, checkpoint: unknown) => Promise<Uint8Array>) {
  const calls: { bytes: number[]; checkpoint: unknown }[] = [];
  const backend = spyOn(workbookRecoveryBackend, 'open').mockImplementation(async (bytes, _options, checkpoint) => {
    calls.push({ bytes: [...bytes], checkpoint });
    const saved = await open(bytes, checkpoint);
    return { replay: async () => { throw new Error('nothing to replay'); }, save: async () => saved, dispose: async () => {} };
  });
  const recovery = createWorkbookSessionRecovery(new Uint8Array([1, 2, 3]), {}, async () => ({
    epoch: 1, throughSeq: 1, version: 1, state: new Uint8Array([9, 9]),
  }), 1);
  return { calls, backend, recovery };
}

test('recovery reopens from the source until the first checkpoint, then from the checkpoint alone', async () => {
  const { calls, backend, recovery } = checkpointed(async () => new Uint8Array([7]));
  try {
    expect(await recovery.recovery.save(() => {})).toEqual(new Uint8Array([7]));
    expect(calls[0]).toEqual({ bytes: [1, 2, 3], checkpoint: undefined });
    admit(recovery, 1);
    expect(await recovery.recovery.capture(() => true)).toBe(true);
    expect(await recovery.recovery.save(() => {})).toEqual(new Uint8Array([7]));
    expect(calls[1].bytes).toEqual([]);
    expect(calls[1].checkpoint).toMatchObject({ throughSeq: 1, version: 1, state: new Uint8Array([9, 9]) });
  } finally { backend.mockRestore(); recovery.dispose(); }
});

test('an unreadable checkpoint fails recovery loudly instead of reopening a released source', async () => {
  const unreadable = new Error('Recovery checkpoint could not be decoded');
  const { calls, backend, recovery } = checkpointed(async (_bytes, checkpoint) => {
    if (checkpoint) throw unreadable;
    return new Uint8Array([7]);
  });
  try {
    admit(recovery, 1);
    expect(await recovery.recovery.capture(() => true)).toBe(true);
    const error = await recovery.recovery.save(() => {}).then(() => null, (failure: unknown) => failure);
    expect(error).toBeInstanceOf(WorkbookRecoveryReplayError);
    expect(error).toMatchObject({ seq: 1, reason: 'checkpoint-unreadable', cause: unreadable,
      message: 'Workbook recovery refused sequence 1: checkpoint-unreadable' });
    expect(calls).toHaveLength(1);
    expect(calls[0].bytes).toEqual([]);
  } finally { backend.mockRestore(); recovery.dispose(); }
});
