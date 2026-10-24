import { expect, spyOn, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session/types';
import { createWorkbookSessionRecovery, WorkbookRecoveryReplayError } from './clientRecovery';

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
