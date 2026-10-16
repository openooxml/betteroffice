import { describe, expect, test } from 'bun:test';
import type { CommandStatus, WorkbookAck } from '@betteroffice/xlsx';
import type { InputDraft, InputSeal } from './inputCoordinator';
import { createPeerlessWorkerInputCoordinator, type PeerlessWorkerCommand, type PeerlessWorkerInputEvent } from './peerlessWorkerInputCoordinator';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, failed) => { resolve = done; reject = failed; });
  return { promise, resolve, reject };
}

function harness(maxOutstanding = 8) {
  const sent: PeerlessWorkerCommand[] = [];
  const events: PeerlessWorkerInputEvent[] = [];
  const errors: unknown[] = [];
  const replies = new Map<number, ReturnType<typeof deferred<CommandStatus>>>();
  const listeners = new Set<(ack: WorkbookAck) => void>();
  const saved: number[] = [];
  const saveEntered = deferred<void>();
  const state = { generation: 1, throughSeq: 0, version: 'engine:initial/α', seal: {} as InputSeal,
    save: Promise.resolve(new Uint8Array([9])), compositionId: undefined as string | undefined };
  const session = {
    commandIdentity: { session: 'book', incarnation: 'worker:α' },
    get confirmedBasis() { return { epoch: 1, version: state.version }; },
    get throughSeq() { return state.throughSeq; },
    submit(command: PeerlessWorkerCommand) {
      sent.push(command);
      const reply = deferred<CommandStatus>();
      replies.set(command.seq, reply);
      return reply.promise;
    },
    onAck(listener: (ack: WorkbookAck) => void) {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    save: () => { saved.push(state.throughSeq); saveEntered.resolve(); return state.save; },
    whenPresented: async (_seq: number) => {},
  };
  const coordinator = createPeerlessWorkerInputCoordinator({
    session: () => session, generation: () => state.generation, isReady: () => true,
    seal: () => state.seal, sync: () => {}, compositionId: () => state.compositionId,
    target: (sheet, row, col) => ({ sheetId: `sheet:${sheet}`, range: {
      kind: 'rowCol', start: { row, col }, end: { row, col },
    } }),
    onEvent: (event) => events.push(event), onError: (error) => errors.push(error), maxOutstanding,
  });
  const draft = (value: string, row = 0, col = 0): InputDraft => ({
    generation: state.generation, sheet: 0, row, col, source: 'cell', value,
  });
  const ack = (seq: number, status: CommandStatus['status'] = 'applied', version = `engine:${seq}/β`) => {
    state.throughSeq = Math.max(state.throughSeq, seq);
    state.version = version;
    const receipt: CommandStatus = { seq, status, ...(status === 'refused'
      ? { reason: 'locked-target', committed: false, affectedSeqs: [seq] } : {}) };
    for (const listener of listeners) listener({ throughSeq: state.throughSeq, version, statuses: [receipt] });
    replies.get(seq)!.resolve(receipt);
    return receipt;
  };
  const dispatched = (seq: number) => {
    if (sent.some((command) => command.seq === seq)) return Promise.resolve();
    return new Promise<void>((resolve) => {
      const off = coordinator.subscribe((event) => {
        if (event.type === 'dispatched' && event.seq === seq) { off(); resolve(); }
      });
    });
  };
  return { coordinator, state, sent, events, errors, saved, saveEntered, draft, ack, dispatched };
}

describe('peerless command swap input boundaries', () => {
  test('pipelines dependent commits on confirmed tokens and omits settled successful dependencies', async () => {
    const { coordinator, sent, draft, ack } = harness();
    try {
      const first = coordinator.submit(draft('A1'));
      const next = coordinator.submit(draft('A2', 1), { after: first });
      expect(sent.map(({ seq, basis, afterSeq }) => ({ seq, basis, afterSeq }))).toEqual([
        { seq: 1, basis: { epoch: 1, version: 'engine:initial/α' }, afterSeq: undefined },
        { seq: 2, basis: { epoch: 1, version: 'engine:initial/α' }, afterSeq: 1 },
      ]);
      ack(1); ack(2);
      await Promise.all([first.result, next.result]);
      const later = coordinator.submit(draft('A3', 2), { after: next });
      expect(sent[2]).toMatchObject({ seq: 3, basis: { epoch: 1, version: 'engine:2/β' } });
      expect(sent[2].afterSeq).toBeUndefined();
      ack(3);
      await later.result;
    } finally { coordinator.retire(); }
  });

  test('unsent same-cell aliases share sequence and outcome without crossing a mode fence', async () => {
    const { coordinator, sent, draft, ack } = harness(1);
    try {
      const first = coordinator.submit(draft('sent'));
      const old = coordinator.submit(draft('old', 1));
      const latest = coordinator.submit(draft('latest', 1));
      expect(latest).toBe(old);
      expect(old.seq).toBeNull();
      const fence = coordinator.setReadOnly(false);
      const afterFence = coordinator.submit(draft('after fence', 1));
      expect(afterFence).not.toBe(old);
      expect(sent).toHaveLength(1);
      ack(1);
      expect(old.seq).toBe(2);
      expect(latest.seq).toBe(2);
      expect(sent[1]).toMatchObject({ kind: 'editCell', payload: { args: [0, 1, 0, 'latest'] } });
      const outcome = ack(2, 'noop');
      expect(await old.result).toBe(outcome);
      expect(await latest.result).toBe(outcome);
      expect(sent[2].kind).toBe('setReadOnly');
      ack(3);
      expect(sent[3]).toMatchObject({ kind: 'editCell',
        payload: { args: [0, 1, 0, 'after fence'] } });
      expect(sent[3].afterSeq).toBeUndefined();
      ack(4);
      await Promise.all([first.result, fence.result, afterFence.result]);
    } finally { coordinator.retire(); }
  });

  test('refused aliases cancel queued descendants while an independent B1 draft retains focus', async () => {
    const { coordinator, sent, events, draft, ack } = harness(1);
    try {
      const blocker = coordinator.submit(draft('blocker', 3));
      const original = coordinator.submit(draft('old A1'));
      const alias = coordinator.submit(draft('restore A1'));
      const predicted = coordinator.submit(draft('dependent A2', 1), { after: alias });
      const descendant = coordinator.submit(draft('dependent A3', 2), { after: predicted });
      const independent = coordinator.submit(draft('independent B1', 0, 1));
      const focused = draft('focused B1', 0, 1);
      coordinator.setDraft(focused);
      ack(1);
      const refused = Promise.allSettled([original.result, alias.result, predicted.result, descendant.result]);
      ack(2, 'refused');
      for (const result of await refused) expect(result).toMatchObject({ status: 'rejected', reason: { code: 'input-failed' } });
      expect(sent.map((command) => command.seq)).toEqual([1, 2, 3]);
      expect(sent[2]).toMatchObject({ payload: { args: [0, 0, 1, 'independent B1'] } });
      expect(coordinator.draft).toEqual(focused);
      expect(coordinator.rejected).toEqual([draft('restore A1')]);
      expect(events.filter((event) => event.type === 'refused')).toHaveLength(1);
      expect(events.filter((event) => event.type === 'cancelled')).toHaveLength(2);
      ack(3);
      await Promise.all([blocker.result, independent.result]);
    } finally { coordinator.retire(); }
  });

  test('clipboard preparation captures target and permission before an immediate read-only fence', async () => {
    const { coordinator, sent, draft, ack, dispatched } = harness();
    try {
      const capture = deferred<string>();
      const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A5' } };
      const pasted = coordinator.clipboard(() => capture.promise, (value) => ({
        kind: 'editCell', payload: { args: [0, 4, 0, value] },
      }), { target });
      target.range.a1 = 'B9';
      const mode = coordinator.setReadOnly(true);
      const blocked = coordinator.submit(draft('blocked', 8));
      await expect(blocked.result).rejects.toMatchObject({ code: 'input-failed' });
      expect(blocked.seq).toBeNull();
      expect(coordinator.readOnly).toBe(true);
      expect(sent).toEqual([]);
      const posted = dispatched(2);
      capture.resolve('captured paste');
      await posted;
      expect(sent).toHaveLength(2);
      expect(sent[0]).toMatchObject({ kind: 'editCell', payload: {
        args: [0, 4, 0, 'captured paste'], target: { sheetId: 'sheet:0', range: { a1: 'A5' } },
      } });
      expect(sent[1]).toMatchObject({ kind: 'setReadOnly', payload: { readOnly: true } });
      ack(1); ack(2);
      await Promise.all([pasted.result, mode.result]);
    } finally { coordinator.retire(); }
  });

  test('captured Save waits for prior acks and holds later dispatch until its ordinary method completes', async () => {
    const { coordinator, state, sent, saved, saveEntered, draft, ack, dispatched } = harness();
    try {
      const savingBytes = deferred<Uint8Array<ArrayBuffer>>();
      state.save = savingBytes.promise;
      const first = coordinator.submit(draft('included'));
      const saving = coordinator.save();
      const later = coordinator.submit(draft('excluded', 1));
      expect(saved).toEqual([]);
      expect(sent).toHaveLength(1);
      ack(1);
      await saveEntered.promise;
      expect(saved).toEqual([1]);
      expect(sent).toHaveLength(1);
      const posted = dispatched(2);
      savingBytes.resolve(new Uint8Array([4, 2]));
      expect(await saving).toEqual(new Uint8Array([4, 2]));
      await posted;
      expect(sent[1]).toMatchObject({ payload: { args: [0, 1, 0, 'excluded'] } });
      ack(2);
      await Promise.all([first.result, later.result]);
    } finally { coordinator.retire(); }
  });

  test('Save seals one local composition at A5 after focus moves and excludes later input', async () => {
    const { coordinator, state, sent, saved, draft, ack, dispatched } = harness();
    try {
      const composition = deferred<boolean>();
      state.seal = { composition: composition.promise };
      state.compositionId = 'ime:captured';
      coordinator.setDraft(draft('preedit', 4));
      const saving = coordinator.save();
      const later = coordinator.submit(draft('later B1', 0, 1));
      coordinator.setDraft({ ...draft('変換', 9), source: 'formula' });
      const posted = dispatched(1);
      composition.resolve(true);
      await posted;
      expect(sent).toHaveLength(1);
      expect(sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 4, 0, '変換'] } });
      const trailing = coordinator.submit(draft('変換', 4), { compositionId: 'ime:captured' });
      expect(trailing.seq).toBe(1);
      expect(sent).toHaveLength(1);
      ack(1);
      await saving;
      expect(saved).toEqual([1]);
      await dispatched(2);
      expect(sent[1]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 1, 'later B1'] } });
      ack(2);
      await Promise.all([trailing.result, later.result]);
    } finally { coordinator.retire(); }
  });

  test('a captured flush rejects refused input before save but allows an explicit independent retry', async () => {
    const { coordinator, sent, saved, draft, ack } = harness();
    try {
      const original = draft('refused');
      const editing = coordinator.submit(original);
      const saving = coordinator.save();
      const settled = Promise.allSettled([editing.result, saving]);
      ack(1, 'refused');
      for (const result of await settled) expect(result).toMatchObject({ status: 'rejected', reason: { code: 'input-failed' } });
      expect(saved).toEqual([]);
      coordinator.discardRejected(original);
      const retry = coordinator.submit(draft('retry', 0, 1));
      expect(sent[1].afterSeq).toBeUndefined();
      ack(2);
      await retry.result;
      await coordinator.save();
      expect(saved).toEqual([2]);
    } finally { coordinator.retire(); }
  });
});
