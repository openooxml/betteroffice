import { describe, expect, test } from 'bun:test';
import type { Proposal } from '@betteroffice/xlsx';
import type { CommandStatus, WorkbookAck, WorkbookCommand } from '@betteroffice/xlsx/worker';
import { createXlsxCommandController } from '../commands/createXlsxCommandStore';
import { createPeerlessWorkerInputCoordinator, type PeerlessWorkerInputSeal } from '../commands/peerlessWorkerInputCoordinator';
import type { InputDraft } from '../commands/inputCoordinator';
import { createAsyncEditorApi } from './createAsyncEditorApi';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function harness(readOnly = false) {
  const sent: WorkbookCommand[] = [];
  const rawCalls: PropertyKey[] = [];
  const replies = new Map<number, ReturnType<typeof deferred<CommandStatus>>>();
  const posted = new Map<WorkbookCommand['kind'], ReturnType<typeof deferred<WorkbookCommand>>>();
  const listeners = new Set<(ack: WorkbookAck) => void>();
  const state = { throughSeq: 0, seal: {} as PeerlessWorkerInputSeal, readOnly };
  const session = {
    commandIdentity: { session: 'book', incarnation: 'worker' },
    get confirmedBasis() { return { epoch: 1, version: `v${state.throughSeq}` }; },
    get throughSeq() { return state.throughSeq; },
    submit(command: WorkbookCommand) {
      sent.push(command);
      const reply = deferred<CommandStatus>();
      replies.set(command.seq, reply);
      posted.get(command.kind)?.resolve(command);
      return reply.promise;
    },
    onAck(listener: (ack: WorkbookAck) => void) {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    save: async () => new Uint8Array(),
    whenPresented: async () => {},
    call: new Proxy({}, { get(_target, key) {
      rawCalls.push(key);
      throw new Error('Proposal creation must use session commands');
    } }),
  };
  const owner = { current: true, retiring: false, ready: true, failure: null, session,
    whenReady: async () => {} } as unknown as PeerlessWorkbookSession;
  const coordinator = createPeerlessWorkerInputCoordinator({
    session: () => session, generation: () => 1, isReady: () => true,
    seal: () => state.seal, sync: () => {}, readOnly: () => state.readOnly,
    target: (sheet, row, col) => ({ sheetId: `sheet:${sheet}`, range: {
      kind: 'rowCol', start: { row, col }, end: { row, col },
    } }),
  });
  const bridge = { focus() {}, exportCoordinator: () => coordinator } as unknown as WorkerEditorApiBridge;
  const api = createAsyncEditorApi(owner, createXlsxCommandController().store, () => bridge);
  const waitFor = (kind: WorkbookCommand['kind']) => {
    const value = deferred<WorkbookCommand>();
    posted.set(kind, value);
    return value.promise;
  };
  const ack = (command: WorkbookCommand, result?: unknown, reason?: string) => {
    state.throughSeq = command.seq;
    const status: CommandStatus = { seq: command.seq, status: reason ? 'refused' : 'applied', result, reason };
    for (const listener of listeners) listener({ throughSeq: command.seq, version: `v${command.seq}`, statuses: [status] });
    replies.get(command.seq)!.resolve(status);
  };
  const draft: InputDraft = { generation: 1, sheet: 0, row: 0, col: 0, source: 'cell', value: 'typed' };
  return { api, owner, coordinator, state, sent, rawCalls, waitFor, ack, draft };
}

const request = () => ({ agentId: 'agent', note: 'totals', edits: [{ sheet: 0, row: 1, col: 0, input: '=SUM(A1:B1)' }] });
const proposal: Proposal = { id: 'proposal:1', agentId: 'agent', note: 'totals',
  cells: [{ sheet: 0, row: 1, col: 0, a1: 'A2', oldText: '', newText: '3' }] };

describe('async worker proposals', () => {
  test('orders a proposal before later typing while an earlier receipt is pending', async () => {
    const h = harness();
    const first = h.coordinator.submit({ ...h.draft, value: 'A' });
    const proposed = h.waitFor('propose');
    const result = h.api.proposeChanges!({ agentId: 'agent',
      edits: [{ sheet: 0, row: 0, col: 0, input: 'agent edit' }] });
    const later = h.coordinator.submit({ ...h.draft, value: 'B' });
    expect(h.sent).toHaveLength(1);
    expect(later.seq).toBeNull();
    h.ack(h.sent[0]!);
    await proposed;
    for (const command of h.sent.slice(1)) h.ack(command, command.kind === 'propose' ? proposal : undefined);
    await Promise.all([first.result, result, later.result]);
    expect(h.sent.map((command) => command.kind === 'editCell' ? command.payload.args[3] : command.kind))
      .toEqual(['A', 'propose', 'B']);
    expect(h.rawCalls).toEqual([]);
  });

  test('drains pending composition and its edit before proposing through the session command layer', async () => {
    const h = harness();
    const composition = deferred<boolean>();
    h.state.seal = { composition: composition.promise, compositionDraft: () => h.draft };
    const edited = h.waitFor('editCell');
    const proposed = h.waitFor('propose');
    const input = request();
    const result = h.api.proposeChanges!(input);
    input.agentId = 'changed';
    input.edits[0]!.input = 'changed';
    expect(h.sent).toEqual([]);
    composition.resolve(true);
    const edit = await edited;
    expect(edit.payload).toMatchObject({ args: [0, 0, 0, 'typed'] });
    expect(h.sent.map((command) => command.kind)).toEqual(['editCell']);
    h.ack(edit);
    const command = await proposed;
    expect(command.payload).toEqual({ args: ['agent', 'totals', request().edits] });
    h.ack(command, proposal);
    await expect(result).resolves.toEqual({ ok: true, proposal });
    expect(h.sent.map((command) => command.kind)).toEqual(['editCell', 'propose']);
    expect(h.rawCalls).toEqual([]);
  });

  for (const state of ['loading', 'failed', 'replaced', 'retiring'] as const) {
    test(`returns an existing admission reason when ${state}`, async () => {
      const h = harness();
      if (state === 'loading') Object.assign(h.owner, { ready: false });
      if (state === 'failed') Object.assign(h.owner, { failure: new Error('worker failed') });
      if (state === 'replaced') Object.assign(h.owner, { current: false });
      if (state === 'retiring') h.owner.retiring = true;
      await expect(h.api.proposeChanges!(request())).resolves.toMatchObject({ ok: false,
        failure: { code: state === 'replaced' || state === 'retiring' ? 'document-replaced' : 'editor-unavailable' } });
      expect(h.sent).toEqual([]);
      expect(h.rawCalls).toEqual([]);
    });

    test(`rechecks admission after pending input drains when ${state}`, async () => {
      const h = harness();
      h.coordinator.setDraft(h.draft);
      const edited = h.waitFor('editCell');
      const result = h.api.proposeChanges!(request());
      const edit = await edited;
      if (state === 'loading') Object.assign(h.owner, { ready: false });
      if (state === 'failed') Object.assign(h.owner, { failure: new Error('worker failed') });
      if (state === 'replaced') Object.assign(h.owner, { current: false });
      if (state === 'retiring') h.owner.retiring = true;
      h.ack(edit);
      await expect(result).resolves.toMatchObject({ ok: false,
        failure: { code: state === 'replaced' || state === 'retiring' ? 'document-replaced' : 'editor-unavailable' } });
      expect(h.sent.map((command) => command.kind)).toEqual(['editCell']);
    });
  }

  for (const state of ['failed', 'replaced'] as const) {
    test(`refuses a proposal reply from a ${state} editor`, async () => {
      const h = harness();
      const proposed = h.waitFor('propose');
      const result = h.api.proposeChanges!(request());
      const command = await proposed;
      if (state === 'failed') Object.assign(h.owner, { failure: new Error('worker failed') });
      else Object.assign(h.owner, { current: false });
      h.ack(command, proposal);
      await expect(result).resolves.toMatchObject({ ok: false,
        failure: { code: state === 'replaced' ? 'document-replaced' : 'editor-unavailable' } });
    });
  }

  test('does not propose when preceding input is refused', async () => {
    const h = harness();
    h.coordinator.setDraft(h.draft);
    const edited = h.waitFor('editCell');
    const result = h.api.proposeChanges!(request());
    h.ack(await edited, undefined, 'locked-target');
    await expect(result).resolves.toMatchObject({ ok: false, failure: { code: 'input-failed' } });
    expect(h.sent.map((command) => command.kind)).toEqual(['editCell']);
  });

  test('returns read-only without submitting a proposal', async () => {
    const h = harness(true);
    await expect(h.api.proposeChanges!(request())).resolves.toMatchObject({ ok: false, failure: { code: 'read-only' } });
    expect(h.sent).toEqual([]);
  });

  for (const [reason, code] of [['read-only', 'read-only'], ['unsupported', 'unsupported-command'],
    ['invalid-command', 'invalid-arguments']] as const) {
    test(`returns the command store reason for ${reason}`, async () => {
      const h = harness();
      const proposed = h.waitFor('propose');
      const result = h.api.proposeChanges!(request());
      h.ack(await proposed, undefined, reason);
      await expect(result).resolves.toMatchObject({ ok: false, failure: { code } });
    });
  }
});
