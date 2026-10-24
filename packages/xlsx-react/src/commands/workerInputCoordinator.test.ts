import { describe, expect, test } from 'bun:test';
import { isClientMessage, SessionFailure } from '../../../../shared/office-session';
import { hydratePeer } from '../../../xlsx/src/session/client';
import { createWorkbookEditPeer, type WorkbookEditPeer } from '../../../xlsx/src/session/editPeer';
import { workbookEditPeerOperations } from '../../../xlsx/src/session/editPeerInternals';
import type { WorkbookReplayOp } from '../../../xlsx/src/session/replay';
import type { CommandStatus, WorkbookAck, WorkbookCommand, WorkbookCommandFrameRequest } from '@betteroffice/xlsx';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from '../../../xlsx/src/session/testHelpers';
import { openWorkbook, type WorkbookHandle } from '../../../xlsx/src/wasm/loader';
import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import type { InputDraft, InputSeal } from './inputCoordinator';
import { WorkerDraftController } from '../worker/WorkerDraftController';
import { holdWorkerCellInputRead } from '../worker/WorkerCellInputReader';
import { createPeerlessPluginDocument } from '../plugins/createPeerlessPluginDocument';
import type { PeerlessWorkbookSession } from '../worker/usePeerlessSessionWorkbook';
import { createWorkerInputCoordinator, WorkerInputNotReadyError, WorkerInputRefusal, type WorkerInputCoordinator } from './workerInputCoordinator';
import {
  createPeerlessWorkerInputCoordinator, type PeerlessWorkerCommand, type PeerlessWorkerCommandInput,
  type PeerlessWorkerInputEvent, type PeerlessWorkerPendingInput,
} from './peerlessWorkerInputCoordinator';

function peerlessHarness(maxOutstanding = 8, maxQueued = 8, controller?: WorkerDraftController) {
  const sent: PeerlessWorkerCommand[] = [];
  const events: PeerlessWorkerInputEvent[] = [];
  const errors: unknown[] = [];
  const listeners = new Set<(ack: WorkbookAck) => void>();
  const failures = new Set<(error: SessionFailure) => void>();
  const replies = new Map<number, ReturnType<typeof deferred<CommandStatus>>>();
  const saved: number[] = [];
  const state = { generation: 1, navigation: 0, throughSeq: 0, version: 'v1', ready: true,
    seal: {} as InputSeal, frame: undefined as WorkbookCommandFrameRequest | undefined };
  const session = {
    commandIdentity: { session: 'book', incarnation: 'worker' },
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
    onFailure(listener: (error: SessionFailure) => void) {
      failures.add(listener);
      return () => { failures.delete(listener); };
    },
    save: async () => { saved.push(state.throughSeq); return new Uint8Array([1]); },
    whenPresented: async (_seq: number) => {},
  };
  const coordinator = createPeerlessWorkerInputCoordinator({
    session: () => session, generation: () => state.generation,
    isReady: () => state.ready,
    selectionRevision: () => state.navigation,
    frame: () => state.frame,
    seal: () => controller?.seal() ?? state.seal, sync: () => controller?.sync(),
    compositionId: () => controller?.compositionId,
    target: (sheet, row, col) => ({ sheetId: `sheet:${sheet}`, range: {
      kind: 'rowCol', start: { row, col }, end: { row, col },
    } }),
    onEvent: (event) => events.push(event), onError: (error) => errors.push(error),
    maxOutstanding, maxQueued,
  });
  const draft = (value: string, row = 0): InputDraft => ({
    generation: state.generation, sheet: 0, row, col: 0, source: 'cell', value,
  });
  const ack = (seq: number, status: CommandStatus['status'] = 'applied', result?: unknown) => {
    state.throughSeq = seq;
    state.version = `v${seq + 1}`;
    const receipt: CommandStatus = { seq, status, ...(result === undefined ? {} : { result }), ...(status === 'refused' ? { reason: 'locked-target' } : {}) };
    for (const listener of listeners) listener({ throughSeq: seq, version: state.version, statuses: [receipt] });
    replies.get(seq)?.resolve(receipt);
  };
  const failSession = (error: SessionFailure) => { for (const listener of failures) listener(error); };
  return { coordinator, session, state, sent, events, errors, saved, draft, ack, replies, listeners, failSession };
}

describe('cell input read hold', () => {
  test('waits for prior input and holds following commands until the seed is applied', async () => {
    const { coordinator, draft, sent, ack } = peerlessHarness();
    const first = coordinator.submit(draft('first', 1));
    const entered = deferred<void>();
    const reply = deferred<string>();
    const received = deferred<void>();
    const applied = deferred<boolean>();
    const posted = deferred<void>();
    const values: string[] = [];
    coordinator.subscribe((event) => { if (event.type === 'dispatched' && event.command?.kind === 'undo') posted.resolve(); });
    let started = false;
    const reading = holdWorkerCellInputRead(coordinator, () => {
      started = true; entered.resolve(); return reply.promise;
    }, (value) => { values.push(value); received.resolve(); }, applied.promise, () => false, (error) => { throw error; });
    const later = coordinator.runAfterPendingInput(() => coordinator.command({ kind: 'undo', payload: { args: [] } }).result);
    expect(started).toBe(false);
    ack(first.seq!);
    await entered.promise;
    reply.resolve('Q3 Sales Dashboard');
    await received.promise;
    expect(values).toEqual(['Q3 Sales Dashboard']);
    expect(sent.map((command) => command.kind)).toEqual(['editCell']);
    applied.resolve(true);
    await reading;
    await posted.promise;
    ack(sent[1].seq);
    await later;
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'undo']);
  });

  test('cancelling the editor releases the read hold before the RPC reply without applying it', async () => {
    const { coordinator, saved } = peerlessHarness();
    const entered = deferred<void>();
    const reply = deferred<string>();
    const applied = deferred<boolean>();
    const values: string[] = [];
    let cancelled = false;
    const reading = holdWorkerCellInputRead(coordinator, () => {
      entered.resolve(); return reply.promise;
    }, (value) => { values.push(value); }, applied.promise, () => cancelled, (error) => { throw error; });
    await entered.promise;
    const saving = coordinator.save();
    cancelled = true;
    applied.resolve(false);
    await Promise.all([reading, saving]);
    expect(saved).toEqual([0]);
    reply.resolve('late');
    await Promise.resolve();
    expect(values).toEqual([]);
  });

  test('closes the editor without failing the session when the read is refused admission', async () => {
    const outcome = deferred<void>();
    const failures: unknown[] = [];
    const reports: unknown[] = [];
    const selected = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
    const controller = new WorkerDraftController({
      generation: () => 1, identity: () => 'book', navigation: () => 0,
      selection: () => selected, sheet: () => 0,
      target: () => ({ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }),
      resolveSheet: () => 0, editable: () => true, prefill: async () => 'value',
      select: () => {}, selectionTarget: () => undefined, dependency: () => null,
      navigate: () => {}, committed: () => {}, focus: () => {},
      fail: (error) => { failures.push(error); outcome.resolve(); },
      report: (error) => { reports.push(error); outcome.resolve(); },
    });
    const { coordinator } = peerlessHarness(8, 1, controller);
    controller.attach(coordinator);
    void coordinator.runAfterPendingInput(() => new Promise<void>(() => {}), true);
    controller.open();
    await outcome.promise;
    expect(failures).toEqual([]);
    expect(reports).toEqual([expect.any(WorkerInputRefusal)]);
    expect(controller.loadingCell).toBe(false);
    expect(controller.getSnapshot().cell).toBeNull();
  });
});

describe('prefill reservation recovery', () => {
  function pending(target: InputDraft) {
    const ready = deferred<boolean>();
    let completed: InputDraft | null = null;
    const input: PeerlessWorkerPendingInput = { target, ready: ready.promise, draft: () => completed };
    return { input, complete(value: string) {
      completed = Object.freeze({ ...target, value, modified: true });
      ready.resolve(true);
    } };
  }

  test('reuses a reservation after a newer draft replaces the input barrier', async () => {
    const { coordinator, draft, failSession } = peerlessHarness();
    const entered = deferred<void>();
    const reading = coordinator.runAfterPendingInput(() => {
      entered.resolve(); return new Promise<void>(() => {});
    }, true);
    await entered.promise;
    const prefill = pending(draft(''));
    const first = coordinator.reserveInput(prefill.input);
    coordinator.setDraft(draft('Y', 1));
    const saving = coordinator.save();
    expect(coordinator.reserveInput(prefill.input)).toBe(first);
    prefill.complete('Q3 Sales DashboardX');
    failSession(new SessionFailure('crash', 'Worker stopped'));
    await Promise.allSettled([reading, first.result, saving]);
    expect(coordinator.recoveryCommands).toMatchObject([
      { kind: 'editCell', payload: { args: [0, 0, 0, 'Q3 Sales DashboardX'] } },
      { kind: 'editCell', payload: { args: [0, 1, 0, 'Y'] } },
    ]);
    expect(coordinator.recoveryCommands).toHaveLength(2);
  });

  for (const failed of ['read', 'session'] as const) for (const cleared of [false, true]) {
    test(`retains a ${cleared ? 'cleared' : 'completed'} prefill once before later edits after ${failed} failure`, async () => {
      const { coordinator, failSession, draft, sent } = peerlessHarness();
      const entered = deferred<void>();
      const reply = deferred<void>();
      const reading = coordinator.runAfterPendingInput(async () => {
        entered.resolve();
        try { await reply.promise; }
        catch (error) { coordinator.fail(error); throw error; }
      }, true);
      await entered.promise;
      const prefill = pending(draft(''));
      coordinator.reserveInput(prefill.input);
      const later = coordinator.submit(draft('Y', 1));
      const saving = coordinator.save();
      const repeated = coordinator.save();
      const value = cleared ? '' : 'Q3 Sales DashboardX';
      prefill.complete(value);
      const failure = new SessionFailure('crash', 'Worker stopped');
      if (failed === 'read') reply.reject(failure);
      else failSession(failure);
      await Promise.allSettled([reading, later.result, saving, repeated]);
      expect(sent).toEqual([]);
      expect(coordinator.recoveryCommands).toMatchObject([
        { kind: 'editCell', payload: { args: [0, 0, 0, value] } },
        { kind: 'editCell', payload: { args: [0, 1, 0, 'Y'] } },
      ]);
      expect(coordinator.recoveryCommands).toHaveLength(2);
    });
  }

  test('a refused reservation stops being the default dependency', async () => {
    const { coordinator, draft, sent, ack } = peerlessHarness();
    const dispatched = deferred<void>();
    coordinator.subscribe((event) => { if (event.type === 'dispatched') dispatched.resolve(); });
    const prefill = pending(draft(''));
    const reserved = coordinator.reserveInput(prefill.input);
    prefill.complete('X');
    await dispatched.promise;
    ack(sent[0].seq, 'refused');
    await reserved.result.catch(() => {});
    coordinator.command({ kind: 'undo', payload: { args: [] } });
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'undo']);
  });
});

describe('prefill refusals and seals', () => {
  function setup() {
    const prefill = deferred<string>();
    const view = { editable: true, navigation: 0, row: 0 };
    let readOnly = () => false;
    const controller = new WorkerDraftController({
      generation: () => 1, identity: () => 'book', navigation: () => view.navigation,
      selection: () => ({ anchor: { row: view.row, col: 0 }, focus: { row: view.row, col: 0 } }), sheet: () => 0,
      target: (_sheet, row, col) => ({ sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row, col }, end: { row, col } } }),
      resolveSheet: () => 0, editable: () => view.editable && !readOnly(), prefill: () => prefill.promise,
      select: (_sheet, cell, move) => { view.row = cell.row + (move === 'down' ? 1 : 0); view.navigation += 1; },
      selectionTarget: () => undefined, dependency: () => null, navigate: () => {},
      committed: (ticket) => { void ticket.result.catch(() => {}); },
      fail: (error) => { throw error; }, report: () => {}, focus: () => {},
    });
    const harness = peerlessHarness(8, 8, controller);
    const { coordinator, sent, ack } = harness;
    readOnly = () => coordinator.readOnly;
    controller.attach(coordinator);
    let acknowledging = false;
    coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.draft && event.seq !== null) controller.submitted(event.draft, event.seq);
      if (event.type === 'refused' && event.draft) controller.refused(event.seq, event.draft);
      if (acknowledging && event.type === 'dispatched' && event.seq !== null) queueMicrotask(() => ack(event.seq!));
    });
    const dispatched = (kind: PeerlessWorkerCommand['kind'], count = 1) => new Promise<void>((resolve) => {
      const check = () => harness.events.filter((event) => event.type === 'dispatched' && event.command?.kind === kind).length >= count;
      if (check()) { resolve(); return; }
      const off = coordinator.subscribe(() => { if (check()) { off(); resolve(); } });
    });
    const edits = () => sent.flatMap((command) => command.kind === 'editCell' ? [command.payload.args[3]] : []);
    const seq = (match: (command: PeerlessWorkerCommand) => boolean) => sent.find(match)!.seq;
    const reserveTypedF2 = async () => {
      controller.open();
      controller.holdCellEdit({ kind: 'insert', text: 'X' });
      controller.commit('cell', 'down');
      prefill.resolve('initial');
      await dispatched('editCell');
    };
    return { ...harness, controller, prefill, view, dispatched, edits, seq, reserveTypedF2,
      acknowledgeLater: () => { acknowledging = true; } };
  }
  const edit = (value: string) => (command: PeerlessWorkerCommand) => command.kind === 'editCell' && command.payload.args[3] === value;
  const mode = (readOnly: boolean) => (command: PeerlessWorkerCommand) => command.kind === 'setReadOnly' && command.payload.readOnly === readOnly;

  test('a reserved refusal while a completed next-cell composition owns focus is not resubmitted by Save', async () => {
    const { coordinator, controller, prefill, ack, dispatched, edits, seq, acknowledgeLater } = setup();
    controller.open();
    controller.holdCellEdit({ kind: 'insert', text: 'X' });
    controller.commit('cell', 'down');
    controller.open('Y');
    const input = { value: 'Y', selectionStart: 1, selectionEnd: 1, setSelectionRange() {} } as unknown as HTMLInputElement;
    controller.cellInput = input;
    controller.start('cell');
    input.value = 'Y変';
    controller.change('cell', 'Y変', input);
    controller.end('変');
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    prefill.resolve('initial');
    await dispatched('editCell', 2);
    expect(edits()).toEqual(['initialX', 'Y変']);
    ack(seq(edit('initialX')), 'refused');
    ack(seq(edit('Y変')), 'cancelled');
    expect(controller.getSnapshot().cell?.value).toBe('Y変');
    expect(controller.getSnapshot().refused.map((record) => record.draft.value)).toEqual(['initialX']);
    acknowledgeLater();
    await expect(coordinator.save()).rejects.toMatchObject({ code: 'input-failed' });
    expect(edits()).toEqual(['initialX', 'Y変']);
  });

  test('a reserved refusal while editing is disabled is not resubmitted after editing resumes', async () => {
    const { coordinator, controller, view, ack, edits, seq, acknowledgeLater, reserveTypedF2 } = setup();
    await reserveTypedF2();
    view.editable = false;
    ack(seq(edit('initialX')), 'refused');
    view.editable = true;
    acknowledgeLater();
    await expect(coordinator.save()).rejects.toMatchObject({ code: 'input-failed' });
    expect(edits()).toEqual(['initialX']);
    expect(controller.getSnapshot().refused.map((record) => record.draft.value)).toEqual(['initialX']);
  });

  test('a reserved refusal during a read-only toggle needs a visible restore before Save sends it', async () => {
    const { coordinator, controller, ack, dispatched, edits, seq, acknowledgeLater, reserveTypedF2 } = setup();
    await reserveTypedF2();
    const locking = coordinator.setReadOnly(true);
    await dispatched('setReadOnly');
    ack(seq(edit('initialX')), 'refused');
    ack(seq(mode(true)));
    await locking.result;
    const unlocking = coordinator.setReadOnly(false);
    await dispatched('setReadOnly', 2);
    ack(seq(mode(false)));
    await unlocking.result;
    acknowledgeLater();
    await expect(coordinator.save()).rejects.toMatchObject({ code: 'input-failed' });
    expect(edits()).toEqual(['initialX']);
    expect(controller.getSnapshot().refused.map((record) => record.draft.value)).toEqual(['initialX']);
    controller.restore();
    expect(controller.getSnapshot().cell?.value).toBe('initialX');
    controller.commit('cell');
    await dispatched('editCell', 2);
    expect(edits()).toEqual(['initialX', 'initialX']);
  });

  test('an untouched F2 stays open across a plugin read that waits for its seed', async () => {
    const { coordinator, controller, prefill, sent } = setup();
    controller.open();
    const reading = coordinator.runAfterPendingInput(() => controller.getSnapshot().cell?.value);
    expect(controller.getSnapshot().cell).not.toBeNull();
    prefill.resolve('initial');
    expect(await reading).toBe('initial');
    expect(controller.getSnapshot().cell?.value).toBe('initial');
    expect(controller.loadingCell).toBe(false);
    expect(sent).toEqual([]);
  });
});

describe('completed composition recovery', () => {
  function setup() {
    const selected = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
    const controller = new WorkerDraftController({
      generation: () => 1, identity: () => 'book', navigation: () => 0,
      selection: () => selected, sheet: () => 0,
      target: () => ({ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }),
      resolveSheet: () => 0, editable: () => true, prefill: async () => '',
      select: () => {}, selectionTarget: () => undefined, dependency: () => null,
      navigate: () => {}, committed: () => {}, fail: (error) => { throw error; }, report: (error) => { throw error; },
      focus: () => {},
    });
    const harness = peerlessHarness(8, 8, controller);
    controller.attach(harness.coordinator);
    return { ...harness, controller };
  }

  for (const source of ['cell', 'formula'] as const) {
    for (const failed of ['read', 'session'] as const) {
      test(`retains ${source} reservations in command order after ${failed} failure`, async () => {
        const { coordinator, controller, sent, errors, draft, failSession } = setup();
        const entered = deferred<void>();
        const reply = deferred<void>();
        const reading = coordinator.runAfterPendingInput(async () => {
          entered.resolve();
          try { await reply.promise; }
          catch (error) { coordinator.fail(error); throw error; }
        }, true);
        await entered.promise;
        const before = coordinator.submit(draft('before', 1));
        controller.start(source);
        const firstValue = source === 'cell' ? '変換' : '=SUM(A2:A4)';
        controller.change(source, firstValue, null);
        const first = coordinator.save();
        const repeated = coordinator.save();
        controller.end();
        await Promise.resolve();
        const middle = coordinator.submit(draft('middle', 2));
        controller.start(source);
        const secondValue = source === 'cell' ? '次' : '=A2*2';
        controller.change(source, secondValue, null);
        const second = coordinator.save();
        controller.end();
        await Promise.resolve();
        const after = coordinator.submit(draft('after', 3));
        expect(sent).toEqual([]);
        const failure = new SessionFailure('crash', `Failed ${failed}`);
        if (failed === 'read') reply.reject(failure);
        else failSession(failure);
        const settled = await Promise.allSettled([
          reading, first, repeated, second, before.result, middle.result, after.result,
        ]);
        for (const outcome of settled) expect(outcome).toEqual({ status: 'rejected', reason: failure });
        expect(errors).toEqual([failure]);
        expect(coordinator.recoveryCommands.map((command) => command.kind === 'editCell' ? command.payload.args : command.kind))
          .toEqual([[0, 1, 0, 'before'], [0, 0, 0, firstValue], [0, 2, 0, 'middle'],
            [0, 0, 0, secondValue], [0, 3, 0, 'after']]);
        expect(coordinator.recoveryCommands[1]).toMatchObject({ payload: { target: {
          sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
        } } });
        expect(sent).toEqual([]);
      });
    }

    test(`retains ${source} input when the session fails before its active gate resumes`, async () => {
      const { coordinator, controller, sent, failSession } = setup();
      controller.start(source);
      controller.change(source, 'completed', null);
      const saving = coordinator.save();
      controller.end();
      const failure = new SessionFailure('crash', 'Worker stopped');
      queueMicrotask(() => failSession(failure));
      await expect(saving).rejects.toBe(failure);
      expect(sent).toEqual([]);
      expect(coordinator.recoveryCommands).toMatchObject([
        { kind: 'editCell', payload: { args: [0, 0, 0, 'completed'] } },
      ]);
    });
  }

  test('omits cancelled and unfinished compositions from recovery', async () => {
    const { coordinator, controller, sent, draft, failSession } = setup();
    const entered = deferred<void>();
    const reading = coordinator.runAfterPendingInput(() => {
      entered.resolve();
      return new Promise<void>(() => {});
    }, true);
    await entered.promise;
    controller.start('cell');
    controller.change('cell', 'cancelled', null);
    const cancelled = coordinator.save();
    controller.cancel();
    controller.start('formula');
    controller.change('formula', '=1+2', null);
    const unfinished = coordinator.save();
    const later = coordinator.submit(draft('later', 1));
    failSession(new SessionFailure('crash', 'Worker stopped'));
    await Promise.allSettled([reading, cancelled, unfinished, later.result]);
    expect(coordinator.recoveryCommands).toMatchObject([
      { kind: 'editCell', payload: { args: [0, 1, 0, 'later'] } },
    ]);
    expect(sent).toEqual([]);
  });

  test('does not duplicate a composition already dispatched when the session fails', async () => {
    const { coordinator, controller, sent, draft, failSession } = setup();
    const posted = deferred<void>();
    coordinator.subscribe((event) => { if (event.type === 'dispatched') posted.resolve(); });
    controller.start('cell');
    controller.change('cell', 'completed', null);
    const saving = coordinator.save();
    const later = coordinator.submit(draft('later', 1));
    controller.end();
    await posted.promise;
    failSession(new SessionFailure('crash', 'Worker stopped'));
    await Promise.allSettled([saving, later.result]);
    expect(sent).toHaveLength(1);
    expect(coordinator.recoveryCommands).toMatchObject([
      { kind: 'editCell', payload: { args: [0, 1, 0, 'later'] } },
    ]);
  });
});

describe('queued composition snapshots', () => {
  for (const source of ['cell', 'formula'] as const) {
    for (const successor of ['composition', 'typing'] as const) {
      test(`preserves completed ${source} input before later ${successor} and saves each boundary`, async () => {
        const selected = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
        const controller = new WorkerDraftController({
          generation: () => 1, identity: () => 'book', navigation: () => 0,
          selection: () => selected, sheet: () => 0,
          target: () => ({ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }),
          resolveSheet: () => 0, editable: () => true, prefill: async () => '',
          select: () => {}, selectionTarget: () => undefined, dependency: () => null,
          navigate: () => {}, committed: () => {}, fail: (error) => { throw error; }, report: (error) => { throw error; },
          focus: () => {},
        });
        const { coordinator, sent, saved, ack } = peerlessHarness(8, 8, controller);
        controller.attach(coordinator);
        coordinator.subscribe((event) => {
          if (event.type === 'dispatched' && event.seq !== null) queueMicrotask(() => ack(event.seq!));
        });
        const entered = deferred<void>();
        const reply = deferred<void>();
        const reading = coordinator.runAfterPendingInput(() => {
          entered.resolve();
          return reply.promise;
        }, true);
        await entered.promise;
        controller.start(source);
        controller.change(source, 'first', null);
        const first = coordinator.save();
        controller.end();
        await Promise.resolve();
        if (successor === 'composition') controller.start(source);
        controller.change(source, 'second', null);
        const second = coordinator.save();
        if (successor === 'composition') controller.end();
        await Promise.resolve();
        expect(saved).toEqual([]);
        reply.resolve();
        await Promise.all([reading, first, second]);
        expect(sent.map((command) => command.kind === 'editCell' ? command.payload.args[3] : command.kind))
          .toEqual(['first', 'second']);
        expect(saved).toEqual([1, 2]);
      });
    }
  }
});

describe('peerless worker input', () => {
  test('refuses overlapping plugin edits behind invalid input and accepts a correction', async () => {
    const { coordinator, session, draft, sent, ack } = peerlessHarness();
    const owner = { current: true, retiring: false, ready: true, failure: null,
      session: { ...session, editState: { readOnly: false } } } as unknown as PeerlessWorkbookSession;
    const document = createPeerlessPluginDocument({
      owner: () => owner, coordinator: () => coordinator, readOnly: () => false,
      commands: () => null, selection: () => null,
      navigator: { selectCells: async () => null, scrollToCell: async () => null },
    });
    const request = { expectVersion: 'v1', steps: [{ op: 'setCellInputs' as const,
      target: { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'B1' } }, inputs: [['host']] }] };
    coordinator.setDraft(draft('invalid'));
    const first = document.admitEdits!(owner, request, () => null, (write) => write());
    const second = document.admitEdits!(owner, request, () => null, (write) => write());
    for (const command of sent) ack(command.seq, command === sent[0] ? 'refused' : 'cancelled');
    const outcomes = await Promise.all([first, second]);
    expect(outcomes).toEqual([
      expect.objectContaining({ ok: false, code: 'input-failed' }),
      expect.objectContaining({ ok: false, code: 'input-failed' }),
    ]);
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'applyEdits', 'applyEdits']);
    expect(sent.slice(1).every((command) => command.afterSeq === 1)).toBe(true);
    expect(coordinator.error).toBeNull();
    coordinator.setDraft(draft('corrected'));
    const corrected = document.admitEdits!(owner, request, () => null, (write) => write());
    for (const command of sent.slice(3)) ack(command.seq, 'applied', { ok: true, version: 'v2', previews: [] });
    expect(await corrected).toMatchObject({ ok: true, value: { ok: true } });
  });

  for (const source of ['cell', 'formula'] as const) {
    for (const completion of ['success', 'cancel'] as const) {
      test(`retains the save reservation through real ${source} IME ${completion}`, async () => {
        const selected = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
        const controller = new WorkerDraftController({
          generation: () => 1, identity: () => 'book', navigation: () => 0,
          selection: () => selected, sheet: () => 0,
          target: () => ({ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }),
          resolveSheet: () => 0, editable: () => true, prefill: async () => '',
          select: () => {}, selectionTarget: () => undefined, dependency: () => null,
          navigate: () => {}, committed: () => {}, fail: (error) => { throw error; }, report: (error) => { throw error; },
          focus: () => {},
        });
        const { coordinator, sent, saved, ack } = peerlessHarness(8, 8, controller);
        controller.attach(coordinator);
        coordinator.subscribe((event) => {
          if (event.type === 'dispatched' && event.seq !== null) queueMicrotask(() => ack(event.seq!));
        });
        const entered = deferred<void>();
        const reply = deferred<void>();
        const reading = coordinator.runAfterPendingInput(() => {
          entered.resolve();
          return reply.promise;
        }, true);
        await entered.promise;
        controller.start(source);
        const first = coordinator.save();
        if (completion === 'success') {
          controller.change(source, 'completed', null);
          controller.end();
          await Promise.resolve();
        } else {
          controller.cancel();
          controller.change(source, 'fresh', null);
        }
        const second = coordinator.save();
        expect(saved).toEqual([]);
        reply.resolve();
        await Promise.all([reading, first, second]);
        expect(saved).toEqual(completion === 'success' ? [1, 1] : [0, 1]);
        expect(sent.map((command) => command.kind === 'editCell' ? command.payload.args[3] : command.kind))
          .toEqual([completion === 'success' ? 'completed' : 'fresh']);
      });
    }
  }

  test('save skips an untouched draft awaiting prefill', async () => {
    const { coordinator, draft, sent, saved, ack } = peerlessHarness();
    coordinator.setDraft({ ...draft(''), modified: false } as InputDraft);
    const saving = coordinator.save();
    const writes = sent.filter((command) => command.kind === 'editCell');
    for (const command of sent) ack(command.seq);
    await saving;
    expect(writes).toEqual([]);
    expect(saved).toEqual([0]);
  });

  test('cut reads after a pending predecessor and reserves its clear before later input', async () => {
    const { coordinator, sent, draft, ack } = peerlessHarness();
    let readThrough: number | undefined;
    const reading = deferred<string>();
    const readPosted = deferred<void>();
    const writing = deferred<void>();
    const posted = deferred<void>();
    const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
    const predecessor = coordinator.submit(draft('original'));
    coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.command?.kind === 'editCell' && event.command.payload.args[1] === 1) posted.resolve();
    });
    const copied = coordinator.runAfterPendingInput(() => {
      readThrough = coordinator.throughSeq;
      readPosted.resolve();
      return reading.promise;
    }, true);
    const written = copied.then(() => writing.promise);
    const cut = coordinator.clipboard(() => written, () => ({ kind: 'applyOps', payload: {
      toolbar: { id: 'clear', args: null, target }, selection: target,
    } }));
    const later = coordinator.submit(draft('later', 1));
    expect(readThrough).toBeUndefined();
    expect(sent).toHaveLength(1);
    expect(cut.seq).toBeNull();
    expect(later.seq).toBeNull();
    ack(predecessor.seq!);
    await readPosted.promise;
    expect(readThrough).toBe(predecessor.seq!);
    expect(sent).toHaveLength(1);
    expect(cut.seq).toBeNull();
    expect(later.seq).toBeNull();
    reading.resolve('original');
    await copied;
    expect(sent).toHaveLength(1);
    expect(cut.seq).toBeNull();
    expect(later.seq).toBeNull();
    writing.resolve();
    await posted.promise;
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'applyOps', 'editCell']);
    expect(sent[1]).toMatchObject({ payload: { toolbar: { id: 'clear', target }, selection: target } });
    expect(sent[2]).toMatchObject({ payload: { args: [0, 1, 0, 'later'] } });
    expect(sent[2].seq).toBeGreaterThan(sent[1].seq);
    ack(sent[1].seq); ack(sent[2].seq);
    await Promise.all([predecessor.result, cut.result, later.result]);
  });

  test('copy posts its read after a pending predecessor without reserving later input', async () => {
    const { coordinator, sent, draft, ack } = peerlessHarness();
    let readThrough: number | undefined;
    const reading = deferred<string>();
    const readPosted = deferred<void>();
    const posted = deferred<void>();
    const predecessor = coordinator.submit(draft('original'));
    coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.command?.kind === 'editCell' && event.command.payload.args[1] === 1) posted.resolve();
    });
    const copied = coordinator.runAfterPendingInput(() => {
      readThrough = coordinator.throughSeq;
      readPosted.resolve();
      return reading.promise;
    }, 'posted');
    const later = coordinator.submit(draft('later', 1));
    expect(readThrough).toBeUndefined();
    expect(sent).toHaveLength(1);
    expect(later.seq).toBeNull();
    ack(predecessor.seq!);
    await readPosted.promise;
    expect(readThrough).toBe(predecessor.seq!);
    await posted.promise;
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'editCell']);
    expect(sent[1]).toMatchObject({ payload: { args: [0, 1, 0, 'later'] } });
    expect(later.seq).toBeGreaterThan(predecessor.seq!);
    reading.resolve('original');
    expect(await copied).toBe('original');
    ack(later.seq!);
    await Promise.all([predecessor.result, later.result]);
  });

  test('replacement cancels pending cut preparation and leaves new input intact', async () => {
    const { coordinator, session, state, sent, draft, ack } = peerlessHarness();
    const writing = deferred<void>();
    const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
    const clears: string[] = [];
    const cut = coordinator.clipboard(() => writing.promise, () => {
      clears.push(target.sheetId);
      return { kind: 'applyOps', payload: { toolbar: { id: 'clear', args: null, target }, selection: target } };
    });
    expect(sent).toEqual([]);
    state.generation += 1;
    session.commandIdentity = { session: 'replacement', incarnation: 'next' };
    state.throughSeq = 0;
    const next = coordinator.submit(draft('new'));
    await expect(cut.result).rejects.toMatchObject({ code: 'document-replaced' });
    writing.resolve();
    await writing.promise;
    expect(clears).toEqual([]);
    expect(sent).toHaveLength(1);
    expect(sent[0]).toMatchObject({ session: 'replacement', payload: { args: [0, 0, 0, 'new'] } });
    ack(next.seq!);
    await next.result;
    expect(coordinator.error).toBeNull();
  });

  test('failure cancels pending cut preparation', async () => {
    const { coordinator, sent } = peerlessHarness();
    const writing = deferred<void>();
    const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
    const clears: string[] = [];
    const cut = coordinator.clipboard(() => writing.promise, () => {
      clears.push(target.sheetId);
      return { kind: 'applyOps', payload: { toolbar: { id: 'clear', args: null, target }, selection: target } };
    });
    const failure = new Error('Worker stopped');
    coordinator.fail(failure);
    await expect(cut.result).rejects.toBe(failure);
    writing.resolve();
    await writing.promise;
    expect(clears).toEqual([]);
    expect(sent).toEqual([]);
  });

  test('overlapping clipboard captures retain their admission frame and selection revision', async () => {
    const { coordinator, state, sent, events, ack } = peerlessHarness();
    const first = deferred<string>();
    const second = deferred<string>();
    const frame = { generation: 1, navigation: 0, width: 800, height: 600, sheet: 0, zoom: 1, dpr: 1,
      viewport: { x: 0, y: 0, width: 800, height: 600 }, requestGeneration: 1 };
    state.frame = frame;
    const pasted = coordinator.clipboard(() => first.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 0, 0, value] },
    }));
    state.navigation = 1;
    state.frame = { ...frame, viewport: { ...frame.viewport, x: 96 }, requestGeneration: 2 };
    const admittedFrame = structuredClone(state.frame);
    const overlapping = coordinator.clipboard(() => second.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 0, 1, value] },
    }));
    state.navigation = 2;
    state.frame.viewport.x = 192;
    state.frame.requestGeneration = 3;
    const firstPosted = deferred<void>();
    const posted = deferred<void>();
    coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.seq === 1) firstPosted.resolve();
      if (event.type === 'dispatched' && event.seq === 2) posted.resolve();
    });
    first.resolve('first'); second.resolve('second');
    await firstPosted.promise;
    ack(1);
    await posted.promise;
    expect(sent[1].frame).toEqual(admittedFrame);
    expect(events.find((event) => event.type === 'dispatched' && event.seq === 2))
      .toMatchObject({ selectionRevision: 1 });
    ack(2);
    await Promise.all([pasted.result, overlapping.result]);
  });

  test('command inputs match the worker protocol', () => {
    type Input<C> = C extends WorkbookCommand ? Pick<C, 'kind' | 'payload'> : never;
    const matches: PeerlessWorkerCommandInput extends Input<WorkbookCommand> ? true : false = true;
    const missingTarget: { kind: 'applyOps'; payload: { toolbar: { id: 'bold'; args: null } } } extends
      PeerlessWorkerCommandInput ? true : false = false;
    const save: { kind: 'applyOps'; payload: { toolbar: { id: 'save'; args: null } } } extends
      PeerlessWorkerCommandInput ? true : false = false;
    const refreshTarget: { kind: 'applyOps'; payload: { toolbar: { id: 'refreshProposals'; args: null;
      target: { sheetId: string; range: { kind: 'a1'; a1: string } } } } } extends
      PeerlessWorkerCommandInput ? true : false = false;
    expect([matches, missingTarget, save, refreshTarget]).toEqual([true, false, false, false]);
  });

  test('read-only admission permits proposal refresh and refuses formatting', async () => {
    const { coordinator, sent, ack } = peerlessHarness();
    const mode = coordinator.setReadOnly(true);
    ack(1);
    await mode.result;
    const refresh = coordinator.command({ kind: 'applyOps', payload: {
      toolbar: { id: 'refreshProposals', args: null },
    } });
    expect(sent[1]).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'refreshProposals' } } });
    const formatting = coordinator.command({ kind: 'patchRangeStyle', payload: { args: [0, 'A1', { bold: true }] } });
    await expect(formatting.result).rejects.toMatchObject({ code: 'input-failed' });
    ack(2);
    await refresh.result;
    expect(sent).toHaveLength(2);
  });

  test('applied frame fence preserves the native draft and composition', async () => {
    const { coordinator, sent, state, draft, ack } = peerlessHarness();
    const first = coordinator.submit(draft('committed'));
    const composing = draft('native', 1);
    coordinator.setDraft(composing);
    state.seal = { composition: deferred<boolean>().promise };
    expect(coordinator.submittedThroughSeq).toBe(1);
    let applied = false;
    const frame = coordinator.whenApplied(1).then(() => { applied = true; });
    await Promise.resolve();
    expect(applied).toBe(false);
    ack(1);
    await Promise.all([first.result, frame]);
    expect(sent).toHaveLength(1);
    expect(coordinator.draft?.value).toBe('native');
  });

  test('posts rapid commits before predecessor ack', async () => {
    const { coordinator, sent, draft, ack } = peerlessHarness();
    const first = coordinator.submit(draft('first'));
    const second = coordinator.submit(draft('second', 1), { after: first });
    expect(sent.map((command) => command.seq)).toEqual([1, 2]);
    expect(sent[1].afterSeq).toBe(1);
    expect(sent.map((command) => command.basis.version)).toEqual(['v1', 'v1']);
    ack(1); ack(2);
    await Promise.all([first.result, second.result]);
  });

  test('fences clipboard mode and save', async () => {
    const { coordinator, sent, saved, draft, ack } = peerlessHarness();
    const capture = deferred<string>();
    const first = coordinator.submit(draft('before'));
    const pasted = coordinator.clipboard(() => capture.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 1, 0, value] },
    }));
    const mode = coordinator.setReadOnly(true);
    const saving = coordinator.save();
    const modePosted = deferred<void>();
    const unsubscribe = coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.command?.kind === 'setReadOnly') modePosted.resolve();
    });
    expect(sent).toHaveLength(1);
    expect(saved).toEqual([]);
    capture.resolve('paste');
    await modePosted.promise;
    unsubscribe();
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'editCell', 'setReadOnly']);
    expect(sent[1].afterSeq).toBe(1);
    expect(sent[2].afterSeq).toBeUndefined();
    ack(1); ack(2); ack(3);
    await Promise.all([first.result, pasted.result, mode.result, saving]);
    expect(saved).toEqual([3]);
  });

  test('stale acceptance refusal sends queued independent navigation and its typed edit', async () => {
    const { coordinator, sent, ack } = peerlessHarness(1);
    const proposal = coordinator.command({ kind: 'acceptProposal', payload: { args: ['stale', {}] } });
    const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'B1' } };
    let navigation!: ReturnType<typeof coordinator.command>;
    const selecting = coordinator.runAfterPendingInput(() => {
      navigation = coordinator.command({ kind: 'select', payload: { selection: selected } });
      return navigation.result;
    });
    const original: InputDraft = { generation: 1, sheet: 0, row: 0, col: 1, source: 'cell', value: 'kept text' };
    const edited = coordinator.submit(original, { after: navigation, selection: {
      sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B2' },
    } });
    expect(sent).toHaveLength(1);
    ack(proposal.seq!, 'refused');
    await expect(proposal.result).rejects.toMatchObject({ code: 'input-failed' });
    expect(sent[1]).toMatchObject({ kind: 'select', payload: { selection: selected } });
    expect(sent[1].afterSeq).toBeUndefined();
    ack(sent[1].seq);
    expect(await selecting).toMatchObject({ status: 'applied' });
    expect(sent[2]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 1, 'kept text'] } });
    ack(sent[2].seq);
    expect(await edited.result).toMatchObject({ status: 'applied' });
    expect(coordinator.rejected).toEqual([]);
    expect(coordinator.pending).toBe(false);
  });

  test('cancels only refused dependents', async () => {
    const { coordinator, sent, events, draft, ack } = peerlessHarness(1);
    const original = draft('restore me');
    const first = coordinator.submit(original);
    const dependent = coordinator.submit(draft('dependent', 1), { after: first });
    const independent = coordinator.submit(draft('newer', 2));
    coordinator.setDraft(draft('focused elsewhere', 3));
    ack(1, 'refused');
    await expect(first.result).rejects.toMatchObject({ code: 'input-failed' });
    await expect(dependent.result).rejects.toMatchObject({ code: 'input-failed' });
    expect(sent.map((command) => command.seq)).toEqual([1, 2]);
    expect(events.find((event) => event.type === 'refused')).toMatchObject({ seq: 1, draft: original });
    expect(coordinator.draft?.value).toBe('focused elsewhere');
    expect(coordinator.rejected).toEqual([original]);
    ack(2);
    await independent.result;
  });

  test('delayed Save fences paste but excludes the later Enter commit', async () => {
    const { coordinator, sent, saved, draft, ack } = peerlessHarness();
    const capture = deferred<string>();
    const pasted = coordinator.clipboard(() => capture.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 0, 0, value] },
    }));
    const saving = coordinator.runAfterPendingInput(() => coordinator.save());
    const later = coordinator.submit(draft('later', 1));
    const posted = deferred<void>();
    const laterPosted = deferred<void>();
    coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.seq === 1) posted.resolve();
      if (event.type === 'dispatched' && event.seq === 2) laterPosted.resolve();
    });
    capture.resolve('paste');
    await posted.promise;
    expect(sent).toHaveLength(1);
    ack(1);
    await saving;
    expect(saved).toEqual([1]);
    await pasted.result;
    await laterPosted.promise;
    expect(sent[1]).toMatchObject({ kind: 'editCell', payload: { args: [0, 1, 0, 'later'] } });
    ack(2);
    await later.result;
  });

  test('selection retains its admission revision while clipboard capture is pending', async () => {
    const { coordinator, state, events, ack } = peerlessHarness();
    const capture = deferred<string>();
    const pasted = coordinator.clipboard(() => capture.promise, (value) => ({
      kind: 'editCell', payload: { args: [0, 0, 0, value] },
    }));
    state.navigation = 1;
    const selected = coordinator.command({ kind: 'select', payload: { selection: {
      sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B2' },
    } } });
    state.navigation = 2;
    const posted = deferred<void>();
    coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.seq === 2) posted.resolve();
    });
    capture.resolve('paste');
    await posted.promise;
    expect(events.find((event) => event.type === 'dispatched' && event.seq === 2))
      .toMatchObject({ selectionRevision: 1 });
    ack(1); ack(2);
    await Promise.all([pasted.result, selected.result]);
  });

  test('discarded refusal releases independent selection formatting and Save', async () => {
    const { coordinator, sent, saved, draft, ack } = peerlessHarness();
    const original = draft('refused');
    const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A2' } };
    const first = coordinator.submit(original, { selection: selected });
    ack(1, 'refused');
    await expect(first.result).rejects.toMatchObject({ code: 'input-failed' });
    coordinator.discardRejected(original);
    const navigation = coordinator.command({ kind: 'select', payload: { selection: selected } });
    const formatting = coordinator.command({ kind: 'patchRangeStyle', payload: { args: [0, 'A2', { bold: true }] } });
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'select', 'patchRangeStyle']);
    expect(sent[1].afterSeq).toBeUndefined();
    expect(sent[2].afterSeq).toBeUndefined();
    ack(2); ack(3);
    await Promise.all([navigation.result, formatting.result]);
    await coordinator.save();
    expect(saved).toEqual([3]);
    expect(coordinator.rejected).toEqual([]);
  });

  test('replacement invalidates late replies', async () => {
    const { coordinator, session, sent, events, state, draft, ack, replies, listeners } = peerlessHarness();
    const old = coordinator.submit(draft('old'));
    const oldReply = replies.get(1)!;
    const oldListener = [...listeners][0];
    state.generation += 1;
    session.commandIdentity = { session: 'replacement', incarnation: 'next' };
    state.throughSeq = 0;
    const next = coordinator.submit(draft('new'));
    await expect(old.result).rejects.toMatchObject({ code: 'document-replaced' });
    const count = events.length;
    const oldReceipt = { seq: 1, status: 'applied' as const };
    oldListener({ throughSeq: 1, version: 'old', statuses: [oldReceipt] });
    oldReply.resolve(oldReceipt);
    await Promise.resolve();
    expect(events).toHaveLength(count);
    expect(sent[1]).toMatchObject({ session: 'replacement', seq: 1 });
    ack(next.seq!);
    await next.result;
  });

  test('coalesces only unsent compatible drafts and bounds storage', async () => {
    const { coordinator, sent, draft, ack } = peerlessHarness(1, 2);
    const first = coordinator.submit(draft('sent'));
    const old = coordinator.submit(draft('old', 1));
    const latest = coordinator.submit(draft('latest', 1));
    const fence = coordinator.setReadOnly(false);
    const overflow = coordinator.submit(draft('overflow', 2));
    await expect(overflow.result).rejects.toMatchObject({ code: 'input-failed' });
    ack(1);
    expect(sent[1]).toMatchObject({ payload: { args: [0, 1, 0, 'latest'] } });
    ack(2); ack(3);
    await Promise.all([first.result, old.result, latest.result, fence.result]);
  });

  test('keeps composition local and commits its captured target once', async () => {
    const { coordinator, sent, draft, ack } = peerlessHarness();
    const composed = draft('変換', 4);
    const first = coordinator.submit(composed, { compositionId: 'ime:1' });
    const trailing = coordinator.submit({ ...composed }, { compositionId: 'ime:1' });
    const cancelled = coordinator.submit(draft(''), { compositionId: 'ime:2' });
    expect(trailing).toBe(first);
    expect(sent).toHaveLength(1);
    expect(sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 4, 0, '変換'] } });
    ack(1);
    await Promise.all([first.result, cancelled.result]);
  });

  test('worker failure rejects pending work without replay', async () => {
    const { coordinator, sent, errors, draft, replies } = peerlessHarness(1);
    const first = coordinator.submit(draft('sent'));
    const queued = coordinator.submit(draft('queued', 1));
    const failure = new Error('Worker stopped');
    replies.get(1)!.reject(failure);
    await expect(first.result).rejects.toBe(failure);
    await expect(queued.result).rejects.toBe(failure);
    await expect(coordinator.save()).rejects.toBe(failure);
    expect(coordinator.readOnly).toBe(true);
    expect(sent).toHaveLength(1);
    expect(errors).toEqual([failure]);
  });

  test('refuses before first paint without consuming a sequence', async () => {
    const { coordinator, state, draft, sent, ack, errors } = peerlessHarness();
    state.ready = false;
    const value = draft('waiting for paint');
    coordinator.setDraft(value);
    const refused = coordinator.submit(value);
    await expect(refused.result).rejects.toMatchObject({ code: 'editor-unavailable' });
    expect(refused.seq).toBeNull();
    expect(sent).toEqual([]);
    expect(coordinator.draft?.value).toBe(value.value);
    state.ready = true;
    const accepted = coordinator.submit(value);
    expect(accepted.seq).toBe(1);
    ack(1);
    await accepted.result;
    expect(errors).toEqual([]);
  });

  test('mode seals composition without waiting for its commit ack', async () => {
    const { coordinator, state, sent, draft, ack } = peerlessHarness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('preedit', 4));
    const mode = coordinator.setReadOnly(true);
    expect(coordinator.readOnly).toBe(true);
    expect(sent).toEqual([]);
    coordinator.setDraft(draft('変換', 9));
    const posted = deferred<void>();
    const unsubscribe = coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.command?.kind === 'setReadOnly') posted.resolve();
    });
    composition.resolve(true);
    await posted.promise;
    unsubscribe();
    expect(sent.map((command) => command.kind)).toEqual(['editCell', 'setReadOnly']);
    expect(sent[0]).toMatchObject({ payload: { args: [0, 4, 0, '変換'] } });
    ack(1); ack(2);
    await mode.result;
  });

  test('save keeps later input behind a composing draft', async () => {
    const { coordinator, state, sent, saved, draft, ack } = peerlessHarness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('変換', 4));
    const saving = coordinator.save();
    const later = coordinator.submit(draft('later', 6));
    const posted = deferred<void>();
    const unsubscribe = coordinator.subscribe((event) => {
      if (event.type === 'dispatched' && event.seq === 1) posted.resolve();
    });
    composition.resolve(true);
    await posted.promise;
    unsubscribe();
    expect(sent).toHaveLength(1);
    expect(sent[0]).toMatchObject({ payload: { args: [0, 4, 0, '変換'] } });
    expect(saved).toEqual([]);
    ack(1);
    await saving;
    expect(saved).toEqual([1]);
    const laterPosted = deferred<void>();
    if (later.seq === null) {
      const off = coordinator.subscribe((event) => {
        if (event.type === 'dispatched' && event.seq === 2) { off(); laterPosted.resolve(); }
      });
    } else laterPosted.resolve();
    await laterPosted.promise;
    ack(2);
    await later.result;
  });
});

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, failed) => {
    resolve = done;
    reject = failed;
  });
  return { promise, resolve, reject };
}

function harness() {
  const log: string[] = [];
  const previews: InputDraft[] = [];
  const writes: InputDraft[] = [];
  const errors: unknown[] = [];
  const hydration: string[] = [];
  const focus: string[] = [];
  const state = {
    generation: 1,
    sheet: 0,
    target: 'A1',
    ready: false,
    peer: deferred<void>(),
    paint: null as Promise<void> | null,
    previewEntered: deferred<void>(),
    shown: null as string | null,
    seal: {} as InputSeal,
    accept: true,
    refuse: new Set<string>(),
    writeError: null as Error | null,
    onSeal: null as (() => void) | null,
    onHydration: null as ((reason: string) => void | Promise<void>) | null,
    onFlushEdits: null as (() => Promise<void>) | null,
    onWrite: null as ((draft: InputDraft) => void) | null,
  };
  let coordinator!: WorkerInputCoordinator;
  const hooks = {
    generation: () => state.generation,
    capture: () => ({ sheet: state.sheet, target: state.target }),
    isReady: () => state.ready,
    whenReady: () => (state.ready ? Promise.resolve() : state.peer.promise),
    seal: () => {
      state.onSeal?.();
      return state.seal;
    },
    sync: () => {
      const current = coordinator.draft;
      if (current && state.shown !== null && current.value !== state.shown) {
        coordinator.setDraft({ ...current, value: state.shown });
      }
    },
    preview: (draft: InputDraft) => {
      previews.push(draft);
      state.previewEntered.resolve();
      return state.paint ?? Promise.resolve();
    },
    write: (draft: InputDraft) => {
      if (state.writeError) throw state.writeError;
      if (!state.accept || state.refuse.has(draft.value)) {
        log.push(`refused ${draft.value}`);
        return false;
      }
      writes.push(draft);
      log.push(`write ${draft.value}`);
      state.onWrite?.(draft);
      return true;
    },
    close: () => focus.push('grid'),
    restore: () => {
      focus.push('cell');
      return true;
    },
    requestHydration: (reason: string) => {
      hydration.push(reason);
      return state.onHydration?.(reason);
    },
    flushEdits: () => state.onFlushEdits?.() ?? Promise.resolve(),
    onError: (error: unknown) => errors.push(error),
  };
  coordinator = createWorkerInputCoordinator(hooks);
  const draft = (
    value: string,
    row = 0,
    sheet = state.sheet,
    source: InputDraft['source'] = 'cell'
  ) => {
    state.shown = value;
    return { generation: state.generation, sheet, row, col: 0, value, source } satisfies InputDraft;
  };
  const ready = () => {
    state.ready = true;
    state.peer.resolve();
  };
  return { coordinator, log, state, draft, ready, previews, writes, errors, hydration, focus };
}

async function errorOf(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
    return null;
  } catch (error) {
    return error instanceof XlsxCommandAdmissionError ? error.code : error;
  }
}

describe('worker input coordinator', () => {
  test('awaits a sealed live draft acknowledgement before a dependent host edit and propagates refusal', async () => {
    for (const refused of [false, true]) {
      const acknowledgement = deferred<void>();
      const waiting = deferred<void>();
      const log: string[] = [];
      const refusal = new WorkerInputRefusal('A1 was refused');
      const coordinator = createWorkerInputCoordinator({
        generation: () => 1,
        capture: () => ({ sheet: 0, target: 'B1' }),
        isReady: () => true,
        whenReady: async () => {},
        seal: () => ({}),
        sync: () => {},
        preview: async () => {},
        write: (draft) => { log.push(`A1:${draft.value}`); return true; },
        acknowledgeEdits: () => {
          if (!log.length) return Promise.resolve();
          waiting.resolve();
          return acknowledgement.promise;
        },
        requestHydration: () => {},
        flushEdits: async () => {},
      });
      coordinator.setDraft({ generation: 1, sheet: 0, row: 0, col: 0, value: 'live', source: 'cell' });
      const dependent = coordinator.runAfterPendingInput(() => { log.push('B1:dependent'); }, { kind: 'host' });
      const result = dependent.catch((error: unknown) => error);
      await waiting.promise;
      expect(log).toEqual(['A1:live']);
      if (refused) acknowledgement.reject(refusal);
      else acknowledgement.resolve();
      expect(await result).toBe(refused ? refusal : undefined);
      expect(log).toEqual(refused ? ['A1:live'] : ['A1:live', 'B1:dependent']);
      expect(coordinator.unapplied).toEqual([]);
    }
  });

  test('replays every typed key accepted before readiness in order', async () => {
    const { coordinator, draft, ready, log, hydration } = harness();
    let text = '';
    const commits: Promise<void>[] = [];
    for (const [row, key] of ['a', 'b', 'c'].entries()) {
      text += key;
      const typed = draft(text, row, 0, row === 2 ? 'formula' : 'cell');
      coordinator.setDraft(typed);
      expect(coordinator.draft?.value).toBe(text);
      commits.push(coordinator.submitAsync(typed));
    }
    expect(log).toEqual([]);
    expect(coordinator.committed.map((entry) => entry.value)).toEqual(['a', 'ab', 'abc']);
    ready();
    await Promise.all(commits);
    expect(log).toEqual(['write a', 'write ab', 'write abc']);
    expect(hydration).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('keeps queued commits separate from repeated live draft updates', async () => {
    const { coordinator, draft, ready, writes } = harness();
    const first = draft('first');
    coordinator.setDraft(first);
    expect(coordinator.submit(first)).toBe(true);
    const second = draft('second', 1);
    coordinator.setDraft(second);
    const committed = coordinator.submitAsync(second);
    coordinator.setDraft(draft('l', 2));
    coordinator.setDraft(draft('la', 2));
    const live = draft('later', 2);
    coordinator.setDraft(live);
    ready();
    await committed;
    expect(writes).toEqual([first, second]);
    expect(coordinator.draft).toBe(live);
  });

  test('awaits peer readiness and preview paint before the facade write', async () => {
    const { coordinator, draft, ready, state, previews, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const committed = coordinator.submitAsync(draft('typed'));
    expect(previews).toEqual([]);
    ready();
    await state.previewEntered.promise;
    expect(previews.map((entry) => entry.value)).toEqual(['typed']);
    expect(log).toEqual([]);
    paint.resolve();
    await committed;
    expect(log).toEqual(['write typed']);
  });

  test('captures clipboard payload and target during the original gesture', async () => {
    const { coordinator, state, ready, log } = harness();
    const payload = { text: 'pasted', cells: [['one']] };
    let inGesture = true;
    let captures = 0;
    const paste = coordinator.clipboard(
      () => {
        expect(inGesture).toBe(true);
        captures += 1;
        return payload;
      },
      (input, intent) => {
        expect(input).toEqual({ text: 'pasted', cells: [['one']] });
        expect(intent.generation).toBe(1);
        expect(intent.sheet).toBe(0);
        expect(intent.target).toBe('A1');
        log.push(input.text);
      }
    );
    inGesture = false;
    payload.text = 'changed';
    payload.cells[0][0] = 'changed';
    state.sheet = 4;
    state.target = 'Z9';
    expect(captures).toBe(1);
    ready();
    await paste;
    expect(log).toEqual(['pasted']);
    expect(captures).toBe(1);
  });

  test('shares one FIFO across commits, clipboard, charts, commands, hosts and plugins', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const commit = coordinator.submitAsync(draft('typed'));
    const paste = coordinator.clipboard(() => 'paste', (input) => log.push(input));
    const chartInput = { chartId: 'chart-1', position: { x: 10, y: 20 } };
    const chart = coordinator.runAfterPendingInput(
      (intent) => {
        expect(intent.input).toEqual({ chartId: 'chart-1', position: { x: 10, y: 20 } });
        expect(intent.sheet).toBe(0);
        expect(intent.target).toBe('A1');
        log.push('chart');
      },
      { kind: 'chart', input: chartInput }
    );
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    const host = coordinator.runAfterPendingInput(() => log.push('host'), { kind: 'host' });
    const plugin = coordinator.runAfterPendingInput(() => log.push('plugin'), { kind: 'plugin' });
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual([
      'commit',
      'clipboard',
      'chart',
      'command',
      'host',
      'plugin',
    ]);
    chartInput.position.x = 999;
    state.sheet = 3;
    state.target = 'B4';
    ready();
    await Promise.all([commit, paste, chart, command, host, plugin]);
    expect(log).toEqual(['write typed', 'paste', 'chart', 'command', 'host', 'plugin']);
  });

  test('does not move focus or overwrite live typing when readiness completes', async () => {
    const { coordinator, draft, ready, focus } = harness();
    const committed = coordinator.submitAsync(draft('cell'));
    const live = draft('formula still typing', 1, 0, 'formula');
    coordinator.setDraft(live);
    focus.push('formula');
    ready();
    await committed;
    expect(focus).toEqual(['formula']);
    expect(coordinator.draft).toBe(live);
  });

  test('reports failure once and retains all unapplied intents for recovery', async () => {
    const { coordinator, draft, state, ready, log, errors, hydration } = harness();
    const commit = coordinator.submitAsync(draft('typed'));
    const paste = coordinator.clipboard(() => 'paste', (input) => log.push(input));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    const accepted = coordinator.unapplied;
    const failed = new Error('worker failed');
    coordinator.fail(failed);
    coordinator.fail(new Error('same session'));
    expect(await errorOf(commit)).toBe(failed);
    expect(await errorOf(paste)).toBe(failed);
    expect(await errorOf(command)).toBe(failed);
    expect(coordinator.unapplied).toEqual(accepted);
    const late = coordinator.submitAsync(draft('accepted after failure', 1));
    expect(await errorOf(late)).toBe(failed);
    expect(errors).toEqual([failed]);
    expect(coordinator.error).toBe(failed);
    state.onHydration = ready;
    await coordinator.recover();
    expect(log).toEqual(['write typed', 'paste', 'command', 'write accepted after failure']);
    expect(hydration).toEqual(['recovery']);
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.error).toBeNull();
  });

  test('drains accepted intents before replacement and keeps new input in its generation', async () => {
    const { coordinator, draft, state, ready, log, previews } = harness();
    const oldPeer = state.peer;
    const commit = coordinator.submitAsync(draft('old'));
    const command = coordinator.runAfterPendingInput(() => log.push('old command'));
    const draining = coordinator.drain();
    expect(log).toEqual([]);
    ready();
    await draining;
    await commit;
    await command;
    expect(log).toEqual(['write old', 'old command']);
    state.generation += 1;
    coordinator.reset();
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.draft).toBeNull();
    expect(coordinator.pending).toBe(false);
    state.ready = false;
    state.peer = deferred<void>();
    const next = coordinator.submitAsync(draft('new'));
    oldPeer.resolve();
    ready();
    await next;
    expect(log).toEqual(['write old', 'old command', 'write new']);
    expect(previews.map((entry) => entry.value)).toEqual(['old', 'new']);
  });

  test('prevents synchronous host mutations from overtaking queued work', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const commit = coordinator.submitAsync(draft('typed'));
    expect(() => coordinator.runSync(() => log.push('early host'))).toThrow(WorkerInputNotReadyError);
    ready();
    await state.previewEntered.promise;
    expect(() => coordinator.runSync(() => log.push('early host'))).toThrow(WorkerInputNotReadyError);
    paint.resolve();
    await commit;
    await coordinator.flush();
    expect(coordinator.pending).toBe(false);
    expect(coordinator.runSync(() => log.push('host'))).toBe(2);
    expect(coordinator.runSync(() => log.push('next host'))).toBe(3);
    expect(log).toEqual(['write typed', 'host', 'next host']);
  });

  test('keeps a refused commit and newer typing until a correction is written', async () => {
    const { coordinator, draft, ready, state, log, focus } = harness();
    const refused = draft('refused', 0, 0);
    const commit = coordinator.submitAsync(refused);
    const live = draft('newer', 1, 1);
    coordinator.setDraft(live);
    state.accept = false;
    ready();
    expect(await errorOf(commit)).toBe('input-failed');
    expect(coordinator.draft).toBe(live);
    expect(coordinator.rejected).toEqual([refused]);
    expect(coordinator.unapplied.map((entry) => entry.input)).toEqual([refused]);
    expect(await errorOf(coordinator.runAfterPendingInput(() => log.push('blocked')))).toBe(
      'input-failed'
    );
    state.accept = true;
    const corrected = draft('corrected', 0, 0);
    coordinator.setDraft(corrected);
    await coordinator.submitAsync(corrected);
    expect(coordinator.rejected).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
    expect(focus).toEqual([]);
    expect(log).toEqual(['refused refused', 'write corrected']);
  });

  test('a queued correction supersedes a refusal without restoring over live input', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    state.refuse.add('bad');
    const refused = coordinator.submitAsync(draft('bad'));
    const corrected = draft('corrected');
    coordinator.setDraft(corrected);
    expect(coordinator.settle()).toBe(true);
    const settled = coordinator.submitAsync(corrected);
    const live = draft('live', 1);
    coordinator.setDraft(live);
    ready();
    expect(await errorOf(refused)).toBe('input-failed');
    await settled;
    expect(coordinator.draft).toBe(live);
    await coordinator.flush();
    expect(coordinator.rejected).toEqual([]);
    expect(coordinator.unapplied).toEqual([]);
    expect(log).toEqual(['refused bad', 'write corrected', 'write live']);
  });

  test('takes final composition text once before queued commands', async () => {
    const { coordinator, draft, state, ready, log, previews } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const first = coordinator.runAfterPendingInput((intent) => {
      expect(intent.sheet).toBe(0);
      expect(intent.target).toBe('A1');
      log.push('first');
    });
    const second = coordinator.runAfterPendingInput(() => log.push('second'));
    state.shown = '日本語';
    state.sheet = 2;
    state.target = 'C3';
    composition.resolve(true);
    ready();
    await Promise.all([first, second]);
    expect(log).toEqual(['write 日本語', 'first', 'second']);
    expect(previews).toHaveLength(1);
  });

  test('refuses commands when composition ends without its input or a gesture is active', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    state.seal = { refused: 'gesture-active' };
    expect(await errorOf(coordinator.runAfterPendingInput(() => log.push('gesture')))).toBe(
      'gesture-active'
    );
    expect(coordinator.unapplied).toEqual([]);
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    composition.resolve(false);
    ready();
    expect(await errorOf(command)).toBe('input-failed');
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
  });

  test('writes a draft once when commands and Enter accept the same draft', async () => {
    const { coordinator, draft, ready, log, previews } = harness();
    const typed = draft('typed');
    coordinator.setDraft(typed);
    const first = coordinator.runAfterPendingInput(() => log.push('first'));
    const second = coordinator.runAfterPendingInput(() => log.push('second'));
    const enter = coordinator.submitAsync(typed);
    ready();
    await Promise.all([first, second, enter]);
    expect(log).toEqual(['write typed', 'first', 'second']);
    expect(previews).toHaveLength(1);
  });

  test('keeps final composed text ahead of a command when Enter commits first', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    coordinator.setDraft(draft('日本語'));
    const enter = coordinator.submitAsync(coordinator.draft!);
    state.seal = {};
    composition.resolve(true);
    ready();
    await Promise.all([command, enter]);
    expect(log).toEqual(['write 日本語', 'command']);
  });

  test('fails a command behind input landed by its own seal', async () => {
    const { coordinator, state, ready, log } = harness();
    state.onSeal = () => {
      void coordinator.input(() => false);
    };
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    ready();
    expect(await errorOf(command)).toBe('input-failed');
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['input', 'command']);
  });

  test('recovers a failed command without writing its already applied draft twice', async () => {
    const { coordinator, draft, ready, log } = harness();
    const failed = new Error('command failed');
    let refuse = true;
    coordinator.setDraft(draft('typed'));
    const command = coordinator.runAfterPendingInput(() => {
      if (refuse) throw failed;
      log.push('command');
    });
    ready();
    expect(await errorOf(command)).toBe(failed);
    expect(log).toEqual(['write typed']);
    expect(coordinator.committed).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
    refuse = false;
    await coordinator.recover();
    expect(log).toEqual(['write typed', 'command']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('does not replay an applied edit whose asynchronous acknowledgement fails', async () => {
    const { coordinator, ready, log } = harness();
    const acknowledgement = deferred<void>();
    const entered = deferred<void>();
    const edit = coordinator.runAfterPendingInput(
      (_intent, markApplied) => {
        log.push('edit');
        markApplied();
        entered.resolve();
        return acknowledgement.promise;
      },
      { kind: 'host' }
    );
    const after = coordinator.runAfterPendingInput(() => log.push('after'));
    ready();
    await entered.promise;
    expect(() => coordinator.runSync(() => log.push('overtake'))).toThrow(WorkerInputNotReadyError);
    const failed = new Error('acknowledgement failed');
    acknowledgement.reject(failed);
    expect(await errorOf(edit)).toBe(failed);
    expect(await errorOf(after)).toBe(failed);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
    await coordinator.recover();
    expect(log).toEqual(['edit', 'after']);
  });

  test('retains a draft when preview fails and retries it during recovery', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const commit = coordinator.submitAsync(draft('typed'));
    ready();
    await state.previewEntered.promise;
    const failed = new Error('preview failed');
    paint.reject(failed);
    expect(await errorOf(commit)).toBe(failed);
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['commit']);
    state.paint = null;
    await coordinator.recover();
    expect(log).toEqual(['write typed']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('retains clipboard capture after failure without reading the clipboard again', async () => {
    const { coordinator, state, ready, log } = harness();
    const captured = deferred<{ text: string }>();
    let reads = 0;
    const paste = coordinator.clipboard(
      () => {
        reads += 1;
        return captured.promise;
      },
      (input) => log.push(input.text)
    );
    const failed = new Error('worker failed');
    coordinator.fail(failed);
    expect(await errorOf(paste)).toBe(failed);
    captured.resolve({ text: 'retained clipboard' });
    expect(await coordinator.unapplied[0].input).toEqual({ text: 'retained clipboard' });
    state.onHydration = ready;
    await coordinator.recover();
    expect(reads).toBe(1);
    expect(log).toEqual(['retained clipboard']);
  });

  test('drains an accepted write across its preview before replacement', async () => {
    const { coordinator, draft, ready, state, log } = harness();
    const paint = deferred<void>();
    state.paint = paint.promise;
    const commit = coordinator.submitAsync(draft('old'));
    ready();
    await state.previewEntered.promise;
    const draining = coordinator.drain();
    expect(log).toEqual([]);
    paint.resolve();
    await draining;
    await commit;
    expect(log).toEqual(['write old']);
    state.generation += 1;
    coordinator.reset();
    state.paint = null;
    await coordinator.submitAsync(draft('new'));
    expect(log).toEqual(['write old', 'write new']);
  });

  test('requests hydration only through explicit flush, request and recovery calls', async () => {
    const { coordinator, draft, state, ready, hydration, log } = harness();
    coordinator.setDraft(draft('live'));
    expect(hydration).toEqual([]);
    state.onHydration = ready;
    await coordinator.flush();
    expect(hydration).toEqual(['flush']);
    expect(log).toEqual(['write live']);
    await coordinator.requestHydration('on-demand');
    await coordinator.recover();
    expect(hydration).toEqual(['flush', 'on-demand', 'recovery']);
  });

  test('drains accepted work before changing generations without an explicit reset', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const commit = coordinator.submitAsync(draft('old'));
    const draining = coordinator.drain();
    ready();
    await draining;
    await commit;
    state.generation += 1;
    await coordinator.submitAsync(draft('new'));
    expect(log).toEqual(['write old', 'write new']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('retains a failed synchronous host operation with its acceptance snapshot', async () => {
    const { coordinator, state, ready, log } = harness();
    ready();
    const failed = new Error('facade failed');
    const input = { value: 'accepted' };
    let refuse = true;
    expect(() =>
      coordinator.runSync(
        (intent) => {
          if (refuse) throw failed;
          expect(intent.sheet).toBe(0);
          expect(intent.target).toBe('A1');
          expect(intent.input).toEqual({ value: 'accepted' });
          log.push('host');
        },
        { input }
      )
    ).toThrow(failed);
    input.value = 'changed';
    state.sheet = 2;
    state.target = 'B2';
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['host']);
    refuse = false;
    await coordinator.recover();
    expect(log).toEqual(['host']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('drains work accepted as the preceding operation completes', async () => {
    const { coordinator, ready, log } = harness();
    ready();
    await coordinator
      .runAfterPendingInput(() => log.push('first'))
      .then(() => coordinator.runAfterPendingInput(() => log.push('second')));
    expect(log).toEqual(['first', 'second']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('retains a draft when the facade throws and writes it during recovery', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const failed = new Error('facade failed');
    state.writeError = failed;
    const commit = coordinator.submitAsync(draft('typed'));
    ready();
    expect(await errorOf(commit)).toBe(failed);
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['commit']);
    state.writeError = null;
    await coordinator.recover();
    expect(log).toEqual(['write typed']);
  });

  test('recovers final composition text that resolves after session failure', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    coordinator.setDraft(draft('日本'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    const failed = new Error('worker failed');
    coordinator.fail(failed);
    expect(await errorOf(command)).toBe(failed);
    state.shown = '日本語';
    composition.resolve(true);
    state.onHydration = ready;
    await coordinator.recover();
    expect(log).toEqual(['write 日本語', 'command']);
  });

  test('does not replay a refused command draft after its correction is written', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    state.refuse.add('bad');
    coordinator.setDraft(draft('bad'));
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    ready();
    expect(await errorOf(command)).toBe('input-failed');
    const corrected = draft('corrected');
    coordinator.setDraft(corrected);
    await coordinator.submitAsync(corrected);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
    await coordinator.recover();
    expect(log).toEqual(['refused bad', 'write corrected', 'command']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('drains accepted flush input through hydration before replacement', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const hydration = deferred<void>();
    state.onHydration = () => hydration.promise;
    coordinator.setDraft(draft('old'));
    const flush = coordinator.flush();
    const draining = coordinator.drain();
    expect(log).toEqual([]);
    ready();
    hydration.resolve();
    await flush;
    await draining;
    state.generation += 1;
    coordinator.reset();
    expect(coordinator.unapplied).toEqual([]);
    expect(log).toEqual(['write old']);
  });

  test('captures clipboard access before sealing and waits for final composition text', async () => {
    const { coordinator, draft, state, ready, log } = harness();
    const composition = deferred<boolean>();
    state.seal = { composition: composition.promise };
    state.onSeal = () => log.push('seal');
    coordinator.setDraft(draft('日本'));
    const paste = coordinator.clipboard(
      () => {
        log.push('capture');
        return 'paste';
      },
      (input) => log.push(input)
    );
    expect(log).toEqual(['capture', 'seal']);
    state.shown = '日本語';
    composition.resolve(true);
    ready();
    await paste;
    expect(log).toEqual(['capture', 'seal', 'write 日本語', 'paste']);
  });

  test('retains refused clipboard data and blocks commands behind it until recovery', async () => {
    const { coordinator, ready, log } = harness();
    let accept = false;
    const paste = coordinator.clipboard(
      () => 'paste',
      (input) => {
        if (!accept) return false;
        log.push(input);
        return true;
      }
    );
    const command = coordinator.runAfterPendingInput(() => log.push('command'));
    ready();
    expect(await errorOf(paste)).toBe('input-failed');
    expect(await errorOf(command)).toBe('input-failed');
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['clipboard', 'command']);
    expect(coordinator.unapplied[0].input).toBe('paste');
    accept = true;
    await coordinator.recover();
    expect(log).toEqual(['paste', 'command']);
    expect(coordinator.unapplied).toEqual([]);
  });

  test('keeps newer commits behind refused clipboard input until FIFO recovery', async () => {
    const { coordinator, draft, ready, log, writes } = harness();
    ready();
    let accept = false;
    const paste = coordinator.clipboard(() => 'paste', (input) => {
      if (!accept) return false;
      log.push(`write ${input}`);
      return true;
    });
    expect(await errorOf(paste)).toBe('input-failed');
    const newer = draft('newer');
    expect(await errorOf(coordinator.submitAsync(newer))).toBe('input-failed');
    expect(writes).toEqual([]);
    expect(log).toEqual([]);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['clipboard', 'commit']);
    accept = true;
    await coordinator.recover();
    expect(log).toEqual(['write paste', 'write newer']);
    expect(writes).toEqual([newer]);
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.pending).toBe(false);
  });

  test('rejects every flush while refused clipboard input remains unapplied', async () => {
    const { coordinator, ready } = harness();
    ready();
    let accept = false;
    expect(await errorOf(coordinator.clipboard(() => 'paste', () => accept))).toBe('input-failed');
    for (let call = 0; call < 2; call += 1) {
      expect(await errorOf(coordinator.flush())).toBe('input-failed');
      expect(coordinator.pending).toBe(true);
      expect(coordinator.unapplied[0]?.kind).toBe('clipboard');
      expect(() => coordinator.runSync(() => {})).toThrow(WorkerInputNotReadyError);
    }
    accept = true;
    await coordinator.recover();
    await coordinator.flush();
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.pending).toBe(false);
  });

  test('blocks draft corrections behind a clipboard entry whose sealed draft was refused', async () => {
    const { coordinator, draft, ready, state, log, writes } = harness();
    coordinator.setDraft(draft('sealed'));
    state.refuse.add('sealed');
    const paste = coordinator.clipboard(() => 'paste', (input) => { log.push(`write ${input}`); });
    ready();
    expect(await errorOf(paste)).toBe('input-failed');
    const newer = draft('newer');
    coordinator.setDraft(newer);
    expect(await errorOf(coordinator.submitAsync(newer))).toBe('input-failed');
    expect(writes).toEqual([]);
    expect(log).toEqual(['refused sealed']);
    expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['clipboard', 'commit']);
    state.refuse.clear();
    await coordinator.recover();
    expect(log).toEqual(['refused sealed', 'write sealed', 'write paste', 'write newer']);
    expect(writes.map((entry) => entry.value)).toEqual(['sealed', 'newer']);
    expect(coordinator.unapplied).toEqual([]);
    expect(coordinator.pending).toBe(false);
  });

  test('waits for input queued after coordinator flush starts through the final worker acknowledgement', async () => {
    const { fixture, wasmBytes } = await loadWorkbookSessionFixtures();
    const acknowledgement = [deferred<void>(), deferred<void>()];
    const entered = [deferred<void>(), deferred<void>()];
    const hydrating = deferred<void>();
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          const sequence = (message.args[0] as { sequence: number }).sequence;
          entered[sequence - 1].resolve();
          void acknowledgement[sequence - 1].promise.then(() => listener(message));
        } else listener(message);
      }),
    }), { wasm: wasmBytes.buffer, retainPeerHydration: true });
    const { coordinator, draft, ready, state, writes } = harness();
    let peer: WorkbookHandle | undefined;
    let edits: WorkbookEditPeer | undefined;
    state.onHydration = async () => {
      await hydrating.promise;
      peer = await hydratePeer(session);
      edits = createWorkbookEditPeer({ session, peer });
      ready();
    };
    state.onWrite = (input) => {
      if (!edits) throw new Error('Missing hydrated edit peer');
      expect(edits.editCell(input.sheet, input.row, input.col, input.value).applied).toBe(true);
    };
    state.onFlushEdits = () => {
      if (!edits) throw new Error('Missing hydrated edit peer');
      return edits.flush();
    };
    const flush = coordinator.flush();
    let flushed = false;
    void flush.then(() => { flushed = true; }).catch(() => {});
    const first = draft('queued after flush starts');
    const commit = coordinator.submitAsync(first);
    hydrating.resolve();
    try {
      await commit;
      await entered[0].promise;
      await session.call.version();
      if (!peer || !edits) throw new Error('Missing hydrated edit peer');
      expect(flushed).toBe(false);
      expect(edits.sentSequence).toBe(1);
      expect(edits.acknowledgedSequence).toBe(0);
      const second = draft('queued during acknowledgement', 1);
      await coordinator.submitAsync(second);
      await entered[1].promise;
      acknowledgement[0].resolve();
      await session.call.version();
      expect(flushed).toBe(false);
      expect(edits.sentSequence).toBe(2);
      expect(edits.acknowledgedSequence).toBe(1);
      acknowledgement[1].resolve();
      await flush;
      expect(flushed).toBe(true);
      expect(edits.acknowledgedSequence).toBe(edits.sentSequence);
      expect(edits.acknowledgedSequence).toBe(2);
      expect(writes).toEqual([first, second]);
      expect(peer.version()).toBe(await session.call.version());
      expect(peer.save()).toEqual(await session.save());
      expect(coordinator.unapplied).toEqual([]);
      expect(coordinator.pending).toBe(false);
    } finally {
      for (const held of acknowledgement) held.resolve();
      edits?.dispose();
      peer?.dispose();
      await session.dispose();
    }
  });

  test('flushes hydration-queued input through peer application and worker acknowledgement', async () => {
    const { fixture, wasmBytes } = await loadWorkbookSessionFixtures();
    const calculation = { nowSerial: 46_000.5, randSeed: 123456789 };
    const acknowledgement = deferred<void>();
    const entered = deferred<void>();
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      listen: (listener) => transport.listen((message) => {
        if (isClientMessage(message) && message.kind === 'call' && message.method === 'replay') {
          entered.resolve();
          void acknowledgement.promise.then(() => listener(message));
        } else listener(message);
      }),
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    const { coordinator, draft, ready, state, writes } = harness();
    let peer: WorkbookHandle | undefined;
    let edits: WorkbookEditPeer | undefined;
    state.onHydration = async () => {
      peer = await hydratePeer(session);
      edits = createWorkbookEditPeer({ session, peer });
      ready();
    };
    state.onWrite = (input) => {
      if (!edits) throw new Error('Missing hydrated edit peer');
      expect(edits.editCell(input.sheet, input.row, input.col, input.value).applied).toBe(true);
    };
    state.onFlushEdits = () => {
      if (!edits) throw new Error('Missing hydrated edit peer');
      return edits.flush();
    };
    try {
      const original = await session.call.cellInputs(0, 'A1');
      const input = draft('queued during hydration');
      const commit = coordinator.submitAsync(input);
      const flush = coordinator.flush();
      let flushed = false;
      void flush.then(() => { flushed = true; }).catch(() => {});
      await Promise.race([
        entered.promise,
        flush.then(() => { throw new Error('Flush resolved before worker replay'); }),
      ]);
      await commit;
      if (!peer || !edits) throw new Error('Missing hydrated edit peer');
      expect(writes).toEqual([input]);
      expect(peer.cell(0, 0, 0).input).toBe(input.value);
      expect(await session.call.cellInputs(0, 'A1')).toEqual(original);
      expect(edits.sentSequence).toBe(1);
      expect(edits.acknowledgedSequence).toBe(0);
      expect(flushed).toBe(false);
      acknowledgement.resolve();
      await flush;
      expect(peer.cell(0, 0, 0).input).toBe(input.value);
      expect((await session.call.cellInputs(0, 'A1')).cells[0]?.[0]?.input).toBe(input.value);
      expect(edits.acknowledgedSequence).toBeGreaterThanOrEqual(1);
      expect((await session.call.frame({ x: 0, y: 0, width: 800, height: 800 })).sequence).toBeGreaterThanOrEqual(1);
      expect(peer.version()).toBe(await session.call.version());
      expect(peer.save()).toEqual(await session.save());
      expect(coordinator.unapplied).toEqual([]);
      expect(coordinator.pending).toBe(false);
    } finally {
      acknowledgement.resolve();
      edits?.dispose();
      peer?.dispose();
      await session.dispose();
    }
  });

  test('keeps input arriving inside recovery hydration in the peer and recovery save', async () => {
    const { fixture, wasmBytes } = await loadWorkbookSessionFixtures();
    const calculation = { nowSerial: 46_000.5, randSeed: 123456789 };
    let crash: ((error: unknown) => void) | undefined;
    const session = await createTestWorkbookSession(fixture, (transport) => ({
      ...transport,
      onError(listener) { crash = listener; return transport.onError(listener); },
    }), { calculation, wasm: wasmBytes.buffer, retainPeerHydration: true });
    const peer = await hydratePeer(session);
    const edits = createWorkbookEditPeer({ session, peer });
    const { coordinator, draft, ready, state, log, writes } = harness();
    ready();
    try {
      if (!crash) throw new Error('Missing worker crash callback');
      const failed = new Promise<SessionFailure>((resolve) => { session.onFailure(resolve); });
      crash(new SessionFailure('crash', 'Worker stopped'));
      expect(await failed).toMatchObject({ code: 'crash', message: 'Worker stopped' });
      expect(edits.state).toBe('failed');
      const operations = workbookEditPeerOperations(edits);
      const paste: WorkbookReplayOp = { method: 'editCell', args: [0, 0, 0, 'paste'] };
      state.onWrite = (value) => {
        expect(operations.applyRecoveryOp({
          method: 'editCell', args: [value.sheet, value.row, value.col, value.value],
        })).toMatchObject({ applied: true });
      };
      let accept = false;
      const original = peer.cell(0, 0, 0).input;
      expect(await errorOf(coordinator.clipboard(() => 'paste', (input) => {
        if (!accept) return false;
        expect(operations.applyRecoveryOp(paste)).toMatchObject({ applied: true });
        log.push(`write ${input}`);
        return true;
      }))).toBe('input-failed');
      const entered = deferred<void>();
      const hydration = deferred<void>();
      const incoming = draft('arrived during recovery');
      state.onHydration = async (reason) => {
        if (reason !== 'recovery') return;
        expect(await errorOf(coordinator.submitAsync(incoming))).toBe('input-failed');
        entered.resolve();
        await hydration.promise;
      };
      accept = true;
      const recovering = coordinator.recover();
      await Promise.race([
        entered.promise,
        recovering.then(() => { throw new Error('Recovery finished before retaining incoming input'); }),
      ]);
      expect(peer.cell(0, 0, 0).input).toBe(original);
      expect(writes).toEqual([]);
      expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['clipboard', 'commit']);
      hydration.resolve();
      await recovering;
      expect(peer.cell(0, 0, 0).input).toBe('arrived during recovery');
      const recoverySave = edits.recoverySave();
      expect<Uint8Array>(new Uint8Array(recoverySave.bytes)).toEqual(peer.save());
      const reopened = openWorkbook(new Uint8Array(recoverySave.bytes), { calculation });
      try {
        expect(reopened.cell(0, 0, 0).input).toBe('arrived during recovery');
      } finally { reopened.dispose(); }
      expect(log).toEqual(['write paste', 'write arrived during recovery']);
      expect(coordinator.unapplied).toEqual([]);
      expect(coordinator.pending).toBe(false);
    } finally {
      edits.dispose();
      peer.dispose();
      await session.dispose();
    }
  });

  test('records synchronous operation application before a failure in the next microtask', async () => {
    for (const kind of ['input', 'clipboard', 'command'] as const) {
      const { coordinator, ready, log } = harness();
      const failed = new Error(`${kind} acknowledgement failed`);
      const apply = () => {
        log.push(kind);
        queueMicrotask(() => coordinator.fail(failed));
        return true;
      };
      let result: Promise<unknown>;
      if (kind === 'input') result = coordinator.input(apply);
      else if (kind === 'clipboard') result = coordinator.clipboard(() => 'payload', apply);
      else result = coordinator.runAfterPendingInput(apply);
      const after = coordinator.runAfterPendingInput(() => log.push('after'));
      ready();
      expect(await errorOf(result)).toBe(failed);
      expect(await errorOf(after)).toBe(failed);
      expect(coordinator.unapplied.map((entry) => entry.kind)).toEqual(['command']);
      await coordinator.recover();
      expect(log).toEqual([kind, 'after']);
    }
  });
});

test('allows current-frame merge geometry only while pending commands preserve merges', async () => {
  const { coordinator, draft, ack } = peerlessHarness();
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const navigation = coordinator.command({ kind: 'select', payload: { selection: selected } });
  expect(coordinator.pending).toBe(true);
  expect(coordinator.pendingSelectionGeometryStable).toBe(true);
  const edit = coordinator.submit(draft('changed'));
  expect(coordinator.pendingSelectionGeometryStable).toBe(true);
  const merge = coordinator.command({ kind: 'applyOps', payload: { toolbar: { id: 'merge', args: { value: 'all' }, target: selected } } });
  expect(coordinator.pendingSelectionGeometryStable).toBe(false);
  ack(navigation.seq!);
  ack(edit.seq!);
  ack(merge.seq!);
  await Promise.all([navigation.result, edit.result, merge.result]);
  expect(coordinator.pending).toBe(false);
});
