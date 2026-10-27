import { expect, test } from 'bun:test';
import { SessionFailure } from '../../../../shared/office-session/types';
import type { EchoEntry, EchoHooks, RecoveryEvent } from '../../../../shared/office-session/commands/types';
import { isTransferResult } from '../../../../shared/office-session/protocol';
import type { WorkbookHandle } from '../wasm/loader';
import { createWorkbookCommandState } from './commandHandlers';
import { createWorkbookCommandHost } from './commandHost';
import type { WorkbookCommand, WorkbookCommandFrame, WorkbookEditConnection, WorkbookSharedAck, WorkbookSharedCommand } from './commandProtocol';
import { createWorkbookEditSession } from './editSession';

const initial: WorkbookSharedAck = {
  session: 'book', incarnation: 1, epoch: 1, throughSeq: 0, version: 0,
  engineVersion: 'opaque:initial', outcomes: [],
};

function controlled() {
  const calls: { command: WorkbookSharedCommand; resolve(ack: WorkbookSharedAck): void }[] = [];
  const failures = new Set<(failure: SessionFailure) => void>();
  let failure: SessionFailure | undefined;
  let checkpoints = 0;
  let closes = 0;
  const connection: WorkbookEditConnection = {
    initial,
    client: {
      call: {
        edit: (command) => new Promise((resolve) => { calls.push({ command, resolve }); }),
        checkpoint: async () => { checkpoints += 1; throw new Error('Unexpected checkpoint'); },
      },
      get failure() { return failure; },
      onFailure(listener) { failures.add(listener); return () => { failures.delete(listener); }; },
      dispose: async () => { closes += 1; },
    },
  };
  return {
    connection, calls,
    get checkpoints() { return checkpoints; },
    get closes() { return closes; },
    fail(value: SessionFailure) {
      failure = value;
      for (const listener of [...failures]) listener(value);
    },
    ack(index: number, withFrame = false, revision = index + 1) {
      const seq = calls[index].command.seq;
      calls[index].resolve({ ...initial, throughSeq: seq, version: seq, engineVersion: `opaque:${seq}`,
        outcomes: [{ seq, status: 'applied' }], ...(withFrame ? { frame: {
          coversThroughSeq: seq, revision, payload: paint(seq, revision),
        } } : {}) });
    },
  };
}

function paint(seq: number, revision: number): WorkbookCommandFrame {
  return {
    session: 'book', incarnation: 'opaque:worker', coversThroughSeq: seq,
    epoch: revision, sequence: seq, version: `opaque:${seq}`, sheet: 0,
    viewport: { x: 0, y: 0, width: 80, height: 40 }, zoom: 1, dpr: 1,
    requestGeneration: 1, displayList: { width: 80, height: 40, commands: [] },
  };
}

async function drain() {
  for (let step = 0; step < 8; step += 1) await Promise.resolve();
}

for (const code of ['crash', 'trap', 'silence', 'out-of-memory'] as const) {
  test(`xlsx shared session terminates ${code} without recovery or settled replay`, async () => {
    const port = controlled();
    const reports: RecoveryEvent[] = [];
    let sourceReads = 0;
    let reopens = 0;
    let checkpointPolicies = 0;
    let recoveryCheckpoints = 0;
    const WeakReference = (globalThis as unknown as {
      WeakRef: new (value: object) => { deref(): object | undefined };
    }).WeakRef;
    const retained: { deref(): object | undefined }[] = [];
    const source = Object.defineProperty({}, 'bytes', { enumerable: true, get() {
      sourceReads += 1;
      throw new Error('Source must never be read');
    } });
    const session = createWorkbookEditSession({
      connection: port.connection,
      get source() { sourceReads += 1; return source; },
      recovery: {
        reopen: async () => { reopens += 1; throw new Error('Unexpected reopen'); },
        checkpoint: async () => { recoveryCheckpoints += 1; throw new Error('Unexpected checkpoint'); },
        onRecovery: (event) => { reports.push(event); },
      },
      checkpointPolicy: () => { checkpointPolicies += 1; return true; },
      commandBytes: (envelope) => { retained.push(new WeakReference(envelope)); return 1; },
      echo: { paint: () => {}, drop: () => {}, present: () => true },
      requestFullFrame: async () => { throw new Error('Unexpected frame request'); },
    });
    const settled = session.submit('editCell', { payload: { args: [0, 0, 0, 'settled'] } });
    await drain();
    port.ack(0, true);
    expect(await settled.ack).toMatchObject({ seq: 1, status: 'applied' });
    expect(await settled.presented).toBe('presented');
    await new Promise<void>((resolve) => { setTimeout(resolve, 0); });
    Bun.gc(true);
    expect(retained).toHaveLength(1);
    expect(retained[0].deref()).toBeUndefined();
    const pending = [session.submit('editCell', { payload: { args: [0, 1, 0, 'pending'] } }),
      session.submit('editCell', { payload: { args: [0, 2, 0, 'dependent'] } }, { after: settled })];
    const rejected: unknown[] = [];
    for (const ticket of pending) void ticket.ack.catch((error) => { rejected.push(error); });
    await drain();
    const failure = new SessionFailure(code, `Synthetic ${code}`);
    port.fail(failure);
    port.fail(failure);
    for (const ticket of pending) {
      await expect(ticket.ack).rejects.toBe(failure);
      expect(await ticket.presented).toBe('dropped');
    }
    expect(await settled.ack).toMatchObject({ seq: 1, status: 'applied' });
    port.ack(1, true);
    await drain();
    expect(rejected).toEqual([failure, failure]);
    expect(reports).toEqual([{ phase: 'failed', incarnation: 1, attempt: 0,
      diverged: [], historyReset: false, failure }]);
    expect(sourceReads).toBe(0);
    expect(reopens).toBe(0);
    expect(recoveryCheckpoints).toBe(0);
    expect(checkpointPolicies).toBe(0);
    expect(port.checkpoints).toBe(0);
    expect(port.calls).toHaveLength(3);
    expect(port.closes).toBe(1);
    expect(session.state.editState?.readOnly).toBe(true);
    await session.dispose();
  });
}

for (const rejectedInstall of ['reject', 'throw'] as const) {
  test(`xlsx shared presentation preserves coverage revision echoes and waiters on ${rejectedInstall}`, async () => {
    const port = controlled();
    let accept = false;
    let visible: readonly EchoEntry<string>[] = [];
    const attempts: { retired: string[]; remaining: string[] }[] = [];
    const echo: EchoHooks<string, ReturnType<typeof paint>> = {
      paint: (entries) => { visible = entries; }, drop: () => {},
      present(_frame, retired, remaining) {
        expect(session.state.presentedThroughSeq).toBe(-1);
        expect(session.presentationRevision).toBeUndefined();
        attempts.push({ retired: retired.map((entry) => entry.value), remaining: remaining.map((entry) => entry.value) });
        if (!accept) {
          if (rejectedInstall === 'throw') throw new Error('Synthetic paint failure');
          return false;
        }
        visible = remaining;
        return true;
      },
    };
    const session = createWorkbookEditSession({ connection: port.connection, echo,
      requestFullFrame: async () => initial });
    const ticket = session.submit('editCell', { payload: { args: [0, 0, 0, 'new'] } }, { echo: 'new' });
    await drain();
    let ticketPresented = false;
    let waiterPresented = false;
    void ticket.presented.then(() => { ticketPresented = true; });
    const waiter = session.whenPresented(1).then((status) => { waiterPresented = true; return status; });
    port.ack(0, true, 5);
    await ticket.ack;
    await drain();
    expect(session.state.presentedThroughSeq).toBe(-1);
    expect(session.presentationRevision).toBeUndefined();
    expect(visible.map((entry) => entry.value)).toEqual(['new']);
    expect(ticketPresented).toBe(false);
    expect(waiterPresented).toBe(false);
    accept = true;
    session.present({ coversThroughSeq: 1, revision: 5, payload: paint(1, 5) });
    expect(await ticket.presented).toBe('presented');
    expect(await waiter).toBe('presented');
    expect(session.state.presentedThroughSeq).toBe(1);
    expect(session.presentationRevision).toBe(5);
    expect(visible).toEqual([]);
    expect(attempts).toEqual([{ retired: ['new'], remaining: [] }, { retired: ['new'], remaining: [] }]);
    await session.dispose();
  });
}

test('xlsx shared echo overflow keeps newest text without dropping displaced tickets', async () => {
  const port = controlled();
  let visible: string[] = [];
  const drops: string[] = [];
  const session = createWorkbookEditSession<string>({ connection: port.connection, maxEchoEntries: 2,
    echo: {
      paint: (entries) => { visible = entries.map((entry) => entry.value); },
      drop: (entries) => { drops.push(...entries.map((entry) => entry.value)); },
      present: (_frame, _retired, remaining) => { visible = remaining.map((entry) => entry.value); return true; },
    }, requestFullFrame: async () => initial });
  const tickets = ['oldest', 'middle', 'newest'].map((text, row) =>
    session.submit('editCell', { payload: { args: [0, row, 0, text] } }, { echo: text }));
  await drain();
  expect(visible).toEqual(['middle', 'newest']);
  expect(drops).toEqual([]);
  let displacedPresented = false;
  void tickets[0].presented.then(() => { displacedPresented = true; });
  port.ack(0);
  port.ack(1);
  port.ack(2, true);
  for (const ticket of tickets) {
    expect(await ticket.ack).toMatchObject({ status: 'applied' });
    expect(await ticket.presented).toBe('presented');
  }
  expect(displacedPresented).toBe(true);
  expect(visible).toEqual([]);
  expect(drops).toEqual([]);
  await session.dispose();
});

function compatibility(seq: number): WorkbookCommand {
  return { session: 'book', incarnation: 'opaque:worker', seq,
    basis: { epoch: 1, version: initial.engineVersion }, kind: 'editCell',
    payload: { args: [0, 0, 0, String(seq)] } };
}

test('simultaneous replies publish after individual shared acceptance with manual presentation', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection,
    identity: { session: 'book', incarnation: 'opaque:worker' },
    requestFullFrame: async () => initial });
  const events: string[] = [];
  session.onAck((ack) => {
    events.push(`ack:${ack.throughSeq}:coverage:${session.state.throughSeq}`);
    events.push(`paint:${ack.throughSeq}:${session.markPresented(ack.frame!)}`);
  });
  const first = session.command(compatibility(1));
  const second = session.command(compatibility(2));
  const presented = session.whenPresented(2);
  await drain();
  expect(port.calls.map((call) => call.command.seq)).toEqual([1, 2]);
  port.ack(0, true);
  port.ack(1, true);
  expect((await Promise.all([first, second])).map((ack) => ack.throughSeq)).toEqual([1, 2]);
  expect(events).toEqual(['ack:1:coverage:1', 'paint:1:true', 'ack:2:coverage:2', 'paint:2:true']);
  expect(await presented).toBe('presented');
  expect(session.state.presentedThroughSeq).toBe(2);
  await session.dispose();
});

for (const sent of [false, true]) {
  test(`compatibility disposal rejects ${sent ? 'sent' : 'queued'} commands without a wire reply`, async () => {
    const port = controlled();
    const session = createWorkbookEditSession({ connection: port.connection,
      identity: { session: 'book', incarnation: 'opaque:worker' },
      requestFullFrame: async () => initial });
    const events: string[] = [];
    const command = session.command(compatibility(1)).then(() => { events.push('resolved'); }, (error: unknown) => {
      expect(error).toBeInstanceOf(SessionFailure);
      expect((error as SessionFailure).code).toBe('disposed');
      events.push('rejected:disposed');
    });
    if (sent) await drain();
    expect(port.calls).toHaveLength(sent ? 1 : 0);
    events.push('dispose');
    await session.dispose();
    await command;
    expect(events).toEqual(['dispose', 'rejected:disposed']);
    if (sent) port.ack(0, true);
    await drain();
    expect(events).toEqual(['dispose', 'rejected:disposed']);
    expect(port.closes).toBe(1);
  });
}

for (const terminal of ['dispose', 'failure'] as const) {
  test(`duplicate compatibility commands reject on ${terminal} before late raw replies`, async () => {
    const port = controlled();
    const session = createWorkbookEditSession({ connection: port.connection,
      identity: { session: 'book', incarnation: 'opaque:worker' },
      requestFullFrame: async () => initial });
    const events: string[] = [];
    session.onAck((ack) => { events.push(`ack:${ack.throughSeq}`); });
    const first = session.command(compatibility(1));
    await drain();
    port.ack(0);
    await first;
    const failure = new SessionFailure('crash', 'Synthetic duplicate failure');
    const duplicate = session.command(compatibility(1)).then(() => { events.push('resolved'); }, (error: unknown) => {
      if (terminal === 'failure') expect(error).toBe(failure);
      else expect(error).toMatchObject({ code: 'disposed' });
      events.push(`rejected:${terminal}`);
    });
    expect(port.calls.map((call) => call.command.seq)).toEqual([1, 1]);
    events.push('duplicate:held');
    events.push(terminal);
    if (terminal === 'dispose') await session.dispose();
    else port.fail(failure);
    await duplicate;
    expect(events).toEqual(['ack:1', 'duplicate:held', terminal, `rejected:${terminal}`]);
    events.push('duplicate:released');
    port.ack(1);
    await drain();
    expect(events).toEqual(['ack:1', 'duplicate:held', terminal, `rejected:${terminal}`, 'duplicate:released']);
    expect(session.state.throughSeq).toBe(1);
    expect(session.engineVersion).toBe('opaque:1');
    await session.dispose();
    expect(port.closes).toBe(1);
  });
}

test('duplicate compatibility commands recheck disposal after raw reply resolution', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection,
    identity: { session: 'book', incarnation: 'opaque:worker' },
    requestFullFrame: async () => initial });
  const events: string[] = [];
  session.onAck((ack) => { events.push(`ack:${ack.throughSeq}`); });
  const first = session.command(compatibility(1));
  await drain();
  port.ack(0);
  await first;
  const duplicate = session.command(compatibility(1));
  const settled = duplicate.catch(() => {});
  expect(port.calls.map((call) => call.command.seq)).toEqual([1, 1]);
  port.ack(1);
  events.push('duplicate:released');
  await session.dispose();
  await settled;
  await expect(duplicate).rejects.toMatchObject({ code: 'disposed' });
  expect(events).toEqual(['ack:1', 'duplicate:released']);
  expect(port.closes).toBe(1);
});

test('null compatibility payload is admitted as a refusal before valid sequence two', async () => {
  const port = controlled();
  const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
  const handle = {
    version: () => initial.engineVersion, sheetCount: () => 1,
    sheetInfo: () => ({ sheetIds: ['sheet:0'], sheetNames: ['Data'], activeSheet: 0 }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }] }),
    historyState: () => ({ canUndo: false, canRedo: false }),
    selectionFormatting: () => ({}), mergedRanges: () => [], calculationStatus: () => ({ limitedCells: [] }),
  } as unknown as WorkbookHandle;
  const state = createWorkbookCommandState({ session: 'book', incarnation: 'opaque:worker' },
    { epoch: 1, version: handle.version() }, selected, ['sheet:0']);
  const host = createWorkbookCommandHost({ handle, state,
    frame: () => { throw new Error('Unexpected frame'); } });
  const events: string[] = [];
  port.connection.client.call.edit = async (input) => {
    events.push(`send:${input.seq}`);
    const reply = await host.handler(null, input);
    const ack = isTransferResult(reply) ? reply.value : reply;
    const { frame, ...decoded } = ack;
    expect(frame).toBeUndefined();
    const { frame: compatibilityFrame, ...receipt } = decoded.compatibility!;
    expect(compatibilityFrame).toBeUndefined();
    return { ...decoded, compatibility: receipt };
  };
  const session = createWorkbookEditSession({ connection: port.connection,
    identity: state.identity, requestFullFrame: async () => initial });
  session.onAck((ack) => { events.push(`${ack.statuses[0].status}:${ack.throughSeq}`); });
  const malformed = Object.assign(compatibility(1), { payload: null });
  const first = await session.command(malformed);
  expect(first.statuses).toMatchObject([{ seq: 1, status: 'refused', reason: 'invalid-command', committed: false }]);
  const second = await session.command({ ...state.identity, seq: 2,
    basis: { epoch: 1, version: initial.engineVersion }, kind: 'setReadOnly', payload: { readOnly: true } });
  expect(second.statuses).toMatchObject([{ seq: 2, status: 'applied', committed: false }]);
  expect(events).toEqual(['send:1', 'refused:1', 'send:2', 'applied:2']);
  expect(session.state.throughSeq).toBe(2);
  expect(session.state.editState?.readOnly).toBe(true);
  expect(session.failure).toBeUndefined();
  await session.dispose();
});

test('default scheduling interleaves native and compatibility sequence allocation', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection,
    identity: { session: 'book', incarnation: 'opaque:worker' },
    requestFullFrame: async () => initial });
  const events: string[] = [];
  session.onAck((ack) => { events.push(`ack:${ack.throughSeq}`); });
  const first = session.submit('editCell', { payload: { args: [0, 0, 0, '1'] } });
  expect(first.seq).toBeUndefined();
  await drain();
  expect(first.seq).toBe(1);
  port.ack(0);
  expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
  const second = session.command(compatibility(2));
  await drain();
  expect(port.calls.map((call) => call.command.seq)).toEqual([1, 2]);
  port.ack(1);
  expect((await second).statuses).toMatchObject([{ seq: 2, status: 'applied' }]);
  const third = session.submit('editCell', { payload: { args: [0, 0, 0, '3'] } });
  await drain();
  expect(third.seq).toBe(3);
  port.ack(2);
  expect(await third.ack).toMatchObject({ seq: 3, status: 'applied' });
  await drain();
  expect(events).toEqual(['ack:1', 'ack:2', 'ack:3']);
  await session.dispose();
});

test('native echo reentry refuses compatibility sequence 65 at capacity before save', async () => {
  const port = controlled();
  const events: string[] = [];
  let reentered = false;
  let rejection: unknown;
  const session = createWorkbookEditSession<string>({ connection: port.connection,
    identity: { session: 'book', incarnation: 'opaque:worker' },
    echo: {
      paint(entries) {
        if (reentered || !entries.some((entry) => entry.value === 'native:65')) return;
        reentered = true;
        events.push('echo:65');
        void session.command(compatibility(65)).then(() => { events.push('accepted:65'); }, (error: unknown) => {
          rejection = error;
          events.push('refused:65');
        });
      },
      drop: () => {}, present: () => true,
    },
    requestFullFrame: async () => initial });
  session.onAck((ack) => { events.push(`ack:${ack.throughSeq}`); });
  const first = Array.from({ length: 64 }, (_, row) =>
    session.submit('editCell', { payload: { args: [0, row, 0, `native:${row + 1}`] } }));
  await drain();
  expect(port.calls.map((call) => call.command.seq)).toEqual(Array.from({ length: 64 }, (_, index) => index + 1));
  const native = session.submit('editCell', { payload: { args: [0, 64, 0, 'native:65'] } }, { echo: 'native:65' });
  const saved = session.runAfterEdits(async () => { events.push('save:65'); return Uint8Array.of(65); });
  await drain();
  expect(reentered).toBe(true);
  expect(rejection).toMatchObject({ message: 'Workbook command sequence is awaiting native dispatch' });
  expect(events).toEqual(['echo:65', 'refused:65']);
  expect(native.seq).toBeUndefined();
  expect(port.calls).toHaveLength(64);
  port.ack(0);
  await drain();
  expect(native.seq).toBe(65);
  expect(port.calls.map((call) => call.command.seq)).toEqual(Array.from({ length: 65 }, (_, index) => index + 1));
  port.ack(1);
  await drain();
  expect(events).toEqual(['echo:65', 'refused:65', 'ack:1', 'ack:2']);
  expect(port.calls).toHaveLength(65);
  for (let index = 2; index < 65; index += 1) port.ack(index);
  for (const ticket of first) expect(await ticket.ack).toMatchObject({ seq: ticket.seq, status: 'applied' });
  expect(await native.ack).toMatchObject({ seq: 65, status: 'applied' });
  expect(await saved).toEqual(Uint8Array.of(65));
  expect(events).toEqual(['echo:65', 'refused:65',
    ...Array.from({ length: 65 }, (_, index) => `ack:${index + 1}`), 'save:65']);
  expect(session.state.throughSeq).toBe(65);
  expect(session.failure).toBeUndefined();
  await session.dispose();
});

test('compatibility admission fences later native edits against lane overtaking at capacity', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection,
    identity: { session: 'book', incarnation: 'opaque:worker' },
    maxInFlight: 1, inputBurstCap: 1,
    requestFullFrame: async () => initial });
  const events: string[] = [];
  session.onAck((ack) => { events.push(`ack:${ack.throughSeq}`); });
  const first = session.submit('editCell', { payload: { args: [0, 0, 0, 'first'] } });
  await drain();
  expect(first.seq).toBe(1);
  const compatibilityEdit = session.command(compatibility(2));
  const native = session.submit('editCell', { payload: { args: [0, 1, 0, 'native'] } }, { lane: 'collab' });
  const saved = session.runAfterEdits(async () => { events.push('save:3'); return Uint8Array.of(3); });
  await drain();
  expect(port.calls.map((call) => call.command.seq)).toEqual([1]);
  expect(native.seq).toBeUndefined();
  port.ack(0);
  expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
  await drain();
  expect(port.calls.map((call) => ({ seq: call.command.seq, compatibility: call.command.payload.compatibility?.seq })))
    .toEqual([{ seq: 1, compatibility: undefined }, { seq: 2, compatibility: 2 }]);
  expect(native.seq).toBeUndefined();
  expect(events).toEqual(['ack:1']);
  port.ack(1);
  expect((await compatibilityEdit).statuses).toMatchObject([{ seq: 2, status: 'applied' }]);
  await drain();
  expect(port.calls.map((call) => ({ seq: call.command.seq, compatibility: call.command.payload.compatibility?.seq })))
    .toEqual([{ seq: 1, compatibility: undefined }, { seq: 2, compatibility: 2 }, { seq: 3, compatibility: undefined }]);
  expect(native.seq).toBe(3);
  expect(events).toEqual(['ack:1', 'ack:2']);
  port.ack(2);
  expect(await native.ack).toMatchObject({ seq: 3, status: 'applied' });
  expect(await saved).toEqual(Uint8Array.of(3));
  expect(events).toEqual(['ack:1', 'ack:2', 'ack:3', 'save:3']);
  expect(session.failure).toBeUndefined();
  await session.dispose();
});

test('prototype client methods preserve their receiver through failure and disposal', async () => {
  const port = controlled();
  const events: string[] = [];
  class Client {
    readonly call = port.connection.client.call;
    readonly failure = undefined;
    onFailure(listener: (failure: SessionFailure) => void) {
      expect<Client>(this).toBe(client);
      events.push('subscribe');
      const off = port.connection.client.onFailure(listener);
      return () => { events.push('unsubscribe'); off(); };
    }
    async dispose() {
      expect<Client>(this).toBe(client);
      events.push('dispose');
    }
  }
  const client = new Client();
  expect(Object.keys(client)).not.toContain('onFailure');
  expect(Object.keys(client)).not.toContain('dispose');
  const session = createWorkbookEditSession({ connection: { initial, client },
    requestFullFrame: async () => initial });
  expect(events).toEqual(['subscribe']);
  const ticket = session.submit('editCell', { payload: { args: [0, 0, 0, '1'] } });
  const rejected = ticket.ack.catch((error: unknown) => error);
  await drain();
  const failure = new SessionFailure('crash', 'Synthetic failure');
  port.fail(failure);
  expect(await rejected).toBe(failure);
  expect(session.failure).toBe(failure);
  await session.dispose();
  expect(events).toEqual(['subscribe', 'unsubscribe', 'dispose']);
});

test('sparse projection patches merge before shared state publication', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection, requestFullFrame: async () => initial });
  const mergedRanges = Array.from({ length: 5000 }, (_, row) => ({ start: { row, col: 0 }, end: { row, col: 1 } }));
  const published: { seq: number; ranges: number; readOnly: boolean | undefined }[] = [];
  session.subscribe((state) => { published.push({ seq: state.throughSeq,
    ranges: state.editState?.mergedRanges?.length ?? 0, readOnly: state.editState?.readOnly }); });
  const first = session.submit('select', { payload: { selection: {
    sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' },
  } } });
  await drain();
  port.calls[0].resolve({ ...initial, throughSeq: 1, outcomes: [{ seq: 1, status: 'noop' }],
    editState: { mergedRanges, readOnly: false, canUndo: false } });
  await first.ack;
  const second = session.submit('setReadOnly', { payload: { readOnly: true } });
  await drain();
  port.calls[1].resolve({ ...initial, throughSeq: 2, outcomes: [{ seq: 2, status: 'noop' }],
    editState: { readOnly: true } });
  await second.ack;
  expect(published).toEqual([{ seq: 1, ranges: 5000, readOnly: false }, { seq: 2, ranges: 5000, readOnly: true }]);
  expect(session.state.editState?.mergedRanges).toEqual(mergedRanges);
  expect(session.state.editState?.canUndo).toBe(false);
  await session.dispose();
});

test('non-committing shared outcomes preserve compatibility receipts', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection,
    requestFullFrame: async () => { throw new Error('Unexpected frame request'); } });
  const acknowledgements: import('./commandProtocol').WorkbookAck[] = [];
  session.onAck((ack) => { acknowledgements.push(ack); });
  const ticket = session.submit('setReadOnly', { payload: { readOnly: true } });
  await drain();
  port.calls[0].resolve({ ...initial, throughSeq: 1,
    outcomes: [{ seq: 1, status: 'applied', committed: false }], editState: { readOnly: true } });
  await ticket.ack;
  await drain();
  expect(acknowledgements[0].statuses).toEqual([{ seq: 1, status: 'applied', committed: false }]);
  expect(session.state.version).toBe(0);
  expect(session.engineVersion).toBe(initial.engineVersion);
  session.dispose();
});

test('an edit waiting behind a read is not overtaken by an edit submitted as the read settles', async () => {
  const port = controlled();
  const session = createWorkbookEditSession({ connection: port.connection,
    identity: { session: 'book', incarnation: 'opaque:worker' }, requestFullFrame: async () => initial });
  const value = (index: number) => (port.calls[index].command.payload.payload as { args: unknown[] }).args[3];
  const first = session.submit('editCell', { payload: { args: [0, 0, 0, 'first'] } });
  await drain();
  expect(port.calls.map((call) => call.command.seq)).toEqual([1]);
  const read = session.runAfterEdits(async () => 'cells');
  const waiting = session.submit('editCell', { payload: { args: [0, 1, 0, 'waiting'] } });
  await drain();
  expect(port.calls).toHaveLength(1);
  const later = read.then(() => session.submit('editCell', { payload: { args: [0, 2, 0, 'later'] } }));
  port.ack(0);
  expect(await first.ack).toMatchObject({ seq: 1, status: 'applied' });
  const laterTicket = await later;
  await drain();
  expect(port.calls.map((_, index) => value(index))).toEqual(['first', 'waiting', 'later']);
  expect(port.calls.map((call) => call.command.seq)).toEqual([1, 2, 3]);
  port.ack(1);
  port.ack(2);
  expect(await waiting.ack).toMatchObject({ seq: 2, status: 'applied' });
  expect(await laterTicket.ack).toMatchObject({ seq: 3, status: 'applied' });
  await session.dispose();
});
