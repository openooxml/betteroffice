import { beforeAll, expect, test } from 'bun:test';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { CollaborationProvider } from '../collaboration/provider';
import { decodeMessages } from '../collaboration/protocol';
import type { CollaborationTransport, CollaborationTransportEvent, CollaborationUpdateOrigin } from '../collaboration/types';
import type { WorkbookSession } from './client';
import type { WorkbookAck } from './commandProtocol';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from './testHelpers';

let bytes: Uint8Array;
beforeAll(async () => { bytes = (await loadWorkbookSessionFixtures()).fixture; });

function transports() {
  const listeners = [new Set<(event: CollaborationTransportEvent) => void>(),
    new Set<(event: CollaborationTransportEvent) => void>()];
  const frames: Uint8Array[][] = [[], []];
  const pair: CollaborationTransport[] = listeners.map((local, index) => ({
    connect() { for (const listener of local) listener({ type: 'open' }); },
    disconnect() { for (const listener of local) listener({ type: 'close' }); },
    send(data) {
      frames[index].push(data.slice());
      for (const listener of listeners[1 - index]) listener({ type: 'message', data: data.slice() });
      return true;
    },
    onEvent(listener) { local.add(listener); return () => { local.delete(listener); }; },
  }));
  return { pair, frames };
}

function whenSynced(provider: CollaborationProvider): Promise<void> {
  if (provider.synced) return Promise.resolve();
  return new Promise((resolve) => {
    const off = provider.onStatus(({ synced }) => { if (synced) { off(); resolve(); } });
  });
}

function whenAcknowledged(session: WorkbookSession, throughSeq = session.throughSeq + 1): Promise<void> {
  if (session.throughSeq >= throughSeq) return Promise.resolve();
  return new Promise((resolve) => {
    const off = session.onAck((ack) => { if (ack.throughSeq >= throughSeq) { off(); resolve(); } });
  });
}

async function quiet(progress: () => number) {
  for (let idle = 0, seen = -1; idle < 64;) {
    await new Promise((resolve) => setTimeout(resolve, 0));
    const moved = progress();
    idle = moved === seen ? idle + 1 : 0;
    seen = moved;
  }
}

async function edit(session: WorkbookSession, input: string) {
  const ticket = session.editSession!.submit('editCell', {
    payload: { args: [0, 2, 1, input] },
    frame: { sheet: 0, viewport: { x: 0, y: 0, width: 800, height: 600 }, zoom: 1, dpr: 1, generation: 1 },
  });
  expect((await ticket.ack).status).toBe('applied');
}

test('collaborative worker replicas paint initial state and serialize edits, merges and undo', async () => {
  const left = await createTestWorkbookSession(bytes, undefined, { collaborative: true, clientId: 101 });
  let right: WorkbookSession | undefined;
  const providers: CollaborationProvider[] = [];
  let hold = false;
  const delayed: (() => void)[] = [];
  let held: { count: number; resolve(): void } | undefined;
  const whenHeld = (count: number) => new Promise<void>((resolve) => {
    if (delayed.length >= count) resolve();
    else held = { count, resolve };
  });
  try {
    expect(left.replica).toBeDefined();
    await edit(left, 'seed');
    const initialUpdate = await left.replica!.async.encodeStateAsUpdate();
    right = await createTestWorkbookSession(bytes, (transport) => ({ ...transport,
      post(message, transfer) {
        if (hold && isHostMessage(message) && message.kind === 'reply' && message.ok &&
          message.value && typeof message.value === 'object' && 'outcomes' in message.value) {
          delayed.push(() => transport.post(message, transfer));
          if (held && delayed.length >= held.count) { held.resolve(); held = undefined; }
        } else transport.post(message, transfer);
      },
    }), {
      collaborative: true, clientId: 202, initialUpdate,
      initialViewport: { width: 800, height: 600, zoom: 1, dpr: 1, requestGeneration: 0 },
    });
    expect(right.initialPaint?.frame.displayList.commands.some((command) =>
      command.op === 'text' && command.text === 'seed')).toBe(true);
    expect(right.initialPaint?.frame.version).toBe(right.confirmedBasis.version);
    const wire = transports();
    providers.push(new CollaborationProvider(left.replica!, wire.pair[0]),
      new CollaborationProvider(right.replica!, wire.pair[1]));
    const acks: number[] = [];
    const remotePaints: WorkbookAck[] = [];
    const updates: CollaborationUpdateOrigin[] = [];
    right.replica!.onUpdate((_, origin) => updates.push(origin));
    right.onAck((ack) => {
      acks.push(ack.throughSeq);
      if (ack.frame) remotePaints.push(ack);
    });
    const synced = Promise.all(providers.map(whenSynced));
    for (const provider of providers) provider.connect();
    await synced;
    await Promise.all([left.call.version(), right.call.version()]);
    await quiet(() => wire.frames[0].length + wire.frames[1].length + left.throughSeq + right!.throughSeq);
    expect(providers.every((provider) => provider.synced)).toBe(true);
    const remoteUpdates = () => wire.frames[1].flatMap((frame) => decodeMessages(frame))
      .filter((message) => message.type === 'update').length;
    const before = remoteUpdates();
    const received = updates.length;
    const paints = remotePaints.length;
    let merged = whenAcknowledged(right);
    await edit(left, 'local');
    await merged;
    expect((await right.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('local');
    expect(updates.slice(received)).toEqual(['remote']);
    expect(remoteUpdates()).toBe(before);
    expect(remotePaints.length).toBeGreaterThan(paints);
    expect(remotePaints.slice(paints).some((ack) => ack.frame && ack.frame.version === ack.version &&
      ack.frame.coversThroughSeq === ack.throughSeq && ack.frame.displayList.commands.some((command) =>
        command.op === 'text' && command.text === 'local'))).toBe(true);
    merged = whenAcknowledged(right);
    const undo = left.editSession!.submit('undo', { payload: { args: [] } });
    expect((await undo.ack).status).toBe('applied');
    await merged;
    expect((await right.call.cellInputs(0, 'B3')).cells[0][0].input).toBe('seed');
    expect(remoteUpdates()).toBe(before);
    const concurrent = Promise.all([whenAcknowledged(left, left.throughSeq + 2),
      whenAcknowledged(right, right.throughSeq + 2)]);
    await Promise.all([edit(left, 'left'), edit(right, 'right')]);
    await concurrent;
    expect(updates).toContain('local');
    expect((await left.call.cellInputs(0, 'B3')).cells).toEqual((await right.call.cellInputs(0, 'B3')).cells);
    hold = true;
    const accepted = acks.length;
    const queued = whenAcknowledged(right, right.throughSeq + 2);
    await edit(left, 'queued remote');
    await whenHeld(1);
    merged = whenAcknowledged(left);
    const local = edit(right, 'queued local');
    await whenHeld(2);
    expect(delayed).toHaveLength(2);
    expect(acks).toHaveLength(accepted);
    hold = false;
    for (const reply of delayed.splice(0)) reply();
    await Promise.all([local, queued, merged]);
    await Promise.all([left.call.version(), right.call.version()]);
    expect((await left.call.cellInputs(0, 'B3')).cells).toEqual((await right.call.cellInputs(0, 'B3')).cells);
    const structural = right.editSession!.submit('applyOps', {
      payload: { args: [[{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]] },
    });
    expect(await structural.ack).toMatchObject({ status: 'refused', reason: 'unsupported' });
    const calculation = right.editSession!.submit('editCell', {
      payload: { args: [0, 2, 1, 'fixed'], calculation: { nowSerial: 0, randSeed: 1 } },
    });
    expect(await calculation.ack).toMatchObject({ status: 'refused', reason: 'unsupported' });
    await expect(right.replica!.async.applyUpdate(Uint8Array.of(0xff))).rejects.toThrow('invalid-update');
    expect(right.failure).toBeUndefined();
    merged = whenAcknowledged(left);
    await edit(right, 'after refusal');
    await merged;
    expect((await left.call.cellInputs(0, 'B3')).cells).toEqual((await right.call.cellInputs(0, 'B3')).cells);
    expect(acks.every((seq, index) => index === 0 || seq === acks[index - 1] + 1)).toBe(true);
    expect(right.throughSeq).toBe(acks[acks.length - 1]);
  } finally {
    for (const provider of providers) provider.destroy();
    await left.dispose();
    await right?.dispose();
  }
});

test('collaborative worker replicas sync freeze panes and undo', async () => {
  const left = await createTestWorkbookSession(bytes, undefined, { collaborative: true, clientId: 301 });
  const right = await createTestWorkbookSession(bytes, undefined, { collaborative: true, clientId: 302 });
  const wire = transports();
  const providers = [
    new CollaborationProvider(left.replica!, wire.pair[0]),
    new CollaborationProvider(right.replica!, wire.pair[1]),
  ];
  try {
    const synced = Promise.all(providers.map(whenSynced));
    for (const provider of providers) provider.connect();
    await synced;
    expect(await left.call.sheetView(0)).toMatchObject({ frozenRows: 0, frozenCols: 0, frozenHeight: 0 });
    expect(await right.call.sheetView(0)).toMatchObject({ frozenRows: 0, frozenCols: 0, frozenHeight: 0 });

    let merged = whenAcknowledged(right);
    const freeze = left.editSession!.submit('applyOps', {
      payload: {
        args: [[{ type: 'setFreezePane', sheet: 0, pane: { rows: 1, cols: 0, top_left: { row: 1, col: 0 } } }]],
      },
    });
    expect((await freeze.ack).status).toBe('applied');
    await merged;
    const leftFrozen = await left.call.sheetView(0);
    const rightFrozen = await right.call.sheetView(0);
    expect(leftFrozen).toMatchObject({ frozenRows: 1, frozenCols: 0 });
    expect(leftFrozen.frozenHeight).toBeGreaterThan(0);
    expect(rightFrozen).toMatchObject({
      frozenRows: leftFrozen.frozenRows,
      frozenCols: leftFrozen.frozenCols,
      frozenWidth: leftFrozen.frozenWidth,
      frozenHeight: leftFrozen.frozenHeight,
    });

    merged = whenAcknowledged(right);
    const undo = left.editSession!.submit('undo', { payload: { args: [] } });
    expect((await undo.ack).status).toBe('applied');
    await merged;
    expect(await left.call.sheetView(0)).toMatchObject({ frozenRows: 0, frozenCols: 0, frozenHeight: 0 });
    expect(await right.call.sheetView(0)).toMatchObject({ frozenRows: 0, frozenCols: 0, frozenHeight: 0 });
  } finally {
    for (const provider of providers) provider.destroy();
    await left.dispose();
    await right.dispose();
  }
});
