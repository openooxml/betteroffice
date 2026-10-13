import { expect, mock, test } from 'bun:test';
import * as session from '../../../../shared/office-session/commands/session';
import { isClientMessage, type ClientMessage } from '../../../../shared/office-session/protocol';
import { detachedBuffers, freezingEdits, type FrozenEnvelope } from '../../../../shared/office-session/testing/frozen';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import type { WorkbookSharedWireAck } from './commandProtocol';
import type { WorkbookSessionOpenReply } from './methods';

const actual = { ...session };
const received: FrozenEnvelope[] = [];
mock.module('../../../../shared/office-session/commands/session', () => ({
  ...actual, createWorkerEditSession: freezingEdits(actual.createWorkerEditSession, received),
}));
const { createWorkbookSession } = await import('./client');

const identity = { session: 'book', incarnation: 'opaque:worker' };
const basis = { epoch: 1, version: 'opaque:0' };
const opened: WorkbookSessionOpenReply = {
  format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0,
  sheets: [{ id: 'sheet:0', index: 0, name: 'Data' }], commandContext: { ...identity, basis },
};

function copyingWorker() {
  const calls: { message: Extract<ClientMessage, { kind: 'call' }>; transfer: Transferable[] }[] = [];
  let listener: (message: unknown) => void = () => {};
  const transport: SessionTransport = {
    post(message, transfer = []) {
      const delivered = structuredClone(message, { transfer });
      if (!isClientMessage(delivered) || delivered.kind !== 'call') return;
      if (delivered.method !== 'open') calls.push({ message: delivered, transfer });
      else queueMicrotask(() => { listener({ protocol: 1, kind: 'reply', id: delivered.id, ok: true, value: opened }); });
    },
    listen(next) { listener = next; return () => { listener = () => {}; }; },
    onError: () => () => {},
    close: () => {},
  };
  return { transport, calls,
    ack(index: number) {
      const { message } = calls[index]!;
      const seq = (message.args[0] as { seq: number }).seq;
      const status = { seq, status: 'applied' as const, committed: true };
      const compatibility = { throughSeq: seq, version: `opaque:${seq}`, statuses: [status] };
      const ack: WorkbookSharedWireAck = { session: identity.session, incarnation: 1, epoch: 1, throughSeq: seq,
        version: seq, engineVersion: `opaque:${seq}`, outcomes: [{ seq, status: 'applied', workbook: status }], compatibility };
      listener({ protocol: 1, kind: 'reply', id: message.id, ok: true,
        value: message.method === 'command' ? { ...compatibility, shared: ack } : ack });
    } };
}

test('workbook edit connections leave frozen envelopes and their buffers untouched on both compatibility branches', async () => {
  const worker = copyingWorker();
  const workbook = await createWorkbookSession(new Uint8Array(), { includeEditState: true }, worker.transport);
  try {
    const native = workbook.editSession!.submit('applyUpdate', { payload: { update: Uint8Array.of(1, 2, 3) } },
      { lane: 'collab', fence: true });
    worker.ack(0);
    expect(await native.ack).toMatchObject({ seq: 1, status: 'applied' });
    const edited = workbook.submit({ ...identity, seq: 2, afterSeq: 1, basis, kind: 'editCell',
      payload: { args: [0, 0, 0, 'x'] } });
    worker.ack(1);
    expect(await edited).toMatchObject({ seq: 2, status: 'applied' });
    const remote = workbook.submit({ ...identity, seq: 3, basis, kind: 'applyUpdate',
      payload: { update: Uint8Array.of(4, 5, 6) } });
    worker.ack(2);
    expect(await remote).toMatchObject({ seq: 3, status: 'applied' });
    expect(received.map(({ envelope }) => [envelope.seq, envelope.kind, Object.isFrozen(envelope)])).toEqual([
      [1, 'applyUpdate', true], [2, 'editCell', true], [3, 'applyUpdate', true],
    ]);
    expect(received.map(({ buffers }) => buffers.length > 0)).toEqual([true, false, true]);
    for (const { buffers } of received) expect(detachedBuffers(buffers)).toEqual([]);
    expect(worker.calls.map(({ message, transfer }) => [message.method, transfer.length])).toEqual([
      ['edit', 0], ['command', 0], ['edit', 0],
    ]);
    expect(worker.calls[1]!.message.args[0]).toMatchObject({ seq: 2, afterSeq: 1, kind: 'editCell' });
    expect(worker.calls[2]!.message.args[0]).toMatchObject({ seq: 3,
      payload: { compatibility: { payload: { update: Uint8Array.of(4, 5, 6) } } } });
    expect(workbook.failure).toBeUndefined();
  } finally { await workbook.dispose(); }
});
