import { expect, test } from 'bun:test';
import { detachedBuffers, freezeDeep } from '../../../../shared/office-session/testing/frozen';
import { createResidentCommandConnection } from './residentCommandConnection';
import type { DocxCommandEnvelope, DocxCommandPayload, ResidentCommandInitRequest } from './residentCommandProtocol';
import type { ResidentEngineWorkerPort } from './residentEngineWorkerClient';
import type { ResidentEngineWorkerResponse } from './residentEngineWorkerProtocol';
import { ResidentSourceWorkerClient } from './residentSourceWorkerClient';

const INIT: ResidentCommandInitRequest = {
  session: 'docx-frozen', incarnation: 1, epoch: 1, readOnly: false,
  fonts: [], layoutInput: '{}', layoutWithRegions: false, layoutRevision: 1, extras: '{}',
};
const identity = { session: INIT.session, incarnation: INIT.incarnation, epoch: INIT.epoch };
const point = { story: 'body', paraId: 'p1', offset: 0 };

function envelope(seq: number, kind: string, payload: DocxCommandPayload): DocxCommandEnvelope {
  return { session: INIT.session, incarnation: INIT.incarnation, seq, basis: { epoch: INIT.epoch, version: 0 }, kind, payload };
}

function copyingPort() {
  const received: DocxCommandEnvelope[] = [];
  const port: ResidentEngineWorkerPort = {
    onmessage: null, onerror: null, onmessageerror: null,
    postMessage(message, transfer = []) {
      const request = structuredClone(message, { transfer });
      if (!('id' in request)) return;
      let response: ResidentEngineWorkerResponse | undefined;
      if (request.type === 'commandInit') {
        response = { id: request.id, ok: true, commandInitial: { ...identity, version: 0, throughSeq: 0, outcomes: [],
          selection: null,
          editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
            canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
          frame: { coversThroughSeq: 0, revision: 0,
            payload: { frame: new ArrayBuffer(0), caret: { frameEpoch: 0, caretRect: null }, layoutRevision: 1 } } } };
      } else if (request.type === 'edit') {
        received.push(request.command);
        const seq = request.command.seq;
        response = { id: request.id, ok: true,
          commandAck: { ...identity, version: seq, throughSeq: seq, outcomes: [{ seq, status: 'applied' }] } };
      }
      if (response) queueMicrotask(() => port.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>));
    },
    terminate() {},
  };
  return { port, received };
}

test('resident command edits leave a frozen envelope and its buffers untouched', async () => {
  const worker = copyingPort();
  const connection = await createResidentCommandConnection(new ResidentSourceWorkerClient(worker.port), INIT);
  try {
    const commands = [
      envelope(1, 'setFontEnvironment', {
        faces: [{ family: 'Probe', data: Uint8Array.of(0, 1, 0, 0) }], chains: { Probe: [0] }, families: ['Probe'],
      }),
      envelope(2, 'insertText', { text: 'a', selection: {
        anchor: { ...point, sticky: Uint8Array.of(7) }, focus: { ...point, sticky: Uint8Array.of(8) },
      } }),
    ];
    for (const command of commands) {
      const buffers = freezeDeep(command);
      expect(buffers.length).toBeGreaterThan(0);
      const ack = await connection.client.call.edit(command);
      expect(ack.outcomes).toEqual([{ seq: command.seq, status: 'applied' }]);
      expect(detachedBuffers(buffers)).toEqual([]);
    }
    expect(worker.received).toEqual(commands);
    expect(connection.client.failure).toBeUndefined();
  } finally { await connection.client.dispose(); }
});
