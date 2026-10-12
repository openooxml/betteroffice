import { expect, mock, test } from 'bun:test';
import type { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import { createPeerlessCollaborationReplica } from './peerlessCollaboration';

test('typing follows the latest unacknowledged merge and the adapter returns its inference', async () => {
  let receive!: Parameters<ResidentEngineWorkerClient['onCollabUpdates']>[0];
  const client = { onCollabUpdates(listener: typeof receive) { receive = listener; return () => {}; } } as unknown as ResidentEngineWorkerClient;
  let resolve!: (outcome: { status: 'applied' }) => void;
  const ticket = { id: 1, seq: 1, ack: new Promise<{ status: 'applied' }>(done => { resolve = done; }),
    presented: Promise.resolve('presented' as const) };
  const submit = mock((..._args: Parameters<DocxWorkerEditSession['submit']>) => ticket as ReturnType<DocxWorkerEditSession['submit']>);
  const session = { submit } as unknown as DocxWorkerEditSession;
  const replica = createPeerlessCollaborationReplica(client, session, 41);
  const merged = replica.async.applyUpdate(new Uint8Array([1]));
  session.submit('insertText', { text: 'queued' });
  expect(submit.mock.calls[1]?.[2]).toEqual({ after: ticket });
  const older = { ...ticket, id: 0, seq: 0 };
  session.submit('insertText', { text: 'after older input' }, { after: older });
  expect(submit.mock.calls[2]?.[2]).toEqual({ after: ticket });
  const later = { ...ticket, id: 2, seq: 2 };
  session.submit('insertText', { text: 'after later input' }, { after: later });
  expect(submit.mock.calls[3]?.[2]).toEqual({ after: later });
  const insertion = { clientId: 42, story: 'body', paraId: 'p', endOffset: 4 };
  receive({ updates: [], inferences: [{ seq: 1, insertion }] });
  resolve({ status: 'applied' });
  expect(await merged).toEqual(insertion);
  replica.retire();
});
