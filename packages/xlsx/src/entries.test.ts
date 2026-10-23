import { expect, test } from 'bun:test';
import * as xlsx from './index';
import * as worker from './worker';

const publicWorker = ['WorkbookEditPeerFailedError', 'WorkbookPeerHydrationError', 'WorkbookPngUnavailableError',
  'WorkbookRecoveryReplayError', 'createWorkbookEditSession', 'openWorkbookSession', 'workbookSessionRecovery'];
const internalWorker = ['WorkbookRecoveryRefusal', 'createWorkbookEditPeer', 'createWorkbookRecoveryMutators',
  'failWorkbookEditPeer', 'hydratePeer', 'localNowSerial', 'workbookEditPeerOperations', 'workbookSessionInternals'];

test('root re-exports the public worker entry', () => {
  expect(Object.keys(worker).sort()).toEqual([...publicWorker, ...internalWorker].sort());
  for (const name of publicWorker) expect(xlsx[name as keyof typeof xlsx]).toBe(worker[name as keyof typeof worker]);
});

test('worker plumbing stays off the root entry', () => {
  for (const name of internalWorker) expect(name in xlsx).toBe(false);
});
