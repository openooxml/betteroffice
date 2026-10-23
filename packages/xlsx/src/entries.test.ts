import { expect, test } from 'bun:test';
import * as xlsx from './index';
import * as worker from './worker';

const publicWorker = ['WorkbookEditPeerFailedError', 'WorkbookPeerHydrationError', 'WorkbookPngUnavailableError',
  'WorkbookRecoveryReplayError', 'createWorkbookEditSession', 'openWorkbookSession', 'workbookSessionRecovery'];
const internalWorker = ['WorkbookRecoveryRefusal', 'createWorkbookEditPeer', 'createWorkbookRecoveryMutators',
  'failWorkbookEditPeer', 'hydratePeer', 'localNowSerial', 'workbookEditPeerOperations', 'workbookSessionInternals'];

test('root re-exports the public worker entry', () => {
  expect(Object.keys(worker).sort()).toEqual([...publicWorker, ...internalWorker].sort());
  const root = new Map<string, unknown>(Object.entries(xlsx));
  const entry = new Map<string, unknown>(Object.entries(worker));
  for (const name of publicWorker) expect(root.get(name)).toBe(entry.get(name));
});

test('worker plumbing stays off the root entry', () => {
  for (const name of internalWorker) expect(name in xlsx).toBe(false);
});
