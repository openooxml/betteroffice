import type { WorkbookEditPeer } from './editPeer';
import type { WorkbookReplayOp, WorkbookReplayReply } from './replay';

export interface WorkbookEditPeerOperations {
  applyQueuedOp(op: WorkbookReplayOp): WorkbookReplayReply['result'];
  applyRecoveryOp(op: WorkbookReplayOp): WorkbookReplayReply['result'];
}

export const workbookEditPeerInternals = new WeakMap<WorkbookEditPeer, WorkbookEditPeerOperations>();

export function workbookEditPeerOperations(peer: WorkbookEditPeer): WorkbookEditPeerOperations {
  const operations = workbookEditPeerInternals.get(peer);
  if (!operations) throw new TypeError('Workbook edit peer does not support retained operations');
  return operations;
}
