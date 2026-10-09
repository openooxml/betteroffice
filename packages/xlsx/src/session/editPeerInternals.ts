import type { WorkbookEditPeer } from './editPeer';
import { WORKBOOK_REPLAY_MUTATORS, type WorkbookReplayMethod, type WorkbookReplayOp, type WorkbookReplayReply } from './replay';

export interface WorkbookEditPeerOperations {
  fail(error: unknown): void;
  whenAcknowledged(): Promise<void>;
  applyQueuedOp(op: WorkbookReplayOp): WorkbookReplayReply['result'];
  applyRecoveryOp(op: WorkbookReplayOp): WorkbookReplayReply['result'];
}

export class WorkbookRecoveryRefusal extends Error {
  constructor(readonly result: WorkbookReplayReply['result']) {
    super(`Engine refused workbook recovery: ${JSON.stringify(result)}`);
    this.name = 'WorkbookRecoveryRefusal';
  }
}

export const workbookEditPeerInternals = new WeakMap<WorkbookEditPeer, WorkbookEditPeerOperations>();

/** @internal */
export function failWorkbookEditPeer(peer: WorkbookEditPeer, error: unknown): void {
  workbookEditPeerInternals.get(peer)?.fail(error);
}

export function workbookEditPeerOperations(peer: WorkbookEditPeer): WorkbookEditPeerOperations {
  const operations = workbookEditPeerInternals.get(peer);
  if (!operations) throw new TypeError('Workbook edit peer does not support retained operations');
  return operations;
}

/** @internal */
export function createWorkbookRecoveryMutators(peer: WorkbookEditPeer): Pick<WorkbookEditPeer, WorkbookReplayMethod> {
  const operations = workbookEditPeerOperations(peer);
  return Object.fromEntries(Object.keys(WORKBOOK_REPLAY_MUTATORS).map((method) => [
    method, (...args: unknown[]) => operations.applyRecoveryOp({ method, args } as WorkbookReplayOp),
  ])) as unknown as Pick<WorkbookEditPeer, WorkbookReplayMethod>;
}
