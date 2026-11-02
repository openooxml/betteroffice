import { WorkerInputNotReadyError } from '../commands/workerInputErrors';
import { asError } from './asError';

/** @experimental */
export class XlsxPeerNotReadyError extends WorkerInputNotReadyError {
  constructor() {
    super();
    this.name = 'XlsxPeerNotReadyError';
  }
}

/** @experimental */
export class XlsxWorkerEditorCollaborationError extends Error {
  readonly code = 'collaboration-unavailable';
  constructor() {
    super('Collaboration is unavailable in the worker editor');
    this.name = 'XlsxWorkerEditorCollaborationError';
  }
}

export class XlsxWorkerSessionFailedError extends Error {
  readonly code = 'worker-session-failed';
  constructor(readonly cause: unknown) {
    super(`Workbook worker session failed: ${asError(cause).message}`);
    this.name = 'XlsxWorkerSessionFailedError';
  }
}
