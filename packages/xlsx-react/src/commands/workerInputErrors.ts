import { XlsxCommandAdmissionError, type XlsxCommandAdmissionCode } from './createXlsxCommandStore';

/** The English notice for an edit sheet protection refuses. */
export const SHEET_PROTECTED_MESSAGE = 'This cell is on a protected sheet.';

/** Whether an engine refusal names a locked cell or a protected sheet. */
export function isSheetProtectedMessage(message: string): boolean {
  return /^(?:[A-Z]+[1-9]\d* is locked on protected sheet |sheet [\s\S]+ is protected$)/.test(message);
}

/** An engine refusal a worker preview carried: its failure code and message. */
export interface WorkerPreviewRefusal {
  readonly code: string;
  readonly message: string;
}

/** The refusal a worker preview of an edit carried, when `error` is one. */
export function previewRefusal(error: unknown): WorkerPreviewRefusal | null {
  if (!(error instanceof Error) || error.name !== 'WorkbookPreviewRefusal') return null;
  const failure = (error as { refusal?: { failure?: { code?: unknown; message?: unknown } } }).refusal?.failure;
  const code = (error as { code?: unknown }).code;
  return {
    code: typeof code === 'string' ? code : typeof failure?.code === 'string' ? failure.code : 'refused',
    message: typeof failure?.message === 'string' ? failure.message : error.message,
  };
}

/**
 * Whether `error` is the engine refusing a user edit for sheet protection. A typed refusal is
 * classified by its `locked-target` code; locked cells also cover merged and array cells, which
 * only the engine's text tells apart, as it does for a thrown engine error that has no code.
 */
export function isSheetProtectedError(error: unknown): boolean {
  const refused = previewRefusal(error);
  if (refused) return refused.code === 'locked-target' && isSheetProtectedMessage(refused.message);
  return error instanceof Error && isSheetProtectedMessage(error.message);
}

export class WorkerInputRefusal extends XlsxCommandAdmissionError {
  constructor(message: string, readonly barrier = false, code: XlsxCommandAdmissionCode = 'input-failed') {
    super(code);
    this.message = message;
  }
}

export class WorkerInputNotReadyError extends Error {
  constructor() {
    super('The edit peer and accepted input must be flushed before a synchronous mutation');
    this.name = 'WorkerInputNotReadyError';
  }
}
