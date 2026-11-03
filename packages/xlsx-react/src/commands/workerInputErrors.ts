import { XlsxCommandAdmissionError, type XlsxCommandAdmissionCode } from './createXlsxCommandStore';

/** The English notice for an edit sheet protection refuses. */
export const SHEET_PROTECTED_MESSAGE = 'This cell is on a protected sheet.';

/** Whether an engine refusal names a locked cell or a protected sheet. */
export function isSheetProtectedMessage(message: string): boolean {
  return /^(?:[A-Z]+[1-9]\d* is locked on protected sheet |sheet [\s\S]+ is protected$)/.test(message);
}

/** Whether `error` is the engine refusing a user edit for sheet protection. */
export function isSheetProtectedError(error: unknown): boolean {
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
