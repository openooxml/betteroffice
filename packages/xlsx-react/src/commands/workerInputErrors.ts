import { XlsxCommandAdmissionError, type XlsxCommandAdmissionCode } from './createXlsxCommandStore';

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
