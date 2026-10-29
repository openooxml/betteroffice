import { XlsxCommandAdmissionError, type XlsxCommandAdmissionCode } from './createXlsxCommandStore';

export class WorkerInputRefusal extends XlsxCommandAdmissionError {
  constructor(message: string, readonly barrier = false, code: XlsxCommandAdmissionCode = 'input-failed') {
    super(code);
    this.message = message;
  }
}
