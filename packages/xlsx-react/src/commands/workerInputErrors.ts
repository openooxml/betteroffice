import { XlsxCommandAdmissionError } from './createXlsxCommandStore';

export class WorkerInputRefusal extends XlsxCommandAdmissionError {
  constructor(message: string, readonly barrier = false) {
    super('input-failed');
    this.message = message;
  }
}
