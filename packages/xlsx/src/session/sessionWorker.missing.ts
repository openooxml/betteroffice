import { missingWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  throw missingWebWorker('openWorkbook');
}
