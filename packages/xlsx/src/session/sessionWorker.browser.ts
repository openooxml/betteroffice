import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  requireWebWorker('openWorkbook');
  return new Worker(new URL('./xlsxSessionWorker.mjs', import.meta.url), { type: 'module' });
}
