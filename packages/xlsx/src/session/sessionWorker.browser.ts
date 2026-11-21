import { attachFontPackSource } from '../../../../shared/office-session/fontPackSource';
import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  requireWebWorker('openWorkbook');
  const worker = new Worker(new URL('./xlsxSessionWorker.mjs', import.meta.url), { type: 'module' });
  attachFontPackSource(worker);
  return worker;
}
