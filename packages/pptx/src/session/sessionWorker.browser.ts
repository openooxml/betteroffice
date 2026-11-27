import { attachFontPackSource } from '../../../../shared/office-session/fontPackSource';
import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  requireWebWorker('openPresentation');
  const worker = new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' });
  attachFontPackSource(worker);
  return worker;
}
