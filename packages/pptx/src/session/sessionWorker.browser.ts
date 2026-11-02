import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  requireWebWorker('openPresentation');
  return new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' });
}
