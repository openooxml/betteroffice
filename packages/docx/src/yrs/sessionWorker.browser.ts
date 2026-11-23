import { attachFontPackSource } from '../../../../shared/office-session/fontPackSource';
import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  requireWebWorker('createYrsSession');
  const worker = new Worker(new URL('./docxWorker.mjs', import.meta.url), {
    type: 'module',
    name: 'betteroffice-docx-worker',
  });
  attachFontPackSource(worker);
  return worker;
}
