import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export function createSessionWorker(): Worker {
  requireWebWorker('createYrsSession');
  return new Worker(new URL('./docxWorker.mjs', import.meta.url), {
    type: 'module',
    name: 'betteroffice-docx-worker',
  });
}
