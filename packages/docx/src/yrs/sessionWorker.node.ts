import { createSessionWorker as createThreadWorker } from '../../../../shared/office-session/nodeSessionWorker';

export function createSessionWorker(): Worker {
  return createThreadWorker(new URL('./docxWorker.node.mjs', import.meta.url), { type: 'module', name: 'betteroffice-docx-worker' }, 'createYrsSession');
}
