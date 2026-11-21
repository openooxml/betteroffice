import { createSessionWorker as createThreadWorker } from '../../../../shared/office-session/nodeSessionWorker';

export function createSessionWorker(): Worker {
  return createThreadWorker(new URL('./xlsxSessionWorker.node.mjs', import.meta.url), { type: 'module' }, 'openWorkbook');
}
