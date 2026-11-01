import { createSessionWorker as createThreadWorker } from '../../../../shared/office-session/nodeSessionWorker';
import { createSessionWorker as createWebWorker } from './sessionWorker.browser';

export function createSessionWorker(): Worker {
  return typeof Worker === 'function' ? createWebWorker() : createThreadWorker(
    new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' }, 'openPresentation');
}
