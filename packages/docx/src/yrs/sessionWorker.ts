import { createDeferredSessionWorker } from '../../../../shared/office-session/deferredWorker';
import { createSessionWorker as createWebWorker } from './sessionWorker.browser';

export function createSessionWorker(): Worker {
  if (typeof Worker === 'function') return createWebWorker();
  return createDeferredSessionWorker(async () => (await import('#session-worker')).createSessionWorker(), 'createYrsSession');
}
