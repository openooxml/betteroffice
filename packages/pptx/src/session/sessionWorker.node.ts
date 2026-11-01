import {
  createSessionWorker as createThreadWorker, sessionWorkerScope as threadScope,
  compileSessionWasm as compileNodeWasm,
} from '../../../../shared/office-session/nodeSessionWorker';
import type { SessionScope } from '../../../../shared/office-session/transport';
import {
  createSessionWorker as createWebWorker, sessionWorkerScope as webScope,
  compileSessionWasm as compileWebWasm,
} from './sessionWorker.browser';

export function createSessionWorker(): Worker {
  return typeof Worker === 'function' ? createWebWorker() : createThreadWorker(
    new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' }, 'openPresentation');
}

export function sessionWorkerScope(): SessionScope {
  return typeof self === 'undefined' ? threadScope() : webScope();
}

export function compileSessionWasm(url: URL): Promise<WebAssembly.Module> {
  return typeof Worker === 'function' ? compileWebWasm(url) : compileNodeWasm(url);
}
