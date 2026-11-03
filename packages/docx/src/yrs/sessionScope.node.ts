import {
  compileSessionWasm as compileNodeWasm, sessionWorkerScope as threadScope,
} from '../../../../shared/office-session/nodeSessionScope';
import type { SessionScope } from '../../../../shared/office-session/transport';
import { compileSessionWasm as compileWebWasm, sessionWorkerScope as webScope } from './sessionScope.browser';

export { readWasmSync } from '../../../../shared/office-session/nodeSessionScope';

export function sessionWorkerScope(): SessionScope {
  return typeof self === 'undefined' ? threadScope() : webScope();
}

export function compileSessionWasm(url: URL): Promise<WebAssembly.Module> {
  return typeof Worker === 'function' ? compileWebWasm(url) : compileNodeWasm(url);
}
