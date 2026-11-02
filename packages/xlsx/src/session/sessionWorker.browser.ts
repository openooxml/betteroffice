import type { SessionScope } from '../../../../shared/office-session/transport';
import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export { compileWasm as compileSessionWasm } from '../../../../shared/office-session/wasm';

export function createSessionWorker(): Worker {
  requireWebWorker('openWorkbook');
  return new Worker(new URL('./xlsxSessionWorker.mjs', import.meta.url), { type: 'module' });
}

export function sessionWorkerScope(): SessionScope {
  return self as unknown as SessionScope;
}
