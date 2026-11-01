import type { SessionScope } from '../../../../shared/office-session/transport';
import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export { compileWasm as compileSessionWasm } from '../../../../shared/office-session/wasm';

export function createSessionWorker(): Worker {
  requireWebWorker('openPresentation');
  return new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' });
}

export function sessionWorkerScope(): SessionScope {
  return self as unknown as SessionScope;
}
