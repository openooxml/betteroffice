import type { SessionScope } from '../../../../shared/office-session/transport';
import { requireWebWorker } from '../../../../shared/office-session/webWorker';

export { compileWasm as compileSessionWasm } from '../../../../shared/office-session/wasm';

export function createSessionWorker(): Worker {
  requireWebWorker('createYrsSession');
  return new Worker(new URL('./docxWorker.mjs', import.meta.url), {
    type: 'module',
    name: 'betteroffice-docx-worker',
  });
}

export function sessionWorkerScope(): SessionScope {
  return self as unknown as SessionScope;
}

export function readWasmSync(_url: URL): Uint8Array | undefined {
  return undefined;
}
