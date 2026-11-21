import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { compileSessionWasm, createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { sessionWasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createPresentationSessionHost, initializePresentationEditorWasm } from './host';
import { listenPptxWorkerRender } from './workerRender';

const WORKER_STARTED = Symbol.for('betteroffice.pptx.worker.started');

export function startSessionWorker(scope: SessionScope): void {
  const flags = scope as unknown as Record<symbol, unknown>;
  if (flags[WORKER_STARTED]) return;
  flags[WORKER_STARTED] = true;
  const workerRender = listenPptxWorkerRender(createScopeTransport(scope));
  const transport = workerRender.transport;
  const host = createPresentationSessionHost(transport, {
    workerRender,
    initWasm: engineKind === 'native' ? initWasm : createWorkerWasmInitializer(transport, sessionWasmAssetUrl(), initWasm, compileSessionWasm),
    initEditorWasm: initializePresentationEditorWasm,
  });
  Object.assign(globalThis, { __betterofficeEngineProgress: () => host.progress() });
}

