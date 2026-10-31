import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createPresentationSessionHost, initializePresentationEditorWasm } from './host';

function startWorker(scope: SessionScope): void {
  const transport = createScopeTransport(scope);
  const host = createPresentationSessionHost(transport, {
    initWasm: engineKind === 'native' ? initWasm : createWorkerWasmInitializer(transport, wasmAssetUrl(), initWasm),
    initEditorWasm: initializePresentationEditorWasm,
  });
  Object.assign(globalThis, { __betterofficeEngineProgress: () => host.progress() });
}

if (typeof self === 'undefined') {
  void import('../../../../shared/office-session/nodeWorker').then(({ sessionWorkerScope }) => startWorker(sessionWorkerScope()));
} else {
  startWorker(self as unknown as SessionScope);
}
