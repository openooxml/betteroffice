import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { sessionWorkerScope, compileSessionWasm } from '#session-scope';
import { createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { sessionWasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createPresentationSessionHost, initializePresentationEditorWasm } from './host';
import { listenPptxWorkerRender } from './workerRender';

function startWorker(scope: SessionScope): void {
  const workerRender = listenPptxWorkerRender(createScopeTransport(scope));
  const transport = workerRender.transport;
  const host = createPresentationSessionHost(transport, {
    workerRender,
    initWasm: engineKind === 'native' ? initWasm : createWorkerWasmInitializer(transport, sessionWasmAssetUrl(), initWasm, compileSessionWasm),
    initEditorWasm: initializePresentationEditorWasm,
  });
  Object.assign(globalThis, { __betterofficeEngineProgress: () => host.progress() });
}

startWorker(sessionWorkerScope());
