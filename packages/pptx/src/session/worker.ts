import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createPresentationSessionHost, initializePresentationEditorWasm } from './host';

const transport = createScopeTransport(self as unknown as SessionScope);
const host = createPresentationSessionHost(transport, {
  initWasm: engineKind === 'native' ? initWasm : createWorkerWasmInitializer(transport, wasmAssetUrl(), initWasm),
  initEditorWasm: initializePresentationEditorWasm,
});
Object.assign(globalThis, { __betterofficeEngineProgress: () => host.progress() });
