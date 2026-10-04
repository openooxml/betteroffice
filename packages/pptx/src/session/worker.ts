import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { initWasm } from '../wasm/loader';
import { createPresentationSessionHost } from './host';

const transport = createScopeTransport(self as unknown as SessionScope);
createPresentationSessionHost(transport, {
  initWasm: createWorkerWasmInitializer(transport, wasmAssetUrl(), initWasm),
});
