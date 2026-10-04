import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { initWasm } from '../wasm/loader';
import { createWorkbookSessionHost } from './host';

const transport = createScopeTransport(self as unknown as SessionScope);
const eager = self.name === 'office-session-wasm-default';
createWorkbookSessionHost(transport, {
  initWasm: createWorkerWasmInitializer(transport, wasmAssetUrl(), initWasm, eager),
});
