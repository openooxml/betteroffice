import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { compileWasm, createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { initWasm } from '../wasm/loader';
import { createWorkbookSessionHost } from './host';

const transport = createScopeTransport(self as unknown as SessionScope);
let module: WebAssembly.Module | undefined;
createWorkbookSessionHost(transport, {
  initWasm: createWorkerWasmInitializer(transport, wasmAssetUrl(), initWasm, async (url) => {
    module = await compileWasm(url);
    return module;
  }),
  wasmModule: () => module,
});
