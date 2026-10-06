import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { compileWasm, createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { initWasm } from '../wasm/loader';
import { createWorkbookSessionHost } from './host';

const transport = createScopeTransport(self as unknown as SessionScope);
let module: WebAssembly.Module | undefined;
let retainedHydration = false;
transport.listen((message) => {
  if (isClientMessage(message) && message.kind === 'call' && message.method === 'open') {
    const input = message.args[1] as { retainPeerHydration?: boolean } | undefined;
    retainedHydration = input?.retainPeerHydration === true;
  }
});
const moduleTransport = {
  ...transport,
  post: (...args: Parameters<typeof transport.post>) => {
    if (retainedHydration && isHostMessage(args[0]) && args[0].kind === 'wasm-module') return;
    transport.post(...args);
  },
};
createWorkbookSessionHost(transport, {
  initWasm: createWorkerWasmInitializer(moduleTransport, wasmAssetUrl(), initWasm, async (url) => {
    module = await compileWasm(url);
    return module;
  }),
  wasmModule: () => module,
});
