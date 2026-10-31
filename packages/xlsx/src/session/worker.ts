import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { compileWasm, createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createWorkbookSessionHost } from './host';
import { renderPng } from './raster';

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
const initialize = createWorkerWasmInitializer(moduleTransport, wasmAssetUrl(), initWasm, async (url) => {
  module = await compileWasm(url);
  return module;
});
let initializeLean: ReturnType<typeof createWorkerWasmInitializer> | undefined;
const host = createWorkbookSessionHost(transport, {
  initWasm: engineKind === 'native' ? initWasm : (input, options) => {
    if (options?.png !== false) return initialize(input);
    initializeLean ??= createWorkerWasmInitializer(moduleTransport, wasmAssetUrl(false), initWasm, async (url) => {
      module = await compileWasm(url);
      return module;
    });
    return initializeLean(input);
  },
  wasmModule: () => module,
  renderPng,
});
Object.assign(globalThis, { __betterofficeEngineProgress: () => host.progress() });
