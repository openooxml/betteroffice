import { createScopeTransport, type SessionScope } from '../../../../shared/office-session/transport';
import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { compileSessionWasm, createWorkerWasmInitializer } from '../../../../shared/office-session/wasm';
import { wasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createWorkbookSessionHost } from './host';

const WORKER_STARTED = Symbol.for('betteroffice.xlsx.worker.started');

export function startSessionWorker(scope: SessionScope): void {
  const flags = scope as unknown as Record<symbol, unknown>;
  if (flags[WORKER_STARTED]) return;
  flags[WORKER_STARTED] = true;
  const transport = createScopeTransport(scope);
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
    module = await compileSessionWasm(url);
    return module;
  });
  let initializeLean: ReturnType<typeof createWorkerWasmInitializer> | undefined;
  const host = createWorkbookSessionHost(transport, {
    initWasm: engineKind === 'native' ? initWasm : (input, options) => {
      if (options?.png !== false) return initialize(input);
      initializeLean ??= createWorkerWasmInitializer(moduleTransport, wasmAssetUrl(false), initWasm, async (url) => {
        module = await compileSessionWasm(url);
        return module;
      });
      return initializeLean(input);
    },
    wasmModule: () => module,
  });
  Object.assign(globalThis, { __betterofficeEngineProgress: () => host.progress() });
}
