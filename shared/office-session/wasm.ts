import { isClientMessage, type HostMessage } from './protocol';
import type { SessionTransport } from './transport';

export async function compileWasm(url: URL): Promise<WebAssembly.Module> {
  const response = await fetch(url);
  if (typeof WebAssembly.compileStreaming === 'function') {
    try {
      return await WebAssembly.compileStreaming(response.clone());
    } catch (error) {
      if (!response.ok || !['basic', 'cors', 'default'].includes(response.type) ||
        response.headers.get('Content-Type') === 'application/wasm') throw error;
    }
  }
  return WebAssembly.compile(await response.arrayBuffer());
}

export function createWorkerWasmInitializer(
  transport: SessionTransport,
  url: URL,
  initialize: (input: ArrayBuffer | WebAssembly.Module) => Promise<void>,
  compile: (url: URL) => Promise<WebAssembly.Module> = compileWasm
): (input?: ArrayBuffer | WebAssembly.Module) => Promise<void> {
  let pending: Promise<WebAssembly.Module> | undefined;
  let advertised = false;

  function module(): Promise<WebAssembly.Module> {
    if (!pending) {
      pending = compile(url);
      void pending.catch(() => {});
    }
    return pending;
  }

  transport.listen((message) => {
    if (isClientMessage(message) && message.kind === 'wasm-compile') module();
  });
  return async (input) => {
    if (input !== undefined) return initialize(input);
    const compiling = module();
    let compiled: WebAssembly.Module;
    try {
      compiled = await compiling;
    } catch (error) {
      if (pending === compiling) pending = undefined;
      throw error;
    }
    await initialize(compiled);
    if (!advertised) {
      advertised = true;
      const message: HostMessage = { protocol: 1, kind: 'wasm-module', url: url.href, module: compiled };
      try { transport.post(message); } catch {}
    }
  };
}
