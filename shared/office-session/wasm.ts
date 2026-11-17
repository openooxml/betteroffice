import { isClientMessage, type HostMessage } from './protocol';
import type { SessionTransport } from './transport';
import { readWasmSync } from './readWasmAsset';

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

export async function compileWasmAsset(url: URL): Promise<WebAssembly.Module> {
  const bytes = readWasmSync(url);
  return bytes ? WebAssembly.compile(bytes as BufferSource) : compileWasm(url);
}

export function compileSessionWasm(url: URL): Promise<WebAssembly.Module> {
  return typeof Worker === 'function' ? compileWasm(url) : compileWasmAsset(url);
}

/** Compiles an engine from a response body, streaming where the runtime can. */
async function compileStream(stream: ReadableStream<Uint8Array>): Promise<WebAssembly.Module> {
  const response = new Response(stream, { headers: { 'Content-Type': 'application/wasm' } });
  if (typeof WebAssembly.compileStreaming === 'function') return WebAssembly.compileStreaming(response);
  return WebAssembly.compile(await response.arrayBuffer());
}

export function createWorkerWasmInitializer(
  transport: SessionTransport,
  url: URL,
  initialize: (input: ArrayBuffer | WebAssembly.Module) => Promise<void>,
  compile: (url: URL) => Promise<WebAssembly.Module> = compileWasm
): (input?: ArrayBuffer | WebAssembly.Module) => Promise<void> {
  let pending: Promise<WebAssembly.Module> | undefined;
  let supply: ((module: Promise<WebAssembly.Module>) => void) | undefined;
  let advertised = false;

  function module(): Promise<WebAssembly.Module> {
    if (!pending) {
      pending = compile(url);
      void pending.catch(() => {});
    }
    return pending;
  }

  transport.listen((message) => {
    if (!isClientMessage(message) || message.kind !== 'wasm-compile') return;
    if (message.pending) {
      if (!pending) {
        pending = new Promise((resolve) => { supply = resolve; });
        void pending.catch(() => {});
        try { transport.post({ protocol: 1, kind: 'wasm-accept' } satisfies HostMessage); } catch {}
      }
      return;
    }
    const waiting = supply;
    supply = undefined;
    if (waiting) {
      const next = message.stream ? compileStream(message.stream)
        : message.bytes ? WebAssembly.compile(message.bytes) : compile(url);
      void next.catch(() => {});
      waiting(next);
    } else if (message.stream || message.bytes) {
      if (!pending) {
        pending = message.stream ? compileStream(message.stream) : WebAssembly.compile(message.bytes!);
        void pending.catch(() => {});
      } else void message.stream?.cancel().catch(() => {});
    } else module();
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
