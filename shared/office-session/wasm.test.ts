import { describe, expect, it, mock, spyOn } from 'bun:test';
import type { HostMessage } from './protocol';
import type { SessionTransport } from './transport';
import { compileWasm, createWorkerWasmInitializer } from './wasm';

const WASM_BYTES = new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]);
const URL_INPUT = new URL('https://example.test/module.wasm');

function transport(messages: HostMessage[]): SessionTransport {
  return {
    post: (message) => { messages.push(message as HostMessage); },
    listen: () => () => {}, onError: () => () => {}, close() {},
  };
}

describe('worker wasm initialization', () => {
  it('starts compilation eagerly and publishes the module before initialization', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    const initialize = mock(() => Promise.resolve());
    const init = createWorkerWasmInitializer(transport(messages), URL_INPUT, initialize, true, compile);
    expect(compile).toHaveBeenCalledTimes(1);
    expect(initialize).not.toHaveBeenCalled();
    await Promise.resolve();
    expect(messages).toEqual([{ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }]);
    await init();
    expect(initialize).toHaveBeenCalledWith(module);
    expect(compile).toHaveBeenCalledTimes(1);
  });

  it('does not compile or publish when a module or bytes are supplied', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    const initialize = mock(() => Promise.resolve());
    const init = createWorkerWasmInitializer(transport(messages), URL_INPUT, initialize, false, compile);
    await init(module);
    await init(WASM_BYTES.buffer);
    expect(initialize).toHaveBeenNthCalledWith(1, module);
    expect(initialize).toHaveBeenNthCalledWith(2, WASM_BYTES.buffer);
    expect(compile).not.toHaveBeenCalled();
    expect(messages).toEqual([]);
  });

  it('retains an eager compile error until open and retries the next attempt', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const error = new Error('Compile failed');
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    compile.mockRejectedValueOnce(error);
    const initialize = mock(() => Promise.resolve());
    const init = createWorkerWasmInitializer(transport(messages), URL_INPUT, initialize, true, compile);
    await Promise.resolve();
    expect(messages).toEqual([]);
    await expect(init()).rejects.toBe(error);
    expect(initialize).not.toHaveBeenCalled();
    expect(compile).toHaveBeenCalledTimes(1);
    await init();
    expect(compile).toHaveBeenCalledTimes(2);
    expect(initialize).toHaveBeenCalledWith(module);
    expect(messages).toEqual([{ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }]);
  });
});

describe('worker wasm compilation', () => {
  it.each(['application/wasm', 'text/plain'])(
    'compiles a %s response without instantiating it', async (mime) => {
      const module = new WebAssembly.Module(WASM_BYTES);
      const response = new Response(WASM_BYTES, { headers: { 'Content-Type': mime } });
      const fetch = spyOn(globalThis, 'fetch').mockResolvedValue(response);
      const streaming = spyOn(WebAssembly, 'compileStreaming');
      const compile = spyOn(WebAssembly, 'compile').mockResolvedValue(module);
      const instantiate = spyOn(WebAssembly, 'instantiate');
      const instantiateStreaming = spyOn(WebAssembly, 'instantiateStreaming');
      if (mime === 'application/wasm') streaming.mockResolvedValue(module);
      else streaming.mockRejectedValue(new Error('MIME'));
      try {
        expect(await compileWasm(URL_INPUT)).toBe(module);
        expect(fetch).toHaveBeenCalledWith(URL_INPUT);
        expect(streaming).toHaveBeenCalledTimes(1);
        expect(compile).toHaveBeenCalledTimes(mime === 'application/wasm' ? 0 : 1);
        expect(instantiate).not.toHaveBeenCalled();
        expect(instantiateStreaming).not.toHaveBeenCalled();
      } finally {
        for (const spy of [fetch, streaming, compile, instantiate, instantiateStreaming]) spy.mockRestore();
      }
    }
  );

  it.each(['http', 'compile', 'opaque'])('propagates a %s failure without retrying as bytes', async (kind) => {
    const response = new Response(WASM_BYTES, {
      status: kind === 'http' ? 404 : 200,
      headers: { 'Content-Type': kind === 'compile' ? 'application/wasm' : 'text/plain' },
    });
    if (kind === 'opaque') Object.defineProperty(response, 'type', { value: 'opaque' });
    const error = new Error('Compile failed');
    const fetch = spyOn(globalThis, 'fetch').mockResolvedValue(response);
    const streaming = spyOn(WebAssembly, 'compileStreaming').mockRejectedValue(error);
    const compile = spyOn(WebAssembly, 'compile');
    try {
      await expect(compileWasm(URL_INPUT)).rejects.toBe(error);
      expect(compile).not.toHaveBeenCalled();
    } finally {
      for (const spy of [fetch, streaming, compile]) spy.mockRestore();
    }
  });
});
