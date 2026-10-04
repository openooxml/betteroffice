import { describe, expect, it, mock, spyOn } from 'bun:test';
import { runInNewContext } from 'node:vm';
import { createSessionClient } from './client';
import { isHostMessage, type HostMessage } from './protocol';
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

function cacheClient(onWasmModule: (url: string, module: WebAssembly.Module) => void) {
  let receive = (_message: unknown) => {};
  const client = createSessionClient<{ open(): string }, {}>({
    post() {},
    listen: (listener) => { receive = listener; return () => { receive = () => {}; }; },
    onError: () => () => {}, close() {},
  }, { methods: { open: true }, onWasmModule });
  return { client, receive: (message: unknown) => receive(message) };
}

describe('worker wasm advertisements', () => {
  it('accepts modules from another realm without interrupting open', async () => {
    const module: WebAssembly.Module = runInNewContext('new WebAssembly.Module(bytes)', { bytes: WASM_BYTES });
    expect(module instanceof WebAssembly.Module).toBe(false);
    expect(Object.prototype.toString.call(module)).toBe('[object WebAssembly.Module]');
    const onWasmModule = mock((_url: string, _module: WebAssembly.Module) => {});
    const h = cacheClient(onWasmModule);
    try {
      const open = h.client.call.open();
      const message = { protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module };
      expect(isHostMessage(message)).toBe(true);
      h.receive(message);
      expect(onWasmModule).toHaveBeenCalledWith(URL_INPUT.href, module);
      h.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'opened' });
      expect(await open).toBe('opened');
      expect(h.client.failure).toBeUndefined();
    } finally { await h.client.dispose(); }
  });

  it('ignores malformed cache advertisements while open is pending', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const onWasmModule = mock((_url: string, _module: WebAssembly.Module) => {});
    const h = cacheClient(onWasmModule);
    try {
      const open = h.client.call.open();
      for (const message of [
        { protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module: {} },
        { protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module: null },
        { protocol: 1, kind: 'wasm-module', url: URL_INPUT.href },
        { protocol: 1, kind: 'wasm-module', module },
        { protocol: 1, kind: 'wasm-module', url: 1, module },
        { protocol: 2, kind: 'wasm-module', url: URL_INPUT.href, module },
        { kind: 'wasm-module', url: URL_INPUT.href, module },
        { protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module: {
          get [Symbol.toStringTag]() { throw new Error('Unusable module'); },
        } },
      ]) {
        expect(isHostMessage(message)).toBe(false);
        expect(() => h.receive(message)).not.toThrow();
        expect(h.client.failure).toBeUndefined();
      }
      expect(onWasmModule).not.toHaveBeenCalled();
      h.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'opened' });
      expect(await open).toBe('opened');
    } finally { await h.client.dispose(); }
  });

  it('isolates a throwing cache listener from open replies', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const onWasmModule = mock(() => { throw new Error('Cache unavailable'); });
    const h = cacheClient(onWasmModule);
    try {
      const open = h.client.call.open();
      expect(() => h.receive({ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }))
        .not.toThrow();
      expect(onWasmModule).toHaveBeenCalledTimes(1);
      h.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'opened' });
      expect(await open).toBe('opened');
      expect(h.client.failure).toBeUndefined();
    } finally { await h.client.dispose(); }
  });
});

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
