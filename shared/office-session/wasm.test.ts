import { describe, expect, it, mock, spyOn } from 'bun:test';
import { runInNewContext } from 'node:vm';
import { createSessionClient, supplyWasmFetch } from './client';
import { isClientMessage, isHostMessage, type ClientMessage, type HostMessage } from './protocol';
import { mainThreadWasmSpies } from './testing/wasmWorker';
import type { SessionTransport } from './transport';
import { compileWasm, compileWasmAsset, createWorkerWasmInitializer } from './wasm';

const WASM_BYTES = new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0]);
const URL_INPUT = new URL('https://example.test/module.wasm');

function transport(messages: HostMessage[]): SessionTransport & { receive(message: unknown): void } {
  let receive = (_message: unknown) => {};
  return {
    post: (message) => { messages.push(message as HostMessage); },
    listen: (listener) => { receive = listener; return () => { receive = () => {}; }; },
    onError: () => () => {}, close() {},
    receive: (message) => receive(message),
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
  it('keeps legacy module callbacks compatible while forwarding optional hydration headers', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const wire = transport([]);
    const onWasmModule = mock((_url: string, _module: WebAssembly.Module) => {});
    const onWasmModuleMessage = mock((_message: HostMessage) => {});
    const client = createSessionClient<{ open(): string }, {}>(wire, {
      methods: { open: true }, onWasmModule, onWasmModuleMessage,
    });
    try {
      const open = client.call.open();
      const legacy = { protocol: 1 as const, kind: 'wasm-module' as const, url: URL_INPUT.href, module };
      wire.receive(legacy);
      expect(onWasmModule.mock.calls).toEqual([[URL_INPUT.href, module]]);
      expect(onWasmModuleMessage).toHaveBeenCalledWith(legacy);
      const retained = { ...legacy, hydration: 'opaque hydration', version: 'v1', sequence: 7 };
      expect(isHostMessage(retained)).toBe(true);
      wire.receive(retained);
      expect(onWasmModule.mock.calls).toEqual([[URL_INPUT.href, module], [URL_INPUT.href, module]]);
      expect(onWasmModuleMessage).toHaveBeenLastCalledWith(retained);
      wire.receive({ protocol: 1, kind: 'reply', id: 1, ok: true, value: 'opened' });
      expect(await open).toBe('opened');
      expect(client.failure).toBeUndefined();
    } finally { await client.dispose(); }
  });

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
  it('starts compilation on the message and open awaits it before initialization', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const messages: HostMessage[] = [];
    let resolve!: (module: WebAssembly.Module) => void;
    const pending = new Promise<WebAssembly.Module>((done) => { resolve = done; });
    const compile = mock(() => pending);
    let finishInitialize!: () => void;
    const initialized = new Promise<void>((done) => { finishInitialize = done; });
    const initialize = mock(() => initialized);
    const t = transport(messages);
    const init = createWorkerWasmInitializer(t, URL_INPUT, initialize, compile);
    expect(compile).not.toHaveBeenCalled();
    t.receive({ protocol: 1, kind: 'wasm-compile' });
    expect(compile).toHaveBeenCalledTimes(1);
    expect(compile).toHaveBeenCalledWith(URL_INPUT);
    expect(initialize).not.toHaveBeenCalled();
    const open = init();
    await Promise.resolve();
    expect(initialize).not.toHaveBeenCalled();
    expect(messages).toEqual([]);
    resolve(module);
    await pending;
    expect(initialize).toHaveBeenCalledWith(module);
    expect(messages).toEqual([]);
    finishInitialize();
    await open;
    expect(messages).toEqual([{ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }]);
    await init();
    expect(initialize).toHaveBeenCalledWith(module);
    expect(compile).toHaveBeenCalledTimes(1);
    expect(messages).toHaveLength(1);
  });

  it('compiles lazily in open when no message or input was supplied', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    const initialize = mock(() => Promise.resolve());
    const init = createWorkerWasmInitializer(transport(messages), URL_INPUT, initialize, compile);
    expect(compile).not.toHaveBeenCalled();
    await init();
    expect(compile).toHaveBeenCalledTimes(1);
    expect(compile).toHaveBeenCalledWith(URL_INPUT);
    expect(initialize).toHaveBeenCalledWith(module);
    expect(messages).toEqual([{ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }]);
  });

  it('does not compile or publish when a module or bytes are supplied', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    const initialize = mock(() => Promise.resolve());
    const init = createWorkerWasmInitializer(transport(messages), URL_INPUT, initialize, compile);
    await init(module);
    await init(WASM_BYTES.buffer);
    expect(initialize).toHaveBeenNthCalledWith(1, module);
    expect(initialize).toHaveBeenNthCalledWith(2, WASM_BYTES.buffer);
    expect(compile).not.toHaveBeenCalled();
    expect(messages).toEqual([]);
  });

  it.each(['module', 'bytes'])('uses explicit %s during a compile and publishes the result once', async (kind) => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const input = kind === 'module' ? new WebAssembly.Module(WASM_BYTES) : WASM_BYTES.buffer;
    const messages: HostMessage[] = [];
    let resolve!: (module: WebAssembly.Module) => void;
    const pending = new Promise<WebAssembly.Module>((done) => { resolve = done; });
    const compile = mock(() => pending);
    const initialize = mock(() => Promise.resolve());
    const t = transport(messages);
    const init = createWorkerWasmInitializer(t, URL_INPUT, initialize, compile);
    t.receive({ protocol: 1, kind: 'wasm-compile' });
    await init(input);
    expect(initialize).toHaveBeenCalledWith(input);
    expect(messages).toEqual([]);
    t.receive({ protocol: 1, kind: 'wasm-compile' });
    expect(compile).toHaveBeenCalledTimes(1);
    resolve(module);
    await pending;
    expect(messages).toEqual([]);
    t.receive({ protocol: 1, kind: 'wasm-compile' });
    await init();
    expect(initialize).toHaveBeenNthCalledWith(2, module);
    expect(compile).toHaveBeenCalledTimes(1);
    expect(messages).toEqual([{ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }]);
  });

  it('propagates initialization failure without advertising and retains the compiled module', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const error = new Error('Initialize failed');
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    const initialize = mock(() => Promise.resolve());
    initialize.mockRejectedValueOnce(error);
    const t = transport(messages);
    const init = createWorkerWasmInitializer(t, URL_INPUT, initialize, compile);
    t.receive({ protocol: 1, kind: 'wasm-compile' });
    await expect(init()).rejects.toBe(error);
    expect(initialize).toHaveBeenCalledWith(module);
    expect(compile).toHaveBeenCalledTimes(1);
    expect(messages).toEqual([]);

    await init();
    expect(initialize).toHaveBeenCalledTimes(2);
    expect(compile).toHaveBeenCalledTimes(1);
    expect(messages).toEqual([{ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module }]);
  });

  it('retains an eager compile error until open and retries the next attempt', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const error = new Error('Compile failed');
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    compile.mockRejectedValueOnce(error);
    const initialize = mock(() => Promise.resolve());
    const t = transport(messages);
    const init = createWorkerWasmInitializer(t, URL_INPUT, initialize, compile);
    t.receive({ protocol: 1, kind: 'wasm-compile' });
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

  it('ignores malformed compile messages until open', async () => {
    const module = new WebAssembly.Module(WASM_BYTES);
    const messages: HostMessage[] = [];
    const compile = mock(() => Promise.resolve(module));
    const initialize = mock(() => Promise.resolve());
    const t = transport(messages);
    const init = createWorkerWasmInitializer(t, URL_INPUT, initialize, compile);
    t.receive({ protocol: 2, kind: 'wasm-compile' });
    t.receive({ kind: 'wasm-compile' });
    t.receive({ protocol: 1, kind: 'wasm-module', url: URL_INPUT.href, module });
    expect(compile).not.toHaveBeenCalled();
    await init();
    expect(compile).toHaveBeenCalledTimes(1);
  });
});

describe('worker wasm compilation', () => {
  it('compiles file assets from disk without fetching or instantiating', async () => {
    const original = Object.getOwnPropertyDescriptor(process, 'getBuiltinModule');
    const url = new URL('file:///worker%20assets/module.wasm');
    const readFileSync = mock(() => WASM_BYTES);
    const fileURLToPath = mock(() => '/worker assets/module.wasm');
    const fetch = spyOn(globalThis, 'fetch');
    const instantiate = spyOn(WebAssembly, 'instantiate');
    Object.defineProperty(process, 'getBuiltinModule', { configurable: true, value: (name: string) => {
      if (name === 'node:fs') return { readFileSync };
      if (name === 'node:url') return { fileURLToPath };
      throw new Error(`Unexpected builtin ${name}`);
    } });
    try {
      expect(await compileWasmAsset(url)).toBeInstanceOf(WebAssembly.Module);
      expect(fileURLToPath).toHaveBeenCalledWith(url.href);
      expect(readFileSync).toHaveBeenCalledWith('/worker assets/module.wasm');
      expect(fetch).not.toHaveBeenCalled();
      expect(instantiate).not.toHaveBeenCalled();
    } finally {
      if (original) Object.defineProperty(process, 'getBuiltinModule', original);
      else Reflect.deleteProperty(process, 'getBuiltinModule');
      fetch.mockRestore();
      instantiate.mockRestore();
    }
  });

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

describe('engine fetched by the client', () => {
  const settle = () => new Promise((resolve) => setTimeout(resolve, 5));
  const engine = () => new Response(WASM_BYTES.slice(), { headers: { 'Content-Type': 'application/wasm' } });
  function workerSide(compile = mock(async (_url: URL) => new WebAssembly.Module(WASM_BYTES))) {
    const wire = transport([]);
    const initialize = mock(async (_input: ArrayBuffer | WebAssembly.Module) => {});
    const init = createWorkerWasmInitializer(wire, URL_INPUT, initialize, compile);
    return { wire, initialize, init, compile };
  }
  const isModule = (value: unknown) => Object.prototype.toString.call(value) === '[object WebAssembly.Module]';

  it('accepts an announced body, then compiles the stream or bytes without fetching', async () => {
    for (const body of ['stream', 'bytes'] as const) {
      const messages: HostMessage[] = [];
      const wire = transport(messages);
      const initialize = mock(async (_input: ArrayBuffer | WebAssembly.Module) => {});
      const compile = mock(async (_url: URL) => new WebAssembly.Module(WASM_BYTES));
      const init = createWorkerWasmInitializer(wire, URL_INPUT, initialize, compile);
      wire.receive({ protocol: 1, kind: 'wasm-compile', pending: true });
      expect(messages).toEqual([{ protocol: 1, kind: 'wasm-accept' }]);
      const opening = init();
      await settle();
      expect(initialize).not.toHaveBeenCalled();
      wire.receive(body === 'stream' ? { protocol: 1, kind: 'wasm-compile', stream: engine().body }
        : { protocol: 1, kind: 'wasm-compile', bytes: WASM_BYTES.slice().buffer });
      await opening;
      expect(isModule(initialize.mock.calls[0]![0])).toBe(true);
      expect(compile).not.toHaveBeenCalled();
    }
  });

  it('compiles itself when the client gives up, and ignores a late body once compiling', async () => {
    const fallback = workerSide();
    fallback.wire.receive({ protocol: 1, kind: 'wasm-compile', pending: true });
    const waiting = fallback.init();
    fallback.wire.receive({ protocol: 1, kind: 'wasm-compile' });
    await waiting;
    expect(fallback.compile).toHaveBeenCalledTimes(1);
    const late = workerSide();
    late.wire.receive({ protocol: 1, kind: 'wasm-compile' });
    const stream = engine().body!;
    late.wire.receive({ protocol: 1, kind: 'wasm-compile', stream });
    await late.init();
    expect(late.compile).toHaveBeenCalledTimes(1);
    await settle();
    expect(stream.locked).toBe(false);
  });

  it('validates the new fields', () => {
    expect(isClientMessage({ protocol: 1, kind: 'wasm-compile', pending: true })).toBe(true);
    expect(isClientMessage({ protocol: 1, kind: 'wasm-compile', stream: engine().body })).toBe(true);
    expect(isClientMessage({ protocol: 1, kind: 'wasm-compile', bytes: new ArrayBuffer(8) })).toBe(true);
    expect(isClientMessage({ protocol: 1, kind: 'wasm-compile', pending: false })).toBe(false);
    expect(isClientMessage({ protocol: 1, kind: 'wasm-compile', stream: {} })).toBe(false);
    expect(isClientMessage({ protocol: 1, kind: 'wasm-compile', bytes: 'x' })).toBe(false);
  });

  function clientSide(cloneFails = false) {
    const posted: { message: ClientMessage; transfer?: Transferable[] }[] = [];
    const listeners = new Set<(message: unknown) => void>();
    let closed = false;
    const wire: SessionTransport = {
      post: (message, transfer) => {
        if (closed) throw new Error('closed');
        if (cloneFails && transfer?.some((item) => item instanceof ReadableStream)) {
          throw new DOMException('stream not transferable', 'DataCloneError');
        }
        posted.push({ message: message as ClientMessage, transfer });
      },
      listen: (listener) => { listeners.add(listener); return () => { listeners.delete(listener); }; },
      onError: () => () => {}, close() { closed = true; },
    };
    const receive = (message: HostMessage) => { for (const listener of [...listeners]) listener(message); };
    return { wire, posted, receive, accept: () => receive({ protocol: 1, kind: 'wasm-accept' }) };
  }

  it('hands the body over as a stream, or as bytes where streams cannot be transferred, never compiling here', async () => {
    const spies = mainThreadWasmSpies();
    const fetched = spyOn(globalThis, 'fetch').mockImplementation(async () => engine());
    try {
      const streaming = clientSide();
      supplyWasmFetch(streaming.wire, new URL('https://example.test/stream.wasm'));
      const copying = clientSide(true);
      supplyWasmFetch(copying.wire, new URL('https://example.test/bytes.wasm'));
      await settle();
      expect(streaming.posted).toHaveLength(1);
      streaming.accept();
      copying.accept();
      await settle();
      expect(streaming.posted[0]!.message).toEqual({ protocol: 1, kind: 'wasm-compile', pending: true });
      const stream = (streaming.posted[1]!.message as { stream?: ReadableStream }).stream;
      expect(stream).toBeInstanceOf(ReadableStream);
      expect(streaming.posted[1]!.transfer).toEqual([stream]);
      const bytes = (copying.posted[1]!.message as { bytes?: ArrayBuffer }).bytes;
      expect(bytes?.byteLength).toBe(WASM_BYTES.byteLength);
      expect(copying.posted[1]!.transfer).toEqual([bytes]);
      expect(fetched).toHaveBeenCalledTimes(2);
      for (const spy of spies) expect(spy).not.toHaveBeenCalled();
    } finally {
      fetched.mockRestore();
      for (const spy of spies) spy.mockRestore();
    }
  });

  it('never hands a body to a worker from before the handoff, and cancels it when that worker compiles its own', async () => {
    const body = engine();
    const cancelled = spyOn(body.body!, 'cancel');
    const fetched = spyOn(globalThis, 'fetch').mockImplementation(async () => body);
    try {
      const old = clientSide();
      supplyWasmFetch(old.wire, new URL('https://example.test/old-worker.wasm'));
      await settle();
      old.receive({ protocol: 1, kind: 'wasm-module', url: 'https://example.test/old-worker.wasm',
        module: new WebAssembly.Module(WASM_BYTES) });
      old.accept();
      await settle();
      expect(old.posted.map((entry) => entry.message)).toEqual([{ protocol: 1, kind: 'wasm-compile', pending: true }]);
      expect(cancelled).toHaveBeenCalledTimes(1);
    } finally { fetched.mockRestore(); }
  });

  it('aborts a fetch still waiting for headers when an old worker compiles first, and cancels a late body', async () => {
    let signal: AbortSignal | undefined;
    let deliver!: (response: Response) => void;
    const fetched = spyOn(globalThis, 'fetch').mockImplementation((_input, init) => {
      signal = init?.signal ?? undefined;
      return new Promise<Response>((resolve) => { deliver = resolve; });
    });
    try {
      const old = clientSide();
      supplyWasmFetch(old.wire, new URL('https://example.test/early-module.wasm'));
      await settle();
      old.receive({ protocol: 1, kind: 'wasm-module', url: 'https://example.test/early-module.wasm',
        module: new WebAssembly.Module(WASM_BYTES) });
      expect(signal?.aborted).toBe(true);
      const late = engine();
      const cancelled = spyOn(late.body!, 'cancel');
      deliver(late);
      await settle();
      expect(cancelled).toHaveBeenCalledTimes(1);
      expect(old.posted.map((entry) => entry.message)).toEqual([{ protocol: 1, kind: 'wasm-compile', pending: true }]);
    } finally { fetched.mockRestore(); }
  });

  it('leaves the worker its own path after a failed fetch or off http(s), and stops on abort', async () => {
    const unhandled = mock(() => {});
    process.on('unhandledRejection', unhandled);
    let signal: AbortSignal | undefined;
    const fetched = spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
      if (String(input).includes('missing')) return new Response('', { status: 404 });
      if (String(input).includes('down')) throw new TypeError('network');
      signal = init?.signal ?? undefined;
      return new Promise<Response>((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')));
      });
    });
    try {
      const plain = [{ protocol: 1, kind: 'wasm-compile', pending: true }, { protocol: 1, kind: 'wasm-compile' }];
      const missing = clientSide();
      supplyWasmFetch(missing.wire, new URL('https://example.test/missing.wasm'));
      const down = clientSide();
      supplyWasmFetch(down.wire, new URL('https://example.test/down.wasm'));
      const local = clientSide();
      supplyWasmFetch(local.wire, new URL('file:///engine.wasm'));
      const aborted = clientSide();
      const abort = supplyWasmFetch(aborted.wire, new URL('https://example.test/slow.wasm'));
      await settle();
      abort();
      await settle();
      expect(missing.posted.map((entry) => entry.message)).toEqual(plain);
      expect(down.posted.map((entry) => entry.message)).toEqual(plain);
      expect(local.posted.map((entry) => entry.message)).toEqual([{ protocol: 1, kind: 'wasm-compile' }]);
      expect(signal?.aborted).toBe(true);
      expect(aborted.posted.map((entry) => entry.message)).toEqual([{ protocol: 1, kind: 'wasm-compile', pending: true }]);
      expect(unhandled).not.toHaveBeenCalled();
    } finally { fetched.mockRestore(); process.off('unhandledRejection', unhandled); }
  });
});
