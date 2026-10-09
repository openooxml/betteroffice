import { spyOn } from 'bun:test';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { ClientMessage, HostMessage } from '../protocol';

const WASM_BYTES = new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00]);

export interface WasmTestWorker {
  readonly requests: ClientMessage[];
  readonly sources: Array<ArrayBuffer | WebAssembly.Module>;
  readonly modules: WebAssembly.Module[];
  readonly terminated: boolean;
  crash(): void;
}

export async function sessionWasmFactory<Client>(clientPath: string, workerPath: string) {
  async function bundle(entry: string, client = false): Promise<string> {
    const result = await Bun.build({
      entrypoints: [client ? import.meta.path : entry],
      target: 'bun',
      format: 'esm',
      plugins: [{
        name: 'session-wasm-worker',
        setup(build) {
          build.onLoad({ filter: /wasmWorker\.ts$/ }, () => ({
            contents: `import * as client from ${JSON.stringify(entry)}; testHarness.client = client;`,
            loader: 'js',
          }));
          build.onLoad({ filter: /\/wasm\/asset\.ts$/ }, () => ({
            contents: 'export const wasmAssetUrl = () => new URL(testHarness.assetUrl);',
            loader: 'js',
          }));
          build.onLoad({ filter: /\/wasm\/loader\.ts$/ }, () => ({
            contents: `
              export class PresentationPeerError extends Error {
                constructor(code, message) { super(message); this.name = 'PresentationPeerError'; this.code = code; }
              }
              export const openPresentationReplayBaseline = () => { throw new Error('openPresentationReplayBaseline is not stubbed'); };
              export const presentationPeerHydration = () => { throw new Error('presentationPeerHydration is not stubbed'); };
              export const presentationPeerMetadata = () => { throw new Error('presentationPeerMetadata is not stubbed'); };
              export const presentationPeerDisplayListJson = () => { throw new Error('presentationPeerDisplayListJson is not stubbed'); };
              export const registerPresentationPeerFonts = () => { throw new Error('registerPresentationPeerFonts is not stubbed'); };
              export const replayPresentation = () => { throw new Error('replayPresentation is not stubbed'); };
              export const initWasm = (source) => testHarness.initialize(source);
              export const openPresentation = (bytes) => testHarness.open(bytes);
              export const openWorkbook = (bytes) => testHarness.open(bytes);
              export const openWorkbookPeer = openWorkbook;
              export class StaleProposalError extends Error {
                constructor(cells, targets = []) {
                  super('stale: ' + cells.join(', '));
                  this.name = 'StaleProposalError';
                  this.cells = cells;
                  this.targets = targets;
                }
              }
              export const decodeTiffImage = () => { throw new Error('decodeTiffImage is not stubbed'); };
              export const presentationDisplayListJson = () => { throw new Error('presentationDisplayListJson is not stubbed'); };
              export const presentationMetadata = () => ({ slides: [], size: { width: 0, height: 0 } });
              export const workbookDisplayListJson = () => { throw new Error('workbookDisplayListJson is not stubbed'); };
              export const workbookPeerHydration = () => { throw new Error('workbookPeerHydration is not stubbed'); };
              export const workbookPeerSnapshot = () => { throw new Error('workbookPeerSnapshot is not stubbed'); };
              export const createWorkbookSnapshotBuilder = () => { throw new Error('createWorkbookSnapshotBuilder is not stubbed'); };
            `,
            loader: 'js',
          }));
        },
      }],
    });
    if (!result.success) throw new Error(`Session worker bundle failed: ${result.logs.join('\n')}`);
    return (await result.outputs[0].text()).split('import.meta.url').join(
      JSON.stringify(pathToFileURL(entry).href)
    );
  }

  const [clientCode, workerCode] = await Promise.all([
    bundle(resolve(clientPath), true), bundle(resolve(workerPath)),
  ]);
  const startClient = new Function('testHarness', 'Worker', clientCode);
  const startWorker = new Function('self', 'WebAssembly', 'fetch', 'testHarness', workerCode);

  return () => {
    const moduleConstructor = WebAssembly.Module;
    const runtimeError = WebAssembly.RuntimeError;
    const module = new moduleConstructor(WASM_BYTES);
    const workers: WasmTestWorker[] = [];
    const trace: string[] = [];
    let compiles = 0;
    let assetUrl = 'https://example.test/session.wasm';
    let nextCompile: (() => Promise<WebAssembly.Module>) | undefined;
    let nextInitializeError: Error | undefined;
    const clientHarness = { client: undefined as unknown as Client, get assetUrl() { return assetUrl; } };

    class TestWorker implements WasmTestWorker {
      readonly requests: ClientMessage[] = [];
      readonly sources: Array<ArrayBuffer | WebAssembly.Module> = [];
      readonly modules: WebAssembly.Module[] = [];
      terminated = false;
      private readonly listeners = new Map<string, Set<EventListener>>();
      private readonly scopeListeners = new Map<string, Set<EventListener>>();

      constructor(url: URL) {
        workers.push(this);
        const worker = this;
        const scope = {
          location: { href: url.href },
          postMessage(message: HostMessage, transfer: Transferable[] = []) {
            const data = structuredClone(message, { transfer });
            if (data.kind === 'wasm-module') worker.modules.push(data.module);
            queueMicrotask(() => worker.dispatch(worker.listeners, 'message', { data }));
          },
          addEventListener(type: string, listener: EventListener) {
            worker.subscribe(worker.scopeListeners, type, listener);
          },
          removeEventListener(type: string, listener: EventListener) {
            worker.scopeListeners.get(type)?.delete(listener);
          },
          close: () => worker.terminate(),
        };
        startWorker(scope, {
          Module: moduleConstructor,
          RuntimeError: runtimeError,
          compileStreaming: () => {
            compiles += 1;
            trace.push('compile');
            const compile = nextCompile;
            nextCompile = undefined;
            return compile ? compile() : Promise.resolve(module);
          },
          compile: () => { throw new Error('Unexpected byte compile'); },
        }, () => {
          trace.push('fetch');
          return Promise.resolve(new Response(WASM_BYTES, {
            headers: { 'Content-Type': 'application/wasm' },
          }));
        }, {
          assetUrl,
          initialize(source: ArrayBuffer | WebAssembly.Module) {
            worker.sources.push(source);
            trace.push('initialize');
            if (nextInitializeError) {
              const error = nextInitializeError;
              nextInitializeError = undefined;
              return Promise.reject(error);
            }
            return Promise.resolve();
          },
          open(bytes: Uint8Array) {
            return {
              snapshot: () => ({ slides: [], widthEmu: 0, heightEmu: 0 }),
              version: () => '0',
              sheetInfo: () => ({ sheetIds: [], sheetNames: [], activeSheet: 0 }),
              save: () => bytes.slice(),
              dispose() {},
            };
          },
        });
      }

      private subscribe(listeners: Map<string, Set<EventListener>>, type: string, listener: EventListener) {
        const set = listeners.get(type) ?? new Set<EventListener>();
        set.add(listener);
        listeners.set(type, set);
      }

      private dispatch(listeners: Map<string, Set<EventListener>>, type: string, fields: object) {
        for (const listener of [...(listeners.get(type) ?? [])]) {
          if (!this.terminated && listeners.get(type)?.has(listener)) listener(fields as Event);
        }
      }

      postMessage(message: ClientMessage, transfer: Transferable[] = []) {
        this.requests.push(message);
        trace.push(`post:${message.kind === 'call' ? message.method : message.kind}`);
        const data = structuredClone(message, { transfer });
        queueMicrotask(() => {
          trace.push(`receive:${data.kind === 'call' ? data.method : data.kind}`);
          this.dispatch(this.scopeListeners, 'message', { data });
        });
      }

      addEventListener(type: string, listener: EventListener) {
        this.subscribe(this.listeners, type, listener);
      }

      removeEventListener(type: string, listener: EventListener) {
        this.listeners.get(type)?.delete(listener);
      }

      terminate() {
        this.terminated = true;
        this.listeners.clear();
        this.scopeListeners.clear();
      }

      crash() { this.dispatch(this.listeners, 'error', { message: 'Worker crashed' }); }
    }

    startClient(clientHarness, TestWorker);
    return {
      client: clientHarness.client, workers, module, trace,
      get compiles() { return compiles; },
      set assetUrl(url: string) { assetUrl = url; },
      worker: () => new TestWorker(new URL('https://example.test/worker.mjs')) as unknown as Worker,
      rejectCompile(error: Error) { nextCompile = () => Promise.reject(error); },
      rejectInitialize(error: Error) { nextInitializeError = error; },
      holdCompile() {
        let release!: (module: WebAssembly.Module) => void;
        const pending = new Promise<WebAssembly.Module>((resolve) => { release = resolve; });
        nextCompile = () => pending;
        return () => release(module);
      },
    };
  };
}

export function mainThreadWasmSpies() {
  const modulePrototype = WebAssembly.Module.prototype;
  const instancePrototype = WebAssembly.Instance.prototype;
  const module = spyOn(WebAssembly, 'Module');
  const instance = spyOn(WebAssembly, 'Instance');
  Object.defineProperty(module, 'prototype', { value: modulePrototype });
  Object.defineProperty(instance, 'prototype', { value: instancePrototype });
  return [module, instance, spyOn(WebAssembly, 'instantiate'),
    spyOn(WebAssembly, 'instantiateStreaming'), spyOn(WebAssembly, 'compile'),
    spyOn(WebAssembly, 'compileStreaming')];
}
