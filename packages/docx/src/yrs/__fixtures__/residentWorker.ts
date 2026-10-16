import { resolve } from 'node:path';
import { encodeDisplayListFrameExtras, loadRustDisplayListQueryEngine, loadedRustDisplayListQueryEngine } from '../../layout/render/rustDisplayList';
import { createResidentEngineSession, type ResidentEngineSession } from '../residentEngineSession';
import { preloadEditWasm, preloadEditWasmFrom } from '../wasm/index';
import { decodeDocxHostJson } from '../yrsSessionFacade';
import { wasmModuleMemories } from '../../wasm/loadWasmAsset';
import type { ResidentEngineWorkerPort } from '../residentEngineWorkerClient';
import type {
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '../residentEngineWorkerProtocol';

/** The real resident worker module running in-process on a real resident session. */
export interface InProcessResidentWorker extends ResidentEngineWorkerPort {
  /** Request types posted to the worker, in order. */
  readonly requests: (ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule)['type'][];
  readonly sessions: ResidentEngineSession[];
  readonly imageBlobs: Blob[];
  readonly fontFaces: Set<FontFace>;
  /** Holds the worker's replies until `release`. */
  hold(): void;
  release(): void;
}

const STUBS: Record<string, string> = {
  '../wasm/loadWasmAsset': 'export const wasmModuleMemories = () => testHarness.wasmModuleMemories();',
  './residentEngineSession':
    'export const createResidentEngineSession = (heapLimitBytes, clientId) => testHarness.createSession(heapLimitBytes, clientId);',
  './wasm/index': `
    export const preloadEditWasm = () => testHarness.preload();
    export const preloadEditWasmFrom = (source) => testHarness.preloadFrom(source);
  `,
  './yrsSessionFacade': 'export const decodeDocxHostJson = (...args) => testHarness.decodeDocxHostJson(...args);',
  '../layout/render/glyphCache': 'export class GlyphCache {}',
  '../layout/render/canvasBackend': `
    export const rasterizeDisplayPageToBackBuffer = async (buffer, page, _options, dpr = 1, zoom = 1) => {
      const scale = dpr * zoom;
      buffer.width = Math.ceil(page.width * scale);
      buffer.height = Math.ceil(page.height * scale);
      for (const primitive of page.primitives) {
        if (primitive.kind === 'image' && _options.resolveImage) await _options.resolveImage(primitive.relId);
      }
      return buffer;
    };
    export const presentOffscreenPageBackBuffer = () => {};
    export const presentOffscreenPageBackBufferWithCaret = () => {};
    export const releaseOffscreenPageCanvas = () => {};
  `,
};

class Bitmap {
  constructor(public width: number, public height: number) {}
  close() { this.width = 0; this.height = 0; }
}

class Canvas {
  constructor(public width: number, public height: number) {}
  getContext(kind: string) { return kind === '2d' ? { drawImage() {} } : null; }
  transferToImageBitmap() { return new Bitmap(this.width, this.height); }
}

/**
 * Bundles the worker module, with canvas output stubbed, and returns a factory that starts one
 * in-process worker per call. Messages cross with structured-clone semantics, asynchronously.
 * `terminate` stops the worker like a real one: no more messages either way, its timers and
 * channels close and its sessions are freed.
 */
export async function residentWorkerFactory(options: { transferBuffers?: boolean } = {}): Promise<(clientId?: number) => InProcessResidentWorker> {
  const result = await Bun.build({
    entrypoints: [resolve(import.meta.dir, '../residentEngineWorker.ts')],
    target: 'bun',
    format: 'iife',
    plugins: [
      {
        name: 'in-process-resident-worker',
        setup(build) {
          build.onLoad({ filter: /\/rustDisplayList\.ts$/ }, () => ({
            contents: `
              export const encodeDisplayListFrameExtras = (...args) => testHarness.encodeDisplayListFrameExtras(...args);
              export const loadRustDisplayListQueryEngine = () => testHarness.loadRustDisplayListQueryEngine();
              export const loadedRustDisplayListQueryEngine = () => testHarness.loadedRustDisplayListQueryEngine();
            `,
            loader: 'js',
          }));
          build.onResolve({ filter: /.*/ }, ({ path, importer }) =>
            importer.endsWith('/residentEngineWorker.ts') && path in STUBS
              ? { path, namespace: 'resident-worker-stub' }
              : undefined
          );
          build.onLoad({ filter: /.*/, namespace: 'resident-worker-stub' }, ({ path }) => ({
            contents: STUBS[path],
            loader: 'js',
          }));
        },
      },
    ],
  });
  if (!result.success) throw new AggregateError(result.logs, 'Resident worker bundle failed');
  const start = new Function(
    'self',
    'OffscreenCanvas',
    'MessageChannel',
    'setTimeout',
    'clearTimeout',
    'testHarness',
    'createImageBitmap',
    'FontFace',
    await result.outputs[0].text()
  ) as (
    scope: unknown,
    canvas: unknown,
    channel: unknown,
    schedule: unknown,
    unschedule: unknown,
    harness: unknown,
    decode: unknown,
    fontFace: unknown
  ) => void;
  return (clientId) => {
    let held: ResidentEngineWorkerResponse[] | null = null;
    let terminated = false;
    const channels: MessageChannel[] = [];
    const timers = new Set<ReturnType<typeof setTimeout>>();
    const fontFaces = new Set<FontFace>();
    const scope = {
      fonts: { add: (face: FontFace) => fontFaces.add(face), delete: (face: FontFace) => fontFaces.delete(face) },
      onmessage: null as ((event: {
        data: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule;
      }) => void) | null,
      postMessage(reply: ResidentEngineWorkerResponse) {
        if (terminated) return;
        if (held) held.push(reply);
        else deliver(reply);
      },
    };
    const deliver = (reply: ResidentEngineWorkerResponse) => {
      const data = structuredClone(reply);
      if (data.ok) {
        for (const frame of [data.commandInitial?.frame, data.commandAck?.frame, data.commandFrame?.frame]) {
          for (const page of frame?.payload.stagedPresentation?.pages ?? []) {
            for (const bitmap of [page.bitmap, ...(page.backgrounds ?? []).map((background) => background.bitmap)]) {
              Object.setPrototypeOf(bitmap, Bitmap.prototype);
            }
          }
        }
      }
      queueMicrotask(() => {
        if (!terminated) worker.onmessage?.({ data } as MessageEvent<ResidentEngineWorkerResponse>);
      });
    };
    const worker: InProcessResidentWorker = {
      onmessage: null,
      onerror: null,
      onmessageerror: null,
      requests: [],
      sessions: [],
      imageBlobs: [],
      fontFaces,
      postMessage(message, transfer = []) {
        worker.requests.push(message.type);
        if (terminated) return;
        const data = structuredClone(message, { transfer: options.transferBuffers ? transfer.filter(value => value instanceof ArrayBuffer) : [] });
        if (data.type === 'attachCanvases') {
          for (const page of data.pages) Object.setPrototypeOf(page.canvas, Canvas.prototype);
        }
        queueMicrotask(() => {
          if (!terminated) scope.onmessage?.({ data });
        });
      },
      terminate() {
        if (terminated) return;
        terminated = true;
        held = null;
        fontFaces.clear();
        scope.onmessage = null;
        for (const timer of timers) clearTimeout(timer);
        timers.clear();
        for (const channel of channels.splice(0)) {
          channel.port1.close();
          channel.port2.close();
        }
        for (const session of worker.sessions) session.destroy();
      },
      hold() {
        held ??= [];
      },
      release() {
        const replies = held ?? [];
        held = null;
        for (const reply of replies) deliver(reply);
      },
    };
    start(
      scope,
      Canvas,
      class extends MessageChannel {
        constructor() {
          super();
          channels.push(this);
        }
      },
      (callback: (...args: unknown[]) => void, delay?: number, ...args: unknown[]) => {
        const timer = setTimeout(() => {
          timers.delete(timer);
          if (!terminated) callback(...args);
        }, delay);
        if (terminated) clearTimeout(timer);
        else timers.add(timer);
        return timer;
      },
      (timer: ReturnType<typeof setTimeout>) => {
        timers.delete(timer);
        clearTimeout(timer);
      },
      {
        encodeDisplayListFrameExtras, loadRustDisplayListQueryEngine, loadedRustDisplayListQueryEngine,
        createSession: async (heapLimitBytes?: number, requested?: number) => {
          const session = await createResidentEngineSession(heapLimitBytes, clientId ?? requested);
          worker.sessions.push(session);
          if (terminated) session.destroy();
          return session;
        },
        preload: preloadEditWasm,
        preloadFrom: preloadEditWasmFrom,
        decodeDocxHostJson,
        wasmModuleMemories,
      },
      async (blob: Blob) => { worker.imageBlobs.push(blob); return new Bitmap(1, 1); },
      class {
        status = 'unloaded';
        async load() { this.status = 'loaded'; return this; }
      }
    );
    return worker;
  };
}
