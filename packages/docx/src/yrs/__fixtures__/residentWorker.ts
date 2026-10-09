import { resolve } from 'node:path';
import { createResidentEngineSession, type ResidentEngineSession } from '../residentEngineSession';
import { preloadEditWasm, preloadEditWasmFrom } from '../wasm/index';
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
  /** Holds the worker's replies until `release`. */
  hold(): void;
  release(): void;
}

const STUBS: Record<string, string> = {
  './residentEngineSession':
    'export const createResidentEngineSession = (heapLimitBytes) => testHarness.createSession(heapLimitBytes);',
  './wasm/index': `
    export const preloadEditWasm = () => testHarness.preload();
    export const preloadEditWasmFrom = (source) => testHarness.preloadFrom(source);
  `,
  '../layout/render/glyphCache': 'export class GlyphCache {}',
  '../layout/render/canvasBackend': `
    export const rasterizeDisplayPageToBackBuffer = async () => {};
    export const presentOffscreenPageBackBuffer = () => {};
    export const presentOffscreenPageBackBufferWithCaret = () => {};
    export const releaseOffscreenPageCanvas = () => {};
  `,
};

/**
 * Bundles the worker module, with canvas output stubbed, and returns a factory that starts one
 * in-process worker per call. Messages cross with structured-clone semantics, asynchronously.
 * `terminate` stops the worker like a real one: no more messages either way, its timers and
 * channels close and its sessions are freed.
 */
export async function residentWorkerFactory(): Promise<(clientId?: number) => InProcessResidentWorker> {
  const result = await Bun.build({
    entrypoints: [resolve(import.meta.dir, '../residentEngineWorker.ts')],
    target: 'bun',
    format: 'iife',
    plugins: [
      {
        name: 'in-process-resident-worker',
        setup(build) {
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
    await result.outputs[0].text()
  ) as (
    scope: unknown,
    canvas: unknown,
    channel: unknown,
    schedule: unknown,
    unschedule: unknown,
    harness: unknown
  ) => void;
  return (clientId) => {
    let held: ResidentEngineWorkerResponse[] | null = null;
    let terminated = false;
    const channels: MessageChannel[] = [];
    const timers = new Set<ReturnType<typeof setTimeout>>();
    const scope = {
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
      postMessage(message) {
        worker.requests.push(message.type);
        if (terminated) return;
        const data = structuredClone(message);
        queueMicrotask(() => {
          if (!terminated) scope.onmessage?.({ data });
        });
      },
      terminate() {
        if (terminated) return;
        terminated = true;
        held = null;
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
      class {},
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
        createSession: async (heapLimitBytes?: number) => {
          const session = await createResidentEngineSession(heapLimitBytes, clientId);
          worker.sessions.push(session);
          if (terminated) session.destroy();
          return session;
        },
        preload: preloadEditWasm,
        preloadFrom: preloadEditWasmFrom,
      }
    );
    return worker;
  };
}
