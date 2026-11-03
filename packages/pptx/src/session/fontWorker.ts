import { createSessionClient, requestWasmCompile } from '../../../../shared/office-session/client';
import { rpcFontDecoder, type FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { createSessionWorker } from '#session-worker';
import { shipsWasmAsset } from '../wasm/asset';
import type { PresentationCommandMethods } from './commandProtocol';

type Methods = Pick<PresentationCommandMethods, 'decodeFontFace'>;
const UNTRACKED_CALL_ID = 2 ** 40;

interface PreparedFontWorker {
  decoder: FontDecoder;
  release(): void;
  take(): Worker;
}

const preparedWorkers = new WeakMap<object, ReturnType<typeof createPreparedFontWorker>>();
const Registry = (globalThis as unknown as {
  FinalizationRegistry?: new (cleanup: (release: () => void) => void) => {
    register(target: object, release: () => void, token: object): void;
    unregister(token: object): boolean;
  };
}).FinalizationRegistry;
const abandonedWorkers = Registry ? new Registry(release => release()) : undefined;

/**
 * Starts the session worker for `source` early and returns its font decoder; the editor that opens the same
 * `source` object adopts the worker. Unless `defaultFonts` is false, the worker also starts fetching the CJK
 * coverage faces the opening slide needs. Pass the `defaultFonts` and `activeSlideIndex` the editor will use.
 */
export function preparePresentationFontDecoder(source: Uint8Array | ArrayBuffer,
  options: { defaultFonts?: boolean; activeSlideIndex?: number } = {}): PreparedFontWorker {
  const existing = preparedWorkers.get(source);
  if (existing && !existing.retired()) return existing.prepared;
  const owner = createPreparedFontWorker();
  preparedWorkers.set(source, owner);
  abandonedWorkers?.register(owner.prepared.decoder, owner.abandon, owner.token);
  if (options.defaultFonts !== false) owner.prefetchCoverage(source, options.activeSlideIndex ?? 0);
  return owner.prepared;
}

function createPreparedFontWorker() {
  const worker = createSessionWorker();
  const transport = createWorkerTransport(worker);
  const subscriptions = new Set<() => void>();
  const subscribe = (stop: () => void) => {
    subscriptions.add(stop);
    return () => { subscriptions.delete(stop); stop(); };
  };
  let retired = false;
  let active = 0;
  let abandoned = false;
  const token = {};
  function detach(): void {
    retired = true;
    abandonedWorkers?.unregister(token);
    for (const stop of subscriptions) stop();
    subscriptions.clear();
  }
  const connection: SessionTransport = { ...transport,
    post(message, transfer) { if (!retired) transport.post(message, transfer); },
    listen: listener => subscribe(transport.listen(listener)),
    onError: listener => subscribe(transport.onError(listener)),
    close() { if (!retired) { detach(); transport.close(); } },
  };
  const rpc = createSessionClient<Methods, {}>(connection, {
    methods: { decodeFontFace: true },
  });
  if (shipsWasmAsset) requestWasmCompile(connection);
  function release(): void {
    if (!retired) { detach(); void rpc.dispose(); transport.close(); }
  }
  async function track<T>(pending: Promise<T>): Promise<T> {
    active += 1;
    try { return await pending; }
    finally { if (--active === 0 && abandoned) release(); }
  }
  const decode = rpcFontDecoder((bytes, options, transfer) => rpc.callWithTransfer('decodeFontFace', [bytes, options], transfer));
  return {
    token,
    retired: () => retired,
    prefetchCoverage(source: Uint8Array | ArrayBuffer, activeSlideIndex: number): void {
      void import('./coverageScan').then((scan) => scan.activeSlideCoverageScripts(source, activeSlideIndex)).then((scripts) => {
        if (scripts.length) connection.post({ protocol: 1, kind: 'call', id: UNTRACKED_CALL_ID, method: 'prefetchCoverageFonts', args: [scripts] }, []);
      }).catch(() => {});
    },
    abandon() { abandoned = true; if (active === 0) release(); },
    prepared: {
      decoder: (bytes, options) => retired
        ? Promise.reject(new Error('Presentation font decoder was released')) : track(decode(bytes, options)),
      release,
      take() {
        if (retired) throw new Error('Presentation font decoder was released');
        detach();
        void rpc.dispose();
        return worker;
      },
    } satisfies PreparedFontWorker,
  };
}

export function takePresentationFontWorker(source: Uint8Array | ArrayBuffer): Worker | undefined {
  const owner = preparedWorkers.get(source);
  preparedWorkers.delete(source);
  return owner && !owner.retired() ? owner.prepared.take() : undefined;
}
