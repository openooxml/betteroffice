import { getFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { shouldSlicePresentationFonts } from '../../../../shared/office-session/render/pptxSlicedFonts';
import { createSessionClient, requestWasmCompile } from '../../../../shared/office-session/client';
import { rpcFontDecoder, type FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import { createSessionWorker } from './sessionWorker';
import { shipsWasmAsset } from '../wasm/asset';
import type { PresentationCommandMethods } from './commandProtocol';

type Methods = Pick<PresentationCommandMethods, 'decodeFontFace'>;
const UNTRACKED_CALL_ID = 2 ** 40;

interface PreparedFontWorker {
  decoder: FontDecoder;
  release(): void;
  take(): Worker;
}

type FontMode = 'disabled' | 'whole' | 'sliced';
const fontMode = ({ defaultFonts, slicedFonts }: { defaultFonts?: boolean; slicedFonts?: boolean }): FontMode =>
  defaultFonts === false ? 'disabled' : shouldSlicePresentationFonts(slicedFonts) ? 'sliced' : 'whole';
const preparedWorkers = new WeakMap<object, Map<FontMode, ReturnType<typeof createPreparedFontWorker>>>();
const Registry = (globalThis as unknown as {
  FinalizationRegistry?: new (cleanup: (release: () => void) => void) => {
    register(target: object, release: () => void, token: object): void;
    unregister(token: object): boolean;
  };
}).FinalizationRegistry;
const abandonedWorkers = Registry ? new Registry(release => release()) : undefined;

/**
 * Starts the session worker for `source` early and returns a handle to its font decoder; the editor that opens the
 * same `source` object with the same font mode adopts the worker. Each shared handle is
 * released on its own, and the worker stops when the last one is released. Unless `defaultFonts` is false, the
 * worker also starts fetching the CJK coverage faces the opening slide needs. Pass the `defaultFonts`,
 * `activeSlideIndex` and `slicedFonts` the editor will use.
 */
export function preparePresentationFontDecoder(source: Uint8Array | ArrayBuffer,
  options: { defaultFonts?: boolean; activeSlideIndex?: number; slicedFonts?: boolean } = {}): PreparedFontWorker {
  const mode = fontMode(options);
  let workers = preparedWorkers.get(source);
  if (!workers) preparedWorkers.set(source, workers = new Map());
  let owner = workers.get(mode);
  if (!owner || owner.retired()) {
    owner = createPreparedFontWorker();
    workers.set(mode, owner);
    if (mode !== 'disabled') {
      owner.prefetchCoverage(source, options.activeSlideIndex ?? 0, mode === 'sliced');
    }
  }
  return owner.handle();
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
  let holders = 0;
  const tokens = new Set<object>();
  function detach(): void {
    retired = true;
    for (const token of tokens) abandonedWorkers?.unregister(token);
    tokens.clear();
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
    finally { if (--active === 0 && abandoned && holders === 0) release(); }
  }
  const decode = rpcFontDecoder((bytes, options, transfer) => rpc.callWithTransfer('decodeFontFace', [bytes, options], transfer));
  const decoder: FontDecoder = (bytes, options) => retired
    ? Promise.reject(new Error('Presentation font decoder was released')) : track(decode(bytes, options));
  function take(): Worker {
    if (retired) throw new Error('Presentation font decoder was released');
    detach();
    void rpc.dispose();
    return worker;
  }
  return {
    retired: () => retired,
    shared: () => holders > 1,
    take,
    prefetchCoverage(source: Uint8Array | ArrayBuffer, activeSlideIndex: number, sliced: boolean): void {
      const fontCdn = getFontCdnConfiguration();
      void import('./coverageScan').then((scan) => scan.activeSlideCoverageScripts(source, activeSlideIndex)).then((scripts) => {
        const args = fontCdn ? [scripts, sliced, fontCdn] : [scripts, sliced];
        if (scripts.length) connection.post({ protocol: 1, kind: 'call', id: UNTRACKED_CALL_ID, method: 'prefetchCoverageFonts', args }, []);
      }).catch(() => {});
    },
    /** A holder's own handle: the worker stops once every handle is released, or collected after its decodes. */
    handle(): PreparedFontWorker {
      holders += 1;
      abandoned = false;
      let held = true;
      const token = {};
      const drop = (collected: boolean) => {
        if (!held) return;
        held = false;
        tokens.delete(token);
        abandonedWorkers?.unregister(token);
        if (--holders > 0) return;
        if (collected) { abandoned = true; if (active === 0) release(); } else release();
      };
      const prepared: PreparedFontWorker = {
        decoder: (bytes, options) => held ? decoder(bytes, options)
          : Promise.reject(new Error('Presentation font decoder was released')),
        release: () => drop(false),
        take() {
          if (!held) throw new Error('Presentation font decoder was released');
          if (holders > 1) throw new Error('Presentation font decoder is shared with another editor');
          held = false;
          holders -= 1;
          return take();
        },
      };
      tokens.add(token);
      abandonedWorkers?.register(prepared.decoder, () => drop(true), token);
      return prepared;
    },
  };
}

/** The prepared worker for `source`, unless another editor still holds it; that editor's fonts keep loading. */
export function takePresentationFontWorker(source: Uint8Array | ArrayBuffer,
  options: { defaultFonts?: boolean; slicedFonts?: boolean } = {}): Worker | undefined {
  const workers = preparedWorkers.get(source);
  const mode = fontMode(options);
  const owner = workers?.get(mode);
  if (owner?.shared() && !owner.retired()) return undefined;
  workers?.delete(mode);
  return owner && !owner.retired() ? owner.take() : undefined;
}
