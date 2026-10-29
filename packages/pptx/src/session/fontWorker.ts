import { createSessionClient } from '../../../../shared/office-session/client';
import { rpcFontDecoder, type FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';
import type { PresentationCommandMethods } from './commandProtocol';

type Methods = Pick<PresentationCommandMethods, 'decodeFontFace' | 'fontFamilies'>;

interface PreparedFontWorker {
  decoder: FontDecoder;
  families: Promise<string[]>;
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

export function preparePresentationFontDecoder(source: Uint8Array | ArrayBuffer): PreparedFontWorker {
  const existing = preparedWorkers.get(source);
  if (existing && !existing.retired()) return existing.prepared;
  const bytes = (source instanceof Uint8Array ? source : new Uint8Array(source)).slice().buffer;
  const owner = createPreparedFontWorker(bytes);
  preparedWorkers.set(source, owner);
  abandonedWorkers?.register(owner.prepared.decoder, owner.abandon, owner.token);
  return owner.prepared;
}

function createPreparedFontWorker(bytes: ArrayBuffer) {
  const worker = new Worker(new URL('./pptxSessionWorker.mjs', import.meta.url), { type: 'module' });
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
    methods: { decodeFontFace: true, fontFamilies: true },
  });
  function release(): void {
    if (!retired) { detach(); void rpc.dispose(); transport.close(); }
  }
  async function track<T>(pending: Promise<T>): Promise<T> {
    active += 1;
    try { return await pending; }
    finally { if (--active === 0 && abandoned) release(); }
  }
  const families = track(rpc.callWithTransfer('fontFamilies', [bytes], [bytes]));
  void families.catch(() => {});
  const decode = rpcFontDecoder((bytes, options, transfer) => rpc.callWithTransfer('decodeFontFace', [bytes, options], transfer));
  return {
    token,
    retired: () => retired,
    abandon() { abandoned = true; if (active === 0) release(); },
    prepared: {
      decoder: (bytes, options) => retired
        ? Promise.reject(new Error('Presentation font decoder was released')) : track(decode(bytes, options)),
      families,
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
