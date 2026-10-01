import { createRoot } from 'react-dom/client';
import { useEffect, useRef, useState } from 'react';
import { DocxEditor, configureDefaultFonts, type DocxEditorRef } from '@betteroffice/docx-react';
import * as fonts from '../../../packages/fonts/src/index';
import '../../../packages/docx-react/src/styles/editor.css';

configureDefaultFonts({ fonts });

type Fields = Record<string, unknown>;

interface TypingProbe {
  editor: DocxEditorRef | null;
  mode: 'default' | 'worker';
  /** Requests the page posted to the resident worker, and their replies, in epoch ms. */
  posts: Fields[];
  replies: Fields[];
  /** keydown events in epoch ms, as the page saw them. */
  keys: { key: string; t: number }[];
  longTasks: { start: number; duration: number }[];
  /** The worker's own log: arrival, start of handling and reply per request id. */
  workerLog(): Promise<Fields[]>;
  /** Asks the worker for per-edit engine phase timings from now on. */
  profileEdits: boolean;
  /** The resident worker's wasm memory at the last `workerLog()`. */
  workerWasmBytes: number;
}

const origin = performance.timeOrigin;
const query = new URLSearchParams(location.search);
const probe: TypingProbe = {
  editor: null,
  mode: query.get('mode') === 'worker' ? 'worker' : 'default',
  posts: [],
  replies: [],
  keys: [],
  longTasks: [],
  profileEdits: false,
  workerWasmBytes: 0,
  workerLog: async () => [],
};
(window as unknown as { __typing: TypingProbe }).__typing = probe;

document.addEventListener(
  'keydown',
  (event) => probe.keys.push({ key: event.key, t: origin + event.timeStamp }),
  { capture: true }
);
try {
  new PerformanceObserver((list) => {
    for (const entry of list.getEntries()) {
      probe.longTasks.push({ start: origin + entry.startTime, duration: entry.duration });
    }
  }).observe({ type: 'longtask', buffered: true });
} catch {
  // longtask entries are Chromium-only
}

// Runs first inside the resident worker: timestamps each request's arrival, the
// start of its handling (the worker's handler reads `type` first) and its reply.
function workerTap() {
  const scope = self as unknown as DedicatedWorkerGlobalScope;
  const base = performance.timeOrigin;
  const now = () => base + performance.now();
  const log: Record<string, unknown>[] = [];
  const post = scope.postMessage.bind(scope);
  const memories: WebAssembly.Memory[] = [];
  const keep = <T extends (...args: never[]) => Promise<unknown>>(instantiate: T): T =>
    (async (...args: never[]) => {
      const result = (await instantiate(...args)) as { instance?: WebAssembly.Instance };
      const instance = (result.instance ?? result) as WebAssembly.Instance;
      for (const value of Object.values(instance.exports)) {
        if (value instanceof WebAssembly.Memory) memories.push(value);
      }
      return result;
    }) as T;
  WebAssembly.instantiate = keep(WebAssembly.instantiate.bind(WebAssembly)) as typeof WebAssembly.instantiate;
  WebAssembly.instantiateStreaming = keep(
    WebAssembly.instantiateStreaming.bind(WebAssembly)
  ) as typeof WebAssembly.instantiateStreaming;
  scope.addEventListener('message', (event) => {
    const data = event.data as Record<string, unknown> | null;
    if (!data || typeof data !== 'object') return;
    if (data.__typingProbe) {
      event.stopImmediatePropagation();
      post({
        __typingProbe: true,
        log: log.splice(0),
        wasmBytes: memories.reduce((sum, memory) => sum + memory.buffer.byteLength, 0),
      });
      return;
    }
    const entry: Record<string, unknown> = { id: data.id, type: data.type, arrive: now() };
    log.push(entry);
    const type = data.type;
    Object.defineProperty(data, 'type', {
      configurable: true,
      enumerable: true,
      get() {
        entry.start ??= now();
        return type;
      },
    });
  });
  scope.postMessage = ((message: Record<string, unknown> | null, transfer?: Transferable[]) => {
    if (message && typeof message === 'object' && typeof message.id === 'number') {
      log.push({ id: message.id, reply: now(), ok: message.ok });
    }
    return post(message, transfer as Transferable[]);
  }) as typeof scope.postMessage;
}

const NativeWorker = window.Worker;
const workers: Worker[] = [];
const probeWaiters = new Map<Worker, ((log: Fields[]) => void)[]>();
class TapWorker extends NativeWorker {
  constructor(url: string | URL, options?: WorkerOptions) {
    const resident = /residentEngineWorker/.test(String(url)) && options?.type === 'module';
    if (resident) {
      const tap = URL.createObjectURL(
        new Blob([`(${workerTap.toString()})();`], { type: 'text/javascript' })
      );
      const entry = new URL(String(url), location.href).href;
      url = URL.createObjectURL(
        new Blob([`import ${JSON.stringify(tap)};\nimport ${JSON.stringify(entry)};`], {
          type: 'text/javascript',
        })
      );
    }
    super(url, options);
    if (!resident) return;
    workers.push(this);
    this.addEventListener('message', (event: MessageEvent) => {
      const data = event.data as Fields | null;
      if (data?.__typingProbe) {
        event.stopImmediatePropagation();
        probe.workerWasmBytes = data.wasmBytes as number;
        probeWaiters.get(this)?.shift()?.((data as { log: Fields[] }).log);
        return;
      }
      probe.replies.push({
        t: origin + performance.now(),
        id: data?.id,
        ok: data?.ok,
        engineMs: data?.engineMs,
        workerTotalMs: data?.workerTotalMs,
        replayMs: data?.replayMs,
        replayedPages: data?.replayedPages,
        caretPainted: data?.caretPainted,
        layoutProvisional: data?.layoutProvisional,
        engineProfile: data?.engineProfile,
        frameBytes: data?.frame instanceof ArrayBuffer ? data.frame.byteLength : undefined,
        error: data?.ok === false ? String(data.error).slice(0, 200) : undefined,
        worker: workers.indexOf(this),
      });
    });
  }

  postMessage(message: unknown, transfer?: unknown): void {
    const request = message as Fields | null;
    if (
      probe.profileEdits &&
      request &&
      (request.type === 'applyInput' || request.type === 'applyDelete')
    ) {
      request.profile = true;
    }
    probe.posts.push({
      t: origin + performance.now(),
      id: request?.id,
      type: request?.type,
      worker: workers.indexOf(this),
      ...(request?.foreground === true ? { foreground: true } : {}),
      ...(typeof request?.provisionalPages === 'number' ? { provisionalPages: request.provisionalPages } : {}),
    });
    (super.postMessage as (message: unknown, transfer?: unknown) => void)(message, transfer);
  }
}
window.Worker = TapWorker as unknown as typeof Worker;

probe.workerLog = async () => {
  const logs = await Promise.all(
    workers.map(
      (worker) =>
        new Promise<Fields[]>((resolve) => {
          probeWaiters.set(worker, [...(probeWaiters.get(worker) ?? []), resolve]);
          NativeWorker.prototype.postMessage.call(worker, { __typingProbe: true });
        })
    )
  );
  return logs.flat();
};

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const editor = useRef<DocxEditorRef>(null);
  useEffect(() => {
    void fetch(query.get('doc') ?? '/__typing/document.docx')
      .then((response) => response.arrayBuffer())
      .then(setBuffer);
  }, []);
  useEffect(() => {
    probe.editor = editor.current;
  });
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={editor}
        documentBuffer={buffer}
        experimentalWorkerOpen={probe.mode === 'worker'}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
