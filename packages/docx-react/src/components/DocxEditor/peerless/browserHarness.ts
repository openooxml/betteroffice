import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import type { DocxCommandFrame, ResidentEngineWorkerResponse } from '@betteroffice/docx/yrs';
import type { PeerlessEchoIntent } from './echoPainter';
import type { PeerlessFrameGeometry } from './viewStore';

type Mode = 'editing' | 'viewing';
type Stage = 'installed' | 'reconciled' | 'microtask' | 'raf';
type Probe = (stage: 'installed' | 'reconciled', frame: DocxCommandFrame,
  retired: readonly EchoEntry<PeerlessEchoIntent>[], remaining: readonly EchoEntry<PeerlessEchoIntent>[],
  canvases: readonly HTMLCanvasElement[]) => void;
type Checkpoint = {
  stage: Stage;
  transaction: number;
  revision: number;
  coversThroughSeq: number;
  retired: number[];
  remaining: number[];
  pixels: string[];
};
type BrowserHarness = {
  readonly capabilities: { presentation: boolean; mode: boolean; export: boolean; viewport: boolean };
  checkpoints: Checkpoint[];
  presentationProbe: Probe;
  armPresentation(coversThroughSeq: number): void;
  setMode(mode: Mode): void;
  registerWorker(worker: Worker): void;
  exportBytes(): Promise<number[]>;
  viewportProbe(geometry: PeerlessFrameGeometry): void;
  setViewport(pages: number[]): Promise<void>;
  tableAction(action: unknown): Promise<unknown>;
};

let tableActionExecutor: ((action: unknown) => Promise<unknown>) | undefined;

export function registerTableActionExecutor(executor: (action: unknown) => Promise<unknown>): () => void {
  if (!(globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean }).__DOCX_PEERLESS_HARNESS__) return () => {};
  tableActionExecutor = executor;
  return () => { if (tableActionExecutor === executor) tableActionExecutor = undefined; };
}

let active: Probe | undefined;
let exportRequest = 0;

export const presentationProbe: Probe = (...args) => active?.(...args);

export function registerBrowserHarness(controls: { setMode(mode: Mode): void }): () => void {
  const scope = globalThis as typeof globalThis & {
    __DOCX_PEERLESS_HARNESS__?: boolean;
    __docxPeerlessHarness?: BrowserHarness;
  };
  if (!(globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean }).__DOCX_PEERLESS_HARNESS__) return () => {};
  let worker: Worker | undefined;
  let coverage: number | undefined;
  let transaction = 0;
  let generation = 0;
  let installed: Checkpoint | undefined;
  let frameEpoch: number | undefined;
  let viewport: readonly number[] | undefined;
  const checkpoints: Checkpoint[] = [];
  const pixels = (canvases: readonly HTMLCanvasElement[]) => {
    try { return canvases.map((canvas) => canvas.toDataURL()); }
    catch { return []; }
  };
  const sequences = (entries: readonly EchoEntry<PeerlessEchoIntent>[]) => entries.map((entry) =>
    entry.seq ?? entry.ticket.seq
  ).filter((seq): seq is number => seq !== undefined);
  const capture: Probe = (stage, frame, retired, remaining, canvases) => {
    if (frame.coversThroughSeq !== coverage) return;
    if (stage === 'installed') {
      if (installed) return;
      installed = { stage, transaction: ++transaction, revision: frame.revision,
        coversThroughSeq: frame.coversThroughSeq, retired: sequences(retired), remaining: sequences(remaining),
        pixels: pixels(canvases) };
      checkpoints.push(installed);
      return;
    }
    if (!installed || installed.revision !== frame.revision) return;
    const sample = installed;
    const current = generation;
    coverage = undefined;
    checkpoints.push({ ...sample, stage, pixels: pixels(canvases) });
    queueMicrotask(() => {
      if (active !== capture || generation !== current) return;
      checkpoints.push({ ...sample, stage: 'microtask', pixels: pixels(canvases) });
      requestAnimationFrame(() => {
        if (active !== capture || generation !== current) return;
        checkpoints.push({ ...sample, stage: 'raf', pixels: pixels(canvases) });
      });
    });
  };
  const api: BrowserHarness = {
    get capabilities() { return { presentation: true, mode: true, export: worker !== undefined,
      viewport: worker !== undefined && frameEpoch !== undefined }; },
    checkpoints,
    presentationProbe,
    armPresentation(coversThroughSeq) {
      generation += 1;
      coverage = coversThroughSeq;
      installed = undefined;
      checkpoints.length = 0;
    },
    setMode: controls.setMode,
    tableAction(action) {
      if (!tableActionExecutor) return Promise.reject(new Error('DOCX harness table actions are not registered'));
      return tableActionExecutor(action);
    },
    registerWorker(value) { worker = value; },
    viewportProbe(geometry) {
      frameEpoch = geometry.frameEpoch;
      const visible = viewport;
      if (visible) geometry.pages = geometry.pages.filter(({ page }) => visible.includes(page.pageIndex));
    },
    setViewport(pages) {
      const target = worker;
      const expectedFrameEpoch = frameEpoch;
      if (!target || expectedFrameEpoch === undefined) return Promise.reject(new Error('DOCX harness viewport is not registered'));
      if (!pages.length || pages.some((page) => !Number.isSafeInteger(page) || page < 0)) {
        return Promise.reject(new Error('DOCX harness viewport requires page indices'));
      }
      const requested = [...new Set(pages)];
      viewport = requested;
      const id = --exportRequest;
      return new Promise<void>((resolve, reject) => {
        const finish = () => {
          clearTimeout(timeout);
          target.removeEventListener('message', receive);
        };
        const receive = (event: MessageEvent<ResidentEngineWorkerResponse>) => {
          const reply = event.data;
          if (reply.id !== id) return;
          finish();
          if (!reply.ok) reject(new Error(reply.error));
          else resolve();
        };
        const timeout = setTimeout(() => {
          finish();
          reject(new Error('DOCX harness viewport timed out'));
        }, 15_000);
        target.addEventListener('message', receive);
        try { target.postMessage({ id, type: 'buildPages', pages: requested, expectedFrameEpoch, paintCaret: false }); }
        catch (error) { finish(); reject(error); }
      });
    },
    exportBytes() {
      const target = worker;
      if (!target) return Promise.reject(new Error('DOCX harness worker is not registered'));
      const id = --exportRequest;
      return new Promise<number[]>((resolve, reject) => {
        const finish = () => {
          clearTimeout(timeout);
          target.removeEventListener('message', receive);
        };
        const receive = (event: MessageEvent<ResidentEngineWorkerResponse>) => {
          const reply = event.data;
          if (reply.id !== id) return;
          finish();
          if (!reply.ok) reject(new Error(reply.error));
          else if (!reply.saved) reject(new Error('DOCX harness export returned no bytes'));
          else resolve(Array.from(new Uint8Array(reply.saved)));
        };
        const timeout = setTimeout(() => {
          finish();
          reject(new Error('DOCX harness export timed out'));
        }, 15_000);
        target.addEventListener('message', receive);
        try { target.postMessage({ id, type: 'save', comments: [] }); }
        catch (error) { finish(); reject(error); }
      });
    },
  };
  active = capture;
  scope.__docxPeerlessHarness = api;
  return () => {
    generation += 1;
    if (active === capture) active = undefined;
    if (scope.__docxPeerlessHarness === api) delete scope.__docxPeerlessHarness;
  };
}
