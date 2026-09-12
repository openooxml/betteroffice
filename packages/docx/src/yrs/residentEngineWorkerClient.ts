import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerSnapshot,
  YrsSelection,
} from './index';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';

export interface ResidentEngineWorkerFrame {
  frame: Uint8Array;
  updates: Uint8Array[];
  engineMs: number;
  workerTotalMs: number;
  engineProfile?: YrsEngineApplyProfile;
  caret: YrsResidentCaretSnapshot;
  selection: YrsSelection | null;
  /** The presented frame carries the worker-painted caret line. */
  caretPainted: boolean;
  replayMs: number;
  replayedPages: number;
  layoutRevision: number;
}

export interface ResidentEngineOffscreenPage {
  pageId: string;
  canvas: OffscreenCanvas;
}

export interface ResidentEngineWorkerApplyResult extends ResidentEngineWorkerFrame {
  applied: true;
}

type PendingRequest = {
  resolve(response: ResidentEngineWorkerResponse & { ok: true }): void;
  reject(error: Error): void;
  timeout: ReturnType<typeof setTimeout>;
};

type AwaitedRequest = Exclude<
  ResidentEngineWorkerRequestWithoutId,
  { type: 'applyUpdate' | 'eraseCaret' | 'destroy' }
>;

/** attachCanvases queues behind a sync, so it shares that budget. */
const REQUEST_TIMEOUT_MS: Record<AwaitedRequest['type'], number> = {
  bootstrap: 15_000,
  sync: 15_000,
  attachCanvases: 15_000,
  buildFrame: 5_000,
  applyInput: 5_000,
  applyDelete: 5_000,
};

export interface ResidentEngineWorkerPort {
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  postMessage(message: ResidentEngineWorkerRequest, transfer?: Transferable[]): void;
  terminate(): void;
}

function spawnResidentEngineWorker(): ResidentEngineWorkerPort {
  return new Worker(new URL('./residentEngineWorker.mjs', import.meta.url), {
    type: 'module',
    name: 'openooxml-resident-engine',
  });
}

/** Dedicated-worker owner for resident input, pagination, and FrameDelta output. */
export class ResidentEngineWorkerClient {
  private readonly pending = new Map<number, PendingRequest>();
  private nextId = 1;
  private terminalError: Error | null = null;
  private ready = false;
  private revision = 0;
  private remoteVector: Uint8Array | null = null;
  private appliedFontsRevision: number | null = null;

  constructor(private readonly worker: ResidentEngineWorkerPort = spawnResidentEngineWorker()) {
    this.worker.onmessage = (event) => {
      const response = event.data;
      if (response.ok && response.stateVector) {
        this.remoteVector = new Uint8Array(response.stateVector);
      }
      if (!response.ok && response.terminal) {
        this.fail(new ResidentWorkerUnavailableError(response.error));
        return;
      }
      const pending = this.pending.get(response.id);
      if (!pending) return;
      this.pending.delete(response.id);
      clearTimeout(pending.timeout);
      if (response.ok) pending.resolve(response);
      else pending.reject(residentWorkerError(response.error, response.residentUnavailable));
    };
    this.worker.onerror = (event) => {
      this.fail(new Error(`Resident engine worker failed: ${event.message}`));
    };
    this.worker.onmessageerror = () => {
      this.fail(new Error('Resident engine worker returned an unreadable message'));
    };
  }

  isReady(): boolean {
    return this.ready;
  }

  layoutRevision(): number {
    return this.revision;
  }

  /** The worker replica's last reported yrs state vector (null before any). */
  remoteStateVector(): Uint8Array | null {
    return this.remoteVector;
  }

  /** The fonts revision this worker last applied (null before bootstrap). */
  syncedFontsRevision(): number | null {
    return this.appliedFontsRevision;
  }

  async bootstrap(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string
  ): Promise<ResidentEngineWorkerFrame> {
    const fontsRevision = snapshot.fontsRevision;
    const response = await this.request(
      { type: 'bootstrap', snapshot, extras, expectedFrameEpoch: 0 },
      snapshotTransfers(snapshot)
    );
    const result = frameResult(response);
    this.recordSync(response, fontsRevision);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async sync(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string,
    expectedFrameEpoch: number,
    paintCaret = false
  ): Promise<ResidentEngineWorkerFrame> {
    const fontsRevision = snapshot.fontsRevision;
    const response = await this.request(
      { type: 'sync', snapshot, extras, expectedFrameEpoch, paintCaret },
      snapshotTransfers(snapshot)
    );
    const result = frameResult(response);
    this.recordSync(response, fontsRevision);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async buildFrame(
    extras: string,
    expectedFrameEpoch: number,
    paintCaret = false
  ): Promise<ResidentEngineWorkerFrame> {
    const result = frameResult(
      await this.request({ type: 'buildFrame', extras, expectedFrameEpoch, paintCaret })
    );
    return result;
  }

  async applyInput(
    text: string,
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false,
    paintCaret = false
  ): Promise<ResidentEngineWorkerApplyResult | { applied: false }> {
    if (!this.ready) return { applied: false };
    try {
      const result = frameResult(
        await this.request({
          type: 'applyInput',
          text,
          selection,
          expectedFrameEpoch,
          profile,
          paintCaret,
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (error instanceof ResidentWorkerUnavailableError) return { applied: false };
      throw error;
    }
  }

  async applyDelete(
    direction: 'backward' | 'forward',
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false,
    paintCaret = false
  ): Promise<ResidentEngineWorkerApplyResult | { applied: false }> {
    if (!this.ready) return { applied: false };
    try {
      const result = frameResult(
        await this.request({
          type: 'applyDelete',
          direction,
          selection,
          expectedFrameEpoch,
          profile,
          paintCaret,
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (error instanceof ResidentWorkerUnavailableError) return { applied: false };
      throw error;
    }
  }

  /** Drop the worker-painted caret line by re-presenting the caret page's
   * retained raster. Fire-and-forget and idempotent. */
  eraseCaret(): void {
    if (this.terminalError) return;
    const id = this.nextId++;
    const message: ResidentEngineWorkerRequest = { id, type: 'eraseCaret' };
    this.worker.postMessage(message);
  }

  invalidate(update: Uint8Array, selection: YrsSelection | null): void {
    if (this.terminalError) return;
    this.ready = false;
    const owned = update.slice();
    const id = this.nextId++;
    const message: ResidentEngineWorkerRequest = {
      id,
      type: 'applyUpdate',
      update: owned,
      selection,
    };
    this.worker.postMessage(message, [owned.buffer]);
  }

  async attachCanvases(
    pages: ResidentEngineOffscreenPage[],
    activePageIds: string[],
    devicePixelRatio: number,
    zoom: number,
    caretStyle: ResidentCaretPaintStyle
  ): Promise<void> {
    const canvases = pages.map((page) => page.canvas);
    await this.request(
      { type: 'attachCanvases', pages, activePageIds, devicePixelRatio, zoom, caretStyle },
      canvases
    );
  }

  destroy(): void {
    if (this.terminalError) return;
    const id = this.nextId++;
    const message: ResidentEngineWorkerRequest = { id, type: 'destroy' };
    this.worker.postMessage(message);
    this.fail(new Error('Resident engine worker was destroyed'));
  }

  private request(
    request: AwaitedRequest,
    transfer: Transferable[] = []
  ): Promise<ResidentEngineWorkerResponse & { ok: true }> {
    if (this.terminalError) return Promise.reject(this.terminalError);
    const id = this.nextId++;
    const timeoutMs = REQUEST_TIMEOUT_MS[request.type];
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.fail(
          new Error(`Resident engine worker did not answer ${request.type} within ${timeoutMs}ms`)
        );
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timeout });
      this.worker.postMessage({ ...request, id } as ResidentEngineWorkerRequest, transfer);
    });
  }

  /** Record a successfully applied bootstrap/sync payload's fonts revision.
   * The state vector is tracked centrally in `onmessage`. */
  private recordSync(
    _response: ResidentEngineWorkerResponse & { ok: true },
    fontsRevision: number
  ): void {
    this.appliedFontsRevision = fontsRevision;
  }

  private fail(error: Error): void {
    this.terminalError = error;
    this.ready = false;
    this.worker.terminate();
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timeout);
      pending.reject(error);
    }
    this.pending.clear();
  }
}

class ResidentWorkerUnavailableError extends Error {}

function residentWorkerError(message: string, unavailable = false): Error {
  return unavailable ? new ResidentWorkerUnavailableError(message) : new Error(message);
}

function snapshotTransfers(snapshot: YrsResidentWorkerSnapshot): Transferable[] {
  return [snapshot.state.buffer, ...snapshot.fonts.map((font) => font.buffer)];
}

function frameResult(
  response: ResidentEngineWorkerResponse & { ok: true }
): ResidentEngineWorkerFrame {
  if (!response.frame) throw new Error('Resident engine worker response omitted its FrameDelta');
  if (!response.caret) throw new Error('Resident engine worker response omitted its caret snapshot');
  if (response.selection === undefined) {
    throw new Error('Resident engine worker response omitted its selection');
  }
  return {
    frame: new Uint8Array(response.frame),
    updates: (response.updates ?? []).map((update) => new Uint8Array(update)),
    engineMs: response.engineMs ?? 0,
    workerTotalMs: response.workerTotalMs ?? 0,
    engineProfile: response.engineProfile,
    caret: response.caret,
    selection: response.selection,
    caretPainted: response.caretPainted ?? false,
    replayMs: response.replayMs ?? 0,
    replayedPages: response.replayedPages ?? 0,
    layoutRevision: response.layoutRevision ?? 0,
  };
}

export function canUseResidentEngineWorker(): boolean {
  return typeof Worker !== 'undefined';
}
