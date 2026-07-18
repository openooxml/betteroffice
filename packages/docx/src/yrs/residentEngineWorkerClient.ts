import type { YrsEngineApplyProfile, YrsResidentWorkerSnapshot, YrsSelection } from './index';
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
  timeout: ReturnType<typeof setTimeout> | null;
};

const RESIDENT_ENGINE_WORKER_STARTUP_TIMEOUT_MS = 15_000;

/** Dedicated-worker owner for resident input, pagination, and FrameDelta output. */
export class ResidentEngineWorkerClient {
  private readonly worker: Worker;
  private readonly pending = new Map<number, PendingRequest>();
  private nextId = 1;
  private destroyed = false;
  private ready = false;
  private revision = 0;

  constructor() {
    this.worker = new Worker(new URL('./residentEngineWorker.mjs', import.meta.url), {
      type: 'module',
      name: 'openooxml-resident-engine',
    });
    this.worker.onmessage = (event: MessageEvent<ResidentEngineWorkerResponse>) => {
      const response = event.data;
      const pending = this.pending.get(response.id);
      if (!pending) return;
      this.pending.delete(response.id);
      if (pending.timeout) clearTimeout(pending.timeout);
      if (response.ok) pending.resolve(response);
      else pending.reject(residentWorkerError(response.error, response.residentUnavailable));
    };
    this.worker.onerror = (event) => {
      this.failAll(new Error(`Resident engine worker failed: ${event.message}`));
      this.ready = false;
    };
    this.worker.onmessageerror = () => {
      this.failAll(new Error('Resident engine worker returned an unreadable message'));
      this.ready = false;
    };
  }

  isReady(): boolean {
    return this.ready;
  }

  layoutRevision(): number {
    return this.revision;
  }

  async bootstrap(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string
  ): Promise<ResidentEngineWorkerFrame> {
    const response = await this.request(
      {
        type: 'bootstrap',
        snapshot,
        extras,
        expectedFrameEpoch: 0,
      },
      snapshotTransfers(snapshot),
      RESIDENT_ENGINE_WORKER_STARTUP_TIMEOUT_MS
    );
    const result = frameResult(response);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async sync(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string,
    expectedFrameEpoch: number
  ): Promise<ResidentEngineWorkerFrame> {
    const response = await this.request(
      { type: 'sync', snapshot, extras, expectedFrameEpoch },
      snapshotTransfers(snapshot)
    );
    const result = frameResult(response);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  async buildFrame(extras: string, expectedFrameEpoch: number): Promise<ResidentEngineWorkerFrame> {
    const result = frameResult(
      await this.request({ type: 'buildFrame', extras, expectedFrameEpoch })
    );
    return result;
  }

  async applyInput(
    text: string,
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false
  ): Promise<ResidentEngineWorkerApplyResult | { applied: false }> {
    if (!this.ready) return { applied: false };
    try {
      const result = frameResult(
        await this.request({ type: 'applyInput', text, selection, expectedFrameEpoch, profile })
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
    profile = false
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
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (error instanceof ResidentWorkerUnavailableError) return { applied: false };
      throw error;
    }
  }

  invalidate(update: Uint8Array, selection: YrsSelection | null): void {
    if (this.destroyed) return;
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
    zoom: number
  ): Promise<void> {
    const canvases = pages.map((page) => page.canvas);
    await this.request(
      { type: 'attachCanvases', pages, activePageIds, devicePixelRatio, zoom },
      canvases
    );
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    const id = this.nextId++;
    const message: ResidentEngineWorkerRequest = { id, type: 'destroy' };
    this.worker.postMessage(message);
    this.worker.terminate();
    this.failAll(new Error('Resident engine worker was destroyed'));
    this.ready = false;
  }

  private request(
    request: ResidentEngineWorkerRequestWithoutId,
    transfer: Transferable[] = [],
    timeoutMs?: number
  ): Promise<ResidentEngineWorkerResponse & { ok: true }> {
    if (this.destroyed) return Promise.reject(new Error('Resident engine worker was destroyed'));
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timeout = timeoutMs
        ? setTimeout(() => {
            const pending = this.pending.get(id);
            if (!pending) return;
            this.pending.delete(id);
            const error = new Error(
              `Resident engine worker did not acknowledge ${request.type} within ${timeoutMs}ms`
            );
            pending.reject(error);
            this.failAll(error);
            this.ready = false;
            this.destroyed = true;
            this.worker.terminate();
          }, timeoutMs)
        : null;
      this.pending.set(id, { resolve, reject, timeout });
      this.worker.postMessage({ ...request, id } as ResidentEngineWorkerRequest, transfer);
    });
  }

  private failAll(error: Error): void {
    for (const pending of this.pending.values()) {
      if (pending.timeout) clearTimeout(pending.timeout);
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
  return {
    frame: new Uint8Array(response.frame),
    updates: (response.updates ?? []).map((update) => new Uint8Array(update)),
    engineMs: response.engineMs ?? 0,
    workerTotalMs: response.workerTotalMs ?? 0,
    engineProfile: response.engineProfile,
    replayMs: response.replayMs ?? 0,
    replayedPages: response.replayedPages ?? 0,
    layoutRevision: response.layoutRevision ?? 0,
  };
}

export function canUseResidentEngineWorker(): boolean {
  return typeof Worker !== 'undefined';
}
