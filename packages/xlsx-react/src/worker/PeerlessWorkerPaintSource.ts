import type { DisplayList, GridMeta, MergedRange, Viewport, WorkbookCommandFrame, WorkbookFrame } from '@betteroffice/xlsx';

export interface WorkerPaintRequest {
  generation: number;
  requestGeneration?: number;
  coversThroughSeq?: number;
  sheet: number;
  navigation: number;
  viewport: Viewport;
  width: number;
  height: number;
  zoom: number;
  dpr: number;
}

export interface WorkerFramePaint {
  displayList: DisplayList;
  mergedRanges: readonly MergedRange[];
  version: string;
}

export interface WorkerPaintResult extends WorkerFramePaint {
  source: 'worker';
  request: WorkerPaintRequest;
  geometry: GridMeta | undefined;
  sequence: number;
  coversThroughSeq?: number;
}

export type WorkerPaintTransaction = (painted: WorkerPaintResult, publish: () => void) => void;

export interface WorkerPaintScheduler {
  request(callback: () => void): number;
  cancel(id: number): void;
}

export interface WorkerPaintSourceOptions {
  generation: number;
  commandIdentity?: { session: string; incarnation: string };
  capture(): WorkerPaintRequest | null;
  requestFrame(request: WorkerPaintRequest): Promise<WorkbookFrame>;
  sentSequence(): number;
  publish(painted: WorkerPaintResult): void;
  onInstalled?(painted: WorkerPaintResult): void;
  onError(error: unknown): void;
  isCurrent?(): boolean;
  scheduler?: WorkerPaintScheduler;
}

interface PendingFrame {
  request: WorkerPaintRequest;
  revision: number;
  sequence: number;
  installation: number;
}

const animationScheduler: WorkerPaintScheduler = {
  request: (callback) => requestAnimationFrame(callback),
  cancel: (id) => cancelAnimationFrame(id),
};

function snapshot(request: WorkerPaintRequest): WorkerPaintRequest {
  return { ...request, viewport: { ...request.viewport } };
}

function sameViewport(left: Viewport, right: Viewport): boolean {
  return left.x === right.x && left.y === right.y &&
    left.width === right.width && left.height === right.height;
}

function sameRequest(left: WorkerPaintRequest | null, right: WorkerPaintRequest): boolean {
  return left !== null && left.generation === right.generation && left.sheet === right.sheet &&
    left.navigation === right.navigation && left.width === right.width && left.height === right.height &&
    left.zoom === right.zoom && left.dpr === right.dpr && sameViewport(left.viewport, right.viewport);
}

export class PeerlessWorkerPaintSource {
  private readonly scheduler: WorkerPaintScheduler;
  private disposed = false;
  private failed = false;
  private raf: number | null = null;
  private revision = 0;
  private pending: PendingFrame | null = null;
  private running = false;
  private lastPaint: WorkerPaintResult | null = null;
  private lastFrameSequence = 0;
  private commandCoverage = 0;
  private installation = 0;
  private readonly transactions = new Set<WorkerPaintTransaction>();

  constructor(private readonly options: WorkerPaintSourceOptions) {
    this.scheduler = options.scheduler ?? animationScheduler;
  }

  get painted(): WorkerPaintResult | null { return this.lastPaint; }
  get coversThroughSeq(): number { return this.commandCoverage; }

  registerPaintTransaction(transaction: WorkerPaintTransaction): () => void {
    if (this.disposed) return () => {};
    this.transactions.add(transaction);
    return () => { this.transactions.delete(transaction); };
  }

  captureCommandFrame(): WorkerPaintRequest | null {
    const request = this.failed ? null : this.capture();
    return request ? { ...request, requestGeneration: this.revision,
      coversThroughSeq: this.options.sentSequence() } : null;
  }

  installInitial(frame: WorkbookCommandFrame | undefined): boolean {
    const current = this.capture();
    if (!frame || !current || this.failed || this.lastPaint || this.running || this.pending ||
      this.raf !== null || this.revision !== 0 || frame.coversThroughSeq !== 0 || frame.sequence !== 0 ||
      !Number.isSafeInteger(frame.requestGeneration) ||
      frame.requestGeneration < 0 || frame.viewport.width !== current.width / frame.zoom ||
      frame.viewport.height !== current.height / frame.zoom ||
      (current.requestGeneration !== undefined && frame.requestGeneration !== current.requestGeneration)) return false;
    const request: WorkerPaintRequest = { ...current, sheet: frame.sheet,
      viewport: { ...frame.viewport }, zoom: frame.zoom, dpr: frame.dpr,
      requestGeneration: frame.requestGeneration };
    this.revision = frame.requestGeneration;
    if (this.install(frame, request)) return true;
    this.revision = 0;
    return false;
  }

  install(frame: WorkbookCommandFrame | undefined, request: WorkerPaintRequest): boolean {
    const identity = this.options.commandIdentity;
    if (!frame || !identity || !this.current || this.failed ||
      frame.session !== identity.session || frame.incarnation !== identity.incarnation ||
      frame.requestGeneration !== (request.requestGeneration ?? request.generation) ||
      (request.requestGeneration !== undefined && request.requestGeneration !== this.revision) ||
      frame.zoom !== request.zoom || frame.dpr !== request.dpr ||
      frame.sheet !== request.sheet || !sameViewport(frame.viewport, request.viewport) ||
      !sameRequest(this.capture(), request) || !Number.isSafeInteger(frame.coversThroughSeq) ||
      !Number.isSafeInteger(frame.sequence) || frame.sequence < 0 ||
      frame.coversThroughSeq < Math.max(this.commandCoverage, request.coversThroughSeq ?? 0)) return false;
    try {
      this.publish('worker', snapshot(request), {
        displayList: frame.displayList, mergedRanges: frame.mergedRanges ?? [], version: frame.version,
      }, frame.sequence, frame.coversThroughSeq);
      this.installation += 1;
      return true;
    } catch (error) {
      this.fail(error);
      return false;
    }
  }

  schedule(): void {
    if (!this.current || this.failed) return;
    this.revision += 1;
    if (this.raf !== null) return;
    this.raf = this.scheduler.request(() => {
      this.raf = null;
      if (!this.current || this.failed) return;
      const request = this.capture();
      if (!request) return;
      this.pending = { request, revision: this.revision, sequence: this.options.sentSequence(),
        installation: this.installation };
      void this.drain().catch((error) => this.fail(error));
    });
  }

  fail(error: unknown): void {
    if (!this.current || this.failed) return;
    this.failed = true;
    this.cancelFrame();
    try { this.options.onError(error); } catch {}
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.cancelFrame();
    this.transactions.clear();
  }

  private get current(): boolean {
    return !this.disposed && (this.options.isCurrent?.() ?? true);
  }

  private capture(): WorkerPaintRequest | null {
    if (!this.current) return null;
    const request = this.options.capture();
    return request?.generation === this.options.generation ? snapshot(request) : null;
  }

  private matches(pending: PendingFrame): boolean {
    return pending.revision === this.revision && pending.installation === this.installation &&
      sameRequest(this.capture(), pending.request);
  }

  private cancelFrame(): void {
    if (this.raf !== null) this.scheduler.cancel(this.raf);
    this.raf = null;
    this.pending = null;
  }

  private retry(pending: PendingFrame): void {
    const request = this.capture();
    if (!request || sameRequest(request, pending.request) &&
      this.options.sentSequence() === pending.sequence) return;
    if (!this.pending && this.raf === null) this.schedule();
  }

  private publish(
    source: WorkerPaintResult['source'], request: WorkerPaintRequest,
    paint: WorkerFramePaint, sequence: number, coversThroughSeq?: number
  ): WorkerPaintResult {
    const painted: WorkerPaintResult = {
      ...paint, source, request, sequence, geometry: paint.displayList.grid,
      ...(coversThroughSeq === undefined ? {} : { coversThroughSeq }),
    };
    const previous = this.lastPaint;
    const previousSequence = this.lastFrameSequence;
    const previousCoverage = this.commandCoverage;
    this.lastPaint = painted;
    this.lastFrameSequence = Math.max(this.lastFrameSequence, sequence);
    if (coversThroughSeq !== undefined) this.commandCoverage = coversThroughSeq;
    try {
      let publish = () => { this.options.publish(painted); };
      for (const transaction of [...this.transactions].reverse()) {
        const next = publish;
        publish = () => { transaction(painted, next); };
      }
      publish();
    } catch (error) {
      if (this.lastPaint === painted) {
        this.lastPaint = previous;
        this.lastFrameSequence = previousSequence;
        this.commandCoverage = previousCoverage;
      }
      throw error;
    }
    this.options.onInstalled?.(painted);
    return painted;
  }

  private async drain(): Promise<void> {
    if (this.running || !this.current || this.failed) return;
    this.running = true;
    try {
      while (this.current && !this.failed && this.pending) {
        const pending = this.pending;
        this.pending = null;
        if (!this.matches(pending)) {
          this.retry(pending);
          continue;
        }
        const { request } = pending;
        let frame: WorkbookFrame;
        try {
          frame = await this.options.requestFrame(snapshot(request));
        } catch (error) {
          if (!this.current || this.failed) break;
          if (pending.installation !== this.installation || !sameRequest(this.capture(), request) ||
            error instanceof Error && error.name === 'SessionSuperseded') {
            this.retry(pending);
          } else {
            this.fail(error);
          }
          continue;
        }
        if (!this.current || this.failed) break;
        const coverage = 'coversThroughSeq' in frame && typeof frame.coversThroughSeq === 'number'
          ? frame.coversThroughSeq : undefined;
        if (!this.matches(pending) || frame.sheet !== request.sheet ||
          !sameViewport(frame.viewport, request.viewport) ||
          (request.coversThroughSeq === undefined
            ? !(frame.sequence >= this.options.sentSequence()) || frame.sequence < this.lastFrameSequence
            : coverage === undefined || coverage < Math.max(request.coversThroughSeq, this.commandCoverage))) {
          this.retry(pending);
          continue;
        }
        try {
          this.publish('worker', request, {
            displayList: frame.displayList, mergedRanges: frame.mergedRanges ?? [], version: frame.version,
          }, frame.sequence, coverage);
        } catch (error) { this.fail(error); }
      }
    } finally { this.running = false; }
  }
}
