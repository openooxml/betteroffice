import type { DisplayList, GridMeta, MergedRange, Viewport, WorkbookFrame } from '@betteroffice/xlsx';

export interface WorkerPaintRequest {
  generation: number;
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
}

export interface WorkerPaintScheduler {
  request(callback: () => void): number;
  cancel(id: number): void;
}

export interface WorkerPaintSourceOptions {
  generation: number;
  capture(): WorkerPaintRequest | null;
  requestFrame(request: WorkerPaintRequest): Promise<WorkbookFrame>;
  sentSequence(): number;
  publish(painted: WorkerPaintResult): void;
  onError(error: unknown): void;
  isCurrent?(): boolean;
  scheduler?: WorkerPaintScheduler;
}

interface PendingFrame {
  request: WorkerPaintRequest;
  revision: number;
}

function publicationOutcome() {
  let resolve!: () => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<void>((done, failed) => { resolve = done; reject = failed; });
  void promise.catch(() => {});
  return { promise, resolve, reject };
}

const MAX_PUBLICATION_RETRIES = 3;

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

export class WorkerPaintSource {
  private readonly scheduler: WorkerPaintScheduler;
  private disposed = false;
  private failed = false;
  private raf: number | null = null;
  private revision = 0;
  private pending: PendingFrame | null = null;
  private running = false;
  private lastPaint: WorkerPaintResult | null = null;
  private publication: ReturnType<typeof publicationOutcome> | null = null;
  private readonly editOutcomes = new Set<ReturnType<typeof publicationOutcome>>();
  private retries = 0;

  constructor(private readonly options: WorkerPaintSourceOptions) {
    this.scheduler = options.scheduler ?? animationScheduler;
  }

  get painted(): WorkerPaintResult | null { return this.lastPaint; }
  get issuedPublicationOutcome(): Promise<void> | null { return this.publication?.promise ?? null; }
  get publicationOutcome(): Promise<void> | null {
    return this.issuedPublicationOutcome ?? [...this.editOutcomes].pop()?.promise ?? null;
  }

  schedule(): void {
    this.enqueue(true);
  }

  scheduleEdit(needsPublication = true): void {
    if (!this.current || this.failed) return;
    if (needsPublication) this.editOutcomes.add(publicationOutcome());
    this.enqueue(false);
  }

  private enqueue(invalidate: boolean): void {
    if (!this.current || this.failed) return;
    if (invalidate) this.revision += 1;
    if (this.raf !== null) return;
    this.raf = this.scheduler.request(() => {
      this.raf = null;
      if (!this.current || this.failed) return;
      const request = this.capture();
      if (!request) return;
      this.pending = { request, revision: this.revision };
      void this.drain().catch((error) => this.fail(error));
    });
  }

  fail(error: unknown): void {
    if (!this.current || this.failed) return;
    this.failed = true;
    this.rejectOutcomes(error);
    this.cancelFrame();
    try { this.options.onError(error); } catch {}
  }

  dispose(error: unknown = new Error('Worker paint source disposed')): void {
    if (this.disposed) return;
    this.disposed = true;
    this.rejectOutcomes(error);
    this.cancelFrame();
  }

  private rejectOutcomes(error: unknown): void {
    this.publication?.reject(error);
    this.publication = null;
    for (const outcome of this.editOutcomes) outcome.reject(error);
    this.editOutcomes.clear();
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
    return pending.revision === this.revision && sameRequest(this.capture(), pending.request);
  }

  private cancelFrame(): void {
    if (this.raf !== null) this.scheduler.cancel(this.raf);
    this.raf = null;
    this.pending = null;
  }

  private retry(
    request: WorkerPaintRequest, error: unknown = new Error('Worker frame publication retry limit exceeded')
  ): void {
    if (!sameRequest(this.capture(), request)) this.retries = 0;
    else if (++this.retries > MAX_PUBLICATION_RETRIES) { this.fail(error); return; }
    if (!this.pending && this.raf === null && this.capture()) this.schedule();
  }

  private publish(
    source: WorkerPaintResult['source'], request: WorkerPaintRequest,
    paint: WorkerFramePaint, sequence: number
  ): WorkerPaintResult {
    const painted: WorkerPaintResult = {
      ...paint, source, request, sequence, geometry: paint.displayList.grid,
    };
    const previous = this.lastPaint;
    this.lastPaint = painted;
    try {
      this.options.publish(painted);
    } catch (error) {
      if (this.lastPaint === painted) this.lastPaint = previous;
      throw error;
    }
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
          this.retry(pending.request);
          continue;
        }
        const { request } = pending;
        const outcome = [...this.editOutcomes].pop() ?? publicationOutcome();
        this.publication = outcome;
        try {
          const frame = await this.options.requestFrame(snapshot(request));
          if (!this.current || this.failed) break;
          if (!this.matches(pending) || frame.sheet !== request.sheet ||
            !sameViewport(frame.viewport, request.viewport) ||
            !(frame.sequence >= this.options.sentSequence())) {
            this.retry(request);
            continue;
          }
          const covered = [...this.editOutcomes];
          this.publish('worker', request, {
            displayList: frame.displayList, mergedRanges: frame.mergedRanges ?? [], version: frame.version,
          }, frame.sequence);
          for (const slot of covered) {
            slot.resolve();
            this.editOutcomes.delete(slot);
          }
          this.retries = 0;
        } catch (error) {
          if (!this.current || this.failed) break;
          if (!this.matches(pending) || error instanceof Error && error.name === 'SessionSuperseded') {
            this.retry(request, error);
          } else {
            this.fail(error);
          }
        } finally {
          if (!this.editOutcomes.has(outcome)) outcome.resolve();
          if (this.publication === outcome) this.publication = null;
        }
      }
    } finally { this.running = false; }
  }
}
