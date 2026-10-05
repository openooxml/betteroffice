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
  sequence: number;
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

export class WorkerPaintSource {
  private readonly scheduler: WorkerPaintScheduler;
  private disposed = false;
  private failed = false;
  private raf: number | null = null;
  private revision = 0;
  private pending: PendingFrame | null = null;
  private running = false;
  private lastPaint: WorkerPaintResult | null = null;

  constructor(private readonly options: WorkerPaintSourceOptions) {
    this.scheduler = options.scheduler ?? animationScheduler;
  }

  get painted(): WorkerPaintResult | null { return this.lastPaint; }

  schedule(): void {
    if (!this.current || this.failed) return;
    this.revision += 1;
    if (this.raf !== null) return;
    this.raf = this.scheduler.request(() => {
      this.raf = null;
      if (!this.current || this.failed) return;
      const request = this.capture();
      if (!request) return;
      this.pending = { request, revision: this.revision, sequence: this.options.sentSequence() };
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

  private retry(pending: PendingFrame): void {
    const request = this.capture();
    if (!request || sameRequest(request, pending.request) &&
      this.options.sentSequence() === pending.sequence) return;
    if (!this.pending && this.raf === null) this.schedule();
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
          this.retry(pending);
          continue;
        }
        const { request } = pending;
        let frame: WorkbookFrame;
        try {
          frame = await this.options.requestFrame(snapshot(request));
        } catch (error) {
          if (!this.current || this.failed) break;
          if (!sameRequest(this.capture(), request) || error instanceof Error && error.name === 'SessionSuperseded') {
            this.retry(pending);
          } else {
            this.fail(error);
          }
          continue;
        }
        if (!this.current || this.failed) break;
        if (!this.matches(pending) || frame.sheet !== request.sheet ||
          !sameViewport(frame.viewport, request.viewport) ||
          !(frame.sequence >= this.options.sentSequence())) {
          this.retry(pending);
          continue;
        }
        try {
          this.publish('worker', request, {
            displayList: frame.displayList, mergedRanges: frame.mergedRanges ?? [], version: frame.version,
          }, frame.sequence);
        } catch (error) { this.fail(error); }
      }
    } finally { this.running = false; }
  }
}
