import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerSnapshot,
  YrsSelection,
} from './index';
import type { CollaborationCursor } from '../collaboration/types';
import type { ResidentCaretPaintStyle } from './residentCaret';
import type {
  ResidentDocumentRead,
  ResidentDocumentReadValues,
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerRequestWithoutId,
  ResidentEngineWorkerResponse,
  ResidentProposalOperation,
  ResidentProposalResponse,
} from './residentEngineWorkerProtocol';
import type { DocxProposalRegistryState } from './proposals';
import type { WasmModuleMemory } from '../wasm/loadWasmAsset';
import { editWasmModule } from './wasm/index';

/** @internal */
export interface ResidentProposalReply
  extends Omit<ResidentProposalResponse, 'updates' | 'stateVector'> {
  updates: Uint8Array[];
  stateVector: Uint8Array;
}

export interface ResidentEngineWorkerFrame {
  frame: Uint8Array;
  pageFrames?: Uint8Array[];
  updates: Uint8Array[];
  engineMs: number;
  workerTotalMs: number;
  engineProfile?: YrsEngineApplyProfile;
  caret: YrsResidentCaretSnapshot;
  selection: YrsSelection | null;
  /** The same selection as sticky positions, for the host to resolve against its content. */
  selectionCursor?: CollaborationCursor | null;
  /** The presented frame carries the worker-painted caret line. */
  caretPainted: boolean;
  replayMs: number;
  replayedPages: number;
  layoutRevision: number;
  /** The worker document version the frame lays out. */
  documentVersion?: string;
  documentPreview?: boolean;
  documentAsOpened?: boolean;
  /** Characters an applyDelete removed. */
  deletedUnits: number;
  /** The region layout the worker ran, when the request handed it the layout. */
  layoutJson?: string;
  /** `layoutJson` covers only the first pages; `completeLayout` finishes it. */
  layoutProvisional?: boolean;
}

/** A bootstrap/sync whose snapshot layout the worker runs as the only layout. */
export interface ResidentEngineWorkerLayoutOptions {
  /** Display extras minus the header/footer payload the worker's layout supplies. */
  layoutExtras?: string;
  /** The host state vector the snapshot brings the worker to. */
  stateVector?: Uint8Array;
  /** Lay out just the body's first pages before replying. */
  provisionalPages?: number;
  /** Bootstrap only: lay out the document {@link ResidentEngineWorkerClient.open} opened. */
  opened?: boolean;
  /** Bootstrap only: the epoch of the frame the host shows; the worker's frames follow it. */
  frameEpoch?: number;
  /**
   * Bootstrap only: the most the worker's editing core may allocate at once.
   * An allocation past it stops the worker with {@link ResidentWorkerOutOfMemoryError}.
   */
  heapLimitBytes?: number;
}

/** What an {@link ResidentEngineWorkerClient.open} seeded in the worker. */
export interface ResidentEngineWorkerOpened {
  /** The package's host metadata JSON, for `decodeDocxHostJson`. */
  hostJson: string;
  /** The worker replica's state vector once seeded. */
  stateVector: Uint8Array;
}

/** How a bootstrap or sync builds its frame. */
export interface ResidentEngineWorkerSnapshotOptions {
  /** Pages `[start, end)` the frame builds; the rest stay unbuilt. */
  displayWindow?: [number, number];
}

export interface ResidentEngineOffscreenPage {
  pageId: string;
  canvas: OffscreenCanvas;
}

export interface ResidentEngineWorkerApplyResult extends ResidentEngineWorkerFrame {
  applied: true;
}

// Not `completeLayout`: page builds run between the steps of a sliced completion.
const FRAME_REQUESTS = new Set<AwaitedRequest['type']>([
  'bootstrap',
  'sync',
  'buildFrame',
  'releasePages',
  'applyInput',
  'applyDelete',
  'proposal',
  'documentRead',
]);

type PendingRequest = {
  type: AwaitedRequest['type'];
  resolve(response: ResidentEngineWorkerResponse & { ok: true }): void;
  reject(error: Error): void;
};

type AwaitedRequest = Exclude<
  ResidentEngineWorkerRequestWithoutId,
  { type: 'applyUpdate' | 'eraseCaret' | 'destroy' }
>;

/**
 * How long the worker may stay silent while requests wait. A large document
 * legitimately takes seconds per request, and the main thread would take just
 * as long, so this only catches a worker that stopped answering altogether.
 */
export const RESIDENT_WORKER_SILENCE_MS = 60_000;

export interface ResidentEngineWorkerPort {
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  postMessage(
    message: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule,
    transfer?: Transferable[]
  ): void;
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
  private watchdog: ReturnType<typeof setTimeout> | null = null;
  private nextId = 1;
  private terminalError: Error | null = null;
  private ready = false;
  private revision = 0;
  private remoteVector: Uint8Array | null = null;
  private appliedFontsRevision: number | null = null;
  private bootstrapped = false;
  /** Set once `open` is sent, with the heap limit it opened under. */
  private openedHeapLimit: { bytes?: number } | null = null;
  private bootstraps = 0;
  /** The font-requirements read still in flight, shared while nothing was posted after it. */
  private fontRead: { layoutInput: string; id: number; reply: Promise<string> } | null = null;
  /** Id of the last snapshot request sent; replies to earlier requests must
   * not replace the state it recorded. */
  private lastSnapshotId = 0;
  private keepSurfaces = false;
  private bootstrapWaiters: Array<() => void> = [];
  private lastMemory: WasmModuleMemory[] | null = null;
  private answeredFrameEpoch = 0;
  private retainBuiltPages = false;
  private failureListener: ((error: Error) => void) | null = null;

  constructor(private readonly worker: ResidentEngineWorkerPort = spawnResidentEngineWorker()) {
    this.worker.onmessage = (event) => {
      const response = event.data;
      if (response.memory) this.lastMemory = response.memory;
      if (response.ok && response.caret) {
        this.answeredFrameEpoch = Math.max(this.answeredFrameEpoch, response.caret.frameEpoch);
      }
      if (response.ok && response.id >= this.lastSnapshotId) {
        const stateVector = response.proposal?.stateVector ?? response.stateVector;
        if (stateVector) this.remoteVector = new Uint8Array(stateVector);
      }
      if (!response.ok && response.terminal) {
        this.fail(
          response.outOfMemory
            ? new ResidentWorkerOutOfMemoryError(response.error, response.memory ?? [])
            : new ResidentWorkerUnavailableError(response.error)
        );
        return;
      }
      if (this.pending.size > 0) this.armWatchdog();
      const pending = this.pending.get(response.id);
      if (!pending) return;
      this.pending.delete(response.id);
      if (this.pending.size === 0) this.disarmWatchdog();
      if (response.ok) pending.resolve(response);
      else pending.reject(residentWorkerError(response.error, response.residentUnavailable));
    };
    this.worker.onerror = (event) => {
      this.fail(new ResidentWorkerFailureError(`Resident engine worker failed: ${event.message}`));
    };
    this.worker.onmessageerror = () => {
      this.fail(new ResidentWorkerFailureError('Resident engine worker returned an unreadable message'));
    };
  }

  isReady(): boolean {
    return this.ready;
  }

  /** @internal */
  hasFailed(): boolean {
    return this.terminalError !== null;
  }

  /** The worker's wasm memories as of its latest reply; null before one. */
  memory(): WasmModuleMemory[] | null {
    return this.lastMemory;
  }

  layoutRevision(): number {
    return this.revision;
  }

  /** @internal The newest frame epoch a reply carried; 0 before any frame. */
  answeredFrame(): number {
    return this.answeredFrameEpoch;
  }

  /** @internal */
  setRetainBuiltPages(retain: boolean): void {
    this.retainBuiltPages = retain;
  }

  /** @internal Called once when the worker fails, whether or not a request was waiting; not on `destroy`. */
  onFailure(listener: ((error: Error) => void) | null): void {
    this.failureListener = listener;
  }

  /** @internal Whether foreground document or frame work awaits its reply; `reads: false` leaves document reads out. */
  frameRequestPending(reads = true): boolean {
    for (const { type } of this.pending.values()) {
      if (FRAME_REQUESTS.has(type) && (reads || type !== 'documentRead')) return true;
    }
    return false;
  }

  /** The worker replica's last reported yrs state vector (null before any). */
  remoteStateVector(): Uint8Array | null {
    return this.remoteVector;
  }

  /** The fonts revision this worker last applied (null before bootstrap). */
  syncedFontsRevision(): number | null {
    return this.appliedFontsRevision;
  }

  /** A bootstrap was sent; later snapshots go as syncs queued behind it. */
  bootstrapSent(): boolean {
    return this.bootstrapped;
  }

  /** Resolves once a bootstrap is sent, at once when one was. */
  whenBootstrapSent(): Promise<void> {
    if (this.bootstrapped) return Promise.resolve();
    return new Promise((resolve) => this.bootstrapWaiters.push(resolve));
  }

  /**
   * Starts over with another document in the same worker: the next request
   * is a bootstrap, which keeps the page surfaces already attached so the
   * new document's pages paint where the old ones were.
   */
  rebootstrap(): void {
    this.bootstrapped = false;
    this.remoteVector = null;
    this.appliedFontsRevision = null;
    this.lastSnapshotId = 0;
    this.revision = 0;
    this.keepSurfaces = true;
    // The bootstrap it asks for frees the worker's document, opened there or not.
    this.openedHeapLimit = null;
  }

  /**
   * Parses and seeds a DOCX in the worker, which then holds the document:
   * its first layout is a bootstrap with `opened`, and the main replica loads
   * {@link encodeState}. A copy of `bytes` is transferred.
   */
  async open(
    bytes: Uint8Array,
    options: { digest?: string; generation?: string; heapLimitBytes?: number } = {}
  ): Promise<ResidentEngineWorkerOpened> {
    if (this.openedHeapLimit || this.bootstrapped) {
      throw new ResidentWorkerFailureError('Resident engine worker already holds a document');
    }
    const reservation = { bytes: options.heapLimitBytes };
    this.openedHeapLimit = reservation;
    let response: ResidentEngineWorkerResponse & { ok: true };
    try {
      const copy = new Uint8Array(bytes);
      response = await this.request(
        {
          type: 'open',
          bytes: copy.buffer,
          ...(options.digest !== undefined ? { digest: options.digest } : {}),
          ...(options.generation !== undefined ? { generation: options.generation } : {}),
          ...(options.heapLimitBytes !== undefined ? { heapLimitBytes: options.heapLimitBytes } : {}),
        },
        [copy.buffer]
      );
    } catch (error) {
      // The worker freed the session a failed open made, so it can open again.
      if (this.openedHeapLimit === reservation) this.openedHeapLimit = null;
      throw error;
    }
    if (response.hostJson === undefined || !response.stateVector) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted the opened document');
    }
    return { hostJson: response.hostJson, stateVector: new Uint8Array(response.stateVector) };
  }

  /**
   * Opens a display-only preview of the first `blocks` body blocks of a DOCX in the worker, or
   * resolves null when the package cannot open as a preview. Its first layout is a bootstrap
   * with `opened`; {@link rebootstrap} then lets {@link open} replace it with the whole document.
   * A copy of `bytes` is transferred.
   */
  async openPreview(
    bytes: Uint8Array,
    blocks: number,
    options: { heapLimitBytes?: number } = {}
  ): Promise<ResidentEngineWorkerOpened | null> {
    if (this.openedHeapLimit || this.bootstrapped) {
      throw new ResidentWorkerFailureError('Resident engine worker already holds a document');
    }
    const reservation = { bytes: options.heapLimitBytes };
    this.openedHeapLimit = reservation;
    let response: ResidentEngineWorkerResponse & { ok: true };
    try {
      const copy = new Uint8Array(bytes);
      response = await this.request(
        {
          type: 'open',
          bytes: copy.buffer,
          previewBlocks: blocks,
          ...(options.heapLimitBytes !== undefined ? { heapLimitBytes: options.heapLimitBytes } : {}),
        },
        [copy.buffer]
      );
    } catch (error) {
      if (this.openedHeapLimit === reservation) this.openedHeapLimit = null;
      throw error;
    }
    if (response.previewRefused) {
      if (this.openedHeapLimit === reservation) this.openedHeapLimit = null;
      return null;
    }
    if (response.hostJson === undefined || !response.stateVector) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted the opened preview');
    }
    return { hostJson: response.hostJson, stateVector: new Uint8Array(response.stateVector) };
  }

  /** The font requirements of a region layout request, read from the opened document. */
  fontRequirements(layoutInput: string): Promise<string> {
    const shared = this.fontRead;
    if (shared && shared.layoutInput === layoutInput && shared.id === this.nextId - 1) {
      return shared.reply;
    }
    const reply = this.request({ type: 'fontRequirements', layoutInput }).then((response) => {
      if (response.requirementsJson === undefined) {
        throw new ResidentWorkerFailureError('Resident engine worker omitted the font requirements');
      }
      return response.requirementsJson;
    });
    const read = { layoutInput, id: this.nextId - 1, reply };
    this.fontRead = read;
    const settle = () => {
      if (this.fontRead === read) this.fontRead = null;
    };
    reply.then(settle, settle);
    return reply;
  }

  /** @internal */
  async proposal(operation: ResidentProposalOperation): Promise<ResidentProposalReply> {
    if (!this.bootstrapped) {
      throw new ResidentWorkerFailureError('Resident engine worker has not laid out its document');
    }
    const response = await this.request({ type: 'proposal', operation });
    if (!response.proposal) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted the proposal result');
    }
    return {
      ...response.proposal,
      updates: response.proposal.updates.map((update) => new Uint8Array(update)),
      stateVector: new Uint8Array(response.proposal.stateVector),
    };
  }

  /** @internal */
  async documentRead<K extends ResidentDocumentRead['kind']>(
    read: ResidentDocumentRead & { kind: K }
  ): Promise<{ version: string; value: ResidentDocumentReadValues[K] }> {
    const response = await this.request({ type: 'documentRead', read });
    if (!response.read) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted the document read');
    }
    return response.read as { version: string; value: ResidentDocumentReadValues[K] };
  }

  /**
   * Reads the document at `expectVersion`. A later message may run first; the
   * read then answers `superseded` when the document moved on meanwhile.
   * @internal
   */
  async documentReadAt<K extends ResidentDocumentRead['kind']>(
    read: ResidentDocumentRead & { kind: K },
    expectVersion: string
  ): Promise<
    | { status: 'ok'; version: string; value: ResidentDocumentReadValues[K] }
    | { status: 'superseded' }
  > {
    const response = await this.request({ type: 'documentRead', read, expectVersion });
    if (response.superseded) return { status: 'superseded' };
    if (!response.read) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted the document read');
    }
    const { version, value } = response.read as {
      version: string;
      value: ResidentDocumentReadValues[K];
    };
    return { status: 'ok', version, value };
  }

  /** @internal */
  async handOver(): Promise<{
    state: Uint8Array;
    version: string;
    proposals: DocxProposalRegistryState;
  }> {
    const response = await this.request({ type: 'encodeState' });
    if (!response.state || response.version === undefined || !response.proposals) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted its document handoff');
    }
    return {
      state: new Uint8Array(response.state),
      version: response.version,
      proposals: response.proposals,
    };
  }

  /** The worker's whole document state as one yrs v1 update. */
  async encodeState(): Promise<Uint8Array> {
    const response = await this.request({ type: 'encodeState' });
    if (!response.state) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted its state');
    }
    return new Uint8Array(response.state);
  }

  async revisionCount(): Promise<number> {
    const response = await this.request({ type: 'revisionCount' });
    if (
      typeof response.revisionCount !== 'number' ||
      !Number.isInteger(response.revisionCount) ||
      response.revisionCount < 0
    ) {
      throw new ResidentWorkerFailureError('Resident engine worker omitted a valid revision count');
    }
    return response.revisionCount;
  }

  async warm(): Promise<void> {
    const response = this.request({ type: 'warm', hostModule: true });
    const postModule = (module: WebAssembly.Module | null): void => {
      if (this.terminalError) return;
      try {
        this.worker.postMessage({ type: 'editModule', module });
      } catch {
        try {
          this.worker.postMessage({ type: 'editModule', module: null });
        } catch {}
      }
    };
    void editWasmModule().then(postModule, () => postModule(null));
    await response;
  }

  async bootstrap(
    snapshot: YrsResidentWorkerSnapshot,
    extras: string,
    options: ResidentEngineWorkerLayoutOptions & ResidentEngineWorkerSnapshotOptions = {}
  ): Promise<ResidentEngineWorkerFrame> {
    if (options.opened && !this.openedHeapLimit) {
      throw new ResidentWorkerFailureError('Resident engine worker has no opened document');
    }
    if (
      options.opened &&
      options.heapLimitBytes !== undefined &&
      options.heapLimitBytes !== this.openedHeapLimit?.bytes
    ) {
      throw new ResidentWorkerFailureError(
        'Resident engine worker opened its document under another heap limit'
      );
    }
    const fontsRevision = snapshot.fontsRevision;
    this.bootstrapped = true;
    for (const resolve of this.bootstrapWaiters.splice(0)) resolve();
    const keepSurfaces = this.keepSurfaces;
    this.keepSurfaces = false;
    const generation = ++this.bootstraps;
    const pending = this.request(
      {
        type: 'bootstrap',
        snapshot,
        extras,
        expectedFrameEpoch: options.frameEpoch ?? 0,
        ...(options.layoutExtras !== undefined ? { layoutExtras: options.layoutExtras } : {}),
        ...(options.displayWindow
          ? {
              displayWindow: options.displayWindow,
              ...(this.retainBuiltPages ? { retainBuiltPages: true } : {}),
            }
          : {}),
        ...(options.provisionalPages !== undefined
          ? { provisionalPages: options.provisionalPages }
          : {}),
        ...(keepSurfaces ? { keepSurfaces: true } : {}),
        ...(options.opened ? { opened: true } : {}),
        ...(options.heapLimitBytes !== undefined ? { heapLimitBytes: options.heapLimitBytes } : {}),
      },
      snapshotTransfers(snapshot)
    );
    this.recordSent(options.stateVector, fontsRevision);
    let response: ResidentEngineWorkerResponse & { ok: true };
    try {
      response = await pending;
    } catch (error) {
      // The open it was queued behind failed, so the worker holds no document to bootstrap.
      if (options.opened && !this.openedHeapLimit && generation === this.bootstraps) {
        this.bootstrapped = false;
        this.remoteVector = null;
        this.appliedFontsRevision = null;
      }
      throw error;
    }
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
    paintCaret = false,
    options: ResidentEngineWorkerLayoutOptions & ResidentEngineWorkerSnapshotOptions = {}
  ): Promise<ResidentEngineWorkerFrame> {
    const fontsRevision = snapshot.fontsRevision;
    const pending = this.request(
      {
        type: 'sync',
        snapshot,
        extras,
        expectedFrameEpoch,
        paintCaret,
        ...(options.layoutExtras !== undefined ? { layoutExtras: options.layoutExtras } : {}),
        ...(options.provisionalPages !== undefined
          ? { provisionalPages: options.provisionalPages }
          : {}),
        ...(options.displayWindow
          ? {
              displayWindow: options.displayWindow,
              ...(this.retainBuiltPages ? { retainBuiltPages: true } : {}),
            }
          : {}),
      },
      snapshotTransfers(snapshot)
    );
    this.recordSent(options.stateVector, fontsRevision);
    const response = await pending;
    const result = frameResult(response);
    this.recordSync(response, fontsRevision);
    this.ready = true;
    this.revision = result.layoutRevision;
    return result;
  }

  /**
   * Lay out the rest of a provisional bootstrap layout: its frame and full
   * layout, or null when a later snapshot already replaced it.
   */
  async completeLayout(
    expectedFrameEpoch: number,
    paintCaret = false,
    sliceBlocks?: number
  ): Promise<ResidentEngineWorkerFrame | null> {
    const response = await this.request({
      type: 'completeLayout',
      expectedFrameEpoch,
      paintCaret,
      ...(sliceBlocks ? { sliceBlocks } : {}),
    });
    return response.frame ? frameResult(response) : null;
  }

  async buildFrame(
    extras: string,
    expectedFrameEpoch: number,
    paintCaret = false,
    displayWindow?: [number, number]
  ): Promise<ResidentEngineWorkerFrame> {
    const result = frameResult(
      await this.request({
        type: 'buildFrame',
        extras,
        expectedFrameEpoch,
        paintCaret,
        ...(displayWindow
          ? { displayWindow, ...(this.retainBuiltPages ? { retainBuiltPages: true } : {}) }
          : {}),
      })
    );
    return result;
  }

  /** Build unbuilt display pages; the reply frame carries them. */
  buildPages(
    pages: number[], expectedFrameEpoch: number, paintCaret?: boolean
  ): Promise<ResidentEngineWorkerFrame>;
  /** @internal */
  buildPages(
    pages: number[], expectedFrameEpoch: number, paintCaret: boolean, background: boolean
  ): Promise<ResidentEngineWorkerFrame | null>;
  async buildPages(
    pages: number[],
    expectedFrameEpoch: number,
    paintCaret = false,
    background = false
  ): Promise<ResidentEngineWorkerFrame | null> {
    const response = await this.request({
      type: 'buildPages', pages, expectedFrameEpoch, paintCaret,
      ...(background ? { background: true } : {}),
    });
    return background && response.pageBuildSuperseded ? null : frameResult(response);
  }

  async releasePages(
    pages: Array<{ index: number; pageId: string }>,
    expectedFrameEpoch: number,
    paintCaret = false
  ): Promise<ResidentEngineWorkerFrame | { superseded: true }> {
    const response = await this.request({
      type: 'releasePages',
      pages,
      expectedFrameEpoch,
      paintCaret,
    });
    return response.superseded ? { superseded: true } : frameResult(response);
  }

  async applyInput(
    text: string,
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false,
    paintCaret = false,
    displayWindow?: [number, number]
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
          ...(displayWindow
            ? { displayWindow, ...(this.retainBuiltPages ? { retainBuiltPages: true } : {}) }
            : {}),
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (inputUnavailable(error)) return { applied: false };
      throw error;
    }
  }

  async applyDelete(
    direction: 'backward' | 'forward',
    selection: YrsSelection,
    expectedFrameEpoch: number,
    profile = false,
    paintCaret = false,
    count = 1,
    displayWindow?: [number, number]
  ): Promise<ResidentEngineWorkerApplyResult | { applied: false }> {
    if (!this.ready) return { applied: false };
    try {
      const result = frameResult(
        await this.request({
          type: 'applyDelete',
          direction,
          count,
          selection,
          expectedFrameEpoch,
          profile,
          paintCaret,
          ...(displayWindow
            ? { displayWindow, ...(this.retainBuiltPages ? { retainBuiltPages: true } : {}) }
            : {}),
        })
      );
      return { applied: true, ...result };
    } catch (error) {
      if (inputUnavailable(error)) return { applied: false };
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
    this.fail(new ResidentWorkerFailureError('Resident engine worker was destroyed'), false);
  }

  private request(
    request: AwaitedRequest,
    transfer: Transferable[] = []
  ): Promise<ResidentEngineWorkerResponse & { ok: true }> {
    if (this.terminalError) return Promise.reject(this.terminalError);
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      if (this.pending.size === 0) this.armWatchdog();
      this.pending.set(id, { type: request.type, resolve, reject });
      this.worker.postMessage({ ...request, id } as ResidentEngineWorkerRequest, transfer);
    });
  }

  /**
   * Requests run in order, so once a snapshot is sent every later request
   * finds its state and fonts in the worker: the next sync can diff against
   * them before this one answers.
   */
  private recordSent(stateVector: Uint8Array | undefined, fontsRevision: number): void {
    if (this.terminalError) return;
    this.lastSnapshotId = this.nextId - 1;
    if (stateVector) this.remoteVector = stateVector.slice();
    this.appliedFontsRevision = fontsRevision;
  }

  /** Restarted by every message, so a queue of slow but answered requests never trips it. */
  private armWatchdog(): void {
    this.disarmWatchdog();
    this.watchdog = setTimeout(() => {
      this.watchdog = null;
      const waiting = this.pending.values().next().value?.type ?? 'a request';
      this.fail(
        new ResidentWorkerFailureError(
          `Resident engine worker did not answer ${waiting} within ${RESIDENT_WORKER_SILENCE_MS}ms`
        )
      );
    }, RESIDENT_WORKER_SILENCE_MS);
  }

  private disarmWatchdog(): void {
    if (this.watchdog === null) return;
    clearTimeout(this.watchdog);
    this.watchdog = null;
  }

  /** Record a successfully applied bootstrap/sync payload's fonts revision.
   * The state vector is tracked centrally in `onmessage`. */
  private recordSync(
    response: ResidentEngineWorkerResponse & { ok: true },
    fontsRevision: number
  ): void {
    if (response.id >= this.lastSnapshotId) this.appliedFontsRevision = fontsRevision;
  }

  private fail(error: Error, notify = true): void {
    const first = this.terminalError === null;
    this.terminalError = error;
    this.ready = false;
    this.disarmWatchdog();
    this.worker.terminate();
    for (const pending of this.pending.values()) pending.reject(error);
    this.pending.clear();
    for (const resolve of this.bootstrapWaiters.splice(0)) resolve();
    if (notify && first) this.failureListener?.(error);
  }
}

class ResidentWorkerUnavailableError extends Error {}

/** The worker itself failed (crash, timeout, torn-down, corrupt reply). */
export class ResidentWorkerFailureError extends Error {}

/** The worker trapped because its wasm memory could not grow any further. */
export class ResidentWorkerOutOfMemoryError extends ResidentWorkerUnavailableError {
  constructor(
    message: string,
    /** The worker's wasm memories when it trapped. */
    readonly memory: WasmModuleMemory[]
  ) {
    super(message);
  }
}

/** Input the worker could not take; running out of memory is left to the caller. */
function inputUnavailable(error: unknown): boolean {
  return (
    error instanceof ResidentWorkerUnavailableError &&
    !(error instanceof ResidentWorkerOutOfMemoryError)
  );
}

function residentWorkerError(message: string, unavailable = false): Error {
  return unavailable ? new ResidentWorkerUnavailableError(message) : new Error(message);
}

function snapshotTransfers(snapshot: YrsResidentWorkerSnapshot): Transferable[] {
  return [
    snapshot.state.buffer,
    ...snapshot.fonts.flatMap((font) => (font instanceof Uint8Array ? [font.buffer] : [])),
  ];
}

function frameResult(
  response: ResidentEngineWorkerResponse & { ok: true }
): ResidentEngineWorkerFrame {
  if (!response.frame)
    throw new ResidentWorkerFailureError('Resident engine worker response omitted its FrameDelta');
  if (!response.caret)
    throw new ResidentWorkerFailureError('Resident engine worker response omitted its caret snapshot');
  if (response.selection === undefined) {
    throw new ResidentWorkerFailureError('Resident engine worker response omitted its selection');
  }
  return {
    frame: new Uint8Array(response.frame),
    ...(response.pageFrames ? { pageFrames: response.pageFrames.map((frame) => new Uint8Array(frame)) } : {}),
    updates: (response.updates ?? []).map((update) => new Uint8Array(update)),
    engineMs: response.engineMs ?? 0,
    workerTotalMs: response.workerTotalMs ?? 0,
    engineProfile: response.engineProfile,
    caret: response.caret,
    selection: response.selection,
    selectionCursor: response.selectionCursor ?? null,
    caretPainted: response.caretPainted ?? false,
    replayMs: response.replayMs ?? 0,
    replayedPages: response.replayedPages ?? 0,
    layoutRevision: response.layoutRevision ?? 0,
    ...(response.documentVersion === undefined ? {} : { documentVersion: response.documentVersion }),
    ...(response.documentPreview === undefined ? {} : { documentPreview: response.documentPreview }),
    ...(response.documentAsOpened === undefined ? {} : { documentAsOpened: response.documentAsOpened }),
    deletedUnits: response.deletedUnits ?? 0,
    ...(response.layoutJson !== undefined ? { layoutJson: response.layoutJson } : {}),
    ...(response.layoutProvisional ? { layoutProvisional: true } : {}),
  };
}

export function canUseResidentEngineWorker(): boolean {
  return typeof Worker !== 'undefined';
}

interface PreloadedWorker {
  client: ResidentEngineWorkerClient;
  factory: typeof Worker;
  ready: Promise<void> | null;
  owners: number;
  idleTimer: ReturnType<typeof setTimeout> | null;
  releaseTimer: ReturnType<typeof setTimeout> | null;
}

let preloadedWorker: PreloadedWorker | null = null;

function clearPreloadedWorkerTimers(worker: PreloadedWorker): void {
  if (worker.idleTimer !== null) clearTimeout(worker.idleTimer);
  if (worker.releaseTimer !== null) clearTimeout(worker.releaseTimer);
  worker.idleTimer = null;
  worker.releaseTimer = null;
}

function discardPreloadedWorker(worker: PreloadedWorker): void {
  if (preloadedWorker !== worker) return;
  preloadedWorker = null;
  clearPreloadedWorkerTimers(worker);
  worker.client.destroy();
}

/** @internal */
export function preloadResidentEngineWorker(): Promise<void> {
  if (!canUseResidentEngineWorker()) return Promise.resolve();
  if (
    preloadedWorker &&
    (preloadedWorker.factory !== Worker || preloadedWorker.client.hasFailed())
  ) {
    discardPreloadedWorker(preloadedWorker);
  }
  if (!preloadedWorker) {
    preloadedWorker = {
      client: new ResidentEngineWorkerClient(),
      factory: Worker,
      ready: null,
      owners: 0,
      idleTimer: null,
      releaseTimer: null,
    };
  }
  const worker = preloadedWorker;
  if (!worker.ready) {
    worker.ready = worker.client.warm().catch((error: unknown) => {
      if (worker.client.hasFailed()) discardPreloadedWorker(worker);
      else worker.ready = null;
      throw error;
    });
  }
  clearPreloadedWorkerTimers(worker);
  if (worker.owners === 0) {
    worker.idleTimer = setTimeout(() => discardPreloadedWorker(worker), RESIDENT_WORKER_SILENCE_MS);
  }
  return worker.ready;
}

/** @internal */
export function retainPreloadedResidentEngineWorker(): () => void {
  const worker = preloadedWorker;
  if (!worker) return () => {};
  clearPreloadedWorkerTimers(worker);
  worker.owners += 1;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    worker.owners -= 1;
    if (worker.owners === 0 && preloadedWorker === worker) {
      worker.releaseTimer = setTimeout(() => discardPreloadedWorker(worker), 0);
    }
  };
}

/** @internal */
export function takePreloadedResidentEngineWorker(): ResidentEngineWorkerClient | null {
  const worker = preloadedWorker;
  if (!worker) return null;
  if (!canUseResidentEngineWorker() || worker.factory !== Worker || worker.client.hasFailed()) {
    discardPreloadedWorker(worker);
    return null;
  }
  preloadedWorker = null;
  clearPreloadedWorkerTimers(worker);
  return worker.client;
}
