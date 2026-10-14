import type { EchoEntry, EchoHooks } from '../../../../../../shared/office-session/commands/types';
import { decodeFrameDelta } from '@betteroffice/docx/layout/render';
import { closeCommandBitmap } from '@betteroffice/docx/peerless';
import type {
  DocxCommandFrame,
  DocxCommandFramePayload,
  DocxCommandPageIdentity,
  DocxStagedPageBitmap,
} from '@betteroffice/docx/peerless';
import { releaseBitmap, retainBitmap } from './bitmapOwnership';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import type { PeerlessFrameGeometry, PeerlessViewStore } from './viewStore';

/** @internal Geometry decoding and bitmap validation precede synchronous installation. */
export interface PeerlessFramePresentationOptions {
  viewStore: PeerlessViewStore;
  echoPainter: PeerlessEchoPainter;
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  prepareGeometry(frame: DocxCommandFrame): PeerlessFrameGeometry | undefined;
  requestFullFrame(): void;
}

/** @internal Shared transactional hooks accept a frame only after staged pixels install. */
export interface PeerlessFramePresentation {
  hooks: EchoHooks<PeerlessEchoIntent, DocxCommandFramePayload>;
  present(
    frame: DocxCommandFrame,
    retired: readonly EchoEntry<PeerlessEchoIntent>[],
    remaining: readonly EchoEntry<PeerlessEchoIntent>[]
  ): boolean;
  dispose(): void;
}

/** @internal */
export function createFramePresentation(
  options: PeerlessFramePresentationOptions
): PeerlessFramePresentation {
  let disposed = false;
  let revision: number | undefined;
  let coverage = -1;
  const authoritative = new Map<HTMLCanvasElement, { page: DocxCommandPageIdentity; bitmap: ImageBitmap }>();
  const handedOff = new WeakSet<ImageBitmap>();
  const owner = {};
  let identity = { ...options.viewStore.getSnapshot().identity };

  function reject(frame: DocxCommandFrame): false {
    for (const page of frame.payload.stagedPresentation?.pages ?? []) {
      for (const bitmap of [page.bitmap, ...(page.backgrounds ?? []).map((item) => item.bitmap)]) {
        if (!handedOff.has(bitmap)) closeCommandBitmap(bitmap);
      }
    }
    return false;
  }

  function setBackup(canvas: HTMLCanvasElement, next?: { page: DocxCommandPageIdentity; bitmap: ImageBitmap }): void {
    const previous = authoritative.get(canvas);
    if (next) {
      retainBitmap(next.bitmap, owner);
      authoritative.set(canvas, next);
    } else authoritative.delete(canvas);
    if (previous && ![...authoritative.values()].some(({ bitmap }) => bitmap === previous.bitmap)) {
      releaseBitmap(previous.bitmap, owner);
    }
  }

  function pruneBackups(geometry?: PeerlessFrameGeometry): void {
    for (const [canvas, backup] of authoritative) {
      const page = geometry?.pages.find((item) => item.page.pageId === backup.page.pageId)?.page;
      if (options.canvasForPage(backup.page) !== canvas || (geometry && !page)) setBackup(canvas);
      else if (page && page.pageIndex !== backup.page.pageIndex) authoritative.set(canvas, { ...backup, page });
    }
  }

  function repaint(canvas: HTMLCanvasElement, width: number, height: number, bitmap?: ImageBitmap): void {
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.save();
    try {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      canvas.width = width;
      canvas.height = height;
      ctx.clearRect(0, 0, width, height);
      if (bitmap) ctx.drawImage(bitmap, 0, 0);
    } finally { ctx.restore(); }
  }

  function install(
    frame: DocxCommandFrame,
    retired: readonly EchoEntry<PeerlessEchoIntent>[],
    remaining: readonly EchoEntry<PeerlessEchoIntent>[]
  ): boolean {
    if (disposed) return reject(frame);
    const current = options.viewStore.getSnapshot().identity;
    if (current.session !== identity.session || current.incarnation !== identity.incarnation ||
      current.epoch !== identity.epoch) {
      coverage = -1;
      revision = undefined;
      identity = { ...current };
      options.echoPainter.installContext(undefined);
    }
    pruneBackups();
    if (frame.coversThroughSeq < coverage ||
      (revision !== undefined && frame.revision <= revision)) {
      return reject(frame);
    }
    let geometry: PeerlessFrameGeometry | undefined;
    const staged = frame.payload.stagedPresentation;
    const targets: { page: DocxStagedPageBitmap; canvas: HTMLCanvasElement;
      ctx: CanvasRenderingContext2D; width: number; height: number }[] = [];
    let missingCanvas = false;
    try {
      if (frame.baseRevision !== undefined && frame.baseRevision !== revision) throw new Error('base-mismatch');
      let base = revision;
      const deltas = [...(frame.payload.pageFrames ?? []), frame.payload.frame].map(decodeFrameDelta);
      const first = deltas[0]!;
      if (first.full !== (frame.baseRevision === undefined)) throw new Error('base-mismatch');
      for (const delta of deltas) {
        if (!delta.full && delta.baseFrameEpoch !== base) throw new Error('base-mismatch');
        base = delta.frameEpoch;
      }
      if (base !== frame.revision || frame.payload.caret.frameEpoch !== frame.revision ||
        !staged || staged.frameEpoch !== frame.revision) throw new Error('frame-mismatch');
      geometry = options.prepareGeometry(frame);
      const identity = options.viewStore.getSnapshot().identity;
      if (!geometry || geometry.frameEpoch !== frame.revision ||
        geometry.coversThroughSeq !== frame.coversThroughSeq ||
        geometry.layoutRevision !== frame.payload.layoutRevision ||
        geometry.identity.session !== identity.session || geometry.identity.incarnation !== identity.incarnation ||
        geometry.identity.epoch !== identity.epoch) throw new Error('geometry-mismatch');
      const pages = new Set<string>();
      for (const page of staged.pages) {
        if (pages.has(page.page.pageId) || page.frameEpoch !== frame.revision ||
          page.bitmap.width <= 0 || page.bitmap.height <= 0) throw new Error('bitmap-mismatch');
        pages.add(page.page.pageId);
        const target = geometry.pages.find((candidate) => candidate.page.pageId === page.page.pageId &&
          candidate.page.pageIndex === page.page.pageIndex);
        if (!target) throw new Error('page-mismatch');
        const canvas = options.canvasForPage(page.page);
        if (!canvas) { missingCanvas = true; continue; }
        const ctx = canvas.getContext('2d');
        if (!ctx) throw new Error('missing-canvas');
        targets.push({ page, canvas, ctx, width: canvas.width, height: canvas.height });
      }
    } catch {
      options.requestFullFrame();
      return reject(frame);
    }
    if (missingCanvas) {
      options.requestFullFrame();
      return reject(frame);
    }
    try {
      for (const { page: { bitmap }, canvas, ctx } of targets) {
        ctx.save();
        try {
          ctx.setTransform(1, 0, 0, 1, 0, 0);
          if (canvas.width !== bitmap.width) canvas.width = bitmap.width;
          if (canvas.height !== bitmap.height) canvas.height = bitmap.height;
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(bitmap, 0, 0);
        } finally { ctx.restore(); }
      }
    } catch {
      for (const canvas of new Set([...authoritative.keys(), ...targets.map((item) => item.canvas)])) {
        const backup = authoritative.get(canvas)?.bitmap;
        const target = targets.find((item) => item.canvas === canvas);
        repaint(canvas, backup?.width ?? target!.width, backup?.height ?? target!.height, backup);
      }
      options.echoPainter.paint([...retired, ...remaining]);
      options.requestFullFrame();
      return reject(frame);
    }
    for (const { page, canvas } of targets) {
      for (const [previous, backup] of authoritative) {
        if (previous !== canvas && backup.page.pageId === page.page.pageId) setBackup(previous);
      }
      setBackup(canvas, { page: page.page, bitmap: page.bitmap });
      handedOff.add(page.bitmap);
      for (const background of page.backgrounds ?? []) handedOff.add(background.bitmap);
    }
    pruneBackups(geometry);
    const harness = (globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean }).__DOCX_PEERLESS_HARNESS__
      ? (globalThis as typeof globalThis & { __docxPeerlessHarness?: {
        presentationProbe(stage: 'installed' | 'reconciled', frame: DocxCommandFrame,
          retired: readonly EchoEntry<PeerlessEchoIntent>[], remaining: readonly EchoEntry<PeerlessEchoIntent>[],
          canvases: readonly HTMLCanvasElement[]): void;
        viewportProbe?(geometry: PeerlessFrameGeometry): void;
      } }).__docxPeerlessHarness : undefined;
    try {
      harness?.viewportProbe?.(geometry);
      options.viewStore.installFrameGeometry(geometry);
      pruneBackups();
      harness?.presentationProbe('installed', frame, retired, remaining, targets.map(({ canvas }) => canvas));
      options.echoPainter.reconcile(frame, retired, remaining);
      harness?.presentationProbe('reconciled', frame, retired, remaining, targets.map(({ canvas }) => canvas));
    } catch {
      options.echoPainter.disable('reconcile');
      for (const [canvas, { bitmap }] of authoritative) repaint(canvas, bitmap.width, bitmap.height, bitmap);
    }
    coverage = frame.coversThroughSeq;
    revision = frame.revision;
    return true;
  }

  return {
    hooks: {
      paint(pending) { if (!disposed) { pruneBackups(); options.echoPainter.paint(pending); } },
      drop(entries, reason) { if (!disposed) { pruneBackups(); options.echoPainter.drop(entries, reason); } },
      present: install,
    },
    present: install,
    dispose() {
      disposed = true;
      for (const canvas of authoritative.keys()) setBackup(canvas);
    },
  };
}

/** @internal Synchronously installs pixels, retires covered echo, and redraws later echo. */
export function present(
  presentation: PeerlessFramePresentation,
  frame: DocxCommandFrame,
  retired: readonly EchoEntry<PeerlessEchoIntent>[],
  remaining: readonly EchoEntry<PeerlessEchoIntent>[]
): boolean {
  return presentation.present(frame, retired, remaining);
}
