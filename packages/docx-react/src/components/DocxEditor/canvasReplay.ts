import type { GlyphCache, ImageResolver, RetainedFrame } from '@betteroffice/docx/layout/render';

export interface CanvasRasterEnvironment {
  dpr: number;
  zoom: number;
  glyphCache?: GlyphCache;
  resolveImage?: ImageResolver;
}

interface CanvasPresentation {
  canvas: HTMLCanvasElement;
  pageId: bigint | undefined;
  revision: number | undefined;
  environment: CanvasRasterEnvironment;
}

export class CanvasReplayState {
  private frame?: RetainedFrame | null;
  private revision = 0;
  private pageRevisions = new Map<bigint, number>();
  private presented = new WeakMap<HTMLCanvasElement, CanvasPresentation>();

  updateFrame(frame: RetainedFrame | null | undefined): void {
    if (this.frame === frame) return;
    const missedFrame = Boolean(
      frame &&
        this.frame &&
        (frame.docEpoch !== this.frame.docEpoch || frame.frameEpoch !== this.frame.frameEpoch + 1)
    );
    this.frame = frame;
    this.revision += 1;
    if (!frame) {
      this.pageRevisions.clear();
      return;
    }
    const revisions = new Map<bigint, number>();
    for (const page of frame.pages) {
      revisions.set(
        page.pageId,
        missedFrame || frame.damagedPageIds.has(page.pageId)
          ? this.revision
          : this.pageRevisions.get(page.pageId) ?? this.revision
      );
    }
    this.pageRevisions = revisions;
  }

  prepare(
    canvas: HTMLCanvasElement,
    pageId: bigint | undefined,
    environment: CanvasRasterEnvironment
  ): CanvasPresentation | null {
    const revision = pageId === undefined ? undefined : this.pageRevisions.get(pageId);
    const previous = this.presented.get(canvas);
    if (
      revision !== undefined &&
      previous !== undefined &&
      previous.pageId === pageId &&
      previous.revision === revision &&
      previous.environment.dpr === environment.dpr &&
      previous.environment.zoom === environment.zoom &&
      previous.environment.glyphCache === environment.glyphCache &&
      previous.environment.resolveImage === environment.resolveImage
    ) {
      return null;
    }
    return { canvas, pageId, revision, environment };
  }

  didPresent(presentation: CanvasPresentation): void {
    this.presented.set(presentation.canvas, presentation);
  }

  release(canvas: HTMLCanvasElement): void {
    this.presented.delete(canvas);
  }
}

export interface CanvasReplayPreparation {
  buffer: HTMLCanvasElement;
  ready: Promise<unknown>;
  present(): void;
}

export async function presentCanvasReplay(
  preparations: CanvasReplayPreparation[],
  isCurrent: () => boolean
): Promise<void> {
  try {
    const results = await Promise.allSettled(preparations.map(({ ready }) => ready));
    const failure = results.find((result) => result.status === 'rejected');
    if (failure?.status === 'rejected') throw failure.reason;
    if (isCurrent()) {
      for (const preparation of preparations) preparation.present();
    }
  } finally {
    for (const { buffer } of preparations) {
      buffer.width = 0;
      buffer.height = 0;
    }
  }
}
