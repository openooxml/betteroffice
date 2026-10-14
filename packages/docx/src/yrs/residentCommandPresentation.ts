import { transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import type {
  DocxCommandEchoContext,
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxCommandStagedPresentation,
} from './residentCommandProtocol';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';

const discardedBitmaps = new WeakSet<ImageBitmap>();

export function closeCommandBitmap(bitmap: ImageBitmap): void {
  if (discardedBitmaps.has(bitmap)) return;
  discardedBitmaps.add(bitmap);
  bitmap.close();
}

export function discardCommandPresentation(frame: DocxCommandFrame | undefined): void {
  for (const page of frame?.payload?.stagedPresentation?.pages ?? []) {
    for (const bitmap of [page.bitmap, ...(page.backgrounds ?? []).map((background) => background.bitmap)]) {
      closeCommandBitmap(bitmap);
    }
  }
}

/** @internal Worker-owned raster inputs; output bitmaps transfer to main. */
export interface StageCommandPresentationInput {
  layout: DocxCommandSelectionLayout;
  pages: readonly DocxCommandPageIdentity[];
  pixelRatio: number;
  echoContext?: DocxCommandEchoContext;
  rasterizePage(page: DocxCommandPageIdentity): Promise<ImageBitmap>;
  rasterizeBackground(
    dependency: DocxCommandEchoContext['damage'][number]
  ): Promise<ImageBitmap>;
}

/** @internal */
export async function stageCommandPresentation(
  input: StageCommandPresentationInput
): Promise<TransferResult<DocxCommandStagedPresentation>> {
  const bitmaps = new Set<ImageBitmap>();
  const pages: DocxCommandStagedPresentation['pages'][number][] = [];
  const frameEpoch = input.layout.frame.frameEpoch;
  try {
    if (!Number.isFinite(input.pixelRatio) || input.pixelRatio <= 0) throw new Error('Invalid pixel ratio');
    if (input.echoContext && input.echoContext.frameEpoch !== frameEpoch) throw new Error('Echo frame mismatch');
    for (const page of input.pages) {
      const bitmap = await input.rasterizePage(page);
      bitmaps.add(bitmap);
      const backgrounds: NonNullable<DocxCommandStagedPresentation['pages'][number]['backgrounds']>[number][] = [];
      for (const dependency of input.echoContext?.damage ?? []) {
        if (dependency.page.pageId !== page.pageId || dependency.page.pageIndex !== page.pageIndex) continue;
        const background = await input.rasterizeBackground(dependency);
        bitmaps.add(background);
        backgrounds.push({ id: dependency.backgroundId, rect: dependency.rect, bitmap: background });
      }
      pages.push({ page, frameEpoch, bitmap, ...(backgrounds.length ? { backgrounds } : {}) });
    }
    return transferable({ frameEpoch, pages }, [...bitmaps]);
  } catch (error) {
    for (const bitmap of bitmaps) bitmap.close();
    throw error;
  }
}
