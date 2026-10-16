import { transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import type {
  DocxCommandEchoContext,
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxCommandStagedPresentation,
} from './residentCommandProtocol';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';
import type { DisplayPage, DisplayPrimitive } from '../layout/render/displayList';
import { displayPrimitiveRect } from '../layout/render/displayListGeometry';
import type { RetainedFrame } from '../layout/render/frameDelta';

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

/**
 * @internal The page an echo background rasters: everything but the echoed line, reduced to what can
 * paint into the damage rect, so the background costs the rect rather than the page.
 */
export function echoBackgroundPage(
  retained: RetainedFrame['pages'][number],
  dependency: DocxCommandEchoContext['damage'][number]
): DisplayPage {
  const omitted = new Set(dependency.omittedPrimitiveIds.map(BigInt));
  const { rect } = dependency;
  const reaches = (primitive: DisplayPrimitive) => {
    if (primitive.kind === 'image' || primitive.kind === 'shape' || primitive.clipGroup ||
      (primitive.kind === 'glyphRun' || primitive.kind === 'text') && (primitive.rotationDeg || primitive.modernEffects ||
        primitive.textShadow)) return true;
    const box = displayPrimitiveRect(primitive);
    const margin = primitive.kind === 'glyphRun' ? primitive.size : primitive.kind === 'text' ? box.h : 1;
    return box.x - margin < rect.x + rect.width && box.x + box.w + margin > rect.x &&
      box.y - margin < rect.y + rect.height && box.y + box.h + margin > rect.y;
  };
  const page = retained.page;
  const kept = (primitive: DisplayPrimitive, index: number) =>
    !omitted.has(retained.primitiveIds[index]!) && reaches(primitive);
  const region = <T extends { primitives: DisplayPrimitive[] }>(value: T | undefined) =>
    value && { ...value, primitives: value.primitives.filter(reaches) };
  return { ...page,
    watermarkPrimitiveCount: page.primitives.slice(0, page.watermarkPrimitiveCount ?? 0).filter(kept).length,
    primitives: page.primitives.filter(kept),
    header: region(page.header), footer: region(page.footer),
    ...(page.noteAreas ? { noteAreas: page.noteAreas.map((area) => ({ ...area,
      ...(area.separatorPrimitives ? { separatorPrimitives: area.separatorPrimitives.filter(reaches) } : {}),
      ...(area.primitives ? { primitives: area.primitives.filter(reaches) } : {}) })) } : {}),
  };
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
