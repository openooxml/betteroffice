import { pageToPng } from '../render/png';
import type { PageDisplayList, VsdxFontFace } from '../types';
import type { VsdxDocument, VsdxRenderer } from './generated/vsdx_wasm.js';

export interface VsdxPngEngine {
  renderer: VsdxRenderer;
  doc: VsdxDocument;
  fonts(): VsdxFontFace[];
  displayList(page: number): PageDisplayList;
  media(assetId: string): Uint8Array;
}

export function renderPng(
  engine: VsdxPngEngine, target: { page: number }, options: { scale?: number; signal?: AbortSignal },
): Promise<Uint8Array> {
  return pageToPng(engine.displayList(target.page), { ...options, fonts: engine.fonts(), media: engine.media });
}
