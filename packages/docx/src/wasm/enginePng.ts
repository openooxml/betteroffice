import { GlyphCache } from '../layout/render/glyphCache';
import { createCanvasImageResolver } from '../layout/render/canvasImageResolver';
import { rasterizeDisplayPageToBackBuffer, type PageCanvasBuffer } from '../layout/render/canvasBackend';
import { canvasPngAvailable, canvasUnavailable, fontUnavailable, pagePngRequest, pictureUnavailable, type DocxPngOptions } from '../layout/render/png';
import type { DisplayList, DisplayPage } from '../layout/render/displayList';

/** What an engine needs to draw one page, from the session that laid it out. */
export interface DocxPngEngine {
  displayList: DisplayList;
  /** The font ids layout measured each family with, for engines that rasterize the list themselves. */
  fontChains: Record<string, number[]>;
  page: DisplayPage;
  outlineGlyph(fontId: number, glyphId: number): string;
  media(token: string): { bytes: Uint8Array; mimeType: string } | null;
  /** The session, for engines that render PNG themselves. */
  session: unknown;
}

/** Which engine this module draws with; the native package swaps in its own. */
export const engineKind: 'wasm' | 'native' = 'wasm';

const SVG_DATA_URL = /^data:image\/svg/i;

/** Paints one page with the editor's canvas backend and encodes it. */
export async function renderPng(engine: DocxPngEngine, options: Pick<DocxPngOptions, 'scale' | 'signal'>): Promise<Uint8Array> {
  const { width, height, scale } = pagePngRequest(engine.page, options);
  if (!canvasPngAvailable()) throw canvasUnavailable();
  options.signal?.throwIfAborted();
  const canvas = createCanvas(width, height);
  if (!canvas.getContext('2d')) throw canvasUnavailable();
  let refused: Error | undefined;
  const resolve = createCanvasImageResolver({ media: engine.media });
  await rasterizeDisplayPageToBackBuffer(canvas, { ...engine.page, background: engine.page.background ?? '#ffffff' }, {
    glyphCache: new GlyphCache({ provider: engine.outlineGlyph }),
    onGlyphFallback: () => { refused ??= fontUnavailable('A glyph on this page has no outline in the face layout measured it with.'); },
    resolveImage: (id) => {
      if (typeof Image === 'undefined' && (SVG_DATA_URL.test(id) || engine.media(id)?.mimeType === 'image/svg+xml')) {
        refused ??= pictureUnavailable();
        return null;
      }
      return resolve(id);
    },
  }, scale);
  if (refused) throw refused;
  options.signal?.throwIfAborted();
  return new Uint8Array(await (await encode(canvas)).arrayBuffer());
}

function createCanvas(width: number, height: number): PageCanvasBuffer {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(width, height);
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function encode(canvas: PageCanvasBuffer): Promise<Blob> {
  if ('convertToBlob' in canvas) return canvas.convertToBlob({ type: 'image/png' });
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('The canvas produced no PNG.'));
    }, 'image/png');
  });
}

export default renderPng;
