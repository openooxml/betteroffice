import { GlyphCache } from '../layout/render/glyphCache';
import { createCanvasImageResolver } from '../layout/render/canvasImageResolver';
import { rasterizeDisplayPageToBackBuffer, type PageCanvasBuffer } from '../layout/render/canvasBackend';
import { canvasPngAvailable, canvasUnavailable, fontUnavailable, pagePngRequest, type DocxPngOptions } from '../layout/render/png';
import type { PngPictureUnavailableReason, PngUnavailablePicture } from '../../../../shared/office-session/pngExportError';
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
const METAFILE_PLACEHOLDER = /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" width="[\d.]+" height="[\d.]+" viewBox="0 0 [\d.]+ [\d.]+" preserveAspectRatio="none"><rect x="0\.5" y="0\.5" width="[\d.]+" height="[\d.]+" fill="#f1f3f4" stroke="#c4c7c5"\/><g transform="translate\([\d.]+ [\d.]+\) scale\([\d.]+\)" fill="none" stroke="#9aa0a6" stroke-width="1\.5" stroke-linejoin="round"><rect x="-12" y="-9" width="24" height="18" rx="2"\/><circle cx="-5" cy="-3\.5" r="2"\/><path d="M-10 7L-3 0L2 5L5 2L10 7"\/><\/g><\/svg>$/;

/** Whether display bytes are the stand-in the parser shows for a metafile it could not convert. */
function metafilePlaceholder(source: { bytes: Uint8Array; mimeType: string } | null): boolean {
  return source?.mimeType === 'image/svg+xml' && source.bytes.byteLength < 4096 && METAFILE_PLACEHOLDER.test(new TextDecoder().decode(source.bytes));
}

/** Paints one page with the editor's canvas backend and encodes it. */
export async function renderPng(engine: DocxPngEngine, options: Pick<DocxPngOptions, 'scale' | 'signal' | 'onDiagnostic'>): Promise<Uint8Array> {
  const { width, height, scale } = pagePngRequest(engine.page, options);
  if (!canvasPngAvailable()) throw canvasUnavailable();
  options.signal?.throwIfAborted();
  const canvas = createCanvas(width, height);
  if (!canvas.getContext('2d')) throw canvasUnavailable();
  let refused: Error | undefined;
  const reasons = new Map<string, PngPictureUnavailableReason>();
  const pictures: PngUnavailablePicture[] = [];
  const resolve = createCanvasImageResolver({ media: engine.media, onUnavailable: (id, reason) => reasons.set(id, reason) });
  await rasterizeDisplayPageToBackBuffer(canvas, { ...engine.page, background: engine.page.background ?? '#ffffff' }, {
    glyphCache: new GlyphCache({ provider: engine.outlineGlyph }),
    onGlyphFallback: () => { refused ??= fontUnavailable('A glyph on this page has no outline in the face layout measured it with.'); },
    resolveImage: (id) => {
      const source = engine.media(id);
      if (metafilePlaceholder(source)) {
        reasons.set(id, 'metafile');
        return null;
      }
      if (typeof Image === 'undefined' && (SVG_DATA_URL.test(id) || source?.mimeType === 'image/svg+xml')) {
        reasons.set(id, 'unsupported');
        return null;
      }
      return resolve(id);
    },
    onPictureUnavailable: (id) => { pictures.push({ id: pictureId(id), reason: reasons.get(id) ?? 'missing' }); },
  }, scale);
  if (refused) throw refused;
  options.signal?.throwIfAborted();
  const png = new Uint8Array(await (await encode(canvas)).arrayBuffer());
  if (pictures.length) options.onDiagnostic?.({ code: 'picture-unavailable', pictures });
  return png;
}

/** A picture's reference as diagnostics name it: a `data:` URL by its header, never its payload. */
export function pictureId(id: string): string {
  const comma = id.indexOf(',');
  return id.startsWith('data:') && comma >= 0 ? id.slice(0, comma) : id;
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
