import { engines, loadCanvas } from '#agents-engines';
import JSZip from 'jszip';
import { imageSize } from 'image-size';
import { createFontProvider, resolveGlyphFallbackFace, type BundledFontSource } from '@betteroffice/fonts';
import type { PresentationHandle, CanvasImageResolver, PptxFontFace, ShapeSnapshot, SlideSnapshot } from '@betteroffice/pptx';
import { DocumentToolError, type RenderedPage } from './types';

export async function renderPptxSlide(
  handle: PresentationHandle,
  page: number,
  scale = 1,
  provider?: BundledFontSource,
  theme?: PptxThemeFonts,
): Promise<RenderedPage> {
  const snapshot = handle.snapshot();
  const slide = snapshot.slides[page - 1];
  if (!slide) throw new DocumentToolError('SLIDE_OUT_OF_RANGE', `Choose a slide index from 1 to ${snapshot.slides.length}.`);
  if (!Number.isFinite(scale) || scale < 0.25 || scale > 3) throw new DocumentToolError('INVALID_ARGUMENT', 'scale must be from 0.25 to 3.');
  try {
    if (provider) await registerPptxFonts(handle, slide, provider, theme);
    const png = await handle.renderPng({ slide: slide.id }, { scale });
    const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
    return { png, page, pageCount: snapshot.slides.length, width: view.getUint32(16), height: view.getUint32(20), warnings: ['Preview uses the presentation engine and bundled substitute fonts.'] };
  } catch (error) {
    if ((error as { code?: unknown } | null)?.code !== 'canvas-unavailable') throw error;
  }
  return renderPptxSlideCanvas(handle, page, scale, provider, theme);
}

async function registerPptxFonts(handle: PresentationHandle, slide: SlideSnapshot, provider: BundledFontSource, theme?: PptxThemeFonts) {
  const faces = pptxFontFaces([slide], theme);
  if (faces.length > 64) throw new DocumentToolError('FONT_LIMIT', 'Slide uses more than 64 font styles. Reduce font variety before previewing.');
  const loaded = await loadPptxFonts(faces, provider);
  for (const font of loaded.fonts) handle.registerFont(font);
  for (const font of loaded.fallbackFonts) handle.registerFallbackFont(font);
  return loaded;
}

async function renderPptxSlideCanvas(handle: PresentationHandle, page: number, scale: number, provider?: BundledFontSource, theme?: PptxThemeFonts): Promise<RenderedPage> {
  const { createCanvas, GlobalFonts, loadImage } = await loadCanvas();
  const { module } = await engines.load('pptx');
  provider ??= createFontProvider(undefined, { decoder: module.decodeFontFace });
  const snapshot = handle.snapshot();
  const slide = snapshot.slides[page - 1]!;
  const warnings = new Set<string>(['Preview uses the presentation engine and bundled substitute fonts.']);
  const { fonts, fallbackFonts } = await registerPptxFonts(handle, slide, provider, theme);
  for (const font of [...fonts, ...fallbackFonts]) GlobalFonts.register(Buffer.from(font.bytes), font.family);
  const list = handle.layoutSlide(page - 1);
  const width = Math.ceil(list.width * scale);
  const height = Math.ceil(list.height * scale);
  if (![width, height].every(n => Number.isSafeInteger(n) && n > 0) || width * height > 16000000) {
    throw new DocumentToolError('PAGE_TOO_LARGE', 'Slide exceeds the 16 megapixel preview limit. Lower scale.');
  }
  const canvas = createCanvas(width, height);
  const cache = new Map<string, CanvasImageSource | null>();
  let pixels = 0;
  const resolveImage: CanvasImageResolver = async asset => {
    if (cache.has(asset)) return cache.get(asset)!;
    let result: CanvasImageSource | null = null;
    try {
      const bytes = handle.mediaBytes(asset);
      if (bytes.length > 8 * 1024 * 1024) throw new Error('Image exceeds the 8 MiB preview limit.');
      const info = imageSize(bytes);
      if (!['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'].includes(info.type ?? '')) throw new Error('Image preview supports raster images only.');
      const size = info.width * info.height;
      if (!Number.isSafeInteger(size) || size <= 0 || size > 16000000 || pixels + size > 32000000) throw new Error('Image exceeds the preview pixel limit.');
      pixels += size;
      result = await loadImage(Buffer.from(bytes)) as unknown as CanvasImageSource;
    } catch (error) { warnings.add(error instanceof Error ? error.message : 'Image could not be decoded.'); }
    cache.set(asset, result);
    return result;
  };
  await module.paintSlide(canvas.getContext('2d') as unknown as CanvasRenderingContext2D, list, 1, scale, { resolveImage, maxShadowPixels: 0 });
  return { png: new Uint8Array(await canvas.encode('png')), page, pageCount: snapshot.slides.length, width, height, warnings: [...warnings, 'Shadow effects are omitted in server previews.'] };
}

export interface PptxFontStyle { family: string; bold: boolean; italic: boolean }

/** The Latin faces of a deck's theme fonts, for runs that name no face of their own. */
export interface PptxThemeFonts { major: string[]; minor: string[] }

/** Reads the theme's Latin faces. */
export async function pptxThemeFonts(bytes: Uint8Array): Promise<PptxThemeFonts> {
  const fonts = { major: new Set<string>(), minor: new Set<string>() };
  let zip: JSZip;
  try { zip = await JSZip.loadAsync(bytes); } catch { return { major: [], minor: [] }; }
  for (const file of Object.values(zip.files)) {
    if (file.dir || !/^ppt\/theme\/[^/]+\.xml$/.test(file.name)) continue;
    const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
    if (size !== undefined && size > 16 * 1024 * 1024) continue;
    let xml: string;
    try { xml = await file.async('string'); } catch { continue; }
    xml = xml.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g, '');
    const stack: { local: string; kind?: 'major' | 'minor' }[] = [];
    const found = { major: false, minor: false };
    for (const [, closing, name, attributes] of xml.matchAll(/<(\/?)([^\s/<>!?]+)(\s(?:[^"'<>]|"[^"<]*"|'[^'<]*')*|\/?)>/g)) {
      if (closing) { stack.pop(); continue; }
      const local = name.slice(name.lastIndexOf(':') + 1);
      const parent = stack.at(-1);
      if (local === 'latin' && parent?.kind && !found[parent.kind] && stack.at(-2)?.local === 'fontScheme') {
        found[parent.kind] = true;
        const family = attributes.match(/(?:^|\s)typeface\s*=\s*(["'])([\s\S]*?)\1/)?.[2];
        if (family) fonts[parent.kind].add(family);
        if (found.major && found.minor) break;
      }
      if (!/\/\s*$/.test(attributes)) stack.push({ local, kind: local === 'majorFont' ? 'major' : local === 'minorFont' ? 'minor' : undefined });
    }
  }
  return { major: [...fonts.major], minor: [...fonts.minor] };
}

/** Every run style the slides use, Arial first so the renderer's first face is the same for every call, at most `limit` styles. */
export function pptxFontFaces(slides: SlideSnapshot[], theme: PptxThemeFonts = { major: [], minor: [] }, limit = Infinity): PptxFontStyle[] {
  const requests = new Map<string, { families: string[]; bold: boolean; italic: boolean }>();
  function collect(shapes: ShapeSnapshot[]) {
    for (const shape of shapes) {
      for (const story of shape.textStories) for (const paragraph of story.paragraphs) for (const run of paragraph.runs) {
        const named = run.style.fontFamily;
        const bold = run.style.bold ?? false;
        const italic = run.style.italic ?? false;
        const source = !named ? 'theme' : named.startsWith('+mj') ? 'major' : named.startsWith('+mn') ? 'minor' : undefined;
        const key = `${source ? `@${source}` : named}:${bold}:${italic}`;
        if (requests.has(key)) continue;
        const families = source === 'theme' ? [...theme.major, ...theme.minor] : source ? theme[source] : [named!];
        requests.set(key, { families: families.length ? families : ['Arial'], bold, italic });
      }
      collect(shape.children);
    }
  }
  for (const slide of slides) collect(slide.shapes);
  const faces = new Map<string, PptxFontStyle>();
  faces.set('Arial:false:false', { family: 'Arial', bold: false, italic: false });
  for (const { families, bold, italic } of requests.values()) {
    for (const family of families) {
      if (faces.size >= limit) return [...faces.values()];
      faces.set(`${family}:${bold}:${italic}`, { family, bold, italic });
    }
  }
  return [...faces.values()].slice(0, limit);
}

/** The bundled face for each style, and the glyph coverage faces they need. */
export async function loadPptxFonts(
  faces: PptxFontStyle[],
  provider: BundledFontSource = createFontProvider(),
): Promise<{ fonts: PptxFontFace[]; fallbackFonts: PptxFontFace[] }> {
  const fonts: PptxFontFace[] = [];
  const fallbackFonts: PptxFontFace[] = [];
  const coverageFaces = new Set<string>();
  for (const face of faces) {
    let buffer: ArrayBuffer | undefined;
    try {
      const load = provider.resolve(face.family, face.bold, face.italic);
      buffer = await load?.();
    } catch {}
    buffer ??= await provider.resolveLastResort(face.family, face.bold, face.italic, 'powerpoint')();
    fonts.push({ ...face, bytes: new Uint8Array(buffer) });
    const coverageFace = resolveGlyphFallbackFace(face.family, face.bold, face.italic, 'powerpoint');
    const coverage = provider.resolveGlyphFallback?.(face.family, face.bold, face.italic, 'powerpoint');
    const coverageKey = coverageFace && `${coverageFace.file}:${coverageFace.weight === 700}:${face.italic}`;
    if (coverage && coverageFace && coverageKey && !coverageFaces.has(coverageKey)) {
      let bytes: Uint8Array;
      try { bytes = new Uint8Array(await coverage()); } catch { continue; }
      coverageFaces.add(coverageKey);
      fallbackFonts.push({ family: coverageFace.family, bold: coverageFace.weight === 700, italic: face.italic, bytes });
    }
  }
  return { fonts, fallbackFonts };
}
