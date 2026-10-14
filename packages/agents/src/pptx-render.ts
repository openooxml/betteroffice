import { createCanvas, GlobalFonts, loadImage } from '@napi-rs/canvas';
import { imageSize } from 'image-size';
import { createFontProvider } from '@betteroffice/fonts';
import { paintSlide, type PresentationHandle, type CanvasImageResolver, type ShapeSnapshot } from '@betteroffice/pptx';
import { DocumentToolError, type RenderedPage } from './types';

export async function renderPptxSlide(handle: PresentationHandle, page: number, scale = 1): Promise<RenderedPage> {
  const snapshot = handle.snapshot();
  const slide = snapshot.slides[page - 1];
  if (!slide) throw new DocumentToolError('SLIDE_OUT_OF_RANGE', `Choose a slide index from 1 to ${snapshot.slides.length}.`);
  if (!Number.isFinite(scale) || scale < 0.25 || scale > 3) throw new DocumentToolError('INVALID_ARGUMENT', 'scale must be from 0.25 to 3.');
  const warnings = new Set<string>(['Preview uses the presentation engine and bundled substitute fonts.']);
  const provider = createFontProvider();
  const faces = new Map<string, { family: string; bold: boolean; italic: boolean }>();
  faces.set('Arial:false:false', { family: 'Arial', bold: false, italic: false });
  function collect(shapes: ShapeSnapshot[]) {
    for (const shape of shapes) {
      for (const story of shape.textStories) for (const paragraph of story.paragraphs) for (const run of paragraph.runs) {
        const family = run.style.fontFamily ?? 'Arial';
        const bold = run.style.bold ?? false;
        const italic = run.style.italic ?? false;
        faces.set(`${family}:${bold}:${italic}`, { family, bold, italic });
      }
      collect(shape.children);
    }
  }
  collect(slide.shapes);
  if (faces.size > 64) throw new DocumentToolError('FONT_LIMIT', 'Slide uses more than 64 font styles. Reduce font variety before previewing.');
  for (const face of faces.values()) {
    const load = provider.resolve(face.family, face.bold, face.italic) ?? provider.resolveLastResort(face.family, face.bold, face.italic, 'powerpoint');
    const bytes = new Uint8Array(await load());
    handle.registerFont({ ...face, bytes });
    GlobalFonts.register(Buffer.from(bytes), face.family);
  }
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
  await paintSlide(canvas.getContext('2d') as unknown as CanvasRenderingContext2D, list, 1, scale, { resolveImage, maxShadowPixels: 0 });
  return { png: new Uint8Array(await canvas.encode('png')), page, pageCount: snapshot.slides.length, width, height, warnings: [...warnings, 'Shadow effects are omitted in server previews.'] };
}
