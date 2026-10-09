import { createCanvas, loadImage, Path2D } from '@napi-rs/canvas';
import { imageSize } from 'image-size';
import { createFontProvider } from '@betteroffice/fonts';
import { buildResidentRegionLayoutRequest, computeLayout, getLayoutKernelInputs } from '@betteroffice/docx/editor';
import { createRustMeasureSource, type ResidentFontRequirement } from '@betteroffice/docx/layout';
import { buildRustDisplayList, drawDisplayPage, GlyphCache } from '@betteroffice/docx/layout/render';
import { DocumentToolError, type DocumentRenderer } from './types';

export const renderDocxPage: DocumentRenderer = async (session, pageNumber) => {
  const document = session.materializeDocx();
  if (!document) throw new DocumentToolError('UNSUPPORTED', 'Rendering requires a session opened from DOCX bytes.');
  const source = createRustMeasureSource({ engine: session, bundled: createFontProvider() });
  const request = buildResidentRegionLayoutRequest(document, 0, {});
  const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[];
  await source.prepareFontRequirements(requirements);
  const measurement = source.measurementConfigForRequirements(requirements);
  if (!measurement) throw new DocumentToolError('FONT_UNAVAILABLE', 'Required fonts could not be loaded. Install @betteroffice/fonts and, for CJK, @betteroffice/fonts-cjk.');
  const { layout, notesConverged } = computeLayout({ document, pageGap: 0, session, renderEnv: {}, measurement });
  const kernel = getLayoutKernelInputs(layout)!;
  const display = await buildRustDisplayList({ ...kernel, layout, fontChains: measurement.fontChains }, session);
  const page = display.pages[pageNumber - 1];
  if (!page) throw new DocumentToolError('PAGE_OUT_OF_RANGE', `Choose a page from 1 to ${display.pages.length}.`, { pageCount: display.pages.length });
  if (!Number.isFinite(page.width) || !Number.isFinite(page.height) || page.width <= 0 || page.height <= 0 || Math.ceil(page.width) * Math.ceil(page.height) > 16_000_000) {
    throw new DocumentToolError('PAGE_TOO_LARGE', 'Page exceeds the 16 megapixel render limit.');
  }
  const canvas = createCanvas(Math.ceil(page.width), Math.ceil(page.height));
  const warnings = new Set<string>(['Preview uses BetterOffice layout and bundled substitute fonts; compare with the target Office application when exact fidelity matters.']);
  if (!notesConverged) warnings.add('Footnote pagination did not converge.');
  const primitives = [...page.primitives, ...(page.header?.primitives ?? []), ...(page.footer?.primitives ?? []), ...(page.noteAreas ?? []).flatMap(area => [...(area.separatorPrimitives ?? []), ...(area.primitives ?? [])])];
  if (primitives.some(primitive => primitive.kind === 'glyphRun' && primitive.glyphs.some(glyph => glyph.id === 0))) {
    warnings.add('Some characters have no glyph in the available fonts. Their preview may show replacement boxes.');
  }
  const images = new Map<string, Awaited<ReturnType<typeof loadImage>> | null>();
  let decodedPixels = 0;
  const media = document.package.media;
  const mediaBySource = new Map([...media?.values() ?? []].map(file => [file.dataUrl, file]));
  const glyphCache = new GlyphCache({
    provider: (font, glyph) => session.outlineGlyphJson(font, glyph),
    createPath: () => new Path2D() as unknown as globalThis.Path2D,
  });
  await drawDisplayPage(canvas.getContext('2d') as unknown as CanvasRenderingContext2D, { ...page, background: page.background ?? '#ffffff' }, {
    glyphCache,
    resolveImage: async (id) => {
      if (images.has(id)) return images.get(id) as unknown as CanvasImageSource | null;
      const relationship = document.package.relationships?.get(id);
      const target = relationship?.target.replace(/^\//, '').replace(/^\.\//, '');
      const file = mediaBySource.get(id) ?? media?.get(id) ?? (target ? media?.get(target) ?? media?.get(`word/${target}`) : undefined);
      if (!file) { warnings.add('Image unavailable in the embedded document media.'); images.set(id, null); return null; }
      try {
        if (file.data.byteLength > 16 * 1024 * 1024) throw new DocumentToolError('IMAGE_LIMIT', 'An embedded image exceeds the preview image limits.');
        const size = imageSize(new Uint8Array(file.data));
        const pixels = size.width * size.height;
        if (!['png', 'jpg', 'gif', 'webp', 'bmp', 'ico'].includes(size.type ?? '')) {
          throw new DocumentToolError('UNSUPPORTED_IMAGE', 'An embedded image uses an unsupported encoding; previews support raster images only.');
        }
        if (!Number.isSafeInteger(pixels) || pixels <= 0 || pixels > 16_000_000 || decodedPixels + pixels > 64_000_000) {
          throw new DocumentToolError('IMAGE_LIMIT', 'An embedded image exceeds the preview image limits.');
        }
        decodedPixels += pixels;
        const image = await loadImage(Buffer.from(file.data));
        images.set(id, image);
        return image as unknown as CanvasImageSource;
      } catch (error) { warnings.add(error instanceof DocumentToolError ? error.message : 'An embedded image could not be decoded.'); images.set(id, null); return null; }
    },
  });
  return { png: await canvas.encode('png'), page: pageNumber, pageCount: display.pages.length, width: canvas.width, height: canvas.height, warnings: [...warnings] };
};
