import { engines, loadCanvas } from '#agents-engines';
import { imageSize } from 'image-size';
import { createFontProvider } from '@betteroffice/fonts';
import type { YrsSession } from '@betteroffice/docx/yrs';
import type { ResidentFontRequirement } from '@betteroffice/docx/layout';
import { cleanupGuard } from './cleanup';
import { DocumentToolError, type DocumentRenderer } from './types';

const previewWarning = 'Preview uses BetterOffice layout and bundled substitute fonts; compare with the target Office application when exact fidelity matters.';

export const renderDocxPage: DocumentRenderer = async (session, pageNumber) => {
  try {
    const png = await session.renderPng({ page: pageNumber - 1 });
    const prepared = await layoutDocxPage(session);
    const release = cleanupGuard(prepared.dispose);
    try {
      const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
      return { png, page: pageNumber, pageCount: prepared.layout.pages.length, width: view.getUint32(16), height: view.getUint32(20), warnings: [previewWarning, ...(!prepared.notesConverged ? ['Footnote pagination did not converge.'] : [])] };
    } catch (error) { throw release.failed(error); } finally { release.run(); }
  } catch (error) {
    const code = (error as { code?: unknown } | null)?.code;
    if (code === 'canvas-unavailable') return renderDocxPageCanvas(session, pageNumber);
    if (code === 'canvas-too-large') throw new DocumentToolError('PAGE_TOO_LARGE', error instanceof Error ? error.message : 'Page exceeds the preview render limit.');
    if (code === 'font-unavailable') throw new DocumentToolError('FONT_UNAVAILABLE', error instanceof Error ? error.message : 'Required fonts could not be loaded.');
    const range = error instanceof RangeError && error.message.match(/^page index .* out of range for (\d+) page\(s\)$/);
    if (range) throw new DocumentToolError('PAGE_OUT_OF_RANGE', `Choose a page from 1 to ${range[1]}.`, { pageCount: Number(range[1]) });
    if (code === 'page-out-of-range' || error instanceof TypeError && error.message === 'page must be a non-negative integer') throw new DocumentToolError('PAGE_OUT_OF_RANGE', 'Choose a one-based page index.');
    throw error;
  }
};

/** Lays out a private copy, so a caller's session keeps its layout and display state. */
async function layoutDocxPage(original: YrsSession) {
  const { module } = await engines.load('docx');
  const sourceDocument = original.materializeDocx();
  if (!sourceDocument) throw new DocumentToolError('UNSUPPORTED', 'Rendering requires a session opened from DOCX bytes.');
  const session = await module.yrs.createYrsSession();
  try {
    if (sourceDocument.originalBuffer) {
      const bytes = new Uint8Array(sourceDocument.originalBuffer);
      const host = session.openDocx(bytes, false);
      session.bootstrapPeer(original.encodeState(), original.encodePeerMetadata(), bytes, host);
    } else {
      session.loadState(original.encodeState());
    }
  } catch (error) { throw cleanupGuard(() => session.destroy()).failed(error); }
  const dispose = () => session.destroy();
  try {
    const document = session.materializeDocx()!;
    const bundled = module.layout.hasConfiguredFontSource()
      ? await module.layout.resolveDefaultFontProvider()
      : createFontProvider(undefined, { decoder: module.yrs.decodeFontFace });
    const source = module.layout.createRustMeasureSource({ engine: session, bundled });
    source.setCompat(document.package.settings?.compatibilityFlags);
    source.setEmbeddedFaces(await module.utils.extractEmbeddedFontFaces(document));
    const renderEnv = module.editor.documentRenderEnv(document);
    const request = module.editor.buildResidentRegionLayoutRequest(document, 0, renderEnv);
    const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[];
    await source.prepareFontRequirements(requirements);
    const measurement = source.measurementConfigForRequirements(requirements);
    if (!measurement) throw new DocumentToolError('FONT_UNAVAILABLE', 'Required fonts could not be loaded. Install @betteroffice/fonts and the optional coverage packages @betteroffice/fonts-cjk and @betteroffice/fonts-intl as needed.');
    const { layout, notesConverged } = module.editor.computeLayout({ document, pageGap: 0, session, renderEnv, measurement });
    return { session, module, document, measurement, layout, notesConverged, dispose };
  } catch (error) { throw cleanupGuard(dispose).failed(error); }
}

const renderDocxPageCanvas: DocumentRenderer = async (original, pageNumber) => {
  const { createCanvas, loadImage, Path2D } = await loadCanvas();
  const prepared = await layoutDocxPage(original);
  const { session, module, document, measurement, notesConverged } = prepared;
  const { openPagedDisplay, drawDisplayPage, GlyphCache } = module.render;
  const release = cleanupGuard(prepared.dispose);
  try {
    const display = openPagedDisplay(session, measurement.fontChains);
    const page = display.page(pageNumber - 1);
    if (!page) throw new DocumentToolError('PAGE_OUT_OF_RANGE', `Choose a page from 1 to ${display.pageCount}.`, { pageCount: display.pageCount });
    if (!Number.isFinite(page.width) || !Number.isFinite(page.height) || page.width <= 0 || page.height <= 0 || Math.ceil(page.width) * Math.ceil(page.height) > 16_000_000) {
      throw new DocumentToolError('PAGE_TOO_LARGE', 'Page exceeds the 16 megapixel render limit.');
    }
    const canvas = createCanvas(Math.ceil(page.width), Math.ceil(page.height));
    const warnings = new Set<string>([previewWarning]);
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
    return { png: await canvas.encode('png'), page: pageNumber, pageCount: display.pageCount, width: canvas.width, height: canvas.height, warnings: [...warnings] };
  } catch (error) { throw release.failed(error); } finally { release.run(); }
};
