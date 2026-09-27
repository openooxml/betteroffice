import { createCanvas, loadImage, Path2D } from '@napi-rs/canvas';
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
  if (!Number.isFinite(page.width) || !Number.isFinite(page.height) || page.width <= 0 || page.height <= 0 || page.width * page.height > 16_000_000) {
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
      const media = document.package.media;
      const file = media?.get(id) ?? (target ? media?.get(target) ?? media?.get(`word/${target}`) : undefined);
      if (!file) { warnings.add(`Image unavailable: ${id}`); images.set(id, null); return null; }
      try {
        const image = await loadImage(Buffer.from(file.data));
        images.set(id, image);
        return image as unknown as CanvasImageSource;
      } catch { warnings.add(`Image could not be decoded: ${id}`); images.set(id, null); return null; }
    },
  });
  return { png: await canvas.encode('png'), page: pageNumber, pageCount: display.pages.length, width: canvas.width, height: canvas.height, warnings: [...warnings] };
};
