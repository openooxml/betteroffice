import { readFileSync } from 'node:fs';
import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import type { DisplayList, ImagePrimitive } from '../../packages/docx/src/layout/render/displayList';
import { PNG_SKIPPED_EFFECT_KINDS, type PngDiagnostic } from '../../shared/office-session/pngExportError';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { skippedEffectKinds as canvasEngineKinds } from '../../packages/docx/src/wasm/engineEdit';
import { engineKind as canvasModuleKind, renderPng as renderCanvasPng, type DocxPngEngine } from '../../packages/docx/src/wasm/enginePng';
import { renderPng as renderNativePng } from './js/enginePng';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { pictureDocument } from '../../e2e/docx-png-pictures';
import { rezipPartsToArrayBuffer } from '../../packages/docx/src/docx/rezip/parts';
import {
  EFFECT_FIXTURES,
  MIXED_FIXTURE,
  ONE_PIXEL_PNG,
  PICTURE_BODY,
  PLAIN_FIXTURE,
  UNDECODABLE_PNG,
  UNDRAWN_PICTURE_FILL_FIXTURES,
  effectDocx,
  effectsSkipped,
  expectedDiagnostics,
  type EffectFixture,
} from '../../packages/docx/src/yrs/pngEffectsSkippedTestSupport';

type Session = import('./index.js').EditSession;

const raw = require('./index.js') as typeof import('./index.js');
const fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));

/** Under the native engine preload the canvas module is swapped for the native one, so only the native half runs. */
const canvasEngine = canvasModuleKind === 'wasm';

/** Header parts only reach the page through session layout, which the session-level PNG test covers on this engine. */
const bodyFixtures = EFFECT_FIXTURES.filter((fixture) => !('word/header1.xml' in fixture.parts));

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stubCanvas() {
  class Canvas {
    constructor(public width: number, public height: number) {}
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          if (property === 'createLinearGradient' || property === 'createRadialGradient' || property === 'createPattern') return () => ({ addColorStop: () => undefined });
          return property in target ? target[property as string] : () => undefined;
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob() { return Promise.resolve(new Blob(['png'])); }
  }
  class Path { moveTo() {} lineTo() {} quadraticCurveTo() {} bezierCurveTo() {} closePath() {} addPath() {} }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Path2D: scope.Path2D, Image: scope.Image, document: scope.document, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, Path2D: Path, Image: undefined, document: undefined,
    createImageBitmap: (source: Blob) => source.size <= UNDECODABLE_PNG.byteLength
      ? Promise.reject(new Error('The source image could not be decoded.'))
      : Promise.resolve({ width: 1, height: 1, close() {} }),
  });
  restore = () => { Object.assign(scope, saved); };
}

function laidOut(session: Session): { list: DisplayList; fontChains: Record<string, number[]> } {
  const font = session.register_measure_font(fontBytes);
  const fontChains = { 'calibri|0|0': [font], 'calibri|1|0': [font] };
  const materialized = JSON.parse(session.materialize_docx()!) as {
    document: { package: { theme: { colorScheme: Record<string, string> } } };
  };
  const layout = session.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{
      sectionId: 'main',
      properties: {
        pageWidth: 4320, pageHeight: 2880,
        marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
      },
    }] },
    measurement: {
      fontChains,
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      authoritativeShaping: true,
    },
    renderEnv: { themeColors: materialized.document.package.theme.colorScheme },
  }));
  const list = JSON.parse(session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }))) as DisplayList;
  return { list, fontChains };
}

function pageEngine(session: Session, list: DisplayList, fontChains: Record<string, number[]>, page: number): DocxPngEngine {
  return {
    displayList: list,
    fontChains,
    page: list.pages[page]!,
    session,
    outlineGlyph: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),
    media: () => ({ bytes: ONE_PIXEL_PNG, mimeType: 'image/png' }),
  };
}

type Render = typeof renderNativePng;

async function diagnosticsFor(render: Render, parts: Record<string, Uint8Array>, pages: number[], seed: number): Promise<PngDiagnostic[]> {
  const session = new raw.EditSession(seed);
  try {
    session.open_docx(raw.rezip_docx(parts), true);
    const { list, fontChains } = laidOut(session);
    const diagnostics: PngDiagnostic[] = [];
    for (const page of pages) await render(pageEngine(session, list, fontChains, page), { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    return diagnostics;
  } finally {
    session.free();
  }
}

/** Lays out a one-page docx with `body`, lets `change` edit page 0's display list, and renders that page. */
async function withPage(body: string, seed: number, change: (page: DisplayList['pages'][number]) => void, render: Render): Promise<PngDiagnostic[]> {
  const session = new raw.EditSession(seed);
  try {
    session.open_docx(raw.rezip_docx(effectDocx({ body })), true);
    const { list, fontChains } = laidOut(session);
    const engine = pageEngine(session, list, fontChains, 0);
    change(engine.page);
    const diagnostics: PngDiagnostic[] = [];
    await render(engine, { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    return diagnostics;
  } finally {
    session.free();
  }
}

/** The kinds the `effects-skipped` diagnostics among `diagnostics` name. */
function reportedKinds(diagnostics: PngDiagnostic[]): Set<string> {
  return new Set(effectsSkipped(diagnostics).flatMap((diagnostic) => diagnostic.effects));
}

const fixtureCases = (fixtures: EffectFixture[]) => fixtures.map((fixture) => [fixture.name, fixture] as const);

describe('native docx PNG effects-skipped diagnostic', () => {
  test.each(fixtureCases(bodyFixtures))('%s reports the kinds the native rasterizer does not draw, once', async (_name, fixture) => {
    const diagnostics = await diagnosticsFor(renderNativePng, fixture.parts, [0], 431);
    expect(effectsSkipped(diagnostics)).toEqual(expectedDiagnostics(0, fixture.skipped.native));
  });

  test('a page without effects reports none', async () => {
    expect(effectsSkipped(await diagnosticsFor(renderNativePng, PLAIN_FIXTURE.parts, [0], 432))).toEqual([]);
  });

  test('a plain page and an effect page report once, for the effect page, by its zero-based index', async () => {
    const diagnostics = await diagnosticsFor(renderNativePng, MIXED_FIXTURE.parts, [0, 1], 433);
    expect(effectsSkipped(diagnostics)).toEqual(expectedDiagnostics(MIXED_FIXTURE.effectPage, MIXED_FIXTURE.skipped));
  });

  test('a picture with a filter reports image.filter and still exports', async () => {
    let png: Uint8Array | undefined;
    const render: Render = async (engine, options) => (png = await renderNativePng(engine, options));
    const diagnostics = await withPage(PICTURE_BODY, 437, (page) => {
      const images = page.primitives.filter((primitive): primitive is ImagePrimitive => primitive.kind === 'image');
      expect(images).toHaveLength(1);
      for (const image of images) image.filter = 'brightness(1.4) contrast(0.4)';
    }, render);
    expect(png?.byteLength).toBeGreaterThan(0);
    expect(diagnostics).toEqual([{ code: 'effects-skipped', page: 0, effects: ['image.filter'] }]);
  });

  test.each(UNDRAWN_PICTURE_FILL_FIXTURES.map((fixture) => [fixture.name, fixture] as const))('%s reports only that the picture is unavailable', async (_name, fixture) => {
    const session = await nativeYrs.createYrsSession({ clientId: 438 });
    try {
      session.openDocx(new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(fixture.parts)))), true);
      const diagnostics: PngDiagnostic[] = [];
      await session.renderPng({ page: 0 }, { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
      expect(diagnostics.map((diagnostic) => diagnostic.code)).toEqual(['picture-unavailable']);
    } finally {
      session.destroy();
    }
  });

  test('a picture whose part is missing reports only that the picture is unavailable', async () => {
    const session = await nativeYrs.createYrsSession({ clientId: 434 });
    try {
      session.openDocx(await pictureDocument('missing'), true);
      const diagnostics: PngDiagnostic[] = [];
      await session.renderPng({ page: 0 }, { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
      expect(diagnostics.map((diagnostic) => diagnostic.code)).toEqual(['picture-unavailable']);
    } finally {
      session.destroy();
    }
  });
});

describe('both docx PNG engines', () => {
  test('name exactly the shared kinds, in the shared order', () => {
    expect(raw.skippedEffectKinds()).toEqual([...PNG_SKIPPED_EFFECT_KINDS]);
    expect(canvasEngineKinds()).toEqual([...PNG_SKIPPED_EFFECT_KINDS]);
  });

  test('tell two pictures of one type apart on the native engine', async () => {
    const dataUrl = (bytes: Uint8Array) => `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`;
    const diagnostics = await withPage(PICTURE_BODY, 442, (page) => {
      const image = page.primitives.find((primitive): primitive is ImagePrimitive => primitive.kind === 'image')!;
      const broken = { ...image, relId: dataUrl(UNDECODABLE_PNG) };
      const filtered = { ...image, relId: dataUrl(ONE_PIXEL_PNG), filter: 'url(#washout)' };
      page.primitives.push(broken, filtered);
    }, renderNativePng);
    expect(diagnostics.map((diagnostic) => diagnostic.code)).toEqual(['picture-unavailable', 'effects-skipped']);
    const [unavailable, skipped] = diagnostics;
    expect(unavailable?.code === 'picture-unavailable' && unavailable.pictures.map((picture) => picture.id)).toEqual(['data:image/png;base64']);
    expect(skipped?.code === 'effects-skipped' && skipped.effects).toEqual(['image.filter']);
  });
});

describe('both docx PNG engines report the kinds of their own table', () => {
  test.each(fixtureCases(bodyFixtures))('%s', async (_name, fixture) => {
    const native = reportedKinds(await diagnosticsFor(renderNativePng, fixture.parts, [0], 435));
    expect(native).toEqual(new Set(fixture.skipped.native));
    if (!canvasEngine) return;
    stubCanvas();
    const canvas = reportedKinds(await diagnosticsFor(renderCanvasPng, fixture.parts, [0], 436));
    expect(canvas).toEqual(new Set(fixture.skipped.wasm));
  });
});
