import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { engineKind } from '../wasm/enginePng';
import { preloadEditWasm } from '../wasm/edit';
import { skippedEffectKinds } from '../wasm/engineEdit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession } from './index';
import * as pngExport from '../../../../shared/office-session/pngExportError';
import type { PngDiagnostic } from '../../../../shared/office-session/pngExportError';
import { rezipPartsToArrayBuffer } from '../docx/rezip/parts';
import {
  EFFECT_FIXTURES,
  MIXED_FIXTURE,
  PLAIN_FIXTURE,
  UNDECODABLE_PNG,
  TWO_PICTURE_FIXTURES,
  UNDRAWN_PICTURE_FILL_FIXTURES,
  effectsSkipped,
  expectedDiagnostics,
} from './pngEffectsSkippedTestSupport';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stubBrowser() {
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
    convertToBlob(request: { type: string }) { return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  class Path { moveTo() {} lineTo() {} quadraticCurveTo() {} bezierCurveTo() {} closePath() {} addPath() {} }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Path2D: scope.Path2D, Image: scope.Image, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, Path2D: Path, Image: undefined,
    createImageBitmap: (source: Blob) => source.size <= UNDECODABLE_PNG.byteLength
      ? Promise.reject(new Error('The source image could not be decoded.'))
      : Promise.resolve({ width: 1, height: 1, close: () => undefined }),
  });
  restore = () => { Object.assign(scope, saved); };
}

async function pageDiagnostics(parts: Record<string, Uint8Array>, pages: number[]): Promise<PngDiagnostic[]> {
  stubBrowser();
  const session = await createYrsSession();
  try {
    session.openDocx(new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(parts)))), true);
    const diagnostics: PngDiagnostic[] = [];
    for (const page of pages) await session.renderPng({ page }, { onDiagnostic: (diagnostic) => diagnostics.push(diagnostic) });
    return diagnostics;
  } finally { session.destroy(); }
}

describe('effects-skipped PNG diagnostic', () => {
  test('the shared module lists every kind the diagnostic names', () => {
    const kinds: readonly string[] = pngExport.PNG_SKIPPED_EFFECT_KINDS ?? [];
    expect(skippedEffectKinds()).toEqual([...kinds]);
    const tableKinds = EFFECT_FIXTURES.flatMap((fixture) => [...fixture.skipped.wasm, ...fixture.skipped.native]);
    for (const kind of tableKinds) expect(kinds).toContain(kind);
  });

  test('every fixture lists its kinds in the shared order', () => {
    const kinds: readonly string[] = pngExport.PNG_SKIPPED_EFFECT_KINDS ?? [];
    const lists = [PLAIN_FIXTURE, ...EFFECT_FIXTURES].flatMap((fixture) => [fixture.skipped.wasm, fixture.skipped.native]).concat([MIXED_FIXTURE.skipped]);
    for (const list of lists) {
      const names: readonly string[] = list;
      expect(names).toEqual(kinds.filter((kind) => names.includes(kind)));
    }
  });

  test.each(EFFECT_FIXTURES.map((fixture) => [fixture.name, fixture] as const))('%s reports the kinds this engine does not draw, once', async (_name, fixture) => {
    const diagnostics = await pageDiagnostics(fixture.parts, [0]);
    expect(effectsSkipped(diagnostics)).toEqual(expectedDiagnostics(0, fixture.skipped[engineKind]));
  });

  test.each(UNDRAWN_PICTURE_FILL_FIXTURES.map((fixture) => [fixture.name, fixture] as const))('%s reports only that the picture is unavailable', async (_name, fixture) => {
    const diagnostics = await pageDiagnostics(fixture.parts, [0]);
    expect(diagnostics.map((diagnostic) => diagnostic.code)).toEqual(['picture-unavailable']);
  });

  test.each(TWO_PICTURE_FIXTURES.map((fixture) => [fixture.name, fixture] as const))('of two pictures of one type, %s', async (_name, fixture) => {
    const diagnostics = await pageDiagnostics(fixture.parts, [0]);
    const [unavailable, ...rest] = diagnostics;
    expect(unavailable?.code === 'picture-unavailable' && unavailable.pictures.length).toBe(1);
    expect(rest).toEqual(expectedDiagnostics(0, fixture.skipped));
  });

  test('a page without effects reports none', async () => {
    expect(effectsSkipped(await pageDiagnostics(PLAIN_FIXTURE.parts, [0]))).toEqual([]);
  });

  test('a plain page and an effect page report once, for the effect page, by its zero-based index', async () => {
    const diagnostics = await pageDiagnostics(MIXED_FIXTURE.parts, [0, 1]);
    expect(effectsSkipped(diagnostics)).toEqual(expectedDiagnostics(MIXED_FIXTURE.effectPage, MIXED_FIXTURE.skipped));
  });

  test('the page a diagnostic names is the target page it was rendered for', async () => {
    const [diagnostic] = effectsSkipped(await pageDiagnostics(MIXED_FIXTURE.parts, [1]));
    expect(diagnostic?.page).toBe(1);
  });
});
