import { afterEach, expect, test } from 'bun:test';
import * as fontManager from '@betteroffice/fonts';
import { configureDefaultFonts } from './defaultFontProvider';
import { createRustMeasureSource } from './rustMeasureSource';
import * as wasm from '../wasm';

afterEach(() => configureDefaultFonts({}));

async function prepare(configured = false) {
  await wasm.preloadLayoutWasm();
  if (configured) configureDefaultFonts({ fonts: fontManager });
  const registered: Uint8Array[] = [];
  const source = createRustMeasureSource({
    engine: {
      registerFont(bytes) { registered.push(bytes); return wasm.registerMeasureFont(bytes); },
      registerSubstituteFont: wasm.registerSubstituteMeasureFont,
      clearFonts() {},
    },
  });
  const requirements = ['Noto Sans Devanagari', 'Calibri'].map((family) => ({
    key: `${family.toLowerCase()}|0|0`, family, bold: false, italic: false,
  }));
  await source.prepareFontRequirements(requirements);
  return { config: source.measurementConfigForRequirements(requirements)!, registered };
}

function measure(fontChains: Record<string, number[]>, family: string, text: string) {
  try {
    return JSON.parse(wasm.measureParagraphJson(JSON.stringify({
      block: { kind: 'paragraph', runs: [{ kind: 'text', text, fontFamily: family }] },
      maxWidth: 400,
      fontChains,
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      authoritativeShaping: true,
    })));
  } catch (error) {
    if (String(error).includes('UNSUPPORTED')) return 'browser';
    throw error;
  }
}

function paint(fontChains: Record<string, number[]>, family: string, text: string) {
  const display = JSON.parse(wasm.buildDisplayListJson(JSON.stringify({
    measured: [{
      block: {
        kind: 'paragraph', id: 'coverage',
        runs: [{ kind: 'text', text, fontFamily: family, fontSize: 11 }],
      },
      measure: {
        kind: 'paragraph', totalHeight: 20,
        lines: [{
          headRun: 0, headChar: 0, tailRun: 0, tailChar: text.length,
          width: 100, ascent: 14, descent: 4, lineHeight: 20, syntheticFallback: true,
        }],
      },
    }],
    fontChains,
    options: {},
    layout: { pages: [{
      size: { w: 816, h: 1056 }, margins: { top: 96, right: 96, bottom: 96, left: 96 },
      number: 1,
      fragments: [{
        kind: 'paragraph', blockId: 'coverage', x: 96, y: 96, width: 624, height: 20,
        fromLine: 0, toLine: 1,
      }],
    }] },
  })));
  return display.pages[0].primitives.filter((primitive: { kind: string }) =>
    primitive.kind === 'text' || primitive.kind === 'glyphRun');
}

test('implicit fonts leave uncovered Devanagari to the no-provider path while Latin uses Carlito', async () => {
  const { config, registered } = await prepare();
  for (const family of ['Noto Sans Devanagari', 'Calibri']) {
    for (const text of ['\u0939', 'Latin \u0939']) {
      const main = measure({}, family, text);
      expect(main).toBe('browser');
      expect(measure(config.fontChains, family, text)).toEqual(main);
      const browser = paint({}, family, text);
      expect(browser.length).toBeGreaterThan(0);
      expect(paint(config.fontChains, family, text)).toEqual(browser);
    }
  }
  expect(measure(config.fontChains, 'Calibri', 'Latin text').kind).toBe('paragraph');
  const carlito = await fontManager.createFontProvider().resolve('Calibri', false, false)!();
  expect(registered).toContainEqual(new Uint8Array(carlito));
  expect(paint(config.fontChains, 'Calibri', 'Latin text').some(
    (primitive: { kind: string }) => primitive.kind === 'glyphRun',
  )).toBe(true);
});

test('an explicitly configured manager keeps authoritative missing-glyph shaping', async () => {
  const { config } = await prepare(true);
  for (const family of ['Noto Sans Devanagari', 'Calibri']) {
    expect(measure(config.fontChains, family, '\u0939').kind).toBe('paragraph');
    expect(paint(config.fontChains, family, '\u0939').some(
      (primitive: { kind: string }) => primitive.kind === 'glyphRun',
    )).toBe(true);
  }
});
