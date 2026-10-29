import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { getRustTextEngine } from '../layout/measure/rustMeasureSource';
import { measureParagraphJson, outlineGlyphJson } from '../layout/wasm/index';
import { preloadEditWasm } from './edit';
import { createYrsSession } from '../yrs/index';

const font = new Uint8Array(readFileSync(resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
const input = (id: number) => JSON.stringify({
  block: { kind: 'paragraph', id: 1, runs: [{ kind: 'text', text: 'Shared engine' }] },
  maxWidth: 400,
  fontChains: { 'liberation sans|0|0': [id] },
  defaults: { fontSize: 12, fontFamily: 'Liberation Sans' },
});

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

test('the public text engine and Yrs session keep independent fonts and measurements', async () => {
  const session = await createYrsSession({ clientId: 76301 });
  const textEngine = await getRustTextEngine();
  try {
    textEngine.clearFonts();
    expect(session.registerFont(font)).toBe(0);
    const sessionMeasurement = session.measureParagraphJson(input(0));
    const sessionOutline = session.outlineGlyphJson(0, 36);

    expect(textEngine.registerFont(font)).toBe(0);
    const substitute = textEngine.registerSubstituteFont!(0, 'Segoe UI Light');
    expect(substitute).toBe(1);
    expect(JSON.parse(measureParagraphJson(input(substitute))).lines.length).toBeGreaterThan(0);
    textEngine.clearFonts();
    expect(session.measureParagraphJson(input(0))).toBe(sessionMeasurement);
    expect(session.outlineGlyphJson(0, 36)).toBe(sessionOutline);
    expect(session.registerFont(font)).toBe(1);

    expect(textEngine.registerFont(font)).toBe(0);
    const moduleMeasurement = measureParagraphJson(input(0));
    const moduleOutline = outlineGlyphJson(0, 36);
    expect(session.registerSubstituteFont(0, 'Segoe UI Light')).toBe(2);
    expect(measureParagraphJson(input(0))).toBe(moduleMeasurement);
    expect(outlineGlyphJson(0, 36)).toBe(moduleOutline);
    session.clearFonts();
    expect(measureParagraphJson(input(0))).toBe(moduleMeasurement);
    expect(outlineGlyphJson(0, 36)).toBe(moduleOutline);
    expect(textEngine.registerFont(font)).toBe(1);
    expect(session.registerFont(font)).toBe(0);
  } finally {
    session.destroy();
    textEngine.clearFonts();
  }
});
