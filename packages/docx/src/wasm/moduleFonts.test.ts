import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createEditSession, preloadEditWasm } from './edit';
import {
  build_display_list_json,
  clear_measure_fonts,
  layout_document_json,
  measure_paragraph_json,
  outline_glyph_json,
  register_measure_font,
  register_substitute_measure_font,
} from './generated/edit/docx_edit.js';

const font = new Uint8Array(readFileSync(resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
)));
const input = (id: number) => JSON.stringify({
  block: { kind: 'paragraph', id: 1, runs: [{ kind: 'text', text: 'Font stores' }] },
  maxWidth: 400,
  fontChains: { 'liberation sans|0|0': [id] },
  defaults: { fontSize: 12, fontFamily: 'Liberation Sans' },
});

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, 'generated/edit/docx_edit_bg.wasm')
))), 30_000);

test('module font exports coexist with an edit session in the same wasm instance', () => {
  clear_measure_fonts();
  const session = createEditSession(7);
  try {
    expect(session.register_measure_font(font)).toBe(0);
    const sessionMeasurement = session.measure_paragraph_json(input(0));
    const sessionOutline = session.outline_glyph_json(0, 36);

    expect(register_measure_font(font)).toBe(0);
    const substitute = register_substitute_measure_font(0, 'Segoe UI Light');
    expect(substitute).toBe(1);
    expect(JSON.parse(measure_paragraph_json(input(substitute))).lines.length).toBeGreaterThan(0);
    expect(JSON.parse(outline_glyph_json(substitute, 36)).cmds.length).toBeGreaterThan(0);
    const options = {
      pageSize: { w: 816, h: 1056 },
      margins: { top: 96, right: 96, bottom: 96, left: 96 },
      pageGap: 20,
    };
    const layout = JSON.parse(layout_document_json(JSON.stringify({ measured: [], options })));
    expect(() => build_display_list_json(JSON.stringify({ measured: [], options, layout })))
      .not.toThrow();

    clear_measure_fonts();
    expect(session.measure_paragraph_json(input(0))).toBe(sessionMeasurement);
    expect(session.outline_glyph_json(0, 36)).toBe(sessionOutline);
    expect(session.register_measure_font(font)).toBe(1);

    expect(register_measure_font(font)).toBe(0);
    const moduleMeasurement = measure_paragraph_json(input(0));
    const moduleOutline = outline_glyph_json(0, 36);
    expect(session.register_substitute_measure_font(0, 'Segoe UI Light')).toBe(2);
    expect(measure_paragraph_json(input(0))).toBe(moduleMeasurement);
    expect(outline_glyph_json(0, 36)).toBe(moduleOutline);
    session.clear_measure_fonts();
    expect(measure_paragraph_json(input(0))).toBe(moduleMeasurement);
    expect(outline_glyph_json(0, 36)).toBe(moduleOutline);
    expect(register_measure_font(font)).toBe(1);
    expect(session.register_measure_font(font)).toBe(0);
  } finally {
    session.free();
    clear_measure_fonts();
  }
});
