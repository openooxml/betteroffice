import { expect, test } from 'bun:test';
import { createEditSession } from './edit';
import {
  build_display_list_json,
  clear_measure_fonts,
  measure_paragraph_json,
  outline_glyph_json,
  register_measure_font,
  register_substitute_measure_font,
} from './generated/edit/docx_edit.js';

function thrown(call: () => unknown): string {
  try {
    call();
  } catch (error) {
    return String(error);
  }
  return 'returned';
}

test('module-level font helpers throw instead of trapping while an edit session is open', () => {
  const session = createEditSession(7);
  try {
    for (const call of [
      () => clear_measure_fonts(),
      () => register_measure_font(new Uint8Array([0, 1, 0, 0])),
      () => register_substitute_measure_font(0, 'Calibri'),
      () => measure_paragraph_json('{}'),
      () => outline_glyph_json(0, 0),
      () => build_display_list_json('{}'),
    ]) {
      expect(thrown(call)).toContain('measurement fonts belong to an open editing session');
    }
    expect(thrown(() => session.clear_measure_fonts())).toBe('returned');
  } finally {
    session.free();
  }
  expect(thrown(() => clear_measure_fonts())).toBe('returned');
});
