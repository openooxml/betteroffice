import { expect, test } from 'bun:test';
import type { PageDisplayList } from '@betteroffice/vsdx';
import { collectDiagnostics } from './VsdxEditor';

const frame: PageDisplayList = {
  contractVersion: 4,
  width: 1,
  height: 1,
  paintTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
  primitives: [{ kind: 'textBox', id: 'text', zOrder: 0, x: 0, y: 0, width: 1, height: 1, paragraphs: [{ runs: [{ text: 'x', family: 'Arial', sizeIn: 12, bold: false, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#000', diagnostics: [{ category: 'integrity', code: 'missing-media', detail: '' }, { category: 'fidelity', code: 'font-substituted', detail: '' }] }]}], lines: [] }],
};

test('collects structured diagnostics without matching their text', () => {
  expect(collectDiagnostics(frame)).toEqual([
    { category: 'integrity', code: 'missing-media', detail: '' },
    { category: 'fidelity', code: 'font-substituted', detail: '' },
  ]);
});
