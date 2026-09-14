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

test('collects defaulted paint diagnostics and tolerates shapes without them', () => {
  const shapes: PageDisplayList = {
    ...frame,
    primitives: [
      { kind: 'shape', id: 'resolved', zOrder: 0, path: [] },
      { kind: 'group', id: 'group', zOrder: 1, primitives: [{ kind: 'shape', id: 'defaulted', zOrder: 2, path: [], diagnostics: [{ category: 'fidelity', code: 'unresolvable-fill-colour', detail: 'unresolvable fill colour: missing colour cell FillForegnd' }] }] },
    ],
  };
  expect(collectDiagnostics(shapes)).toEqual([
    { category: 'fidelity', code: 'unresolvable-fill-colour', detail: 'unresolvable fill colour: missing colour cell FillForegnd' },
  ]);
});

function textBox(frame: PageDisplayList) {
  const box = frame.primitives[0];
  if (box.kind !== 'textBox') throw new Error('fixture must start with a text box');
  return box;
}

function withoutDiagnostics(): PageDisplayList {
  const { diagnostics: _omitted, ...run } = textBox(frame).paragraphs[0].runs[0];
  return { ...frame, primitives: [{ ...textBox(frame), paragraphs: [{ runs: [run] }] }] };
}

test('tolerates text runs that omit diagnostics', () => {
  expect(collectDiagnostics(withoutDiagnostics())).toEqual([]);
});

test('collects diagnostics alongside runs that omit them', () => {
  const mixed = withoutDiagnostics();
  textBox(mixed).paragraphs[0].runs.push({ ...textBox(frame).paragraphs[0].runs[0], diagnostics: [{ category: 'fidelity', code: 'font-substituted', detail: '' }] });
  expect(collectDiagnostics(mixed)).toEqual([{ category: 'fidelity', code: 'font-substituted', detail: '' }]);
});
