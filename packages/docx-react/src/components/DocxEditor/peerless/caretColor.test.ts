import { expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from '@betteroffice/docx/layout/render';
import { peerlessCaretColor } from './caretColor';

const caret = { x: 130, y: 100, width: 1, height: 20 };

function run(color: string, left: number, right: number, extra: Partial<DisplayPrimitive> = {}): DisplayPrimitive {
  return { kind: 'glyphRun', fontId: 1, size: 16, color, text: 'ab',
    glyphs: [{ id: 1, x: left, y: 115, cluster: 0, advance: (right - left) / 2 }, { id: 2, x: (left + right) / 2, y: 115, cluster: 1, advance: (right - left) / 2 }],
    ...extra } as DisplayPrimitive;
}

function page(...primitives: DisplayPrimitive[]): DisplayPage {
  return { pageIndex: 0, width: 816, height: 1056, primitives } as unknown as DisplayPage;
}

test('caret takes the colour of the run it sits in', () => {
  expect(peerlessCaretColor(page(run('#c00000', 100, 160)), caret)).toBe('#c00000');
});

test('caret at a run boundary follows the run before it', () => {
  expect(peerlessCaretColor(page(run('#0000ff', 100, 130), run('#00aa00', 130, 160)), caret)).toBe('#0000ff');
});

test('caret at the start of a line follows the run after it', () => {
  expect(peerlessCaretColor(page(run('#0000ff', 130, 160)), caret)).toBe('#0000ff');
});

test('white text on dark cell shading gives a white caret', () => {
  const shading = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#1f1f1f' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(shading, run('#ffffff', 100, 160)), caret)).toBe('#ffffff');
});

test('white text on a dark highlight gives a white caret', () => {
  const highlight = { kind: 'decoration', deco: 'highlight', x: 100, y: 98, w: 60, h: 22, color: '#000080' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(highlight, run('#ffffff', 100, 160)), caret)).toBe('#ffffff');
});

test('text that would vanish against its fill falls back to a contrasting caret', () => {
  expect(peerlessCaretColor(page(run('#ffffff', 100, 160)), caret)).toBe('#000');
  const shading = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#101010' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(shading, run('#000000', 100, 160)), caret)).toBe('#fff');
});

test('an empty line in a dark cell gets a white caret, an empty page line a black one', () => {
  const shading = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#1f1f1f' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(shading), caret)).toBe('#fff');
  expect(peerlessCaretColor(page(), caret)).toBe('#000');
  expect(peerlessCaretColor(undefined, caret)).toBe('#000');
});

test('runs on other lines and list markers do not colour the caret', () => {
  const below = { ...run('#ff0000', 100, 160), glyphs: [{ id: 1, x: 100, y: 160, cluster: 0, advance: 60 }] } as DisplayPrimitive;
  expect(peerlessCaretColor(page(below, run('#ff0000', 100, 160, { listMarker: true })), caret)).toBe('#000');
});
