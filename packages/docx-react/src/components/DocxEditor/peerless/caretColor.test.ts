import { expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from '@betteroffice/docx/layout/render';
import { peerlessCaretColor, peerlessCaretPaint } from './caretColor';

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

test('caret just past a dark highlight reads against the page, not the highlight', () => {
  const highlight = { kind: 'decoration', deco: 'highlight', x: 100, y: 98, w: 30, h: 22, color: '#000080' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(highlight, run('#ffffff', 100, 130)), caret)).toBe('#000');
});

test('header and footer shading and text colour the caret', () => {
  const shading = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#1f1f1f' } as DisplayPrimitive;
  const withHeader = { ...page(), header: { rId: 'rId1', kind: 'header', y: 0, height: 200, primitives: [shading, run('#ffffff', 100, 160)] } } as DisplayPage;
  expect(peerlessCaretColor(withHeader, caret)).toBe('#ffffff');
  const withFooter = { ...page(), footer: { rId: 'rId2', kind: 'footer', y: 0, height: 200, primitives: [shading] } } as DisplayPage;
  expect(peerlessCaretColor(withFooter, caret)).toBe('#fff');
});

test('a dark page background gives a light caret', () => {
  expect(peerlessCaretColor({ ...page(), background: '#101010' } as DisplayPage, caret)).toBe('#fff');
  expect(peerlessCaretColor({ ...page(run('#e0e0e0', 100, 160)), background: '#101010' } as DisplayPage, caret)).toBe('#e0e0e0');
});

test('a highlight clipped away from the caret does not colour it', () => {
  const highlight = { kind: 'decoration', deco: 'highlight', x: 60, y: 98, w: 100, h: 22, color: '#000080',
    clipGroup: { clip: { x: 140, y: 0, w: 200, h: 400 } } } as DisplayPrimitive;
  expect(peerlessCaretColor(page(highlight, run('#ffffff', 100, 160)), caret)).toBe('#000');
});

test('text over a shape or image keeps its own colour', () => {
  const shape = { kind: 'shape', x: 90, y: 90, w: 200, h: 40 } as DisplayPrimitive;
  expect(peerlessCaretColor(page(shape, run('#ffffff', 100, 160)), caret)).toBe('#ffffff');
  const image = { kind: 'image', x: 90, y: 90, w: 200, h: 40 } as DisplayPrimitive;
  expect(peerlessCaretColor(page(image, run('#ffffff', 100, 160)), caret)).toBe('#ffffff');
  const cover = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#ffffff' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(shape, cover, run('#ffffff', 100, 160)), caret)).toBe('#000');
});

test('transparent and translucent fills are ignored', () => {
  const clear = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: 'transparent' } as DisplayPrimitive;
  const faint = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: 'rgba(0, 0, 0, 0.1)' } as DisplayPrimitive;
  expect(peerlessCaretColor(page(clear, faint, run('#ffffff', 100, 160)), caret)).toBe('#000');
});

test('an unknown fill gives the caret a contrasting halo', () => {
  const shape = { kind: 'shape', x: 90, y: 90, w: 200, h: 40 } as DisplayPrimitive;
  expect(peerlessCaretPaint(page(shape), caret)).toEqual({ color: '#000', halo: '#fff' });
  expect(peerlessCaretPaint(page(shape, run('#ffffff', 100, 160)), caret)).toEqual({ color: '#ffffff', halo: '#000' });
  expect(peerlessCaretPaint(page(run('#c00000', 100, 160)), caret)).toEqual({ color: '#c00000' });
});

test('a rotated shape behind the caret still makes the fill unknown', () => {
  const rotated = { kind: 'shape', x: 100, y: 150, w: 200, h: 40, transform: { rotation: 90 } } as DisplayPrimitive;
  const at = { x: 200, y: 100, width: 1, height: 20 };
  expect(peerlessCaretPaint(page(rotated, run('#ffffff', 190, 230)), at)).toEqual({ color: '#ffffff', halo: '#000' });
  const image = { kind: 'image', x: 100, y: 150, w: 200, h: 40, rotationDeg: 90 } as DisplayPrimitive;
  expect(peerlessCaretPaint(page(image), at)).toEqual({ color: '#000', halo: '#fff' });
});

test('translucent fills are composited over what is beneath them', () => {
  const grey = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#ffffff80' } as DisplayPrimitive;
  expect(peerlessCaretPaint({ ...page(grey, run('#808080', 100, 160)), background: '#000000' } as DisplayPage, caret)).toEqual({ color: '#000' });
  const half = { kind: 'rect', x: 90, y: 90, w: 200, h: 40, fill: '#000000', opacity: 0.5 } as DisplayPrimitive;
  expect(peerlessCaretPaint(page(half, run('#808080', 100, 160)), caret)).toEqual({ color: '#000' });
});
