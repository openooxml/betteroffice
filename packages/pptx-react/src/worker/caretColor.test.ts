import { expect, test } from 'bun:test';
import type { Paint, PositionedTextLine, SlideDisplayList, SlidePrimitive } from '@betteroffice/pptx';
import { pptxCaretColors } from './caretColor';

const caret = { x: 120, y: 100, width: 1, height: 20 };

function line(...runs: Array<{ start: number; end: number; color: string }>): PositionedTextLine {
  return { runs: runs.map((run) => ({ ...run, text: 'x'.repeat(run.end - run.start) })) } as unknown as PositionedTextLine;
}

function slide(primitives: SlidePrimitive[], background?: SlideDisplayList['background']): SlideDisplayList {
  return { contractVersion: 1, width: 960, height: 540, background, primitives };
}

const shape = (color: string, box = { x: 100, y: 80, w: 300, h: 80 }) =>
  ({ kind: 'shape', objectId: 1, ...box, name: 'Box', geometry: 'rect', path: [], fill: { kind: 'solid', color } }) as SlidePrimitive;

test('caret takes the run colour on a white slide', () => {
  expect(pptxCaretColors(slide([]), line({ start: 0, end: 5, color: '#c00000' }), 3, caret))
    .toEqual({ color: '#c00000', caret: '#000000', background: '#ffffff' });
});

test('caret at a run boundary follows the run before it', () => {
  expect(pptxCaretColors(slide([]), line({ start: 0, end: 3, color: '#0000ff' }, { start: 3, end: 6, color: '#00aa00' }), 3, caret).color)
    .toBe('#0000ff');
  expect(pptxCaretColors(slide([]), line({ start: 0, end: 3, color: '#0000ff' }), 0, caret).color).toBe('#0000ff');
});

test('white text on a dark slide background keeps a white caret', () => {
  expect(pptxCaretColors(slide([], { kind: 'solid', color: '#0b1f3a' }), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret))
    .toEqual({ color: '#ffffff', caret: '#ffffff', background: '#0b1f3a' });
});

test('white text on a dark shape fill keeps a white caret', () => {
  expect(pptxCaretColors(slide([shape('#202020')]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret))
    .toEqual({ color: '#ffffff', caret: '#ffffff', background: '#202020' });
});

test('fills that do not cover the caret are ignored', () => {
  expect(pptxCaretColors(slide([shape('#000000', { x: 500, y: 300, w: 50, h: 50 })]), line(), 0, caret))
    .toEqual({ color: '#000000', caret: '#000000', background: '#ffffff' });
});

test('dark slide with no run or unreadable text falls back to white', () => {
  const dark = slide([], { kind: 'gradient', gradientType: 'linear', stops: [{ position: 0, color: '#000000' }, { position: 1, color: '#202020' }] });
  expect(pptxCaretColors(dark, null, null, caret).color).toBe('#ffffff');
  expect(pptxCaretColors(slide([], { kind: 'solid', color: '#101010' }), line({ start: 0, end: 5, color: '#111111' }), 2, caret).color).toBe('#ffffff');
});

test('table cell shapes colour the caret background', () => {
  const table = { kind: 'table', objectId: 2, x: 100, y: 80, w: 300, h: 80, name: 'Table', label: 'Table', primitives: [shape('#1f3864')] } as SlidePrimitive;
  expect(pptxCaretColors(slide([table]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret).background).toBe('#1f3864');
});

test('a gradient fill trusts the run colour, wherever the caret sits on it', () => {
  const gradient = { kind: 'gradient', gradientType: 'linear', angleDeg: 0, stops: [{ position: 0, color: '#000000' }, { position: 0.6, color: '#eeeeee' }, { position: 0.8, color: '#f8f8f8' }, { position: 1, color: '#ffffff' }] } satisfies Paint;
  const box = { kind: 'shape', objectId: 1, x: 100, y: 80, w: 300, h: 80, name: 'Box', geometry: 'rect', path: [], fill: gradient } as SlidePrimitive;
  expect(pptxCaretColors(slide([box]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret)).toEqual({ color: '#ffffff', caret: '#ffffff', background: '#000000' });
  expect(pptxCaretColors(slide([], gradient), line({ start: 0, end: 5, color: '#1f1f1f' }), 2, caret)).toEqual({ color: '#1f1f1f', caret: '#000000', background: '#ffffff' });
});

test('rotated or non-rectangular shapes and images make the fill unknown, so the run colour is trusted', () => {
  const dark = { kind: 'solid', color: '#000000' } satisfies Paint;
  const rotated = { ...shape('#ffffff'), transform: { rotationDeg: 45 } } as SlidePrimitive;
  expect(pptxCaretColors(slide([rotated], dark), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret).color).toBe('#ffffff');
  const ellipse = { ...shape('#ffffff'), geometry: 'ellipse' } as SlidePrimitive;
  expect(pptxCaretColors(slide([ellipse], dark), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret).color).toBe('#ffffff');
  const image = { kind: 'image', objectId: 3, x: 100, y: 80, w: 300, h: 80 } as SlidePrimitive;
  expect(pptxCaretColors(slide([image]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret).color).toBe('#ffffff');
  expect(pptxCaretColors(slide([shape('#ffffff')], dark), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret).color).toBe('#000000');
});

test('an empty line over a shape of unknown extent reads against the last known fill', () => {
  const dark = { kind: 'solid', color: '#000000' } satisfies Paint;
  const rotated = { ...shape('#ffffff'), transform: { rotationDeg: 45 } } as SlidePrimitive;
  expect(pptxCaretColors(slide([rotated], dark), line(), 0, caret).color).toBe('#ffffff');
  const ellipse = { ...shape('#000000'), geometry: 'ellipse' } as SlidePrimitive;
  expect(pptxCaretColors(slide([ellipse]), line(), 0, caret).color).toBe('#000000');
});

test('an empty line over an unknown fill follows the insertion colour when known', () => {
  const ellipse = { ...shape('#000000'), geometry: 'ellipse' } as SlidePrimitive;
  expect(pptxCaretColors(slide([ellipse]), line(), 0, caret)).toEqual({ color: '#000000', caret: '#ffffff', background: '#ffffff' });
  expect(pptxCaretColors(slide([ellipse]), line(), 0, caret, '#ffffff')).toEqual({ color: '#ffffff', caret: '#ffffff', background: '#000000' });
  expect(pptxCaretColors(slide([]), line(), 0, caret, '#c00000')).toEqual({ color: '#c00000', caret: '#000000', background: '#ffffff' });
  expect(pptxCaretColors(slide([], { kind: 'solid', color: '#000000' }), line(), 0, caret, '#000000').color).toBe('#ffffff');
});

test('translucent fills are composited over what is beneath them', () => {
  const dark = { kind: 'solid', color: '#000000' } satisfies Paint;
  expect(pptxCaretColors(slide([shape('#ffffff80')], dark), line(), 0, caret, '#808080')).toEqual({ color: '#000000', caret: '#000000', background: '#808080' });
  expect(pptxCaretColors(slide([shape('#00000000')]), line({ start: 0, end: 5, color: '#1f1f1f' }), 2, caret)).toEqual({ color: '#1f1f1f', caret: '#000000', background: '#ffffff' });
});

test('a translucent fill over an image keeps the fill unknown', () => {
  const photo = { kind: 'image', objectId: 3, x: 100, y: 80, w: 300, h: 80 } as SlidePrimitive;
  expect(pptxCaretColors(slide([photo, shape('#ffffff00')]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret))
    .toEqual({ color: '#ffffff', caret: '#ffffff', background: '#000000' });
  expect(pptxCaretColors(slide([photo, shape('#202020')]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret))
    .toEqual({ color: '#ffffff', caret: '#ffffff', background: '#202020' });
});

test('the caret is black or white against the fill, whatever the text colour', () => {
  const dark = { kind: 'solid', color: '#0b1f3a' } satisfies Paint;
  expect(pptxCaretColors(slide([]), line({ start: 0, end: 5, color: '#315efb' }), 2, caret).caret).toBe('#000000');
  expect(pptxCaretColors(slide([], dark), line({ start: 0, end: 5, color: '#c8f56a' }), 2, caret).caret).toBe('#ffffff');
  expect(pptxCaretColors(slide([], dark), line({ start: 0, end: 5, color: '#6b7280' }), 2, caret).caret).toBe('#ffffff');
});

test('a rounded shape colours the caret by its own fill', () => {
  const card = (color: string) => ({ ...shape(color), geometry: 'roundRect' }) as SlidePrimitive;
  expect(pptxCaretColors(slide([card('#fff0ea')]), line({ start: 0, end: 4, color: '#e0603a' }), 2, caret).caret).toBe('#000000');
  expect(pptxCaretColors(slide([card('#101828')]), line({ start: 0, end: 5, color: '#c8f56a' }), 2, caret).caret).toBe('#ffffff');
});

test('the fill, not the text, decides the caret where both shades would show', () => {
  expect(pptxCaretColors(slide([shape('#808080')]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret).caret).toBe('#000000');
});

test('a shape colours the caret only where its outline covers it', () => {
  const diamond = { ...shape('#000000'), geometry: 'diamond', path: [
    { type: 'move', x: 0.5, y: 0 }, { type: 'line', x: 1, y: 0.5 }, { type: 'line', x: 0.5, y: 1 }, { type: 'line', x: 0, y: 0.5 }, { type: 'close' },
  ] } as SlidePrimitive;
  const text = line({ start: 0, end: 5, color: '#000000' });
  expect(pptxCaretColors(slide([diamond]), text, 2, caret).caret).toBe('#000000');
  expect(pptxCaretColors(slide([diamond]), text, 2, { ...caret, x: 250 }).caret).toBe('#ffffff');
  const rotated = { ...diamond, transform: { rotationDeg: 30 } } as SlidePrimitive;
  expect(pptxCaretColors(slide([rotated]), text, 2, caret).caret).toBe('#000000');
});
