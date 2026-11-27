import { expect, test } from 'bun:test';
import type { Paint, PositionedTextLine, SlideDisplayList, SlidePrimitive } from '@betteroffice/pptx';
import { paintedCaretInk, pptxCaretColors } from './caretColor';

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
  expect(pptxCaretColors(slide([box]), line({ start: 0, end: 5, color: '#ffffff' }), 2, caret)).toEqual({ color: '#ffffff', caret: '#000000', background: '#000000' });
  expect(pptxCaretColors(slide([], gradient), line({ start: 0, end: 5, color: '#1f1f1f' }), 2, caret)).toEqual({ color: '#1f1f1f', caret: '#000000', background: '#ffffff' });
});

test('a gradient estimates the caret by its stops, not by the text', () => {
  const dark = { kind: 'gradient', gradientType: 'linear', angleDeg: 0, stops: [{ position: 0, color: '#101010' }, { position: 1, color: '#202020' }] } satisfies Paint;
  const box = { kind: 'shape', objectId: 1, x: 100, y: 80, w: 300, h: 80, name: 'Box', geometry: 'rect', path: [], fill: dark } as SlidePrimitive;
  expect(pptxCaretColors(slide([box]), line({ start: 0, end: 5, color: '#000000' }), 2, caret).caret).toBe('#ffffff');
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

function painted(width: number, height: number, pixel: (x: number, y: number) => readonly number[], fail?: Error) {
  const reads: number[][] = [];
  const canvas = { width, height } as HTMLCanvasElement;
  const read = (_source: HTMLCanvasElement, left: number, top: number, w: number, h: number) => {
    reads.push([left, top, w, h]);
    if (fail) throw fail;
    const data = new Uint8ClampedArray(w * h * 4);
    for (let y = 0; y < h; y += 1) for (let x = 0; x < w; x += 1) data.set(pixel(left + x, top + y), (y * w + x) * 4);
    return data;
  };
  return { canvas, reads, read };
}

const list = { width: 960, height: 540 };
const bar = (x: number, y = 100, height = 20) => [{ x: x + 0.5, y }, { x: x + 0.5, y: y + height }] as const;

test('the painted caret reads the pixels along its bar at the device pixel ratio and zoom', () => {
  const white = () => [255, 255, 255, 255];
  const zoomed = painted(960 * 3, 540 * 3, white);
  expect(paintedCaretInk(zoomed.canvas, list, ...bar(120), zoomed.read)).toBe('#000000');
  expect(zoomed.reads).toEqual([[360, 298, 3, 64]]);
  const half = painted(480, 270, white);
  paintedCaretInk(half.canvas, list, ...bar(121, 101), half.read);
  expect(half.reads).toEqual([[60, 50, 2, 11]]);
});

test('the painted caret is black or white by what is painted behind it, not by a nearby glyph', () => {
  const photo = (x: number) => x < 1440 ? [16, 16, 16, 255] : [240, 240, 240, 255];
  const dark = painted(2880, 1620, photo);
  expect(paintedCaretInk(dark.canvas, list, ...bar(120), dark.read)).toBe('#ffffff');
  const light = painted(2880, 1620, photo);
  expect(paintedCaretInk(light.canvas, list, ...bar(700), light.read)).toBe('#000000');
  const grey = painted(1920, 1080, () => [128, 128, 128, 255]);
  expect(paintedCaretInk(grey.canvas, list, ...bar(120), grey.read)).toBe('#000000');
  const glyph = painted(1920, 1080, (_x, y) => y < 215 ? [255, 255, 255, 255] : [11, 31, 75, 255]);
  expect(paintedCaretInk(glyph.canvas, list, ...bar(120), glyph.read)).toBe('#ffffff');
});

test('a rotated caret reads the pixels along its own bar, not its bounding box', () => {
  const stripe = painted(960, 540, (x, y) => Math.abs(x - y) <= 1 ? [16, 16, 16, 255] : [255, 255, 255, 255]);
  expect(paintedCaretInk(stripe.canvas, list, { x: 100.5, y: 100.5 }, { x: 114.5, y: 114.5 }, stripe.read)).toBe('#ffffff');
  expect(stripe.reads).toEqual([[100, 100, 15, 15]]);
});

test('unreadable pixels leave the caret to the estimate', () => {
  const tainted = painted(1920, 1080, () => [0, 0, 0, 255], new DOMException('tainted', 'SecurityError'));
  expect(paintedCaretInk(tainted.canvas, list, ...bar(120), tainted.read)).toBeNull();
  expect(tainted.reads).toHaveLength(1);
  const clear = painted(1920, 1080, () => [0, 0, 0, 0]);
  expect(paintedCaretInk(clear.canvas, list, ...bar(120), clear.read)).toBeNull();
  expect(paintedCaretInk(null, list, ...bar(120))).toBeNull();
  const empty = painted(0, 0, () => [0, 0, 0, 255]);
  expect(paintedCaretInk(empty.canvas, list, ...bar(120), empty.read)).toBeNull();
  expect(empty.reads).toEqual([]);
});

test('the slide canvas is copied into a separate read-optimised canvas and never read back itself', () => {
  const calls: unknown[][] = [];
  const context = {
    clearRect: () => {},
    drawImage: (...args: unknown[]) => { calls.push(['drawImage', ...args]); },
    getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(w * h * 4).fill(255) }),
  };
  const scratch = { width: 0, height: 0, getContext: (...args: unknown[]) => { calls.push(['scratch', ...args]); return context; } };
  const source = { width: 2880, height: 1620, ownerDocument: { createElement: () => scratch },
    getContext: (...args: unknown[]) => { calls.push(['source', ...args]); return null; } } as unknown as HTMLCanvasElement;
  expect(paintedCaretInk(source, list, ...bar(120))).toBe('#000000');
  expect(paintedCaretInk(source, list, ...bar(140))).toBe('#000000');
  expect(calls).toEqual([['scratch', '2d', { willReadFrequently: true }],
    ['drawImage', source, 360, 298, 3, 64, 0, 0, 3, 64], ['drawImage', source, 420, 298, 3, 64, 0, 0, 3, 64]]);
});
