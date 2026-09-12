import { expect, test } from 'bun:test';
import { paintPage } from './canvas';
import type { PageDisplayList } from '../types';

function context(log: string[]): CanvasRenderingContext2D {
  return new Proxy({
    createLinearGradient: () => ({ addColorStop: () => {} }),
  }, {
    get(target, key) {
      if (key in target) return Reflect.get(target, key);
      return (...args: unknown[]) => { log.push(`${String(key)}:${args.join(',')}`); };
    },
    set(_, key, value) { log.push(`${String(key)}=${String(value)}`); return true; },
  }) as unknown as CanvasRenderingContext2D;
}

const transform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

test('replays primitives in z order and paints placeholders', async () => {
  const log: string[] = [];
  const list: PageDisplayList = {
    contractVersion: 3, width: 100, height: 100, paintTransform: transform,
    primitives: [
      { kind: 'placeholder', id: 'late', zOrder: 2, x: 10, y: 10, width: 20, height: 20, reason: 'missing image' },
      { kind: 'shape', id: 'early', zOrder: 1, path: [{ type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 1 }], fill: { kind: 'solid', color: '#000' } },
    ],
  };
  await paintPage(context(log), list);
  expect(log.findIndex(entry => entry.startsWith('moveTo'))).toBeLessThan(log.findIndex(entry => entry.startsWith('strokeRect')));
  expect(log.some(entry => entry.startsWith('strokeRect'))).toBe(true);
  expect(log.some(entry => entry.startsWith('fillText:missing image'))).toBe(true);
});

test('rejects display-list versions other than v3', async () => {
  await expect(paintPage(context([]), { contractVersion: 2, width: 1, height: 1, paintTransform: transform, primitives: [] } as unknown as PageDisplayList)).rejects.toThrow('unsupported VSDX display-list contract version 2');
});

test('replays positioned text runs at their line caret positions', async () => {
  const log: string[] = [];
  const list: PageDisplayList = {
    contractVersion: 3, width: 100, height: 100, paintTransform: transform,
    primitives: [{
      kind: 'textBox', id: 'text', zOrder: 1, x: 1, y: 2, width: 90, height: 80,
      paragraphs: [
        { runs: [{ text: 'left', family: 'Arial', sizeIn: 12, bold: false, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#111', diagnostics: [] }] },
        { runs: [{ text: 'right', family: 'Arial', sizeIn: 12, bold: true, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#222', diagnostics: [] }] },
      ],
      lines: [
        { x: 30, y: 20, width: 20, height: 12, start: 0, end: 4, caretStops: [{ position: 0, x: 30, y: 20 }, { position: 4, x: 50, y: 20 }] },
        { x: 60, y: 45, width: 25, height: 12, start: 4, end: 9, caretStops: [{ position: 4, x: 60, y: 45 }, { position: 9, x: 85, y: 45 }] },
      ],
    }],
  };
  await paintPage(context(log), list);
  expect(log.filter(entry => entry.startsWith('fillText:'))).toEqual(['fillText:left,30,20', 'fillText:right,60,45']);
});

test('paints a rotated text box through its own transform', async () => {
  const log: string[] = [];
  const rotated = { a: 0, b: 1, c: -1, d: 0, e: 10, f: 20 };
  const list: PageDisplayList = {
    contractVersion: 3, width: 100, height: 100, paintTransform: transform,
    primitives: [{
      kind: 'textBox', id: 'text', zOrder: 1, x: 0, y: 0, width: 50, height: 20, transform: rotated,
      paragraphs: [{ runs: [{ text: 'turn', family: 'Arial', sizeIn: 12, bold: false, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#111', diagnostics: [] }] }],
      lines: [{ x: 0, y: 10, width: 30, height: 12, start: 0, end: 4, caretStops: [{ position: 0, x: 0, y: 10 }, { position: 4, x: 30, y: 10 }] }],
    }],
  };
  await paintPage(context(log), list);
  expect(log).toContain('transform:0,1,-1,0,10,20');
  expect(log).toContain('fillText:turn,0,10');
});
