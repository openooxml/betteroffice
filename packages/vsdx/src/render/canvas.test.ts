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
