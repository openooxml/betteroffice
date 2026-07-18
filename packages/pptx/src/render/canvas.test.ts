import { describe, expect, test } from 'bun:test';
import type { SlideDisplayList } from '../types';
import { paintSlide } from './canvas';

describe('PPTX canvas replay', () => {
  test('paints shape geometry and positioned text in Rust order', async () => {
    const calls: string[] = [];
    const ctx = new Proxy(
      {
        createLinearGradient: () => ({ addColorStop: () => undefined }),
        createRadialGradient: () => ({ addColorStop: () => undefined }),
        fillText: (text: string) => calls.push(`text:${text}`),
        moveTo: () => calls.push('move'),
        lineTo: () => calls.push('line'),
        fill: () => calls.push('fill'),
        stroke: () => calls.push('stroke'),
      } as Record<string, unknown>,
      {
        get(target, property) {
          if (property in target) return target[property as string];
          return () => undefined;
        },
        set(target, property, value) {
          target[property as string] = value;
          return true;
        },
      }
    ) as unknown as CanvasRenderingContext2D;
    const list: SlideDisplayList = {
      contractVersion: 1,
      width: 320,
      height: 180,
      background: { kind: 'solid', color: '#ffffff' },
      primitives: [
        {
          kind: 'shape',
          objectId: 1,
          shapeId: 'shape:1',
          name: 'Card',
          x: 20,
          y: 20,
          w: 280,
          h: 140,
          geometry: 'rect',
          path: [
            { type: 'move', x: 0, y: 0 },
            { type: 'line', x: 1, y: 0 },
            { type: 'line', x: 1, y: 1 },
            { type: 'close' },
          ],
          fill: { kind: 'solid', color: '#325ee6' },
          stroke: { color: '#10235b', width: 2 },
        },
        {
          kind: 'textBox',
          objectId: 1,
          shapeId: 'shape:1',
          storyId: 'story:1',
          x: 40,
          y: 50,
          w: 240,
          h: 80,
          anchor: 'top',
          paragraphs: [],
          lines: [
            {
              x: 40,
              y: 50,
              width: 60,
              height: 24,
              baseline: 68,
              start: 0,
              end: 5,
              caretStops: [
                { position: 0, x: 40 },
                { position: 5, x: 100 },
              ],
              runs: [
                {
                  text: 'Hello',
                  start: 0,
                  end: 5,
                  x: 40,
                  width: 60,
                  fontId: 1,
                  fontFamily: 'Liberation Sans',
                  fontSizePx: 20,
                  bold: false,
                  italic: false,
                  underline: false,
                  color: '#ffffff',
                  glyphs: [],
                },
              ],
            },
          ],
        },
      ],
    };

    await paintSlide(ctx, list, 2);
    expect(calls).toContain('move');
    expect(calls).toContain('line');
    expect(calls).toContain('fill');
    expect(calls).toContain('stroke');
    expect(calls).toContain('text:Hello');
  });
});
