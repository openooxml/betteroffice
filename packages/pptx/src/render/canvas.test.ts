import { describe, expect, test } from 'bun:test';
import type { SlideDisplayList } from '../types';
import { paintSlide } from './canvas';

describe('PPTX canvas replay', () => {
  test('paints shape geometry and positioned text in display-list order', async () => {
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

  test('paints chart parts clipped to the chart rectangle', async () => {
    const calls: string[] = [];
    const ctx = new Proxy(
      {
        clip: () => calls.push('clip'),
        fillText: (text: string) => calls.push(`text:${text}`),
        fill: () => calls.push('fill'),
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
      primitives: [
        {
          kind: 'chart',
          objectId: 4,
          shapeId: 'slide:0:1',
          name: 'Revenue chart',
          label: 'Revenue, column chart, 2 series, 3 categories',
          x: 10,
          y: 10,
          w: 300,
          h: 160,
          primitives: [
            {
              kind: 'shape',
              objectId: 4,
              name: '',
              x: 20,
              y: 20,
              w: 40,
              h: 100,
              geometry: 'rect',
              path: [
                { type: 'move', x: 0, y: 0 },
                { type: 'line', x: 1, y: 0 },
                { type: 'close' },
              ],
              fill: { kind: 'solid', color: '#6254e7' },
            },
            {
              kind: 'textBox',
              objectId: 4,
              x: 20,
              y: 130,
              w: 60,
              h: 14,
              anchor: 'top',
              paragraphs: [],
              lines: [
                {
                  x: 20,
                  y: 130,
                  width: 20,
                  height: 14,
                  baseline: 141,
                  start: 0,
                  end: 2,
                  caretStops: [],
                  runs: [
                    {
                      text: 'Q1',
                      start: 0,
                      end: 2,
                      x: 20,
                      width: 20,
                      fontId: 1,
                      fontFamily: 'Liberation Sans',
                      fontSizePx: 10,
                      bold: false,
                      italic: false,
                      underline: false,
                      color: '#222222',
                      glyphs: [],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    await paintSlide(ctx, list, 1);
    expect(calls).toContain('clip');
    expect(calls).toContain('fill');
    expect(calls).toContain('text:Q1');
  });
});
