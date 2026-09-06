import { describe, expect, test } from 'bun:test';
import type { GeometryPathCommand, ImageEffect, ShapePrimitive, SlideDisplayList } from '../types';
import { applyImageEffects, paintSlide } from './canvas';

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

  test('strokes a gradient outline with a gradient sized to the shape box', async () => {
    const gradients: Array<{ args: number[]; stops: Array<[number, string]> }> = [];
    let strokeStyle: unknown;
    let fillStyle: unknown;
    const objects: unknown[] = [];
    const ctx = new Proxy(
      {
        createLinearGradient: (...args: number[]) => {
          const entry = { args, stops: [] as Array<[number, string]> };
          gradients.push(entry);
          const gradient = {
            addColorStop: (position: number, color: string) => entry.stops.push([position, color]),
          };
          objects.push(gradient);
          return gradient;
        },
      } as Record<string, unknown>,
      {
        get(target, property) {
          if (property in target) return target[property as string];
          return () => undefined;
        },
        set(target, property, value) {
          if (property === 'strokeStyle') strokeStyle = value;
          if (property === 'fillStyle') fillStyle = value;
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
          kind: 'shape',
          objectId: 1,
          name: 'Spoke',
          x: 20,
          y: 40,
          w: 200,
          h: 0,
          geometry: 'line',
          path: [
            { type: 'move', x: 0, y: 0 },
            { type: 'line', x: 1, y: 0 },
          ],
          stroke: {
            color: '#c00000',
            width: 3,
            paint: {
              kind: 'gradient',
              gradientType: 'linear',
              angleDeg: 0,
              stops: [
                { position: 0, color: '#c00000' },
                { position: 1, color: '#c2c2c2' },
              ],
            },
          },
        },
      ],
    };

    await paintSlide(ctx, list, 1);
    expect(gradients).toHaveLength(1);
    expect(gradients[0]?.args).toEqual([20, 40, 220, 40]);
    expect(gradients[0]?.stops).toEqual([
      [0, '#c00000'],
      [1, '#c2c2c2'],
    ]);
    expect(strokeStyle).toBe(objects[0]);
    const shape = list.primitives[0];
    if (shape?.kind !== 'shape' || !shape.stroke) throw new Error('shape');
    shape.stroke.tailEnd = { kind: 'triangle', width: 9, length: 9 };
    await paintSlide(ctx, list, 1);
    expect(fillStyle).toBe(objects[2]);
    expect(strokeStyle).toBe(objects[2]);
    list.primitives = [{
      kind: 'image', objectId: 2, name: 'Outline', x: 20, y: 40, w: 200, h: 0,
      stroke: shape.stroke,
    }];
    await paintSlide(ctx, list, 1);
    expect(strokeStyle).toBe(objects[3]);
    expect(gradients[3]?.args).toEqual([20, 40, 220, 40]);
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

  test('paints justified word starts at the engine caret positions', async () => {
    const calls: Array<{ text: string; x: number }> = [];
    const ctx = new Proxy(
      {
        fillText: (text: string, x: number) => calls.push({ text, x }),
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
    const caretStops = [
      { position: 0, x: 40 },
      { position: 1, x: 50 },
      { position: 2, x: 60 },
      { position: 3, x: 70 },
      { position: 4, x: 100 },
      { position: 5, x: 110 },
      { position: 6, x: 120 },
      { position: 7, x: 130 },
    ];
    const glyphs = [
      { glyphId: 1, cluster: 0, x: 40, advance: 10, xOffset: 0, yOffset: 68 },
      { glyphId: 2, cluster: 1, x: 50, advance: 10, xOffset: 0, yOffset: 68 },
      { glyphId: 3, cluster: 2, x: 60, advance: 10, xOffset: 0, yOffset: 68 },
      { glyphId: 4, cluster: 3, x: 70, advance: 5, xOffset: 0, yOffset: 68 },
      { glyphId: 5, cluster: 4, x: 100, advance: 10, xOffset: 0, yOffset: 68 },
      { glyphId: 6, cluster: 5, x: 110, advance: 10, xOffset: 0, yOffset: 68 },
      { glyphId: 7, cluster: 6, x: 120, advance: 10, xOffset: 0, yOffset: 68 },
    ];
    const list: SlideDisplayList = {
      contractVersion: 1,
      width: 320,
      height: 180,
      primitives: [
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
              width: 90,
              height: 24,
              baseline: 68,
              start: 0,
              end: 7,
              caretStops,
              runs: [
                {
                  text: 'one two',
                  start: 0,
                  end: 7,
                  x: 40,
                  width: 90,
                  fontId: 1,
                  fontFamily: 'Liberation Sans',
                  fontSizePx: 20,
                  bold: false,
                  italic: false,
                  underline: false,
                  color: '#ffffff',
                  glyphs,
                },
              ],
            },
          ],
        },
      ],
    };

    await paintSlide(ctx, list, 1);
    const painted = calls.find((call) => call.text === 'two');
    const engine = caretStops.find((stop) => stop.position === 4);
    expect(painted?.x).toBe(engine?.x);
    expect(calls).toEqual([
      { text: 'one ', x: 40 },
      { text: 'two', x: 100 },
    ]);
  });
});

type Call = [string, ...number[]];

function recordingContext(calls: Call[]): CanvasRenderingContext2D {
  const round = (value: number): number => Math.round(value * 1e6) / 1e6 || 0;
  const target: Record<string, unknown> = {};
  for (const name of ['beginPath', 'moveTo', 'lineTo', 'closePath', 'fill', 'stroke', 'ellipse', 'rotate']) {
    target[name] = (...args: unknown[]) =>
      calls.push([name, ...args.filter((arg): arg is number => typeof arg === 'number').map(round)]);
  }
  return new Proxy(target, {
    get(record, property) {
      if (property in record) return record[property as string];
      return () => undefined;
    },
    set(record, property, value) {
      record[property as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

async function paintLine(
  line: Pick<ShapePrimitive, 'x' | 'y' | 'w' | 'h' | 'path' | 'stroke' | 'transform'>
): Promise<Call[]> {
  const calls: Call[] = [];
  const list: SlideDisplayList = {
    contractVersion: 1,
    width: 640,
    height: 720,
    primitives: [{ kind: 'shape', objectId: 1, shapeId: 'shape:1', name: 'Line', geometry: 'line', ...line }],
  };
  await paintSlide(recordingContext(calls), list, 1);
  return calls;
}

function afterPathStroke(calls: Call[]): Call[] {
  return calls.slice(calls.findIndex(([name]) => name === 'stroke') + 1);
}

describe('PPTX canvas line ends', () => {
  const straight = {
    x: 10,
    y: 50,
    w: 100,
    h: 0,
    path: [
      { type: 'move' as const, x: 0, y: 0 },
      { type: 'line' as const, x: 1, y: 0 },
    ],
  };

  test('draws each kind at the head, sized by the end width and length', async () => {
    const marks: Record<string, Call[]> = {
      triangle: [
        ['beginPath'],
        ['moveTo', 10, 50],
        ['lineTo', 19, 47],
        ['lineTo', 19, 53],
        ['closePath'],
        ['fill'],
      ],
      stealth: [
        ['beginPath'],
        ['moveTo', 10, 50],
        ['lineTo', 19, 47],
        ['lineTo', 15.4, 50],
        ['lineTo', 19, 53],
        ['closePath'],
        ['fill'],
      ],
      arrow: [['beginPath'], ['moveTo', 19, 47], ['lineTo', 10, 50], ['lineTo', 19, 53], ['stroke']],
      diamond: [
        ['beginPath'],
        ['moveTo', 5.5, 50],
        ['lineTo', 10, 47],
        ['lineTo', 14.5, 50],
        ['lineTo', 10, 53],
        ['closePath'],
        ['fill'],
      ],
      oval: [['beginPath'], ['ellipse', 10, 50, 4.5, 3, 3.141593, 0, 6.283185], ['fill']],
    };
    for (const [kind, expected] of Object.entries(marks)) {
      const calls = await paintLine({
        ...straight,
        stroke: { color: '#ff0000', width: 2, headEnd: { kind, width: 6, length: 9 } },
      });
      expect(afterPathStroke(calls)).toEqual(expected);
    }
  });

  test('paints nothing for missing or unknown ends', async () => {
    const plain = await paintLine({ ...straight, stroke: { color: '#ff0000', width: 2 } });
    expect(afterPathStroke(plain)).toEqual([]);
    const unknown = await paintLine({
      ...straight,
      stroke: {
        color: '#ff0000',
        width: 2,
        headEnd: { kind: 'squiggle', width: 6, length: 9 },
        tailEnd: { kind: 'none', width: 6, length: 9 },
      },
    });
    expect(afterPathStroke(unknown)).toEqual([]);
  });

  test('orients both ends along a rotated diagonal line in its own frame', async () => {
    const calls = await paintLine({
      x: 0,
      y: 0,
      w: 30,
      h: 40,
      path: [
        { type: 'move', x: 0, y: 0 },
        { type: 'line', x: 1, y: 1 },
      ],
      transform: { rotationDeg: 30 },
      stroke: {
        color: '#ff0000',
        width: 2,
        headEnd: { kind: 'triangle', width: 10, length: 5 },
        tailEnd: { kind: 'triangle', width: 10, length: 5 },
      },
    });
    const rotation = calls.findIndex(([name]) => name === 'rotate');
    expect(calls[rotation]).toEqual(['rotate', 0.523599]);
    expect(rotation).toBeLessThan(calls.findIndex(([name]) => name === 'moveTo'));
    expect(afterPathStroke(calls)).toEqual([
      ['beginPath'],
      ['moveTo', 0, 0],
      ['lineTo', 7, 1],
      ['lineTo', -1, 7],
      ['closePath'],
      ['fill'],
      ['beginPath'],
      ['moveTo', 30, 40],
      ['lineTo', 23, 39],
      ['lineTo', 31, 33],
      ['closePath'],
      ['fill'],
    ]);
  });

  test('follows the first and last segments of a bent connector', async () => {
    const calls = await paintLine({
      x: 80,
      y: 480,
      w: 320,
      h: 160,
      path: [
        { type: 'move', x: 0, y: 0 },
        { type: 'line', x: 0.5, y: 0 },
        { type: 'line', x: 0.5, y: 1 },
        { type: 'line', x: 1, y: 1 },
      ],
      stroke: {
        color: '#ff0000',
        width: 2,
        headEnd: { kind: 'triangle', width: 8, length: 12 },
        tailEnd: { kind: 'triangle', width: 8, length: 12 },
      },
    });
    expect(afterPathStroke(calls)).toEqual([
      ['beginPath'],
      ['moveTo', 80, 480],
      ['lineTo', 92, 476],
      ['lineTo', 92, 484],
      ['closePath'],
      ['fill'],
      ['beginPath'],
      ['moveTo', 400, 640],
      ['lineTo', 388, 644],
      ['lineTo', 388, 636],
      ['closePath'],
      ['fill'],
    ]);
  });
  test('uses curve tangents and skips repeated endpoint coordinates', async () => {
    const paths: ShapePrimitive['path'][] = [
      [
        { type: 'move', x: 0, y: 0 },
        { type: 'quad', cpx: 0, cpy: 1, x: 1, y: 1 },
      ],
      [
        { type: 'move', x: 0, y: 0 },
        { type: 'cubic', cp1x: 0, cp1y: 1, cp2x: 0, cp2y: 1, x: 1, y: 1 },
      ],
      [
        { type: 'move', x: 0, y: 0 },
        { type: 'line', x: 0, y: 0 },
        { type: 'line', x: 0, y: 1 },
        { type: 'line', x: 1, y: 1 },
        { type: 'line', x: 1, y: 1 },
      ],
    ];
    for (const path of paths) {
      const calls = await paintLine({
        x: 10,
        y: 50,
        w: 100,
        h: 100,
        path,
        stroke: {
          color: '#ff0000',
          width: 2,
          headEnd: { kind: 'triangle', width: 8, length: 12 },
          tailEnd: { kind: 'triangle', width: 8, length: 12 },
        },
      });
      expect(afterPathStroke(calls)).toEqual([
        ['beginPath'],
        ['moveTo', 10, 50],
        ['lineTo', 14, 62],
        ['lineTo', 6, 62],
        ['closePath'],
        ['fill'],
        ['beginPath'],
        ['moveTo', 110, 150],
        ['lineTo', 98, 154],
        ['lineTo', 98, 146],
        ['closePath'],
        ['fill'],
      ]);
    }
  });

});

describe('PPTX picture cropping', () => {
  function harness() {
    const calls: string[] = [];
    const ctx = new Proxy(
      {
        drawImage: (...args: unknown[]) => calls.push(`draw:${args.slice(1).join(',')}`),
        clip: () => calls.push('clip'),
        save: () => calls.push('save'),
        restore: () => calls.push('restore'),
        beginPath: () => calls.push('beginPath'),
        rect: (...args: unknown[]) => calls.push(`rect:${args.join(',')}`),
        moveTo: (...args: unknown[]) => calls.push(`move:${args.join(',')}`),
        lineTo: (...args: unknown[]) => calls.push(`line:${args.join(',')}`),
        bezierCurveTo: (...args: unknown[]) => calls.push(`cubic:${args.join(',')}`),
        closePath: () => calls.push('close'),
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
    return { calls, ctx };
  }

  function list(image: Record<string, unknown>): SlideDisplayList {
    return {
      contractVersion: 1,
      width: 320,
      height: 180,
      primitives: [
        {
          kind: 'image',
          objectId: 1,
          name: 'Screenshot',
          x: 10,
          y: 20,
          w: 200,
          h: 100,
          assetId: 'ppt/media/image1.png',
          ...image,
        },
      ],
    } as SlideDisplayList;
  }

  const source = { width: 400, height: 300 } as unknown as CanvasImageSource;

  const ellipse: GeometryPathCommand[] = [
    { type: 'move', x: 1, y: 0.5 },
    { type: 'cubic', cp1x: 1, cp1y: 0.75, cp2x: 0.75, cp2y: 1, x: 0.5, y: 1 },
    { type: 'cubic', cp1x: 0.25, cp1y: 1, cp2x: 0, cp2y: 0.75, x: 0, y: 0.5 },
    { type: 'cubic', cp1x: 0, cp1y: 0.25, cp2x: 0.25, cp2y: 0, x: 0.5, y: 0 },
    { type: 'cubic', cp1x: 0.75, cp1y: 0, cp2x: 1, cp2y: 0.25, x: 1, y: 0.5 },
    { type: 'close' },
  ];

  const ellipseOutline = [
    'beginPath',
    'move:210,70',
    'cubic:210,95,160,120,110,120',
    'cubic:60,120,10,95,10,70',
    'cubic:10,45,60,20,110,20',
    'cubic:160,20,210,45,210,70',
    'close',
  ];

  test('a cropped picture with its own outline is clipped and stroked along that outline', async () => {
    const { calls, ctx } = harness();
    await paintSlide(
      ctx,
      list({
        crop: { left: 0.25, top: 0.5, right: 0.25, bottom: 0.25 },
        path: ellipse,
        stroke: { color: '#ff00ff', width: 2 },
      }),
      1,
      1,
      { resolveImage: async () => source }
    );
    expect(calls).toEqual([
      'save',
      'save',
      'save',
      ...ellipseOutline,
      'clip',
      'draw:100,150,200,75,10,20,200,100',
      'restore',
      ...ellipseOutline,
      'stroke',
      'restore',
      'restore',
    ]);
  });

  test('an uncropped picture with its own outline still draws the whole source through it', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list({ path: ellipse }), 1, 1, { resolveImage: async () => source });
    expect(calls).toEqual([
      'save',
      'save',
      'save',
      ...ellipseOutline,
      'clip',
      'draw:0,0,400,300,10,20,200,100',
      'restore',
      'restore',
      'restore',
    ]);
  });

  test('a picture without its own outline is stroked along its frame', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list({ stroke: { color: '#10235b', width: 1 } }), 1, 1, {
      resolveImage: async () => source,
    });
    expect(calls).toEqual([
      'save',
      'save',
      'draw:0,0,400,300,10,20,200,100',
      'beginPath',
      'rect:10,20,200,100',
      'stroke',
      'restore',
      'restore',
    ]);
  });

  test('draws only the kept sub-rectangle, masked to the frame', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list({ crop: { top: 0.1, bottom: 0.2 } }), 1, 1, {
      resolveImage: async () => source,
    });
    expect(calls).toContain('save');
    expect(calls).toContain('clip');
    expect(calls).toContain('restore');
    expect(calls).toContain('draw:0,30,400,210,10,20,200,100');
  });

  test('an uncropped picture draws the whole source and needs no mask', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list({}), 1, 1, { resolveImage: async () => source });
    expect(calls).toContain('draw:0,0,400,300,10,20,200,100');
    expect(calls).not.toContain('clip');
  });

  test('a two-contour mask keeps both contours, so a counter can be punched out', async () => {
    const { calls, ctx } = harness();
    const ring: GeometryPathCommand[] = [
      { type: 'move', x: 0, y: 0 },
      { type: 'line', x: 1, y: 0 },
      { type: 'line', x: 1, y: 1 },
      { type: 'line', x: 0, y: 1 },
      { type: 'close' },
      { type: 'move', x: 0.25, y: 0.25 },
      { type: 'line', x: 0.25, y: 0.75 },
      { type: 'line', x: 0.75, y: 0.75 },
      { type: 'line', x: 0.75, y: 0.25 },
      { type: 'close' },
    ];
    await paintSlide(ctx, list({ path: ring }), 1, 1, { resolveImage: async () => source });
    expect(calls).toEqual([
      'save',
      'save',
      'save',
      'beginPath',
      'move:10,20',
      'line:210,20',
      'line:210,120',
      'line:10,120',
      'close',
      'move:60,45',
      'line:60,95',
      'line:160,95',
      'line:160,45',
      'close',
      'clip',
      'draw:0,0,400,300,10,20,200,100',
      'restore',
      'restore',
      'restore',
    ]);
  });

  test('a crop that keeps nothing draws nothing', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list({ crop: { left: 0.6, right: 0.6 } }), 1, 1, {
      resolveImage: async () => source,
    });
    expect(calls.some((call) => call.startsWith('draw:'))).toBe(false);
  });
});

describe('PPTX shape shadows', () => {
  function harness(supportsFilters = true) {
    const calls: string[] = [];
    const transforms: number[][] = [];
    const surfaces: number[][] = [];
    const makeContext = (name: string) => {
      const state: Record<string, unknown> = {
        canvas: { width: 480, height: 480 }, filter: supportsFilters ? 'none' : undefined,
        getTransform: () => ({ a: 3, b: 0, c: 0, d: 3, e: 0, f: 0 }),
        setTransform: (...matrix: number[]) => { if (name === 'mask') transforms.push(matrix); },
        fill: () => calls.push(`${name}:fill`),
        stroke: () => calls.push(`${name}:stroke`),
        fillRect: () => calls.push(`${name}:tint:${state.globalCompositeOperation}:${state.fillStyle}`),
        drawImage: (_image: unknown, x: number, y: number) => calls.push(state.shadowColor
          ? `${name}:native:${state.shadowColor},${state.shadowBlur},${state.shadowOffsetX},${state.shadowOffsetY}:${x},${y}`
          : `${name}:shadow:${state.filter}:${x},${y}`),
      };
      const filters: unknown[] = [];
      state.save = () => filters.push(state.filter);
      state.restore = () => { state.filter = filters.pop(); };
      return new Proxy(state, {
        get: (target, key) => target[key as string] ?? (() => undefined),
        set: (target, key, value) => { target[key as string] = value; return true; },
      }) as unknown as CanvasRenderingContext2D;
    };
    const previous = globalThis.OffscreenCanvas;
    Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, writable: true, value: class {
      constructor(public width: number, public height: number) { surfaces.push([width, height]); }
      getContext() { return makeContext('mask'); }
    } });
    return { calls, transforms, surfaces, ctx: makeContext('main'), restore: () => {
      Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, writable: true, value: previous });
    } };
  }

  function list(shadow: ShapePrimitive['shadow']): SlideDisplayList {
    return {
      contractVersion: 1, width: 160, height: 160,
      primitives: [{
        kind: 'shape', objectId: 1, name: 'card', x: 40, y: 40, w: 40, h: 40,
        geometry: 'rect', path: [
          { type: 'move', x: 0, y: 0 }, { type: 'line', x: 1, y: 0 },
          { type: 'line', x: 1, y: 1 }, { type: 'line', x: 0, y: 1 }, { type: 'close' },
        ],
        fill: { kind: 'solid', color: '#4472c480' },
        stroke: { color: '#10235b', width: 2 }, shadow,
      }],
    };
  }

  test('scaled shadows transform both axes and retain scaled outline margins', async () => {
    const { calls, transforms, surfaces, ctx, restore } = harness();
    try {
      await paintSlide(ctx, list({ color: '#00000066', scaleX: 2, scaleY: 0.5, dx: -40, dy: 80 }), 3, 1);
      expect(transforms[0]).toEqual([6, 0, 0, 1.5, -216, -36]);
      expect(surfaces).toEqual([[288, 108]]);
      expect(calls).toContain('main:shadow:blur(0px):96,276');
    } finally { restore(); }
  });

  test('shadow work shares a slide budget, fails before allocation, and resets for each paint', async () => {
    const { surfaces, ctx, restore } = harness();
    try {
      const display = list({ color: '#00000066' });
      const shape = display.primitives[0] as ShapePrimitive;
      shape.stroke = undefined;
      const options = { maxShadowPixels: 120 * 120 };
      await paintSlide(ctx, display, 3, 1, options);
      await paintSlide(ctx, display, 3, 1, options);
      expect(surfaces).toHaveLength(2);
      await expect(paintSlide(ctx, display, 3, 1, { maxShadowPixels: 120 * 120 - 1 })).rejects.toThrow('pixel budget');
      expect(surfaces).toHaveLength(2);
      const many = { ...display, primitives: Array.from({ length: 10_000 }, () => shape) };
      await expect(paintSlide(ctx, many, 3, 1, options)).rejects.toThrow('pixel budget');
      expect(surfaces).toHaveLength(3);
      const chart = { kind: 'chart', objectId: 9, x: 0, y: 0, w: 160, h: 160, primitives: [shape, shape] };
      await expect(paintSlide(ctx, { ...display, primitives: [chart] } as SlideDisplayList, 3, 1, options)).rejects.toThrow('pixel budget');
      expect(surfaces).toHaveLength(4);
    } finally { restore(); }
  });

  test('a shadow combines fill and outline alpha and scales blur and offset to the device', async () => {
    const { calls, ctx, restore } = harness();
    try {
      await paintSlide(ctx, list({ color: '#00000066', blur: 8, dx: 6, dy: 6 }), 2, 1.5);
      expect(calls).toEqual([
        'mask:fill', 'mask:stroke', 'mask:tint:source-in:#00000066',
        'main:shadow:blur(12px):126,126', 'main:fill', 'main:stroke',
      ]);
      expect(ctx.filter).toBe('none');
    } finally { restore(); }
  });

  test('a context without filters paints one shadow from the combined source alpha', async () => {
    const { calls, ctx, restore } = harness(false);
    try {
      await paintSlide(ctx, list({ color: '#00000066', blur: 8, dx: 6, dy: 6 }), 2, 1.5);
      expect(calls).toEqual([
        'mask:fill', 'mask:stroke', 'main:native:#00000066,24,-355,18:481,108',
        'main:fill', 'main:stroke',
      ]);
    } finally { restore(); }
  });

  test('an unshadowed shape leaves the shadow state alone', async () => {
    const { calls, ctx, restore } = harness();
    try {
      await paintSlide(ctx, list(undefined), 1, 1);
      expect(calls).toEqual(['main:fill', 'main:stroke']);
    } finally { restore(); }
  });

  test('an unfilled outline casts a shadow even at zero blur and offset', async () => {
    const { calls, ctx, restore } = harness();
    try {
      const display = list({ color: '#00000066' });
      (display.primitives[0] as ShapePrimitive).fill = undefined;
      await paintSlide(ctx, display, 2, 1.5);
      expect(calls).toEqual([
        'mask:stroke', 'mask:tint:source-in:#00000066',
        'main:shadow:blur(0px):108,108', 'main:stroke',
      ]);
    } finally { restore(); }
  });
});

describe('blip colour effects', () => {
  test('biLevel thresholds on Rec. 601 luma and leaves alpha alone', () => {
    const data = new Uint8ClampedArray([0x03, 0xa7, 0xdf, 0x80]);
    applyImageEffects(data, [{ kind: 'biLevel', threshold: 0.5 }]);
    expect([...data]).toEqual([0, 0, 0, 0x80]);

    const light = new Uint8ClampedArray([0x03, 0xa7, 0xdf, 0xff]);
    applyImageEffects(light, [{ kind: 'biLevel', threshold: 0.25 }]);
    expect([...light]).toEqual([255, 255, 255, 0xff]);
  });

  test('duotone interpolates between the two colours by luma', () => {
    const data = new Uint8ClampedArray([0, 0, 0, 0xff, 255, 255, 255, 0xff]);
    applyImageEffects(data, [{ kind: 'duotone', shadow: '#737373ff', highlight: '#ffffffff' }]);
    expect([...data]).toEqual([0x73, 0x73, 0x73, 0xff, 255, 255, 255, 0xff]);
  });

  test('effects apply in list order', () => {
    const ordered: ImageEffect[] = [
      { kind: 'colorChange', from: '#ffffffff', to: '#ffffff00' },
      { kind: 'duotone', shadow: '#000000ff', highlight: '#ff0000ff' },
    ];
    const data = new Uint8ClampedArray([255, 255, 255, 0xff]);
    applyImageEffects(data, ordered);
    expect(data[3]).toBe(0);

    const reversed = new Uint8ClampedArray([255, 255, 255, 0xff]);
    applyImageEffects(reversed, [...ordered].reverse());
    expect(reversed[3]).toBe(0xff);
  });
});

test('colour changes respect useA and preserve transparent and antialiased pixels', () => {
  for (const useAlpha of [undefined, true, false]) {
    const data = new Uint8ClampedArray([
      255, 255, 255, 255, 255, 255, 255, 128, 255, 255, 255, 0, 255, 254, 255, 255,
    ]);
    applyImageEffects(data, [{ kind: 'colorChange', from: '#ffffffff', to: '#ff000000', useAlpha }]);
    expect([...data]).toEqual(useAlpha === false
      ? [255, 0, 0, 255, 255, 0, 0, 128, 255, 0, 0, 0, 255, 254, 255, 255]
      : [255, 0, 0, 0, 255, 255, 255, 128, 255, 255, 255, 0, 255, 254, 255, 255]);
  }
  const data = new Uint8ClampedArray([3, 167, 223, 128]);
  applyImageEffects(data, [{ kind: 'grayscale' }]);
  expect([...data]).toEqual([124, 124, 124, 128]);
});

test('a duotone endpoint modulates alpha instead of replacing it', () => {
  const data = new Uint8ClampedArray([255, 255, 255, 200, 0, 0, 0, 0]);
  applyImageEffects(data, [{ kind: 'duotone', shadow: '#000000ff', highlight: '#ffffff80' }]);
  // The white pixel takes half the highlight's alpha; the transparent one stays transparent.
  expect([...data]).toEqual([255, 255, 255, 100, 0, 0, 0, 0]);

  const opaque = new Uint8ClampedArray([255, 255, 255, 200]);
  applyImageEffects(opaque, [{ kind: 'duotone', shadow: '#000000', highlight: '#ffffff' }]);
  expect([...opaque]).toEqual([255, 255, 255, 200]);
});

test('picture effects reach canvas before cropping without changing the shared source', async () => {
  const original = globalThis.OffscreenCanvas;
  try {
    for (const failure of [null, 'read', 'context'] as const) {
      // A source of its own per case: a recoloured bitmap is cached against the source it
      // came from, so sharing one here would answer the later cases from the first.
      const source = { width: 4, height: 1, pixels: new Uint8ClampedArray([3, 167, 223, 128]) };
      class Surface {
        pixels = new Uint8ClampedArray();
        constructor(public width: number, public height: number) {}
        getContext() {
          if (failure === 'context') throw new Error('context unavailable');
          return {
            drawImage: () => { this.pixels = source.pixels.slice(); },
            getImageData: () => {
              if (failure === 'read') throw new Error('tainted');
              return { data: this.pixels };
            },
            putImageData: () => {},
          };
        }
      }
      globalThis.OffscreenCanvas = Surface as unknown as typeof OffscreenCanvas;
      const draws: unknown[][] = [];
      const ctx = new Proxy({} as CanvasRenderingContext2D, {
        get: (_, key) => key === 'drawImage' ? (...args: unknown[]) => draws.push(args) : () => {},
        set: () => true,
      });
      await paintSlide(ctx, {
        contractVersion: 1,
        width: 100,
        height: 100,
        primitives: [
          { kind: 'image', objectId: 1, name: 'Effect', x: 10, y: 20, w: 40, h: 10,
            assetId: 'image', crop: { left: 0.25 }, effects: [{ kind: 'biLevel', threshold: 0.25 }] },
          { kind: 'image', objectId: 2, name: 'Control', x: 10, y: 40, w: 40, h: 10, assetId: 'image' },
        ],
      }, 1, 1, { resolveImage: async () => source as unknown as CanvasImageSource });
      expect(draws).toHaveLength(2);
      expect(draws[0].slice(1)).toEqual([1, 0, 3, 1, 10, 20, 40, 10]);
      expect([...(draws[0][0] as typeof source).pixels]).toEqual(
        failure ? [3, 167, 223, 128] : [255, 255, 255, 128]
      );
      expect(draws[1][0]).toBe(source);
      expect([...source.pixels]).toEqual([3, 167, 223, 128]);
    }
  } finally {
    if (original === undefined) Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
    else globalThis.OffscreenCanvas = original;
  }
});

/** Installs a fake OffscreenCanvas that records each surface's size and does no pixel work. */
async function withSurfaces(
  run: (surfaces: { width: number; height: number }[]) => Promise<void>
): Promise<void> {
  const original = globalThis.OffscreenCanvas;
  const surfaces: { width: number; height: number }[] = [];
  class Surface {
    constructor(public width: number, public height: number) {
      surfaces.push({ width, height });
    }
    getContext() {
      return {
        drawImage: () => {},
        getImageData: () => ({ data: new Uint8ClampedArray(4) }),
        putImageData: () => {},
      };
    }
  }
  globalThis.OffscreenCanvas = Surface as unknown as typeof OffscreenCanvas;
  try {
    await run(surfaces);
  } finally {
    if (original === undefined) Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
    else globalThis.OffscreenCanvas = original;
  }
}

/** Paints one picture with `effects` from `source` and returns the canvas's drawImage calls. */
async function paintEffect(source: object, effects: ImageEffect[]): Promise<unknown[][]> {
  const draws: unknown[][] = [];
  const ctx = new Proxy({} as CanvasRenderingContext2D, {
    get: (_, key) => key === 'drawImage' ? (...args: unknown[]) => draws.push(args) : () => {},
    set: () => true,
  });
  await paintSlide(ctx, {
    contractVersion: 1,
    width: 100,
    height: 100,
    primitives: [
      { kind: 'image', objectId: 1, name: 'Effect', x: 10, y: 20, w: 40, h: 10, assetId: 'image', effects },
    ],
  }, 1, 1, { resolveImage: async () => source as unknown as CanvasImageSource });
  return draws;
}

test('a recolouring is reused for the same source and effects and redone for another list', async () => {
  await withSurfaces(async (surfaces) => {
    const source = { width: 4, height: 1 };
    const biLevel: ImageEffect[] = [{ kind: 'biLevel', threshold: 0.25 }];
    const first = await paintEffect(source, biLevel);
    const second = await paintEffect(source, biLevel);
    expect(surfaces).toHaveLength(1);
    expect(first[0][0]).not.toBe(source);
    expect(second[0][0]).toBe(first[0][0]);
    await paintEffect(source, [{ kind: 'grayscale' }]);
    expect(surfaces).toHaveLength(2);
    await paintEffect({ width: 4, height: 1 }, biLevel);
    expect(surfaces).toHaveLength(3);
  });
});

test('an oversized bitmap is recoloured within the pixel cap and drawn back at picture size', async () => {
  await withSurfaces(async (surfaces) => {
    const draws = await paintEffect({ width: 8192, height: 8192 }, [{ kind: 'grayscale' }]);
    expect(surfaces).toEqual([{ width: 5792, height: 5792 }]);
    expect(draws[0].slice(1)).toEqual([0, 0, 5792, 5792, 10, 20, 40, 10]);
    await paintEffect({ width: 8192, height: 4096 }, [{ kind: 'grayscale' }]);
    expect(surfaces[1]).toEqual({ width: 8192, height: 4096 });
  });
});

test('a video frame is recoloured on every paint rather than kept', async () => {
  await withSurfaces(async (surfaces) => {
    const video = { videoWidth: 4, videoHeight: 1 };
    const first = await paintEffect(video, [{ kind: 'grayscale' }]);
    const second = await paintEffect(video, [{ kind: 'grayscale' }]);
    expect(surfaces).toEqual([{ width: 4, height: 1 }, { width: 4, height: 1 }]);
    expect(first[0][0]).not.toBe(video);
    expect(second[0][0]).not.toBe(first[0][0]);
  });
});

test('retained recolourings stay within the pixel budget, dropping the least recently used', async () => {
  await withSurfaces(async (surfaces) => {
    const effects: ImageEffect[] = [{ kind: 'grayscale' }];
    const a = { width: 8192, height: 4096 };
    const b = { width: 8192, height: 4096 };
    const c = { width: 8192, height: 4096 };
    for (const source of [a, b, c]) await paintEffect(source, effects);
    expect(surfaces).toHaveLength(3);
    await paintEffect(b, effects);
    expect(surfaces).toHaveLength(3);
    await paintEffect(a, effects);
    expect(surfaces).toHaveLength(4);
    await paintEffect(b, effects);
    expect(surfaces).toHaveLength(4);
    await paintEffect(c, effects);
    expect(surfaces).toHaveLength(5);
  });
});

describe('PPTX text overflow', () => {
  function harness() {
    const calls: string[] = [];
    const ctx = new Proxy(
      {
        clip: () => calls.push('clip'),
        fillText: (t: string) => calls.push(`text:${t}`),
        rect: () => calls.push('rect'),
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
    return { calls, ctx };
  }

  function list(overflow: boolean): SlideDisplayList {
    return {
      contractVersion: 1,
      width: 320,
      height: 180,
      primitives: [
        {
          kind: 'textBox',
          objectId: 1,
          x: 10,
          y: 10,
          w: 100,
          h: 20,
          anchor: 'top',
          paragraphs: [],
          overflow,
          lines: [
            {
              x: 10,
              y: 10,
              width: 100,
              height: 40,
              baseline: 30,
              start: 0,
              end: 5,
              caretStops: [],
              runs: [
                {
                  text: 'spill',
                  start: 0,
                  end: 5,
                  x: 10,
                  width: 100,
                  fontId: 0,
                  fontFamily: 'Arial',
                  fontSizePx: 40,
                  bold: false,
                  italic: false,
                  underline: false,
                  color: '#000000',
                  glyphs: [],
                },
              ],
            },
          ],
        },
      ],
    } as unknown as SlideDisplayList;
  }

  test('text taller than its box is not clipped to it', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list(true), 1, 1, {});
    expect(calls).not.toContain('clip');
    expect(calls).toContain('text:spill');
  });

  test('text that fits is still clipped to its box', async () => {
    const { calls, ctx } = harness();
    await paintSlide(ctx, list(false), 1, 1, {});
    expect(calls).toContain('clip');
  });
});

test('paints script glyphs and underlines at their shifted baselines', async () => {
  const text: [string, number][] = [];
  const underlines: number[] = [];
  const ctx = new Proxy({
    fillText: (value: string, _x: number, y: number) => text.push([value, y]),
    fillRect: (_x: number, y: number) => underlines.push(y),
  } as Record<string, unknown>, {
    get: (target, key) => target[key as string] ?? (() => undefined),
    set: (target, key, value) => { target[key as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const list: SlideDisplayList = {
    contractVersion: 1, width: 200, height: 100,
    primitives: [{
      kind: 'textBox', objectId: 1, x: 0, y: 0, w: 200, h: 100,
      anchor: 'top', paragraphs: [], lines: [{
        x: 10, y: 20, width: 60, height: 50, baseline: 50, start: 0, end: 3,
        caretStops: [], runs: [6, -5, undefined].map((offset, i) => ({
          text: ['up', 'down', 'base'][i], start: i, end: i + 1, x: 10 + i * 20, width: 20,
          fontId: 1, fontFamily: 'Arial', fontSizePx: 10, bold: false, italic: false,
          underline: true, color: '#475467', baselineOffsetPx: offset, glyphs: [],
        })),
      }],
    }],
  };
  await paintSlide(ctx, list, 1);
  expect(text).toEqual([['up', 44], ['down', 55], ['base', 50]]);
  expect(underlines).toHaveLength(3);
  [44.8, 55.8, 50.8].forEach((y, i) => expect(underlines[i]).toBeCloseTo(y));
});
