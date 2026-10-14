import { describe, expect, it } from 'bun:test';
import fixture from '../../test-fixtures/path-backend-agreement.json';
import type { DisplayList, GeometryPathCommand, PatternType, TextCmd } from '../display-list/types';
import { paintDisplayList } from './canvas2d';
import { PATTERN_MASKS } from './patterns';

describe('Canvas2D pattern masks', () => {
  it('cover the specified share of each 8x8 tile', () => {
    // gray shares are the spec's percentages of 64 pixels; hatches are 1-px (light) or
    // 2-px (dark) lines every 4 px, and grids and trellises count each crossing once
    const expected: Record<PatternType, number> = {
      gray0625: 64 / 16, gray125: 64 / 8, lightGray: 64 / 4, mediumGray: 64 / 2, darkGray: 64 * 3 / 4,
      lightHorizontal: 16, lightVertical: 16, lightDown: 16, lightUp: 16,
      lightGrid: 16 + 16 - 4, lightTrellis: 16 + 16 - 8,
      darkHorizontal: 32, darkVertical: 32, darkDown: 32, darkUp: 32,
      darkGrid: 32 + 32 - 16, darkTrellis: 32 + 32 - 16,
    };
    const set = (rows: readonly number[]) =>
      rows.reduce((sum, row) => sum + row.toString(2).split('').filter((bit) => bit === '1').length, 0);
    expect(Object.fromEntries(Object.entries(PATTERN_MASKS).map(([pattern, rows]) => [pattern, set(rows)])))
      .toEqual(expected);
  });
});

describe('Canvas2D double lines', () => {
  it('keep the released rails and gap when the command carries no rail width', () => {
    for (const dpr of [1, 2]) for (const width of [0.5, 1, 2]) {
      const strokes: number[][] = [];
      let points: number[] = [];
      const context = {
        lineWidth: 0, strokeStyle: '', lineCap: '',
        save() {}, restore() {}, setTransform() {}, clearRect() {}, rect() {}, clip() {}, setLineDash() {},
        beginPath() { points = []; },
        moveTo(x: number, y: number) { points.push(x, y); },
        lineTo(x: number, y: number) { points.push(x, y); },
        stroke() { strokes.push([this.lineWidth, ...points]); },
      };
      const dl = { width: 100, height: 40, commands: [
        { op: 'line', x1: 0, y1: 10, x2: 100, y2: 10, width, color: '#000000', style: 'double' },
        { op: 'line', x1: 20, y1: 0, x2: 20, y2: 40, width, color: '#000000', style: 'double' },
      ] } as DisplayList;
      paintDisplayList(context as unknown as CanvasRenderingContext2D, dl, dpr);
      const rail = Math.max(width * 0.6, 0.5);
      const offset = Math.max(width * 0.8, 0.8);
      expect(strokes).toEqual([
        [rail, 0, 10 - offset, 100, 10 - offset],
        [rail, 0, 10 + offset, 100, 10 + offset],
        [rail, 20 - offset, 0, 20 - offset, 40],
        [rail, 20 + offset, 0, 20 + offset, 40],
      ]);
    }
  });
});

describe('Canvas2D pattern borders', () => {
  it('paints all six dash styles and a three-pixel double border', () => {
    const dashes: number[][] = [];
    const widths: number[] = [];
    const context = {
      lineWidth: 0, strokeStyle: '', lineCap: '',
      save() {}, restore() {}, setTransform() {}, clearRect() {},
      beginPath() {}, rect() {}, clip() {}, moveTo() {}, lineTo() {},
      setLineDash(value: number[]) { dashes.push(value); },
      stroke() { widths.push(this.lineWidth); },
    };
    const styles = ['mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot', 'double'];
    const dl = { width: 100, height: 40, commands: styles.map((style) => ({
      op: 'line', x1: 0, y1: 10, x2: 100, y2: 10,
      width: style === 'double' ? 3 : style.startsWith('medium') ? 2 : 1,
      color: '#000000', style, ...(style === 'double' ? { rail: 1 } : {}),
    })) } as DisplayList;
    paintDisplayList(context as unknown as CanvasRenderingContext2D, dl, 1);
    expect(dashes).toEqual([[8, 4], [4, 2, 1, 2], [8, 4, 2, 4], [4, 2, 1, 2, 1, 2], [8, 4, 2, 4, 2, 4], [4, 2, 1, 2], []]);
    expect(widths).toEqual([2, 1, 2, 1, 2, 1, 1, 1]);
  });

  it('caches eight device-pixel tiles and anchors them to the sheet at every dpr', () => {
    const tiles: number[][] = [];
    const transforms: number[][] = [];
    const patternTransforms: object[] = [];
    const patterns: object[] = [];
    const original = globalThis.OffscreenCanvas;
    class TileCanvas {
      constructor(public width: number, public height: number) {}
      getContext() {
        return {
          createImageData: () => ({ data: new Uint8ClampedArray(256) }),
          putImageData: (image: ImageData) => tiles.push(Array.from(image.data)),
        };
      }
    }
    globalThis.OffscreenCanvas = TileCanvas as unknown as typeof OffscreenCanvas;
    const context = {
      fillStyle: '',
      save() {}, restore() {}, clearRect() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
      setTransform(...value: number[]) { transforms.push(value); },
      createPattern() {
        const pattern = { setTransform(value: object) { patternTransforms.push(value); } };
        patterns.push(pattern);
        return pattern;
      },
    };
    const dl = { width: 32, height: 8, commands: [0, 13].map((x) => ({
      op: 'patternFill', x, y: 0, w: x === 0 ? 13 : 19, h: 8,
      pattern: 'lightUp', fg: '#000000', bg: '#ffffff', originX: -3, originY: -2,
    })) } as unknown as DisplayList;
    try {
      for (const dpr of [1, 2]) paintDisplayList(context as unknown as CanvasRenderingContext2D, dl, dpr);
      paintDisplayList(context as unknown as CanvasRenderingContext2D, dl, 2, { x: 37.5, y: 19.5 });
      expect(tiles).toHaveLength(2);
      expect(patterns).toHaveLength(2);
      for (const tile of tiles) {
        for (let y = 0; y < 8; y += 1) for (let x = 0; x < 8; x += 1) {
          expect(tile[(y * 8 + x) * 4]).toBe((x + y) % 4 === 0 ? 0 : 255);
        }
      }
      expect(patternTransforms).toContainEqual({ a: 1, b: 0, c: 0, d: 1, e: -3, f: -2 });
      expect(patternTransforms).toContainEqual({ a: 0.5, b: 0, c: 0, d: 0.5, e: -3, f: -2 });
      expect(patternTransforms).toContainEqual({ a: 0.5, b: 0, c: 0, d: 0.5, e: -2.75, f: -1.75 });
      expect(transforms).toContainEqual([2, 0, 0, 2, 0, 0]);
    } finally {
      globalThis.OffscreenCanvas = original;
    }
  });
});

interface PathTrace {
  verb: string;
  points: number[];
}

describe('Canvas2D paths', () => {
  it('follows the shared raster and canvas trace contract', () => {
    const trace: PathTrace[] = [];
    let fills = 0;
    let strokes = 0;
    const context = {
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 0,
      save() {},
      restore() {},
      setTransform() {},
      clearRect() {},
      beginPath() {},
      rect() {},
      clip() {},
      setLineDash() {},
      moveTo(x: number, y: number) {
        trace.push({ verb: 'move', points: [x, y] });
      },
      lineTo(x: number, y: number) {
        trace.push({ verb: 'line', points: [x, y] });
      },
      quadraticCurveTo(cpx: number, cpy: number, x: number, y: number) {
        trace.push({ verb: 'quad', points: [cpx, cpy, x, y] });
      },
      bezierCurveTo(
        cp1x: number,
        cp1y: number,
        cp2x: number,
        cp2y: number,
        x: number,
        y: number
      ) {
        trace.push({ verb: 'cubic', points: [cp1x, cp1y, cp2x, cp2y, x, y] });
      },
      closePath() {
        trace.push({ verb: 'close', points: [] });
      },
      fill() {
        fills += 1;
      },
      stroke() {
        strokes += 1;
      },
    };
    const displayList: DisplayList = {
      width: 24,
      height: 24,
      commands: [
        {
          op: 'path',
          commands: fixture.commands as GeometryPathCommand[],
          fill: '#123456',
          stroke: { color: '#654321', width: 2 },
          clip: { x: 1, y: 1, w: 22, h: 22 },
        },
      ],
    };

    paintDisplayList(context as unknown as CanvasRenderingContext2D, displayList, 1);

    expect(trace).toEqual(fixture.trace);
    expect({ fills, strokes }).toEqual({ fills: 1, strokes: 1 });
  });
});

describe('Canvas2D text', () => {
  function formattedTrace(text: string, numberFormat: unknown, width: number, measured: string[] = []) {
    const trace: { text: string; x: number; align: string }[] = [];
    const context = {
      font: '', fillStyle: '', textAlign: '', textBaseline: '',
      save() {}, restore() {}, setTransform() {}, clearRect() {},
      beginPath() {}, rect() {}, clip() {},
      measureText(value: string) {
        measured.push(value);
        return { width: [...value].reduce((sum, char) => sum + (char === ' ' ? 3 : char === '(' ? 5 : char === ')' ? 7 : 10), 0) };
      },
      fillText(value: string, x: number) { trace.push({ text: value, x, align: this.textAlign }); },
    };
    paintDisplayList(context as unknown as CanvasRenderingContext2D, {
      width, height: 20,
      commands: [{ op: 'text', x: width - 2, y: 15, text, fontSize: 11,
        color: '#000000', align: 'right', clip: { x: 0, y: 0, w: width, h: 20 },
        numberFormat,
      } as unknown as TextCmd],
    }, 1);
    return trace;
  }

  it('uses character widths and keeps accounting currency and digits at both edges', () => {
    for (const width of [100, 160]) {
      const trace = formattedTrace(' €12 ', {
        left: 2, width: width - 4,
        parts: [
          { offset: 0, kind: 'padding', character: '(' },
          { offset: 4, kind: 'fill', character: ' ' },
          { offset: 6, kind: 'padding', character: ')' },
        ],
      }, width);
      expect(trace.find((run) => run.text === '€')?.x).toBe(7);
      expect(trace.find((run) => run.text === '12')?.x).toBe(width - 29);
      expect(trace.every((run) => run.align === 'left')).toBe(true);
    }
  });

  it('repeats non-space fill at its original position', () => {
    const trace = formattedTrace('$12', {
      left: 2, width: 96, parts: [{ offset: 1, kind: 'fill', character: '.' }],
    }, 100);
    expect(trace.map((run) => run.text)).toEqual(['$', '......', '12']);
    expect(trace[2].x).toBe(78);
  });

  it('pads question placeholders with digit advances', () => {
    const trace = formattedTrace(' 1', {
      left: 2, width: 96, parts: [{ offset: 0, kind: 'padding', character: '0' }],
    }, 100);
    expect(trace[0]).toEqual({ text: '1', x: 88, align: 'left' });
  });

  it('fills a narrow date cell with hashes', () => {
    expect(formattedTrace('1/1/2020', { left: 2, width: 26, parts: [], overflow: true }, 30))
      .toEqual([{ text: '##', x: 8, align: 'left' }]);
  });

  it('shrinks General decimals before scientific and hashes', () => {
    expect(formattedTrace('1234.56789', { left: 2, width: 76, parts: [], general: 1234.56789, overflow: true }, 80)[0].text).toBe('1234.57');
    expect(formattedTrace('123456789', { left: 2, width: 66, parts: [], general: 123456789, overflow: true }, 70)[0].text).toBe('1E+08');
    expect(formattedTrace('123456789', { left: 2, width: 26, parts: [], general: 123456789, overflow: true }, 30)[0].text).toBe('##');
  });

  it('rounds fitted scientific Generals from fifteen significant digits', () => {
    for (const value of [1.005e12, 1.005e12 - 2 ** -13]) {
      expect(formattedTrace('1.005E+12', { left: 2, width: 86, parts: [], general: value, overflow: true }, 90)[0].text).toBe('1.01E+12');
    }
  });

  it('paints a fitting number and its underline over the extent its right-aligned text had', () => {
    const runs: { text: string; x: number; align: string }[] = [];
    const rects: number[][] = [];
    const context = {
      font: '', fillStyle: '', textAlign: '', textBaseline: '',
      save() {}, restore() {}, setTransform() {}, clearRect() {},
      beginPath() {}, rect() {}, clip() {},
      measureText(value: string) { return { width: value.length * 10 }; },
      fillText(value: string, x: number) { runs.push({ text: value, x, align: this.textAlign }); },
      fillRect(x: number, y: number, w: number, h: number) { rects.push([x, y, w, h]); },
    };
    paintDisplayList(context as unknown as CanvasRenderingContext2D, {
      width: 100, height: 20,
      commands: [{ op: 'text', x: 98, y: 15, text: '1234', fontSize: 9, color: '#000000',
        align: 'right', underline: true, clip: { x: 0, y: 0, w: 100, h: 20 },
        numberFormat: { left: 2, width: 96, parts: [], overflow: true },
      } as unknown as TextCmd],
    }, 1);
    const px = 9 * (96 / 72);
    expect(runs).toEqual([{ text: '1234', x: 98 - 40, align: 'left' }]);
    expect(rects).toEqual([[98 - 40, 15 + px * 0.1, 40, Math.max(px * 0.05, 0.5)]]);
  });

  it('paints a layout without padding, fill or overflow at its aligned edge without measuring', () => {
    for (const [align, x] of [['left', 2], ['center', 50], ['right', 98]] as const) {
      const runs: { text: string; x: number; align: string }[] = [];
      const context = {
        font: '', fillStyle: '', textAlign: '', textBaseline: '',
        save() {}, restore() {}, setTransform() {}, clearRect() {},
        beginPath() {}, rect() {}, clip() {},
        measureText(value: string): never { throw new Error(`measured ${value}`); },
        fillText(value: string, at: number) { runs.push({ text: value, x: at, align: this.textAlign }); },
      };
      paintDisplayList(context as unknown as CanvasRenderingContext2D, {
        width: 100, height: 20,
        commands: [{ op: 'text', x, y: 15, text: '1234', fontSize: 11, color: '#000000', align,
          clip: { x: 0, y: 0, w: 100, h: 20 }, numberFormat: { left: 2, width: 96, parts: [] },
        } as unknown as TextCmd],
      }, 1);
      expect(runs).toEqual([{ text: '1234', x, align }]);
    }
  });

  it('measures padded, filled and overflow-checked layouts', () => {
    for (const numberFormat of [
      { left: 2, width: 96, parts: [{ offset: 0, kind: 'padding', character: '(' }] },
      { left: 2, width: 96, parts: [{ offset: 1, kind: 'fill', character: '.' }] },
      { left: 2, width: 96, parts: [], general: 1234, overflow: true },
    ]) {
      const measured: string[] = [];
      formattedTrace(' 1234', numberFormat, 100, measured);
      expect(measured.length).toBeGreaterThan(0);
    }
  });

  it('uses point-sized fonts for cells, ghosts, and charts at every display scale', () => {
    const fonts: string[] = [];
    const transforms: number[][] = [];
    const context = {
      font: '',
      fillStyle: '',
      textAlign: '',
      textBaseline: '',
      beginPath() {}, rect() {}, clip() {},
      save() {},
      restore() {},
      setTransform(...values: number[]) { transforms.push(values); },
      clearRect() {},
      fillText() {
        fonts.push(this.font);
      },
    };
    const displayList: DisplayList = {
      width: 120,
      height: 80,
      commands: [undefined, false, true].map((chart) => ({
        op: 'text',
        x: 0,
        y: 20,
        text: 'Quarterly revenue',
        fontSize: 12,
        fontFamily: 'Calibri',
        color: '#000000',
        chart,
      })),
    };
    for (const scale of [1, 2]) {
      paintDisplayList(
        context as unknown as CanvasRenderingContext2D,
        displayList,
        scale
      );
    }
    paintDisplayList(context as unknown as CanvasRenderingContext2D, { width: 1, height: 1, commands: [] }, 2, { x: 37.5, y: 37.5 });
    expect(transforms[transforms.length - 1]).toEqual([2, 0, 0, 2, 37.5, 37.5]);
    expect(fonts.map((font) => Number.parseFloat(font))).toEqual([
      16, 16, 16, 16, 16, 16,
    ]);
  });
});
