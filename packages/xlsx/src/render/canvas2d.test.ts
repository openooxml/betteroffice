import { describe, expect, it } from 'bun:test';
import fixture from '../../test-fixtures/path-backend-agreement.json';
import type { DisplayList, GeometryPathCommand } from '../display-list/types';
import { paintDisplayList } from './canvas2d';

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
      color: '#000000', style,
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
