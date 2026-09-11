import { describe, expect, it } from 'bun:test';
import fixture from '../../test-fixtures/path-backend-agreement.json';
import type { DisplayList, GeometryPathCommand } from '../display-list/types';
import { paintDisplayList } from './canvas2d';

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
