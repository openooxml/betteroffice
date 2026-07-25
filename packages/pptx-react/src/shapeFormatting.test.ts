import { describe, expect, it } from 'bun:test';
import type { ShapeSnapshot } from '@betteroffice/pptx';
import { shapeFormattingFromShape } from './shapeFormatting';

describe('shapeFormattingFromShape', () => {
  it('exposes direct shape styling in toolbar units', () => {
    const formatting = shapeFormattingFromShape({
      ...shape,
      geometry: 'roundRect',
      adjustValues: { adj: 0.25 },
      fill: { type: 'solid', color: { rgb: '3367D6' } },
      resolvedFillColor: '#3367D6',
      outline: { width: 38_100, color: { rgb: 'EA4335' } },
      resolvedOutlineColor: '#EA4335',
    });

    expect(formatting).toEqual({
      geometry: 'roundRect',
      fillColor: '#3367D6',
      strokeColor: '#EA4335',
      strokeWidthPt: 3,
      adjustments: { adj: 0.25 },
    });
  });

  it('uses presentation-resolved colors while retaining raw theme references', () => {
    const formatting = shapeFormattingFromShape({
      ...shape,
      fill: { type: 'solid', color: { themeColor: 'accent1' } },
      resolvedFillColor: '#123456',
      outline: { color: { themeColor: 'accent2' } },
      resolvedOutlineColor: '#ABCDEF',
    });

    expect(formatting.fillColor).toBe('#123456');
    expect(formatting.strokeColor).toBe('#ABCDEF');
  });

  it('recognizes explicit no-fill and no-line styling', () => {
    expect(shapeFormattingFromShape({
      ...shape,
      fill: { type: 'none' },
      outline: {},
    })).toMatchObject({
      fillColor: null,
      strokeColor: null,
      strokeWidthPt: null,
    });
  });
});

const shape: ShapeSnapshot = {
  id: 'shape',
  sourceId: 0,
  kind: 'shape',
  name: 'Shape',
  x: 0,
  y: 0,
  width: 1_000_000,
  height: 1_000_000,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  geometry: 'rect',
  adjustValues: {},
  placeholder: null,
  fill: null,
  resolvedFillColor: null,
  outline: null,
  resolvedOutlineColor: null,
  mediaPartPath: null,
  graphic: null,
  textStories: [],
  children: [],
};
