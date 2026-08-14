import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  buildDisplayListJson,
  preloadLayoutWasm,
  registerMeasureFont,
} from '../../wasm/layout';
import { drawPrimitive } from './canvasBackend';
import type {
  DisplayList,
  DisplayPaintClip,
  GlyphRunPrimitive,
  TextRunPrimitive,
} from './displayList';
import type { GlyphCache } from './glyphCache';

interface Bounds {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

interface PaintedBounds extends Bounds {
  naturalLeft: number;
  naturalRight: number;
  naturalTop: number;
  naturalBottom: number;
}

interface TextInk {
  width: number;
  ascent: number;
  descent: number;
}

interface TransformState {
  x: number;
  y: number;
  scaleX: number;
  scaleY: number;
}

interface CanvasState {
  clip: Bounds;
  transform: TransformState;
}

interface FakeGlyphPath {
  width: number;
  height: number;
}

const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

let fontId = -1;

beforeAll(async () => {
  await preloadLayoutWasm();
  fontId = registerMeasureFont(new Uint8Array(readFileSync(FONT)));
});

function textRun(
  text: string,
  x: number,
  width: number,
  family: string,
  rtl: boolean = false,
  paintClip?: DisplayPaintClip
): TextRunPrimitive {
  return {
    kind: 'text',
    text,
    x,
    baselineY: 20,
    width,
    paintClip,
    font: `16px "${family}"`,
    color: '#000000',
    rtl,
  };
}

function intersect(a: Bounds, b: Bounds): Bounds {
  const left = Math.max(a.left, b.left);
  const right = Math.max(left, Math.min(a.right, b.right));
  const top = Math.max(a.top, b.top);
  const bottom = Math.max(top, Math.min(a.bottom, b.bottom));
  return { left, right, top, bottom };
}

function recordingContext(textInk: ReadonlyMap<string, TextInk>): {
  ctx: CanvasRenderingContext2D;
  paints: PaintedBounds[];
} {
  const paints: PaintedBounds[] = [];
  const stack: CanvasState[] = [];
  let state: CanvasState = {
    clip: { left: -Infinity, right: Infinity, top: -Infinity, bottom: Infinity },
    transform: { x: 0, y: 0, scaleX: 1, scaleY: 1 },
  };
  let path: Bounds | undefined;

  const transformedBounds = (x: number, y: number, width: number, height: number): Bounds => {
    const x1 = state.transform.x + x * state.transform.scaleX;
    const x2 = state.transform.x + (x + width) * state.transform.scaleX;
    const y1 = state.transform.y + y * state.transform.scaleY;
    const y2 = state.transform.y + (y + height) * state.transform.scaleY;
    return {
      left: Math.min(x1, x2),
      right: Math.max(x1, x2),
      top: Math.min(y1, y2),
      bottom: Math.max(y1, y2),
    };
  };

  const record = (natural: Bounds): void => {
    const painted = intersect(natural, state.clip);
    paints.push({
      ...painted,
      naturalLeft: natural.left,
      naturalRight: natural.right,
      naturalTop: natural.top,
      naturalBottom: natural.bottom,
    });
  };

  const ctx = {
    save(): void {
      stack.push({
        clip: { ...state.clip },
        transform: { ...state.transform },
      });
    },
    restore(): void {
      state = stack.pop() ?? state;
    },
    beginPath(): void {
      path = undefined;
    },
    rect(x: number, y: number, width: number, height: number): void {
      path = transformedBounds(x, y, width, height);
    },
    clip(): void {
      if (path) state.clip = intersect(state.clip, path);
    },
    translate(x: number, y: number): void {
      state.transform.x += x * state.transform.scaleX;
      state.transform.y += y * state.transform.scaleY;
    },
    scale(x: number, y: number): void {
      state.transform.scaleX *= x;
      state.transform.scaleY *= y;
    },
    fill(pathToFill: Path2D): void {
      const glyph = pathToFill as unknown as FakeGlyphPath;
      record(transformedBounds(0, 0, glyph.width, glyph.height));
    },
    fillText(text: string, x: number, y: number): void {
      const ink = textInk.get(text) ?? { width: 0, ascent: 0, descent: 0 };
      record(transformedBounds(x, y - ink.ascent, ink.width, ink.ascent + ink.descent));
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, paints };
}

function syntheticMixedFamilyDisplayList(): DisplayList {
  const first = '@@@@@@@@@@@@@@@@@@@@';
  const input = {
    measured: [
      {
        block: {
          kind: 'paragraph',
          id: 0,
          runs: [
            {
              kind: 'text',
              text: first,
              pmStart: 1,
              pmEnd: 21,
              fontFamily: 'Liberation Sans',
              fontSize: 12,
            },
            {
              kind: 'text',
              text: 'next',
              pmStart: 21,
              pmEnd: 25,
              fontFamily: 'Unregistered Face',
              fontSize: 12,
            },
          ],
          attrs: { defaultFontFamily: 'Liberation Sans', defaultFontSize: 12 },
          pmStart: 0,
          pmEnd: 26,
        },
        measure: {
          kind: 'paragraph',
          totalHeight: 18.4,
          lines: [
            {
              headRun: 0,
              headChar: 0,
              tailRun: 1,
              tailChar: 4,
              width: 384,
              ascent: 12.8,
              descent: 3.2,
              lineHeight: 18.4,
              syntheticFallback: true,
            },
          ],
        },
      },
    ],
    fontChains: { 'liberation sans|0|0': [fontId] },
    options: {},
    layout: {
      pages: [
        {
          size: { w: 816, h: 1056 },
          margins: { top: 96, right: 96, bottom: 96, left: 96 },
          number: 1,
          fragments: [
            {
              kind: 'paragraph',
              blockId: 0,
              x: 96,
              y: 96,
              width: 624,
              height: 18.4,
              fromLine: 0,
              toLine: 1,
              pmStart: 0,
              pmEnd: 26,
            },
          ],
        },
      ],
    },
  };
  return JSON.parse(buildDisplayListJson(JSON.stringify(input))) as DisplayList;
}

describe('Canvas text-run slot clipping', () => {
  it('clips a Rust-shaped run from a synthetic mixed-family line', async () => {
    const list = syntheticMixedFamilyDisplayList();
    const glyphRun = list.pages[0].primitives.find(
      (primitive): primitive is GlyphRunPrimitive =>
        primitive.kind === 'glyphRun' && primitive.text.startsWith('@')
    );
    const nextRun = list.pages[0].primitives.find(
      (primitive): primitive is TextRunPrimitive =>
        primitive.kind === 'text' && primitive.text === 'next'
    );
    if (!glyphRun?.paintClip || !nextRun?.paintClip) {
      throw new Error('synthetic line did not emit paint clips');
    }

    const clipRight = glyphRun.paintClip.x + glyphRun.paintClip.w;
    const shapedRight = Math.max(
      ...glyphRun.glyphs.map((glyph) => glyph.x + (glyph.advance ?? 0))
    );
    expect(glyphRun.paintClip).toEqual({ x: 96, w: 320 });
    expect(nextRun.paintClip).toEqual({ x: 416, w: 64 });
    expect(shapedRight).toBeGreaterThan(clipRight);

    const glyphPath = { width: 2048, height: 2048 } as unknown as Path2D;
    const glyphCache = {
      get: () => ({ path: glyphPath, upem: 2048 }),
    } as unknown as GlyphCache;
    const { ctx, paints } = recordingContext(
      new Map([['next', { width: 64, ascent: 12.8, descent: 3.2 }]])
    );
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await drawPrimitive(ctx, glyphRun, { glyphCache });
      await drawPrimitive(ctx, nextRun);
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }

    const glyphPaints = paints.slice(0, glyphRun.glyphs.length);
    expect(glyphPaints.some((paint) => paint.naturalRight > nextRun.x)).toBe(true);
    expect(glyphPaints.every((paint) => paint.right <= nextRun.x)).toBe(true);
    expect(paints.at(-1)?.left).toBe(nextRun.x);
  });

  it('keeps wide browser glyphs out of the following run', async () => {
    const probes = [
      {
        text: '@@@@@@@@@@',
        family: 'Liberation Sans',
        reservedEm: 10,
        naturalEm: 10.151,
      },
      { text: '⸻', family: 'Noto Sans SC', reservedEm: 1, naturalEm: 2.459 },
      {
        text: '﷽',
        family: 'Noto Sans Arabic',
        reservedEm: 1,
        naturalEm: 9.779,
        rtl: true,
      },
    ];

    for (const probe of probes) {
      const slotWidth = probe.reservedEm * 16;
      const naturalWidth = probe.naturalEm * 16;
      const { ctx, paints } = recordingContext(
        new Map([
          [probe.text, { width: naturalWidth, ascent: 12.8, descent: 3.2 }],
          ['next', { width: 16, ascent: 12.8, descent: 3.2 }],
        ])
      );

      await drawPrimitive(
        ctx,
        textRun(probe.text, 0, slotWidth, probe.family, probe.rtl, {
          x: 0,
          w: slotWidth,
        })
      );
      await drawPrimitive(
        ctx,
        textRun('next', slotWidth, 16, probe.family, false, { x: slotWidth, w: 16 })
      );

      expect(paints).toHaveLength(2);
      expect(paints[0].naturalRight).toBeGreaterThan(paints[1].left);
      expect(paints[0].right).toBeLessThanOrEqual(paints[1].left);
    }
  });

  it('leaves accented capitals and emphasis marks vertically unbounded', async () => {
    const run = textRun('É', 0, 16, 'Liberation Sans', false, { x: 0, w: 16 });
    run.baselineY = 100;
    run.emphasisMark = 'dot';
    const { ctx, paints } = recordingContext(
      new Map([
        ['É', { width: 12, ascent: 13.872, descent: 3.2 }],
        ['•', { width: 8, ascent: 4, descent: 4 }],
      ])
    );

    await drawPrimitive(ctx, run);

    expect(paints).toHaveLength(2);
    expect(paints[0].naturalTop).toBeLessThan(run.baselineY - 12.8);
    expect(paints[0].top).toBe(paints[0].naturalTop);
    expect(paints[0].bottom).toBe(paints[0].naturalBottom);
    expect(paints[1].naturalTop).toBeCloseTo(82.4);
    expect(paints[1].top).toBe(paints[1].naturalTop);
    expect(paints[1].bottom).toBe(paints[1].naturalBottom);
  });

  it('leaves registered-face browser effects unclipped', async () => {
    const { ctx, paints } = recordingContext(
      new Map([['wide', { width: 64, ascent: 12.8, descent: 3.2 }]])
    );
    const run = textRun('wide', 0, 16, 'Liberation Sans');
    run.textShadow = 'shadow';

    await drawPrimitive(ctx, run);

    expect(paints).toHaveLength(1);
    expect(paints[0].right).toBe(paints[0].naturalRight);
    expect(paints[0].right).toBeGreaterThan(run.width);
  });
});
