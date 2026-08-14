import { describe, expect, it } from 'bun:test';

import { drawPrimitive } from './canvasBackend';
import type { TextRunPrimitive } from './displayList';

interface PaintedInterval {
  left: number;
  right: number;
  naturalRight: number;
}

interface HorizontalClip {
  left: number;
  right: number;
}

function textRun(
  text: string,
  x: number,
  width: number,
  family: string,
  rtl: boolean = false
): TextRunPrimitive {
  return {
    kind: 'text',
    text,
    x,
    baselineY: 20,
    width,
    font: `16px "${family}"`,
    color: '#000000',
    rtl,
  };
}

function recordingContext(naturalWidths: ReadonlyMap<string, number>): {
  ctx: CanvasRenderingContext2D;
  paints: PaintedInterval[];
} {
  const paints: PaintedInterval[] = [];
  const clips: HorizontalClip[] = [];
  let clip: HorizontalClip = { left: -Infinity, right: Infinity };
  let path: HorizontalClip | undefined;
  const ctx = {
    save(): void {
      clips.push({ ...clip });
    },
    restore(): void {
      clip = clips.pop() ?? clip;
    },
    beginPath(): void {
      path = undefined;
    },
    rect(x: number, _y: number, width: number, _height: number): void {
      path = {
        left: Math.min(x, x + width),
        right: Math.max(x, x + width),
      };
    },
    clip(): void {
      if (!path) return;
      clip = {
        left: Math.max(clip.left, path.left),
        right: Math.min(clip.right, path.right),
      };
    },
    fillText(text: string, x: number): void {
      const naturalRight = x + (naturalWidths.get(text) ?? 0);
      paints.push({
        left: Math.max(x, clip.left),
        right: Math.max(
          Math.max(x, clip.left),
          Math.min(naturalRight, clip.right)
        ),
        naturalRight,
      });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, paints };
}

describe('Canvas text-run slot clipping', () => {
  it('keeps wide fallback glyphs out of the following run', async () => {
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
          [probe.text, naturalWidth],
          ['next', 16],
        ])
      );

      await drawPrimitive(
        ctx,
        textRun(probe.text, 0, slotWidth, probe.family, probe.rtl)
      );
      await drawPrimitive(ctx, textRun('next', slotWidth, 16, probe.family));

      expect(paints).toHaveLength(2);
      expect(paints[0].naturalRight).toBeGreaterThan(paints[1].left);
      expect(paints[0].right).toBeLessThanOrEqual(paints[1].left);
    }
  });
});
