import { expect, test } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import type {} from './xlsx-worker-editor-probe';

const FILL_CELLS: [number, number][] = [[1, 1], [1, 2], [1, 3]];
const BORDER_CELLS: [number, number][] = [3, 4, 5, 6, 7, 8].map((row) => [row, 1]);

test('worker-only editor paints device-pixel fills and all extended border styles', async ({ browser, baseURL }) => {
  for (const dpr of [1, 2]) {
    const context = await browser.newContext({ baseURL, deviceScaleFactor: dpr });
    const page = await context.newPage();
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    try {
      await page.goto('/xlsx-worker-editor.html?asyncApi=1&variant=patterns');
      await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
      await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.ready);
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
      for (const zoom of [1, 2]) {
        await pinnedEvaluate(page, (scale) => window.__xlsxWorkerEditor.zoom(scale), zoom);
        await expect(page.locator('[data-paint-source="worker"]')).toHaveAttribute('data-worker-zoom', String(zoom));
        const cells = [...FILL_CELLS, ...BORDER_CELLS];
        await expect.poll(() => pinnedEvaluate(page, async (cells) =>
          (await window.__xlsxWorkerEditor.peerless!.cellRects(cells))?.zoom ?? null, cells)).toBe(zoom);
        const samples = await pinnedEvaluate(page, async ({ cells, fillCount, zoom }) => {
          const { rects } = (await window.__xlsxWorkerEditor.peerless!.cellRects(cells))!;
          const canvas = document.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]')!;
          const ctx = canvas.getContext('2d')!;
          const ratio = canvas.width / canvas.clientWidth;
          const scale = ratio * zoom;
          const pixel = (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data);
          const fills = rects.slice(0, fillCount).map((rect) => {
            const x = Math.ceil(rect.x * ratio) + 4;
            const y = Math.ceil(rect.y * ratio) + 4;
            return { x, y, colors: Array.from({ length: 8 }, (_, dy) => Array.from({ length: 8 }, (_, dx) => pixel(x + dx, y + dy))) };
          });
          const borders = rects.slice(fillCount).map((rect) => {
            const x = Math.round(rect.x * ratio);
            const y = Math.round(rect.y * ratio);
            return Array.from({ length: 64 }, (_, dx) => {
              const band = Array.from({ length: Math.max(2, 2 * scale) + 1 }, (_, dy) => pixel(x + dx, y - Math.ceil(scale) + dy));
              return band.some((color) => color[0] === 18 && color[1] === 52 && color[2] === 86);
            });
          });
          return { fills, borders, scale };
        }, { cells, fillCount: FILL_CELLS.length, zoom });
        for (const [i, fill] of samples.fills.entries()) {
          for (let y = 0; y < 8; y += 1) for (let x = 0; x < 8; x += 1) {
            const px = fill.x + x;
            const py = fill.y + y;
            const fg = i < 2 ? (px + py) % 4 === 0 : py % 2 === 0 && px % 4 === (py % 4 === 0 ? 0 : 2);
            expect(fill.colors[y]![x], `fill ${i}, pixel ${px},${py}, dpr ${dpr}, zoom ${zoom}`)
              .toEqual(fg ? [18, 52, 86, 255] : [254, 220, 186, 255]);
          }
        }
        const sequences = [[8, 4], [4, 2, 1, 2], [8, 4, 2, 4], [4, 2, 1, 2, 1, 2], [8, 4, 2, 4, 2, 4], [4, 2, 1, 2]];
        for (const [i, pixels] of samples.borders.entries()) {
          const runs = sequences[i]!.map((run) => run * samples.scale);
          const period = runs.reduce((a, b) => a + b, 0);
          for (let x = 0; x < 64; x += 1) {
            let offset = (x + 0.5) % period;
            let run = 0;
            while (offset >= runs[run]!) offset -= runs[run++]!;
            expect(pixels[x], `border ${i}, pixel ${x}, dpr ${dpr}, zoom ${zoom}`).toBe(run % 2 === 0);
          }
        }
      }
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.mainThreadEngines())).toBe(0);
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries)).toEqual([]);
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
      expect(errors).toEqual([]);
    } finally {
      await context.close();
    }
  }
});
