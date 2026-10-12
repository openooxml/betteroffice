import { expect, test } from 'playwright/test';
import type {} from './xlsx-worker-editor-probe';

test('worker-only editor paints device-pixel fills and all extended border styles', async ({ browser, baseURL }) => {
  for (const dpr of [1, 2]) {
    const context = await browser.newContext({ baseURL, deviceScaleFactor: dpr });
    const page = await context.newPage();
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    try {
      await page.goto('/xlsx-worker-editor.html?asyncApi=1&variant=patterns');
      await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
      await page.evaluate(() => window.__xlsxWorkerEditor.ready);
      await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
      for (const zoom of [1, 2]) {
        await page.evaluate((scale) => window.__xlsxWorkerEditor.zoom(scale), zoom);
        await expect(page.locator('[data-paint-source="worker"]')).toHaveAttribute('data-worker-zoom', String(zoom));
        const samples = await page.evaluate((zoom) => {
          const canvas = document.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]')!;
          const ctx = canvas.getContext('2d')!;
          const scale = canvas.width / canvas.clientWidth * zoom;
          const col = Math.round((24 * 7 + 5) * scale);
          const row = Math.round(32 * scale);
          const pixel = (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data);
          const fills = [1, 2, 3].map((column) => {
            const x = col * column + 16;
            const y = row + 8;
            return { x, y, colors: Array.from({ length: 8 }, (_, dy) => Array.from({ length: 8 }, (_, dx) => pixel(x + dx, y + dy))) };
          });
          const borders = Array.from({ length: 6 }, (_, i) => {
            const y = row * (i + 3);
            return Array.from({ length: 64 }, (_, dx) => {
              const band = Array.from({ length: Math.max(2, 2 * scale) + 1 }, (_, dy) => pixel(col + dx, y - Math.ceil(scale) + dy));
              return band.some((color) => color[0] === 18 && color[1] === 52 && color[2] === 86);
            });
          });
          return { fills, borders, scale };
        }, zoom);
        for (const [i, fill] of samples.fills.entries()) {
          for (let y = 0; y < 8; y += 1) for (let x = 0; x < 8; x += 1) {
            const px = fill.x + x;
            const py = fill.y + y;
            const fg = i < 2 ? (px + py) % 4 === 0 : py % 2 === 0 && px % 4 === (py % 4 === 0 ? 0 : 2);
            expect(fill.colors[y]![x]).toEqual(fg ? [18, 52, 86, 255] : [254, 220, 186, 255]);
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
            expect(pixels[x], `border ${i}, pixel ${x}, dpr ${dpr}`).toBe(run % 2 === 0);
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
