import { test, expect, type Route } from 'playwright/test';

type Centres = Record<'parse' | 'renderer', [number, number][]>;

test('the loading spinner holds its position until the first page paints', async ({ page }) => {
  const hold = async (route: Route) => {
    await new Promise((done) => setTimeout(done, 1_000));
    await route.continue();
  };
  await page.context().route(/docx_edit_bg\.wasm/, hold);
  await page.context().route(/residentEngineWorker/, hold);
  await page.addInitScript(() => {
    const centres: Centres = { parse: [], renderer: [] };
    (window as unknown as { __centres: Centres }).__centres = centres;
    const sample = () => {
      const ring = document.querySelector('[style*="docx-spin"]');
      if (ring) {
        const box = ring.getBoundingClientRect();
        const phase = ring.closest('[data-testid="canvas-renderer-loading"]')
          ? 'renderer'
          : 'parse';
        centres[phase].push([box.left + box.width / 2, box.top + box.height / 2]);
      }
      if (!document.querySelector('canvas[data-page-index="0"]')) requestAnimationFrame(sample);
    };
    requestAnimationFrame(sample);
  });
  await page.goto('/');
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  const centres = await page.evaluate(
    () => (window as unknown as { __centres: Centres }).__centres
  );
  expect(centres.parse.length).toBeGreaterThan(0);
  expect(centres.renderer.length).toBeGreaterThan(0);
  const [x, y] = centres.parse[0];
  const drift = [...centres.parse, ...centres.renderer].map(([cx, cy]) =>
    Math.max(Math.abs(cx - x), Math.abs(cy - y))
  );
  expect(Math.max(...drift)).toBeLessThan(0.5);
  await expect(page.getByTestId('canvas-renderer-loading')).toHaveCount(0);
});
