import { test, expect, type Page } from 'playwright/test';
import type { PixelResult } from './pptx-worker-editor-probe';

test.use({ headless: true, launchOptions: { executablePath: process.env.CHROME || undefined } });

const failures = new WeakMap<Page, { console: string[]; page: string[] }>();
test.beforeEach(async ({ page }) => {
  const errors = { console: [] as string[], page: [] as string[] };
  failures.set(page, errors);
  page.on('console', (message) => {
    if (message.type() === 'error') errors.console.push(message.text());
  });
  page.on('pageerror', (error) => errors.page.push(error.message));
});
test.afterEach(async ({ page }) => {
  const errors = failures.get(page)!;
  expect.soft(errors.console, 'Console errors').toEqual([]);
  expect.soft(errors.page, 'Page errors').toEqual([]);
});

function assertParity(results: PixelResult[], dpr: number) {
  expect(results).toHaveLength(2);
  expect(results.map((result) => result.thumbnail).sort()).toEqual([false, true]);
  for (const result of results) {
    const label = `${result.thumbnail ? 'thumbnail' : 'slide'} DPR ${dpr}, sequence ${result.sequence}`;
    expect(result.dpr, label).toBe(dpr);
    expect(result.differingPixels, label).toBe(0);
    expect(result.maxChannelDelta, label).toBe(0);
  }
}

async function recovery(page: Page, index: number, text: string) {
  const saved = await page.evaluate((index) => window.__pptxWorkerEditor.recover(index), index);
  expect(saved.recovery).toBe(true);
  expect(saved.bytes).toBeGreaterThan(0);
  expect(saved.text).toBe(text);
}

for (const dpr of [1, 2]) {
  test.describe(`worker editor DPR ${dpr}`, () => {
    test.use({ deviceScaleFactor: dpr });
    test('worker_editing_hydration_paint_save_and_recovery', async ({ page }) => {
      await page.goto('/pptx-worker-editor.html');
      await page.waitForFunction(() => '__pptxWorkerEditor' in window);
      await page.evaluate(() => window.__pptxWorkerEditor.mounted);
      await expect(page.getByTestId('pptx-worker-status')).toHaveText('Preparing editor. Editing will be available shortly.');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.state()))
        .toEqual({ hydrated: false, sequence: 0, acknowledged: 0, held: 1 });
      const stage = page.getByRole('application');
      await stage.focus();
      await page.keyboard.press('x');
      await expect(page.getByText('Presentation editor is not ready', { exact: true })).toBeVisible();
      expect(await page.evaluate(() => window.__pptxWorkerEditor.refuse())).toBe('PptxPeerNotReadyError');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.unchanged())).toBe(true);
      await expect(page.getByTestId('pptx-slide-canvas').locator('..').locator('canvas[aria-hidden="true"]')).toHaveCount(0);
      expect(await page.evaluate(() => window.__pptxWorkerEditor.state().sequence)).toBe(0);

      await page.evaluate(() => window.__pptxWorkerEditor.releaseHydration());
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('initial');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.state().hydrated)).toBe(true);
      assertParity(await page.evaluate(() => window.__pptxWorkerEditor.parity()), dpr);
      expect(await page.evaluate(() => window.__pptxWorkerEditor.select(0, 7))).toBe(true);
      const overlay = page.getByTestId('pptx-slide-canvas').locator('..').locator('canvas[aria-hidden="true"]');
      await expect(overlay).toHaveCount(1);
      await expect.poll(() => overlay.evaluate((canvas) => {
        const target = canvas as HTMLCanvasElement;
        return target.getContext('2d')!.getImageData(0, 0, target.width, target.height).data
          .some((value, index) => index % 4 === 3 && value > 0);
      })).toBe(true);
      await page.evaluate(async () => { await window.__pptxWorkerEditor.baseline(); window.__pptxWorkerEditor.hold(); });
      await page.keyboard.type('worker edit');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('worker edit');
      await expect.poll(() => page.evaluate(() => window.__pptxWorkerEditor.state().held)).toBeGreaterThan(0);
      const lag = await page.evaluate(() => window.__pptxWorkerEditor.state());
      expect(lag.sequence).toBeGreaterThan(lag.acknowledged);
      expect(await page.evaluate(() => window.__pptxWorkerEditor.unchanged())).toBe(true);
      await expect(overlay).toHaveCount(0);
      expect(await page.evaluate(() => window.__pptxWorkerEditor.select(0, 4))).toBe(true);
      await expect(overlay).toHaveCount(0);
      const point = await page.evaluate(() => window.__pptxWorkerEditor.point(0));
      expect(point.position).toBe(0);
      await page.mouse.click(point.x, point.y);
      await page.keyboard.press('!');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('!worker edit');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.unchanged())).toBe(true);
      await expect(overlay).toHaveCount(0);

      await page.evaluate(() => window.__pptxWorkerEditor.release());
      assertParity(await page.evaluate(() => window.__pptxWorkerEditor.parity()), dpr);
      expect(await page.evaluate(() => window.__pptxWorkerEditor.unchanged())).toBe(false);
      const caret = await page.evaluate(() => window.__pptxWorkerEditor.overlay(1));
      expect(caret.pixels).toBeGreaterThan(0);
      expect(caret.differingPixels).toBe(0);
      const provenance = await page.evaluate(() => window.__pptxWorkerEditor.provenance());
      expect(provenance.total).toBeGreaterThan(0);
      expect(provenance.thumbnails).toBeGreaterThan(0);
      expect(provenance.unknown).toBe(0);
      const saved = await page.evaluate(() => window.__pptxWorkerEditor.save());
      expect(saved.bytes).toBeGreaterThan(0);
      expect(saved.text).toBe('!worker edit');

      const reason = 'This feature is unavailable in this editing mode.';
      for (const [id, title] of [['pptx-export-png', `Export PNG: ${reason}`], ['pptx-present', reason]]) {
        const control = page.getByTestId(id);
        await expect(control).toHaveAttribute('aria-disabled', 'true');
        await expect(control).toHaveAttribute('title', title);
      }
      await page.evaluate(() => window.__pptxWorkerEditor.proposal());
      await page.getByTestId('pptx-proposals-button').click();
      const preview = page.getByTestId('pptx-proposal-preview');
      await expect(preview).toBeDisabled();
      await expect(preview).toHaveAttribute('title', 'This feature is unavailable in this editing mode.');
      await expect(page.getByTestId('pptx-proposals-panel').getByText('This feature is unavailable in this editing mode.', { exact: true })).toBeVisible();
      await preview.click({ force: true });
      await expect(page.getByTestId('pptx-proposal-preview-dialog')).toHaveJSProperty('open', false);
      await page.getByTestId('pptx-proposals-panel').getByRole('button', { name: 'Close', exact: true }).click();

      expect(await page.evaluate(() => window.__pptxWorkerEditor.select(0, 12))).toBe(true);
      await page.evaluate(() => window.__pptxWorkerEditor.hold());
      await page.keyboard.type('accepted before failure');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('accepted before failure');
      await page.evaluate(() => window.__pptxWorkerEditor.fail());
      await expect.poll(() => page.evaluate(() => window.__pptxWorkerEditor.errors.filter((error) => error.typed)),
        { timeout: 75_000, intervals: [250, 1000] }).toEqual([
        { name: 'PptxWorkerEditorFailedError', code: 'editor-failed', cause: 'silence', typed: true },
      ]);
      expect(await page.evaluate(() => window.__pptxWorkerEditor.rejected(0)))
        .toEqual(Array(7).fill('PptxWorkerEditorFailedError'));
      await recovery(page, 0, 'accepted before failure');

      await page.evaluate(() => window.__pptxWorkerEditor.replace());
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('initial');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.rejected(0)))
        .toEqual(Array(7).fill('PptxWorkerEditorDisposedError'));
      await recovery(page, 0, 'accepted before failure');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.select(0, 7))).toBe(true);
      await page.evaluate(() => window.__pptxWorkerEditor.hold());
      await page.keyboard.type('accepted before replacement');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('accepted before replacement');
      await page.evaluate(() => window.__pptxWorkerEditor.replace());
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('initial');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.rejected(1)))
        .toEqual(Array(7).fill('PptxWorkerEditorDisposedError'));
      await recovery(page, 1, 'accepted before replacement');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.select(0, 7))).toBe(true);
      await page.evaluate(() => window.__pptxWorkerEditor.hold());
      await page.keyboard.type('accepted before unmount');
      expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('accepted before unmount');
      await page.evaluate(() => window.__pptxWorkerEditor.unmount());
      for (const [index, text] of ['accepted before failure', 'accepted before replacement', 'accepted before unmount'].entries()) {
        expect(await page.evaluate((index) => window.__pptxWorkerEditor.rejected(index), index))
          .toEqual(Array(7).fill('PptxWorkerEditorDisposedError'));
        await recovery(page, index, text);
      }
      expect(await page.evaluate(() => window.__pptxWorkerEditor.errors)).toEqual([
        { name: 'PptxPeerNotReadyError', code: 'peer-not-ready', cause: '', typed: false },
        { name: 'PptxWorkerEditorFailedError', code: 'editor-failed', cause: 'silence', typed: true },
      ]);
    });
  });
}
