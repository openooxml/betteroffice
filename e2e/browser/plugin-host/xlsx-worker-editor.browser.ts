import { test, expect, type Page } from 'playwright/test';
import type { WorkerEditorProbe } from './xlsx-worker-editor-harness';

declare global {
  interface Window { __xlsxWorkerEditor: WorkerEditorProbe }
}

test.setTimeout(120_000);

async function open(page: Page) {
  await page.goto('/xlsx-worker-editor.html');
  await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
  await page.evaluate(() => window.__xlsxWorkerEditor.ready);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
}

async function edit(page: Page, value: string) {
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await input.fill('');
  await input.pressSequentially(value);
  await expect(input).toHaveValue(value);
  await expect(input).toBeFocused();
  expect(await input.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(value.length);
  await input.press('Enter');
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe(value);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain(value);
  await expect(page.getByTestId('xlsx-commit-preview')).toHaveCount(0);
}

test('xlsx worker editor echoes the caret, paints a commit preview and saves editable bytes for reopening', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await open(page);
  await edit(page, 'worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.previews)).toContain('worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.previewFrames)).toContain('worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.saveAndReopen())).toBeGreaterThan(0);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('worker edit');
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('worker edit');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
  expect(errors).toEqual([]);
});

test('xlsx worker editor undo restores the original cell and canvas value', async ({ page }) => {
  await open(page);
  await edit(page, 'undo value');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.undo())).toBe(true);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('initial');
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).not.toContain('undo value');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});
