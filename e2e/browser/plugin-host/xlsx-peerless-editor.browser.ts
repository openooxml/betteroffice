import { test, expect, type Page } from 'playwright/test';
import type {} from './xlsx-worker-editor-probe';

const pageErrors = new WeakMap<Page, string[]>();

async function open(page: Page, query = 'asyncApi=1') {
  await page.goto(`/xlsx-worker-editor.html?${query}`);
  await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
  await page.evaluate(() => window.__xlsxWorkerEditor.ready);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
}

async function cells(page: Page, range = 'A1', sheet = 0) {
  return page.evaluate(({ range, sheet }) => window.__xlsxWorkerEditor.peerless!.cells(sheet, range), { range, sheet });
}

async function selection(page: Page, row: number, col: number, sheetIndex = 0) {
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.peerless!.selection())).toMatchObject({
    sheetIndex, cells: { anchor: { row, col }, focus: { row, col } },
  });
}

async function edit(page: Page, value: string, commit = 'Enter') {
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await input.fill('');
  await input.pressSequentially(value);
  await expect(input).toHaveValue(value);
  await expect(input).toBeFocused();
  expect(await input.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(value.length);
  await input.press(commit);
  await expect(input).toHaveCount(0);
}

test.beforeEach(async ({ page }) => {
  const errors: string[] = [];
  pageErrors.set(page, errors);
  page.on('pageerror', (error) => errors.push(error.message));
});

test.afterEach(async ({ page }) => {
  expect(pageErrors.get(page)).toEqual([]);
  if (!new URL(page.url()).searchParams.has('asyncApi')) return;
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.mainThreadEngines())).toBe(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerEntries)).toEqual([]);
  await expect(page.getByTestId('xlsx-input-refusal')).toHaveCount(0);
});

test('asyncApi mounts a peer-less editor without a main-thread engine', async ({ page }) => {
  await open(page);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.mainThreadEngines())).toBe(0);
  expect(await cells(page)).toEqual([['initial']]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.hydrated())).toBe(true);
  await expect(page.locator('[data-paint-source="worker"]')).toBeVisible();
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('hybrid control still creates a main-thread engine', async ({ page }) => {
  await open(page, '');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.mainThreadEngines())).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.cell())).toBe('initial');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less typing commits the cell and formula bar before navigation', async ({ page }) => {
  await open(page);
  await edit(page, 'worker edit', 'Tab');
  await expect.poll(() => cells(page)).toEqual([['worker edit']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('worker edit');
  await selection(page, 0, 1);
  const formula = page.getByTestId('xlsx-formula-input');
  await formula.fill('=24');
  await formula.press('Enter');
  await expect.poll(() => cells(page, 'B1')).toEqual([['=24']]);
  await selection(page, 1, 1);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('24');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less keyboard navigation and sheet tabs update the selection and canvas', async ({ page }) => {
  await open(page);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('ArrowRight');
  await selection(page, 0, 1);
  await scroll.press('ArrowDown');
  await selection(page, 1, 1);
  await scroll.press('ArrowLeft');
  await selection(page, 1, 0);
  await page.getByRole('tab', { name: 'Other' }).click();
  await selection(page, 0, 0, 1);
  await expect(page.getByRole('tab', { name: 'Other' })).toHaveAttribute('aria-selected', 'true');
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('other');
  await page.getByRole('tab', { name: 'Sheet', exact: true }).click();
  await selection(page, 0, 0);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less undo and redo restore cell inputs and painted values', async ({ page }) => {
  await open(page);
  await edit(page, 'history value');
  await expect.poll(() => cells(page)).toEqual([['history value']]);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('ControlOrMeta+z');
  await expect.poll(() => cells(page)).toEqual([['initial']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).not.toContain('history value');
  await scroll.press('ControlOrMeta+Shift+z');
  await expect.poll(() => cells(page)).toEqual([['history value']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('history value');
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less clipboard pastes, copies and cuts a rectangular range', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await open(page);
  const text = 'paste one\tpaste two\r\npaste three\tpaste four';
  await page.evaluate((value) => navigator.clipboard.writeText(value), text);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('ControlOrMeta+v');
  await expect.poll(() => cells(page, 'A1:B2')).toEqual([['paste one', 'paste two'], ['paste three', 'paste four']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.peerless!.selection())).toMatchObject({
    cells: { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } },
  });
  await page.evaluate(() => navigator.clipboard.writeText('copy sentinel'));
  await scroll.press('ControlOrMeta+c');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(text);
  expect(await cells(page, 'A1:B2')).toEqual([['paste one', 'paste two'], ['paste three', 'paste four']]);
  await page.evaluate(() => navigator.clipboard.writeText('cut sentinel'));
  await scroll.press('ControlOrMeta+x');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(text);
  await expect.poll(() => cells(page, 'A1:B2')).toEqual([['', ''], ['', '']]);
  await scroll.press('ControlOrMeta+z');
  await expect.poll(() => cells(page, 'A1:B2')).toEqual([['paste one', 'paste two'], ['paste three', 'paste four']]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less plugin reads, edits and navigation survive save and reopen', async ({ page }) => {
  await open(page);
  await edit(page, 'saved edit');
  await page.evaluate(() => window.__xlsxWorkerEditor.pluginRoutes());
  await expect.poll(() => cells(page, 'A4')).toEqual([['plugin']]);
  await selection(page, 0, 0);
  const generation = await page.evaluate(() => window.__xlsxWorkerEditor.generation());
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerless!.saveAndReopen())).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.generation())).toBeGreaterThan(generation);
  expect(await cells(page)).toEqual([['saved edit']]);
  expect(await cells(page, 'A4')).toEqual([['plugin']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('saved edit');
  await edit(page, 'reopened edit');
  await expect.poll(() => cells(page)).toEqual([['reopened edit']]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less recovery replays edits in a new worker after termination', async ({ page }) => {
  await open(page);
  await edit(page, 'recovered edit');
  await expect.poll(() => cells(page)).toEqual([['recovered edit']]);
  await page.evaluate(() => window.__xlsxWorkerEditor.pluginRoutes());
  await expect.poll(() => cells(page, 'A4')).toEqual([['plugin']]);
  const workers = await page.evaluate(() => window.__xlsxWorkerEditor.peerless!.workerCount());
  await page.evaluate(() => window.__xlsxWorkerEditor.fail());
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.peerless!.failure())).toContain('Browser worker failure');
  await expect(page.getByTestId('xlsx-error')).toBeVisible();
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toHaveLength(1);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerless!.recoverAndReopen())).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerless!.workerCount())).toBeGreaterThan(workers);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
  expect(await cells(page)).toEqual([['recovered edit']]);
  expect(await cells(page, 'A4')).toEqual([['plugin']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('recovered edit');
  await edit(page, 'after recovery');
  await expect.poll(() => cells(page)).toEqual([['after recovery']]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toHaveLength(1);
});
