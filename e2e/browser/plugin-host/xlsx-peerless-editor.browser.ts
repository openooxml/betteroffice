import { test, expect, type Page } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import type {} from './xlsx-worker-editor-probe';

const pageErrors = new WeakMap<Page, string[]>();

async function open(page: Page, query = 'asyncApi=1') {
  await page.goto(`/xlsx-worker-editor.html?${query}`);
  await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.ready);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('initial');
}

async function cells(page: Page, range = 'A1', sheet = 0) {
  return pinnedEvaluate(page, ({ range, sheet }) => window.__xlsxWorkerEditor.peerless!.cells(sheet, range), { range, sheet });
}

async function selection(page: Page, row: number, col: number, sheetIndex = 0) {
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.selection())).toMatchObject({
    sheetIndex, cells: { anchor: { row, col }, focus: { row, col } },
  });
}

async function edit(page: Page, value: string, commit = 'Enter') {
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await input.press('ControlOrMeta+A');
  await input.press('Delete');
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
  expect(await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.cell())).toBe('initial');
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

for (const route of ['select', 'execCommand', 'none'] as const) {
  test(`peer-less ${
    route === 'none' ? 'typing without selection appends to' : `native select-all via ${route} replaces`
  } a delayed prefill`, async ({ page }) => {
    await open(page);
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('initial');
    await page.evaluate(() => window.__xlsxWorkerEditor.holdCellInputs());
    await page.getByTestId('xlsx-scroll').focus();
    await page.getByTestId('xlsx-scroll').press('F2');
    const input = page.getByTestId('xlsx-cell-editor');
    await expect(input).toBeFocused();
    await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.heldCellInputs())).toBeGreaterThan(0);
    await input.evaluate((element, route) => {
      if (route === 'select') (element as HTMLInputElement).select();
      else if (route === 'execCommand') document.execCommand('selectAll');
    }, route);
    await input.pressSequentially('replacement');
    await page.evaluate(async () => {
      await new Promise((resolve) => setTimeout(resolve, 1000));
      window.__xlsxWorkerEditor.releaseCellInputs();
    });
    await expect(input).not.toHaveValue('');
    await expect(input).not.toHaveValue('initial');
    const expected = route === 'none' ? 'initialreplacement' : 'replacement';
    expect(await input.inputValue()).toBe(expected);
    await input.press('Enter');
    await expect.poll(() => cells(page)).toEqual([[expected]]);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
  });
}

for (const route of ['keyboard', 'native', 'paste', 'unicode'] as const) {
  test(`peer-less uncached pending prefill keeps native input for ${route}`, async ({ page, context }) => {
    await open(page);
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('initial');
    await page.evaluate(() => window.__xlsxWorkerEditor.holdCellInputs());
    const scroll = page.getByTestId('xlsx-scroll');
    await scroll.focus();
    await scroll.press('ArrowRight');
    await selection(page, 0, 1);
    await scroll.press('F2');
    const input = page.getByTestId('xlsx-cell-editor');
    await expect(input).toBeFocused();
    await expect(input).toHaveValue('');
    await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.heldCellInputs())).toBeGreaterThan(0);
    if (route === 'keyboard') await input.press('ControlOrMeta+A');
    if (route === 'native') await input.evaluate((element) => (element as HTMLInputElement).select());
    if (route === 'paste') {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      await pinnedEvaluate(page, () => navigator.clipboard.writeText('replacement'));
      await input.press('ControlOrMeta+v');
    } else if (route === 'unicode') await page.keyboard.insertText('😀');
    else await input.pressSequentially('replacement');
    await page.evaluate(async () => {
      await new Promise((resolve) => setTimeout(resolve, 1000));
      window.__xlsxWorkerEditor.releaseCellInputs();
    });
    const expected = route === 'keyboard' ? 'replacement' : route === 'unicode' ? '1😀' : '1replacement';
    await expect(input).toHaveValue(expected);
    await input.press('Enter');
    await expect.poll(() => cells(page, 'B1')).toEqual([[expected]]);
  });
}

for (const release of ['F2', 'select all', 'Delete', 'typing'] as const) {
  test(`peer-less F2 keeps the edit keys when its held seed arrives after ${release}`, async ({ page }) => {
    await open(page);
    const text = 'held edit';
    const releaseAfter = async (step: typeof release) => {
      if (step !== release) return;
      expect(await page.evaluate(() => window.__xlsxWorkerEditor.heldCellInputs())).toBeGreaterThan(0);
      await page.evaluate(() => window.__xlsxWorkerEditor.releaseCellInputs());
    };
    await page.evaluate(() => window.__xlsxWorkerEditor.holdCellInputs());
    const scroll = page.getByTestId('xlsx-scroll');
    await scroll.focus();
    await scroll.press('F2');
    const input = page.getByTestId('xlsx-cell-editor');
    await expect(input).toBeFocused();
    await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.heldCellInputs())).toBeGreaterThan(0);
    await releaseAfter('F2');
    await input.press('ControlOrMeta+A');
    await releaseAfter('select all');
    await input.press('Delete');
    await releaseAfter('Delete');
    await input.pressSequentially(text);
    await releaseAfter('typing');
    await expect(input).toHaveValue(text);
    expect(await input.evaluate((element) => (element as HTMLInputElement).selectionStart)).toBe(text.length);
    await input.press('Enter');
    await expect(input).toHaveCount(0);
    await expect.poll(() => cells(page)).toEqual([[text]]);
    expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
  });
}

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
  await selection(page, 1, 0);
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
  await pinnedEvaluate(page, (value) => navigator.clipboard.writeText(value), text);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.focus();
  await scroll.press('ControlOrMeta+v');
  await expect.poll(() => cells(page, 'A1:B2')).toEqual([['paste one', 'paste two'], ['paste three', 'paste four']]);
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.selection())).toMatchObject({
    cells: { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } },
  });
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('copy sentinel'));
  await scroll.press('ControlOrMeta+c');
  await expect.poll(() => pinnedEvaluate(page, () => navigator.clipboard.readText())).toBe(text);
  expect(await cells(page, 'A1:B2')).toEqual([['paste one', 'paste two'], ['paste three', 'paste four']]);
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('cut sentinel'));
  await scroll.press('ControlOrMeta+x');
  await expect.poll(() => pinnedEvaluate(page, () => navigator.clipboard.readText())).toBe(text);
  await expect.poll(() => cells(page, 'A1:B2')).toEqual([['', ''], ['', '']]);
  await scroll.press('ControlOrMeta+z');
  await expect.poll(() => cells(page, 'A1:B2')).toEqual([['paste one', 'paste two'], ['paste three', 'paste four']]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
});

test('peer-less plugin reads, edits and navigation survive save and reopen', async ({ page }) => {
  await open(page);
  await edit(page, 'saved edit');
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.pluginRoutes());
  await expect.poll(() => cells(page, 'A4')).toEqual([['plugin']]);
  await selection(page, 0, 0);
  const generation = await page.evaluate(() => window.__xlsxWorkerEditor.generation());
  expect(await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.saveAndReopen())).toBeGreaterThan(0);
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
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.pluginRoutes());
  await expect.poll(() => cells(page, 'A4')).toEqual([['plugin']]);
  const workers = await page.evaluate(() => window.__xlsxWorkerEditor.peerless!.workerCount());
  await page.evaluate(() => window.__xlsxWorkerEditor.fail());
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.peerless!.failure())).toContain('Browser worker failure');
  await expect(page.getByTestId('xlsx-error')).toBeVisible();
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toHaveLength(1);
  expect(await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.recoverAndReopen())).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.peerless!.workerCount())).toBeGreaterThan(workers);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
  expect(await cells(page)).toEqual([['recovered edit']]);
  expect(await cells(page, 'A4')).toEqual([['plugin']]);
  await expect.poll(() => page.evaluate(() => window.__xlsxWorkerEditor.paintedTexts())).toContain('recovered edit');
  await edit(page, 'after recovery');
  await expect.poll(() => cells(page)).toEqual([['after recovery']]);
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toHaveLength(1);
});

test('onReady host methods commit drafts and save edited workbook bytes', async ({ page }) => {
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await open(page);
  const initial = await pinnedEvaluate(page, async () => {
    const api = window.__xlsxWorkerEditor.peerless!.editor;
    return { version: await api.version(), cells: await api.readCells({ ranges: [{
      sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' },
    }] }) };
  });
  expect(initial.version).toBeTruthy();
  expect(initial.cells).toMatchObject({ ok: true });
  await page.getByTestId('xlsx-scroll').focus();
  await page.getByTestId('xlsx-scroll').press('F2');
  await page.getByTestId('xlsx-cell-editor').fill('pending host draft');
  const draft = await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.editor.readCells({
    ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }],
  }));
  expect(JSON.stringify(draft)).toContain('pending host draft');
  const result = await pinnedEvaluate(page, async () => {
    const api = window.__xlsxWorkerEditor.peerless!.editor;
    const applied = await api.applyEdits({ expectVersion: await api.version(), steps: [{ op: 'setCellInputs',
      target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, inputs: [['host batch text']] }] });
    const found = await api.findText({ text: 'host batch text' });
    const bytes = await api.save();
    return { applied, found, bytes: Array.from(bytes), saved: window.__xlsxWorkerEditor.saves.at(-1)?.length };
  });
  expect(result.applied).toMatchObject({ ok: true });
  expect(JSON.stringify(result.found)).toContain('host batch text');
  expect(result.saved).toBe(result.bytes.length);
  expect(downloads).toEqual([]);
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(new Uint8Array(result.bytes));
  expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('host batch text');
});
