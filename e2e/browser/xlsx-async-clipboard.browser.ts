import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { clipboardWorkbook, selectCell, workbookXml } from './xlsx-async-clipboard-helpers';

test.beforeEach(async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await open(page, await clipboardWorkbook());
});

async function open(page: Page, file: string | Awaited<ReturnType<typeof clipboardWorkbook>>, a1 = '11', activeSheet = 'Sheet') {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(file);
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByRole('tab', { name: activeSheet, exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(a1);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
}

async function save(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.xlsx');
  await (await downloading).saveAs(file);
  return { file, sheets: await workbookXml(page, await readFile(file)) };
}

async function roundTrip(page: Page, info: TestInfo) {
  const activeSheet = await page.getByRole('tab', { selected: true }).innerText();
  const saved = await save(page, info, 'pasted');
  const a1 = saved.sheets[activeSheet === 'Other' ? other : sheet].cells.A1;
  await open(page, saved.file, a1?.formula ? '=' + a1.formula : a1?.value ?? '', activeSheet);
  expect((await save(page, info, 'reopened')).sheets).toEqual(saved.sheets);
  return saved.sheets;
}

async function input(page: Page, row: number, col: number, value: string) {
  await selectCell(page, row, col);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(value);
}

async function range(page: Page) {
  await selectCell(page, 0, 0);
  await page.getByTestId('xlsx-scroll').press('Shift+ArrowRight');
  await page.getByTestId('xlsx-scroll').press('Shift+ArrowDown');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByRole('gridcell', { selected: true })).toHaveCount(4);
}

async function copy(page: Page, expected?: string, cut = false) {
  await page.evaluate(() => navigator.clipboard.writeText('clipboard-sentinel'));
  await page.getByTestId('xlsx-scroll').press(cut ? 'ControlOrMeta+x' : 'ControlOrMeta+c');
  const clipboard = expect.poll(() => page.evaluate(() => navigator.clipboard.readText()));
  if (expected === undefined) await clipboard.not.toBe('clipboard-sentinel');
  else await clipboard.toBe(expected);
}

async function pasteText(page: Page, text: string) {
  await page.evaluate((text) => navigator.clipboard.writeText(text), text);
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+v');
}

async function history(page: Page, info: TestInfo, row: number, col: number, value: string, previous = '') {
  await input(page, row, col, value);
  await expect(page.getByRole('gridcell').filter({ hasText: value }).first()).toHaveCount(1);
  const after = (await save(page, info, 'after-paste')).sheets;
  await page.getByTestId('xlsx-undo').click();
  await input(page, row, col, previous);
  const undone = (await save(page, info, 'undone')).sheets;
  expect(undone).not.toEqual(after);
  await page.getByTestId('xlsx-redo').click();
  await input(page, row, col, value);
  expect((await save(page, info, 'redone')).sheets).toEqual(after);
  return roundTrip(page, info);
}

const sheet = 'xl/worksheets/sheet1.xml';
const other = 'xl/worksheets/sheet2.xml';

test('copy a two-dimensional range into another sheet and preserve the source', async ({ page }, info) => {
  await range(page);
  await copy(page, '11\t12\r\n21\t22');
  await page.getByRole('tab', { name: 'Other', exact: true }).click();
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('other');
  await selectCell(page, 2, 1);
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+v');
  const saved = await history(page, info, 2, 1, '11');
  expect(saved[other].cells.B3.value).toBe('11');
  expect(saved[other].cells.C3.value).toBe('12');
  expect(saved[other].cells.B4.value).toBe('21');
  expect(saved[other].cells.C4.value).toBe('22');
  expect(saved[sheet].cells.A1.value).toBe('11');
  expect(saved[sheet].cells.B2.value).toBe('22');
});

test('paste beyond the current sheet extent expands the grid without dropping cells', async ({ page }, info) => {
  const before = (await save(page, info, 'before-extent')).sheets;
  expect(before[sheet].dimension).toBe('A1:H8');
  expect(before[sheet].cells.H8.value).toBe('extent');
  await selectCell(page, 0, 9);
  await pasteText(page, 'edge-one\tedge-two\r\nedge-three\tedge-four');
  await expect(page.getByRole('gridcell').filter({ hasText: 'edge-one' })).toHaveCount(1);
  const saved = await history(page, info, 0, 9, 'edge-one');
  expect(saved[sheet].cells.J1.value).toBe('edge-one');
  expect(saved[sheet].cells.K1.value).toBe('edge-two');
  expect(saved[sheet].cells.J2.value).toBe('edge-three');
  expect(saved[sheet].cells.K2.value).toBe('edge-four');
  for (const [address, cell] of Object.entries(before[sheet].cells)) expect(saved[sheet].cells[address]).toEqual(cell);
  expect(saved[sheet].merges).toEqual(before[sheet].merges);
  expect(saved[other]).toEqual(before[other]);
  expect(saved[sheet].dimension).toBe('A1:K8');
  await input(page, 1, 10, 'edge-four');
});

test('paste external TSV with quoted tabs, newlines, empty fields and a trailing newline', async ({ page }, info) => {
  await selectCell(page, 3, 0);
  await pasteText(page, 'outside\t"a\tb"\t\r\n"line 1\nline 2"\t"say ""hi"""\tlast\r\n');
  const saved = await history(page, info, 3, 0, 'outside');
  expect(saved[sheet].cells.B4.value).toBe('a\tb');
  expect(saved[sheet].cells.C4?.value ?? '').toBe('');
  expect(saved[sheet].cells.A5.value).toBe('line 1\nline 2');
  expect(saved[sheet].cells.B5.value).toBe('say "hi"');
  expect(saved[sheet].cells.C5.value).toBe('last');
  expect(saved[sheet].cells.A6?.value ?? '').toBe('');
});

test('paste a single value over a merged range preserves the merge and replaces its anchor', async ({ page }, info) => {
  await selectCell(page, 0, 5);
  await pasteText(page, 'merged replacement');
  const saved = await history(page, info, 0, 5, 'merged replacement', 'merged');
  expect(saved[sheet].merges).toContain('F1:G1');
  expect(saved[sheet].cells.F1.value).toBe('merged replacement');
  expect(saved[sheet].cells.G1?.value ?? '').toBe('');
});

test('copy exposes TSV without changing the workbook', async ({ page }, info) => {
  const before = (await save(page, info, 'before-copy')).sheets;
  await range(page);
  await copy(page, '11\t12\r\n21\t22');
  expect(await roundTrip(page, info)).toEqual(before);
});

test('one undo removes the entire pasted block and preserves the preceding edit', async ({ page }, info) => {
  await selectCell(page, 2, 0);
  await page.getByTestId('xlsx-scroll').press('F2');
  const editor = page.getByTestId('xlsx-cell-editor');
  await editor.fill('preceding edit');
  await editor.press('Enter');
  await input(page, 2, 0, 'preceding edit');
  const before = (await save(page, info, 'before-block')).sheets;
  await selectCell(page, 3, 0);
  await pasteText(page, 'block-a\tblock-b\r\nblock-c\tblock-d');
  await input(page, 4, 1, 'block-d');
  const after = (await save(page, info, 'after-block')).sheets;
  await page.getByTestId('xlsx-undo').click();
  await input(page, 2, 0, 'preceding edit');
  for (const [row, col] of [[3, 0], [3, 1], [4, 0], [4, 1]]) await input(page, row, col, '');
  expect((await save(page, info, 'one-undo')).sheets).toEqual(before);
  await page.getByTestId('xlsx-redo').click();
  await input(page, 4, 1, 'block-d');
  expect(await roundTrip(page, info)).toEqual(after);
});

test('copy formulas shifts relative references and keeps absolute and mixed references', async ({ page }, info) => {
  await selectCell(page, 0, 3);
  await copy(page, '=A1+$B$1+A$2+$B2');
  await selectCell(page, 2, 3);
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+v');
  await input(page, 2, 3, '=A3+$B$1+A$2+$B4');
  await expect(page.getByRole('gridcell').filter({ hasText: /^D3, / })).toHaveText(/^D3, 33(?:, selected)?$/);
  const after = (await save(page, info, 'formula-paste')).sheets;
  expect(after[sheet].cells.D3.formula).toBe('A3+$B$1+A$2+$B4');
  await page.getByTestId('xlsx-undo').click();
  await input(page, 2, 3, '');
  await page.getByTestId('xlsx-redo').click();
  await input(page, 2, 3, '=A3+$B$1+A$2+$B4');
  expect(await roundTrip(page, info)).toEqual(after);
});

test('copying a formula shifts only its relative references at the destination', async ({ page }, info) => {
  await input(page, 0, 3, '=A1+$B$1+A$2+$B2');
  await copy(page);
  await selectCell(page, 2, 4);
  await page.getByTestId('xlsx-scroll').press('ControlOrMeta+v');
  await input(page, 2, 4, '=B3+$B$1+B$2+$B4');
  const saved = await roundTrip(page, info);
  expect(saved[sheet].cells.E3.formula).toBe('B3+$B$1+B$2+$B4');
  expect(saved[sheet].cells.D1.formula).toBe('A1+$B$1+A$2+$B2');
});
