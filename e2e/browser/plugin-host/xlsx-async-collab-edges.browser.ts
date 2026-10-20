import { test, expect, type Page } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import { workbookXml } from '../xlsx-async-clipboard-helpers';
import type {} from './xlsx-worker-editor-probe';

const sheet = 'xl/worksheets/sheet1.xml';
const other = 'xl/worksheets/sheet2.xml';
const peer = (page: Page, index: number) => page.locator(`[data-peerless-editor="${index}"]`);

test.beforeEach(async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.setViewportSize({ width: 2560, height: 1000 });
  await page.goto('/xlsx-worker-editor.html?asyncApi=1&collaboration=1');
  await page.waitForFunction(() => '__xlsxWorkerEditor' in window);
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.ready);
  for (const index of [0, 1]) {
    await expect(peer(page, index).locator('[data-paint-source="worker"]')).toBeVisible();
    await expect(peer(page, index).getByTestId('xlsx-formula-input')).toHaveValue('initial');
  }
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.mainThreadEngines())).toBe(0);
});

test.afterEach(async ({ page }) => {
  expect(await page.evaluate(() => window.__xlsxWorkerEditor.errors)).toEqual([]);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-input-refusal')).toHaveCount(0);
});

async function select(page: Page, index: number, row = 0, col = 0) {
  const editor = peer(page, index);
  const scroll = editor.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  for (let offset = 0; offset < row; offset++) await scroll.press('ArrowDown');
  for (let offset = 0; offset < col; offset++) await scroll.press('ArrowRight');
  await expect(editor.getByTestId('xlsx-name-box')).toHaveValue(String.fromCharCode(65 + col) + (row + 1));
}

async function edit(page: Page, index: number, value: string, row = 0, col = 0) {
  await select(page, index, row, col);
  const formula = peer(page, index).getByTestId('xlsx-formula-input');
  await formula.fill(value);
  await formula.press('Enter');
}

async function converged(page: Page, range: string, expected: string[][], sheetIndex = 0) {
  await expect.poll(() => pinnedEvaluate(page, ({ range, sheetIndex }) => Promise.all([
    window.__xlsxWorkerEditor.peerless!.cells(sheetIndex, range, 0),
    window.__xlsxWorkerEditor.peerless!.cells(sheetIndex, range, 1),
  ]), { range, sheetIndex })).toEqual([expected, expected]);
}

async function painted(page: Page, ...values: string[]) {
  await expect.poll(() => page.evaluate(() => [
    window.__xlsxWorkerEditor.paintedTexts(0),
    window.__xlsxWorkerEditor.paintedTexts(1),
  ])).toEqual([expect.arrayContaining(values), expect.arrayContaining(values)]);
}

async function queued(page: Page, sender: number) {
  await expect.poll(() => page.evaluate((sender) => window.__xlsxWorkerEditor.pendingUpdates(sender), sender)).toBeGreaterThan(0);
}

async function resume(page: Page) {
  await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.resumeDelivery());
}

async function savedPeers(page: Page) {
  const saved = [];
  for (const index of [0, 1]) {
    await peer(page, index).getByTestId('xlsx-save').click();
    await expect.poll(() => page.evaluate((index) => window.__xlsxWorkerEditor.savedBytes(index)?.length ?? 0, index)).toBeGreaterThan(0);
    const bytes = await page.evaluate((index) => window.__xlsxWorkerEditor.savedBytes(index)!, index);
    saved.push(await workbookXml(page, Buffer.from(bytes)));
  }
  expect(saved[0]).toEqual(saved[1]);
  return saved;
}

async function history(page: Page, index: number, direction: 'undo' | 'redo') {
  const button = peer(page, index).getByTestId('xlsx-' + direction);
  await expect(button).toBeEnabled();
  await button.click();
}

test('concurrent formatting and a value edit in the same cell preserve both changes', async ({ page }) => {
  await page.evaluate(() => window.__xlsxWorkerEditor.pauseDelivery());
  await peer(page, 0).getByRole('button', { name: 'Bold', exact: true }).click();
  await expect(peer(page, 0).getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await edit(page, 1, 'formatted remote value');
  await queued(page, 0);
  await queued(page, 1);
  await resume(page);
  await converged(page, 'A1', [['formatted remote value']]);
  await painted(page, 'formatted remote value');
  for (const index of [0, 1]) {
    await select(page, index);
    await expect(peer(page, index).getByTestId('xlsx-formula-input')).toHaveValue('formatted remote value');
    await expect(peer(page, index).getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  }
  await history(page, 0, 'undo');
  await converged(page, 'A1', [['formatted remote value']]);
  for (const index of [0, 1]) await expect(peer(page, index).getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'false');
  await history(page, 0, 'redo');
  for (const index of [0, 1]) await expect(peer(page, index).getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  for (const saved of await savedPeers(page)) expect(saved[sheet].cells.A1).toMatchObject({ value: 'formatted remote value', bold: true });
});

test('switching sheets on one peer does not redirect edits or navigation on the other', async ({ page }) => {
  await peer(page, 0).getByRole('tab', { name: 'Other', exact: true }).click();
  await expect(peer(page, 0).getByTestId('xlsx-formula-input')).toHaveValue('other');
  await edit(page, 1, 'first sheet remote');
  await converged(page, 'A1', [['first sheet remote']]);
  await expect(peer(page, 0).getByRole('tab', { name: 'Other', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(peer(page, 1).getByRole('tab', { name: 'Sheet', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(peer(page, 0).getByTestId('xlsx-formula-input')).toHaveValue('other');
  await edit(page, 0, 'second sheet local', 0, 1);
  await converged(page, 'A1:B1', [['other', 'second sheet local']], 1);
  await history(page, 0, 'undo');
  await converged(page, 'B1', [['']], 1);
  await converged(page, 'A1', [['first sheet remote']]);
  await history(page, 0, 'redo');
  await converged(page, 'B1', [['second sheet local']], 1);
  for (const index of [0, 1]) {
    await peer(page, index).getByRole('tab', { name: 'Other', exact: true }).click();
    await select(page, index, 0, 1);
    await expect(peer(page, index).getByTestId('xlsx-formula-input')).toHaveValue('second sheet local');
  }
  await painted(page, 'other', 'second sheet local');
  for (const saved of await savedPeers(page)) {
    expect(saved[sheet].cells.A1.value).toBe('first sheet remote');
    expect(saved[other].cells.B1.value).toBe('second sheet local');
  }
});

test('undo after a remote overwrite of the same cell preserves the remote value', async ({ page }) => {
  await edit(page, 0, 'local before remote');
  await converged(page, 'A1', [['local before remote']]);
  await edit(page, 1, 'remote replacement');
  await converged(page, 'A1', [['remote replacement']]);
  await history(page, 0, 'undo');
  await converged(page, 'A1', [['remote replacement']]);
  for (const index of [0, 1]) {
    await select(page, index);
    await expect(peer(page, index).getByTestId('xlsx-formula-input')).toHaveValue('remote replacement');
  }
  await painted(page, 'remote replacement');
  for (const saved of await savedPeers(page)) expect(saved[sheet].cells.A1.value).toBe('remote replacement');
});

test('a concurrent range paste and typing inside that range converge without losing other pasted cells', async ({ page }) => {
  await page.evaluate(() => window.__xlsxWorkerEditor.pauseDelivery());
  await page.evaluate(() => navigator.clipboard.writeText('paste-a\tpaste-b\r\npaste-c\tpaste-d'));
  await peer(page, 0).getByTestId('xlsx-scroll').press('ControlOrMeta+v');
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B2', 0))).toEqual([
    ['paste-a', 'paste-b'], ['paste-c', 'paste-d'],
  ]);
  await edit(page, 1, 'typed conflict', 0, 1);
  await queued(page, 0);
  await queued(page, 1);
  await resume(page);
  await expect.poll(async () => {
    const [a, b] = await pinnedEvaluate(page, () => Promise.all([
      window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B2', 0),
      window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B2', 1),
    ]));
    return JSON.stringify(a) === JSON.stringify(b) && a[0][0] === 'paste-a' &&
      ['paste-b', 'typed conflict'].includes(a[0][1]) && a[1][0] === 'paste-c' && a[1][1] === 'paste-d';
  }).toBe(true);
  await history(page, 0, 'undo');
  await converged(page, 'A1:B2', [['initial', 'typed conflict'], ['', '']]);
  await painted(page, 'initial', 'typed conflict');
  await history(page, 0, 'redo');
  await expect.poll(() => pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B2', 0))).toEqual([
    ['paste-a', expect.stringMatching(/^(paste-b|typed conflict)$/)], ['paste-c', 'paste-d'],
  ]);
  const expected = await pinnedEvaluate(page, () => window.__xlsxWorkerEditor.peerless!.cells(0, 'A1:B2', 0));
  await converged(page, 'A1:B2', expected);
  await painted(page, 'paste-a', expected[0][1], 'paste-c', 'paste-d');
  for (const saved of await savedPeers(page)) {
    expect(saved[sheet].cells.A1.value).toBe('paste-a');
    expect(saved[sheet].cells.B1.value).toBe(expected[0][1]);
    expect(saved[sheet].cells.A2.value).toBe('paste-c');
    expect(saved[sheet].cells.B2.value).toBe('paste-d');
  }
});

test('remote edits preserve a focused local cell editor on a different cell', async ({ page }) => {
  await page.evaluate(() => window.__xlsxWorkerEditor.pauseDelivery());
  await edit(page, 0, 'remote while typing');
  await queued(page, 0);
  await select(page, 1, 0, 1);
  await peer(page, 1).getByTestId('xlsx-scroll').press('F2');
  const editor = peer(page, 1).getByTestId('xlsx-cell-editor');
  await editor.fill('local-');
  await expect(editor).toBeFocused();
  await resume(page);
  await converged(page, 'A1', [['remote while typing']]);
  await expect(editor).toBeVisible();
  await expect(editor).toBeFocused();
  await expect(editor).toHaveValue('local-');
  await editor.pressSequentially('completed');
  await expect(editor).toHaveValue('local-completed');
  await editor.press('Enter');
  await converged(page, 'A1:B1', [['remote while typing', 'local-completed']]);
  await painted(page, 'remote while typing', 'local-completed');
  await history(page, 1, 'undo');
  await converged(page, 'A1:B1', [['remote while typing', '1']]);
  await history(page, 1, 'redo');
  await converged(page, 'A1:B1', [['remote while typing', 'local-completed']]);
  for (const saved of await savedPeers(page)) {
    expect(saved[sheet].cells.A1.value).toBe('remote while typing');
    expect(saved[sheet].cells.B1.value).toBe('local-completed');
  }
});

test('both peers save OOXML containing edits from each peer after local undo and redo', async ({ page }) => {
  await edit(page, 0, 'saved from first peer');
  await edit(page, 1, 'saved from second peer', 0, 1);
  await converged(page, 'A1:B1', [['saved from first peer', 'saved from second peer']]);
  await history(page, 1, 'undo');
  await converged(page, 'A1:B1', [['saved from first peer', '1']]);
  await history(page, 1, 'redo');
  await converged(page, 'A1:B1', [['saved from first peer', 'saved from second peer']]);
  for (const index of [0, 1]) {
    await select(page, index);
    await expect(peer(page, index).getByTestId('xlsx-formula-input')).toHaveValue('saved from first peer');
    await select(page, index, 0, 1);
    await expect(peer(page, index).getByTestId('xlsx-formula-input')).toHaveValue('saved from second peer');
  }
  await painted(page, 'saved from first peer', 'saved from second peer');
  for (const saved of await savedPeers(page)) {
    expect(saved[sheet].cells.A1.value).toBe('saved from first peer');
    expect(saved[sheet].cells.B1.value).toBe('saved from second peer');
  }
});
