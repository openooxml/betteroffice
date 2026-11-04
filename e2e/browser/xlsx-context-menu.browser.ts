import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { clipboardWorkbook, selectCell, workbookXml } from './xlsx-async-clipboard-helpers';

const sheet = 'xl/worksheets/sheet1.xml';

test.beforeEach(async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(await clipboardWorkbook());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('11');
});

async function save(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.xlsx');
  await (await downloading).saveAs(file);
  return workbookXml(page, await readFile(file));
}

const cellMenu = (page: Page) => page.getByRole('menu', { name: 'Cell menu', exact: true });
const tabMenu = (page: Page) => page.getByRole('menu', { name: 'Sheet tab menu', exact: true });

async function openCellMenu(page: Page) {
  await page.getByTestId('xlsx-scroll').press('Shift+F10');
  await expect(cellMenu(page)).toBeVisible();
  return cellMenu(page);
}

async function choose(page: Page, menu: ReturnType<typeof cellMenu>, ...path: string[]) {
  let scope = menu;
  for (const [index, name] of path.entries()) {
    const item = scope.getByRole('menuitem', { name, exact: true })
      .or(scope.getByRole('menuitemradio', { name, exact: true }));
    await expect(item).not.toHaveAttribute('aria-disabled', 'true');
    if (index < path.length - 1) {
      await item.hover();
      scope = page.getByRole('menu', { name, exact: true });
      await expect(scope).toBeVisible();
    } else await item.click();
  }
  await expect(menu).toBeHidden();
}

test('cell menu opens at the pointer, selects the cell under it and lists Excel items in order', async ({ page }, info) => {
  await selectCell(page, 2, 2);
  const canvas = page.locator('.editor-stage canvas').first();
  const box = (await canvas.boundingBox())!;
  const pointer = { x: box.x + 16, y: box.y + 16 };
  await page.mouse.click(pointer.x, pointer.y, { button: 'right' });
  const menu = cellMenu(page);
  await expect(menu).toBeVisible();
  const placed = (await menu.boundingBox())!;
  expect(Math.abs(placed.x - pointer.x)).toBeLessThan(2);
  expect(Math.abs(placed.y - pointer.y)).toBeLessThan(2);
  await expect(menu).toBeFocused();
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  const aria = await menu.ariaSnapshot();
  await info.attach('cell-menu-aria', { body: aria, contentType: 'text/plain' });
  const names = await menu.locator(':scope > div > [role=menuitem]').evaluateAll((items) => items.map((item) => item.getAttribute('data-label')));
  expect(names).toEqual(['Cut', 'Copy', 'Paste', 'Paste as plain text', 'Insert', 'Delete', 'Clear contents', 'Number format']);
  await expect(menu.getByRole('menuitem', { name: /^Copy/ })).toContainText(/Ctrl\+C|⌘C/);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(page.getByTestId('xlsx-scroll')).toBeFocused();
});

test('Shift+F10 opens the cell menu from the keyboard with the first item focused', async ({ page }) => {
  const menu = await openCellMenu(page);
  await expect(menu.getByRole('menuitem', { name: /^Cut/ })).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(menu.getByRole('menuitem', { name: /^Copy/ })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('xlsx-scroll')).toBeFocused();
});

test('copy, paste, plain paste, cut and clear contents from the cell menu', async ({ page }, info) => {
  await choose(page, await openCellMenu(page), 'Copy');
  await selectCell(page, 2, 2);
  await choose(page, await openCellMenu(page), 'Paste');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('11');
  await selectCell(page, 3, 2);
  await choose(page, await openCellMenu(page), 'Paste as plain text');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('11');
  await selectCell(page, 0, 1);
  await choose(page, await openCellMenu(page), 'Cut');
  await selectCell(page, 4, 4);
  await choose(page, await openCellMenu(page), 'Paste');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('12');
  await selectCell(page, 1, 0);
  await choose(page, await openCellMenu(page), 'Clear contents');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('');
  const cells = (await save(page, info, 'clipboard'))[sheet].cells;
  expect(cells.C3?.value).toBe('11');
  expect(cells.C4).toMatchObject({ value: '11', bold: false, fill: '' });
  expect(cells.E5?.value).toBe('12');
  expect(cells.B1?.value ?? '').toBe('');
  expect(cells.A2?.value ?? '').toBe('');
});

test('insert and delete rows and columns from the cell menu', async ({ page }, info) => {
  await choose(page, await openCellMenu(page), 'Insert', 'Insert rows above');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('');
  await choose(page, await openCellMenu(page), 'Insert', 'Insert columns left');
  let cells = (await save(page, info, 'inserted'))[sheet].cells;
  expect(cells.B2?.value).toBe('11');
  expect(cells.C2?.value).toBe('12');
  await choose(page, await openCellMenu(page), 'Delete', 'Delete rows');
  await choose(page, await openCellMenu(page), 'Delete', 'Delete columns');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('11');
  await choose(page, await openCellMenu(page), 'Insert', 'Insert rows below');
  await choose(page, await openCellMenu(page), 'Insert', 'Insert columns right');
  cells = (await save(page, info, 'deleted'))[sheet].cells;
  expect(cells.A1?.value).toBe('11');
  expect(cells.C1?.value).toBe('12');
  expect(cells.A3?.value).toBe('21');
});

test('number format from the cell menu', async ({ page }, info) => {
  await choose(page, await openCellMenu(page), 'Number format', 'Percent');
  const menu = await openCellMenu(page);
  await menu.getByRole('menuitem', { name: 'Number format', exact: true }).hover();
  await expect(page.getByRole('menu', { name: 'Number format' }).getByRole('menuitemradio', { name: 'Percent' }))
    .toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');
  expect((await save(page, info, 'format'))[sheet].cells.A1?.value).toBe('11');
});

test('sheet tab menu inserts, renames, moves and deletes sheets', async ({ page }, info) => {
  const tab = (name: string) => page.getByRole('tab', { name, exact: true });
  await tab('Other').click({ button: 'right' });
  await expect(tab('Other')).toHaveAttribute('aria-selected', 'true');
  await info.attach('tab-menu-aria', { body: await tabMenu(page).ariaSnapshot(), contentType: 'text/plain' });
  await expect(tabMenu(page).getByRole('menuitem', { name: 'Move right' })).toHaveAttribute('aria-disabled', 'true');
  await choose(page, tabMenu(page), 'Insert sheet');
  await expect(tab('Sheet3')).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('tablist').getByRole('tab')).toHaveText(['Sheet', 'Sheet3', 'Other']);
  await tab('Sheet3').click({ button: 'right' });
  await choose(page, tabMenu(page), 'Rename');
  const rename = page.getByRole('textbox', { name: 'Rename sheet' });
  await expect(rename).toBeFocused();
  await rename.fill('Bad:Name');
  await rename.press('Enter');
  await expect(page.getByTestId('xlsx-input-refusal')).toContainText('Sheet names need 1 to 31 characters');
  await tab('Sheet3').dblclick();
  await page.getByRole('textbox', { name: 'Rename sheet' }).fill('Budget');
  await page.getByRole('textbox', { name: 'Rename sheet' }).press('Enter');
  await expect(tab('Budget')).toBeFocused();
  await tab('Budget').press('Shift+F10');
  await expect(tabMenu(page).getByRole('menuitem').first()).toBeFocused();
  await choose(page, tabMenu(page), 'Move right');
  await expect(page.getByRole('tablist').getByRole('tab')).toHaveText(['Sheet', 'Other', 'Budget']);
  await tab('Budget').click({ button: 'right' });
  await choose(page, tabMenu(page), 'Move left');
  await expect(page.getByRole('tablist').getByRole('tab')).toHaveText(['Sheet', 'Budget', 'Other']);
  await tab('Other').click({ button: 'right' });
  await choose(page, tabMenu(page), 'Delete sheet');
  await expect(page.getByRole('tablist').getByRole('tab')).toHaveText(['Sheet', 'Budget']);
  await tab('Budget').click({ button: 'right' });
  await choose(page, tabMenu(page), 'Delete sheet');
  await tab('Sheet').click({ button: 'right' });
  await expect(tabMenu(page).getByRole('menuitem', { name: 'Delete sheet' })).toHaveAttribute('aria-disabled', 'true');
  await page.keyboard.press('Escape');
  const saved = await save(page, info, 'sheets');
  expect(Object.keys(saved)).toHaveLength(1);
});
