import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');

async function open(page: Page) {
  await page.goto('/?format=docx');
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles(resolve(root, 'apps/demo/public/betteroffice-demo.docx'));
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({
    timeout: 60_000,
  });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
}

async function focusDocument(page: Page) {
  await page.locator('.canvas-page canvas').first().click({
    position: { x: 150, y: 150 },
  });
  const input = page.getByTestId('yrs-input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
}

async function savedTableShape(page: Page, info: TestInfo) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('inserted-table.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((source) => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    const table = document.getElementsByTagNameNS('*', 'tbl')[0];
    if (!table) return null;
    return Array.from(table.getElementsByTagNameNS('*', 'tr')).map((row) =>
      row.getElementsByTagNameNS('*', 'tc').length
    );
  }, xml);
}

function insertTrigger(page: Page) {
  return page.getByRole('menubar').getByRole('menuitem', { name: 'Insert' });
}

test('docx Insert menu opens from Enter, Space, ArrowDown, and ArrowUp', async ({ page }) => {
  await open(page);
  const trigger = insertTrigger(page);
  const menu = page.getByRole('menu', { name: 'Insert' });

  for (const key of ['Enter', 'Space', 'ArrowDown', 'ArrowUp']) {
    await trigger.focus();
    await page.keyboard.press(key);
    await expect(menu).toBeVisible();
    await expect(
      key === 'ArrowUp' ? menu.getByRole('menuitem').last() : menu.getByRole('menuitem').first()
    ).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(menu).toBeHidden();
  }
});

test('docx Insert table grid supports keyboard navigation and inserts its dimensions', async ({
  page,
}, info) => {
  await open(page);
  await focusDocument(page);
  const trigger = insertTrigger(page);
  await trigger.focus();
  await page.keyboard.press('Enter');

  const menu = page.getByRole('menu', { name: 'Insert' });
  const tableItem = menu.getByRole('menuitem', { name: 'Table', exact: true });
  await tableItem.focus();
  await page.keyboard.press('ArrowRight');

  const grid = page.getByRole('grid', { name: 'Table size selector' });
  await expect(grid).toBeVisible();
  const cells = grid.getByRole('gridcell');
  await expect(cells.first()).toBeFocused();

  await page.keyboard.press('ArrowRight');
  await page.getByText('2 × 1', { exact: true }).waitFor();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  await page.getByText('3 × 2', { exact: true }).waitFor();
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();

  expect(await savedTableShape(page, info)).toEqual([3, 3]);
});

test('docx Break submenu supports keyboard opening and arrow navigation', async ({ page }) => {
  await open(page);
  const trigger = insertTrigger(page);
  const menu = page.getByRole('menu', { name: 'Insert' });
  const submenu = page.getByRole('menu', { name: 'Break' });

  await trigger.focus();
  await page.keyboard.press('Enter');
  const breakItem = menu.getByRole('menuitem', { name: 'Break' });
  await breakItem.focus();
  await page.keyboard.press('ArrowRight');
  await expect(submenu).toBeVisible();

  const choices = submenu.getByRole('menuitem');
  await expect(choices.first()).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(choices.nth(1)).toBeFocused();
  await page.keyboard.press('ArrowUp');
  await expect(choices.first()).toBeFocused();

  await page.keyboard.press('Escape');
  await expect(submenu).toBeHidden();
  await expect(menu).toBeVisible();
  await expect(breakItem).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(trigger).toBeFocused();
});

test('docx Insert submenus continue to open on mouse hover', async ({ page }) => {
  await open(page);
  await insertTrigger(page).click();
  const menu = page.getByRole('menu', { name: 'Insert' });
  const grid = page.getByRole('grid', { name: 'Table size selector' });
  const breakMenu = page.getByRole('menu', { name: 'Break' });

  await menu.getByRole('menuitem', { name: 'Table', exact: true }).hover();
  await expect(grid).toBeVisible();
  await page.mouse.move(0, 0);
  await expect(grid).toBeHidden();

  await menu.getByRole('menuitem', { name: 'Break' }).hover();
  await expect(breakMenu).toBeVisible();
});

test('docx menubar arrows move between triggers with a single tab stop', async ({ page }) => {
  await open(page);
  const menubar = page.getByRole('menubar');
  const triggers = menubar.getByRole('menuitem');
  await expect(menubar.locator('[tabindex="0"]')).toHaveCount(1);
  await insertTrigger(page).focus();
  await page.keyboard.press('ArrowLeft');
  await expect(menubar.getByRole('menuitem', { name: 'Format', exact: true })).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(insertTrigger(page)).toBeFocused();
  await expect(page.getByRole('menu')).toHaveCount(0);
  await page.keyboard.press('Home');
  await expect(triggers.first()).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(triggers.last()).toBeFocused();
  await expect(menubar.locator('[tabindex="0"]')).toHaveCount(1);
});

test('docx table grid keeps row and boundary navigation inside the picker', async ({ page }) => {
  await open(page);
  await insertTrigger(page).focus();
  await page.keyboard.press('Enter');
  const menu = page.getByRole('menu', { name: 'Insert' });
  const table = menu.getByRole('menuitem', { name: 'Table', exact: true });
  await table.focus();
  await page.keyboard.press('Space');
  const grid = page.getByRole('grid', { name: 'Table size selector' });
  const cells = grid.getByRole('gridcell');
  await expect(grid.getByRole('row')).toHaveCount(6);
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('ArrowUp');
  await expect(cells.first()).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('End');
  await expect(cells.nth(11)).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(cells.nth(11)).toBeFocused();
  await page.keyboard.press('Home');
  await expect(cells.nth(6)).toBeFocused();
  await page.keyboard.press('Control+End');
  await expect(cells.last()).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(cells.last()).toBeFocused();
  await page.keyboard.press('Control+Home');
  await expect(cells.first()).toBeFocused();
  await expect(grid.locator('[tabindex="0"]')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(grid).toBeHidden();
  await expect(menu).toBeVisible();
  await expect(table).toBeFocused();
});

test('docx Tab and Shift+Tab leave Insert submenus and close the menu tree', async ({ page }) => {
  await open(page);
  for (const key of ['Tab', 'Shift+Tab']) {
    await insertTrigger(page).focus();
    await page.keyboard.press('Enter');
    const menu = page.getByRole('menu', { name: 'Insert' });
    await menu.getByRole('menuitem', { name: 'Break', exact: true }).focus();
    await page.keyboard.press('Enter');
    await expect(page.getByRole('menu', { name: 'Break' })).toBeVisible();
    await page.keyboard.press(key);
    await expect(menu).toBeHidden();
    await expect(page.getByRole('menubar').locator(':focus')).toHaveCount(0);
  }
});
