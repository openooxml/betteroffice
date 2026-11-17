import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { inputReady, inspectDeck, slideCanvas } from './pptx-async-helpers';

const input = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });
const thumbnails = (page: Page) => page.getByRole('complementary', { name: 'Slides', exact: true }).getByRole('button');
const menu = (page: Page, name: string) => page.getByRole('menu', { name, exact: true });

async function fixture() {
  const zip = await JSZip.loadAsync(await readFile(new URL('../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url)));
  const path = 'ppt/slides/slide1.xml';
  const shape = (id: number, x: number, text?: string) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Object ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="2743200"/><a:ext cx="${text ? 3657600 : 1828800}" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill></p:spPr>${text ? `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400" cap="all"><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody>` : ''}</p:sp>`;
  zip.file(path, (await zip.file(path)!.async('string')).replace(/ (gridSpan|rowSpan|hMerge|vMerge)="[^"]*"/gu, '')
    .replace('</p:spTree>', `${shape(10, 914400, 'Menu source')}${shape(11, 6400800)}${shape(12, 8686800)}</p:spTree>`));
  return { name: 'context-menu.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

test.beforeEach(async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles(await fixture());
  await inputReady(page);
  await expect(thumbnails(page)).toHaveCount(1);
});

async function point(page: Page, x: number, y: number, right = false) {
  const box = (await slideCanvas(page).boundingBox())!;
  const location = { x: box.x + box.width * x, y: box.y + box.height * y };
  await page.mouse.click(location.x, location.y, { button: right ? 'right' : 'left' });
  return location;
}

async function selectText(page: Page) {
  await point(page, 0.15, 0.42);
  await expect(input(page)).toBeFocused();
  await input(page).press('ControlOrMeta+Home');
  await input(page).press('ControlOrMeta+Shift+End');
  await expect(page.getByRole('status').filter({ hasText: 'Text selected' })).toBeVisible();
}

async function openText(page: Page, table = false) {
  const result = menu(page, table ? 'Table menu' : 'Text menu');
  await expect(async () => {
    await input(page).press('Shift+F10');
    await expect(result).toBeVisible({ timeout: 1_000 });
  }).toPass();
  return result;
}

async function choose(page: Page, root: Locator, ...path: string[]) {
  let scope = root;
  for (const [index, name] of path.entries()) {
    const item = scope.locator(`[role="menuitem"][data-label=${JSON.stringify(name)}]`);
    await expect(item).not.toHaveAttribute('aria-disabled', 'true');
    if (index < path.length - 1) {
      await item.hover();
      scope = menu(page, name);
      await expect(scope).toBeVisible();
    } else await item.click();
  }
  await expect(root).toBeHidden();
}

async function save(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const path = info.outputPath(`${label}.pptx`);
  await (await downloading).saveAs(path);
  const bytes = await readFile(path);
  const zip = await JSZip.loadAsync(bytes);
  return { deck: await inspectDeck(page, bytes), xml: await zip.file('ppt/slides/slide1.xml')!.async('string') };
}

test('pointer menu selects its target, keeps the range, exposes a11y roles and restores focus', async ({ page }, info) => {
  await selectText(page);
  const location = await point(page, 0.15, 0.42, true);
  const root = menu(page, 'Text menu');
  await expect(root).toBeVisible();
  const box = (await root.boundingBox())!;
  expect(Math.abs(box.x - location.x)).toBeLessThan(2);
  expect(Math.abs(box.y - location.y)).toBeLessThan(2);
  await expect(root).toBeFocused();
  expect(await root.locator('[role="menuitem"]').evaluateAll((items) => items.map((item) => item.getAttribute('data-label'))))
    .toEqual(['Cut', 'Copy', 'Paste']);
  await expect(root.locator('[data-label="Copy"]')).toContainText(/Ctrl\+C|⌘C/);
  await info.attach('text-menu-aria', { body: await root.ariaSnapshot(), contentType: 'text/plain' });
  await page.keyboard.press('Escape');
  await expect(root).toBeHidden();
  await expect(input(page)).toBeFocused();
  await openText(page);
  await expect(root.locator('[data-label="Cut"]')).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(root.locator('[data-label="Copy"]')).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(input(page)).toBeFocused();
});

test('text menu copy, cut and paste use the clipboard and preserve paragraph breaks', async ({ page }, info) => {
  await selectText(page);
  await input(page).press('ControlOrMeta+C');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Menu source');
  await page.evaluate(() => navigator.clipboard.writeText(''));
  await choose(page, await openText(page), 'Copy');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Menu source');
  await choose(page, await openText(page), 'Cut');
  expect((await save(page, info, 'cut')).deck[0]!.shapes[0]!.text).toBe('');
  await point(page, 0.15, 0.42);
  await page.evaluate(() => navigator.clipboard.writeText('First line\nSecond line'));
  await choose(page, await openText(page), 'Paste');
  expect((await save(page, info, 'paste')).deck[0]!.shapes[0]!.paragraphs).toEqual(['First line', 'Second line']);
});

test('denied clipboard reads close the menu without changing text or breaking subsequent commands', async ({ page, context }, info) => {
  await selectText(page);
  await context.clearPermissions();
  await context.grantPermissions(['clipboard-write']);
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.evaluate(() => {
    Object.defineProperty(navigator.clipboard, 'readText', { configurable: true,
      value: () => Promise.reject(new DOMException('Denied', 'NotAllowedError')) });
  });
  await choose(page, await openText(page), 'Paste');
  expect((await save(page, info, 'denied')).deck[0]!.shapes[0]!.text).toBe('Menu source');
  expect(errors).toEqual([]);
});

test('shape menu selects an unselected object and executes all four paint-order moves and Delete', async ({ page }, info) => {
  const open = async () => { await point(page, 0.59, 0.46, true); return menu(page, 'Shape menu'); };
  await selectText(page);
  const root = await open();
  await info.attach('shape-menu-aria', { body: await root.ariaSnapshot(), contentType: 'text/plain' });
  await choose(page, root, 'Bring to front', 'Bring to front');
  expect((await save(page, info, 'front')).deck[0]!.ids).toEqual(['1', '2', '10', '12', '11']);
  await choose(page, await open(), 'Send to back', 'Send backward');
  expect((await save(page, info, 'backward')).deck[0]!.ids).toEqual(['1', '2', '10', '11', '12']);
  await choose(page, await open(), 'Send to back', 'Send to back');
  expect((await save(page, info, 'back')).deck[0]!.ids).toEqual(['1', '11', '2', '10', '12']);
  await choose(page, await open(), 'Bring to front', 'Bring forward');
  expect((await save(page, info, 'forward')).deck[0]!.ids).toEqual(['1', '2', '11', '10', '12']);
  await choose(page, await open(), 'Delete');
  expect((await save(page, info, 'deleted-object')).deck[0]!.ids).not.toContain('11');
});

test('table menu selects a cell and executes every insert and delete operation', async ({ page }, info) => {
  await point(page, 0.24, 0.17, true);
  await expect(menu(page, 'Table menu')).toBeVisible();
  await info.attach('table-menu-aria', { body: await menu(page, 'Table menu').ariaSnapshot(), contentType: 'text/plain' });
  await page.keyboard.press('Escape');
  await point(page, 0.24, 0.17);
  await expect(input(page)).toBeFocused();
  const caretMenu = await openText(page, true);
  for (const name of ['Cut', 'Copy']) await expect(caretMenu.locator(`[data-label="${name}"]`)).toHaveAttribute('aria-disabled', 'true');
  await page.keyboard.press('Escape');
  await expect(input(page)).toBeFocused();
  for (let cell = 0; cell < 3; cell += 1) await input(page).press('Tab');
  await expect(page.getByRole('status').filter({ hasText: 'Text selected' })).toBeVisible();
  await choose(page, await openText(page, true), 'Copy');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Middle');
  await choose(page, await openText(page, true), 'Cut');
  expect((await save(page, info, 'table-cut')).xml).not.toContain('>Middle<');
  await choose(page, await openText(page, true), 'Paste');
  expect((await save(page, info, 'table-paste')).xml).toContain('>Middle<');
  const dimensions = async (label: string) => {
    const { xml } = await save(page, info, label);
    return { rows: (xml.match(/<a:tr\b/gu) ?? []).length, columns: (xml.match(/<a:gridCol\b/gu) ?? []).length };
  };
  const open = async () => {
    await point(page, 0.24, 0.17);
    return openText(page, true);
  };
  await choose(page, await open(), 'Insert', 'Insert rows above');
  expect(await dimensions('row-above')).toEqual({ rows: 4, columns: 3 });
  await choose(page, await open(), 'Insert', 'Insert rows below');
  expect(await dimensions('row-below')).toEqual({ rows: 5, columns: 3 });
  await choose(page, await open(), 'Insert', 'Insert columns to the left');
  expect(await dimensions('column-left')).toEqual({ rows: 5, columns: 4 });
  await choose(page, await open(), 'Insert', 'Insert columns to the right');
  expect(await dimensions('column-right')).toEqual({ rows: 5, columns: 5 });
  await choose(page, await open(), 'Delete', 'Delete rows');
  expect(await dimensions('delete-row')).toEqual({ rows: 4, columns: 5 });
  await choose(page, await open(), 'Delete', 'Delete columns');
  expect(await dimensions('delete-column')).toEqual({ rows: 4, columns: 4 });
  await choose(page, await open(), 'Delete', 'Delete table');
  expect((await save(page, info, 'delete-table')).xml).not.toContain('<a:tbl>');
});

test('thumbnail menu activates the target and inserts, duplicates and deletes slides with last-slide gating', async ({ page }, info) => {
  await thumbnails(page).first().click({ button: 'right' });
  const root = menu(page, 'Slide menu');
  await expect(root.locator('[data-label="Delete slide"]')).toHaveAttribute('aria-disabled', 'true');
  await info.attach('slide-menu-aria', { body: await root.ariaSnapshot(), contentType: 'text/plain' });
  await page.keyboard.press('Escape');
  await expect(thumbnails(page).first()).toBeFocused();
  await thumbnails(page).first().press('Shift+F10');
  await expect(root.locator('[data-label="New slide"]')).toBeFocused();
  await choose(page, root, 'New slide');
  await expect(thumbnails(page)).toHaveCount(2);
  await expect(thumbnails(page).nth(1)).toHaveAttribute('aria-current', 'page');
  await thumbnails(page).first().click({ button: 'right' });
  await expect(thumbnails(page).first()).toHaveAttribute('aria-current', 'page');
  await choose(page, root, 'Duplicate slide');
  await expect(thumbnails(page)).toHaveCount(3);
  await expect(thumbnails(page).nth(1)).toHaveAttribute('aria-current', 'page');
  const duplicated = (await save(page, info, 'duplicate')).deck;
  expect(duplicated[1]!.shapes).toEqual(duplicated[0]!.shapes);
  await thumbnails(page).nth(2).click({ button: 'right' });
  await choose(page, root, 'Delete slide');
  await expect(thumbnails(page)).toHaveCount(2);
  await thumbnails(page).nth(1).click({ button: 'right' });
  await choose(page, root, 'Delete slide');
  await expect(thumbnails(page)).toHaveCount(1);
  await thumbnails(page).first().click({ button: 'right' });
  await expect(root.locator('[data-label="Delete slide"]')).toHaveAttribute('aria-disabled', 'true');
  await page.keyboard.press('Escape');
  expect((await save(page, info, 'one-slide')).deck).toHaveLength(1);
});

test('missing clipboard API disables Paste before execution', async ({ page }) => {
  await selectText(page);
  await page.evaluate(() => Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined }));
  const root = await openText(page);
  await expect(root.locator('[data-label="Paste"]')).toHaveAttribute('aria-disabled', 'true');
  await page.keyboard.press('Escape');
});

test('empty slide space leaves the native context menu event available', async ({ page }) => {
  const prevented = await slideCanvas(page).evaluate((canvas) => {
    const bounds = canvas.getBoundingClientRect();
    const event = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2,
      clientX: bounds.left + bounds.width * 0.9, clientY: bounds.top + bounds.height * 0.9 });
    canvas.dispatchEvent(event);
    return event.defaultPrevented;
  });
  expect(prevented).toBe(false);
  await expect(page.getByRole('menu')).toHaveCount(0);
});
