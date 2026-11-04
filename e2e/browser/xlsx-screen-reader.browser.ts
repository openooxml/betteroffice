import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function workbook() {
  const zip = new JSZip();
  const part = (n: number) => `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${part(1)}${part(2)}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>21</v></c><c r="B1"><f>A1*2</f><v>42</v></c><c r="C1" t="inlineStr"><is><t>Total</t></is></c></row></sheetData></worksheet>`);
  zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page) {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'screen-reader.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await workbook(),
  });
  await expect(page.getByTestId('xlsx-scroll').locator('canvas').first()).toHaveAttribute('data-worker-zoom', '1', { timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
}

const live = (page: Page) => page.getByTestId('xlsx-live-region');
const heard = (page: Page, text: string) => expect(live(page)).toHaveText(new RegExp(`^${text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s?$`));

test('xlsx screen reader: the grid is a labelled focus target and announces the active cell', async ({ page }, info) => {
  await open(page);
  const grid = page.getByTestId('xlsx-scroll');
  await info.attach('aria-before.yml', { body: await page.locator('body').ariaSnapshot(), contentType: 'text/yaml' });
  await expect(grid).toMatchAriaSnapshot('- application "Budget"');
  await expect(grid).toHaveAttribute('aria-roledescription', 'worksheet');
  await grid.focus();
  await heard(page, 'A1, 21');
  await grid.press('ArrowRight');
  await heard(page, 'B1, 42, has formula');
  await grid.press('ArrowRight');
  await heard(page, 'C1, Total');
  await grid.press('Shift+ArrowLeft');
  await heard(page, 'B1 to C1 selected, 1 row by 2 columns');
  await grid.press('ArrowDown');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(/^[BC]2$/);
  const address = await page.getByTestId('xlsx-name-box').inputValue();
  await heard(page, `${address}, empty`);
  await expect(page.getByRole('status').filter({ hasText: `${address}, empty` })).toHaveCount(1);
  await info.attach('aria-after.yml', { body: await page.locator('body').ariaSnapshot(), contentType: 'text/yaml' });
});

test('xlsx screen reader: rapid navigation is announced once, for the cell it stops on', async ({ page }) => {
  await open(page);
  const grid = page.getByTestId('xlsx-scroll');
  await grid.focus();
  await heard(page, 'A1, 21');
  await page.evaluate(() => {
    const region = document.querySelector('[data-testid="xlsx-live-region"]')!;
    const seen: string[] = [];
    new MutationObserver(() => seen.push(region.textContent ?? '')).observe(region, { childList: true, characterData: true, subtree: true });
    (window as unknown as { announcements: string[] }).announcements = seen;
  });
  for (let index = 0; index < 6; index++) await grid.press('ArrowDown', { delay: 20 });
  await heard(page, 'A7, empty');
  const seen = await page.evaluate(() => (window as unknown as { announcements: string[] }).announcements);
  expect(seen.map((text) => text.trim())).toEqual(['A7, empty']);
});

test('xlsx screen reader: sheets switch by keyboard and name the new sheet', async ({ page }) => {
  await open(page);
  const grid = page.getByTestId('xlsx-scroll');
  await grid.focus();
  await heard(page, 'A1, 21');
  await grid.press('Control+Alt+PageDown');
  await heard(page, 'Notes, sheet 2 of 2, A1, empty');
  await grid.press('Control+Shift+PageUp');
  await heard(page, 'Budget, sheet 1 of 2, A1, 21');
  await grid.press('Control+Alt+PageDown');
  await heard(page, 'Notes, sheet 2 of 2, A1, empty');
  await expect(page.getByRole('tab', { name: 'Notes', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(grid).toHaveAccessibleName('Notes');
  await grid.press('F6');
  const tab = page.getByRole('tab', { name: 'Notes', exact: true });
  await expect(tab).toBeFocused();
  await tab.press('ArrowLeft');
  await expect(page.getByRole('tab', { name: 'Budget', exact: true })).toBeFocused();
  await expect(page.getByRole('tab', { name: 'Budget', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('tab', { name: 'Notes', exact: true })).toHaveAttribute('tabindex', '-1');
  await page.keyboard.press('F6');
  await expect(page.getByTestId('xlsx-toolbar').locator(':focus')).toHaveCount(1);
  await page.keyboard.press('F6');
  await expect(grid).toBeFocused();
  await heard(page, 'A1, 21');
  await grid.press('F2');
  await expect(page.getByRole('textbox', { name: 'Edit cell A1' })).toBeFocused();
});
