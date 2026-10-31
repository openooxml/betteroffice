import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page, arm: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Layout menu</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'layout-menu.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }),
  });
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('formatting-bar')).toBeVisible();
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
}

async function openLayoutSubmenu(page: Page, index: number, name: string) {
  const trigger = page.getByRole('menubar').getByRole('menuitem', { name: 'Layout', exact: true });
  await trigger.focus();
  await page.keyboard.press('Enter');
  await expect(page.getByRole('menuitem', { name: /^margins$/i })).toBeFocused();
  for (let i = 0; i < index; i++) await page.keyboard.press('ArrowDown');
  const label = new RegExp(`^${name}$`, 'i');
  await expect(page.getByRole('menuitem', { name: label })).toBeFocused();
  await expect(page.getByRole('menuitem', { name: label })).not.toHaveAttribute('aria-disabled', 'true');
  await page.keyboard.press('ArrowRight');
  await expect(page.getByRole('menu', { name: label })).toBeVisible();
}

async function rulerMargin(page: Page) {
  return page.locator('.docx-horizontal-ruler > div').first().evaluate((zone) => zone.getBoundingClientRect().width);
}

async function rulerWidth(page: Page) {
  return page.locator('.docx-horizontal-ruler').evaluate((ruler) => ruler.getBoundingClientRect().width);
}

async function savedSection(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(name + '.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate(source => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    const section = document.getElementsByTagNameNS('*', 'sectPr')[0];
    const margins = section.getElementsByTagNameNS('*', 'pgMar')[0];
    const size = section.getElementsByTagNameNS('*', 'pgSz')[0];
    const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return {
      top: margins.getAttributeNS(namespace, 'top'),
      bottom: margins.getAttributeNS(namespace, 'bottom'),
      left: margins.getAttributeNS(namespace, 'left'),
      right: margins.getAttributeNS(namespace, 'right'),
      width: size.getAttributeNS(namespace, 'w'),
      height: size.getAttributeNS(namespace, 'h'),
      orientation: size.getAttributeNS(namespace, 'orient'),
    };
  }, xml);
}

for (const arm of ['&peerless=0', '']) {
  test(`Layout keyboard presets save margins and landscape (${arm ? 'legacy' : 'peerless'})`, async ({ page }, info) => {
    await open(page, arm);
    await openLayoutSubmenu(page, 0, 'Margins');
    const normal = page.getByRole('menuitemradio', { name: 'Normal', exact: true });
    const narrow = page.getByRole('menuitemradio', { name: 'Narrow', exact: true });
    await expect(normal).toBeFocused();
    await expect(normal).toBeChecked();
    await page.keyboard.press('ArrowDown');
    await expect(narrow).toBeFocused();
    const margin = await rulerMargin(page);
    await page.keyboard.press('Enter');
    await expect(page.getByRole('menu', { name: 'Layout', exact: true })).toBeHidden();
    await expect.poll(() => rulerMargin(page)).toBeLessThan(margin);
    await openLayoutSubmenu(page, 0, 'Margins');
    await expect(narrow).toBeChecked();
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    expect(await savedSection(page, info, 'narrow')).toMatchObject({ top: '720', bottom: '720', left: '720', right: '720', width: '12240', height: '15840' });

    await openLayoutSubmenu(page, 1, 'Orientation');
    await expect(page.getByRole('menuitemradio', { name: 'Portrait', exact: true })).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(page.getByRole('menuitemradio', { name: 'Landscape', exact: true })).toBeFocused();
    const width = await rulerWidth(page);
    await page.keyboard.press('Enter');
    await expect(page.getByRole('menu', { name: 'Layout', exact: true })).toBeHidden();
    await expect.poll(() => rulerWidth(page)).toBeGreaterThan(width);
    await openLayoutSubmenu(page, 1, 'Orientation');
    await expect(page.getByRole('menuitemradio', { name: 'Landscape', exact: true })).toBeChecked();
    await page.keyboard.press('Escape');
    await page.keyboard.press('Escape');
    expect(await savedSection(page, info, 'landscape')).toEqual({ top: '720', bottom: '720', left: '720', right: '720', width: '15840', height: '12240', orientation: 'landscape' });
  });
}
