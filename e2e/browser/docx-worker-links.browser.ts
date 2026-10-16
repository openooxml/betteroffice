import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const external = 'https://example.com/original';

async function open(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="link" Type="${r}/hyperlink" Target="${external}" TargetMode="External"/><Relationship Id="header" Type="${r}/header" Target="header.xml"/></Relationships>`);
  zip.file('word/_rels/header.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="headLink" Type="${r}/hyperlink" Target="https://example.com/header" TargetMode="External"/></Relationships>`);
  zip.file('word/header.xml', `<w:hdr xmlns:w="${w}" xmlns:r="${r}"><w:p><w:hyperlink r:id="headLink"><w:r><w:t>Header link</w:t></w:r></w:hyperlink></w:p></w:hdr>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>` +
    '<w:p><w:r><w:t>Bare text</w:t></w:r></w:p>' +
    '<w:p><w:hyperlink r:id="link" w:tooltip="Original tip"><w:r><w:t>External link</w:t></w:r></w:hyperlink></w:p>' +
    '<w:p><w:hyperlink w:anchor="target"><w:r><w:t>Jump to target</w:t></w:r></w:hyperlink></w:p>' +
    '<w:p><w:hyperlink w:anchor="cellTarget"><w:r><w:t>Jump to cell</w:t></w:r></w:hyperlink></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:hyperlink r:id="link"><w:r><w:t>Cell link</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl>' +
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:bookmarkStart w:id="1" w:name="target"/><w:r><w:t>Destination</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:bookmarkStart w:id="2" w:name="cellTarget"/><w:r><w:t>Cell destination</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p></w:tc></w:tr></w:tbl>' +
    '<w:p/><w:sectPr><w:headerReference w:type="default" r:id="header"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'links.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function saved(page: Page, info: TestInfo) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('links.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  const links = await page.evaluate(({ xml, rels, w, r }) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');
    const relationships = parser.parseFromString(rels, 'application/xml');
    return Array.from(doc.getElementsByTagNameNS(w, 'hyperlink')).map(link => ({
      text: link.textContent,
      href: Array.from(relationships.getElementsByTagName('Relationship')).find(rel => rel.getAttribute('Id') === link.getAttributeNS(r, 'id'))?.getAttribute('Target') ?? `#${link.getAttributeNS(w, 'anchor')}`,
    }));
  }, { xml, rels, w, r });
  return { zip, xml, rels, links };
}

async function clickText(page: Page, text: string) {
  const run = page.locator('.canvas-page-mirror .layout-paragraph').filter({ hasText: text }).first().locator('a[href]').first();
  await expect(run).toBeAttached();
  await run.evaluate(node => node.scrollIntoView({ block: 'center' }));
  const box = (await run.boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
}

test('worker Mod+K inserts links on a selection and at a bare caret', async ({ page }, info) => {
  const input = await open(page);
  await input.press('Shift+End');
  await input.press('ControlOrMeta+k');
  let dialog = page.getByRole('dialog', { name: 'Insert Hyperlink', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByLabel('Display Text', { exact: true })).toHaveValue('Bare text');
  await dialog.getByLabel('URL', { exact: true }).fill('https://example.com/selection');
  await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
  await expect(dialog).toHaveCount(0);
  expect((await saved(page, info)).links).toContainEqual({ text: 'Bare text', href: 'https://example.com/selection' });
  await input.press('ControlOrMeta+End');
  await input.press('ControlOrMeta+k');
  dialog = page.getByRole('dialog', { name: 'Insert Hyperlink', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('Display Text', { exact: true }).fill('New link');
  await dialog.getByLabel('URL', { exact: true }).fill('https://example.com/caret');
  await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
  await expect(dialog).toHaveCount(0);
  expect((await saved(page, info)).links).toContainEqual({ text: 'New link', href: 'https://example.com/caret' });
  await input.pressSequentially(' after');
  const result = await saved(page, info);
  expect(result.xml).toContain(' after');
});

test('worker link under the caret edits and removes from the dialog', async ({ page }, info) => {
  const input = await open(page);
  await clickText(page, 'External link');
  await page.keyboard.press('Escape');
  await input.press('ControlOrMeta+k');
  const dialog = page.getByRole('dialog', { name: 'Edit Hyperlink', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByLabel('Display Text', { exact: true })).toBeEditable();
  await expect(dialog.getByLabel('Display Text', { exact: true })).toHaveValue('External link');
  await expect(dialog.getByLabel('URL', { exact: true })).toHaveValue(external);
  await dialog.getByLabel('Display Text', { exact: true }).fill('Changed link');
  await dialog.getByLabel('URL', { exact: true }).fill('https://example.com/changed');
  await expect(dialog.getByLabel('Display Text', { exact: true })).toHaveValue('Changed link');
  await expect(dialog.getByLabel('URL', { exact: true })).toHaveValue('https://example.com/changed');
  await dialog.getByRole('button', { name: 'Update', exact: true }).click();
  await expect(dialog).toHaveCount(0);
  expect((await saved(page, info)).links).toContainEqual({ text: 'Changed link', href: 'https://example.com/changed' });
  await input.press('ControlOrMeta+k');
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Remove Link', exact: true }).click();
  await expect(dialog).toHaveCount(0);
  const result = await saved(page, info);
  expect(result.xml).toContain('Changed link');
  expect(result.links.some(link => link.text === 'Changed link')).toBe(false);
});

test('worker clicking a body link supports popup copy edit and remove', async ({ page, context }, info) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.addInitScript(() => {
    const state = window as typeof window & { linkNavigation?: unknown[] };
    window.open = (...args) => { state.linkNavigation = args; return null; };
  });
  await open(page);
  await clickText(page, 'External link');
  const popup = page.locator('.oox-hyperlink-popup');
  await expect(popup).toBeVisible();
  await expect(popup).toContainText(external);
  await popup.getByRole('link', { name: external, exact: true }).click();
  expect(await page.evaluate(() => (window as typeof window & { linkNavigation?: unknown[] }).linkNavigation))
    .toEqual([external, '_blank', 'noopener,noreferrer']);
  await popup.getByRole('button', { name: 'Copy link', exact: true }).click();
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe(external);
  await popup.getByRole('button', { name: 'Edit link', exact: true }).click();
  await popup.getByRole('textbox').nth(0).fill('Popup link');
  await popup.getByRole('textbox').nth(1).fill('https://example.com/popup');
  await popup.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(popup).toHaveCount(0);
  expect((await saved(page, info)).links).toContainEqual({ text: 'Popup link', href: 'https://example.com/popup' });
  await clickText(page, 'Popup link');
  await popup.getByRole('button', { name: 'Remove link', exact: true }).click();
  await expect(popup).toHaveCount(0);
  const result = await saved(page, info);
  expect(result.xml).toContain('Popup link');
  expect(result.links.some(link => link.text === 'Popup link')).toBe(false);
});

test('worker table cell links open the popup and viewing hides edit and remove', async ({ page }, info) => {
  await open(page);
  await clickText(page, 'Cell link');
  const popup = page.locator('.oox-hyperlink-popup');
  await expect(popup).toBeVisible();
  await expect(popup).toContainText(external);
  await page.keyboard.press('Escape');
  const mode = page.getByRole('button', { name: 'Editing', exact: true });
  if (await mode.isVisible()) await mode.click();
  else {
    await page.getByRole('button', { name: 'More actions', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Editing mode', exact: true }).hover();
  }
  await page.getByRole('menuitemradio', { name: /^Viewing\b/ }).click();
  await clickText(page, 'Cell link');
  await expect(popup).toBeVisible();
  await expect(popup.getByRole('button', { name: 'Copy link', exact: true })).toBeVisible();
  await expect(popup.getByRole('button', { name: 'Edit link', exact: true })).toHaveCount(0);
  await expect(popup.getByRole('button', { name: 'Remove link', exact: true })).toHaveCount(0);
  await page.keyboard.press('Escape');
  await clickText(page, 'External link');
  await expect(popup).toBeVisible();
  await expect(popup.getByRole('button', { name: 'Edit link', exact: true })).toHaveCount(0);
  await expect(popup.getByRole('button', { name: 'Remove link', exact: true })).toHaveCount(0);
  expect((await saved(page, info)).links).toContainEqual({ text: 'Cell link', href: external });
});

test('worker header link opens a popup only in an open editing band', async ({ page }, info) => {
  await open(page);
  await clickText(page, 'Header link');
  const popup = page.locator('.oox-hyperlink-popup');
  await expect(popup).toHaveCount(0);
  await expect(page.locator('.hf-inline-editor')).toHaveCount(0);
  const box = (await page.locator('.paged-editor__page').first().boundingBox())!;
  await page.mouse.dblclick(box.x + 200, box.y + 30);
  await expect(page.locator('.hf-inline-editor')).toContainText('Header');
  await clickText(page, 'Header link');
  await expect(popup).toBeVisible();
  await expect(popup).toContainText('https://example.com/header');
  const result = await saved(page, info);
  expect(await result.zip.file('word/header.xml')!.async('string')).toContain('<w:hyperlink');
  expect(await result.zip.file('word/_rels/header.xml.rels')!.async('string')).toContain('https://example.com/header');
});

for (const [link, target] of [['Jump to target', 'Destination'], ['Jump to cell', 'Cell destination']]) {
  test(`worker bookmark link ${link} moves the viewport and caret without a popup`, async ({ page }, info) => {
    const input = await open(page);
    await clickText(page, link!);
    await expect(page.locator('.oox-hyperlink-popup')).toHaveCount(0);
    const targetRun = page.locator('.canvas-page-mirror .layout-paragraph').filter({ hasText: target! }).first().locator('.layout-run').first();
    await expect(targetRun).toBeInViewport();
    await expect.poll(() => page.locator('.docx-editor__scroll-container').evaluate(node => node.scrollTop)).toBeGreaterThan(500);
    await input.pressSequentially('Arrived');
    const result = await saved(page, info);
    expect(result.xml).toContain('Arrived');
    expect(result.links).toContainEqual({ text: link, href: link === 'Jump to target' ? '#target' : '#cellTarget' });
    const paragraph = await page.evaluate(({ xml, target }) => {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      return Array.from(doc.getElementsByTagNameNS('*', 'p')).find(p => p.textContent?.replace('Arrived', '') === target)?.textContent;
    }, { xml: result.xml, target });
    expect(paragraph).toContain('Arrived');
  });
}
