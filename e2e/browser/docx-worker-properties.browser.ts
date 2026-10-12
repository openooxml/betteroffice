import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page, twoSections = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const properties = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>';
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001">' +
    (twoSections ? `<w:pPr>${properties}</w:pPr>` : '') + '<w:r><w:t>Seed</w:t></w:r></w:p>' +
    (twoSections ? '<w:p><w:r><w:t>Second</w:t></w:r></w:p>' : '') + properties + '</w:body></w:document>');
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-properties.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  return input;
}

test('page setup on page one updates the first section without source IDs', async ({ page }, info) => {
  await open(page, true);
  await page.getByRole('menubar').getByRole('menuitem', { name: 'File', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Page setup', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Page Setup', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.locator('input[type=number]').nth(2).fill('2');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const xml = await savedXml(page, info);
  const sections = xml.match(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)!;
  expect(sections).toHaveLength(2);
  expect(sections[0]).toContain('w:left="2880"');
  expect(sections[1]).toContain('w:left="1440"');
});

async function savedXml(page: Page, info: TestInfo, part = 'word/document.xml') {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('properties.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file(part)!.async('string');
}

test('worker page setup preserves admitted text and saves authored margins', async ({ page }, info) => {
  const input = await open(page);
  await input.pressSequentially('Before ');
  await page.getByRole('menubar').getByRole('menuitem', { name: 'File', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Page setup', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Page Setup', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.locator('input[type=number]').nth(2).fill('1.5');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const xml = await savedXml(page, info);
  expect(xml).toContain('Before Seed');
  expect(xml).toContain('w:left="2160"');
});

test('worker watermark dialog saves automatic size without dropping body edits', async ({ page }, info) => {
  const input = await open(page);
  await input.pressSequentially('Watermark ');
  await page.getByRole('menubar').getByRole('menuitem', { name: 'Insert', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Watermark', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Watermark', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.getByText('Text watermark', { exact: true }).click();
  await dialog.getByRole('textbox').first().fill('Draft');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('watermark.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  expect(await zip.file('word/document.xml')!.async('string')).toContain('Watermark Seed');
  const headers = await Promise.all(zip.file(/^word\/header.*\.xml$/).map(file => file.async('string')));
  expect(headers.join('')).toContain('Draft');
});

test('worker hyperlink dialog edits the captured text range and preserves its display text', async ({ page }, info) => {
  const input = await open(page);
  await input.press('ControlOrMeta+End');
  await input.press('Shift+Home');
  const inline = page.getByRole('button', { name: 'Insert link', exact: true });
  if (await inline.isVisible()) await inline.click();
  else {
    await page.getByRole('button', { name: 'More actions', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Insert link', exact: true }).click();
  }
  const dialog = page.getByRole('dialog', { name: 'Insert Hyperlink', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByLabel('Display Text', { exact: true })).toHaveAttribute('readonly');
  await dialog.getByLabel('URL', { exact: true }).fill('https://example.com/link');
  await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('hyperlink.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  expect(await zip.file('word/document.xml')!.async('string')).toContain('Seed');
  expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain('https://example.com/link');
});
