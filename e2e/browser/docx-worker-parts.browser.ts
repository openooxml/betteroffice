import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page, notes = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  if (notes) {
    zip.file('word/document.xml', (await zip.file('word/document.xml')!.async('string')).replace('Seed</w:t></w:r>', 'Seed</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r>'));
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>');
    zip.file('word/footnotes.xml', '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>First note</w:t></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Second note</w:t></w:r></w:p></w:footnote></w:footnotes>');
  }
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-properties.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function savedXml(page: Page, info: TestInfo, part = 'word/document.xml') {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('properties.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file(part)!.async('string');
}

test('worker empty headers open once, save, close and remove through part chrome', async ({ page }, info) => {
  await open(page);
  const surface = page.locator('.paged-editor__page').first();
  const box = (await surface.boundingBox())!;
  await page.mouse.dblclick(box.x + 200, box.y + 30);
  const chrome = page.locator('.hf-inline-editor');
  await expect(chrome).toContainText('Header');
  await page.keyboard.type('New header');
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('header.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const headers = await Promise.all(zip.file(/^word\/header.*\.xml$/).map(file => file.async('string')));
  expect(headers.join('')).toContain('New header');
  await chrome.getByRole('button', { name: /Options/ }).click();
  await chrome.getByRole('button', { name: 'Close header editing' }).click();
  await expect(chrome).toHaveCount(0);
  await page.mouse.dblclick(box.x + 200, box.y + 30);
  await expect(chrome).toContainText('Header');
  await chrome.getByRole('button', { name: /Options/ }).click();
  await chrome.getByRole('button', { name: 'Remove header' }).click();
  await expect(chrome).toHaveCount(0);
  expect(await savedXml(page, info)).not.toContain('w:headerReference');
});

test('worker note clicks edit the second native story in a shared note area', async ({ page }, info) => {
  await open(page, true);
  const note = page.locator('.layout-note[data-note-id="2"] .layout-run[data-doc-start="5"]').first();
  await expect(note).toBeAttached();
  await note.evaluate(node => node.scrollIntoView({ block: 'center' }));
  const box = (await note.boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await expect(page.getByRole('button', { name: 'Close note', exact: true })).toBeVisible();
  await page.keyboard.press('ControlOrMeta+Home');
  await page.keyboard.type('Edited ');
  const xml = await savedXml(page, info, 'word/footnotes.xml');
  const second = xml.split('<w:footnote w:id="2"')[1]!;
  expect(second).toContain('Edited');
  expect(xml.split('<w:footnote w:id="2"')[0]).not.toContain('Edited');
  await page.getByRole('button', { name: 'Close note', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Close note', exact: true })).toHaveCount(0);
});
