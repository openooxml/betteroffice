import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

async function open(page: Page, arm: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Paragraph spacing</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.setViewportSize({ width: 1600, height: 1000 });
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'paragraph-spacing.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  const canvas = page.locator('.canvas-page canvas, .paged-editor__page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60000 });
  await canvas.click({ position: { x: 150, y: 150 } });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
}

async function savedXml(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(name);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file('word/document.xml')!.async('string');
}

for (const arm of ['&peerless=0', '']) {
  test(`paragraph spacing menu saves before spacing and opens options (${arm ? 'legacy' : 'peerless'})`, async ({ page }, info) => {
    await open(page, arm);
    const trigger = page.getByRole('button', { name: 'Line spacing', exact: true });
    await trigger.focus();
    await page.keyboard.press('ArrowDown');
    await page.getByRole('menuitem', { name: 'Add Space Before Paragraph', exact: true }).focus();
    await page.keyboard.press('Enter');
    expect(await savedXml(page, info, 'before.docx')).toMatch(/<w:spacing\b[^>]*\bw:before="240"/);
    await trigger.click();
    await expect(page.getByRole('menuitem', { name: 'Remove Space Before Paragraph', exact: true })).toBeVisible();
    await page.getByRole('menuitem', { name: 'Line Spacing Options...', exact: true }).click();
    const dialog = page.getByRole('dialog', { name: 'Paragraph spacing', exact: true });
    await expect(dialog.getByLabel('Spacing Before', { exact: true })).toHaveValue('12');
    await dialog.getByLabel('Spacing Before', { exact: true }).fill('6');
    await dialog.getByLabel('Spacing After', { exact: true }).fill('18');
    await dialog.getByLabel('Line spacing', { exact: true }).selectOption('1.5');
    await dialog.getByLabel('Spacing After', { exact: true }).press('Enter');
    await expect(dialog).toBeHidden();
    const xml = await savedXml(page, info, 'dialog.docx');
    expect(xml).toMatch(/<w:spacing\b[^>]*\bw:before="120"/);
    expect(xml).toMatch(/<w:spacing\b[^>]*\bw:after="360"/);
    expect(xml).toMatch(/<w:spacing\b[^>]*\bw:line="360"/);
  });
}
