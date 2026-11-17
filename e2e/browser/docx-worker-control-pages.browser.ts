import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const PAGES = 6;

async function open(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const checkbox = '<w:sdt><w:sdtPr><w:id w:val="7"/><w:alias w:val="Far flag"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>';
  const pages = Array.from({ length: PAGES - 1 }, (_, index) => `<w:p><w:r><w:t>Page ${index + 1}</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>`).join('');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
    `${pages}<w:p><w:r><w:t xml:space="preserve">Tick here </w:t></w:r>${checkbox}</w:p>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'control-pages.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.locator('.paged-editor__page')).toHaveCount(PAGES, { timeout: 60_000 });
}

test('a checkbox on a page built by scrolling gets its handle over the control', async ({ page }) => {
  await open(page);
  const sheet = page.locator('.paged-editor__page').nth(PAGES - 1);
  await sheet.evaluate(node => node.scrollIntoView({ block: 'center' }));
  const flag = page.locator('[data-worker-control]').getByRole('checkbox', { name: 'Far flag', exact: true });
  await expect(flag).toBeVisible({ timeout: 10_000 });
  const handle = (await flag.boundingBox())!;
  const sheetBox = (await sheet.boundingBox())!;
  expect(handle.y).toBeGreaterThanOrEqual(sheetBox.y);
  expect(handle.y + handle.height).toBeLessThanOrEqual(sheetBox.y + sheetBox.height);
  const line = (await page.locator('.canvas-page-mirror .layout-paragraph').filter({ hasText: 'Tick here' }).first().boundingBox())!;
  expect(handle.x < line.x + line.width && line.x < handle.x + handle.width).toBe(true);
  expect(handle.y < line.y + line.height && line.y < handle.y + handle.height).toBe(true);
});
