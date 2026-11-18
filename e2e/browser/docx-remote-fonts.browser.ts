import { test, expect } from 'playwright/test';
import JSZip from 'jszip';

const families = ['Calibri', 'Tahoma', 'Roboto', 'Century Gothic', 'MS Mincho'];

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    families.map((family) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:eastAsia="${family}"/></w:rPr><w:t>${family} sample text</w:t></w:r></w:p>`).join('') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('a DOCX with fonts that are not bundled makes no Google Fonts request by default', async ({ page }) => {
  const remote: string[] = [];
  page.on('request', (request) => {
    const host = new URL(request.url()).hostname;
    if (/(^|\.)(googleapis\.com|gstatic\.com|bunny\.net)$/.test(host)) remote.push(request.url());
  });
  await page.goto('http://127.0.0.1:4186/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'remote-fonts.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 60000 });
  await expect(page.locator('.paged-editor__page .canvas-page-mirror').first()).toContainText('Roboto sample text');
  await page.waitForTimeout(3000);
  expect(await page.evaluate(() => [...document.querySelectorAll('link[rel=stylesheet]')]
    .map((link) => (link as HTMLLinkElement).href).filter((href) => /fonts\.googleapis\.com/.test(href)))).toEqual([]);
  expect(remote).toEqual([]);
});
