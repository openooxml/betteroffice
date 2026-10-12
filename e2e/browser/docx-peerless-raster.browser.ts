import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:r><w:t>Device pixel text</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second page</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function backingSizes(page: Page, dpr: number) {
  return page.locator('[data-peerless-docx] [data-page-id] canvas').evaluateAll((canvases, ratio) => {
    const visible = canvases.filter((canvas) => {
      const rect = canvas.getBoundingClientRect();
      return rect.bottom > 0 && rect.top < innerHeight && rect.right > 0 && rect.left < innerWidth;
    });
    return visible.length > 0 && visible.every((element) => {
      const canvas = element as HTMLCanvasElement;
      const rect = canvas.getBoundingClientRect();
      return Math.abs(canvas.width - Math.round(rect.width * ratio)) <= 1 &&
        Math.abs(canvas.height - Math.round(rect.height * ratio)) <= 1;
    });
  }, dpr);
}

for (const dpr of [1, 2]) {
  test(`worker pages rasterize at device pixels at DPR ${dpr}`, async ({ browser }) => {
    const context = await browser.newContext({ deviceScaleFactor: dpr, viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    try {
      await page.goto('/?format=docx&peerless=1');
      await page.locator('input[type=file]').last().setInputFiles({ name: 'raster.docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
      await expect.poll(() => page.evaluate(() => window.__docxPeerlessHarness?.capabilities), {
        message: 'Build with DOCX_PEERLESS_HARNESS=1',
      }).toMatchObject({ presentation: true, mode: true });
      await expect(page.locator('[data-peerless-docx] [data-page-id] canvas').first()).toBeVisible();
      await expect.poll(() => backingSizes(page, dpr)).toBe(true);
      await page.getByRole('combobox', { name: /Zoom/ }).click();
      await page.getByRole('option', { name: '150%', exact: true }).click();
      await expect(page.getByRole('combobox', { name: /Zoom/ })).toHaveText('150%');
      await expect.poll(() => backingSizes(page, dpr)).toBe(true);
    } finally { await context.close(); }
  });
}
