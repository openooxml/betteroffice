import { expect, test, type Locator, type Page } from 'playwright/test';
import JSZip from 'jszip';

const checkbox = (id: number, alias: string) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:alias w:val="${alias}"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`;

async function open(page: Page, count: number, near = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const pages = Array.from({ length: count - 1 }, (_, index) => near && index === 0
    ? `<w:p><w:r><w:t xml:space="preserve">Page 1 </w:t></w:r>${checkbox(8, 'Near flag')}<w:r><w:br w:type="page"/></w:r></w:p>`
    : `<w:p><w:r><w:t>Page ${index + 1}</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>`).join('');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
    `${pages}<w:p><w:r><w:t xml:space="preserve">Tick here </w:t></w:r>${checkbox(7, 'Far flag')}</w:p>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'control-pages.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  if (count <= 12) {
    await expect(page.locator('.paged-editor__page')).toHaveCount(count, { timeout: 60_000 });
    return;
  }
  const h = (await page.locator('.paged-editor__page').first().boundingBox())!.height;
  await expect.poll(() => page.locator('[data-docx-caret-host]').evaluate(n => n.getBoundingClientRect().height), { timeout: 60_000 }).toBeGreaterThan((count - 1) * h);
}

const handleFor = (page: Page, name: string) => page.locator('[data-worker-control]').getByRole('checkbox', { name, exact: true });

async function expectOverControl(flag: Locator, sheet: Locator) {
  await expect(flag).toBeVisible({ timeout: 10_000 });
  await expect(async () => {
    const handle = (await flag.boundingBox())!;
    const sheetBox = (await sheet.boundingBox())!;
    expect(handle.y).toBeGreaterThanOrEqual(sheetBox.y);
    expect(handle.y + handle.height).toBeLessThanOrEqual(sheetBox.y + sheetBox.height);
    const line = (await sheet.locator('.canvas-page-mirror .layout-run-text').filter({ hasText: '☐' }).first().boundingBox())!;
    expect(handle.x < line.x + line.width && line.x < handle.x + handle.width).toBe(true);
    expect(handle.y < line.y + line.height && line.y < handle.y + handle.height).toBe(true);
  }).toPass({ timeout: 10_000 });
}

test('a checkbox on the last page of a short document gets its handle over the control', async ({ page }) => {
  const count = 6;
  await open(page, count);
  const sheet = page.locator('.paged-editor__page').nth(count - 1);
  await sheet.evaluate(node => node.scrollIntoView({ block: 'center' }));
  const flag = handleFor(page, 'Far flag');
  await expect(flag).toBeVisible({ timeout: 10_000 });
  const handle = (await flag.boundingBox())!;
  const sheetBox = (await sheet.boundingBox())!;
  expect(handle.y).toBeGreaterThanOrEqual(sheetBox.y);
  expect(handle.y + handle.height).toBeLessThanOrEqual(sheetBox.y + sheetBox.height);
  const line = (await sheet.locator('.canvas-page-mirror .layout-run-text').filter({ hasText: '☐' }).first().boundingBox())!;
  expect(handle.x < line.x + line.width && line.x < handle.x + handle.width).toBe(true);
  expect(handle.y < line.y + line.height && line.y < handle.y + handle.height).toBe(true);
});

test('a checkbox on a page built by scrolling a long document gets its handle over the control', async ({ page }) => {
  test.setTimeout(120_000);
  const count = 24;
  await open(page, count, true);
  const first = page.locator('.paged-editor__page[data-page-index="0"]');
  const far = page.locator(`.paged-editor__page[data-page-index="${count - 1}"]`);
  const near = handleFor(page, 'Near flag');
  await expectOverControl(near, first);
  await page.evaluate(() => document.fonts.ready);
  await expect.poll(() => near.evaluate(element => new Promise<boolean>(resolve => {
    const before = element.getBoundingClientRect();
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const after = element.getBoundingClientRect();
      resolve(before.x === after.x && before.y === after.y && before.width === after.width && before.height === after.height);
    }));
  }))).toBe(true);
  await expect(far, `precondition: page ${count} must not be mounted before scrolling, or this test cannot exercise the viewport build`).toHaveCount(0);
  await page.locator('[data-docx-caret-host]').evaluate(n => n.scrollIntoView({ block: 'end' }));
  await expect(far.locator('canvas')).toHaveCount(1, { timeout: 30_000 });
  await expectOverControl(handleFor(page, 'Far flag'), far);
});
