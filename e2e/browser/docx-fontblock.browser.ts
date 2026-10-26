import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const finalText = 'FINAL PARAGRAPH COMPLETE';

async function documentBytes(explicitBreaks: boolean) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    Array.from({ length: explicitBreaks ? 8 : 100 }, (_, i) => {
      const family = i % 2 ? 'Calibri' : 'Arial';
      return `<w:p><w:pPr>${explicitBreaks && i > 0 ? '<w:pageBreakBefore/>' : ''}<w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="22"/></w:rPr><w:t>` +
        `Paragraph ${i + 1}. ` + 'The quiet river runs beside the green field. '.repeat(24) + '</w:t></w:r></w:p>';
    }).join('') + `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>${finalText}</w:t></w:r></w:p>` + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function settledPages(page: Page) {
  let previous = 0;
  let stableSince = Date.now();
  let count = 0;
  await expect.poll(async () => {
    const indicators = await page.getByRole('status').allTextContents();
    const total = indicators.map(text => /^\s*\d+ of (\d+)\s*$/.exec(text)).find(Boolean);
    count = total ? Number(total[1]) : 0;
    if (count !== previous || count === 0) { previous = count; stableSince = Date.now(); }
    return count > 0 && Date.now() - stableSince >= 3000;
  }, { timeout: 60000, intervals: [250] }).toBe(true);
  return count;
}

for (const explicitBreaks of [true, false]) {
  test(`default DocxEditor retains ${explicitBreaks ? 'explicit pages' : 'wrapped content'} when font fetches fail`, async ({ browser }) => {
    const buffer = await documentBytes(explicitBreaks);
    const counts: number[] = [];
    for (const blocked of [false, true]) {
      const context = await browser.newContext({ serviceWorkers: 'block', locale: 'en-US' });
      const page = await context.newPage();
      const warnings: string[] = [];
      const errors: string[] = [];
      let aborted = 0;
      page.on('console', message => {
        if (message.type() === 'warning' && /fontRegistry/.test(message.text())) warnings.push(message.text());
      });
      page.on('pageerror', error => errors.push(error.message));
      if (blocked) {
        await page.route('https://cdn.jsdelivr.net/**', route => { aborted++; return route.abort(); });
        await page.route(/\.(?:ttf|otf|woff2?)(?:\?.*)?$/, route => { aborted++; return route.abort(); });
      }
      try {
        await page.goto('http://127.0.0.1:4186/?format=docx');
        await page.locator('input[type=file]').last().setInputFiles({ name: 'fontblock.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
        await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
        await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 60000 });
        const count = await settledPages(page);
        expect(count).toBeGreaterThanOrEqual(6);
        if (explicitBreaks) expect(count).toBe(8);
        counts.push(count);
        await page.locator('.docx-editor__scroll-container').evaluate(element => { element.scrollTop = element.scrollHeight; });
        const lastPage = page.locator(`.paged-editor__page[data-page-index="${count - 1}"]`);
        await expect(lastPage.locator('canvas')).toBeVisible({ timeout: 30000 });
        await expect.poll(() => lastPage.locator('canvas').evaluate((canvas: HTMLCanvasElement) => {
          const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
          for (let i = 0; i < data.length; i += 4) {
            if (data[i + 3]! > 0 && Math.min(data[i]!, data[i + 1]!, data[i + 2]!) < 128) return true;
          }
          return false;
        }), { timeout: 30000 }).toBe(true);
        await expect(lastPage.locator('.canvas-page-mirror')).toContainText(finalText);
        await expect(page.locator('.paged-editor__input')).toBeEditable();
        await expect(page.locator('.notice')).toHaveCount(0);
        expect(errors).toEqual([]);
        if (blocked) {
          expect(aborted).toBeGreaterThan(0);
          // One registry warning; font fallback is never reported through onError, so hosts
          // that treat onError as fatal keep laying out.
          expect(warnings).toHaveLength(1);
          expect(warnings[0]!.startsWith('[fontRegistry]')).toBe(true);
          expect(warnings[0]).toContain('falling back');
        } else {
          expect(warnings).toEqual([]);
        }
      } finally { await context.close(); }
    }
    if (explicitBreaks) expect(counts[1]).toBe(counts[0]);
    else expect(counts[1]!).toBeGreaterThanOrEqual(counts[0]!);
  });
}
