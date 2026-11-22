import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { pinnedEvaluate } from './pinned-evaluate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function words(count: number): string[] {
  return Array.from({ length: count }, (_, index) => `w${String(index).padStart(4, '0')}${index + 1 < count ? ' ' : ''}`);
}

async function open(page: Page, count: number): Promise<Locator> {
  const runs = words(count).map((text, index) =>
    `<w:r><w:rPr>${index % 2 === 0 ? '<w:b/>' : '<w:i/>'}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body><w:p>${runs}</w:p><w:p/>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'many-runs.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 90_000 });
  const input = page.locator('.paged-editor__input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

async function savedParagraphs(page: Page, info: TestInfo, name: string): Promise<{ text: string; boldRuns: number }[]> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const xml = await (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
  return page.evaluate(xml => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    return Array.from(document.getElementsByTagNameNS('*', 'p')).map(paragraph => {
      const runs = Array.from(paragraph.getElementsByTagNameNS('*', 'r'));
      return {
        text: runs.map(run => Array.from(run.getElementsByTagNameNS('*', 't')).map(text => text.textContent ?? '').join('')).join(''),
        boldRuns: runs.filter(run => run.getElementsByTagNameNS('*', 'rPr')[0]?.getElementsByTagNameNS('*', 'b')[0]).length,
      };
    });
  }, xml);
}

for (const count of [1_000, 4_000]) {
  test(`copying and pasting a paragraph of ${count} alternating bold and italic runs finishes promptly`, async ({ page }, info) => {
    test.setTimeout(240_000);
    const input = await open(page, count);
    const text = words(count).join('');
    await pinnedEvaluate(page, () => navigator.clipboard.writeText('stale'));
    await input.press('ControlOrMeta+End');
    await input.press('ArrowLeft');
    await input.press('ControlOrMeta+Shift+Home');
    await input.press('ControlOrMeta+c');
    await expect.poll(() => pinnedEvaluate(page, async () => {
      try {
        return (await navigator.clipboard.readText()).length;
      } catch {
        return 0;
      }
    }), { timeout: 30_000 }).toBe(text.length);
    await input.press('ControlOrMeta+End');
    const start = Date.now();
    await input.press('ControlOrMeta+v');
    await expect.poll(async () => (await savedParagraphs(page, info, `pasted-${count}`))[1]?.text ?? '',
      { timeout: 120_000, intervals: [250] }).toBe(text);
    const elapsed = Date.now() - start;
    console.log(`PASTE_MANY_RUNS_MS runs=${count} ms=${elapsed}`);
    expect(elapsed).toBeLessThan(5_000);
    expect((await savedParagraphs(page, info, `saved-${count}`))[1]!.boldRuns).toBe(0);
  });
}
