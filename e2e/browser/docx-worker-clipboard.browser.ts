import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import { pinnedEvaluate } from './pinned-evaluate';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function documentBytes(): Promise<Buffer> {
  const cell = (text: string, bold = false) => `<w:tc><w:p><w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid>` +
    `<w:tr>${cell('RIGHT', true)}${cell('INDEX')}</w:tr></w:tbl>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>${table}<w:p/>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page): Promise<Locator> {
  await page.addInitScript(() => {
    const denied = () => Promise.reject(new DOMException('Write permission denied.', 'NotAllowedError'));
    Object.defineProperty(Navigator.prototype, 'clipboard', { configurable: true,
      get: () => ({ write: denied, writeText: denied, read: denied, readText: denied }) });
  });
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-clipboard.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  await page.evaluate(() => {
    const sink = document.createElement('div');
    sink.id = 'paste-sink';
    sink.contentEditable = 'true';
    document.body.append(sink);
    sink.addEventListener('paste', (event) => {
      event.preventDefault();
      (window as unknown as { pasted: { text: string; html: string } }).pasted = {
        text: event.clipboardData?.getData('text/plain') ?? '', html: event.clipboardData?.getData('text/html') ?? '' };
    });
  });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

async function selectAllSettled(page: Page, input: Locator) {
  await input.press('ControlOrMeta+a');
  await expect.poll(() => pinnedEvaluate(page, () => {
    const data = new DataTransfer();
    document.querySelector('.paged-editor__peerless-input')!.dispatchEvent(new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }));
    return data.getData('text/plain');
  })).toContain('RIGHT');
}

async function pasteIntoSink(page: Page): Promise<{ text: string; html: string }> {
  await page.evaluate(() => { (window as unknown as { pasted?: unknown }).pasted = undefined; });
  await page.locator('#paste-sink').focus();
  await page.keyboard.press('ControlOrMeta+v');
  return expect.poll(() => page.evaluate(() => (window as unknown as { pasted?: { text: string; html: string } }).pasted ?? null))
    .not.toBeNull().then(() => page.evaluate(() => (window as unknown as { pasted: { text: string; html: string } }).pasted));
}

async function savedXml(page: Page, info: TestInfo, name: string): Promise<string> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  return (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
}

test('copy and cut fill the native clipboard when the async Clipboard API is denied', async ({ page }, info) => {
  const input = await open(page);
  await selectAllSettled(page, input);
  await input.press('ControlOrMeta+c');
  const copied = await pasteIntoSink(page);
  expect(copied.text).toContain('Before');
  expect(copied.text).toContain('RIGHT');
  expect(copied.html).toContain('<table');
  expect(copied.html).toContain('font-weight:bold');
  await input.focus();
  await selectAllSettled(page, input);
  await input.press('ControlOrMeta+x');
  const cut = await pasteIntoSink(page);
  expect(cut.text).toBe(copied.text);
  await expect.poll(async () => (await savedXml(page, info, 'cut')).includes('RIGHT')).toBe(false);
});
