import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function documentBytes(plain = false): Promise<Buffer> {
  const cell = (text: string, bold = false) => `<w:tc><w:p><w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid>` +
    `<w:tr>${cell('RIGHT', true)}${cell('INDEX')}</w:tr></w:tbl>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const body = plain ? '<w:p><w:r><w:t>Immediate clipboard text</w:t></w:r></w:p>'
    : `<w:p><w:r><w:t>Before</w:t></w:r></w:p>${table}<w:p/>`;
  zip.file('word/document.xml', `<w:document ${W}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, deny = true, plain = false): Promise<Locator> {
  if (deny) await page.addInitScript(() => {
    const denied = () => Promise.reject(new DOMException('Write permission denied.', 'NotAllowedError'));
    Object.defineProperty(Navigator.prototype, 'clipboard', { configurable: true,
      get: () => ({ write: denied, writeText: denied, read: denied, readText: denied }) });
  });
  await page.addInitScript(() => {
    const record = (event: ClipboardEvent) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text) (window as unknown as { nativeClipboardText: string }).nativeClipboardText = text;
    };
    document.addEventListener('copy', record);
    document.addEventListener('cut', record);
  });
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'worker-clipboard.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes(plain) });
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
  await input.press('ControlOrMeta+a');
  await input.press('ControlOrMeta+c');
  await expect.poll(() => page.evaluate(() => (window as unknown as { nativeClipboardText?: string }).nativeClipboardText)).toContain('RIGHT');
  const copied = await pasteIntoSink(page);
  expect(copied.text).toContain('Before');
  expect(copied.text).toContain('RIGHT');
  expect(copied.html).toContain('<table');
  expect(copied.html).toContain('font-weight:bold');
  await input.focus();
  await input.press('ControlOrMeta+a');
  await page.evaluate(() => { (window as unknown as { nativeClipboardText: string }).nativeClipboardText = ''; });
  await input.press('ControlOrMeta+x');
  await expect.poll(() => page.evaluate(() => (window as unknown as { nativeClipboardText?: string }).nativeClipboardText)).toContain('RIGHT');
  const cut = await pasteIntoSink(page);
  expect(cut.text).toBe(copied.text);
  await expect.poll(async () => (await savedXml(page, info, 'cut')).includes('RIGHT')).toBe(false);
});


for (const cut of [false, true]) test(`immediate native ${cut ? 'cut' : 'copy'} completes a cache miss with the async Clipboard API denied`, async ({ page }, info) => {
  await page.addInitScript(() => {
    let waiting = true;
    const held: (() => void)[] = [];
    const OriginalWorker = window.Worker;
    window.Worker = class extends OriginalWorker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        const reads = new Set<number>();
        const send = this.postMessage.bind(this);
        this.postMessage = ((message: { id?: number; type?: string; read?: { kind?: string } }, ...args: unknown[]) => {
          if (message.type === 'documentRead' && message.read?.kind === 'clipboardSlice' && message.id !== undefined) reads.add(message.id);
          Reflect.apply(send, this, [message, ...args]);
        }) as typeof this.postMessage;
        const replayed = new WeakSet<Event>();
        this.addEventListener('message', event => {
          if (!waiting || replayed.has(event) || !reads.has(event.data.id)) return;
          event.stopImmediatePropagation();
          held.push(() => { const replay = new MessageEvent('message', { data: event.data }); replayed.add(replay); this.dispatchEvent(replay); });
        });
      }
    };
    const release = () => { waiting = false; for (const send of held.splice(0)) queueMicrotask(send); };
    document.addEventListener('copy', release, { capture: true, once: true });
    document.addEventListener('cut', release, { capture: true, once: true });
  });
  const input = await open(page, true, true);
  await page.evaluate(() => { const sink = document.createElement('textarea'); sink.id = 'plain-sink'; document.body.append(sink); });
  await input.press('ControlOrMeta+a');
  await input.press(cut ? 'ControlOrMeta+x' : 'ControlOrMeta+c');
  await expect.poll(() => page.evaluate(() => (window as unknown as { nativeClipboardText?: string }).nativeClipboardText)).toBe('Immediate clipboard text');
  await expect.poll(async () => (await savedXml(page, info, 'immediate')).includes('Immediate clipboard text')).toBe(!cut);
  await page.locator('#plain-sink').focus();
  await page.keyboard.press('ControlOrMeta+v');
  await expect(page.locator('#plain-sink')).toHaveValue('Immediate clipboard text');
});

test('native copy retains table HTML with the async Clipboard API granted', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  const input = await open(page, false);
  await input.press('ControlOrMeta+a');
  await input.press('ControlOrMeta+c');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toContain('RIGHT');
  const copied = await pasteIntoSink(page);
  expect(copied.text).toContain('Before');
  expect(copied.text).toContain('RIGHT');
  expect(copied.html).toContain('<table');
});
