import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { pinnedEvaluate } from './pinned-evaluate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function open(page: Page): Promise<Locator> {
  const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const body = `<w:p>${run('Bold', '<w:b/>')}${run(' source')}</w:p><w:p>${run('Target')}</w:p>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'paste-options.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  const input = page.locator('.paged-editor__input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

async function copyBoldWord(page: Page, input: Locator) {
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('stale'));
  await input.press('ControlOrMeta+Home');
  for (let index = 0; index < 4; index += 1) await input.press('Shift+ArrowRight');
  await input.press('ControlOrMeta+c');
  await expect.poll(() => pinnedEvaluate(page, async () => {
    try {
      const items = await navigator.clipboard.read();
      const read = async (type: string) => {
        const item = items.find(entry => entry.types.includes(type));
        return item ? (await item.getType(type)).text() : '';
      };
      return { html: await read('text/html'), text: await read('text/plain') };
    } catch (error) {
      if (error instanceof DOMException && error.name === 'InvalidStateError') return { html: '', text: '' };
      throw error;
    }
  })).toMatchObject({ html: expect.stringContaining('data-betteroffice-docx'), text: 'Bold' });
}

async function savedParagraphs(page: Page, info: TestInfo, name: string): Promise<{ text: string; bold: string }[]> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const xml = await (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
  return page.evaluate(xml => {
    const on = (node: Element | undefined) => !!node && !['0', 'false', 'off'].includes(node.getAttributeNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val') ?? '');
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const runText = (run: Element) => Array.from(run.getElementsByTagNameNS('*', 't')).map(text => text.textContent ?? '').join('');
    return Array.from(document.getElementsByTagNameNS('*', 'p')).map(paragraph => {
      const runs = Array.from(paragraph.getElementsByTagNameNS('*', 'r'));
      return {
        text: runs.map(runText).join(''),
        bold: runs.filter(run => on(run.getElementsByTagNameNS('*', 'rPr')[0]?.getElementsByTagNameNS('*', 'b')[0])).map(runText).join(''),
      };
    }).filter(paragraph => paragraph.text);
  }, xml);
}

test('Keep Text Only after pasting a bold word copied in the worker editor keeps the text and drops the bold', async ({ page }, info) => {
  const input = await open(page);
  await copyBoldWord(page, input);
  await input.press('ControlOrMeta+End');
  await input.press('ControlOrMeta+v');
  const control = page.getByRole('button', { name: 'Paste Options (Ctrl)' });
  await expect(control).toBeVisible({ timeout: 30_000 });
  await control.click();
  await page.getByRole('menuitemradio', { name: 'Keep Text Only' }).click();
  await expect.poll(() => savedParagraphs(page, info, 'text-only'), { timeout: 30_000 }).toEqual([
    { text: 'Bold source', bold: 'Bold' },
    { text: 'TargetBold', bold: '' },
  ]);
});

test('typing after a paste removes the Paste Options control', async ({ page }, info) => {
  const input = await open(page);
  await copyBoldWord(page, input);
  await input.press('ControlOrMeta+End');
  await input.press('ControlOrMeta+v');
  const control = page.getByRole('button', { name: 'Paste Options (Ctrl)' });
  await expect(control).toBeVisible({ timeout: 30_000 });
  await page.keyboard.type('x');
  await expect(control).toBeHidden();
  await expect.poll(async () => (await savedParagraphs(page, info, 'typed')).map(paragraph => paragraph.text), { timeout: 30_000 })
    .toEqual(['Bold source', 'TargetBoldx']);
});
