import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const overlay = '[data-peerless-docx] div[aria-hidden="true"][style*="rgba(59, 130, 246"]';
const caret = '[data-peerless-docx] div[aria-hidden="true"][style*="currentcolor" i]';

async function open(page: Page) {
  const run = (text: string, rPr = '') => `<w:r><w:rPr>${rPr}<w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r>`;
  const body = '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:u w:val="single"/><w:sz w:val="24"/></w:rPr></w:pPr>' +
    `${run('Document Title', '<w:b/><w:u w:val="single"/>')}</w:p>` +
    `<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr>${run('First paragraph.')}</w:p>` +
    `<w:p>${run('Second paragraph.', '<w:b/>')}</w:p><w:p>${run('Third paragraph.')}</w:p>`;
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.setViewportSize({ width: 2560, height: 1200 });
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'paste-selection.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  await expect(page.locator(caret)).toHaveCount(1);
}

async function box(page: Page, selector: string) {
  const bounds = await page.locator(selector).first().boundingBox();
  if (!bounds) throw new Error(`${selector} has no box`);
  return bounds;
}

async function highlightedLines(page: Page) {
  const rects = await page.locator(overlay).evaluateAll(nodes => nodes.map(node => {
    const rect = node.getBoundingClientRect();
    return { top: rect.top, bottom: rect.bottom };
  }).filter(rect => rect.bottom > rect.top).sort((a, b) => a.top - b.top));
  const lines: { top: number; bottom: number }[] = [];
  for (const rect of rects) {
    const last = lines.at(-1);
    if (last && rect.top < last.bottom - 1) last.bottom = Math.max(last.bottom, rect.bottom);
    else lines.push({ ...rect });
  }
  return lines;
}

async function drag(page: Page, from: { x: number; y: number }, to: { x: number; y: number }) {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 8 });
  await page.mouse.up();
}

async function boldParagraphs(page: Page, info: TestInfo) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('saved.docx');
  await (await downloading).saveAs(file);
  const xml = await (await JSZip.loadAsync(await readFile(file))).file('word/document.xml')!.async('string');
  return page.evaluate(xml => {
    const on = (node: Element | undefined) => !!node && !['0', 'false', 'off'].includes(node.getAttributeNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'val') ?? '');
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    return Array.from(document.getElementsByTagNameNS('*', 'p')).map(paragraph => {
      const runs = Array.from(paragraph.getElementsByTagNameNS('*', 'r')).filter(run =>
        Array.from(run.getElementsByTagNameNS('*', 't')).some(text => text.textContent));
      return {
        text: runs.map(run => Array.from(run.getElementsByTagNameNS('*', 't')).map(text => text.textContent).join('')).join(''),
        bold: runs.length > 0 && runs.every(run => on(run.getElementsByTagNameNS('*', 'rPr')[0]?.getElementsByTagNameNS('*', 'b')[0])),
      };
    }).filter(paragraph => paragraph.text);
  }, xml);
}

test('bold after a plain-text paste follows the dragged selection in the toolbar and the shortcut', async ({ page }, info) => {
  await open(page);
  const sheet = await box(page, '.paged-editor__page');
  const scale = sheet.width / 816;
  const left = sheet.x + 96 * scale + 2;
  const right = sheet.x + (816 - 96) * scale - 2;
  const start = await box(page, caret);
  await drag(page, { x: start.x + 1, y: start.y + start.height / 2 }, { x: right, y: start.y + start.height * 12 });
  await expect.poll(async () => (await highlightedLines(page)).length).toBe(4);
  const lines = await highlightedLines(page);
  const center = (index: number) => (lines[index]!.top + lines[index]!.bottom) / 2;

  const input = page.locator('.paged-editor__peerless-input');
  expect(await input.evaluate(element => element === document.activeElement)).toBe(true);
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', 'Line one.\nLine two.\nLine three.\nLast line.');
    document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await expect(page.locator(overlay)).toHaveCount(0);
  await expect.poll(async () => {
    const at = await page.locator(caret).first().boundingBox();
    return !!at && Math.abs(at.y + at.height / 2 - center(3)) < (lines[3]!.bottom - lines[3]!.top) / 2;
  }).toBe(true);
  const bold = page.getByTestId('formatting-bar').getByRole('button', { name: 'Bold', exact: true });
  await expect.soft(bold).toHaveAttribute('aria-pressed', 'true');

  await drag(page, { x: left, y: center(0) }, { x: right, y: center(2) });
  await expect.poll(async () => (await highlightedLines(page)).length).toBe(3);
  const selected = await highlightedLines(page);
  expect(Math.abs((selected[2]!.top + selected[2]!.bottom) / 2 - center(2))).toBeLessThan(2);

  await expect.soft(bold).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('ControlOrMeta+b');
  await expect.soft(bold).toHaveAttribute('aria-pressed', 'false');
  expect(await boldParagraphs(page, info)).toEqual([
    { text: 'Line one.', bold: false },
    { text: 'Line two.', bold: false },
    { text: 'Line three.', bold: false },
    { text: 'Last line.', bold: true },
  ]);
});

test('Enter at the end of bold text keeps bold for the new paragraph', async ({ page }, info) => {
  await open(page);
  const input = page.locator('.paged-editor__peerless-input');
  await input.press('End');
  await input.press('Enter');
  await input.pressSequentially('Next');
  const bold = page.getByTestId('formatting-bar').getByRole('button', { name: 'Bold', exact: true });
  await expect.soft(bold).toHaveAttribute('aria-pressed', 'true');
  expect(await boldParagraphs(page, info)).toEqual([
    { text: 'Document Title', bold: true },
    { text: 'Next', bold: true },
    { text: 'First paragraph.', bold: false },
    { text: 'Second paragraph.', bold: true },
    { text: 'Third paragraph.', bold: false },
  ]);
});

test('a paste and Enter take the formatting of the run at the caret, not the paragraph mark', async ({ page }, info) => {
  await open(page);
  const input = page.locator('.paged-editor__peerless-input');
  await input.press('ArrowDown');
  await input.press('Home');
  await input.press('Shift+End');
  await expect(page.locator(overlay)).not.toHaveCount(0);
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', 'One.\nTwo.');
    document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await expect(page.locator(overlay)).toHaveCount(0);
  await input.press('ArrowDown');
  await input.press('End');
  await input.press('Enter');
  await input.pressSequentially('Next');
  const bold = page.getByTestId('formatting-bar').getByRole('button', { name: 'Bold', exact: true });
  await expect.soft(bold).toHaveAttribute('aria-pressed', 'true');
  expect(await boldParagraphs(page, info)).toEqual([
    { text: 'Document Title', bold: true },
    { text: 'One.', bold: false },
    { text: 'Two.', bold: false },
    { text: 'Second paragraph.', bold: true },
    { text: 'Next', bold: true },
    { text: 'Third paragraph.', bold: false },
  ]);
});
