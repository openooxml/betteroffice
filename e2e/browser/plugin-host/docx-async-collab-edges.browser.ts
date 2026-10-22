import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';
import type {} from './docx-peerless-collab-probe';
import { renderedBox } from '../rendered-box';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const editor = (page: Page, index: number) => page.locator(`[data-editor="${index}"]`);
const input = (page: Page, index: number) => editor(page, index).locator('textarea.paged-editor__peerless-input');
const text = (page: Page, index: number) => page.evaluate(index => window.__docxPeerlessCollab.text(index), index);
const paragraph = (value: string, newPage = false) => `<w:p>${newPage ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}<w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;

async function fixture(body: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${wordNamespace}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function load(page: Page, bytes: Buffer) {
  await page.evaluate(bytes => window.__docxPeerlessCollab.load(bytes), Array.from(bytes));
  await expect.poll(() => page.evaluate(() => [0, 1].every(index => window.__docxPeerlessCollab.synced(index)))).toBe(true);
  for (const index of [0, 1]) {
    await expect(editor(page, index).locator('canvas').first()).toBeVisible({ timeout: 60_000 });
    await expect(input(page, index)).toBeEditable();
  }
}

async function open(page: Page, body = paragraph('abcdef')) {
  await page.goto('/docx-peerless-collab.html');
  await page.waitForFunction(() => '__docxPeerlessCollab' in window);
  await page.evaluate(() => window.__docxPeerlessCollab.ready);
  await load(page, await fixture(body));
}

async function at(page: Page, index: number, offset: number) {
  const target = input(page, index);
  await target.focus();
  await target.press('ControlOrMeta+Home');
  for (let step = 0; step < offset; step++) await target.press('ArrowRight');
  return target;
}

async function vectorsMatch(page: Page) {
  await expect.poll(() => page.evaluate(async () => {
    const vectors = await Promise.all([0, 1].map(index => window.__docxPeerlessCollab.stateVector(index)));
    return JSON.stringify(vectors[0]) === JSON.stringify(vectors[1]);
  })).toBe(true);
}

async function converged(page: Page, paragraphs: string[], paint = paragraphs) {
  for (const index of [0, 1]) {
    await expect.poll(() => text(page, index)).toBe(paragraphs.join('\n'));
    for (const value of paint) {
      await expect.poll(() => page.evaluate(index => window.__docxPeerlessCollab.painted(index), index)).toContain(value);
    }
  }
  await vectorsMatch(page);
}

async function hold(page: Page) {
  await page.evaluate(() => {
    window.__docxPeerlessCollab.hold(0);
    window.__docxPeerlessCollab.hold(1);
  });
}

async function release(page: Page, reverse = false) {
  for (const index of [0, 1]) {
    await expect.poll(() => page.evaluate(index => window.__docxPeerlessCollab.heldUpdates(index), index)).toBeGreaterThan(0);
  }
  await page.evaluate(reverse => {
    for (const index of reverse ? [1, 0] : [0, 1]) window.__docxPeerlessCollab.release(index);
  }, reverse);
}

async function save(page: Page, index: number) {
  const downloading = page.waitForEvent('download');
  await input(page, index).press('ControlOrMeta+s');
  const stream = await (await downloading).createReadStream();
  expect(stream).not.toBeNull();
  const chunks: Buffer[] = [];
  for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
  const bytes = Buffer.concat(chunks);
  const zip = await JSZip.loadAsync(bytes);
  const part = zip.file('word/document.xml');
  expect(part).not.toBeNull();
  const xml = await part!.async('string');
  const content = await page.evaluate(({ xml, ns }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved document.xml');
    const elements = (root: Document | Element, name: string) => Array.from(root.getElementsByTagNameNS(ns, name));
    const value = (root: Element) => elements(root, 't').map(node => node.textContent ?? '').join('');
    return {
      paragraphs: elements(document, 'p').map(value),
      cells: elements(document, 'tc').map(value),
      bold: elements(document, 'r').filter(run => {
        const mark = elements(run, 'rPr')[0]?.getElementsByTagNameNS(ns, 'b')[0];
        return mark && !['0', 'false', 'off'].includes(mark.getAttributeNS(ns, 'val') ?? '1');
      }).map(value).join(''),
    };
  }, { xml, ns: wordNamespace });
  return { bytes, xml, ...content };
}

async function saved(page: Page, paragraphs: string[], cells?: string[]) {
  const left = await save(page, 0);
  const right = await save(page, 1);
  for (const result of [left, right]) {
    expect(result.paragraphs).toEqual(paragraphs);
    if (cells) expect(result.cells).toEqual(cells);
  }
  expect(right.xml).toBe(left.xml);
  return left;
}

async function undoRedo(page: Page, index: number, undone: string[], redone: string[]) {
  await input(page, index).press('ControlOrMeta+z');
  await converged(page, undone);
  await saved(page, undone);
  await input(page, index).press('ControlOrMeta+Shift+z');
  await converged(page, redone);
}

async function paintedBold(page: Page, expected: string) {
  for (const index of [0, 1]) {
    await expect.poll(() => editor(page, index).locator('.canvas-page-mirror').evaluateAll(mirrors => {
      let value = '';
      for (const mirror of mirrors) {
        const walker = document.createTreeWalker(mirror, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (Number.parseInt(getComputedStyle(node.parentElement!).fontWeight, 10) >= 600) value += node.textContent;
        }
      }
      return value;
    })).toBe(expected);
  }
}

const tableBody = () => `${paragraph('Top')}<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr>${['abcdef', 'South'].map(value => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph(value)}</w:tc>`).join('')}</w:tr></w:tbl>${paragraph('Bottom')}`;

async function cellAt(page: Page, index: number, cell: number, offset: number) {
  const target = editor(page, index).locator('.layout-table-cell').nth(cell).locator('.layout-run-text').first();
  const bounds = await renderedBox(target);
  await page.mouse.click(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
  await input(page, index).press('Home');
  for (let step = 0; step < offset; step++) await input(page, index).press('ArrowRight');
  return input(page, index);
}

async function cellsConverged(page: Page, cells: string[]) {
  for (const index of [0, 1]) await expect(editor(page, index).locator('.layout-table-cell')).toHaveText(cells);
  await vectorsMatch(page);
}

async function remoteType(page: Page, index: number, value: string) {
  await input(page, index).evaluate((element, value) => element.dispatchEvent(new InputEvent('beforeinput', {
    bubbles: true, cancelable: true, inputType: 'insertText', data: value,
  })), value);
}

test.afterEach(async ({ page }) => {
  if (await page.evaluate(() => '__docxPeerlessCollab' in window)) {
    expect(await page.evaluate(() => window.__docxPeerlessCollab.errors)).toEqual([]);
  }
});

test('concurrent typing at different offsets keeps both insertions in order', async ({ page }) => {
  await open(page);
  await converged(page, ['abcdef']);
  await hold(page);
  await (await at(page, 0, 1)).pressSequentially('A');
  await (await at(page, 1, 5)).pressSequentially('B');
  await expect.poll(() => text(page, 0)).toBe('aAbcdef');
  await expect.poll(() => text(page, 1)).toBe('abcdeBf');
  await release(page);
  await converged(page, ['aAbcdeBf']);
  await undoRedo(page, 0, ['abcdeBf'], ['aAbcdeBf']);
  await saved(page, ['aAbcdeBf']);
});

test('deleting a paragraph preserves text concurrently typed inside it', async ({ page }) => {
  await open(page, ['Start', 'Delete me', 'End'].map(value => paragraph(value)).join(''));
  await converged(page, ['Start', 'Delete me', 'End']);
  await hold(page);
  const deleting = await at(page, 0, 6);
  for (let step = 0; step < 10; step++) await deleting.press('Shift+ArrowRight');
  await deleting.press('Backspace');
  await (await at(page, 1, 9)).pressSequentially('X');
  await expect.poll(() => text(page, 0)).toBe('Start\nEnd');
  await expect.poll(() => text(page, 1)).toBe('Start\nDelXete me\nEnd');
  await release(page);
  await converged(page, ['Start', 'XEnd']);
  await undoRedo(page, 0, ['Start', 'DelXete me', 'End'], ['Start', 'XEnd']);
  await saved(page, ['Start', 'XEnd']);
});

test('bold formatting merges with typing inside the formatted range', async ({ page }) => {
  await open(page);
  await converged(page, ['abcdef']);
  await hold(page);
  const formatting = await at(page, 0, 1);
  for (let step = 0; step < 4; step++) await formatting.press('Shift+ArrowRight');
  await formatting.press('ControlOrMeta+b');
  await (await at(page, 1, 3)).pressSequentially('X');
  await release(page);
  await converged(page, ['abcXdef']);
  await paintedBold(page, 'bcXde');
  expect((await saved(page, ['abcXdef'])).bold).toBe('bcXde');
  await formatting.press('ControlOrMeta+z');
  await converged(page, ['abcXdef']);
  await paintedBold(page, '');
  expect((await saved(page, ['abcXdef'])).bold).toBe('');
  await formatting.press('ControlOrMeta+Shift+z');
  await paintedBold(page, 'bcXde');
  await vectorsMatch(page);
  expect((await saved(page, ['abcXdef'])).bold).toBe('bcXde');
});

test('simultaneous edits in different table cells stay in their cells', async ({ page }) => {
  await open(page, tableBody());
  await cellsConverged(page, ['abcdef', 'South']);
  await hold(page);
  await (await cellAt(page, 0, 0, 6)).pressSequentially('A');
  await (await cellAt(page, 1, 1, 5)).pressSequentially('B');
  await release(page);
  await cellsConverged(page, ['abcdefA', 'SouthB']);
  await input(page, 0).press('ControlOrMeta+z');
  await cellsConverged(page, ['abcdef', 'SouthB']);
  await saved(page, ['Top', 'abcdef', 'SouthB', 'Bottom'], ['abcdef', 'SouthB']);
  await input(page, 0).press('ControlOrMeta+Shift+z');
  await cellsConverged(page, ['abcdefA', 'SouthB']);
  const result = await saved(page, ['Top', 'abcdefA', 'SouthB', 'Bottom'], ['abcdefA', 'SouthB']);
  await load(page, result.bytes);
  await cellsConverged(page, ['abcdefA', 'SouthB']);
  await saved(page, ['Top', 'abcdefA', 'SouthB', 'Bottom'], ['abcdefA', 'SouthB']);
});

test('simultaneous edits in the same table cell preserve both offsets', async ({ page }) => {
  await open(page, tableBody());
  await cellsConverged(page, ['abcdef', 'South']);
  await hold(page);
  await (await cellAt(page, 0, 0, 1)).pressSequentially('A');
  await (await cellAt(page, 1, 0, 5)).pressSequentially('B');
  await release(page);
  await cellsConverged(page, ['aAbcdeBf', 'South']);
  await input(page, 1).press('ControlOrMeta+z');
  await cellsConverged(page, ['aAbcdef', 'South']);
  await saved(page, ['Top', 'aAbcdef', 'South', 'Bottom'], ['aAbcdef', 'South']);
  await input(page, 1).press('ControlOrMeta+Shift+z');
  await cellsConverged(page, ['aAbcdeBf', 'South']);
  await saved(page, ['Top', 'aAbcdeBf', 'South', 'Bottom'], ['aAbcdeBf', 'South']);
});

test('Enter moves concurrent typing after the split point into the second paragraph', async ({ page }) => {
  await open(page);
  await converged(page, ['abcdef']);
  await hold(page);
  await (await at(page, 0, 3)).press('Enter');
  await (await at(page, 1, 5)).pressSequentially('X');
  await expect.poll(() => text(page, 0)).toBe('abc\ndef');
  await expect.poll(() => text(page, 1)).toBe('abcdeXf');
  await release(page);
  await converged(page, ['abc', 'deXf']);
  await undoRedo(page, 0, ['abcdeXf'], ['abc', 'deXf']);
  const result = await saved(page, ['abc', 'deXf']);
  await load(page, result.bytes);
  await converged(page, ['abc', 'deXf']);
  await saved(page, ['abc', 'deXf']);
});

test('each peer undoes and redoes only its own concurrent insertion', async ({ page }) => {
  await open(page, paragraph('Seed'));
  await converged(page, ['Seed']);
  await hold(page);
  await (await at(page, 0, 0)).pressSequentially('A');
  await (await at(page, 1, 4)).pressSequentially('B');
  await release(page);
  await converged(page, ['ASeedB']);
  await input(page, 0).press('ControlOrMeta+z');
  await converged(page, ['SeedB']);
  await saved(page, ['SeedB']);
  await input(page, 1).press('ControlOrMeta+z');
  await converged(page, ['Seed']);
  await saved(page, ['Seed']);
  await input(page, 0).press('ControlOrMeta+Shift+z');
  await converged(page, ['ASeed']);
  await input(page, 1).press('ControlOrMeta+Shift+z');
  await converged(page, ['ASeedB']);
  await saved(page, ['ASeedB']);
});

test('remote typing rebases an open IME composition without losing its buffer', async ({ page }) => {
  await open(page);
  await converged(page, ['abcdef']);
  await at(page, 1, 1);
  const composing = await at(page, 0, 3);
  await composing.evaluate(element => {
    const textarea = element as HTMLTextAreaElement;
    textarea.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
    textarea.value = 'かな';
    textarea.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: 'かな' }));
    textarea.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: 'かな', isComposing: true }));
    textarea.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: 'かな', isComposing: true }));
  });
  await remoteType(page, 1, 'R');
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessCollab.painted(0))).toContain('aRbcdef');
  await expect(composing).toBeFocused();
  await expect(composing).toHaveValue('かな');
  await expect.poll(() => text(page, 1)).toBe('aRbcdef');
  await composing.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '漢字' })));
  await converged(page, ['aRbc漢字def']);
  await expect(composing).toHaveValue('');
  await undoRedo(page, 0, ['aRbcdef'], ['aRbc漢字def']);
  await saved(page, ['aRbc漢字def']);
});

test('a remote edit on another page preserves the local viewport and caret', async ({ page }) => {
  await open(page, paragraph('First page') + paragraph('Middle page', true) + paragraph('Last page', true));
  const remote = input(page, 1);
  await remote.focus();
  await remote.press('ControlOrMeta+End');
  await expect(editor(page, 1).locator('.canvas-page[data-page-index="2"]')).toBeInViewport();
  const local = await at(page, 0, 5);
  await expect.poll(() => text(page, 0)).toBe('First page\nMiddle page\nLast page');
  const scroller = editor(page, 0).locator('.docx-editor__scroll-container');
  await expect(editor(page, 0).locator('.canvas-page[data-page-index="0"]')).toBeInViewport();
  await expect.poll(() => local.evaluate(element => Number.parseFloat((element as HTMLElement).style.height))).toBeGreaterThan(1);
  const before = await scroller.evaluate(element => ({ top: element.scrollTop, left: element.scrollLeft }));
  const caret = await local.evaluate(element => ({ left: (element as HTMLElement).style.left, top: (element as HTMLElement).style.top }));
  await remoteType(page, 1, 'R');
  for (const index of [0, 1]) await expect.poll(() => text(page, index)).toBe('First page\nMiddle page\nLast pageR');
  await vectorsMatch(page);
  for (const index of [0, 1]) await expect.poll(() => page.evaluate(index => window.__docxPeerlessCollab.painted(index), index)).toContain('Last pageR');
  await expect(local).toBeFocused();
  await expect.poll(() => scroller.evaluate(element => ({ top: element.scrollTop, left: element.scrollLeft }))).toEqual(before);
  await expect.poll(() => local.evaluate(element => ({ left: (element as HTMLElement).style.left, top: (element as HTMLElement).style.top }))).toEqual(caret);
  await local.pressSequentially('L');
  await converged(page, ['FirstL page', 'Middle page', 'Last pageR'], ['FirstL page']);
  await local.press('ControlOrMeta+z');
  await converged(page, ['First page', 'Middle page', 'Last pageR'], ['First page']);
  await saved(page, ['First page', 'Middle page', 'Last pageR']);
  await local.press('ControlOrMeta+Shift+z');
  await converged(page, ['FirstL page', 'Middle page', 'Last pageR'], ['FirstL page']);
  await saved(page, ['FirstL page', 'Middle page', 'Last pageR']);
});

test('both peers save identical document XML after convergence and reopening', async ({ page }) => {
  await open(page, paragraph('Opening') + paragraph('Closing'));
  await converged(page, ['Opening', 'Closing']);
  await hold(page);
  await (await at(page, 0, 0)).pressSequentially('A');
  const right = input(page, 1);
  await right.focus();
  await right.press('ControlOrMeta+End');
  await right.pressSequentially('B');
  await release(page, true);
  await converged(page, ['AOpening', 'ClosingB']);
  await undoRedo(page, 1, ['AOpening', 'Closing'], ['AOpening', 'ClosingB']);
  const result = await saved(page, ['AOpening', 'ClosingB']);
  await load(page, result.bytes);
  await converged(page, ['AOpening', 'ClosingB']);
  expect((await saved(page, ['AOpening', 'ClosingB'])).xml).toBe(result.xml);
});
