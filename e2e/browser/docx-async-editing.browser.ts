import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const root = '[data-peerless-docx]';
const caretSelector = `${root} div[aria-hidden="true"][style*="currentcolor" i]`;
const selectionSelector = `${root} div[aria-hidden="true"][style*="rgba(59, 130, 246"]`;
const input = (page: Page) => page.locator('.paged-editor__peerless-input');
const compactPage = { width: 7200, height: 4320, margin: 720 };

function escapeXml(text: string) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function fixture(paragraphs: string[], options: {
  page?: { width: number; height: number; margin: number };
  pageBreaks?: number[];
} = {}) {
  const zip = new JSZip();
  const size = options.page ?? { width: 12240, height: 15840, margin: 1440 };
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${wordNamespace}"><w:body>` +
    paragraphs.map((text, index) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/>' +
      (options.pageBreaks?.includes(index) ? '<w:pageBreakBefore/>' : '') +
      '</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>' +
      `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`).join('') +
    `<w:sectPr><w:pgSz w:w="${size.width}" w:h="${size.height}"/>` +
    `<w:pgMar w:top="${size.margin}" w:right="${size.margin}" w:bottom="${size.margin}" w:left="${size.margin}"/>` +
    '</w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, buffer: Buffer) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'editing.docx', mimeType, buffer });
  await expect(page.locator(`${root} [data-page-index="0"] canvas`)).toBeVisible();
  await expect(input(page)).toBeEditable();
  await input(page).focus();
  await input(page).press('ControlOrMeta+Home');
  await expect(page.locator(caretSelector)).toHaveCount(1);
  await expect(page.locator('.notice')).toHaveCount(0);
}

async function mirrorText(page: Page) {
  return page.locator(`${root} .canvas-page-mirror`).evaluateAll(nodes =>
    nodes.map(node => node.textContent ?? '').join('').replace(/\s/g, '')
  );
}

async function expectMirror(page: Page, paragraphs: string[]) {
  await expect.poll(() => mirrorText(page)).toBe(paragraphs.join('').replace(/\s/g, ''));
}

async function save(page: Page) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const stream = await (await downloading).createReadStream();
  if (!stream) throw new Error('Save did not produce a download');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  const buffer = Buffer.concat(chunks);
  const zip = await JSZip.loadAsync(buffer);
  const part = zip.file('word/document.xml');
  if (!part) throw new Error('Saved document has no word/document.xml');
  const xml = await part.async('string');
  const document = await page.evaluate(({ xml, wordNamespace }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved document.xml');
    const body = document.getElementsByTagNameNS(wordNamespace, 'body')[0];
    if (!body) throw new Error('Saved document has no body');
    const paragraphs = Array.from(body.children).filter(node => node.localName === 'p').map(paragraph =>
      Array.from(paragraph.getElementsByTagNameNS(wordNamespace, '*')).map(node =>
        node.localName === 't' ? node.textContent ?? '' : node.localName === 'br' ? '\n' : node.localName === 'tab' ? '\t' : ''
      ).join('')
    );
    return {
      paragraphs,
      breaks: Array.from(body.getElementsByTagNameNS(wordNamespace, 'br')).map(node => node.getAttributeNS(wordNamespace, 'type') ?? 'textWrapping'),
      tabs: body.getElementsByTagNameNS(wordNamespace, 'tab').length,
    };
  }, { xml, wordNamespace });
  return { buffer, ...document };
}

async function expectDocument(page: Page, paragraphs: string[]) {
  await expectMirror(page, paragraphs);
  const saved = await save(page);
  expect(saved.paragraphs).toEqual(paragraphs);
  expect(saved.breaks).toEqual(Array.from({ length: paragraphs.join('').split('\n').length - 1 }, () => 'textWrapping'));
  expect(saved.tabs).toBe(paragraphs.join('').split('\t').length - 1);
  return saved;
}

async function history(page: Page, before: string[], after: string[]) {
  await expectDocument(page, after);
  await input(page).press('ControlOrMeta+z');
  await expectDocument(page, before);
  await input(page).press('ControlOrMeta+Shift+z');
  await expectDocument(page, after);
}

async function roundTrip(page: Page, paragraphs: string[]) {
  const saved = await expectDocument(page, paragraphs);
  await open(page, saved.buffer);
  await expectDocument(page, paragraphs);
}

async function caretBox(page: Page) {
  await expect(page.locator(caretSelector)).toHaveCount(1);
  const box = await page.locator(caretSelector).boundingBox();
  if (!box) throw new Error('Document caret has no bounds');
  return box;
}

async function expectSelection(page: Page) {
  await expect(page.locator(selectionSelector).first()).toBeVisible();
}

for (const modifier of ['Control', 'Alt']) {
  test(`${modifier} word arrows move between space-separated words in both directions`, async ({ page }) => {
    const original = ['Alpha beta gamma'];
    await open(page, await fixture(original));
    await input(page).press(`${modifier}+ArrowRight`);
    await page.keyboard.insertText('X');
    const forward = ['Alpha Xbeta gamma'];
    await history(page, original, forward);
    await input(page).press('ControlOrMeta+End');
    await input(page).press(`${modifier}+ArrowLeft`);
    await page.keyboard.insertText('Y');
    const backward = [forward[0]!.replace('gamma', 'Ygamma')];
    await history(page, forward, backward);
    await roundTrip(page, backward);
  });

  test(`${modifier} shift word arrows select space-separated words for replacement`, async ({ page }) => {
    const original = ['Alpha beta gamma'];
    await open(page, await fixture(original));
    await input(page).press(`${modifier}+Shift+ArrowRight`);
    await expectSelection(page);
    await page.keyboard.insertText('New ');
    const forward = ['New beta gamma'];
    await history(page, original, forward);
    await input(page).press('ControlOrMeta+End');
    await input(page).press(`${modifier}+Shift+ArrowLeft`);
    await expectSelection(page);
    await page.keyboard.insertText('last');
    const backward = [forward[0]!.replace('gamma', 'last')];
    await history(page, forward, backward);
    await roundTrip(page, backward);
  });

  test(`${modifier} word Backspace and Delete remove space-separated words in either direction`, async ({ page }) => {
    const original = ['Alpha beta gamma'];
    await open(page, await fixture(original));
    await input(page).press('ControlOrMeta+End');
    await input(page).press(`${modifier}+Backspace`);
    const backward = ['Alpha beta '];
    await history(page, original, backward);
    await input(page).press('ControlOrMeta+Home');
    await input(page).press(`${modifier}+Delete`);
    const forward = ['beta '];
    await history(page, backward, forward);
    await roundTrip(page, forward);
  });
}

for (const key of ['Home', 'End']) {
  test(`shift ${key} selects to the edge of a middle visual line in a wrapped paragraph`, async ({ page }) => {
    const text = 'Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega.';
    const original = [text];
    await open(page, await fixture(original, { page: compactPage }));
    const lines = await page.locator(`${root} .canvas-page-mirror .layout-run-text[data-doc-start][data-doc-end]`).evaluateAll(nodes => {
      const lines: { from: number; to: number; top: number; bottom: number }[] = [];
      for (const node of nodes) {
        const box = node.getBoundingClientRect();
        const from = Number(node.getAttribute('data-doc-start'));
        const to = Number(node.getAttribute('data-doc-end'));
        const line = lines.find(line => Math.abs(line.top - box.top) < 1);
        if (line) {
          line.from = Math.min(line.from, from);
          line.to = Math.max(line.to, to);
          line.bottom = Math.max(line.bottom, box.bottom);
        } else lines.push({ from, to, top: box.top, bottom: box.bottom });
      }
      return lines.sort((a, b) => a.top - b.top);
    });
    expect(lines.length).toBeGreaterThanOrEqual(3);
    const line = lines[1]!;
    const base = lines[0]!.from;
    const from = line.from - base;
    const to = line.to - base;
    expect(from).toBeGreaterThan(0);
    expect(to).toBeLessThan(text.length);
    expect(to - from).toBeGreaterThan(3);
    await input(page).press('ArrowDown');
    for (let index = 0; index < 3; index += 1) await input(page).press('ArrowRight');
    await expect.poll(async () => {
      const box = await page.locator(caretSelector).boundingBox();
      return box ? box.y + box.height / 2 : -1;
    }).toBeGreaterThan(line.top);
    const caret = await caretBox(page);
    expect(caret.y + caret.height / 2).toBeLessThan(line.bottom);
    await input(page).press(`Shift+${key}`);
    await expectSelection(page);
    const selection = await page.locator(selectionSelector).evaluateAll(nodes => nodes.map(node => {
      const box = node.getBoundingClientRect();
      return { y: box.y, height: box.height };
    }));
    expect(selection.length).toBeGreaterThan(0);
    for (const box of selection) {
      expect(box.y + box.height / 2).toBeGreaterThan(line.top);
      expect(box.y + box.height / 2).toBeLessThan(line.bottom);
    }
    await page.keyboard.insertText('X');
    const edited = [key === 'Home' ? text.slice(0, from) + 'X' + text.slice(from + 3)
      : text.slice(0, from + 3) + 'X' + text.slice(to)];
    await history(page, original, edited);
    await roundTrip(page, edited);
  });
}

test('document Home and End reach the first and last pages', async ({ page }) => {
  const original = ['First page', 'Middle page', 'Last page'];
  await open(page, await fixture(original, { page: compactPage, pageBreaks: [1, 2] }));
  await input(page).press('ControlOrMeta+End');
  await page.keyboard.insertText(' END');
  const atEnd = ['First page', 'Middle page', 'Last page END'];
  await history(page, original, atEnd);
  await input(page).press('ControlOrMeta+Home');
  await page.keyboard.insertText('START ');
  const both = ['START First page', 'Middle page', 'Last page END'];
  await history(page, atEnd, both);
  await roundTrip(page, both);
});

test('double click selects exactly one word', async ({ page }) => {
  const original = ['Alpha beta gamma'];
  await open(page, await fixture(original));
  const caret = await caretBox(page);
  await page.mouse.dblclick(caret.x + 12, caret.y + caret.height / 2);
  await expectSelection(page);
  await page.keyboard.insertText('Omega');
  const edited = ['Omega beta gamma'];
  await history(page, original, edited);
  await roundTrip(page, edited);
});

test('triple click selects a whole paragraph without selecting its neighbor', async ({ page }) => {
  const original = ['Alpha beta gamma. Another sentence stays in this paragraph.', 'Second paragraph remains.'];
  await open(page, await fixture(original));
  const caret = await caretBox(page);
  await page.mouse.click(caret.x + 12, caret.y + caret.height / 2, { clickCount: 3 });
  await expectSelection(page);
  await page.keyboard.insertText('Replaced paragraph');
  const edited = ['Replaced paragraph', original[1]!];
  await history(page, original, edited);
  await roundTrip(page, edited);
});

test('select all and typing replace the entire document body', async ({ page }) => {
  const original = ['First body paragraph', 'Second body paragraph', 'Last body paragraph'];
  await open(page, await fixture(original, { page: compactPage, pageBreaks: [1, 2] }));
  await input(page).press('ControlOrMeta+A');
  await expectSelection(page);
  await page.keyboard.insertText('Only replacement remains.');
  const edited = ['Only replacement remains.'];
  await history(page, original, edited);
  await roundTrip(page, edited);
});

test('drag across two pages replaces the selected text and merges paragraphs', async ({ page }) => {
  const original = ['Alpha on page one', 'Beta on page two', 'Untouched final paragraph'];
  await open(page, await fixture(original, { page: compactPage, pageBreaks: [1] }));
  await expect(page.locator(`${root} [data-page-index="1"] canvas`)).toBeVisible();
  const start = await caretBox(page);
  const first = await page.locator(`${root} .canvas-page[data-page-index="0"]`).boundingBox();
  const second = await page.locator(`${root} .canvas-page[data-page-index="1"]`).boundingBox();
  if (!first || !second) throw new Error('Document pages have no bounds');
  const scale = second.width / (compactPage.width / 15);
  await page.mouse.move(start.x + 0.25, start.y + start.height / 2);
  await page.mouse.down();
  await page.mouse.move(second.x + second.width - (compactPage.margin / 15) * scale,
    second.y + start.y - first.y + start.height / 2, { steps: 12 });
  await page.mouse.up();
  await expect.poll(() => page.locator(selectionSelector).count()).toBeGreaterThanOrEqual(2);
  await page.keyboard.insertText('Across pages');
  const edited = ['Across pages', original[2]!];
  await history(page, original, edited);
  await roundTrip(page, edited);
});

test('typing across a paragraph boundary keeps the unselected prefix and suffix', async ({ page }) => {
  const original = ['Alpha first', 'Beta second', 'Untouched'];
  await open(page, await fixture(original));
  for (let index = 0; index < 'Alpha '.length; index += 1) await input(page).press('ArrowRight');
  await input(page).press('Shift+ArrowDown');
  await input(page).press('Shift+Home');
  for (let index = 0; index < 'Beta '.length; index += 1) await input(page).press('Shift+ArrowRight');
  await expectSelection(page);
  await page.keyboard.insertText('joined ');
  const edited = ['Alpha joined second', 'Untouched'];
  await history(page, original, edited);
  await roundTrip(page, edited);
});

for (const { title, key, edited } of [
  { title: 'Enter splits a paragraph at the caret', key: 'Enter', edited: ['Alpha', 'Beta'] },
  { title: 'Shift Enter inserts a wrapping break within the same paragraph', key: 'Shift+Enter', edited: ['Alpha\nBeta'] },
]) {
  test(title, async ({ page }) => {
    const original = ['AlphaBeta'];
    await open(page, await fixture(original));
    for (let index = 0; index < 'Alpha'.length; index += 1) await input(page).press('ArrowRight');
    await input(page).press(key);
    await expect(input(page)).toBeFocused();
    await history(page, original, edited);
    await roundTrip(page, edited);
  });
}

test('typing in an empty document groups a sentence and clears redo after new typing', async ({ page }) => {
  await open(page, await fixture(['']));
  await expectDocument(page, ['']);
  await input(page).pressSequentially('A short sentence.');
  const sentence = ['A short sentence.'];
  await history(page, [''], sentence);
  await input(page).press('ControlOrMeta+z');
  await expectDocument(page, ['']);
  await page.keyboard.insertText('Different sentence.');
  const replacement = ['Different sentence.'];
  await expectDocument(page, replacement);
  await input(page).press('ControlOrMeta+Shift+z');
  await expectDocument(page, replacement);
  await history(page, [''], replacement);
  await roundTrip(page, replacement);
});

test('typing a long paragraph flows across a page boundary without splitting it', async ({ page }) => {
  const prefix = 'A long paragraph keeps flowing. '.repeat(8);
  await open(page, await fixture([prefix], { page: compactPage }));
  await expect(page.locator(`${root} [data-page-id]`)).toHaveCount(1);
  await input(page).press('ControlOrMeta+End');
  const suffix = 'Additional words cross the page boundary. '.repeat(12) + 'Final marker.';
  await page.keyboard.insertText(suffix);
  await expect.poll(() => page.locator(`${root} [data-page-id]`).count()).toBeGreaterThan(1);
  await expect(page.locator(`${root} .canvas-page-mirror`).last()).toContainText('Final marker.');
  const edited = [prefix + suffix];
  await history(page, [prefix], edited);
  await roundTrip(page, edited);
});
