import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

type Paragraph = { runs: string[]; pageBreakBefore?: boolean };

const dialog = (page: Page) => page.getByTestId('find-replace-dialog');
const findInput = (page: Page) => dialog(page).getByRole('textbox', { name: 'Find text', exact: true });
const replaceInput = (page: Page) => dialog(page).getByRole('textbox', { name: 'Replace text', exact: true });
const documentInput = (page: Page) => page.locator('[data-docx-worker-editor] .paged-editor__input');
const status = (page: Page) => dialog(page).locator('.docx-find-replace-dialog-status');
const highlights = (page: Page) => page.locator('[data-docx-pages-host] > div[aria-hidden="true"][style*="255,"]');
const currentHighlight = (page: Page) => page.locator('[data-docx-pages-host] > div[aria-hidden="true"][style*="255, 150, 0"]');

function escapeXml(value: string) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function fixture(paragraphs: Paragraph[]) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>');
  zip.file('_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  zip.file('word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    paragraphs.map(({ runs, pageBreakBefore }) => '<w:p>' +
      (pageBreakBefore ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '') +
      runs.map((text, index) => '<w:r>' + (index % 2 ? '<w:rPr><w:b/></w:rPr>' : '') +
        '<w:t xml:space="preserve">' + escapeXml(text) + '</w:t></w:r>').join('') + '</w:p>').join('') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
    '</w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function expectRendered(page: Page, paragraphs: string[]) {
  await expect.poll(async () => (await page.locator('[data-docx-worker-editor] .canvas-page-mirror').allTextContents()).join(''))
    .toBe(paragraphs.join(''));
}

async function open(page: Page, paragraphs: Paragraph[], savedFile?: string) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(savedFile ?? {
    name: 'find-replace.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await fixture(paragraphs),
  });
  await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false');
  const canvas = page.locator('[data-docx-worker-editor] .canvas-page canvas').first();
  await expect(canvas).toBeVisible();
  await expectRendered(page, paragraphs.map(paragraph => paragraph.runs.join('')));
  await expect(page.locator('.notice')).toHaveCount(0);
  await canvas.click({ position: { x: 150, y: 150 } });
  await expect(documentInput(page)).toBeEditable();
  await documentInput(page).press('ControlOrMeta+Home');
}

async function openFind(page: Page, replace = false) {
  await documentInput(page).press(replace ? 'ControlOrMeta+h' : 'ControlOrMeta+f');
  await expect(dialog(page)).toBeVisible();
  await expect(dialog(page).getByRole('heading', { name: replace ? 'Find and Replace' : 'Find', exact: true })).toBeVisible();
  await expect(findInput(page)).toBeFocused();
  if (replace) await expect(replaceInput(page)).toBeVisible();
}

async function expectMatches(page: Page, current: number, total: number) {
  await expect(status(page)).toHaveText(`${current} of ${total} matches`);
  await expect(highlights(page)).toHaveCount(total);
  await expect(currentHighlight(page)).toHaveCount(1);
}

async function find(page: Page, query: string, total: number) {
  await findInput(page).fill(query);
  await expectMatches(page, 1, total);
}

async function closeFind(page: Page) {
  await dialog(page).getByRole('button', { name: 'Close dialog', exact: true }).click();
  await expect(dialog(page)).toBeHidden();
  await expect(highlights(page)).toHaveCount(0);
}

async function save(page: Page, info: TestInfo, label: string, paragraphs: string[]) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const download = await downloading;
  const file = info.outputPath(label + '.docx');
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const documentPart = zip.file('word/document.xml');
  expect(documentPart).not.toBeNull();
  const xml = await documentPart!.async('string');
  const savedParagraphs = await page.evaluate(source => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved document.xml');
    const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return Array.from(document.getElementsByTagNameNS(namespace, 'p'), paragraph =>
      Array.from(paragraph.getElementsByTagNameNS(namespace, 't'), node => node.textContent ?? '').join(''));
  }, xml);
  expect(savedParagraphs).toEqual(paragraphs);
  return file;
}

async function reopenAndSave(page: Page, info: TestInfo, file: string, paragraphs: string[]) {
  await open(page, paragraphs.map(text => ({ runs: [text] })), file);
  await save(page, info, 'reopened', paragraphs);
}

async function expectHistoryAndReopen(page: Page, info: TestInfo, before: string[], after: string[]) {
  await expectRendered(page, after);
  await save(page, info, 'replaced', after);
  await documentInput(page).press('ControlOrMeta+z');
  await expectRendered(page, before);
  await save(page, info, 'undone', before);
  await documentInput(page).press('ControlOrMeta+Shift+z');
  await expectRendered(page, after);
  const redone = await save(page, info, 'redone', after);
  await reopenAndSave(page, info, redone, after);
}

async function currentLeft(page: Page) {
  return currentHighlight(page).evaluate(element => (element as HTMLElement).offsetLeft);
}

async function expectCurrentOnPage(page: Page, pageIndex: number) {
  await expect(currentHighlight(page)).toBeInViewport();
  await expect.poll(async () => {
    const highlight = await currentHighlight(page).boundingBox();
    const canvas = await page.locator(`[data-docx-worker-editor] .canvas-page[data-page-index="${pageIndex}"] canvas`).boundingBox();
    return !!highlight && !!canvas && highlight.x >= canvas.x && highlight.y >= canvas.y &&
      highlight.x + highlight.width <= canvas.x + canvas.width &&
      highlight.y + highlight.height <= canvas.y + canvas.height;
  }).toBe(true);
}

test('docx: Mod+F finds and highlights case-insensitive canvas matches', async ({ page }, info) => {
  const text = 'cat Cat dog cat';
  await open(page, [{ runs: [text] }]);
  await openFind(page);
  await expect(replaceInput(page)).toHaveCount(0);
  await find(page, 'cat', 3);
  await expect(currentHighlight(page)).toBeInViewport();
  for (const highlight of await highlights(page).all()) {
    await expect(highlight).toBeVisible();
    expect((await highlight.boundingBox())!.width).toBeGreaterThan(0);
    expect((await highlight.boundingBox())!.height).toBeGreaterThan(0);
  }
  await closeFind(page);
  const file = await save(page, info, 'unchanged', [text]);
  await reopenAndSave(page, info, file, [text]);
});

test('docx: Enter and Shift+Enter cycle matches and wrap', async ({ page }, info) => {
  const text = 'cat dog cat fox cat';
  await open(page, [{ runs: [text] }]);
  await openFind(page);
  await find(page, 'cat', 3);
  const firstLeft = await currentLeft(page);
  await findInput(page).press('Enter');
  await expectMatches(page, 2, 3);
  await expect.poll(() => currentLeft(page)).toBeGreaterThan(firstLeft);
  const secondLeft = await currentLeft(page);
  await findInput(page).press('Shift+Enter');
  await expectMatches(page, 1, 3);
  await expect.poll(() => currentLeft(page)).toBe(firstLeft);
  await findInput(page).press('Shift+Enter');
  await expectMatches(page, 3, 3);
  await expect.poll(() => currentLeft(page)).toBeGreaterThan(secondLeft);
  await findInput(page).press('Enter');
  await expectMatches(page, 1, 3);
  await expect.poll(() => currentLeft(page)).toBe(firstLeft);
  await closeFind(page);
  await save(page, info, 'unchanged', [text]);
});

test('docx: find buttons reveal matches on later pages', async ({ page }, info) => {
  const paragraphs = ['cat first', 'cat middle', 'cat last'];
  await open(page, paragraphs.map((text, index) => ({ runs: [text], pageBreakBefore: index > 0 })));
  await expect(page.locator('[data-docx-worker-editor] .canvas-page')).toHaveCount(3);
  await openFind(page);
  await find(page, 'cat', 3);
  await expectCurrentOnPage(page, 0);
  await dialog(page).getByRole('button', { name: 'Find next', exact: true }).click();
  await expectMatches(page, 2, 3);
  await expectCurrentOnPage(page, 1);
  await findInput(page).press('Enter');
  await expectMatches(page, 3, 3);
  await expectCurrentOnPage(page, 2);
  await dialog(page).getByRole('button', { name: 'Find next', exact: true }).click();
  await expectMatches(page, 1, 3);
  await expectCurrentOnPage(page, 0);
  await dialog(page).getByRole('button', { name: 'Find previous', exact: true }).click();
  await expectMatches(page, 3, 3);
  await expectCurrentOnPage(page, 2);
  await findInput(page).press('Shift+Enter');
  await expectMatches(page, 2, 3);
  await expectCurrentOnPage(page, 1);
  await closeFind(page);
  await save(page, info, 'unchanged', paragraphs);
});

test('docx: Match case filters and restores highlights', async ({ page }, info) => {
  const text = 'cat Cat CAT catapult bobcat cat.';
  await open(page, [{ runs: [text] }]);
  await openFind(page);
  await find(page, 'cat', 6);
  const matchCase = dialog(page).getByRole('checkbox', { name: 'Match case', exact: true });
  await matchCase.check();
  await expect(matchCase).toBeChecked();
  await expectMatches(page, 1, 4);
  await find(page, 'CAT', 1);
  await matchCase.uncheck();
  await expect(matchCase).not.toBeChecked();
  await expectMatches(page, 1, 6);
  await closeFind(page);
  await save(page, info, 'unchanged', [text]);
});

test('docx: Whole words excludes substrings and respects punctuation', async ({ page }, info) => {
  const text = 'cat catapult bobcat cat. (cat) cat-cat';
  await open(page, [{ runs: [text] }]);
  await openFind(page);
  await find(page, 'cat', 7);
  const wholeWords = dialog(page).getByRole('checkbox', { name: 'Whole words', exact: true });
  await wholeWords.check();
  await expect(wholeWords).toBeChecked();
  await expectMatches(page, 1, 5);
  await wholeWords.uncheck();
  await expect(wholeWords).not.toBeChecked();
  await expectMatches(page, 1, 7);
  await closeFind(page);
  await save(page, info, 'unchanged', [text]);
});

test('docx: no matches disables navigation and replacement', async ({ page }, info) => {
  const text = 'cat dog cat';
  await open(page, [{ runs: [text] }]);
  await openFind(page, true);
  await find(page, 'cat', 2);
  await findInput(page).fill('unfindable');
  await expect(status(page)).toHaveText('No results found');
  await expect(highlights(page)).toHaveCount(0);
  for (const name of ['Find next', 'Find previous', 'Replace', 'Replace All']) {
    await expect(dialog(page).getByRole('button', { name, exact: true })).toBeDisabled();
  }
  await findInput(page).press('Enter');
  await expect(status(page)).toHaveText('No results found');
  await find(page, 'dog', 1);
  await expect(dialog(page).getByRole('button', { name: 'Replace', exact: true })).toBeEnabled();
  await closeFind(page);
  await save(page, info, 'unchanged', [text]);
});

test('docx: clearing a query removes stale counts and highlights', async ({ page }, info) => {
  const text = 'cat dog cat';
  await open(page, [{ runs: [text] }]);
  await openFind(page, true);
  await find(page, 'cat', 2);
  for (const query of ['', '   ']) {
    await findInput(page).fill(query);
    await expect(status(page)).toHaveCount(0);
    await expect(highlights(page)).toHaveCount(0);
    for (const name of ['Find next', 'Find previous', 'Replace', 'Replace All']) {
      await expect(dialog(page).getByRole('button', { name, exact: true })).toBeDisabled();
    }
  }
  await find(page, 'dog', 1);
  await closeFind(page);
  await save(page, info, 'unchanged', [text]);
});

test('docx: find treats punctuation as literal text', async ({ page }, info) => {
  const text = 'a.b aXb [cat] cat';
  await open(page, [{ runs: [text] }]);
  await openFind(page);
  await find(page, 'a.b', 1);
  const firstLeft = await currentLeft(page);
  await find(page, '[cat]', 1);
  await expect.poll(() => currentLeft(page)).toBeGreaterThan(firstLeft);
  await closeFind(page);
  await save(page, info, 'unchanged', [text]);
});

test('docx: replace toggle replaces only the chosen match', async ({ page }, info) => {
  const before = ['cat dog cat fox cat'];
  const after = ['cat dog otter fox cat'];
  await open(page, [{ runs: before }]);
  await openFind(page);
  await find(page, 'cat', 3);
  await dialog(page).getByRole('button', { name: 'Find next', exact: true }).click();
  await expectMatches(page, 2, 3);
  await dialog(page).getByRole('button', { name: '+ Replace', exact: true }).click();
  await expect(replaceInput(page)).toBeVisible();
  await expect(replaceInput(page)).toBeFocused();
  await replaceInput(page).fill('otter');
  await dialog(page).getByRole('button', { name: 'Replace', exact: true }).click();
  await expectRendered(page, after);
  await expect(status(page)).toHaveText(/^[12] of 2 matches$/);
  await expect(highlights(page)).toHaveCount(2);
  await closeFind(page);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: Mod+H and Enter replace the current match', async ({ page }, info) => {
  const before = ['cat dog cat'];
  const after = ['lynx dog cat'];
  await open(page, [{ runs: before }]);
  await openFind(page, true);
  await find(page, 'cat', 2);
  await replaceInput(page).fill('lynx');
  await replaceInput(page).press('Enter');
  await expectRendered(page, after);
  await expectMatches(page, 1, 1);
  await closeFind(page);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: an empty replacement deletes the match', async ({ page }, info) => {
  const before = ['start cat end'];
  const after = ['start end'];
  await open(page, [{ runs: before }]);
  await openFind(page, true);
  await find(page, 'cat ', 1);
  await expect(replaceInput(page)).toHaveValue('');
  await dialog(page).getByRole('button', { name: 'Replace', exact: true }).click();
  await expectRendered(page, after);
  await expect(status(page)).toHaveText('No results found');
  await expect(highlights(page)).toHaveCount(0);
  await closeFind(page);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: Replace All is one undo step separate from earlier typing', async ({ page }, info) => {
  await open(page, [{ runs: ['cat dog cat'] }, { runs: ['cat fox'] }]);
  await documentInput(page).press('ControlOrMeta+End');
  await page.keyboard.insertText('!');
  const before = ['cat dog cat', 'cat fox!'];
  const after = ['otter dog otter', 'otter fox!'];
  await expectRendered(page, before);
  await openFind(page, true);
  await find(page, 'cat', 3);
  await replaceInput(page).fill('otter');
  await dialog(page).getByRole('button', { name: 'Replace All', exact: true }).click();
  await expectRendered(page, after);
  await expect(highlights(page)).toHaveCount(0);
  await closeFind(page);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: Replace All respects Match case and Whole words together', async ({ page }, info) => {
  const before = ['cat Cat catapult cat.'];
  const after = ['dog Cat catapult dog.'];
  await open(page, [{ runs: before }]);
  await openFind(page, true);
  await find(page, 'cat', 4);
  await dialog(page).getByRole('checkbox', { name: 'Match case', exact: true }).check();
  await expectMatches(page, 1, 3);
  await dialog(page).getByRole('checkbox', { name: 'Whole words', exact: true }).check();
  await expectMatches(page, 1, 2);
  await replaceInput(page).fill('dog');
  await dialog(page).getByRole('button', { name: 'Replace All', exact: true }).click();
  await expectRendered(page, after);
  await expect(highlights(page)).toHaveCount(0);
  await closeFind(page);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: replacement crosses formatted runs and preserves literal XML characters', async ({ page }, info) => {
  const runs = ['pre nee', 'dle', ' post needle.'];
  const before = [runs.join('')];
  const after = ['pre naïve & <fox> post naïve & <fox>.'];
  await open(page, [{ runs }]);
  await openFind(page, true);
  await findInput(page).fill('needle');
  await expect(status(page)).toHaveText('1 of 2 matches');
  await expect.poll(() => highlights(page).count()).toBeGreaterThanOrEqual(2);
  await expect(currentHighlight(page).first()).toBeVisible();
  await replaceInput(page).fill('naïve & <fox>');
  await dialog(page).getByRole('button', { name: 'Replace All', exact: true }).click();
  await expectRendered(page, after);
  await expect(highlights(page)).toHaveCount(0);
  await closeFind(page);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: Escape restores document focus and the current match selection', async ({ page }, info) => {
  const before = ['cat cat'];
  const after = ['cat otter'];
  await open(page, [{ runs: before }]);
  await openFind(page);
  await find(page, 'cat', 2);
  await findInput(page).press('Enter');
  await expectMatches(page, 2, 2);
  await findInput(page).press('Escape');
  await expect(dialog(page)).toBeHidden();
  await expect(highlights(page)).toHaveCount(0);
  await expect(documentInput(page)).toBeFocused();
  await page.keyboard.insertText('otter');
  await expectRendered(page, after);
  await expectHistoryAndReopen(page, info, before, after);
});

test('docx: Escape immediately after editing the query does not select the abandoned match', async ({ page }, info) => {
  const before = ['cat dog cat'];
  const after = ['ottercat dog cat'];
  await open(page, [{ runs: before }]);
  await openFind(page);
  await find(page, 'cat', 2);
  await findInput(page).press('Enter');
  await expectMatches(page, 2, 2);
  await findInput(page).fill('unfindable');
  await findInput(page).press('Escape');
  await expect(dialog(page)).toBeHidden();
  await expect(highlights(page)).toHaveCount(0);
  await expect(documentInput(page)).toBeFocused();
  await page.keyboard.insertText('otter');
  await expectRendered(page, after);
  await expectHistoryAndReopen(page, info, before, after);
});
