import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { pinnedEvaluate } from '../pinned-evaluate';

type Box = { x: number; y: number; width: number; height: number };
interface MiniToolbarProbe {
  editor: {
    readParagraphs(request: unknown): Promise<{ ok: true; version: string; paragraphs: { paraId: string; text: string }[] } | { ok: false }>;
    readSelectionInfo(): Promise<{ selectedText?: string } | null>;
  } | null;
  geometry: {
    readAnchorGeometries(targets: unknown[]): Promise<({ ok: true; rects: Box[] } | { ok: false })[]>;
  } | null;
  load(bytes: number[]): void;
  saved(): Promise<number[]>;
}
declare global { interface Window { __miniToolbar: MiniToolbarProbe } }

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: number, runs: string, pPr = '') =>
  `<w:p w14:paraId="${String(id).padStart(8, '0')}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const BODY = [
  paragraph(1, run('Alpha Bravo Charlie')),
  paragraph(2, run('Delta echo foxtrot. Golf hotel india.')),
  paragraph(3, run('Plain ') + run('Heavy', '<w:b/>') + run(' words')),
  paragraph(4, run('Middle'), '<w:jc w:val="center"/>'),
  paragraph(5, run('Edge'), '<w:jc w:val="right"/>'),
].join('');

async function docx(): Promise<number[]> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body>${BODY}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

const toolbar = (page: Page) => page.getByRole('toolbar', { name: 'Selection formatting' });
const input = (page: Page) => page.locator('.paged-editor__input');
const commentButton = (page: Page) => page.getByRole('button', { name: 'Add comment' });

async function open(page: Page, query = '') {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto(`/docx-mini-toolbar.html${query}`);
  await expect.poll(() => page.evaluate(() => typeof window.__miniToolbar?.load === 'function')).toBe(true);
  await pinnedEvaluate(page, (bytes) => window.__miniToolbar.load(bytes), await docx());
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  await expect.poll(() => page.evaluate(() => !!window.__miniToolbar.geometry && !!window.__miniToolbar.editor), { timeout: 60_000 }).toBe(true);
}

async function wordBox(page: Page, word: string): Promise<Box> {
  let box: Box | null = null;
  await expect.poll(async () => {
    box = await pinnedEvaluate(page, async (text) => {
      const probe = window.__miniToolbar;
      const read = await probe.editor!.readParagraphs({ story: 'body', view: 'accepted' });
      if (!read.ok) return null;
      const found = read.paragraphs.find((value) => value.text.includes(text));
      if (!found) return null;
      const offset = found.text.indexOf(text);
      const [result] = await probe.geometry!.readAnchorGeometries([{ kind: 'range', version: read.version,
        range: { story: 'body', view: 'accepted', start: { paraId: found.paraId, offset }, end: { paraId: found.paraId, offset: offset + text.length } } }]);
      if (!result?.ok || !result.rects[0]) return null;
      const rect = result.rects[0];
      const layer = document.querySelector('[data-testid="plugin-overlays"]')!.getBoundingClientRect();
      return { x: layer.left + rect.x, y: layer.top + rect.y, width: rect.width, height: rect.height };
    }, word);
    return box !== null;
  }, { timeout: 30_000 }).toBe(true);
  return box!;
}

async function doubleClick(page: Page, word: string): Promise<Box> {
  const box = await wordBox(page, word);
  await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
  return box;
}

const selectedText = (page: Page) => pinnedEvaluate(page, async () => (await window.__miniToolbar.editor!.readSelectionInfo())?.selectedText ?? null);

async function savedRuns(page: Page, text: string): Promise<string[]> {
  const bytes = await pinnedEvaluate(page, () => window.__miniToolbar.saved());
  const xml = await (await JSZip.loadAsync(Uint8Array.from(bytes))).file('word/document.xml')!.async('string');
  return (xml.match(/<w:r[ >](?:(?!<\/w:r>).)*<\/w:r>/gs) ?? []).filter((entry) => entry.includes(text));
}

function overlaps(a: Box, b: Box): boolean {
  return a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height;
}

test('a double click selects exactly the word and the toolbar shows above it', async ({ page }) => {
  await open(page);
  const word = await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  expect((await selectedText(page))?.trim()).toBe('Bravo');
  const box = (await toolbar(page).boundingBox())!;
  expect(box.y + box.height).toBeLessThanOrEqual(word.y);
  await expect(input(page)).toBeFocused();
});

test('a triple click selects the paragraph and the toolbar shows', async ({ page }) => {
  await open(page);
  const word = await wordBox(page, 'echo');
  await page.mouse.click(word.x + word.width / 2, word.y + word.height / 2, { clickCount: 3 });
  await expect(toolbar(page)).toBeVisible();
  expect((await selectedText(page))?.trim()).toBe('Delta echo foxtrot. Golf hotel india.');
});

test('Bold toggles the selected word and is saved', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Charlie');
  const bold = toolbar(page).getByRole('button', { name: 'Bold' });
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await bold.click();
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expect(input(page)).toBeFocused();
  await expect.poll(async () => (await savedRuns(page, 'Charlie')).at(-1) ?? '').toContain('<w:b/>');
  expect((await savedRuns(page, 'Alpha')).at(-1)).not.toContain('<w:b/>');
});

test('Clear formatting removes the direct formatting', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Heavy');
  await expect(toolbar(page).getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'true');
  await toolbar(page).getByRole('button', { name: 'Clear formatting' }).click();
  await expect(toolbar(page).getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'false');
  await expect.poll(async () => (await savedRuns(page, 'Heavy')).at(-1) ?? '<w:b/>').not.toContain('<w:b/>');
});

test('Paste as plain text replaces the selection with the clipboard text and no marks', async ({ page }) => {
  await open(page);
  await page.evaluate(() => navigator.clipboard.write([new ClipboardItem({
    'text/plain': new Blob(['PastedX'], { type: 'text/plain' }),
    'text/html': new Blob(['<p><b>PastedX</b></p>'], { type: 'text/html' }),
  })]));
  await doubleClick(page, 'Alpha');
  await toolbar(page).getByRole('button', { name: 'Paste as Plain Text' }).click();
  await expect.poll(async () => (await savedRuns(page, 'PastedX')).length).toBe(1);
  expect((await savedRuns(page, 'PastedX'))[0]).not.toContain('<w:b/>');
  expect(await savedRuns(page, 'Alpha')).toEqual([]);
});

test('typing hides the toolbar', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  await page.keyboard.type('Z');
  await expect(toolbar(page)).toHaveCount(0);
});

test('Escape in the toolbar hides it and returns focus to the document with the selection kept', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  const selected = await selectedText(page);
  await page.keyboard.press('Alt+F10');
  await expect(toolbar(page).getByRole('button', { name: 'Bold' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(toolbar(page)).toHaveCount(0);
  await expect(input(page)).toBeFocused();
  expect(await selectedText(page)).toBe(selected);
});

test('Alt+F10 moves focus into the toolbar, arrows rove and Tab returns it to the document', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  const selected = await selectedText(page);
  await page.keyboard.press('Alt+F10');
  const bold = toolbar(page).getByRole('button', { name: 'Bold' });
  await expect(bold).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(toolbar(page).getByRole('button', { name: 'Italic' })).toBeFocused();
  await page.keyboard.press('End');
  await expect(toolbar(page).getByRole('button', { name: 'Paste as Plain Text' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(input(page)).toBeFocused();
  expect(await selectedText(page)).toBe(selected);
  await expect(toolbar(page)).toBeVisible();
});

for (const word of ['Edge', 'Middle']) {
  test(`the toolbar and the comment button do not overlap (${word})`, async ({ page }) => {
    await open(page);
    await doubleClick(page, word);
    await expect(toolbar(page)).toBeVisible();
    await expect(commentButton(page)).toBeVisible();
    const bar = (await toolbar(page).boundingBox())!;
    const comment = (await commentButton(page).boundingBox())!;
    expect(overlaps(bar, comment), JSON.stringify({ bar, comment })).toBe(false);
  });
}

test('an open context menu hides the toolbar', async ({ page }) => {
  await open(page);
  const word = await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  await input(page).dispatchEvent('contextmenu', { button: 2, clientX: word.x + word.width / 2, clientY: word.y + word.height / 2 });
  await expect(page.getByRole('menu')).toBeVisible();
  await expect(toolbar(page)).toHaveCount(0);
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu')).toHaveCount(0);
});

test('a mouse press on an editor toolbar menu hides the toolbar until a new selection', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  const selected = await selectedText(page);
  const trigger = page.getByRole('button', { name: 'Line spacing', exact: true });
  await trigger.click();
  await expect(page.getByRole('menu', { name: 'Line spacing' })).toBeVisible();
  await expect(input(page)).toBeFocused();
  await expect(toolbar(page)).toHaveCount(0);
  await trigger.click();
  await expect(page.getByRole('menu', { name: 'Line spacing' })).toHaveCount(0);
  await expect(input(page)).toBeFocused();
  expect(await selectedText(page)).toBe(selected);
  await page.waitForTimeout(1_000);
  await expect(toolbar(page)).toHaveCount(0);
  await doubleClick(page, 'Charlie');
  await expect(toolbar(page)).toBeVisible();
  expect((await selectedText(page))?.trim()).toBe('Charlie');
});

test('a mouse press inside the toolbar keeps it', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  const box = (await toolbar(page).boundingBox())!;
  await page.mouse.click(box.x + 2, box.y + 2);
  const italic = toolbar(page).getByRole('button', { name: 'Italic' });
  await italic.click();
  await expect(italic).toHaveAttribute('aria-pressed', 'true');
  await page.waitForTimeout(1_000);
  await expect(toolbar(page)).toBeVisible();
  await expect(input(page)).toBeFocused();
});

test('an editor toolbar menu opened from the keyboard hides the toolbar', async ({ page }) => {
  await open(page);
  await doubleClick(page, 'Bravo');
  await expect(toolbar(page)).toBeVisible();
  const trigger = page.getByRole('button', { name: 'Line spacing', exact: true });
  await trigger.focus();
  await page.keyboard.press('ArrowDown');
  await expect(page.getByRole('menu', { name: 'Line spacing' })).toBeVisible();
  await expect(input(page)).not.toBeFocused();
  await expect(toolbar(page)).toHaveCount(0);
});

for (const query of ['?readonly=1', '?mode=viewing']) {
  test(`never shows in ${query.slice(1)}`, async ({ page }) => {
    await open(page, query);
    await doubleClick(page, 'Bravo');
    await expect.poll(async () => (await selectedText(page))?.trim()).toBe('Bravo');
    await page.waitForTimeout(1_000);
    await expect(toolbar(page)).toHaveCount(0);
  });
}
