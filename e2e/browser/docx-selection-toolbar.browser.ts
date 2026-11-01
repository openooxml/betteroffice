import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const PARAGRAPH = 'Alpha BoldWord omega text for the selection toolbar.';
const SECOND = 'Second paragraph stays.';

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>'
    + '<w:p w14:paraId="00000001"><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">Alpha </w:t></w:r>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:b/></w:rPr><w:t>BoldWord</w:t></w:r>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> omega text for the selection toolbar.</w:t></w:r></w:p>'
    + `<w:p w14:paraId="00000002"><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>${SECOND}</w:t></w:r></w:p>`
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, arm: string) {
  await page.setViewportSize({ width: 1600, height: 1000 });
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'selection-toolbar.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await fixture(),
  });
  const canvas = page.locator('.canvas-page canvas, .paged-editor__page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60000 });
  await canvas.click({ position: { x: 150, y: 150 } });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  return input;
}

async function savedXml(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(name);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  return zip.file('word/document.xml')!.async('string');
}

/** The text of each paragraph, with its runs joined. */
function paragraphTexts(xml: string): string[] {
  return [...xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)]
    .map((paragraph) => [...paragraph[1].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((text) => text[1]).join(''));
}

function toolbar(page: Page) {
  return page.getByRole('toolbar', { name: 'Selection toolbar' });
}

async function caretBox(page: Page) {
  const caret = page.locator('[data-testid="caret"], [data-docx-worker-editor] div[aria-hidden="true"][style*="currentcolor" i]').first();
  await expect(caret).toHaveCount(1);
  const box = await caret.boundingBox();
  if (!box) throw new Error('Caret has no bounds');
  return box;
}

/** Double-clicks the bold word, which starts just after the caret at Home. */
async function selectBoldWord(page: Page) {
  const caret = await caretBox(page);
  await page.mouse.dblclick(caret.x + 55, caret.y + caret.height / 2);
  await expect(toolbar(page)).toBeVisible({ timeout: 15000 });
  return caret;
}

for (const arm of ['&peerless=0', '']) {
  const label = arm ? 'legacy' : 'peerless';

  test(`selection mini toolbar appears above a double-click selection and the clicks still select the word and paragraph (${label})`, async ({ page }, info) => {
    const input = await open(page, arm);
    const caret = await selectBoldWord(page);
    const bar = await toolbar(page).boundingBox();
    expect(bar).not.toBeNull();
    expect(bar!.y + bar!.height).toBeLessThanOrEqual(caret.y + 4);
    await page.keyboard.insertText('Chosen');
    expect(paragraphTexts(await savedXml(page, info, 'word.docx'))[0]).toBe('Alpha Chosen omega text for the selection toolbar.');

    await open(page, arm);
    const start = await caretBox(page);
    await page.mouse.click(start.x + 55, start.y + start.height / 2, { clickCount: 3 });
    await expect(toolbar(page)).toBeVisible({ timeout: 15000 });
    await page.keyboard.insertText('Whole paragraph');
    expect(paragraphTexts(await savedXml(page, info, 'paragraph.docx'))).toEqual(['Whole paragraph', SECOND]);
    await expect(input).toBeEditable();
  });

  test(`Clear formatting in the mini toolbar removes bold and keeps the selection (${label})`, async ({ page }, info) => {
    await open(page, arm);
    await selectBoldWord(page);
    await toolbar(page).getByRole('button', { name: 'Clear formatting' }).click();
    const xml = await savedXml(page, info, 'cleared.docx');
    expect(xml).toContain('BoldWord');
    expect(xml).not.toMatch(/<w:b[\s/>]/);
    const selection = page.locator('[data-testid^="selection-rect-"], [data-docx-worker-editor] div[aria-hidden="true"][style*="rgba(59, 130, 246"]');
    await expect.poll(() => selection.count()).toBeGreaterThan(0);
  });

  test(`Paste as plain text in the mini toolbar inserts unformatted clipboard text (${label})`, async ({ page, context }, info) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    const input = await open(page, arm);
    await page.evaluate(() => navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob(['<b style="font-family: Georgia">Plain pasted text</b>'], { type: 'text/html' }),
      'text/plain': new Blob(['Plain pasted text'], { type: 'text/plain' }),
    })]));
    for (let step = 0; step < 5; step++) await input.press('Shift+ArrowRight');
    await expect(toolbar(page)).toBeVisible({ timeout: 15000 });
    await toolbar(page).getByRole('button', { name: 'Paste as Plain Text' }).click();
    const xml = await savedXml(page, info, 'pasted.docx');
    expect(paragraphTexts(xml)[0]).toBe('Plain pasted text BoldWord omega text for the selection toolbar.');
    const pastedRun = [...xml.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)].map((run) => run[1]).find((run) => run.includes('Plain pasted text'));
    expect(pastedRun).toBeDefined();
    expect(pastedRun).not.toMatch(/<w:b[\s/>]/);
    expect(pastedRun).not.toContain('Georgia');
  });

  test(`Alt+F10 moves focus into the mini toolbar and Escape returns it to the document (${label})`, async ({ page }) => {
    const input = await open(page, arm);
    await selectBoldWord(page);
    const buttons = toolbar(page).getByRole('button');
    await page.keyboard.press('Alt+F10');
    await expect(buttons.first()).toBeFocused();
    await page.keyboard.press('ArrowRight');
    await expect(buttons.nth(1)).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(toolbar(page)).toBeHidden();
    await expect(input).toBeFocused();
  });

  test(`typing and Escape hide the mini toolbar (${label})`, async ({ page }) => {
    const input = await open(page, arm);
    await selectBoldWord(page);
    await page.keyboard.type('z');
    await expect(toolbar(page)).toBeHidden();
    await selectBoldWord(page);
    await input.press('Escape');
    await expect(toolbar(page)).toBeHidden();
  });
}
