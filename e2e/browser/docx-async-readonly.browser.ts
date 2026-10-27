import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { pinnedEvaluate } from './pinned-evaluate';

const errors = new WeakMap<Page, string[]>();
const docxInput = (page: Page) => page.locator('[data-docx-worker-editor] textarea');
const docxCanvas = (page: Page) => page.locator('[data-docx-worker-editor] [data-page-id] canvas').first();

async function save(page: Page): Promise<Buffer> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const stream = await (await downloading).createReadStream();
  if (!stream) throw new Error('Save produced no document');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks);
}

async function parts(bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const result: Record<string, string> = {};
  for (const name of Object.keys(zip.files).filter((name) => !zip.files[name]!.dir).sort()) {
    const text = await zip.file(name)!.async('string');
    result[name] = name === 'docProps/core.xml'
      ? text.replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/g, '$1$2')
      : text;
  }
  return result;
}

async function savedText(page: Page, bytes: Buffer) {
  const source = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
  return page.evaluate((source) => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Saved XML is invalid');
    return Array.from(document.getElementsByTagNameNS('*', 't')).map((node) => node.textContent ?? '').join('');
  }, source);
}

async function unchanged(page: Page, baseline: Buffer) {
  const bytes = await save(page);
  expect(await parts(bytes), 'Viewing-mode save must preserve every part').toEqual(await parts(baseline));
  expect(await savedText(page, bytes)).toBe(await savedText(page, baseline));
}

async function clipboard(page: Page, text: string) {
  await pinnedEvaluate(page, (text) => navigator.clipboard.writeText(text), text);
}

async function copied(page: Page, text: string) {
  await expect.poll(() => pinnedEvaluate(page, () => navigator.clipboard.readText())).toBe(text);
}

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:r><w:t>abcd</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>second page</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function openDocx(page: Page, buffer?: Buffer) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'readonly.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: buffer ?? await documentBytes(),
  });
  await expect(docxCanvas(page)).toBeVisible();
  await docxCanvas(page).click({ position: { x: 150, y: 150 } });
  await expect(docxInput(page)).toBeEditable();
  await docxInput(page).press('ControlOrMeta+Home');
  await expect(page.locator('.canvas-page-mirror').first()).toContainText('abcd');
}

async function docxMode(page: Page, viewing = true) {
  const controls = page.getByTestId('peerless-harness-mode');
  const next = viewing ? 'Viewing' : 'Editing';
  if (await controls.isVisible()) {
    const button = controls.getByRole('button', { name: next, exact: true });
    await button.click();
    await expect(button).toHaveAttribute('aria-pressed', 'true');
  } else {
    const current = page.getByRole('button', { name: viewing ? 'Editing' : 'Viewing', exact: true });
    if (await current.isVisible()) await current.click();
    else {
      await page.getByRole('button', { name: 'More actions', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Editing mode', exact: true }).hover();
    }
    await page.getByRole('menuitemradio', { name: new RegExp(`^${next}\\b`) }).click();
  }
  if (viewing) await expect(docxInput(page)).not.toBeEditable();
  else await expect(docxInput(page)).toBeEditable();
}

test.beforeEach(async ({ page, context }) => {
  const messages: string[] = [];
  errors.set(page, messages);
  page.on('pageerror', (error) => messages.push(error.message));
  page.on('console', (message) => { if (message.type() === 'error') messages.push(message.text()); });
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
});

test.afterEach(async ({ page }, testInfo) => {
  const messages = errors.get(page) ?? [];
  errors.delete(page);
  if (testInfo.status === 'skipped') return;
  expect.soft(messages, 'Console and page errors').toEqual([]);
});

test('docx viewing refuses typing, paste, cut and deletion', async ({ page }) => {
  await openDocx(page);
  const baseline = await save(page);
  await docxMode(page);
  await docxInput(page).press('ControlOrMeta+Home');
  await docxInput(page).press('Shift+ArrowRight');
  await clipboard(page, 'blocked paste');
  for (const key of ['x', 'ControlOrMeta+v', 'ControlOrMeta+x', 'Delete', 'Backspace']) {
    await docxInput(page).press(key);
    await expect(page.locator('.canvas-page-mirror').first()).toHaveText('abcd');
    await unchanged(page, baseline);
  }
});

test('docx viewing refuses formatting shortcuts and toolbar mutations', async ({ page }) => {
  await openDocx(page);
  const baseline = await save(page);
  await docxInput(page).press('Shift+End');
  await docxMode(page);
  for (const key of ['ControlOrMeta+b', 'ControlOrMeta+i', 'ControlOrMeta+u']) {
    await docxInput(page).press(key);
    await unchanged(page, baseline);
  }
  for (const name of ['Bold', 'Italic', 'Underline']) {
    const button = page.getByRole('button', { name, exact: true });
    if (await button.count()) {
      await expect(button).toBeDisabled();
      await button.click({ force: true });
    } else await expect(button).toHaveCount(0);
    await unchanged(page, baseline);
  }
  await expect(page.locator('.canvas-page-mirror').first()).toHaveText('abcd');
});

test('docx viewing preserves keyboard selection, copy, scrolling and zoom', async ({ page }) => {
  await openDocx(page);
  const baseline = await save(page);
  await docxMode(page);
  await docxInput(page).press('ControlOrMeta+Home');
  await docxInput(page).press('ArrowRight');
  await docxInput(page).press('Shift+ArrowRight');
  await clipboard(page, 'copy sentinel');
  await docxInput(page).press('ControlOrMeta+c');
  await copied(page, 'b');
  const scroller = page.locator('.editor-stage .docx-editor__scroll-container');
  await scroller.hover({ position: { x: 300, y: 300 } });
  await page.mouse.wheel(0, 600);
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await unchanged(page, baseline);
  const width = (await docxCanvas(page).boundingBox())!.width;
  const zoom = page.getByRole('button', { name: 'Zoom out', exact: true });
  await expect(zoom).toBeEnabled();
  await zoom.click();
  await expect.poll(async () => (await docxCanvas(page).boundingBox())!.width).toBeLessThan(width);
  await unchanged(page, baseline);
});

test('docx edits survive viewing toggles, history, save and reopen', async ({ page }) => {
  await openDocx(page);
  await page.keyboard.insertText('retained ');
  await expect(page.locator('.canvas-page-mirror').first()).toHaveText('retained abcd');
  const baseline = await save(page);
  expect(await savedText(page, baseline)).toBe('retained abcdsecond page');
  await docxMode(page);
  await docxInput(page).press('ControlOrMeta+z');
  await unchanged(page, baseline);
  await docxInput(page).press('ControlOrMeta+Shift+z');
  await unchanged(page, baseline);
  await docxMode(page, false);
  await docxInput(page).press('ControlOrMeta+Home');
  await clipboard(page, 'again ');
  await docxInput(page).press('ControlOrMeta+v');
  await expect(page.locator('.canvas-page-mirror').first()).toHaveText('again retained abcd');
  await docxInput(page).press('ControlOrMeta+z');
  await expect(page.locator('.canvas-page-mirror').first()).toHaveText('retained abcd');
  expect(await savedText(page, await save(page))).toBe('retained abcdsecond page');
  await docxInput(page).press('ControlOrMeta+Shift+z');
  await expect(page.locator('.canvas-page-mirror').first()).toHaveText('again retained abcd');
  const edited = await save(page);
  expect(await savedText(page, edited)).toBe('again retained abcdsecond page');
  await docxMode(page);
  await unchanged(page, edited);
  await openDocx(page, edited);
  await expect(page.locator('.canvas-page-mirror').first()).toHaveText('again retained abcd');
  expect(await savedText(page, await save(page))).toBe('again retained abcdsecond page');
});
