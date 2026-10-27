import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const sample = resolve(root, 'apps/demo/public/betteroffice-demo.docx');
const text = 'Review sample';
const title = 'Welcome to BetterOffice';

async function documentXml(bytes: Buffer) {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

async function counts(page: Page, xml: string, texts: string[]) {
  return page.evaluate(({ xml, texts }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const runs = [...document.getElementsByTagNameNS('*', 't'), ...document.getElementsByTagNameNS('*', 'delText')];
    const tracked = (node: Element | null): boolean => !!node && (node.localName === 'ins' || node.localName === 'del' || tracked(node.parentElement));
    return {
      insertions: document.getElementsByTagNameNS('*', 'ins').length,
      deletions: document.getElementsByTagNameNS('*', 'del').length,
      texts: Object.fromEntries(texts.map(text => {
        const matching = runs.filter(node => node.textContent?.includes(text));
        return [text, { found: matching.length > 0, tracked: matching.some(tracked) }];
      })),
    };
  }, { xml, texts });
}

async function saved(page: Page, info: TestInfo, name: string, texts: string[]) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  return counts(page, await documentXml(await readFile(file)), texts);
}

async function suggest(page: Page, replace = false) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(sample);
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await page.getByTestId('formatting-bar').getByRole('button', { name: 'Editing', exact: true }).click();
  await page.getByRole('menuitemradio', { name: /^Suggesting/ }).click();
  await page.locator('.paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
  const input = page.locator('.paged-editor__input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  if (replace) await input.press('Shift+End');
  await input.pressSequentially(text);
  const card = page.locator('.docx-tracked-change-card');
  await expect(card).toHaveCount(1, { timeout: 15_000 });
  await card.click();
  return card;
}

test.beforeEach(async ({ page }) => { await page.setViewportSize({ width: 2560, height: 1200 }); });

test('repeated Delete advances past retained tracked deletions', async ({ page }, info) => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Alpha beta gamma</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'tracked.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }),
  });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  await page.getByTestId('formatting-bar').getByRole('button', { name: 'Editing', exact: true }).click();
  await page.getByRole('menuitemradio', { name: /^Suggesting/ }).click();
  await page.locator('.paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
  const input = page.locator('.paged-editor__input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  for (let index = 0; index < 3; index++) await input.press('ArrowRight');
  const deletions = page.locator('.canvas-page-mirror').first().locator('.docx-deletion');
  await input.press('Delete');
  await expect.poll(async () => (await deletions.allTextContents()).join('')).toBe('h');
  await input.press('Delete');
  await expect.poll(async () => (await deletions.allTextContents()).join('')).toBe('ha');
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('forward-deleted.docx');
  await (await downloading).saveAs(file);
  const xml = await documentXml(await readFile(file));
  const state = await page.evaluate(xml => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    return {
      text: [...document.getElementsByTagNameNS('*', 't')].map(node => node.textContent ?? '').join(''),
      deleted: [...document.getElementsByTagNameNS('*', 'delText')].map(node => node.textContent ?? '').join(''),
    };
  }, xml);
  expect(state).toEqual({ text: 'Alp beta gamma', deleted: 'ha' });
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(file);
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect.poll(async () => (await deletions.allTextContents()).join('')).toBe('ha');
});

const cases = [
  { name: 'accepting a suggestion', replace: false, decision: 'Accept', texts: { [text]: { found: true, tracked: false } } },
  { name: 'rejecting a suggestion', replace: false, decision: 'Reject', texts: { [text]: { found: false, tracked: false } } },
  { name: 'accepting a replacement', replace: true, decision: 'Accept',
    texts: { [text]: { found: true, tracked: false }, [title]: { found: false, tracked: false } } },
  { name: 'rejecting a replacement', replace: true, decision: 'Reject',
    texts: { [text]: { found: false, tracked: false }, [title]: { found: true, tracked: false } } },
] as const;

for (const { name, replace, decision, texts } of cases) {
  test(`${name} from its sidebar card saves no tracked insertion or deletion`, async ({ page }, info) => {
    const before = await counts(page, await documentXml(await readFile(sample)), []);
    const card = await suggest(page, replace);
    await card.getByTitle(`${decision} change`, { exact: true }).click();
    await expect(card).toHaveCount(0);
    expect(await saved(page, info, `${decision}-${replace ? 'replacement' : 'insertion'}`, Object.keys(texts)))
      .toEqual({ insertions: before.insertions, deletions: before.deletions, texts });
  });
}
