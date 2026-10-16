import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const sample = resolve(root, 'apps/demo/public/betteroffice-demo.docx');
const text = 'Review sample';

async function documentXml(bytes: Buffer) {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

async function saved(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const xml = await documentXml(await readFile(file));
  return page.evaluate(({ xml, text }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const runs = Array.from(document.getElementsByTagNameNS('*', 't')).filter(node => node.textContent?.includes(text));
    const tracked = (node: Element | null): boolean => !!node && (node.localName === 'ins' || tracked(node.parentElement));
    return { insertions: document.getElementsByTagNameNS('*', 'ins').length, found: runs.length > 0, tracked: runs.some(tracked) };
  }, { xml, text });
}

async function suggest(page: Page) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(sample);
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await page.getByTestId('formatting-bar').getByRole('button', { name: 'Editing', exact: true }).click();
  await page.getByRole('menuitemradio', { name: /^Suggesting/ }).click();
  await page.locator('.paged-editor__page canvas').first().click({ position: { x: 150, y: 150 } });
  const input = page.locator('.paged-editor__peerless-input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await input.pressSequentially(text);
  const card = page.locator('.docx-tracked-change-card', { hasText: text });
  await expect(card).toHaveCount(1, { timeout: 15_000 });
  await card.click();
  return card;
}

test.beforeEach(async ({ page }) => { await page.setViewportSize({ width: 2560, height: 1200 }); });

test('accepting a suggestion from its sidebar card saves it untracked', async ({ page }, info) => {
  const before = await page.evaluate(xml => new DOMParser().parseFromString(xml, 'application/xml')
    .getElementsByTagNameNS('*', 'ins').length, await documentXml(await readFile(sample)));
  const card = await suggest(page);
  await card.getByTitle('Accept', { exact: true }).click();
  await expect(card).toHaveCount(0);
  expect(await saved(page, info, 'accepted')).toEqual({ insertions: before, found: true, tracked: false });
});

test('rejecting a suggestion from its sidebar card removes the text', async ({ page }, info) => {
  const before = await page.evaluate(xml => new DOMParser().parseFromString(xml, 'application/xml')
    .getElementsByTagNameNS('*', 'ins').length, await documentXml(await readFile(sample)));
  const card = await suggest(page);
  await card.getByTitle('Reject', { exact: true }).click();
  await expect(card).toHaveCount(0);
  expect(await saved(page, info, 'rejected')).toEqual({ insertions: before, found: false, tracked: false });
});
