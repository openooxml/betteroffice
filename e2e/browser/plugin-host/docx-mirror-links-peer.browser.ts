import { expect, test, type Page } from 'playwright/test';
import type {} from './docx-peerless-collab-probe';

const editor = (page: Page, index: number) => page.locator(`[data-editor="${index}"]`);

async function open(page: Page) {
  await page.addInitScript(() => {
    const state = window as typeof window & { opened?: unknown[][] };
    state.opened = [];
    window.open = (...args) => { state.opened!.push(args); return null; };
  });
  await page.goto('/docx-peerless-collab.html');
  await page.waitForFunction(() => '__docxPeerlessCollab' in window);
  await page.evaluate(() => window.__docxPeerlessCollab.ready);
  await expect.poll(() => page.evaluate(() => [0, 1].every(index => window.__docxPeerlessCollab.synced(index)))).toBe(true);
  await expect(editor(page, 0).locator('canvas').first()).toBeVisible();
}

async function linkFromPeer(page: Page, url: string) {
  const input = editor(page, 1).locator('textarea.paged-editor__input');
  await input.focus();
  await input.press('ControlOrMeta+Home');
  await input.press('Shift+End');
  await input.press('ControlOrMeta+k');
  const dialog = page.getByRole('dialog', { name: 'Insert Hyperlink', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('URL', { exact: true }).fill(url);
  const insert = dialog.getByRole('button', { name: 'Insert', exact: true });
  if (await insert.isEnabled()) await insert.click();
  if (!await dialog.waitFor({ state: 'hidden', timeout: 3000 }).then(() => true, () => false)) {
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  }
  await expect(dialog).toHaveCount(0);
}

const peerMirrorHrefs = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll('[data-editor="0"] .canvas-page-mirror a[href], [data-editor="0"] .canvas-page-mirror [data-href]'),
    element => [element.getAttribute('href'), element.getAttribute('data-href')]).flat().filter((value): value is string => value !== null));

test.afterEach(async ({ page }) => {
  if (await page.evaluate(() => '__docxPeerlessCollab' in window)) expect(await page.evaluate(() => window.__docxPeerlessCollab.errors)).toEqual([]);
});

test('a link a peer sets reaches the mirror and opens without an opener', async ({ page }) => {
  await open(page);
  await linkFromPeer(page, 'https://example.com/peer');
  const anchor = editor(page, 0).locator('.canvas-page-mirror a[href="https://example.com/peer"]').first();
  await expect(anchor).toBeAttached();
  await anchor.focus();
  await page.keyboard.press('Enter');
  await expect.poll(() => page.evaluate(() => (window as typeof window & { opened?: unknown[][] }).opened))
    .toContainEqual(['https://example.com/peer', '_blank', 'noopener,noreferrer']);
});

test('a hostile link a peer sets never reaches the mirror as a link', async ({ page }) => {
  await open(page);
  await linkFromPeer(page, 'javascript:window.__pwned=1;void 0');
  await expect.poll(async () => {
    const [a, b] = await page.evaluate(() => Promise.all([0, 1].map(index => window.__docxPeerlessCollab.stateVector(index))));
    return JSON.stringify(a) === JSON.stringify(b);
  }).toBe(true);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessCollab.painted(0))).toContain('Seed');
  for (const href of await peerMirrorHrefs(page)) expect(href).not.toMatch(/script|data:/i);
  expect(await page.evaluate(() => (window as typeof window & { __pwned?: number }).__pwned ?? 0)).toBe(0);
});
