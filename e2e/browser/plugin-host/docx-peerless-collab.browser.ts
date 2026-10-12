import { expect, test, type Locator, type Page } from 'playwright/test';
import type {} from './docx-peerless-collab-probe';

const editor = (page: Page, index: number) => page.locator(`[data-editor="${index}"]`);
const input = (page: Page, index: number) => {
  const worker = editor(page, index).locator('textarea.paged-editor__peerless-input');
  return index === 1 ? worker.or(page.locator('textarea[data-testid="yrs-input"]')).last() : worker;
};
const text = (page: Page, index: number) => page.evaluate(index => window.__docxPeerlessCollab.text(index), index);
async function open(page: Page, mixed = false) {
  await page.goto(`/docx-peerless-collab.html${mixed ? '?mixed=1' : ''}`);
  await page.waitForFunction(() => '__docxPeerlessCollab' in window);
  await page.evaluate(() => window.__docxPeerlessCollab.ready);
  await expect.poll(() => page.evaluate(() => [0, 1].every(index => window.__docxPeerlessCollab.synced(index)))).toBe(true);
  await expect(editor(page, 0).locator('canvas').first()).toBeVisible();
}
async function select(page: Page, index: number, end = true) {
  const target = input(page, index);
  await target.focus(); await target.press(end ? 'ControlOrMeta+End' : 'ControlOrMeta+Home');
  return target;
}
async function type(target: Locator, value: string, withoutFocus = false) {
  if (!withoutFocus) { await target.pressSequentially(value); return; }
  await target.evaluate((element, value) => element.dispatchEvent(new InputEvent('beforeinput', {
    bubbles: true, cancelable: true, inputType: 'insertText', data: value,
  })), value);
}
async function converged(page: Page, expected?: string) {
  const value = expected ?? await text(page, 0);
  await expect.poll(async () => { const a = await text(page, 0); return a === await text(page, 1) ? a : null; }).toBe(value);
  for (const index of [0, 1]) await expect.poll(() => page.evaluate(index => window.__docxPeerlessCollab.painted(index), index)).toContain(value);
}

test.afterEach(async ({ page }) => {
  if (await page.evaluate(() => '__docxPeerlessCollab' in window)) expect(await page.evaluate(() => window.__docxPeerlessCollab.errors)).toEqual([]);
});

test('concurrent typing converges, paints and undo removes only the local text', async ({ page }) => {
  await open(page);
  await page.evaluate(() => { window.__docxPeerlessCollab.hold(0); window.__docxPeerlessCollab.hold(1); });
  await type(await select(page, 0), 'A'); await type(await select(page, 1), 'B');
  await expect.poll(() => text(page, 0)).toBe('SeedA'); await expect.poll(() => text(page, 1)).toBe('SeedB');
  await page.evaluate(() => { window.__docxPeerlessCollab.release(0); window.__docxPeerlessCollab.release(1); });
  await expect.poll(async () => (await text(page, 0)).length).toBe(6);
  await converged(page);
  await input(page, 0).press('ControlOrMeta+z'); await converged(page, 'SeedB');
});

test('remote caret paints at the typed position and rebases after an insertion before it', async ({ page }) => {
  await open(page);
  await type(await select(page, 0), 'A'); await converged(page, 'SeedA');
  const caret = editor(page, 1).getByTestId('remote-presence-caret');
  await expect(caret).toBeVisible(); const before = await caret.boundingBox();
  await type(await select(page, 1, false), 'prefix'); await converged(page, 'prefixSeedA');
  await expect.poll(async () => (await caret.boundingBox())?.x ?? 0).toBeGreaterThan(before!.x);
});

test('disconnected edits converge after reconnect', async ({ page }) => {
  await open(page); await page.evaluate(() => window.__docxPeerlessCollab.disconnect(1));
  await type(await select(page, 0), 'A'); await type(await select(page, 1), 'B');
  await expect.poll(() => text(page, 0)).toBe('SeedA'); await expect.poll(() => text(page, 1)).toBe('SeedB');
  await page.evaluate(() => window.__docxPeerlessCollab.reconnect(1));
  await expect.poll(async () => (await text(page, 0)).length).toBe(6); await converged(page);
});

test('room change reopens with the shared state and document swap stays consistent', async ({ page }) => {
  await open(page); await type(await select(page, 0), 'A'); await converged(page, 'SeedA');
  const before = await page.evaluate(() => [0, 1].map(index => window.__docxPeerlessCollab.generation(index)));
  await page.evaluate(() => window.__docxPeerlessCollab.room()); await converged(page, 'SeedA');
  expect(await page.evaluate(() => [0, 1].map(index => window.__docxPeerlessCollab.generation(index)))).toEqual(before.map(n => n + 1));
  await page.evaluate(() => window.__docxPeerlessCollab.swap()); await converged(page, 'Replacement');
});

test('composition survives a remote merge and commits at its rebased position', async ({ page }) => {
  await open(page); await select(page, 1, false); const target = await select(page, 0);
  await target.evaluate(element => {
    const textarea = element as HTMLTextAreaElement;
    textarea.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
    textarea.value = 'かな';
    textarea.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: 'かな' }));
    textarea.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: 'かな', isComposing: true }));
    textarea.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: 'かな', isComposing: true }));
  });
  await type(input(page, 1), 'R', true);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessCollab.painted(0))).toContain('RSeed');
  await expect(target).toHaveValue('かな');
  await target.evaluate(element => element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '漢字' })));
  await converged(page, 'RSeed漢字');
});

test('worker and hybrid replicas converge and paint each other cursors', async ({ page }) => {
  await open(page, true);
  await type(await select(page, 0), 'A'); await converged(page, 'SeedA');
  await expect(editor(page, 1).getByTestId('remote-presence-overlay').locator(':scope > div').first()).toBeVisible();
  await type(await select(page, 1), 'B'); await converged(page, 'SeedAB');
  await expect(editor(page, 0).getByTestId('remote-presence-caret')).toBeVisible();
});
