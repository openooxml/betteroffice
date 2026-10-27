import { test, expect, type Page } from 'playwright/test';

type ProbeWindow = { __probe: { editor: { readParagraphs(request: { view: 'accepted' }): Promise<{ paragraphs: { text: string }[] }> } | null } };

const documentText = (page: Page) => page.evaluate(async () =>
  (await (window as unknown as ProbeWindow).__probe.editor!.readParagraphs({ view: 'accepted' })).paragraphs.map(({ text }) => text).join('\n'));

/** Holds the engine wasm until the returned release runs, so the editor stays in its opening state. */
async function holdOpen(page: Page) {
  let release!: () => void;
  const gate = new Promise<void>((done) => { release = done; });
  await page.context().route(/docx_edit_bg\.wasm/, async (route) => {
    await gate;
    await route.continue();
  });
  return release;
}

/** Dispatches `type` on the focused element and reports whether the editor cancelled it. */
const dispatchToFocus = (page: Page, type: 'paste' | 'compositionstart') => page.evaluate((type) => {
  const event = type === 'paste'
    ? new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: Object.assign(new DataTransfer(), {}) })
    : new CompositionEvent('compositionstart', { bubbles: true, cancelable: true, data: '' });
  if (event instanceof ClipboardEvent) event.clipboardData?.setData('text/plain', 'zqxj');
  (document.activeElement ?? document.body).dispatchEvent(event);
  return event.defaultPrevented;
}, type);

const kinds: [string, (page: Page) => Promise<void>][] = [
  ['printable keys', (page) => page.keyboard.type('zqxj')],
  ['Enter', (page) => page.keyboard.press('Enter')],
  ['Backspace', (page) => page.keyboard.press('Backspace')],
  ['an AltGr or Option character', (page) => page.keyboard.press('Alt+KeyQ')],
  ['a paste', async (page) => { expect(await dispatchToFocus(page, 'paste')).toBe(true); }],
  ['an IME composition', async (page) => { expect(await dispatchToFocus(page, 'compositionstart')).toBe(true); }],
];

for (const [kind, act] of kinds) {
  for (const press of kind === 'printable keys' ? [true, false] : [true]) {
    test(`${kind} while the document opens${press ? ' after a press' : ''} show that typing is paused and change nothing`, async ({ page }) => {
      const release = await holdOpen(page);
      await page.goto('/?peerless=1');
      const scroller = page.locator('.docx-editor__scroll-container');
      await expect(scroller).toBeVisible({ timeout: 60_000 });
      await expect(page.locator('canvas[data-page-index="0"]')).toHaveCount(0);
      if (press) {
        const box = (await scroller.boundingBox())!;
        await page.mouse.click(box.x + box.width / 2, box.y + 250);
      }
      const paused = page.getByTestId('docx-typing-paused');
      await expect(paused).toHaveCount(0);
      await act(page);
      await expect(paused).toHaveText('This document is still opening. Typing will work once it is ready.');
      await expect(paused).toHaveAttribute('role', 'status');
      release();
      await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
      await expect(paused).toHaveCount(0);
      const opened = await documentText(page);
      await page.keyboard.type('vwk');
      await expect.poll(() => documentText(page), { timeout: 20_000 }).toContain('vwk');
      await expect(paused).toHaveCount(0);
      await page.context().unroute(/docx_edit_bg\.wasm/);
      await page.goto('/?peerless=1');
      await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
      expect(opened).toBe(await documentText(page));
    });
  }
}
