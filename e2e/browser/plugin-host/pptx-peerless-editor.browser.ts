import { test, expect, type Page } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import type {} from './pptx-peerless-editor-probe';

test.use({ launchOptions: { executablePath: process.env.CHROME || undefined } });

const pageErrors = new WeakMap<Page, string[]>();
const editor = (page: Page, index = 0) => page.locator(`[data-peerless-editor="${index}"]`);
const input = (page: Page, index = 0) => editor(page, index).getByRole('textbox', { name: 'Slide text' });

async function open(page: Page, query = 'asyncApi=1') {
  await page.goto(`/pptx-worker-editor.html?${query}`);
  await page.waitForFunction(() => '__pptxWorkerEditor' in window);
  await expect.poll(() => page.evaluate(() => '__pptxPeerlessEditor' in window), { timeout: 5000 }).toBe(true);
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.ready);
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial');
}

async function content(page: Page, index = 0) {
  return pinnedEvaluate(page, (index) => window.__pptxPeerlessEditor.content(index), index);
}

async function text(page: Page, index = 0) {
  return (await content(page, index)).stories[0]!.text;
}

async function select(page: Page, start: number, end = start, index = 0) {
  expect(await pinnedEvaluate(page, ({ start, end, index }) => window.__pptxPeerlessEditor.select(start, end, index),
    { start, end, index })).toBe(true);
  await input(page, index).focus();
}

async function replace(page: Page, value: string, index = 0) {
  await select(page, 0, (await text(page, index)).length, index);
  await input(page, index).pressSequentially(value);
  await pinnedEvaluate(page, (index) => window.__pptxPeerlessEditor.flush(index), index);
  await expect.poll(() => text(page, index)).toBe(value);
}

test.beforeEach(async ({ page }) => {
  const errors: string[] = [];
  pageErrors.set(page, errors);
  page.on('pageerror', (error) => errors.push(error.message));
});

test.afterEach(async ({ page }) => {
  expect.soft(pageErrors.get(page), 'Page errors').toEqual([]);
  if (!new URL(page.url()).searchParams.has('asyncApi')) return;
  expect.soft(await page.evaluate(() => window.__pptxMainEngineProbe?.count()), 'Main-thread engines').toBe(0);
  if (await page.evaluate(() => '__pptxPeerlessEditor' in window)) {
    expect.soft(await page.evaluate(() => window.__pptxPeerlessEditor.errors), 'Editor errors').toEqual([]);
  }
});

test('asyncApi mounts a peer-less editor without a main-thread engine', async ({ page }) => {
  await open(page);
  await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
  await expect(page.getByTestId('pptx-slide-canvas')).toBeVisible();
  expect(await page.evaluate(() => window.__pptxPeerlessEditor.hydrated())).toBe(true);
  expect(await page.evaluate(() => window.__pptxPeerlessEditor.mainThreadEngines())).toBe(0);
  expect((await content(page)).slides).toHaveLength(2);
  expect(await text(page)).toBe('initial');
});

test('hybrid control still creates a main-thread engine', async ({ page }) => {
  await page.goto('/pptx-worker-editor.html');
  await page.waitForFunction(() => '__pptxWorkerEditor' in window);
  await pinnedEvaluate(page, () => window.__pptxWorkerEditor.mounted);
  await pinnedEvaluate(page, () => window.__pptxWorkerEditor.releaseHydration());
  expect(await page.evaluate(() => window.__pptxMainEngineProbe!.count())).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.__pptxWorkerEditor.text())).toBe('initial');
  expect(await page.evaluate(() => window.__pptxWorkerEditor.errors)).toEqual([]);
});

test('peer-less slide buttons and keyboard navigation adopt the active frame', async ({ page }) => {
  await open(page);
  const slides = editor(page).getByRole('complementary').getByRole('button');
  await expect(slides).toHaveCount(2);
  await slides.nth(1).click();
  await expect(slides.nth(1)).toHaveAttribute('aria-current', 'page');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.activeSlide())).toBe(2);
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('other');
  const stage = editor(page).getByRole('application');
  await stage.press('Home');
  await expect(slides.nth(0)).toHaveAttribute('aria-current', 'page');
  await stage.press('PageDown');
  await expect(slides.nth(1)).toHaveAttribute('aria-current', 'page');
  await stage.press('PageUp');
  await expect(slides.nth(0)).toHaveAttribute('aria-current', 'page');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial');
});

test('peer-less pointer typing commits before leaving and returning to the slide', async ({ page }) => {
  await open(page);
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.point(7));
  await page.mouse.click(point.x, point.y);
  await expect(input(page)).toBeFocused();
  await page.keyboard.type(' committed');
  await editor(page).getByRole('complementary').getByRole('button').nth(1).click();
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.activeSlide())).toBe(2);
  await expect.poll(() => text(page)).toBe('initial committed');
  await editor(page).getByRole('application').press('PageUp');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial committed');
});

test('peer-less shape drag precedes an immediate text click and typing', async ({ page }) => {
  await open(page);
  const before = (await content(page)).slides[0]!.shapes[0]!;
  const shape = await page.evaluate(() => window.__pptxPeerlessEditor.shapePoint());
  const point = await page.evaluate(() => window.__pptxPeerlessEditor.point(2));
  await page.evaluate(() => window.__pptxPeerlessEditor.hold());
  await page.mouse.move(shape.x, shape.y);
  await page.mouse.down();
  await page.mouse.move(shape.x + 30, shape.y, { steps: 4 });
  await page.mouse.up();
  await page.mouse.click(point.x, point.y);
  await page.keyboard.type('!');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.held())).toBeGreaterThan(0);
  await page.evaluate(() => window.__pptxPeerlessEditor.release());
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.flush());
  const result = await content(page);
  expect(result.stories[0]!.text).toBe('in!itial');
  expect(result.slides[0]!.shapes[0]!.x).toBe(before.x + Math.round(30 * shape.units.x));
  expect(result.slides[0]!.shapes[0]!.y).toBe(before.y);
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('in!itial');
});

test('peer-less text box insertion focuses input and commits the new story', async ({ page }) => {
  await open(page);
  await editor(page).getByRole('button', { name: 'Text box', exact: true }).click();
  const canvas = await editor(page).getByTestId('pptx-slide-canvas').boundingBox();
  if (!canvas) throw new Error('Slide canvas is missing');
  await page.mouse.move(canvas.x + canvas.width * 0.2, canvas.y + canvas.height * 0.65);
  await page.mouse.down();
  await page.mouse.move(canvas.x + canvas.width * 0.6, canvas.y + canvas.height * 0.85, { steps: 4 });
  await page.mouse.up();
  await expect(input(page)).toBeFocused();
  await page.keyboard.type('inserted text');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.flush());
  const result = await content(page);
  expect(result.slides[0]!.shapes).toHaveLength(2);
  expect(result.stories.map((story) => story.text)).toEqual(['initial', 'inserted text', 'other']);
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('inserted text');
});

test('peer-less slide insertion updates the strip and worker document', async ({ page }) => {
  await open(page);
  await editor(page).getByRole('button', { name: 'Add slide', exact: true }).click();
  const slides = editor(page).getByRole('complementary').getByRole('button');
  await expect(slides).toHaveCount(3);
  expect((await content(page)).slides).toHaveLength(3);
  await expect(slides.nth(1)).toHaveAttribute('aria-current', 'page');
  await slides.nth(2).click();
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('other');
});

test('peer-less undo and redo restore committed text and the painted frame', async ({ page }) => {
  await open(page);
  await replace(page, 'history value');
  await input(page).press('ControlOrMeta+z');
  await expect.poll(() => text(page)).toBe('initial');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial');
  await input(page).press('ControlOrMeta+Shift+z');
  await expect.poll(() => text(page)).toBe('history value');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('history value');
});

test('peer-less paste replaces the selected text exactly once', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await open(page);
  await select(page, 0, 7);
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('pasted value'));
  await input(page).press('ControlOrMeta+v');
  await expect.poll(() => text(page)).toBe('pasted value');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('pasted value');
});

test('peer-less copy and paste transfer selected text between slides', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await open(page);
  await select(page, 0, 7);
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('copy sentinel'));
  await input(page).press('ControlOrMeta+c');
  await expect.poll(() => pinnedEvaluate(page, () => navigator.clipboard.readText())).toBe('initial');
  expect(await text(page)).toBe('initial');
  await editor(page).getByRole('complementary').getByRole('button').nth(1).click();
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.activeSlide())).toBe(2);
  await select(page, 0, 5);
  await input(page).press('ControlOrMeta+v');
  await expect.poll(async () => (await content(page)).stories[1]!.text).toBe('initial');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial');
});

test('peer-less cut copies before removing selected text and undo restores it', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await open(page);
  await select(page, 0, 7);
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('cut sentinel'));
  await input(page).press('ControlOrMeta+x');
  await expect.poll(() => pinnedEvaluate(page, () => navigator.clipboard.readText())).toBe('initial');
  await expect.poll(() => text(page)).toBe('');
  await input(page).press('ControlOrMeta+z');
  await expect.poll(() => text(page)).toBe('initial');
});

test('peer-less composition commits once with replies pending', async ({ page }) => {
  await open(page);
  await select(page, 7);
  await page.evaluate(() => window.__pptxPeerlessEditor.hold());
  await input(page).evaluate((element) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    element.dispatchEvent(new CompositionEvent('compositionupdate', { data: '漢字', bubbles: true }));
    element.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertCompositionText', data: '漢字',
      isComposing: true, bubbles: true, cancelable: true }));
    element.dispatchEvent(new CompositionEvent('compositionend', { data: '漢字', bubbles: true }));
    element.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertFromComposition', data: '漢字',
      bubbles: true, cancelable: true }));
  });
  await page.evaluate(() => window.__pptxPeerlessEditor.release());
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.flush());
  expect(await text(page)).toBe('initial漢字');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('initial漢字');
});

test('peer-less save and reopen preserve edits and accept further typing', async ({ page }) => {
  await open(page);
  await replace(page, 'saved edit');
  const generation = await page.evaluate(() => window.__pptxPeerlessEditor.generation());
  expect(await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.saveAndReopen())).toBeGreaterThan(0);
  expect(await page.evaluate(() => window.__pptxPeerlessEditor.generation())).toBeGreaterThan(generation);
  expect(await text(page)).toBe('saved edit');
  expect((await content(page)).stories[1]!.text).toBe('other');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('saved edit');
  await replace(page, 'reopened edit');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.paintedTexts())).toContain('reopened edit');
});

test('two peer-less editors merge concurrent typing and undo only the local edit', async ({ page }) => {
  await open(page, 'asyncApi=1&collaboration=1');
  await expect(page.getByTestId('pptx-worker-editor')).toHaveCount(2);
  await select(page, 7, 7, 0);
  await input(page, 0).pressSequentially('L');
  await select(page, 7, 7, 1);
  await input(page, 1).pressSequentially('R');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.merge());
  const merged = await text(page, 0);
  expect(merged).toMatch(/^initial(?:LR|RL)$/);
  expect(await text(page, 1)).toBe(merged);
  for (const index of [0, 1]) {
    await expect.poll(() => page.evaluate((index) => window.__pptxPeerlessEditor.paintedTexts(index), index)).toContain(merged);
  }
  await input(page, 0).focus();
  await input(page, 0).press('ControlOrMeta+z');
  await expect.poll(() => text(page, 0)).toBe('initialR');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.merge());
  expect(await text(page, 1)).toBe('initialR');
});

test('peer-less remote merge preserves typing queued behind the update', async ({ page }) => {
  await open(page, 'asyncApi=1&collaboration=1');
  await select(page, 0, 0, 1);
  await input(page, 1).pressSequentially('REMOTE');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.flush(1));
  await select(page, 2, 2, 0);
  await page.evaluate(() => window.__pptxPeerlessEditor.hold(0));
  await input(page, 0).pressSequentially('L');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.queueRemoteMerge());
  await input(page, 0).pressSequentially('NEXT');
  await expect.poll(() => page.evaluate(() => window.__pptxPeerlessEditor.held(0))).toBeGreaterThan(0);
  await page.evaluate(() => window.__pptxPeerlessEditor.release(0));
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.settleRemoteMerge());
  expect(await text(page, 0)).toBe('REMOTEinLNEXTitial');
  await pinnedEvaluate(page, () => window.__pptxPeerlessEditor.merge());
  expect(await text(page, 1)).toBe('REMOTEinLNEXTitial');
});
