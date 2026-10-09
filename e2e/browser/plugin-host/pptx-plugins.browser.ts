import { test, expect, type Page } from 'playwright/test';

interface Probe {
  editor: {
    commands: { execute(id: string, args: unknown): Promise<unknown> };
    readContent(): Promise<{
      version: string;
      slides: { id: string; notes?: string; shapes: { id: string; x: number; y: number }[] }[];
      stories: { slideId: string; shapeId: string; storyId: string; text: string }[];
    }>;
    applyEdits(request: unknown): Promise<unknown>;
    selectText(target: {
      slide: number;
      shapeId: string;
      storyId: string;
      start: number;
      end: number;
    }): boolean;
  } | null;
  navigation: unknown;
  selection: { target: { kind: string } } | null;
  marks: number;
}

const probe = (page: Page) =>
  page.evaluate(() => {
    const value = (window as unknown as { __probe: Probe }).__probe;
    return { navigation: value.navigation, selection: value.selection, marks: value.marks };
  });

async function open(page: Page) {
  await page.goto('/pptx.html');
  await expect(page.getByTestId('pptx-slide-canvas')).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('[data-probe-slide="1"]')).toBeAttached({ timeout: 60_000 });
}

async function zoom(page: Page, scale: number | 'fit') {
  await page.evaluate(
    (value) =>
      (window as unknown as { __probe: Probe }).__probe.editor!.commands.execute('zoom', {
        scale: value,
      }),
    scale
  );
}

/** How far the overlay boxes are from where the canvas shows the slide and a shape on it. */
async function misalignment(page: Page): Promise<number> {
  return page.evaluate(async () => {
    const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
    const canvas = document
      .querySelector('[data-testid="pptx-slide-canvas"]')!
      .getBoundingClientRect();
    const slide = document.querySelector('[data-probe-slide]')?.getBoundingClientRect();
    const marker = document.querySelector<HTMLElement>('[data-probe-shape]');
    if (!slide || !marker) return Number.POSITIVE_INFINITY;
    const read = await editor.readContent();
    const shape = read.slides
      .flatMap((entry) => entry.shapes)
      .find((entry) => entry.id === marker.dataset.probeShape)!;
    const box = marker.getBoundingClientRect();
    const scale = canvas.width / 1280;
    return Math.max(
      Math.abs(slide.left - canvas.left),
      Math.abs(slide.top - canvas.top),
      Math.abs(slide.width - canvas.width),
      Math.abs(slide.height - canvas.height),
      Math.abs(box.left - (canvas.left + (shape.x / 9525) * scale)),
      Math.abs(box.top - (canvas.top + (shape.y / 9525) * scale))
    );
  });
}

test('pptx: overlays land on the presented slide at every zoom and while scrolled', async ({
  page,
}) => {
  await open(page);
  for (const scale of [0.5, 1, 2] as const) {
    await zoom(page, scale);
    await expect(page.locator(`[data-probe-slide][data-zoom="${scale}"]`)).toBeAttached();
    await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
  }
  await page
    .getByTestId('pptx-slide-canvas')
    .evaluate((canvas) => canvas.closest('div[style*="overflow: auto"]')!.scrollBy(120, 80));
  await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
  await zoom(page, 'fit');
  await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
});

test('pptx: overlays leave the pointer to the slide beneath them', async ({ page }) => {
  await open(page);
  const shape = page.locator('[data-probe-shape]');
  await expect(shape).toBeAttached();
  const box = (await shape.boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await expect.poll(async () => (await probe(page)).selection?.target.kind).not.toBe('slide');
});

test('pptx: navigation shows the slide and keeps focus in the panel', async ({ page }) => {
  await open(page);
  const input = page.getByTestId('probe-input');
  await input.focus();
  await page.getByTestId('probe-navigate').focus();
  await page.keyboard.press('Enter');
  await expect.poll(async () => (await probe(page)).navigation).toEqual({ ok: true });
  await expect(page.getByTestId('probe-navigate')).toBeFocused();
  await expect(page.locator('[data-probe-slide="3"]')).toBeAttached();
  await expect
    .poll(async () => {
      const canvas = await page.getByTestId('pptx-slide-canvas').boundingBox();
      const slide = await page.locator('[data-probe-slide="3"]').boundingBox();
      return canvas && slide ? Math.abs(canvas.x - slide.x) + Math.abs(canvas.y - slide.y) : 99;
    })
    .toBeLessThan(1.5);
});

test('pptx: dock tabs work by keyboard, and narrow docks open panels as drawers', async ({
  page,
}) => {
  await open(page);
  const dock = page.getByTestId('plugin-dock-right');
  const probeTab = dock.getByRole('tab', { name: 'Probe' });
  const secondTab = dock.getByRole('tab', { name: 'Second' });
  await probeTab.focus();
  await page.keyboard.press('ArrowRight');
  await expect(secondTab).toBeFocused();
  await expect(secondTab).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('second-panel')).toBeVisible();
  const collapse = dock.getByRole('button', { name: 'Collapse Second' });
  await collapse.focus();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('second-panel')).toBeHidden();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('second-panel')).toBeVisible();

  await page.setViewportSize({ width: 520, height: 900 });
  await expect(page.getByTestId('second-panel')).toBeHidden();
  await secondTab.focus();
  await page.keyboard.press('Enter');
  const drawer = page.getByTestId('plugin-drawer-right');
  await expect(drawer).toBeVisible();
  await expect(drawer).toBeFocused();
  await expect(page.getByTestId('second-panel')).toBeVisible();
  const box = (await drawer.boundingBox())!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  await page.keyboard.press('Escape');
  await expect(drawer).toBeHidden();
  await expect(secondTab).toBeFocused();
});

test('pptx: plugin toolbar commands work by pointer, shortcut and the overflow menu', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1920, height: 1000 });
  await open(page);
  const marks = async () => (await probe(page)).marks;
  const rail = page.getByTestId('pptx-formatting-toolbar');
  await rail.getByRole('button', { name: 'Probe mark' }).click();
  await expect.poll(marks).toBe(1);

  await page.getByRole('application').focus();
  await page.keyboard.press('ControlOrMeta+Shift+K');
  await expect.poll(marks).toBe(2);

  await page.setViewportSize({ width: 480, height: 800 });
  const more = rail.getByTestId('pptx-toolbar-more');
  await expect(more).toBeVisible();
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'More' });
  await expect(menu).toBeVisible();
  await page.keyboard.press('End');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Probe mark');
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
  await expect.poll(marks).toBe(3);
});

test('pptx: plugin text fields keep their keys; built-in shortcuts in plugin chrome act on the editor', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1920, height: 1000 });
  await open(page);
  const state = () =>
    page.evaluate(async () => {
      const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
      const read = await editor.readContent();
      return { notes: read.slides[0].notes ?? '', text: read.stories[0].text };
    });
  await page.evaluate(async () => {
    const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
    const read = await editor.readContent();
    await editor.applyEdits({
      expectVersion: read.version,
      steps: [{ op: 'setSlideNotes', target: { slideId: read.slides[0].id }, text: 'Edited' }],
    });
    const story = read.stories[0];
    editor.selectText({
      slide: 1,
      shapeId: story.shapeId,
      storyId: story.storyId,
      start: 0,
      end: 0,
    });
  });
  const before = await state();
  expect(before.notes).toBe('Edited');

  const overlay = page.getByTestId('overlay-input');
  await overlay.click();
  await page.keyboard.type('xyz');
  await page.keyboard.press('Backspace');
  await expect(overlay).toHaveValue('xy');
  await page.keyboard.press('ControlOrMeta+Z');
  const portal = page.getByTestId('portal-input');
  await portal.click();
  await page.keyboard.type('abc');
  await page.keyboard.press('Backspace');
  await expect(portal).toHaveValue('ab');
  await page.keyboard.press('ControlOrMeta+Z');
  expect(await state()).toEqual(before);

  const notes = async () => (await state()).notes;
  await page.getByTestId('probe-navigate').focus();
  await page.keyboard.press('ControlOrMeta+Z');
  await expect.poll(notes).not.toBe('Edited');
  await page.getByTestId('plugin-dock-right').getByRole('tab', { name: 'Probe' }).focus();
  await page.keyboard.press('ControlOrMeta+Shift+Z');
  await expect.poll(notes).toBe('Edited');
  await page
    .getByTestId('pptx-formatting-toolbar')
    .getByRole('button', { name: 'Probe mark' })
    .focus();
  await page.keyboard.press('ControlOrMeta+Z');
  await expect.poll(notes).not.toBe('Edited');

  await page.getByTestId('probe-navigate').focus();
  await page.keyboard.press('ControlOrMeta+Shift+K');
  await expect.poll(async () => (await probe(page)).marks).toBe(1);
});
