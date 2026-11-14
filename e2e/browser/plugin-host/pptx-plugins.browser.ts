import { test, expect, type Page } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';
import { renderedBox } from '../rendered-box';

type ProbeRect = { x: number; y: number; width: number; height: number };
type ProbeAnchorResult =
  | { ok: true; rects: readonly ProbeRect[]; anchor: ProbeRect; pageRect: ProbeRect; version: string; layoutId: string }
  | { ok: false; failure: { code: string; message: string } };

type ProbeTarget = { kind: 'selection' } | { kind: 'range'; version: string; slideId: string;
  shapeId: string; storyId: string; anchor: number; focus: number };

interface Probe {
  mutation: unknown;
  runMutation(): Promise<unknown>;
  editor: {
    commands: {
      execute(id: string, args: unknown): Promise<unknown>;
      getState(id: string): { active?: boolean | 'mixed' };
    };
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
    }): Promise<boolean>;
  } | null;
  navigation: unknown;
  selection: { target: { kind: string } } | null;
  geometry: {
    layout: { zoom: number; version: string; width: number; height: number };
    getAnchorGeometry(target: ProbeTarget): ProbeAnchorResult;
    readAnchorGeometries(targets: ProbeTarget[]): Promise<ProbeAnchorResult[]>;
    getShapeRect(shapeId: string): ProbeRect | null;
  } | null;
  marks: number;
}

const probe = (page: Page) =>
  page.evaluate(() => {
    const value = (window as unknown as { __probe: Probe }).__probe;
    return { navigation: value.navigation, selection: value.selection, marks: value.marks };
  });

async function open(page: Page, engine: 'legacy' | 'worker') {
  await page.goto(`/pptx.html?engine=${engine}`);
  await expect(page.getByTestId('pptx-slide-canvas')).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('[data-probe-slide="1"]')).toBeAttached({ timeout: 60_000 });
}

async function zoom(page: Page, scale: number | 'fit') {
  await pinnedEvaluate(page,
    (value) =>
      (window as unknown as { __probe: Probe }).__probe.editor!.commands.execute('zoom', {
        scale: value,
      }),
    scale
  );
}

/** How far the overlay boxes are from where the canvas shows the slide and a shape on it. */
async function misalignment(page: Page): Promise<number> {
  return pinnedEvaluate(page, async () => {
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

type ProbeEvent = { type: string; focused?: boolean; viewport?: { x: number; y: number; width: number; height: number } | null };

for (const engine of ['legacy', 'worker'] as const) {
  test.describe(`pptx plugins (${engine})`, () => {
    test.beforeEach(async ({ page }) => {
      await page.addInitScript(() => {
        const probe = { count: 0, capture() { this.count++; } };
        (globalThis as unknown as { __pptxMainEngineProbe: typeof probe }).__pptxMainEngineProbe = probe;
      });
    });
    test.afterEach(async ({ page }) => {
      const count = await page.evaluate(() =>
        (globalThis as unknown as { __pptxMainEngineProbe: { count: number } }).__pptxMainEngineProbe.count);
      if (engine === 'worker') expect(count).toBe(0);
      else expect(count).toBeGreaterThan(0);
    });
test('pptx: a plugin panel gesture admits a granted built-in mutation', async ({ page }) => {
  await open(page, engine);
  expect(await pinnedEvaluate(page, async () => {
    const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
    const read = await editor.readContent();
    const story = read.stories.find((entry) => entry.text.length >= 3)!;
    return editor.selectText({ slide: read.slides.findIndex((slide) => slide.id === story.slideId) + 1,
      shapeId: story.shapeId, storyId: story.storyId, start: 0, end: 3 });
  })).toBe(true);
  const state = () => pinnedEvaluate(page, async () => {
    const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
    return { version: (await editor.readContent()).version, bold: editor.commands.getState('bold').active };
  });
  const before = await state();
  await expect(page.getByTestId('probe-mutation')).toHaveAttribute('data-version', before.version);
  await page.getByTestId('probe-mutation').click();
  await expect.poll(() => page.evaluate(() =>
    (window as unknown as { __probe: Probe }).__probe.mutation)).toMatchObject({ ok: true, status: 'executed' });
  await expect.poll(async () => (await state()).bold).toBe(before.bold !== true);
  expect((await state()).version).not.toBe(before.version);
  const after = await state();
  expect(await pinnedEvaluate(page, () =>
    (window as unknown as { __probe: Probe }).__probe.runMutation())).toMatchObject({
    ok: false, failure: { code: 'unsupported-policy' },
  });
  expect(await state()).toEqual(after);
  await page.getByTestId('probe-timer-mutation').click();
  await expect.poll(() => page.evaluate(() =>
    (window as unknown as { __probe: Probe }).__probe.mutation)).toMatchObject({
    ok: false, failure: { code: 'unsupported-policy' },
  });
  expect(await state()).toEqual(after);
});

test('document focus excludes the plugin panel', async ({ page }) => {
  await open(page, engine);
  const events = () => page.evaluate(() =>
    (window as unknown as { __probe: { events: ProbeEvent[] } }).__probe.events
      .filter((event) => event.type === 'focus-change'));
  await page.getByTestId('probe-input').click();
  await page.evaluate(() => {
    (window as unknown as { __probe: { events: ProbeEvent[] } }).__probe.events = [];
  });
  await page.getByTestId('pptx-slide-canvas').click({ position: { x: 100, y: 100 } });
  await expect.poll(async () => (await events()).at(-1)?.focused).toBe(true);
  await page.getByTestId('probe-input').click();
  await expect.poll(async () => (await events()).at(-1)?.focused).toBe(false);
  expect((await events()).map((event) => event.focused)).toEqual([true, false]);
});

if (engine === 'worker') {
  test('pptx: slide navigation is announced with worker plugins registered', async ({ page }) => {
    await open(page, 'worker');
    const stage = page.getByRole('application');
    await stage.focus();
    await expect(stage).toBeFocused();
    await page.keyboard.press('PageDown');
    await expect(page.getByTestId('pptx-live-region')).toContainText('Slide 2 of 3');
  });
}

test('resizing the page changes the viewport', async ({ page }) => {
  await open(page, engine);
  const viewport = () => page.evaluate(() =>
    (window as unknown as { __probe: { events: ProbeEvent[] } }).__probe.events
      .filter((event) => event.type === 'viewport-change' && event.viewport !== null).at(-1));
  await expect.poll(viewport).toMatchObject({ type: 'viewport-change' });
  const before = (await viewport())!;
  if (before.type !== 'viewport-change') throw new Error('Missing viewport');
  const size = page.viewportSize()!;
  await page.setViewportSize({ width: size.width - 140, height: size.height - 80 });
  await expect.poll(async () => {
    const next = await viewport();
    return next?.type === 'viewport-change' ? next.viewport : null;
  }).not.toEqual(before.viewport);
});

test('pptx: overlays land on the presented slide at every zoom and while scrolled', async ({
  page,
}) => {
  await open(page, engine);
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

test('selection and range anchors follow stage scrolling at every zoom', async ({ page }) => {
  await page.setViewportSize({ width: 1100, height: 760 });
  await open(page, engine);
  const story = await pinnedEvaluate(page, async () => {
    const content = await (window as unknown as { __probe: Probe }).__probe.editor!.readContent();
    return { version: content.version, ...content.stories.find(candidate =>
      candidate.slideId === content.slides[0].id && candidate.text.length > 4)! };
  });
  for (const [anchor, focus] of [[1, 4], [4, 1], [2, 2]]) {
    let baseline: { rects: ProbeRect[]; anchor: ProbeRect } | undefined;
    for (const scale of [0.5, 1, 2] as const) {
      await zoom(page, scale);
      await expect(page.locator(`[data-probe-slide][data-zoom="${scale}"]`)).toBeAttached();
      expect(await pinnedEvaluate(page, async ({ story, anchor, focus }) =>
        await (window as unknown as { __probe: Probe }).__probe.editor!.selectText({
          slide: 1, shapeId: story.shapeId, storyId: story.storyId, start: anchor, end: focus,
        }), { story, anchor, focus }), `selectText anchor=${anchor} focus=${focus} scale=${scale}`).toBe(true);
      await expect(page.locator('[data-probe-anchor="selection"]')).toBeAttached();
      const read = () => pinnedEvaluate(page, async ({ story, anchor, focus }) => {
        const geometry = (window as unknown as { __probe: Probe }).__probe.geometry;
        if (!geometry) return null;
        const targets: ProbeTarget[] = [{ kind: 'selection' }, { kind: 'range', version: story.version,
          slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId, anchor, focus }];
        const answers = targets.map(target => geometry.getAnchorGeometry(target));
        const batch = await geometry.readAnchorGeometries(targets);
        const canvas = document.querySelector('[data-testid="pptx-slide-canvas"]')!.getBoundingClientRect();
        const layer = document.querySelector<HTMLElement>('[data-testid="plugin-overlays"]')!;
        const layerRect = layer.getBoundingClientRect();
        const origin = { x: canvas.left - layerRect.left - layer.clientLeft + layer.scrollLeft,
          y: canvas.top - layerRect.top - layer.clientTop + layer.scrollTop };
        const normalize = (rect: ProbeRect) => ({ x: (rect.x - origin.x) * geometry.layout.width / canvas.width,
          y: (rect.y - origin.y) * geometry.layout.height / canvas.height,
          width: rect.width * geometry.layout.width / canvas.width,
          height: rect.height * geometry.layout.height / canvas.height });
        const markers = ['selection', 'range'].map(name => {
          const marker = document.querySelector(`[data-probe-anchor="${name}"]`)!.getBoundingClientRect();
          return { x: marker.x, y: marker.y, width: marker.width, height: marker.height };
        });
        return { answers, batch, markers, canvas: { x: canvas.x, y: canvas.y },
          normalized: answers[0].ok ? { rects: answers[0].rects.map(normalize), anchor: normalize(answers[0].anchor) } : null };
      }, { story, anchor, focus });
      await expect.poll(async () => (await read())?.answers.every(answer => answer.ok)).toBe(true);
      await page.getByRole('application').evaluate(async stage => {
        stage.style.overflow = 'auto';
        const spacer = document.createElement('div');
        spacer.dataset.anchorScrollSpacer = 'true';
        spacer.style.position = 'absolute';
        spacer.style.left = `${stage.clientWidth + 600}px`;
        spacer.style.top = `${stage.clientHeight + 600}px`;
        spacer.style.width = '1px';
        spacer.style.height = '1px';
        stage.append(spacer);
        stage.scrollTo(0, 0);
        await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
      });
      const before = (await read())!;
      expect(before.batch).toEqual(before.answers);
      expect(before.answers[0]).toEqual(before.answers[1]);
      expect(before.normalized!.rects.length).toBe(anchor === focus ? 0 : 1);
      if (!baseline) baseline = before.normalized!;
      for (const key of ['x', 'y', 'width', 'height'] as const) {
        expect(before.normalized!.anchor[key]).toBeCloseTo(baseline.anchor[key], 3);
      }
      const scrolled = await page.getByRole('application').evaluate(stage => {
        stage.style.overflow = 'auto';
        stage.scrollLeft = 120;
        stage.scrollTop = 80;
        return { x: stage.scrollLeft, y: stage.scrollTop };
      });
      expect(scrolled.x).toBeGreaterThan(0);
      expect(scrolled.y).toBeGreaterThan(0);
      await expect.poll(async () => {
        const after = await read();
        return after ? Math.abs(after.markers[0].x - before.markers[0].x + scrolled.x) +
          Math.abs(after.markers[0].y - before.markers[0].y + scrolled.y) : 999;
      }).toBeLessThan(1.5);
      const after = (await read())!;
      expect(after.batch).toEqual(after.answers);
      expect(after.answers[0]).toEqual(after.answers[1]);
      for (const key of ['x', 'y', 'width', 'height'] as const) {
        expect(after.normalized!.anchor[key]).toBeCloseTo(baseline.anchor[key], 3);
      }
      expect(after.normalized!.rects).toEqual(before.normalized!.rects);
      for (let index = 0; index < 2; index++) {
        expect(after.markers[index].x - before.markers[index].x).toBeCloseTo(after.canvas.x - before.canvas.x, 1);
        expect(after.markers[index].y - before.markers[index].y).toBeCloseTo(after.canvas.y - before.canvas.y, 1);
      }
      await page.getByRole('application').evaluate(stage => {
        stage.scrollTo(0, 0);
        stage.querySelector('[data-anchor-scroll-spacer]')!.remove();
      });
    }
  }
});

test('pptx: overlays leave the pointer to the slide beneath them', async ({ page }) => {
  await open(page, engine);
  const box = await renderedBox(page.locator('[data-probe-shape]'));
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await expect.poll(async () => (await probe(page)).selection?.target.kind).not.toBe('slide');
});

test('pptx: text selection anchor rectangles lie inside the selected shape', async ({ page }) => {
  await open(page, engine);
  const selected = await pinnedEvaluate(page, async () => {
    const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
    const content = await editor.readContent();
    const story = content.stories.find((candidate) => candidate.slideId === content.slides[0].id && candidate.text.length > 3)!;
    const ok = await editor.selectText({ slide: 1, shapeId: story.shapeId, storyId: story.storyId, start: 0, end: 3 });
    return { ok, shapeId: story.shapeId };
  });
  expect(selected.ok).toBe(true);
  await expect.poll(async () => (await probe(page)).selection?.target.kind).toBe('text');
  const readGeometry = () => pinnedEvaluate(page, (shapeId) => {
    const geometry = (window as unknown as { __probe: Probe }).__probe.geometry;
    return geometry ? { answer: geometry.getAnchorGeometry({ kind: 'selection' }), shape: geometry.getShapeRect(shapeId) } : null;
  }, selected.shapeId);
  await expect.poll(async () => {
    const result = await readGeometry();
    if (!result?.answer.ok || !result.shape || !result.answer.rects.length) return false;
    const shape = result.shape;
    return result.answer.rects.every((rect) => rect.x >= shape.x - 1 && rect.y >= shape.y - 1 &&
      rect.x + rect.width <= shape.x + shape.width + 1 && rect.y + rect.height <= shape.y + shape.height + 1);
  }).toBe(true);
  const result = (await readGeometry())!;
  expect(result.shape).not.toBeNull();
  if (!result.answer.ok || !result.shape) throw new Error('Missing selection geometry');
  expect(result.answer.rects.length).toBeGreaterThan(0);
  for (const rect of result.answer.rects) {
    expect(rect.x).toBeGreaterThanOrEqual(result.shape.x - 1);
    expect(rect.y).toBeGreaterThanOrEqual(result.shape.y - 1);
    expect(rect.x + rect.width).toBeLessThanOrEqual(result.shape.x + result.shape.width + 1);
    expect(rect.y + rect.height).toBeLessThanOrEqual(result.shape.y + result.shape.height + 1);
  }
});

test('pptx: navigation shows the slide and keeps focus in the panel', async ({ page }) => {
  await open(page, engine);
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
  await open(page, engine);
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
  await open(page, engine);
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
  await open(page, engine);
  const state = () =>
    pinnedEvaluate(page, async () => {
      const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
      const read = await editor.readContent();
      return { notes: read.slides[0].notes ?? '', text: read.stories[0].text };
    });
  await pinnedEvaluate(page, async () => {
    const editor = (window as unknown as { __probe: Probe }).__probe.editor!;
    const read = await editor.readContent();
    await editor.applyEdits({
      expectVersion: read.version,
      steps: [{ op: 'setSlideNotes', target: { slideId: read.slides[0].id }, text: 'Edited' }],
    });
    const story = read.stories[0];
    await editor.selectText({
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

  });
}
