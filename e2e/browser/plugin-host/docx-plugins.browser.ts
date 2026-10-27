import { test, expect, type Page } from 'playwright/test';

let peerless = false;

async function open(page: Page) {
  await page.goto(peerless ? '/?peerless=1' : '/');
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('[data-probe-page="0"]')).toBeAttached({ timeout: 60_000 });
}

// Chromium's inspector holds an awaited promise weakly: a GC after it settles but before the reply
// fails the call ("Promise was collected", reported as a navigation). The holder keeps it reachable.
async function operate<R, A = undefined>(page: Page, operation: (arg: A) => Promise<R>, arg?: A): Promise<R> {
  const holder = await page.evaluateHandle<{ run(arg: unknown): Promise<R>; promise?: Promise<R> }>(
    `({ run: ${operation} })`
  );
  try {
    const result = await holder.evaluateHandle((holder, arg) => (holder.promise = holder.run(arg)), arg);
    try {
      return (await result.jsonValue()) as R;
    } finally {
      await result.dispose();
    }
  } finally {
    await holder.dispose();
  }
}

async function offsets(page: Page) {
  return page.evaluate(() => {
    const canvas = document.querySelector('canvas[data-page-index="0"]')?.getBoundingClientRect();
    const box = document.querySelector('[data-probe-page="0"]')?.getBoundingClientRect();
    if (!canvas || !box) return [Infinity];
    return [
      box.left - canvas.left,
      box.top - canvas.top,
      box.width - canvas.width,
      box.height - canvas.height,
    ].map((value) => Math.abs(value));
  });
}

for (const mode of ['hybrid', 'peerless'] as const) {
  test.describe(mode, () => {
    test.beforeEach(() => { peerless = mode === 'peerless'; });
test('managed overlays land on the rendered pages at every zoom', async ({ page }) => {
  await open(page);
  await expect(page.getByTestId('probe-card')).toBeVisible();
  for (const zoom of [0.5, 1, 2]) {
    await page.evaluate(
      (value) =>
        (
          window as unknown as { __probe: { editor: { setZoom(zoom: number): void } } }
        ).__probe.editor.setZoom(value),
      zoom
    );
    await expect
      .poll(async () => Math.max(...(await offsets(page))), { timeout: 30_000 })
      .toBeLessThan(1.5);
    const width = await page.evaluate(
      () => document.querySelector('[data-probe-page="0"]')!.getBoundingClientRect().width
    );
    expect(width).toBeGreaterThan(100 * zoom);
  }
});

test('after an edit, overlays and sidebar cards follow the new version; cards keep their state', async ({
  page,
}) => {
  await open(page);
  const card = page.getByTestId('probe-card');
  await expect(card).toBeVisible();
  await card.click();
  await expect(card).toHaveAttribute('data-clicks', '1');
  const before = await card.getAttribute('data-version');
  await page.getByTestId('probe-append').click();
  await expect(card).not.toHaveAttribute('data-version', before!);
  const version = await operate(page, async () => {
    const editor = (
      window as unknown as {
        __probe: {
          editor: { readParagraphs(request: { view: 'accepted' }): Promise<{ version: string }> };
        };
      }
    ).__probe.editor;
    return (await editor.readParagraphs({ view: 'accepted' })).version;
  });
  await expect(card).toHaveAttribute('data-version', version);
  await expect(card).toHaveAttribute('data-clicks', '1');
  await expect
    .poll(async () => Math.max(...(await offsets(page))), { timeout: 30_000 })
    .toBeLessThan(1.5);
});

test('navigation scrolls the paragraph into view and reports success', async ({ page }) => {
  await open(page);
  const scroller = page.locator('.docx-editor__scroll-container');
  const before = await scroller.evaluate((element) => element.scrollTop);
  await page.getByTestId('probe-navigate').click();
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as unknown as { __probe: { navigation: unknown } }).__probe.navigation
      )
    )
    .toEqual({ ok: true });
  await expect
    .poll(() => scroller.evaluate((element) => element.scrollTop))
    .toBeGreaterThan(before);
});

test('navigation keeps the selection unless focus is requested', async ({ page }) => {
  await open(page);
  const paragraphs = () => operate(page, async () =>
    (await (window as unknown as LooseProbe).__probe.editor.readParagraphs({ view: 'accepted' })).paragraphs);
  const text = async (paraId: string) => (await paragraphs()).find(paragraph => paragraph.paraId === paraId)?.text;
  const navigate = async (id: string) => {
    await page.evaluate(() => { (window as unknown as { __probe: { navigation: unknown } }).__probe.navigation = null; });
    await page.getByTestId(id).click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { __probe: { navigation: unknown } }).__probe.navigation)).toEqual({ ok: true });
  };
  const before = await paragraphs();
  const origin = before[0]!;
  const destination = before.filter(paragraph => paragraph.text.trim()).at(-1)!;
  expect(destination.paraId).not.toBe(origin.paraId);
  await page.evaluate(() => (window as unknown as LooseProbe).__probe.editor.focus());
  await page.keyboard.press('ControlOrMeta+Home');
  await navigate('probe-navigate');
  await page.evaluate(() => (window as unknown as LooseProbe).__probe.editor.focus());
  await page.keyboard.type('Q7');
  await expect.poll(() => text(origin.paraId)).toBe(`Q7${origin.text}`);
  expect(await text(destination.paraId)).toBe(destination.text);
  await navigate('probe-navigate-focus');
  await page.keyboard.type('R8');
  await expect.poll(() => text(destination.paraId)).toBe(`R8${destination.text}`);
  expect(await text(origin.paraId)).toBe(`Q7${origin.text}`);
});

test('navigation builds and reaches a paragraph on an unresident page', async ({ page }) => {
  await open(page);
  const filled = await operate(page, async () => {
    const { editor } = (window as unknown as LooseProbe).__probe;
    const read = await editor.readParagraphs({ view: 'accepted' });
    return editor.applyEdits({ expectVersion: read.version, steps: [{ op: 'insertParagraphs',
      target: { story: 'body', paraId: read.paragraphs.at(-1)!.paraId }, at: 'end',
      paragraphs: Array.from({ length: 100 }, (_, index) => ({ text: `Paragraph ${index + 1} needs a later resident page.` })) }] });
  });
  expect(filled).toMatchObject({ ok: true, applied: true });
  // Wait for the probe plugin's async document-change read to update its panel state.
  await expect(page.getByTestId('probe-card')).toHaveAttribute('data-version', (filled as { version: string }).version);
  await page.getByTestId('probe-navigate').click();
  await expect.poll(() => page.evaluate(() => (window as unknown as { __probe: { navigation: unknown } }).__probe.navigation)).toEqual({ ok: true });
  await expect.poll(() => page.locator('.docx-editor__scroll-container').evaluate(element => element.scrollTop)).toBeGreaterThan(1000);
});

test('dock tabs work by keyboard, and narrow docks open panels as drawers', async ({ page }) => {
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

test('plugin toolbar commands work by pointer, shortcut and the keyboard overflow menu', async ({
  page,
}) => {
  await open(page);
  const marks = () =>
    page.evaluate(() => (window as unknown as { __probe: { marks: number } }).__probe.marks);
  const bar = page.getByTestId('formatting-bar');
  await bar.getByRole('button', { name: 'Probe mark' }).click();
  await expect.poll(marks).toBe(1);

  await page
    .locator('.canvas-page canvas')
    .first()
    .click({ position: { x: 150, y: 150 } });
  await expect(page.locator(peerless ? '.paged-editor__input' : '[data-testid=yrs-input]')).toBeFocused();
  await page.keyboard.press('ControlOrMeta+Alt+Shift+P');
  await expect.poll(marks).toBe(2);

  await page.setViewportSize({ width: 360, height: 800 });
  const more = bar.getByTestId('toolbar-more');
  await expect(more).toBeVisible();
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'More actions' });
  await expect(menu).toBeVisible();
  await page.keyboard.press('End');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Probe mark');
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
  await expect(more).toBeFocused();
  await expect.poll(marks).toBe(3);
});

test('sidebar cards make room for comments and start afresh after replacement', async ({
  page,
}) => {
  await open(page);
  const card = page.getByTestId('probe-card');
  await expect(card).toBeVisible();
  await operate(page, async (worker) => {
    const editor = (
      window as unknown as {
        __probe: {
          editor: {
            readParagraphs(request: {
              view: 'accepted';
            }): Promise<{ paragraphs: { paraId: string; text: string }[] }>;
            addComment(options: { paraId: string; text: string; author: string }): number | null;
            insertComment(options: { paraId: string; text: string; author: string }): Promise<number | null>;
          };
        };
      }
    ).__probe.editor;
    const read = await editor.readParagraphs({ view: 'accepted' });
    const first = read.paragraphs.find((paragraph) => paragraph.text.trim().length > 0)!;
    const options = { paraId: first.paraId, text: 'Collision note', author: 'Probe' };
    if (worker) await editor.insertComment(options);
    else editor.addComment(options);
  }, peerless);
  const comment = page.locator('.docx-unified-sidebar').getByText('Collision note');
  await expect(comment).toBeVisible();
  await expect
    .poll(async () => {
      const a = await card.boundingBox();
      const b = await comment.boundingBox();
      if (!a || !b) return false;
      return a.y + a.height <= b.y || b.y + b.height <= a.y;
    })
    .toBe(true);

  await card.click();
  await expect(card).toHaveAttribute('data-expanded', 'true');
  await expect(card).toHaveAttribute('data-clicks', '1');
  const before = await card.getAttribute('data-version');
  await operate(page, () =>
    (window as unknown as { __probe: { reload(): Promise<void> } }).__probe.reload()
  );
  await expect(card).toBeVisible({ timeout: 60_000 });
  await expect(card).not.toHaveAttribute('data-version', before!);
  await expect(card).toHaveAttribute('data-expanded', 'false');
  await expect(card).toHaveAttribute('data-clicks', '0');
});

interface PointHit {
  position: number;
  pageIndex: number;
  region: string;
  version: string;
  layoutId?: string;
  target: {
    kind: 'range';
    story: string;
    start: { paraId: string; offset: number };
    end: { paraId: string; offset: number };
    view: 'accepted';
  };
}

interface LooseProbe {
  __probe: {
    editor: {
      setZoom(zoom: number): void;
      focus(): void;
      readParagraphs(request: {
        story?: string;
        view: 'accepted';
      }): Promise<{ version: string; paragraphs: { paraId: string; text: string }[] }>;
      applyEdits(request: { expectVersion: string; steps: unknown[] }): Promise<unknown>;
      whenLayoutComplete(options?: { timeoutMs?: number }): Promise<number>;
      getPositionAtPoint(clientX: number, clientY: number): PointHit | null;
      readPositionAtPoint(clientX: number, clientY: number): Promise<PointHit | null>;
      readSelectionInfo(): Promise<{ paraId: string; before: string } | null>;
      getEditorRef(): {
        getYrsSession(): { selection(): { head: unknown } | null };
        displayPositionToYrsLoc(hit: PointHit): unknown;
      };
    };
    geometry: {
      layout: { version: string; pageCount: number };
      getPositionAtPoint(clientX: number, clientY: number): PointHit | null;
      readPositionAtPoint(clientX: number, clientY: number): Promise<PointHit | null>;
    };
  };
}

test('page overlays stay on the previous layout while an edit lays out, then follow the new one', async ({
  page,
}) => {
  await open(page);
  const added = page.locator(`[data-probe-page="${await page.evaluate(() => (window as unknown as LooseProbe).__probe.geometry.layout.pageCount)}"]`);
  await expect(added).not.toBeAttached();
  const version = await operate(page, async () => {
    const { editor } = (window as unknown as LooseProbe).__probe;
    const read = await editor.readParagraphs({ view: 'accepted' });
    const gaps = { count: 0 };
    const observer = new MutationObserver(() => {
      if (!document.querySelector('[data-probe-page="0"]')) gaps.count += 1;
    });
    observer.observe(document.body, { childList: true, subtree: true });
    Object.assign(window, {
      __gaps: gaps,
      __gapObserver: observer,
      __probePage: document.querySelector('[data-probe-page="0"]'),
    });
    const paragraphs = Array.from({ length: 80 }, (_, index) => ({
      text: `Filler paragraph ${index + 1} moves the document onto a second page.`,
    }));
    const applied = (await editor.applyEdits({
      expectVersion: read.version,
      steps: [
        {
          op: 'insertParagraphs',
          target: { story: 'body', paraId: read.paragraphs.at(-1)!.paraId },
          at: 'end',
          paragraphs,
        },
      ],
    })) as { version: string };
    await editor.whenLayoutComplete({ timeoutMs: 60_000 });
    return applied.version;
  });
  await expect(added).toBeAttached({ timeout: 60_000 });
  await expect.poll(() => page.evaluate(() => (window as unknown as LooseProbe).__probe.geometry.layout.version), { timeout: 30_000 }).toBe(version);
  const after = await page.evaluate(() => {
    const { __gaps, __gapObserver, __probe, __probePage } = window as unknown as LooseProbe & {
      __gaps: { count: number };
      __gapObserver: MutationObserver;
      __probePage: Element | null;
    };
    __gapObserver.disconnect();
    return {
      gaps: __gaps.count,
      same: document.querySelector('[data-probe-page="0"]') === __probePage,
      version: __probe.geometry.layout.version,
    };
  });
  expect(after).toEqual({ gaps: 0, same: true, version });
  await expect
    .poll(async () => Math.max(...(await offsets(page))), { timeout: 30_000 })
    .toBeLessThan(1.5);
});

test('a drop point on page 2 resolves like a caret click and inserts there at every zoom', async ({
  page,
}) => {
  await open(page);
  const filled = await operate(page, async () => {
    const { editor } = (window as unknown as LooseProbe).__probe;
    const read = await editor.readParagraphs({ view: 'accepted' });
    const paragraphs = Array.from({ length: 80 }, (_, index) => ({
      text: `Filler paragraph ${index + 1} moves the drop target onto the second page.`,
    }));
    return editor.applyEdits({
      expectVersion: read.version,
      steps: [
        {
          op: 'insertParagraphs',
          target: { story: 'body', paraId: read.paragraphs.at(-1)!.paraId },
          at: 'end',
          paragraphs,
        },
      ],
    });
  });
  expect(filled).toMatchObject({ ok: true, applied: true });
  const second = page.locator('canvas[data-page-index="1"]');
  await expect(second).toBeAttached({ timeout: 60_000 });
  for (const zoom of [0.75, 1.5]) {
    await page.evaluate(
      (value) => (window as unknown as LooseProbe).__probe.editor.setZoom(value),
      zoom
    );
    await expect
      .poll(async () => Math.max(...(await offsets(page))), { timeout: 30_000 })
      .toBeLessThan(1.5);
    await second.scrollIntoViewIfNeeded();
    await expect(second).toBeInViewport();
    const query = await operate(page, async (worker) => {
      const { editor, geometry } = (window as unknown as LooseProbe).__probe;
      const surface = document.querySelector('canvas[data-page-index="1"]')!;
      const canvas = surface.getBoundingClientRect();
      const session = worker ? null : editor.getEditorRef().getYrsSession();
      const selection = JSON.stringify(worker ? await editor.readSelectionInfo() : session!.selection());
      const active = document.activeElement;
      let found: { x: number; y: number; hit: PointHit } | null = null;
      for (let fy = 0.1; fy < 0.9 && !found; fy += 0.01) {
        for (let fx = 0.15; fx < 0.8 && !found; fx += 0.02) {
          const x = canvas.left + canvas.width * fx;
          const y = canvas.top + canvas.height * fy;
          if (y < 0 || y > window.innerHeight) continue;
          if (document.elementFromPoint(x, y) !== surface) continue;
          const hit = worker ? await editor.readPositionAtPoint(x, y) : editor.getPositionAtPoint(x, y);
          if (hit?.region === 'body' && hit.pageIndex === 1 && hit.target.start.offset > 4) {
            found = { x, y, hit };
          }
        }
      }
      return {
        found,
        loc: found && (worker ? { paraId: found.hit.target.start.paraId, offset: found.hit.target.start.offset } : editor.getEditorRef().displayPositionToYrsLoc(found.hit)),
        plugin: found && (worker ? await geometry.readPositionAtPoint(found.x, found.y) : geometry.getPositionAtPoint(found.x, found.y)),
        outside: worker ? await editor.readPositionAtPoint(canvas.left - 12, canvas.top + canvas.height / 2) : editor.getPositionAtPoint(canvas.left - 12, canvas.top + canvas.height / 2),
        unchanged:
          JSON.stringify(worker ? await editor.readSelectionInfo() : session!.selection()) === selection && document.activeElement === active,
      };
    }, peerless);
    expect(query.found).not.toBeNull();
    const { x, y, hit } = query.found!;
    expect(query.outside).toBeNull();
    expect(query.unchanged).toBe(true);
    expect(query.plugin).toEqual({ ...hit, layoutId: expect.any(String) });

    await page.mouse.click(x, y);
    await expect
      .poll(() =>
        operate(page, async (worker) => {
          const { editor } = (window as unknown as LooseProbe).__probe;
          if (!worker) return editor.getEditorRef().getYrsSession().selection()?.head ?? null;
          const info = await editor.readSelectionInfo();
          return info ? { paraId: info.paraId, offset: info.before.length } : null;
        }, peerless)
      )
      .toEqual(query.loc);

    const inserted = await operate(page, async ({ target, version }: PointHit) => {
      const { editor } = (window as unknown as LooseProbe).__probe;
      const text = async (): Promise<string> => {
        const read = await editor.readParagraphs({ story: target.story, view: 'accepted' });
        return read.paragraphs.find((paragraph) => paragraph.paraId === target.start.paraId)!
          .text;
      };
      const before = await text();
      const steps = [{ op: 'insertText', target, at: 'start', text: '†' }];
      const applied = await editor.applyEdits({ expectVersion: version, steps });
      const stale = await editor.applyEdits({ expectVersion: version, steps });
      return { before, after: await text(), applied, stale };
    }, hit);
    const offset = hit.target.start.offset;
    expect(inserted.applied).toMatchObject({ ok: true, applied: true });
    expect(inserted.stale).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(inserted.after).toBe(
      `${inserted.before.slice(0, offset)}†${inserted.before.slice(offset)}`
    );
  }
});

test('a header anchor has a rect in the header band of every page at 100% and 150% zoom', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 1600 });
  await page.goto(`/?doc=repeated-header${peerless ? '&peerless=1' : ''}`);
  await expect(page.locator('canvas[data-page-index="2"]')).toBeVisible({ timeout: 120_000 });
  for (const zoom of [1, 1.5]) {
    await page.evaluate(
      (value) => (window as unknown as LooseProbe).__probe.editor.setZoom(value),
      zoom
    );
    await expect
      .poll(
        () =>
          page.evaluate((scale) =>
            Array.from(document.querySelectorAll<HTMLElement>('[data-probe-header-rect]'), (element) => {
              const pageIndex = Number(element.dataset.probeHeaderRect);
              const rect = element.getBoundingClientRect();
              const canvas = document
                .querySelector(`canvas[data-page-index="${pageIndex}"]`)
                ?.getBoundingClientRect();
              const top = canvas ? (rect.top - canvas.top) / scale : -1;
              const left = canvas ? (rect.left - canvas.left) / scale : -1;
              return {
                pageIndex,
                inHeaderBand: rect.width > 0 && top >= 0 && top + rect.height / scale <= 72 && left > 96,
              };
            }), zoom),
        { timeout: 30_000 }
      )
      .toEqual([0, 1, 2].map((pageIndex) => ({ pageIndex, inHeaderBand: true })));
  }
});

  });
}
