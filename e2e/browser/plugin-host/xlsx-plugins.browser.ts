import { test, expect, type Page } from 'playwright/test';
import { pinnedEvaluate } from '../pinned-evaluate';

interface ProbeContext {
  snapshot: { version: string };
  read: {
    version(): Promise<{ ok: boolean; version: string }>;
    readCells(request: unknown): Promise<{ ranges?: { cells: { displayText: string }[][] }[] }>;
  };
  edits: { applyEdits(request: unknown): Promise<unknown> } | null;
  navigation: {
    selectCells(target: unknown, options: { expectVersion: string }): Promise<unknown>;
    scrollToCell(target: unknown, options: { expectVersion: string; align: 'start' }): Promise<unknown>;
  };
}

interface Probe {
  editor: { commands: { execute(id: string, args: unknown): Promise<unknown> } } | null;
  act<T>(action: (context: ProbeContext) => Promise<T>): Promise<T>;
  navigation: unknown;
  selection: { sheetId: string; cells: { focus: { row: number; col: number } } | null } | null;
  marks: number;
  mainThreadEngines: number;
}

const probe = (page: Page) =>
  page.evaluate(() => {
    const value = (window as unknown as { __probe: Probe }).__probe;
    return { navigation: value.navigation, selection: value.selection, marks: value.marks };
  });

async function open(page: Page, engine: 'legacy' | 'worker') {
  await page.goto(`/xlsx.html?engine=${engine}`);
  await expect(page.getByTestId('xlsx-scroll').locator('canvas')).toBeVisible({
    timeout: 120_000,
  });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.locator('[data-probe-cell="0:0"]')).toBeAttached({ timeout: 60_000 });
}

async function select(page: Page, row: number, col: number) {
  const results = await pinnedEvaluate(page,
    ([r, c]) =>
      (window as unknown as { __probe: Probe }).__probe.act(async (context) => {
        const expectVersion = context.snapshot.version;
        return [
          await context.navigation.selectCells(
            { sheetId: 'sheet:0', selection: { anchor: { row: r, col: c }, focus: { row: r, col: c } } },
            { expectVersion }
          ),
          await context.navigation.scrollToCell({ sheetId: 'sheet:0', row: r, col: c }, { expectVersion, align: 'start' }),
        ];
      }),
    [row, col] as const
  );
  expect(results).toEqual([{ ok: true }, { ok: true }]);
}

async function zoom(page: Page, scale: number) {
  await pinnedEvaluate(page,
    (value) =>
      (window as unknown as { __probe: Probe }).__probe.editor!.commands.execute('zoom', {
        scale: value,
      }),
    scale
  );
}

/** How far the overlay's box for the selection is from the editor's own selection box. */
async function misalignment(page: Page): Promise<number> {
  return page.evaluate(() => {
    const cell = document.querySelector('[data-probe-cell]')?.getBoundingClientRect();
    const native = document
      .querySelector('[data-testid="xlsx-selection"]')
      ?.getBoundingClientRect();
    if (!cell || !native) return Number.POSITIVE_INFINITY;
    return Math.max(
      Math.abs(cell.left - native.left),
      Math.abs(cell.top - native.top),
      Math.abs(cell.width - native.width),
      Math.abs(cell.height - native.height)
    );
  });
}

for (const engine of ['legacy', 'worker'] as const) {
  test.describe(`xlsx plugins (${engine})`, () => {
    test.afterEach(async ({ page }) => {
      const count = await page.evaluate(
        () => (window as unknown as { __probe: Probe }).__probe.mainThreadEngines
      );
      if (engine === 'worker') expect(count).toBe(0);
      else expect(count).toBeGreaterThan(0);
    });
test('xlsx: overlays land on the painted cells at every zoom, scrolled and in frozen panes', async ({
  page,
}) => {
  await open(page, engine);
  await select(page, 6, 3);
  for (const scale of [0.5, 1, 2] as const) {
    await zoom(page, scale);
    await expect(page.locator(`[data-probe-cell="6:3"][data-zoom="${scale}"]`)).toBeAttached();
    await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
  }
  await zoom(page, 1);
  await select(page, 30, 3);
  await expect(page.locator('[data-probe-cell="30:3"]')).toBeAttached();
  await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.evaluate((element) => element.scrollBy(0, -40));
  await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
  expect(await scroll.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  const corner = (await page.getByTestId('overlay-input').boundingBox())!;
  const canvas = (await scroll.locator('canvas').boundingBox())!;
  expect(Math.abs(corner.x - (canvas.x + 4))).toBeLessThan(1.5);
  expect(Math.abs(corner.y - (canvas.y + 2))).toBeLessThan(1.5);
});

test('xlsx: overlays leave the pointer to the grid beneath them', async ({ page }) => {
  await open(page, engine);
  const box = (await page.locator('[data-probe-target]').boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('C5');
});

test('xlsx: navigation selects on another sheet and keeps focus in the panel', async ({ page }) => {
  await open(page, engine);
  await page.getByTestId('probe-navigate').focus();
  await page.keyboard.press('Enter');
  await expect.poll(async () => (await probe(page)).navigation).toEqual({ ok: true });
  await expect(page.getByTestId('probe-navigate')).toBeFocused();
  await expect(page.getByTestId('xlsx-sheet-tabs').getByRole('tab', { selected: true })).toHaveText(
    'Summary'
  );
  await expect(page.locator('[data-probe-cell="1:1"][data-probe-sheet="sheet:1"]')).toBeAttached();
  await expect.poll(() => misalignment(page), { timeout: 30_000 }).toBeLessThan(1.5);
  await expect
    .poll(async () => (await probe(page)).selection?.cells?.focus)
    .toEqual({ row: 1, col: 1 });
});

test('xlsx: dock tabs work by keyboard, and narrow docks open panels as drawers', async ({
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

test('xlsx: plugin toolbar commands work by pointer, shortcut and the overflow menu', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1920, height: 1000 });
  await open(page, engine);
  const marks = async () => (await probe(page)).marks;
  const rail = page.getByTestId('xlsx-formatting-toolbar');
  await rail.getByRole('button', { name: 'Probe mark' }).click();
  await expect.poll(marks).toBe(1);

  await page.getByTestId('xlsx-scroll').focus();
  await page.keyboard.press('ControlOrMeta+Shift+K');
  await expect.poll(marks).toBe(2);

  await page.setViewportSize({ width: 480, height: 800 });
  const more = rail.getByTestId('xlsx-toolbar-more');
  await expect(more).toBeVisible();
  await more.focus();
  await page.keyboard.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'More toolbar items' });
  await expect(menu).toBeVisible();
  await page.keyboard.press('End');
  await expect(page.locator(':focus')).toHaveAttribute('data-label', 'Probe mark');
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
  await expect.poll(marks).toBe(3);
});

test('xlsx: plugin text fields keep their keys; built-in shortcuts in plugin chrome act on the editor', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1920, height: 1000 });
  await open(page, engine);
  const b3 = () =>
    pinnedEvaluate(page, () =>
      (window as unknown as { __probe: Probe }).__probe.act(async (context) => {
        const read = await context.read.readCells({
          ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }],
        });
        return read.ranges?.[0]?.cells[0]?.[0]?.displayText;
      })
    );
  const edited = await pinnedEvaluate(page, () =>
    (window as unknown as { __probe: Probe }).__probe.act(async (context) =>
      context.edits!.applyEdits({
        expectVersion: (await context.read.version()).version,
        steps: [
          {
            op: 'setCellInputs',
            target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } },
            inputs: [['Edited']],
          },
        ],
      })
    )
  );
  expect(edited).toMatchObject({ ok: true });
  await select(page, 2, 1);
  const version = () =>
    pinnedEvaluate(page, () =>
      (window as unknown as { __probe: Probe }).__probe.act((context) => context.read.version())
    );
  const before = await version();
  expect(before.ok).toBe(true);

  for (const id of ['overlay-input', 'probe-input', 'portal-input']) {
    const field = page.getByTestId(id);
    await field.click();
    await page.keyboard.type('xyz');
    await page.keyboard.press('Backspace');
    await page.keyboard.press('Delete');
    await page.keyboard.press('Enter');
    await expect(field).toHaveValue('xy');
    await page.keyboard.press('ControlOrMeta+Z');
    await page.keyboard.press('ControlOrMeta+V');
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
  }
  expect(await version()).toEqual(before);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B3');

  await page.getByTestId('probe-navigate').focus();
  await page.keyboard.press('ControlOrMeta+Z');
  await expect.poll(b3).not.toBe('Edited');
  await page.getByTestId('plugin-dock-right').getByRole('tab', { name: 'Probe' }).focus();
  await page.keyboard.press('ControlOrMeta+Shift+Z');
  await expect.poll(b3).toBe('Edited');
  await page
    .getByTestId('xlsx-formatting-toolbar')
    .getByRole('button', { name: 'Probe mark' })
    .focus();
  await page.keyboard.press('ControlOrMeta+Z');
  await expect.poll(b3).not.toBe('Edited');

  await page.getByTestId('probe-navigate').focus();
  await page.keyboard.press('ControlOrMeta+Shift+K');
  await expect.poll(async () => (await probe(page)).marks).toBe(1);
});

  });
}
