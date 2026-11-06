import { test, expect, type Page } from 'playwright/test';

type Probe = {
  issues: { code: string; id: string; chord?: string; other?: string }[];
  hostKeys: { key: string; prevented: boolean }[];
  pluginRuns: number;
  load(shortcuts: unknown): Promise<void>;
  bold(): boolean | 'mixed' | null;
  italic(): boolean | 'mixed' | null;
};

const read = <T>(page: Page, pick: (probe: Probe) => T) =>
  page.evaluate((source) => new Function('probe', `return (${source})(probe)`)(
    (window as unknown as { __shortcuts: Probe }).__shortcuts), pick.toString()) as Promise<T>;

async function open(page: Page, mode: string, shortcuts: unknown) {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto(`/xlsx-shortcuts.html?mode=${mode}`);
  await page.evaluate((map) => (window as unknown as { __shortcuts: Probe }).__shortcuts.load(map), shortcuts);
  const grid = page.getByTestId('xlsx-scroll');
  await expect(grid.locator('canvas').first()).toBeVisible({ timeout: 120_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1', { timeout: 60_000 });
  await expect.poll(() => read(page, (probe) => probe.bold()), { timeout: 60_000 }).not.toBeNull();
  await grid.focus();
  return grid;
}

const cellMenu = (page: Page) => page.getByRole('menu', { name: 'Cell menu', exact: true });

for (const mode of ['async', 'legacy', 'legacy-worker']) {
  test.describe(`xlsx host shortcuts, ${mode}`, () => {
    test('moving bold frees its chord for another command', async ({ page }) => {
      const grid = await open(page, mode, { bold: 'Mod+Alt+B', italic: 'Mod+B' });
      expect(await read(page, (probe) => probe.issues)).toEqual([]);
      const italic = await read(page, (probe) => probe.italic());
      const bold = await read(page, (probe) => probe.bold());
      await grid.press('Control+B');
      await expect.poll(() => read(page, (probe) => probe.italic())).toBe(!italic);
      expect(await read(page, (probe) => probe.bold())).toBe(bold);
      await grid.press('Control+Alt+B');
      await expect.poll(() => read(page, (probe) => probe.bold())).toBe(!bold);
      const button = page.getByRole('button', { name: 'Bold', exact: true }).first();
      await expect(button).toHaveAttribute('title', /Ctrl\+Alt\+B/);
      await expect(button).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+B');
    });

    test('a disabled chord passes through to the host', async ({ page }) => {
      const grid = await open(page, mode, { bold: null });
      const bold = await read(page, (probe) => probe.bold());
      await grid.press('Control+B');
      await page.waitForTimeout(300);
      expect(await read(page, (probe) => probe.bold())).toBe(bold);
      expect((await read(page, (probe) => probe.hostKeys)).at(-1)).toEqual({ key: 'b', prevented: false });
      await expect(page.getByRole('button', { name: 'Bold', exact: true }).first()).not.toHaveAttribute('aria-keyshortcuts');
    });

    test('a remapped paste leaves the old chord to the host and shows in the menu', async ({ page }) => {
      const grid = await open(page, mode, { paste: 'Mod+Alt+V', copy: null });
      await page.evaluate(() => navigator.clipboard.writeText('RICH'));
      const before = await page.getByTestId('xlsx-formula-input').inputValue();
      await grid.press('Control+V');
      await page.waitForTimeout(400);
      await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(before);
      expect((await read(page, (probe) => probe.hostKeys)).at(-1)).toEqual({ key: 'v', prevented: false });
      await grid.press('Control+Alt+V');
      await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('RICH');
      await grid.press('Shift+F10');
      const menu = cellMenu(page);
      await expect(menu).toBeVisible();
      const paste = menu.getByRole('menuitem', { name: 'Paste', exact: true });
      await expect(paste).toContainText('Ctrl+Alt+V');
      await expect(paste).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+V');
      await expect(menu.getByRole('menuitem', { name: 'Copy', exact: true })).not.toContainText('Ctrl');
    });

    test('refused entries are reported and plugin chords follow the map', async ({ page }) => {
      const grid = await open(page, mode, {
        bold: 'Mod+ArrowLeft',
        nope: 'Mod+J',
        'plugin:acme.keys/run': 'Mod+Alt+K',
      });
      expect((await read(page, (probe) => probe.issues)).map(({ code, id }) => [code, id])).toEqual([
        ['reserved-chord', 'bold'],
        ['unknown-id', 'nope'],
      ]);
      await grid.press('Control+Shift+K');
      await page.waitForTimeout(300);
      expect(await read(page, (probe) => probe.pluginRuns)).toBe(0);
      await grid.press('Control+Alt+K');
      await expect.poll(() => read(page, (probe) => probe.pluginRuns)).toBe(1);
    });
  });
}
