import { test, expect, type Page } from 'playwright/test';

type Probe = {
  issues: { code: string; id: string }[];
  hostKeys: { key: string; prevented: boolean }[];
  saves: number;
  pluginRuns: number;
  load(shortcuts: unknown): Promise<void>;
  state(id: 'bold' | 'italic'): boolean | 'mixed' | null | undefined;
  texts(): Promise<string[]>;
};

const probe = <T>(page: Page, pick: string) =>
  page.evaluate((body) => new Function('probe', `return ${body}`)(
    (window as unknown as { __shortcuts: Probe }).__shortcuts), pick) as Promise<T>;
const lastKey = (page: Page) => probe<{ key: string; prevented: boolean } | undefined>(page, 'probe.hostKeys.at(-1)');
const canvas = (page: Page) => page.getByTestId('pptx-slide-canvas');
const input = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });

async function open(page: Page, mode: string, shortcuts: unknown) {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto(`/pptx-shortcuts.html?mode=${mode}`);
  await page.evaluate((map) => (window as unknown as { __shortcuts: Probe }).__shortcuts.load(map), shortcuts);
  await expect(canvas(page)).toBeVisible({ timeout: 120_000 });
  await expect.poll(() => probe(page, 'probe.editor !== null'), { timeout: 60_000 }).toBe(true);
}

async function selectTarget(page: Page) {
  await expect(canvas(page)).toHaveAttribute('data-pptx-input-ready', 'true');
  const box = (await canvas(page).boundingBox())!;
  await page.mouse.click(box.x + box.width * 0.15, box.y + box.height * 0.62);
  await expect(input(page)).toBeFocused();
  await input(page).press('ControlOrMeta+Home');
  await input(page).press('ControlOrMeta+Shift+End');
  await expect(page.getByRole('status').filter({ hasText: 'Text selected' })).toBeVisible();
}

test.describe('pptx host shortcuts, worker', () => {
  test('a remapped paste leaves the old chord to the host', async ({ page }) => {
    await open(page, 'worker', { paste: 'Mod+Alt+V' });
    await selectTarget(page);
    await page.evaluate(() => navigator.clipboard.writeText('Outside'));
    await input(page).press('Control+V');
    await page.waitForTimeout(400);
    expect(await probe<string[]>(page, 'probe.texts()')).toContain('Target');
    expect(await lastKey(page)).toEqual({ key: 'v', prevented: false });
    await input(page).press('Control+Alt+V');
    await expect.poll(() => probe<string[]>(page, 'probe.texts()')).toContain('Outside');
  });

  test('a released plain paste chord pastes nothing; the default one pastes', async ({ page }) => {
    await open(page, 'worker', { pasteAsPlainText: null });
    await selectTarget(page);
    await page.evaluate(() => navigator.clipboard.writeText('Plain'));
    await input(page).press('Control+Shift+V');
    await page.waitForTimeout(400);
    expect(await probe<string[]>(page, 'probe.texts()')).toContain('Target');
    expect(await lastKey(page)).toEqual({ key: 'v', prevented: false });
    await input(page).press('Control+V');
    await expect.poll(() => probe<string[]>(page, 'probe.texts()')).toContain('Plain');
  });

  test('moving bold frees its chord for another command', async ({ page }) => {
    await open(page, 'worker', { bold: 'Mod+Alt+B', italic: 'Mod+B' });
    expect(await probe(page, 'probe.issues')).toEqual([]);
    await selectTarget(page);
    await input(page).press('Control+B');
    await expect.poll(() => probe(page, 'probe.state("italic")')).toBe(true);
    expect(await probe(page, 'probe.state("bold")')).toBe(false);
    await input(page).press('Control+Alt+B');
    await expect.poll(() => probe(page, 'probe.state("bold")')).toBe(true);
  });

  test('a released copy chord is left to the host', async ({ page }) => {
    await open(page, 'worker', { copy: null, cut: 'Mod+Alt+X' });
    await selectTarget(page);
    await page.evaluate(() => navigator.clipboard.writeText('sentinel'));
    await input(page).press('Control+C');
    await page.waitForTimeout(400);
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('sentinel');
    await input(page).press('Control+Shift+X');
    await page.waitForTimeout(400);
    expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('sentinel');
    expect(await probe<string[]>(page, 'probe.texts()')).toContain('Target');
    await input(page).press('Control+Alt+X');
    await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Target');
  });

  test('the context menu shows the remapped chord', async ({ page }) => {
    await open(page, 'worker', { paste: 'Mod+Alt+V', copy: null });
    await selectTarget(page);
    await input(page).press('Shift+F10');
    const paste = page.getByRole('menuitem', { name: 'Paste', exact: true });
    await expect(paste).toContainText('Ctrl+Alt+V');
    await expect(paste).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+V');
    await expect(page.getByRole('menuitem', { name: 'Copy', exact: true })).not.toContainText('Ctrl');
  });
});

for (const mode of ['worker', 'legacy']) {
  test.describe(`pptx host shortcuts, ${mode}`, () => {
    test('a moved save chord runs; the old one reaches the host', async ({ page }) => {
      await open(page, mode, { save: 'Mod+Alt+S', bold: 'Mod+Alt+B' });
      await canvas(page).click();
      await page.keyboard.press('Control+S');
      await page.waitForTimeout(300);
      expect(await probe(page, 'probe.saves')).toBe(0);
      expect(await lastKey(page)).toEqual({ key: 's', prevented: false });
      await page.keyboard.press('Control+Alt+S');
      await expect.poll(() => probe(page, 'probe.saves')).toBe(1);
      const bold = page.getByTestId('pptx-bold');
      await expect(bold).toHaveAttribute('title', /Ctrl\+Alt\+B/);
      await expect(bold).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+B');
    });

    test('clipboard ids are refused only where there is no clipboard', async ({ page }) => {
      await open(page, mode, { paste: 'Mod+Alt+V' });
      const issues = await probe<{ code: string; id: string }[]>(page, 'probe.issues');
      expect(issues.map(({ code, id }) => [code, id])).toEqual(mode === 'legacy' ? [['unknown-id', 'paste']] : []);
    });

    test('refused entries are reported and plugin chords follow the map', async ({ page }) => {
      await open(page, mode, { bold: 'Mod+ArrowLeft', nope: 'Mod+J', 'plugin:acme.keys/run': 'Mod+Alt+K' });
      expect((await probe<{ code: string; id: string }[]>(page, 'probe.issues')).map(({ code, id }) => [code, id])).toEqual([
        ['reserved-chord', 'bold'],
        ['unknown-id', 'nope'],
      ]);
      await canvas(page).click();
      await page.keyboard.press('Control+Shift+K');
      await page.waitForTimeout(300);
      expect(await probe(page, 'probe.pluginRuns')).toBe(0);
      await page.keyboard.press('Control+Alt+K');
      await expect.poll(() => probe(page, 'probe.pluginRuns')).toBe(1);
    });
  });
}
