import { test, expect, type Page } from 'playwright/test';

interface RowProbe {
  editor: { commands: { execute(id: string, args: unknown): Promise<unknown> } } | null;
  activations: string[];
}

interface GutterReport {
  outside: string[];
  buttons: { id: string; width: number; height: number; reachable: boolean }[];
}

const rowProbe = (page: Page) =>
  page.evaluate(() => (window as unknown as { __rowProbe: RowProbe }).__rowProbe.activations);

function measure(gutter: Element): GutterReport {
  const bounds = gutter.getBoundingClientRect();
  const outside = [...gutter.querySelectorAll('*')].filter((child) => {
    const rect = child.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return false;
    return rect.left < bounds.left - 0.5 || rect.right > bounds.right + 0.5 ||
      rect.top < bounds.top - 0.5 || rect.bottom > bounds.bottom + 0.5;
  }).map((child) => child.outerHTML.slice(0, 160));
  const buttons = [...gutter.querySelectorAll<HTMLButtonElement>('button')].map((button) => {
    const rect = button.getBoundingClientRect();
    const hit = document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2);
    return { id: button.dataset.decorationId ?? '', width: rect.width, height: rect.height,
      reachable: hit !== null && button.contains(hit) };
  });
  return { outside, buttons };
}

async function painted(page: Page, engine: 'legacy' | 'worker', scale: number, base: number) {
  if (engine === 'worker') {
    await expect(page.getByTestId('xlsx-scroll').locator('canvas')).toHaveAttribute('data-worker-zoom', String(scale));
  }
  const row = page.getByTestId('xlsx-row-gutter').locator('[data-row="6"]');
  await expect.poll(async () => Number.parseFloat(await row.evaluate((element) => (element as HTMLElement).style.height)))
    .toBeCloseTo(base * scale, 3);
}

for (const engine of ['legacy', 'worker'] as const) {
  test(`crowded row decorations stay inside the gutter at zoom 1 and 2 (${engine})`, async ({ page }) => {
    await page.goto(`/xlsx-row-decorations.html?engine=${engine}`);
    const gutter = page.getByTestId('xlsx-row-gutter');
    await expect(gutter.locator('[data-row="2"] button')).toHaveCount(2, { timeout: 120_000 });
    await expect(gutter.locator('[data-row="8"] button')).toHaveCount(2);
    const base = Number.parseFloat(await gutter.locator('[data-row="6"]').evaluate((element) => (element as HTMLElement).style.height));
    expect(base).toBeGreaterThan(0);
    const short = await gutter.locator('[data-row="8"]').boundingBox();
    expect(short!.height).toBeCloseTo(10, 0);
    await expect(gutter.locator('[data-row="8"] [data-testid="xlsx-row-decoration-more"]')).toHaveText('+8998');
    const second = await gutter.locator('[data-decoration-id="short1"]').boundingBox();
    await page.mouse.click(second!.x + second!.width / 2, second!.y + second!.height / 2);
    await expect.poll(() => rowProbe(page)).toEqual(['short1']);
    for (const scale of [1, 2]) {
      if (scale !== 1) {
        await page.evaluate((value) =>
          (window as unknown as { __rowProbe: RowProbe }).__rowProbe.editor!.commands.execute('zoom', { scale: value }), scale);
      }
      await painted(page, engine, scale, base);
      await expect(gutter).toHaveCSS('width', `${24 * scale}px`);
      await expect(gutter.locator('[data-row="6"] button')).toHaveCount(1);
      const more = gutter.locator('[data-row="2"] [data-testid="xlsx-row-decoration-more"]');
      await expect(more).toHaveText('+1');
      await expect(more).toHaveAttribute('title', 'Third');
      const report = await gutter.evaluate(measure);
      expect(report.outside).toEqual([]);
      expect(report.buttons.map((button) => button.id)).toEqual(['a', 'b', 'd', 'e', 'f', 'short0', 'short1']);
      for (const button of report.buttons) {
        expect(button.width, button.id).toBeGreaterThanOrEqual(10);
        expect(button.height, button.id).toBeGreaterThanOrEqual(10);
        expect(button.reachable, button.id).toBe(true);
      }
    }
    await gutter.locator('[data-decoration-id="b"]').click();
    await gutter.locator('[data-decoration-id="d"]').click();
    await expect.poll(() => rowProbe(page)).toEqual(['short1', 'b', 'd']);
  });
}
