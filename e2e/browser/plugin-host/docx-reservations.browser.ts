import { test, expect, type Page } from 'playwright/test';
import { renderedBox } from '../rendered-box';

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface ReservationProbe {
  editor: { save(): Promise<ArrayBuffer | null> } | null;
  geometry: {
    readAnchorGeometries(targets: unknown[]): Promise<({ ok: true; anchor: Rect; rects: Rect[] } | { ok: false; failure: { code: string } })[]>;
  } | null;
  target: { version: string; paraId: string; text: string } | null;
  setPlugins(enabled: boolean): void;
  setReadOnly(readOnly: boolean): void;
}

type ProbeWindow = { __reservations: ReservationProbe; __saved?: Uint8Array };

async function box(page: Page, id: string): Promise<Rect> {
  return renderedBox(page.locator(`[data-reserved="${id}"]`), 120_000);
}

async function open(page: Page, readOnly = false) {
  await page.goto(`/docx-reservations.html${readOnly ? '?readOnly' : ''}`);
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  return { chip: await box(page, 'chip'), gap: await box(page, 'gap') };
}

test('a viewer opened read-only lays out a plugin overlay in its reservations', async ({ page }) => {
  const { chip, gap } = await open(page, true);
  expect(chip.height / chip.width).toBeCloseTo(0.5, 1);
  expect(gap.height / chip.height).toBeCloseTo(80 / 24, 1);
  expect(gap.y).toBeGreaterThan(chip.y);
  await page.evaluate(() => (window as unknown as ProbeWindow).__reservations.setPlugins(false));
  await expect(page.locator('[data-reserved]')).toHaveCount(0);
});

test('a plugin draws in the space it reserves, read-only too, and saves leave it out', async ({ page }) => {
  const { chip, gap } = await open(page);
  expect(chip.height / chip.width).toBeCloseTo(0.5, 1);
  expect(gap.height / chip.height).toBeCloseTo(80 / 24, 1);
  expect(gap.y).toBeGreaterThan(chip.y);

  const [reserved, word] = await page.evaluate(async () => {
    const probe = (window as unknown as ProbeWindow).__reservations;
    const target = probe.target!;
    const start = target.text.indexOf(' ') + 1;
    const end = target.text.indexOf(' ', start) < 0 ? target.text.length : target.text.indexOf(' ', start);
    const at = (offset: number) => ({ paraId: target.paraId, offset });
    return probe.geometry!.readAnchorGeometries([
      { kind: 'reservation', id: 'chip' },
      { kind: 'range', version: target.version, range: { story: 'body', start: at(start), end: at(end), view: 'accepted' } },
    ]);
  });
  if (!reserved.ok || !word.ok) throw new Error('reservation or word geometry failed');
  expect(word.rects[0]!.x).toBeGreaterThanOrEqual(reserved.anchor.x + reserved.anchor.width - 1);

  const withPlugin = await page.evaluate(async () => {
    const saved = await (window as unknown as ProbeWindow).__reservations.editor!.save();
    (window as unknown as ProbeWindow).__saved = new Uint8Array(saved!);
    return saved!.byteLength;
  });
  expect(withPlugin).toBeGreaterThan(0);

  await page.evaluate(() => (window as unknown as ProbeWindow).__reservations.setReadOnly(true));
  await expect(page.locator('[data-reserved="chip"]')).toBeVisible({ timeout: 60_000 });
  await expect(page.locator('[data-reserved="gap"]')).toBeVisible();

  await page.evaluate(() => (window as unknown as ProbeWindow).__reservations.setPlugins(false));
  await expect(page.locator('[data-reserved]')).toHaveCount(0);
  const same = await page.evaluate(async () => {
    const window_ = window as unknown as ProbeWindow;
    const saved = new Uint8Array((await window_.__reservations.editor!.save())!);
    const before = window_.__saved!;
    return saved.length === before.length && saved.every((value, index) => value === before[index]);
  });
  expect(same).toBe(true);
});
