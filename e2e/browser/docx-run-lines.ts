import { expect, type Page } from 'playwright/test';

export type Run = { text: string; left: number; top: number; width: number; height: number; fontStyle: string; fontWeight: string };
export type Mark = { className: string; left: number; top: number; width: number; height: number };
export type Snapshot = { runs: Run[]; marks: Mark[] };
export type FontRequest = { ms: number; afterShortcut: boolean; path: string; status?: number; done?: boolean };

export function snapshot(page: Page): Promise<Snapshot> {
  return page.evaluate(() => {
    const mirror = document.querySelector('.canvas-page-mirror');
    const geometry = (element: HTMLElement) => ({
      left: parseFloat(element.style.left), top: parseFloat(element.style.top),
      width: parseFloat(element.style.width), height: parseFloat(element.style.height),
    });
    const runs = Array.from(mirror?.querySelectorAll<HTMLElement>('.layout-run-text') ?? [], (element) => ({
      text: element.textContent ?? '', ...geometry(element),
      fontStyle: element.style.fontStyle, fontWeight: element.style.fontWeight,
    }));
    const marks = Array.from(mirror?.querySelectorAll<HTMLElement>('.layout-decoration, .layout-mirror-line, .layout-mirror-rect') ?? [], (element) => ({
      className: element.className, ...geometry(element),
    }));
    return { runs, marks };
  });
}

export function lines(runs: Run[]): Run[][] {
  const sorted = [...runs].sort((a, b) => a.top + a.height / 2 - (b.top + b.height / 2) || a.left - b.left);
  const grouped: Run[][] = [];
  for (const entry of sorted) {
    const last = grouped.at(-1);
    const middle = entry.top + entry.height / 2;
    if (last && Math.abs(middle - (last[0]!.top + last[0]!.height / 2)) < last[0]!.height / 2) last.push(entry);
    else grouped.push([entry]);
  }
  return grouped.map((line) => line.sort((a, b) => a.left - b.left));
}

export function measure(runs: Run[]) {
  const grouped = lines(runs.filter((entry) => entry.text.length > 0));
  const gaps = grouped.flatMap((line) => line.slice(1).map((entry, index) => entry.left - (line[index]!.left + line[index]!.width)));
  return {
    lineCount: grouped.length,
    text: grouped.map((line) => line.map((entry) => entry.text).join('')).join('\n'),
    maxGap: gaps.length ? Math.max(...gaps) : 0,
    right: grouped.length ? Math.max(...grouped.flat().map((entry) => entry.left + entry.width)) : 0,
    runs: grouped.map((line) => line.map((entry) => ({
      text: entry.text, x: +entry.left.toFixed(2), right: +(entry.left + entry.width).toFixed(2),
      top: +entry.top.toFixed(2), style: entry.fontStyle, weight: entry.fontWeight,
    }))),
  };
}

export async function settled(page: Page): Promise<Snapshot> {
  let previous = '';
  let stableSince = Date.now();
  let latest: Snapshot = { runs: [], marks: [] };
  await expect.poll(async () => {
    latest = await snapshot(page);
    const key = JSON.stringify(latest);
    if (key !== previous || latest.runs.length === 0) { previous = key; stableSince = Date.now(); }
    return latest.runs.length > 0 && Date.now() - stableSince >= 2_000;
  }, { timeout: 30_000, intervals: [250] }).toBe(true);
  return latest;
}

export function recordFontRequests(page: Page) {
  const started = Date.now();
  let shortcutAt = Number.POSITIVE_INFINITY;
  const fonts: FontRequest[] = [];
  page.context().on('request', (request) => {
    const path = new URL(request.url()).pathname;
    if (!/\.(?:ttf|otf|woff2?|zst)$/i.test(path)) return;
    const record: FontRequest = { ms: Date.now() - started, afterShortcut: Date.now() >= shortcutAt, path };
    fonts.push(record);
    void request.response().then(async (response) => {
      record.status = response?.status();
      record.done = response ? !(await response.finished()) : false;
    }, () => { record.status = -1; });
  });
  return {
    fonts,
    markShortcut() { shortcutAt = Date.now(); },
    loaded(face: string, afterShortcut: boolean) {
      const pattern = new RegExp(`/${face}(?:[.-][\\w-]*)?\\.(?:ttf|otf|woff2?|zst)$`);
      return fonts.some((font) => (!afterShortcut || font.afterShortcut) && pattern.test(font.path) && font.status === 200 && font.done === true);
    },
  };
}

export function underlinesUnder(marks: Mark[], selected: Run[]): Mark[] {
  return marks.filter((mark) => mark.height <= 4 && selected.some((item) =>
    mark.left < item.left + item.width && mark.left + mark.width > item.left && Math.abs(mark.top - (item.top + item.height)) < item.height));
}
