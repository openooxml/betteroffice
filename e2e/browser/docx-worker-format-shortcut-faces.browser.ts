import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

const BEFORE = 'The quick ';
const SELECTED = 'brown fox jumps';
const AFTER = ' over the lazy dog near the river bank.';
const TEXT = BEFORE + SELECTED + AFTER;
const CONTENT_RIGHT = 816 - 96;

type Run = { text: string; left: number; top: number; width: number; height: number; fontStyle: string; fontWeight: string };
type Mark = { className: string; left: number; top: number; width: number; height: number };
type Snapshot = { runs: Run[]; marks: Mark[] };

const arial = (extra = '') => `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>${extra}</w:rPr>`;
const run = (text: string, extra = '') => `<w:r>${arial(extra)}<w:t xml:space="preserve">${text}</w:t></w:r>`;

async function bareDocument(underlined: boolean) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const runs = underlined ? run(BEFORE) + run(SELECTED, '<w:b/><w:u w:val="single"/>') + run(AFTER) : run(TEXT);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${runs}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

function snapshot(page: Page): Promise<Snapshot> {
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

function lines(runs: Run[]): Run[][] {
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

function measure(runs: Run[]) {
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

async function settled(page: Page): Promise<Snapshot> {
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

const cases = [
  { name: 'italic', key: 'ControlOrMeta+I', button: 'Italic', underlined: false, face: 'LiberationSans-Italic' },
  { name: 'bold', key: 'ControlOrMeta+B', button: 'Bold', underlined: false, face: 'LiberationSans-Bold' },
  { name: 'italic over a bold underlined run', key: 'ControlOrMeta+I', button: 'Italic', underlined: true, face: 'LiberationSans-BoldItalic' },
] as const;

for (const arm of ['', '&peerless=0']) {
  const editor = arm ? 'main-thread' : 'worker';
  for (const entry of cases) {
    test(`docx ${editor} editor ${entry.name} shortcut loads its face and keeps the line`, async ({ page }, info) => {
      test.setTimeout(150_000);
      const label = `${editor}-${entry.name.replace(/\s+/g, '-')}`;
      const started = Date.now();
      let shortcutAt = Number.POSITIVE_INFINITY;
      const fonts: Array<{ ms: number; afterShortcut: boolean; path: string; status?: number }> = [];
      page.context().on('request', (request) => {
        const path = new URL(request.url()).pathname;
        if (!/\.(?:ttf|otf|woff2?|zst)$/i.test(path)) return;
        const record: (typeof fonts)[number] = { ms: Date.now() - started, afterShortcut: Date.now() >= shortcutAt, path };
        fonts.push(record);
        void request.response().then((response) => { record.status = response?.status(); }, () => { record.status = -1; });
      });

      await page.goto('/?format=docx' + arm);
      await page.locator('input[type=file]').last().setInputFiles({
        name: 'shortcut-faces.docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        buffer: await bareDocument(entry.underlined),
      });
      await expect(page.locator('.canvas-page-mirror').first()).toContainText('river bank', { timeout: 60_000 });
      if (!arm) await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 60_000 });
      const before = measure((await settled(page)).runs);
      const pageView = page.locator('.canvas-page, .paged-editor__page').first();
      await pageView.screenshot({ path: info.outputPath(`${label}-before.png`) });

      const input = page.locator('[data-testid=yrs-input], .paged-editor__input').first();
      await input.focus();
      await input.press('ControlOrMeta+Home');
      for (let step = 0; step < BEFORE.length; step += 1) await input.press('ArrowRight');
      for (let step = 0; step < SELECTED.length; step += 1) await input.press('Shift+ArrowRight');
      shortcutAt = Date.now();
      await input.press(entry.key);

      const button = page.getByTestId('formatting-bar').getByRole('button', { name: entry.button, exact: true });
      await expect.soft(button).toHaveAttribute('aria-pressed', 'true');
      await expect.poll(() => fonts.some((font) => font.path.includes(`/${entry.face}.ttf`)), { timeout: 15_000 })
        .toBe(true).catch(() => undefined);
      const after = await settled(page);
      const measured = measure(after.runs);
      await pageView.screenshot({ path: info.outputPath(`${label}-after.png`) });

      const selected = after.runs.filter((item) => item.text.includes('brown') || item.text.includes('jumps'));
      const selectedRight = selected.length ? Math.max(...selected.map((item) => item.left + item.width)) : 0;
      const underlines = after.marks.filter((mark) => mark.height <= 4 && selected.some((item) =>
        mark.left < item.left + item.width && mark.left + mark.width > item.left && Math.abs(mark.top - (item.top + item.height)) < item.height));

      console.log(`FONT-LOG ${label} ${JSON.stringify(fonts)}`);
      console.log(`LINES ${label} before=${JSON.stringify({ lineCount: before.lineCount, maxGap: +before.maxGap.toFixed(2), right: +before.right.toFixed(2) })} after=${JSON.stringify({ lineCount: measured.lineCount, maxGap: +measured.maxGap.toFixed(2), right: +measured.right.toFixed(2) })}`);
      console.log(`RUNS ${label} before=${JSON.stringify(before.runs)}`);
      console.log(`RUNS ${label} after=${JSON.stringify(measured.runs)}`);
      console.log(`MARKS ${label} ${JSON.stringify(after.marks.map((mark) => ({ c: mark.className, x: +mark.left.toFixed(2), w: +mark.width.toFixed(2), y: +mark.top.toFixed(2), h: +mark.height.toFixed(2) })))}`);

      expect.soft(fonts.some((font) => font.path.includes(`/${entry.face}.ttf`)), `${entry.face} requested`).toBe(true);
      expect.soft(before.lineCount, 'line count before the shortcut').toBe(1);
      expect.soft(measured.lineCount, 'line count after the shortcut').toBe(before.lineCount);
      expect.soft(measured.text.replace(/\s+/g, ''), 'paragraph text after the shortcut').toBe(TEXT.replace(/\s+/g, ''));
      expect.soft(measured.maxGap, 'largest gap between adjacent runs').toBeLessThanOrEqual(Math.max(before.maxGap, 0) + 6);
      expect.soft(measured.right, 'right edge of the line').toBeLessThanOrEqual(CONTENT_RIGHT + 1);
      expect.soft(Math.abs(measured.right - before.right), 'line width change').toBeLessThanOrEqual(12);
      expect.soft(selected.length, 'mirror runs holding the selected words').toBeGreaterThan(0);
      if (entry.name === 'bold') expect.soft(selected.every((item) => /bold|[6-9]00/.test(item.fontWeight)), 'selected runs are bold').toBe(true);
      else expect.soft(selected.every((item) => item.fontStyle === 'italic'), 'selected runs are italic').toBe(true);
      for (const mark of underlines) {
        expect.soft(mark.left + mark.width, 'underline right edge').toBeLessThanOrEqual(selectedRight + 2);
      }
    });
  }
}
