import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { measure, recordFontRequests, settled, snapshot, underlinesUnder } from '../docx-run-lines';

const BEFORE = 'The committee reviewed the quarterly figures in detail and agreed that the regional offices should report their results in one format from next year on, so that ';
const SELECTED = 'every branch manager can compare';
const AFTER = ' the numbers without converting them first. The finance team will prepare a template, circulate it to all offices before the end of the month and collect the first reports in the spring, after which the board expects a short summary of the main differences between the regions and a proposal for the next budget cycle.';
const TEXT = BEFORE + SELECTED + AFTER;
const LAST_WORD = 'cycle.';
const CONTENT_RIGHT = 11906 / 15 - 96;

declare global { interface Window { __configuredFonts: { errors: string[]; ready: boolean; fontsReady: number[]; load(bytes: number[]): void } } }

const times = (extra = '') => `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${extra}<w:sz w:val="24"/></w:rPr>`;
const run = (text: string, extra = '') => `<w:r>${times(extra)}<w:t xml:space="preserve">${text}</w:t></w:r>`;

async function a4Document(selectedFormat: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const runs = selectedFormat ? run(BEFORE) + run(SELECTED, selectedFormat) + run(AFTER) : run(TEXT);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${runs}</w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

async function open(page: Page, selectedFormat: string) {
  await page.goto('/docx-configured-fonts.html');
  await page.waitForFunction(() => window.__configuredFonts?.ready === true, undefined, { timeout: 60_000 });
  await page.evaluate((bytes) => window.__configuredFonts.load(bytes), await a4Document(selectedFormat));
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(LAST_WORD, { timeout: 60_000 });
  await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 60_000 });
}

const bold = (weight: string) => /bold|[6-9]00/.test(weight);

// A font environment committed after the face's bytes arrived carries the face.
function faceInUse(page: Page, face: string) {
  return page.evaluate((source) => {
    const pattern = new RegExp(source);
    const arrivals = performance.getEntriesByType('resource')
      .filter((entry) => pattern.test(new URL(entry.name).pathname))
      .map((entry) => (entry as PerformanceResourceTiming).responseEnd);
    return arrivals.length > 0 && window.__configuredFonts.fontsReady.some((at) => at >= Math.max(...arrivals));
  }, `/${face}(?:[.-][\\w-]*)?\\.(?:ttf|otf|woff2?|zst)$`);
}

const cases = [
  { name: 'bold shortcut', key: 'ControlOrMeta+B', button: 'Bold', format: '', face: 'LiberationSerif-Bold' },
  { name: 'italic shortcut', key: 'ControlOrMeta+I', button: 'Italic', format: '', face: 'LiberationSerif-Italic' },
  { name: 'italic run present at open', key: '', button: 'Italic', format: '<w:i/>', face: 'LiberationSerif-Italic' },
  { name: 'italic shortcut over a bold underlined run', key: 'ControlOrMeta+I', button: 'Italic', format: '<w:b/><w:u w:val="single"/>', face: 'LiberationSerif-BoldItalic' },
] as const;

for (const entry of cases) {
  test(`docx worker editor with configured fonts, ${entry.name}, keeps a wrapping Times New Roman paragraph`, async ({ page }, info) => {
    test.setTimeout(180_000);
    const label = entry.name.replace(/\s+/g, '-');
    const fonts = recordFontRequests(page);
    const registryWarnings: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'warning' && message.text().startsWith('[fontRegistry]')) registryWarnings.push(message.text());
    });
    const formatted = (style: string, weight: string) => entry.button === 'Bold' ? bold(weight) : style === 'italic';
    await open(page, entry.format);
    const before = measure((await settled(page)).runs);
    const pageView = page.locator('.canvas-page, .paged-editor__page').first();
    await pageView.screenshot({ path: info.outputPath(`${label}-before.png`) });

    if (entry.key) {
      const input = page.locator('[data-testid=yrs-input], .paged-editor__input').first();
      await input.focus();
      await input.press('ControlOrMeta+Home');
      for (let step = 0; step < BEFORE.length; step += 1) await input.press('ArrowRight');
      for (let step = 0; step < SELECTED.length; step += 1) await input.press('Shift+ArrowRight');
      fonts.markShortcut();
      await input.press(entry.key);
      const button = page.getByTestId('formatting-bar').getByRole('button', { name: entry.button, exact: true });
      await expect.soft(button).toHaveAttribute('aria-pressed', 'true');
    }
    await expect.poll(() => fonts.loaded(entry.face, entry.key !== ''), {
      message: `${entry.face} fetched`, timeout: 30_000,
    }).toBe(true);
    await expect.poll(async () => await faceInUse(page, entry.face) &&
      (!entry.key || (await snapshot(page)).runs.some((item) => item.text.length > 0 && formatted(item.fontStyle, item.fontWeight))), {
      message: `${entry.face} registered and in use`, timeout: 30_000,
    }).toBe(true);
    const after = await settled(page);
    const measured = measure(after.runs);
    await pageView.screenshot({ path: info.outputPath(`${label}-after.png`) });

    const ordered = measured.runs.flat();
    const selected = after.runs.filter((item) => item.text.length > 0 && formatted(item.fontStyle, item.fontWeight));
    const selectedLeft = selected.length ? Math.min(...selected.map((item) => item.left)) : 0;
    const selectedRight = selected.length ? Math.max(...selected.map((item) => item.left + item.width)) : 0;
    const underlines = underlinesUnder(after.marks, selected);
    const baselineGap = entry.key ? Math.max(before.maxGap, 0) : 0;

    console.log(`FONT-LOG ${label} ${JSON.stringify(fonts.fonts)}`);
    console.log(`LINES ${label} before=${JSON.stringify({ lineCount: before.lineCount, maxGap: +before.maxGap.toFixed(2), right: +before.right.toFixed(2) })} after=${JSON.stringify({ lineCount: measured.lineCount, maxGap: +measured.maxGap.toFixed(2), right: +measured.right.toFixed(2) })}`);
    console.log(`RUNS ${label} after=${JSON.stringify(measured.runs)}`);
    console.log(`MARKS ${label} ${JSON.stringify(after.marks.map((mark) => ({ c: mark.className, x: +mark.left.toFixed(2), w: +mark.width.toFixed(2), y: +mark.top.toFixed(2), h: +mark.height.toFixed(2) })))}`);

    expect.soft(before.lineCount, 'line count at open').toBeGreaterThanOrEqual(3);
    expect.soft(measured.lineCount, 'line count after formatting').toBe(before.lineCount);
    expect.soft(measured.text.replace(/\s+/g, ''), 'paragraph text after formatting').toBe(TEXT.replace(/\s+/g, ''));
    expect.soft(measured.text.trimEnd().endsWith(LAST_WORD), 'last word still drawn').toBe(true);
    expect.soft(measured.maxGap, 'largest gap between adjacent runs').toBeLessThanOrEqual(baselineGap + 6);
    expect.soft(measured.right, 'right edge of the paragraph').toBeLessThanOrEqual(CONTENT_RIGHT + 1);
    expect.soft(ordered.filter((item) => formatted(item.style, item.weight)).map((item) => item.text).join(''), `text set ${entry.button.toLowerCase()}`).toBe(SELECTED);
    if (entry.format.includes('<w:u ')) {
      expect.soft(ordered.filter((item) => bold(item.weight)).map((item) => item.text).join(''), 'text kept bold').toBe(SELECTED);
      expect.soft(underlines.length, 'underline marks under the selected words').toBeGreaterThan(0);
    }
    for (const mark of underlines) {
      expect.soft(mark.left, 'underline left edge').toBeGreaterThanOrEqual(selectedLeft - 2);
      expect.soft(mark.left + mark.width, 'underline right edge').toBeLessThanOrEqual(selectedRight + 2);
    }
    expect(await page.evaluate(() => window.__configuredFonts.errors)).toEqual([]);
    expect(registryWarnings, 'font registry fallbacks').toEqual([]);
  });
}
