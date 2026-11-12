import { test, expect } from 'playwright/test';
import JSZip from 'jszip';
import { measure, recordFontRequests, settled, underlinesUnder } from './docx-run-lines';

const BEFORE = 'The quick ';
const SELECTED = 'brown fox jumps';
const AFTER = ' over the lazy dog near the river bank.';
const TEXT = BEFORE + SELECTED + AFTER;
const CONTENT_RIGHT = 816 - 96;

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
      const fonts = recordFontRequests(page);

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
      fonts.markShortcut();
      await input.press(entry.key);

      const button = page.getByTestId('formatting-bar').getByRole('button', { name: entry.button, exact: true });
      await expect.soft(button).toHaveAttribute('aria-pressed', 'true');
      await expect.poll(() => fonts.loaded(entry.face, true), {
        message: `${entry.face} fetched after the shortcut`, timeout: 30_000,
      }).toBe(true);
      const after = await settled(page);
      const measured = measure(after.runs);
      await pageView.screenshot({ path: info.outputPath(`${label}-after.png`) });

      const bold = (weight: string) => /bold|[6-9]00/.test(weight);
      const formatted = (style: string, weight: string) => entry.name === 'bold' ? bold(weight) : style === 'italic';
      const ordered = measured.runs.flat();
      const selected = after.runs.filter((item) => item.text.length > 0 && formatted(item.fontStyle, item.fontWeight));
      const selectedLeft = selected.length ? Math.min(...selected.map((item) => item.left)) : 0;
      const selectedRight = selected.length ? Math.max(...selected.map((item) => item.left + item.width)) : 0;
      const underlines = underlinesUnder(after.marks, selected);

      console.log(`FONT-LOG ${label} ${JSON.stringify(fonts.fonts)}`);
      console.log(`LINES ${label} before=${JSON.stringify({ lineCount: before.lineCount, maxGap: +before.maxGap.toFixed(2), right: +before.right.toFixed(2) })} after=${JSON.stringify({ lineCount: measured.lineCount, maxGap: +measured.maxGap.toFixed(2), right: +measured.right.toFixed(2) })}`);
      console.log(`RUNS ${label} before=${JSON.stringify(before.runs)}`);
      console.log(`RUNS ${label} after=${JSON.stringify(measured.runs)}`);
      console.log(`MARKS ${label} ${JSON.stringify(after.marks.map((mark) => ({ c: mark.className, x: +mark.left.toFixed(2), w: +mark.width.toFixed(2), y: +mark.top.toFixed(2), h: +mark.height.toFixed(2) })))}`);

      expect.soft(before.lineCount, 'line count before the shortcut').toBe(1);
      expect.soft(measured.lineCount, 'line count after the shortcut').toBe(before.lineCount);
      expect.soft(measured.text.replace(/\s+/g, ''), 'paragraph text after the shortcut').toBe(TEXT.replace(/\s+/g, ''));
      expect.soft(measured.maxGap, 'largest gap between adjacent runs').toBeLessThanOrEqual(Math.max(before.maxGap, 0) + 6);
      expect.soft(measured.right, 'right edge of the line').toBeLessThanOrEqual(CONTENT_RIGHT + 1);
      expect.soft(Math.abs(measured.right - before.right), 'line width change').toBeLessThanOrEqual(12);
      expect.soft(ordered.filter((item) => formatted(item.style, item.weight)).map((item) => item.text).join(''), `text set ${entry.button.toLowerCase()}`).toBe(SELECTED);
      if (entry.underlined) {
        expect.soft(ordered.filter((item) => bold(item.weight)).map((item) => item.text).join(''), 'text kept bold').toBe(SELECTED);
        expect.soft(underlines.length, 'underline marks under the selected words').toBeGreaterThan(0);
      }
      for (const mark of underlines) {
        expect.soft(mark.left, 'underline left edge').toBeGreaterThanOrEqual(selectedLeft - 2);
        expect.soft(mark.left + mark.width, 'underline right edge').toBeLessThanOrEqual(selectedRight + 2);
      }
    });
  }
}
