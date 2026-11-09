import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

type Probe = {
  issues: { code: string; id: string; other?: string }[];
  hostKeys: { key: string; prevented: boolean }[];
  swallow: string | null;
  pluginRuns: number;
  errors: string[];
  load(bytes: number[], shortcuts: unknown): void;
  saved(): Promise<number[]>;
};
const PAGE = 'canvas[data-page-index="0"] >> nth=0';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function docx(): Promise<number[]> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document ${W}><w:body><w:p><w:r><w:t>Base text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

async function open(page: Page, mode: string, shortcuts: unknown, query = '') {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto(`/docx-shortcuts.html?mode=${mode}${query}`);
  const bytes = await docx();
  await page.evaluate(([data, map]) => (window as unknown as { __shortcuts: Probe }).__shortcuts.load(data as number[], map), [bytes, shortcuts] as const);
  await expect(page.locator(PAGE)).toBeVisible({ timeout: 120_000 });
  const input = page.locator('.paged-editor__input, [data-testid=yrs-input]').first();
  await input.focus();
  await input.press('ControlOrMeta+End');
  return input;
}

async function clipboard(page: Page) {
  await page.evaluate(() => navigator.clipboard.write([new ClipboardItem({
    'text/plain': new Blob(['RICH'], { type: 'text/plain' }),
    'text/html': new Blob(['<table><tr><td>RICH</td></tr></table>'], { type: 'text/html' }),
  })]));
}

async function documentXml(page: Page): Promise<string> {
  await page.waitForTimeout(400);
  const bytes = await page.evaluate(() => (window as unknown as { __shortcuts: Probe }).__shortcuts.saved());
  const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
  return zip.file('word/document.xml')!.async('string');
}

const richTable = /<w:tbl\b/;

for (const mode of ['worker', 'legacy']) {
  test.describe(`docx host shortcuts, ${mode}`, () => {
    test('a remapped plain paste leaves the old chord to the host', async ({ page }) => {
      const input = await open(page, mode, { pasteAsPlainText: 'Mod+Alt+V' });
      await clipboard(page);
      await input.press('Control+Shift+V');
      expect(await documentXml(page)).not.toContain('RICH');
      const keys = await page.evaluate(() => (window as unknown as { __shortcuts: Probe }).__shortcuts.hostKeys);
      expect(keys.at(-1)).toEqual({ key: 'v', prevented: false });
      await input.press('Control+Alt+V');
      const xml = await documentXml(page);
      expect(xml).toContain('RICH');
      expect(xml).not.toMatch(richTable);
    });

    test('the default plain paste drops formatting', async ({ page }) => {
      const input = await open(page, mode, {});
      await clipboard(page);
      await input.press('Control+Shift+V');
      const xml = await documentXml(page);
      expect(xml).toContain('RICH');
      expect(xml).not.toMatch(richTable);
    });

    test('a chord the host swallowed leaves the next paste rich', async ({ page }) => {
      const input = await open(page, mode, {});
      await clipboard(page);
      await page.evaluate(() => { (window as unknown as { __shortcuts: Probe }).__shortcuts.swallow = 'v'; });
      await input.press('Control+Shift+V');
      expect(await documentXml(page)).not.toContain('RICH');
      await page.evaluate(() => { (window as unknown as { __shortcuts: Probe }).__shortcuts.swallow = null; });
      await input.press('Control+V');
      expect(await documentXml(page)).toMatch(richTable);
    });

    test('a released paste chord leaves the context menu working', async ({ page }) => {
      const input = await open(page, mode, { paste: null });
      await clipboard(page);
      await input.press('Control+V');
      expect(await documentXml(page)).not.toContain('RICH');
      await page.locator(PAGE).click({ button: 'right', position: { x: 200, y: 200 } });
      const paste = page.getByRole('menuitem', { name: /^Paste(?! as)/ });
      await expect(paste).not.toContainText('Ctrl');
      await paste.click();
      await expect.poll(() => documentXml(page)).toContain('RICH');
    });

    test('menus show the remapped chord and aria form', async ({ page }) => {
      await open(page, mode, { pasteAsPlainText: 'Mod+Alt+V', selectAll: null });
      await page.locator(PAGE).click({ button: 'right', position: { x: 200, y: 200 } });
      const plain = page.getByRole('menuitem', { name: /plain text/i });
      await expect(plain).toContainText('Ctrl+Alt+V');
      await expect(plain).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+V');
      await expect(page.getByRole('menuitem', { name: /select all/i })).not.toContainText('Ctrl');
    });

    test('moving bold frees its chord for another command', async ({ page }) => {
      const input = await open(page, mode, { bold: 'Mod+Alt+B', italic: 'Mod+B' });
      const issues = await page.evaluate(() => (window as unknown as { __shortcuts: Probe }).__shortcuts.issues);
      expect(issues).toEqual([]);
      await input.press('Control+A');
      await input.press('Control+B');
      const italic = await documentXml(page);
      expect(italic).toContain('<w:i');
      expect(italic).not.toMatch(/<w:b\b/);
      await input.press('Control+Alt+B');
      expect(await documentXml(page)).toMatch(/<w:b\b/);
    });

    test('remapped select all, copy, cut and paste keep working', async ({ page }) => {
      const input = await open(page, mode, { selectAll: 'Mod+Alt+L', copy: 'Mod+Alt+A', cut: 'Mod+Alt+X', paste: 'Mod+Alt+P' });
      await input.press('Control+Alt+L');
      await input.press('Control+Alt+A');
      await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Base text');
      await input.press('ControlOrMeta+End');
      await input.press('Control+Alt+P');
      await expect.poll(() => documentXml(page)).toContain('Base textBase text');
      await input.press('Control+Alt+L');
      await input.press('Control+Alt+X');
      await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('Base textBase text');
      await expect.poll(() => documentXml(page)).not.toContain('Base text');
    });

    test('a released copy chord leaves the clipboard alone', async ({ page }) => {
      const input = await open(page, mode, { copy: null, cut: null });
      await page.evaluate(() => navigator.clipboard.writeText('KEEP'));
      await input.press('Control+A');
      await input.press('Control+C');
      await input.press('Control+X');
      await page.waitForTimeout(300);
      expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('KEEP');
      expect(await documentXml(page)).toContain('Base text');
    });

    test('a plugin takes a chord the host freed and follows a host remap', async ({ page }) => {
      const runs = () => page.evaluate(() => (window as unknown as { __shortcuts: Probe }).__shortcuts.pluginRuns);
      const freed = await open(page, mode, { insertLink: null });
      await freed.press('Control+K');
      await expect.poll(runs).toBe(1);
      const moved = await open(page, mode, { 'plugin:acme.keys/run': 'Mod+Alt+R' });
      await moved.press('Control+Alt+R');
      await expect.poll(runs).toBe(1);
    });

    test('a chord the input also handles runs the command it was given', async ({ page }) => {
      const input = await open(page, mode, { pasteAsPlainText: 'Mod+Z' });
      await clipboard(page);
      await input.press('q');
      await input.press('Control+Z');
      await expect.poll(() => documentXml(page)).toContain('RICH');
      expect(await documentXml(page)).toContain('q');
    });

    test('a native chord held by an unavailable plugin does nothing', async ({ page }) => {
      const input = await open(page, mode, { 'plugin:absent.keys/run': 'Mod+X' });
      await page.evaluate(() => navigator.clipboard.writeText('KEEP'));
      await input.press('Control+A');
      await input.press('Control+X');
      await page.waitForTimeout(300);
      expect(await documentXml(page)).toContain('Base text');
      expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('KEEP');
    });

    test('a second editor on the page keeps a released chord released', async ({ page }) => {
      const input = await open(page, mode, { paste: null }, '&two=1');
      await clipboard(page);
      await input.press('Control+V');
      expect(await documentXml(page)).not.toContain('RICH');
    });

    test('a remapped paste whose target moved is refused and reported', async ({ page }) => {
      const input = await open(page, mode, { paste: 'Mod+Alt+P' });
      await clipboard(page);
      await page.evaluate(() => {
        const { clipboard } = navigator;
        const later = <T,>(value: Promise<T>) => new Promise<T>((resolve) => setTimeout(() => resolve(value), 600));
        const read = clipboard.read.bind(clipboard);
        const readText = clipboard.readText.bind(clipboard);
        clipboard.read = () => later(read());
        clipboard.readText = () => later(readText());
      });
      await input.press('Control+Alt+P');
      await input.press('q');
      await expect.poll(() => page.evaluate(() => (window as unknown as { __shortcuts: Probe }).__shortcuts.errors.length)).toBe(1);
      const xml = await documentXml(page);
      expect(xml).toContain('q');
      expect(xml).not.toContain('RICH');
    });

    test('a host chord over a default reports the displaced command', async ({ page }) => {
      await open(page, mode, { pasteAsPlainText: 'Mod+K' });
      await expect.poll(() => page.evaluate(() => (window as unknown as { __shortcuts: Probe }).__shortcuts.issues))
        .toEqual([{ code: 'displaced-default', id: 'pasteAsPlainText', chord: 'Mod+k', other: 'insertLink' }]);
    });
  });
}
