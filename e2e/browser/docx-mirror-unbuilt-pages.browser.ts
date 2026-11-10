import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const heading = (number: number) => `Reading heading ${number}`;
const paragraph = (number: number) => `Searchable paragraph on reading page ${number}.`;

async function docxFile() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  const body = Array.from({ length: 48 }, (_, index) =>
    `<w:p><w:pPr>${index ? '<w:pageBreakBefore/>' : ''}<w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>${heading(index + 1)}</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t>${paragraph(index + 1)}</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml', `<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, worker: boolean, off = false) {
  await page.addInitScript(() => {
    document.addEventListener('copy', (event) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text) (window as unknown as { copied: string }).copied = text;
    });
  });
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}${off ? '&unbuiltPageMirror=off' : ''}`);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'unbuilt-pages.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await docxFile() });
  await expect(page.locator('[data-page-index="0"] canvas').first()).toBeVisible({ timeout: 60000 });
  await expect(page.locator('.canvas-page-mirror').getByText(heading(1), { exact: true })).toBeAttached();
  if (worker && !off) await expect(page.locator('[data-docx-unbuilt-page="30"]')).toBeAttached({ timeout: 60000 });
}

async function findFromStart(page: Page, text: string) {
  return page.evaluate((query) => {
    (document.activeElement as HTMLElement | null)?.blur();
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    const range = document.createRange();
    range.setStart(document.body, 0);
    range.collapse(true);
    selection.addRange(range);
    const found = (window as unknown as Window & { find(text: string): boolean }).find(query);
    const node = selection.anchorNode;
    const element = node instanceof Element ? node : node?.parentElement;
    return { found, unbuilt: element?.closest('[data-docx-unbuilt-page]')?.getAttribute('data-docx-unbuilt-page') ?? null,
      built: element?.closest('.canvas-page-mirror')?.closest('[data-page-index]')?.getAttribute('data-page-index') ?? null };
  }, text);
}

test('an unbuilt paragraph navigates to its page and hands focus to the matching built paragraph', async ({ page }) => {
  test.setTimeout(120000);
  await open(page, true);
  const section = page.getByRole('region', { name: 'Page 31', exact: true });
  const target = section.getByText(paragraph(31), { exact: true });
  await expect(target).toBeAttached();
  await expect(target).toHaveJSProperty('tagName', 'P');
  await expect(target).toHaveAttribute('tabindex', '-1');
  await expect(section.getByRole('heading')).toHaveCount(0);
  await target.focus();
  const mirror = page.locator('[data-page-index="30"] .canvas-page-mirror');
  const block = mirror.getByRole('paragraph').filter({ hasText: paragraph(31) });
  await expect(page.locator('[data-page-index="30"] canvas')).toBeVisible();
  await expect(section).toHaveCount(0);
  await expect(block).toBeFocused();
  await expect(block).toHaveText(paragraph(31));
  await expect(block).toHaveAttribute('tabindex', '-1');
  expect(await block.getAttribute('aria-level')).toBeNull();
  const handle = await block.elementHandle();
  await page.locator('.paged-editor__input').focus();
  expect(await handle!.getAttribute('tabindex')).toBeNull();
});

test('native find scrolls and builds an unbuilt page and a second find lands in its built mirror', async ({ page }) => {
  test.setTimeout(120000);
  await open(page, true);
  expect(await findFromStart(page, paragraph(31))).toEqual({ found: true, unbuilt: '30', built: null });
  await expect(page.locator('[data-page-index="30"] canvas')).toBeVisible();
  await expect(page.locator('[data-docx-unbuilt-page="30"]')).toHaveCount(0);
  expect(await page.evaluate(() => !document.activeElement?.closest('.canvas-page-mirror'))).toBe(true);
  await expect.poll(() => page.evaluate(() => getSelection()?.anchorNode?.parentElement?.closest('[data-page-index="30"] .canvas-page-mirror') ? 'built' : 'other')).toBe('built');
  await expect(page.locator('[data-page-index="30"] .canvas-page-mirror').getByText(paragraph(31), { exact: true })).toBeAttached();
  expect(await findFromStart(page, paragraph(31))).toEqual({ found: true, unbuilt: null, built: '30' });
  expect(await findFromStart(page, paragraph(31))).toEqual({ found: true, unbuilt: null, built: '30' });
});

test('select-all across unbuilt sections neither scrolls nor builds pages', async ({ page }) => {
  test.setTimeout(120000);
  await open(page, true);
  await expect(page.locator('[data-docx-unbuilt-page]').first()).toBeAttached();
  const state = () => page.evaluate(() => {
    let scroller: HTMLElement | null = document.querySelector<HTMLElement>('[data-page-index]');
    while (scroller && scroller.scrollHeight <= scroller.clientHeight) scroller = scroller.parentElement;
    return { top: (scroller ?? document.scrollingElement)?.scrollTop ?? 0,
      built: Array.from(document.querySelectorAll<HTMLElement>('[data-page-index]'), (el) => el.dataset.pageIndex).join(',') };
  });
  const before = await state();
  await page.evaluate(() => getSelection()?.selectAllChildren(document.body));
  await page.waitForTimeout(1500);
  expect(await state()).toEqual(before);
});

for (const worker of [false, true]) {
  test(`unbuilt mirror off preserves copy and pointer hit targets (${worker ? 'worker' : 'in-thread'} editor)`, async ({ page }) => {
    test.setTimeout(180000);
    const results: { copied: string; hit: string }[] = [];
    for (const off of [false, true]) {
      await open(page, worker, off);
      if (off || !worker) {
        await expect(page.locator('[data-docx-unbuilt-page]')).toHaveCount(0);
        await expect(page.locator('[data-docx-unbuilt-range]')).toHaveCount(0);
      }
      if (worker && off) expect((await findFromStart(page, paragraph(31))).found).toBe(false);
      const canvas = page.locator('[data-page-index="0"] canvas').first();
      await canvas.scrollIntoViewIfNeeded();
      const hit = await canvas.evaluate((element) => {
        const rect = element.getBoundingClientRect();
        const target = document.elementFromPoint(rect.left + 150, rect.top + 110);
        return `${target?.tagName}:${target?.className ?? ''}:${target?.getAttribute('data-page-index') ?? ''}`;
      });
      await canvas.click({ position: { x: 150, y: 110 } });
      const input = page.locator(worker ? '.paged-editor__input' : '[data-testid="yrs-input"]');
      await input.focus();
      await page.keyboard.press('ControlOrMeta+Home');
      await page.keyboard.press('Shift+End');
      await page.keyboard.press('ControlOrMeta+c');
      await expect.poll(() => page.evaluate(() => (window as unknown as { copied?: string }).copied ?? '')).toContain(heading(1));
      const copied = await page.evaluate(() => (window as unknown as { copied: string }).copied);
      results.push({ copied, hit });
    }
    expect(results[0]).toEqual(results[1]);
  });
}
